# 02-AzureAD-Functions.ps1
# Azure AD/Entra ID role audit functions - Certificate Authentication ONLY

# Optimized Azure AD Role Audit Function - Role-First Approach
# This approach iterates through roles first, then gets assignments, then resolves users

# Clean Get-AzureADRoleAudit function with unknown principal filtering
function Get-AzureADRoleAudit {
    param(
        [bool]$IncludePIM = $true,
        [string]$TenantId,
        [string]$ClientId,
        [string]$CertificateThumbprint,
        [switch]$IncludeUnknownPrincipals,  # New parameter to include deleted/unknown objects
        [bool]$IncludeAllPrincipalTypes = $true,   # Include groups and service principals
        [switch]$IncludeServiceRoles
    )
    
    $results = @()
    
    try {
        # Certificate authentication setup
        if ($TenantId -and $ClientId -and $CertificateThumbprint) {
            Set-M365AuditCertCredentials -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
        }
        
        # Verify certificate authentication is configured
        if (-not $script:AppConfig.UseAppAuth -or $script:AppConfig.AuthType -ne "Certificate") {
            throw "Certificate authentication is required for Azure AD role audit. Use Set-M365AuditCertCredentials first."
        }
        
        # Connect to Microsoft Graph with certificate authentication
        $context = Get-MgContext
        if (-not $context -or $context.AuthType -ne "AppOnly") {
            Write-Host "Connecting to Microsoft Graph with certificate authentication..." -ForegroundColor Yellow
            
            Connect-MgGraph -TenantId $script:AppConfig.TenantId -ClientId $script:AppConfig.ClientId -CertificateThumbprint $script:AppConfig.CertificateThumbprint -NoWelcome
            
            # Verify app-only authentication
            $context = Get-MgContext
            if ($context.AuthType -ne "AppOnly") {
                throw "Expected app-only authentication but got: $($context.AuthType). Check certificate configuration."
            }
            
            Write-Host "✓ Connected with certificate authentication" -ForegroundColor Green
        }
        
        # === STEP 1: Get all role definitions ===
        $ServiceRoles = @(
            'Teams Administrator',
            "Intune Administrator",
            "Compliance Data Administrator",
            "Exchange Recipient Administrator",
            "Compliance Administrator",
            "Exchange Administrator",
            "SharePoint Administrator",
            "Teams Communications Administrator",
            "Teams Communications Support Engineer", 
            "Teams Communications Support Specialist",
            "Teams Devices Administrator",
            "Teams Telephony Administrator",
            "Security Operator",
            "Power Platform Administrator"
        )

        Write-Host "Retrieving Azure AD role definitions..." -ForegroundColor Cyan
        $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition -All

        # Only include service roles if requested.
        if (-not $IncludeServiceRoles) {
            $roleDefinitions = $roleDefinitions | Where-Object {$_.DisplayName -notIn $ServiceRoles}
        }

        Write-Host "Found $($roleDefinitions.Count) role definitions" -ForegroundColor Green
        
        # === STEP 2: Get ALL types of role assignments using the helper function ===
        Write-Host "Retrieving all Azure AD role assignments (active, PIM eligible, PIM active)..." -ForegroundColor Cyan
        
        # Use the helper function to get all assignment types
        $allAssignments = Get-RoleAssignmentsForService -RoleDefinitions $roleDefinitions -ServiceName "Azure AD" -IncludePIM:$IncludePIM
        
        # === STEP 3: Convert assignments to standardized results using the helper function ===
        Write-Host "Processing Azure AD role assignments..." -ForegroundColor Cyan
        
        # Define overarching roles (for Azure AD, all roles are overarching by definition)
        $overarchingRoles = $roleDefinitions.DisplayName  # All Azure AD roles are overarching
        
        # Convert to standardized results using the helper function with filtering
        $convertParams = @{
            Assignments = $allAssignments
            RoleDefinitions = $roleDefinitions
            ServiceName = "Azure AD/Entra ID"
            OverarchingRoles = $overarchingRoles
        }
        
        # Add optional parameters if specified
        if ($IncludeAllPrincipalTypes) {
            $convertParams.IncludeAllPrincipalTypes = $true
        }
        
        if ($IncludeUnknownPrincipals) {
            $convertParams.IncludeUnknownPrincipals = $true
        }
        
        $results = ConvertTo-ServiceAssignmentResults @convertParams
        
        Write-Host "✓ Azure AD role audit completed. Found $($results.Count) role assignments" -ForegroundColor Green
        
        # === STEP 4: Enhanced summary reporting ===
        if ($results.Count -gt 0) {
            $userResults = $results | Where-Object { $_.PrincipalType -eq "User" }
            $serviceResults = $results | Where-Object { $_.PrincipalType -eq "ServicePrincipal" }
            $groupResults = $results | Where-Object { $_.PrincipalType -eq "Group" }
            $unknownResults = $results | Where-Object { $_.PrincipalType -eq "Unknown" }
            $activeResults = $results | Where-Object { $_.AssignmentType -eq "Active" }
            $pimEligibleResults = $results | Where-Object { $_.AssignmentType -eq "Eligible (PIM)" }
            $pimActiveResults = $results | Where-Object { $_.AssignmentType -eq "Active (PIM)" }
            
            Write-Host ""
            Write-Host "=== COMPREHENSIVE AUDIT SUMMARY ===" -ForegroundColor Cyan
            Write-Host "Total assignments: $($results.Count)" -ForegroundColor White
            Write-Host "User assignments: $($userResults.Count)" -ForegroundColor White
            
            if ($IncludeAllPrincipalTypes) {
                Write-Host "Service principal assignments: $($serviceResults.Count)" -ForegroundColor White
                Write-Host "Group assignments: $($groupResults.Count)" -ForegroundColor White
            }
            
            if ($unknownResults.Count -gt 0) {
                Write-Host "Unknown principal assignments: $($unknownResults.Count)" -ForegroundColor Yellow
                Write-Host "  (These are likely deleted users/groups/apps with orphaned assignments)" -ForegroundColor Gray
            }
            
            Write-Host "Active assignments: $($activeResults.Count)" -ForegroundColor White
            Write-Host "PIM eligible assignments: $($pimEligibleResults.Count)" -ForegroundColor Green
            Write-Host "PIM active assignments: $($pimActiveResults.Count)" -ForegroundColor Green
            
            # Show top roles across all assignment types
            $topRoles = $results | Group-Object RoleName | Sort-Object Count -Descending | Select-Object -First 5
            Write-Host "Top roles (all assignment types):" -ForegroundColor Cyan
            foreach ($role in $topRoles) {
                $roleBreakdown = $results | Where-Object { $_.RoleName -eq $role.Name } | Group-Object AssignmentType
                $breakdown = ($roleBreakdown | ForEach-Object { "$($_.Name): $($_.Count)" }) -join ", "
                Write-Host "  $($role.Name): $($role.Count) total [$breakdown]" -ForegroundColor White
            }
            
            # Recommendations based on findings
            Write-Host ""
            Write-Host "=== RECOMMENDATIONS ===" -ForegroundColor Yellow
            
            $globalAdmins = $results | Where-Object { $_.RoleName -eq "Global Administrator" }
            if ($globalAdmins.Count -gt 5) {
                Write-Host "⚠️ Consider reducing Global Administrator count from $($globalAdmins.Count) to 5 or fewer" -ForegroundColor Yellow
            }
            
            $disabledUsers = $results | Where-Object { $_.UserEnabled -eq $false }
            if ($disabledUsers.Count -gt 0) {
                Write-Host "⚠️ $($disabledUsers.Count) disabled users still have role assignments - consider cleanup" -ForegroundColor Yellow
            }
            
            if ($pimEligibleResults.Count -eq 0 -and $activeResults.Count -gt 0) {
                Write-Host "💡 Consider implementing PIM for eligible assignments to reduce standing privileges" -ForegroundColor Cyan
            }
            
            # Show cleanup recommendation if we filtered unknown principals
            if ($IncludeUnknownPrincipals -eq $false) {
                Write-Host "💡 Use -IncludeUnknownPrincipals to see orphaned assignments for cleanup" -ForegroundColor Cyan
            }
        }
        
    }
    catch {
        Write-Error "Error in Azure AD role audit: $($_.Exception.Message)"
        Write-Error "Stack trace: $($_.ScriptStackTrace)"
        
        # Provide specific troubleshooting guidance
        if ($_.Exception.Message -like "*certificate*") {
            Write-Host ""
            Write-Host "Certificate Authentication Setup Required:" -ForegroundColor Red
            Write-Host "1. Create certificate: New-M365AuditCertificate" -ForegroundColor White
            Write-Host "2. Upload .cer file to Azure AD app registration" -ForegroundColor White
            Write-Host "3. Configure: Set-M365AuditCertCredentials -TenantId <id> -ClientId <id> -CertificateThumbprint <thumbprint>" -ForegroundColor White
            Write-Host "4. Verify: Get-M365AuditCertificate" -ForegroundColor White
        }
        elseif ($_.Exception.Message -like "*permission*") {
            Write-Host ""
            Write-Host "Required Application Permissions (Grant Admin Consent):" -ForegroundColor Red
            Write-Host "• Directory.Read.All" -ForegroundColor White
            Write-Host "• RoleManagement.Read.All" -ForegroundColor White
            Write-Host "• User.Read.All" -ForegroundColor White
            Write-Host "Run: Get-M365AuditRequiredPermissions for complete list" -ForegroundColor White
        }
        
        throw
    }
    
    return $results
}

function Get-TeamsRoleAudit {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$CertificateThumbprint,
        [switch]$IncludeAzureADRoles  # New parameter to control inclusion of overarching roles
    )
    
    $results = @()
    
    try {
        # Certificate authentication is required for this function
        if ($TenantId -and $ClientId -and $CertificateThumbprint) {
            Set-M365AuditCertCredentials -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
        }
        
        # Verify certificate authentication is configured
        if (-not $script:AppConfig.UseAppAuth -or $script:AppConfig.AuthType -ne "Certificate") {
            throw "Certificate authentication is required for Teams role audit. Use Set-M365AuditCertCredentials first."
        }
        
        # Connect to Microsoft Graph if not already connected with certificate auth
        $context = Get-MgContext
        if (-not $context -or $context.AuthType -ne "AppOnly") {
            Write-Host "Connecting to Microsoft Graph for Teams roles..." -ForegroundColor Yellow
            
            Connect-MgGraph -TenantId $script:AppConfig.TenantId -ClientId $script:AppConfig.ClientId -CertificateThumbprint $script:AppConfig.CertificateThumbprint #-NoWelcome
            
            # Verify app-only authentication
            $context = Get-MgContext
            if ($context.AuthType -ne "AppOnly") {
                throw "Expected app-only authentication but got: $($context.AuthType). Check certificate configuration."
            }
            
            Write-Host "✓ Connected with certificate authentication" -ForegroundColor Green
        }
        
        # Teams-specific Azure AD roles (NOT overarching roles)
        $teamsSpecificRoles = @(
            "Teams Administrator",
            "Teams Communications Administrator",
            "Teams Communications Support Engineer", 
            "Teams Communications Support Specialist",
            "Teams Devices Administrator",
            "Teams Telephony Administrator"
        )
        
        # Overarching roles that should only appear in Azure AD audit
        $overarchingRoles = @(
            "Global Administrator",
            "Security Administrator",
            "Security Reader",
            "Cloud Application Administrator",
            "Application Administrator",
            "Privileged Authentication Administrator",
            "Privileged Role Administrator"
        )
        
        # Determine which roles to include based on parameter
        $rolesToInclude = if ($IncludeAzureADRoles) {
            $teamsSpecificRoles + $overarchingRoles
        } else {
            $teamsSpecificRoles
        }
        
        # FIX 1: Add -All parameter to get ALL role definitions
        Write-Host "Retrieving Teams-related Azure AD roles..." -ForegroundColor Cyan
        $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition -All | Where-Object { $_.DisplayName -in $rolesToInclude }
        Write-Host "Found $($roleDefinitions.Count) Teams role definitions" -ForegroundColor Green

        $allAssignments = Get-RoleAssignmentsForService -RoleDefinitions $roleDefinitions -ServiceName "Teams" -IncludePIM

        Write-Host "Total Teams assignments across all types: $($allAssignments.Count)" -ForegroundColor Green
        
        $convertParams = @{
            Assignments = $allAssignments
            RoleDefinitions = $roleDefinitions
            ServiceName = "Microsoft Teams"
            OverarchingRoles = $overarchingRoles
        }

        $results = ConvertTo-ServiceAssignmentResults @convertParams
        
        
        # Process all assignments
<#         foreach ($assignment in $allAssignments) {
            try {
                $role = $roleDefinitions | Where-Object { $_.Id -eq $assignment.RoleDefinitionId }
                
                # Determine assignment type based on source
                $assignmentType = switch ($assignment.AssignmentSource) {
                    "Active" { "Active" }
                    "PIMEligible" { "Eligible (PIM)" }
                    "PIMActive" { "Active (PIM)" }
                    default { "Active" }
                }
                
                # Resolve principal (users, groups, service principals)
                $principalInfo = @{
                    UserPrincipalName = "Unknown"
                    DisplayName = "Unknown"
                    UserId = $assignment.PrincipalId
                    UserEnabled = $null
                    #LastSignIn = $null
                    OnPremisesSyncEnabled = $null
                    PrincipalType = "Unknown"
                }
                
                # Try as user
                try {
                    $user = Get-MgUser -UserId $assignment.PrincipalId -Property "UserPrincipalName,DisplayName,AccountEnabled,OnPremisesSyncEnabled" -ErrorAction SilentlyContinue
                    if ($user) {
                        $principalInfo.UserPrincipalName = $user.UserPrincipalName
                        $principalInfo.DisplayName = $user.DisplayName
                        $principalInfo.UserEnabled = $user.AccountEnabled
                        $principalInfo.OnPremisesSyncEnabled = $user.OnPremisesSyncEnabled
                        $principalInfo.PrincipalType = "User"
                    }
                }
                catch { }
                
                # Try as service principal if not user
                if ($principalInfo.PrincipalType -eq "Unknown") {
                    try {
                        $app = Get-MgServicePrincipal -ServicePrincipalId $assignment.PrincipalId -Property "AppId,DisplayName,AccountEnabled" -ErrorAction SilentlyContinue
                        if ($app) {
                            $principalInfo.UserPrincipalName = $app.AppId
                            $principalInfo.DisplayName = "$($app.DisplayName) (Application)"
                            $principalInfo.UserEnabled = $app.AccountEnabled
                            $principalInfo.PrincipalType = "ServicePrincipal"
                        }
                    }
                    catch { }
                }
                
                # Try as group if still unknown
                if ($principalInfo.PrincipalType -eq "Unknown") {
                    try {
                        $group = Get-MgGroup -GroupId $assignment.PrincipalId -Property "Mail,DisplayName" -ErrorAction SilentlyContinue
                        if ($group) {
                            $principalInfo.UserPrincipalName = $group.Mail
                            $principalInfo.DisplayName = "$($group.DisplayName) (Group)"
                            $principalInfo.PrincipalType = "Group"
                        }
                    }
                    catch { }
                }
                
                # Determine role scope for enhanced deduplication
                $roleScope = if ($role.DisplayName -in $overarchingRoles) { "Overarching" } else { "Service-Specific" }
                
                $results += [PSCustomObject]@{
                    Service = "Microsoft Teams"
                    UserPrincipalName = $principalInfo.UserPrincipalName
                    DisplayName = $principalInfo.DisplayName
                    UserId = $principalInfo.UserId
                    RoleName = $role.DisplayName
                    RoleDefinitionId = $assignment.RoleDefinitionId
                    RoleScope = $roleScope  # New property for enhanced deduplication
                    AssignmentType = $assignmentType
                    AssignedDateTime = $assignment.CreatedDateTime
                    UserEnabled = $principalInfo.UserEnabled
                    #LastSignIn = $principalInfo.LastSignIn
                    Scope = $assignment.DirectoryScopeId
                    AssignmentId = $assignment.Id
                    #AuthenticationType = "Certificate"
                    PrincipalType = $principalInfo.PrincipalType
                    OnPremisesSyncEnabled = $principalInfo.OnPremisesSyncEnabled
                    PIMStartDateTime = $assignment.ScheduleInfo.StartDateTime
                    PIMEndDateTime = $assignment.ScheduleInfo.Expiration.EndDateTime
                }
                
            }
            catch {
                Write-Verbose "Error processing Teams assignment: $($_.Exception.Message)"
            }
        }
 #>        
        Write-Host "✓ Teams role audit completed. Found $($results.Count) role assignments (including PIM)" -ForegroundColor Green
        
        # Provide feedback about role filtering
        if (-not $IncludeAzureADRoles) {
            Write-Host "  (Excluding overarching Azure AD roles - use -IncludeAzureADRoles to include)" -ForegroundColor Yellow
        }
        
        # Show breakdown
        if ($results.Count -gt 0) {
            $typeSummary = $results | Group-Object PrincipalType
            $assignmentTypeSummary = $results | Group-Object AssignmentType
            $scopeSummary = $results | Group-Object RoleScope
            
            Write-Host "Principal types:" -ForegroundColor Cyan
            foreach ($type in $typeSummary) {
                Write-Host "  $($type.Name): $($type.Count)" -ForegroundColor White
            }
            
            Write-Host "Assignment types:" -ForegroundColor Cyan
            foreach ($type in $assignmentTypeSummary) {
                Write-Host "  $($type.Name): $($type.Count)" -ForegroundColor White
            }
            
            Write-Host "Role scope:" -ForegroundColor Cyan
            foreach ($scope in $scopeSummary) {
                Write-Host "  $($scope.Name): $($scope.Count)" -ForegroundColor White
            }
        }
        
    }
    catch {
        Write-Error "Error auditing Teams roles: $($_.Exception.Message)"
        
        if ($_.Exception.Message -like "*certificate*") {
            Write-Host "Certificate authentication required for Teams role audit" -ForegroundColor Red
            Write-Host "Use: Set-M365AuditCertCredentials -TenantId <id> -ClientId <id> -CertificateThumbprint <thumbprint>" -ForegroundColor White
        }
        
        throw
    }
    
    return $results
}

# Enhanced Microsoft Defender Role Audit Function with Azure AD Role Filtering
# Add to 03-AzureAD-Functions.ps1

function Get-DefenderRoleAudit {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$CertificateThumbprint,
        [switch]$IncludeAzureADRoles  # New parameter to control inclusion of overarching roles
    )
    
    $results = @()
    
    try {
        # Certificate authentication is required for this function
        if ($TenantId -and $ClientId -and $CertificateThumbprint) {
            Set-M365AuditCertCredentials -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
        }
        
        # Verify certificate authentication is configured
        if (-not $script:AppConfig.UseAppAuth -or $script:AppConfig.AuthType -ne "Certificate") {
            throw "Certificate authentication is required for Defender role audit. Use Set-M365AuditCertCredentials first."
        }
        
        # Connect to Microsoft Graph if not already connected with certificate auth
        $context = Get-MgContext
        if (-not $context -or $context.AuthType -ne "AppOnly") {
            Write-Host "Connecting to Microsoft Graph for Defender roles..." -ForegroundColor Yellow
            
            Connect-MgGraph -TenantId $script:AppConfig.TenantId -ClientId $script:AppConfig.ClientId -CertificateThumbprint $script:AppConfig.CertificateThumbprint -NoWelcome
            
            # Verify app-only authentication
            $context = Get-MgContext
            if ($context.AuthType -ne "AppOnly") {
                throw "Expected app-only authentication but got: $($context.AuthType). Check certificate configuration."
            }
            
            Write-Host "✓ Connected with certificate authentication" -ForegroundColor Green
        }
        
        # Defender-specific roles (limited set)
        $defenderSpecificRoles = @(
            "Security Operator"
        )
        
        # Overarching security roles that should only appear in Azure AD audit
        $overarchingRoles = @(
            "Security Administrator",
            "Security Reader",
            "Global Administrator",
            "Cloud Application Administrator",
            "Application Administrator",
            "Privileged Authentication Administrator",
            "Privileged Role Administrator"
        )
        
        # Determine which roles to include based on parameter
        $rolesToInclude = if ($IncludeAzureADRoles) {
            $defenderSpecificRoles + $overarchingRoles
        } else {
            $defenderSpecificRoles
        }
        
        # FIX 1: Add -All parameter to get ALL role definitions
        Write-Host "Retrieving Defender-related Azure AD roles..." -ForegroundColor Cyan
        $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition -All | Where-Object { $_.DisplayName -in $rolesToInclude }
        Write-Host "Found $($roleDefinitions.Count) Defender role definitions" -ForegroundColor Green

        $AllAssignments = Get-RoleAssignmentsForService -RoleDefinitions $roleDefinitions -ServiceName "Defender" -IncludePIM
                    
        Write-Host "Total Defender assignments across all types: $($allAssignments.Count)" -ForegroundColor Green

        $convertParams = @{
            Assignments = $AllAssignments
            RoleDefinitions = $roleDefinitions
            ServiceName = "Microsoft Defender"
            OverArchingRoles = $overarchingRoles
        }

        $results = ConvertTo-ServiceAssignmentResults @convertParams
        
        # Process all assignments
<#         foreach ($assignment in $allAssignments) {
            try {
                $role = $roleDefinitions | Where-Object { $_.Id -eq $assignment.RoleDefinitionId }
                
                # Determine assignment type based on source
                $assignmentType = switch ($assignment.AssignmentSource) {
                    "Active" { "Active" }
                    "PIMEligible" { "Eligible (PIM)" }
                    "PIMActive" { "Active (PIM)" }
                    default { "Active" }
                }
                
                # Resolve principal (users, groups, service principals)
                $principalInfo = @{
                    UserPrincipalName = "Unknown"
                    DisplayName = "Unknown"
                    UserId = $assignment.PrincipalId
                    UserEnabled = $null
                    OnPremisesSyncEnabled = $null
                    PrincipalType = "Unknown"
                }
                
                # Try as user
                try {
                    $user = Get-MgUser -UserId $assignment.PrincipalId -Property "UserPrincipalName,DisplayName,AccountEnabled,OnPremisesSyncEnabled" -ErrorAction SilentlyContinue
                    if ($user) {
                        $principalInfo.UserPrincipalName = $user.UserPrincipalName
                        $principalInfo.DisplayName = $user.DisplayName
                        $principalInfo.UserEnabled = $user.AccountEnabled
                        $PrincipalInfo.OnPremisesSyncEnabled = $user.OnPremisesSyncEnabled
                        $principalInfo.PrincipalType = "User"
                    }
                }
                catch { }
                
                # Try as service principal if not user
                if ($principalInfo.PrincipalType -eq "Unknown") {
                    try {
                        $app = Get-MgServicePrincipal -ServicePrincipalId $assignment.PrincipalId -Property "AppId,DisplayName,AccountEnabled" -ErrorAction SilentlyContinue
                        if ($app) {
                            $principalInfo.UserPrincipalName = $app.AppId
                            $principalInfo.DisplayName = "$($app.DisplayName) (Application)"
                            $principalInfo.UserEnabled = $app.AccountEnabled
                            $principalInfo.PrincipalType = "ServicePrincipal"
                        }
                    }
                    catch { }
                }
                
                # Try as group if still unknown
                if ($principalInfo.PrincipalType -eq "Unknown") {
                    try {
                        $group = Get-MgGroup -GroupId $assignment.PrincipalId -Property "Mail,DisplayName" -ErrorAction SilentlyContinue
                        if ($group) {
                            $principalInfo.UserPrincipalName = $group.Mail
                            $principalInfo.DisplayName = "$($group.DisplayName) (Group)"
                            $principalInfo.PrincipalType = "Group"
                        }
                    }
                    catch { }
                }
                
                # Determine role scope for enhanced deduplication
                $roleScope = if ($role.DisplayName -in $overarchingRoles) { "Overarching" } else { "Service-Specific" }
                
                $results += [PSCustomObject]@{
                    Service = "Microsoft Defender"
                    UserPrincipalName = $principalInfo.UserPrincipalName
                    DisplayName = $principalInfo.DisplayName
                    UserId = $principalInfo.UserId
                    RoleName = $role.DisplayName
                    RoleDefinitionId = $assignment.RoleDefinitionId
                    RoleScope = $roleScope  # New property for enhanced deduplication
                    AssignmentType = $assignmentType
                    AssignedDateTime = $assignment.CreatedDateTime
                    UserEnabled = $principalInfo.UserEnabled
                    #LastSignIn = $principalInfo.LastSignIn
                    Scope = $assignment.DirectoryScopeId
                    AssignmentId = $assignment.Id
                    #AuthenticationType = "Certificate"
                    PrincipalType = $principalInfo.PrincipalType
                    OnPremisesSyncEnabled = $principalInfo.OnPremisesSyncEnabled
                    PIMStartDateTime = $assignment.ScheduleInfo.StartDateTime
                    PIMEndDateTime = $assignment.ScheduleInfo.Expiration.EndDateTime
                }
                
            }
            catch {
                Write-Verbose "Error processing Defender assignment: $($_.Exception.Message)"
            }
        } #>
        
        Write-Host "✓ Defender role audit completed. Found $($results.Count) role assignments (including PIM)" -ForegroundColor Green
        
        # Provide feedback about role filtering
        if (-not $IncludeAzureADRoles) {
            Write-Host "  (Excluding overarching Azure AD roles - use -IncludeAzureADRoles to include)" -ForegroundColor Yellow
        }
        
        # Show breakdown
        if ($results.Count -gt 0) {
            $typeSummary = $results | Group-Object PrincipalType
            $assignmentTypeSummary = $results | Group-Object AssignmentType
            $scopeSummary = $results | Group-Object RoleScope
            
            Write-Host "Principal types:" -ForegroundColor Cyan
            foreach ($type in $typeSummary) {
                Write-Host "  $($type.Name): $($type.Count)" -ForegroundColor White
            }
            
            Write-Host "Assignment types:" -ForegroundColor Cyan
            foreach ($type in $assignmentTypeSummary) {
                Write-Host "  $($type.Name): $($type.Count)" -ForegroundColor White
            }
            
            Write-Host "Role scope:" -ForegroundColor Cyan
            foreach ($scope in $scopeSummary) {
                Write-Host "  $($scope.Name): $($scope.Count)" -ForegroundColor White
            }
        }
        
    }
    catch {
        Write-Error "Error auditing Defender roles: $($_.Exception.Message)"
        
        if ($_.Exception.Message -like "*certificate*") {
            Write-Host "Certificate authentication required for Defender role audit" -ForegroundColor Red
            Write-Host "Use: Set-M365AuditCertCredentials -TenantId <id> -ClientId <id> -CertificateThumbprint <thumbprint>" -ForegroundColor White
        }
        
        throw
    }
    
    return $results
}

# Enhanced Power Platform Azure AD Role Audit Function with Azure AD Role Filtering
# Add to 03-AzureAD-Functions.ps1

function Get-PowerPlatformAzureADRoleAudit {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$CertificateThumbprint,
        [switch]$IncludeAzureADRoles  # New parameter to control inclusion of overarching roles
    )
    
    $results = @()
    
    try {
        # Certificate authentication is required for this function
        if ($TenantId -and $ClientId -and $CertificateThumbprint) {
            Set-M365AuditCertCredentials -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
        }
        
        # Verify certificate authentication is configured
        if (-not $script:AppConfig.UseAppAuth -or $script:AppConfig.AuthType -ne "Certificate") {
            throw "Certificate authentication is required for Power Platform role audit. Use Set-M365AuditCertCredentials first."
        }
        
        # Connect to Microsoft Graph if not already connected with certificate auth
        $context = Get-MgContext
        if (-not $context -or $context.AuthType -ne "AppOnly") {
            Write-Host "Connecting to Microsoft Graph for Power Platform Azure AD roles..." -ForegroundColor Yellow
            
            Connect-MgGraph -TenantId $script:AppConfig.TenantId -ClientId $script:AppConfig.ClientId -CertificateThumbprint $script:AppConfig.CertificateThumbprint -NoWelcome
            
            # Verify app-only authentication
            $context = Get-MgContext
            if ($context.AuthType -ne "AppOnly") {
                throw "Expected app-only authentication but got: $($context.AuthType). Check certificate configuration."
            }
            
            Write-Host "✓ Connected with certificate authentication" -ForegroundColor Green
        }
        
        # Power Platform-specific Azure AD roles (NOT overarching roles)
        $powerPlatformSpecificRoles = @(
            "Power Platform Administrator"            
        )
        
        # Overarching roles that should only appear in Azure AD audit
        $overarchingRoles = @(
            "Global Administrator",
            "Security Administrator",
            "Security Reader",
            "Cloud Application Administrator",
            "Application Administrator",
            "Privileged Authentication Administrator",
            "Privileged Role Administrator" # This is overarching - covers all Power Platform services
        )
        
        # Determine which roles to include based on parameter
        $rolesToInclude = if ($IncludeAzureADRoles) {
            $powerPlatformSpecificRoles + $overarchingRoles
        } else {
            $powerPlatformSpecificRoles
        }
        
        # FIX 1: Add -All parameter to get ALL role definitions
        Write-Host "Retrieving Power Platform-related Azure AD roles..." -ForegroundColor Cyan
        $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition -All | Where-Object { $_.DisplayName -in $rolesToInclude }

        Write-Host "Found $($roleDefinitions.Count) Power Platform role definitions" -ForegroundColor Green

        $allAssignments = Get-RoleAssignmentsForService -RoleDefinitions $roleDefinitions -ServiceName "Power Platform" -IncludePIM
               
        Write-Host "Total Power Platform assignments across all types: $($allAssignments.Count)" -ForegroundColor Green

          $convertParams = @{
            Assignments = $allAssignments
            RoleDefinitions = $roleDefinitions
            ServiceName = "Power Platform"
            OverarchingRoles = $overarchingRoles
        }
        
        $results = ConvertTo-ServiceAssignmentResults @convertParams

        # Process all assignments (regular + PIM eligible + PIM active)
<#         foreach ($assignment in $allAssignments) {
            try {
                $role = $roleDefinitions | Where-Object { $_.Id -eq $assignment.RoleDefinitionId }
                
                # Determine assignment type based on source
                $assignmentType = switch ($assignment.AssignmentSource) {
                    "Active" { "Active" }
                    "PIMEligible" { "Eligible (PIM)" }
                    "PIMActive" { "Active (PIM)" }
                    default { "Active" }
                }
                
                # Resolve principal (users, groups, service principals)
                $principalInfo = @{
                    UserPrincipalName = "Unknown"
                    DisplayName = "Unknown"
                    UserId = $assignment.PrincipalId
                    UserEnabled = $null
                    OnPremisesSyncEnabled = $Null
                    PrincipalType = "Unknown"
                }
                
                # Try as user first
                try {
                    $user = Get-MgUser -UserId $assignment.PrincipalId -Property "UserPrincipalName,DisplayName,AccountEnabled,OnPremisesSyncEnabled" -ErrorAction SilentlyContinue
                    if ($user) {
                        $principalInfo.UserPrincipalName = $user.UserPrincipalName
                        $principalInfo.DisplayName = $user.DisplayName
                        $principalInfo.UserEnabled = $user.AccountEnabled
                        $principalInfo.OnPremisesSyncEnabled = $User.OnPremisesSyncEnabled
                        $principalInfo.PrincipalType = "User"
                    }
                }
                catch { }
                
                # Try as service principal if not user
                if ($principalInfo.PrincipalType -eq "Unknown") {
                    try {
                        $servicePrincipal = Get-MgServicePrincipal -ServicePrincipalId $assignment.PrincipalId -Property "AppId,DisplayName,AccountEnabled" -ErrorAction SilentlyContinue
                        if ($servicePrincipal) {
                            $principalInfo.UserPrincipalName = $servicePrincipal.AppId
                            $principalInfo.DisplayName = "$($servicePrincipal.DisplayName) (Application)"
                            $principalInfo.UserEnabled = $servicePrincipal.AccountEnabled                            
                            $principalInfo.PrincipalType = "ServicePrincipal"
                        }
                    }
                    catch { }
                }
                
                # Try as group if still unknown
                if ($principalInfo.PrincipalType -eq "Unknown") {
                    try {
                        $group = Get-MgGroup -GroupId $assignment.PrincipalId -Property "Mail,DisplayName" -ErrorAction SilentlyContinue
                        if ($group) {
                            $principalInfo.UserPrincipalName = $group.Mail
                            $principalInfo.DisplayName = "$($group.DisplayName) (Group)"
                            $principalInfo.PrincipalType = "Group"
                        }
                    }
                    catch { }
                }
                
                # Try as directory object if still unknown
                if ($principalInfo.PrincipalType -eq "Unknown") {
                    try {
                        $directoryObject = Get-MgDirectoryObject -DirectoryObjectId $assignment.PrincipalId -ErrorAction SilentlyContinue
                        if ($directoryObject) {
                            $principalInfo.DisplayName = "$($directoryObject.DisplayName) ($($directoryObject.'@odata.type'))"
                            $principalInfo.PrincipalType = $directoryObject.'@odata.type'
                        }
                    }
                    catch { }
                }
                
                # Determine role scope for enhanced deduplication
                $roleScope = if ($role.DisplayName -in $overarchingRoles) { "Overarching" } else { "Service-Specific" }
                
                $results += [PSCustomObject]@{
                    Service = "Power Platform"
                    UserPrincipalName = $principalInfo.UserPrincipalName
                    DisplayName = $principalInfo.DisplayName
                    UserId = $principalInfo.UserId
                    RoleName = $role.DisplayName
                    RoleDefinitionId = $assignment.RoleDefinitionId
                    RoleScope = $roleScope  # New property for enhanced deduplication
                    AssignmentType = $assignmentType
                    AssignedDateTime = $assignment.CreatedDateTime
                    UserEnabled = $principalInfo.UserEnabled
                    #LastSignIn = $principalInfo.LastSignIn
                    Scope = $assignment.DirectoryScopeId
                    AssignmentId = $assignment.Id
                    #AuthenticationType = "Certificate"
                    PrincipalType = $principalInfo.PrincipalType
                    OnPremisesSyncEnabled = $principalInfo.OnPremisesSyncEnabled
                    PIMStartDateTime = $assignment.ScheduleInfo.StartDateTime
                    PIMEndDateTime = $assignment.ScheduleInfo.Expiration.EndDateTime
                }
                
            }
            catch {
                Write-Verbose "Error processing Power Platform assignment: $($_.Exception.Message)"
                
                # Still add with limited info to avoid losing data
                $role = $roleDefinitions | Where-Object { $_.Id -eq $assignment.RoleDefinitionId }
                $results += [PSCustomObject]@{
                    Service = "Power Platform"
                    UserPrincipalName = "Error resolving principal"
                    DisplayName = "Principal ID: $($assignment.PrincipalId)"
                    UserId = $assignment.PrincipalId
                    RoleName = if ($role) { $role.DisplayName } else { "Unknown Role" }
                    RoleDefinitionId = $assignment.RoleDefinitionId
                    RoleScope = "Unknown"
                    AssignmentType = "Error"
                    AssignedDateTime = $assignment.CreatedDateTime
                    UserEnabled = $null
                    #LastSignIn = $null
                    Scope = $assignment.DirectoryScopeId
                    AssignmentId = $assignment.Id
                    #AuthenticationType = "Certificate"
                    PrincipalType = "Error"
                    OnPremisesSyncEnabled = $principalInfo.OnPremisesSyncEnabled
                    PIMStartDateTime = $null
                    PIMEndDateTime = $null
                }
            }
        } #>
        
        Write-Host "✓ Power Platform Azure AD role audit completed. Found $($results.Count) role assignments (including PIM)" -ForegroundColor Green
        
        # Provide feedback about role filtering
        if (-not $IncludeAzureADRoles) {
            Write-Host "  (Excluding overarching Azure AD roles - use -IncludeAzureADRoles to include)" -ForegroundColor Yellow
        }
        
        # Show detailed breakdown
        if ($results.Count -gt 0) {
            Write-Host ""
            Write-Host "Assignment breakdown:" -ForegroundColor Cyan
            
            # By principal type
            $typeSummary = $results | Group-Object PrincipalType
            Write-Host "Principal types:" -ForegroundColor Yellow
            foreach ($type in $typeSummary) {
                Write-Host "  $($type.Name): $($type.Count)" -ForegroundColor White
            }
            
            # By assignment type  
            $assignmentTypeSummary = $results | Group-Object AssignmentType
            Write-Host "Assignment types:" -ForegroundColor Yellow
            foreach ($type in $assignmentTypeSummary) {
                Write-Host "  $($type.Name): $($type.Count)" -ForegroundColor White
            }
            
            # By role scope
            $scopeSummary = $results | Group-Object RoleScope
            Write-Host "Role scope:" -ForegroundColor Yellow
            foreach ($scope in $scopeSummary) {
                Write-Host "  $($scope.Name): $($scope.Count)" -ForegroundColor White
            }
            
            # By role
            $roleSummary = $results | Group-Object RoleName
            Write-Host "Roles:" -ForegroundColor Yellow
            foreach ($role in $roleSummary) {
                Write-Host "  $($role.Name): $($role.Count)" -ForegroundColor White
            }
        }
        
    }
    catch {
        Write-Error "Error auditing Power Platform Azure AD roles: $($_.Exception.Message)"
        
        if ($_.Exception.Message -like "*certificate*") {
            Write-Host "Certificate authentication required for Power Platform role audit" -ForegroundColor Red
            Write-Host "Use: Set-M365AuditCertCredentials -TenantId <id> -ClientId <id> -CertificateThumbprint <thumbprint>" -ForegroundColor White
        }
        
        throw
    }
    
    return $results
}

