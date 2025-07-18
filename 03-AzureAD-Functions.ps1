# 02-AzureAD-Functions.ps1
# Azure AD/Entra ID role audit functions - Certificate Authentication ONLY

function Get-AzureADRoleAudit {
    param(
        [switch]$IncludePIM,
        [string]$TenantId,
        [string]$ClientId,
        [string]$CertificateThumbprint
    )
    
    $results = @()
    
    try {
        # Certificate authentication is required for this function
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
                throw $_
            }
            
            Write-Host "✓ Connected with certificate authentication" -ForegroundColor Green
            Write-Host "  App Name: $($context.AppName)" -ForegroundColor Gray
            #Write-Host "  Certificate Thumbprint: $($script:AppConfig.CertificateThumbprint)" -ForegroundColor Gray
        }
        
        # Verify required permissions
        try {
            [void](Get-MgUser -Top 1 -ErrorAction Stop)
        }
        catch {
            throw $_
        }
        
        try {
            [void](Get-MgRoleManagementDirectoryRoleDefinition -ErrorAction Stop)
            Write-Verbose "RoleManagement.Read.All permission verified"
        }
        catch {
            throw $_
        }
        
        # Get role definitions
        Write-Host "Retrieving Azure AD role definitions..." -ForegroundColor Cyan
        $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition
        $roleDefinitionHash = @{}
        foreach ($roleDef in $roleDefinitions) {
            $roleDefinitionHash[$roleDef.Id] = $roleDef.DisplayName
        }
        Write-Host "Found $($roleDefinitions.Count) role definitions" -ForegroundColor Green
        
        # Get active role assignments
        Write-Host "Retrieving active role assignments..." -ForegroundColor Cyan
        $activeAssignments = Get-MgRoleManagementDirectoryRoleAssignment
        Write-Host "Found $($activeAssignments.Count) active assignments" -ForegroundColor Green
        
        foreach ($assignment in $activeAssignments) {
            try {
                $user = Get-MgUser -UserId $assignment.PrincipalId -ErrorAction SilentlyContinue
                if (-not $user) { 
                    Write-Verbose "Skipping non-user principal: $($assignment.PrincipalId)"
                    continue 
                }
                
                $roleName = $roleDefinitionHash[$assignment.RoleDefinitionId]
                if (-not $roleName) { 
                    Write-Verbose "Role definition not found for ID: $($assignment.RoleDefinitionId)"
                    continue 
                }
                
                $results += [PSCustomObject]@{
                    Service = "Azure AD/Entra ID"
                    UserPrincipalName = $user.UserPrincipalName
                    DisplayName = $user.DisplayName
                    UserId = $user.Id
                    RoleName = $roleName
                    RoleDefinitionId = $assignment.RoleDefinitionId
                    AssignmentType = "Active"
                    AssignedDateTime = $assignment.CreatedDateTime
                    UserEnabled = $user.AccountEnabled
                    LastSignIn = $user.SignInActivity.LastSignInDateTime
                    Scope = $assignment.DirectoryScopeId
                    AssignmentId = $assignment.Id
                    AuthenticationType = "Certificate"
                }
            }
            catch {
                Write-Warning "Error processing assignment $($assignment.Id): $($_.Exception.Message)"
                continue
            }
        }
        
        # Include PIM eligible assignments if requested
        if ($IncludePIM) {
            try {
                Write-Host "Retrieving PIM eligible assignments..." -ForegroundColor Cyan
                $eligibleAssignments = Get-MgRoleManagementDirectoryRoleEligibilitySchedule
                Write-Host "Found $($eligibleAssignments.Count) eligible assignments" -ForegroundColor Green
                
                foreach ($assignment in $eligibleAssignments) {
                    try {
                        $user = Get-MgUser -UserId $assignment.PrincipalId -ErrorAction SilentlyContinue
                        if (-not $user) { 
                            Write-Verbose "Skipping non-user principal in PIM: $($assignment.PrincipalId)"
                            continue 
                        }
                        
                        $roleName = $roleDefinitionHash[$assignment.RoleDefinitionId]
                        if (-not $roleName) { 
                            Write-Verbose "PIM role definition not found for ID: $($assignment.RoleDefinitionId)"
                            continue 
                        }
                        
                        $results += [PSCustomObject]@{
                            Service = "Azure AD/Entra ID"
                            UserPrincipalName = $user.UserPrincipalName
                            DisplayName = $user.DisplayName
                            UserId = $user.Id
                            RoleName = $roleName
                            RoleDefinitionId = $assignment.RoleDefinitionId
                            AssignmentType = "Eligible (PIM)"
                            AssignedDateTime = $assignment.CreatedDateTime
                            UserEnabled = $user.AccountEnabled
                            LastSignIn = $user.SignInActivity.LastSignInDateTime
                            Scope = $assignment.DirectoryScopeId
                            AssignmentId = $assignment.Id
                            AuthenticationType = "Certificate"
                            PIMEndDateTime = $assignment.ScheduleInfo.Expiration.EndDateTime
                            PIMStartDateTime = $assignment.ScheduleInfo.Expiration.StartDateTime
                        }
                    }
                    catch {
                        Write-Warning "Error processing PIM assignment $($assignment.Id): $($_.Exception.Message)"
                        continue
                    }
                }
            }
            catch {
                Write-Warning "Error retrieving PIM eligible assignments: $($_.Exception.Message)"
                Write-Host "Note: PIM may require additional permissions or licensing" -ForegroundColor Yellow
            }
        }
        
        Write-Host "✓ Azure AD role audit completed. Found $($results.Count) role assignments" -ForegroundColor Green
    }
    catch {
        Write-Error "Error in Azure AD role audit: $($_.Exception.Message)"
        
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

# Enhanced Teams Role Audit Function with Azure AD Role Filtering
# Add to 03-AzureAD-Functions.ps1

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
            
            Connect-MgGraph -TenantId $script:AppConfig.TenantId -ClientId $script:AppConfig.ClientId -CertificateThumbprint $script:AppConfig.CertificateThumbprint -NoWelcome
            
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
        
        Write-Host "Retrieving Teams-related Azure AD roles..." -ForegroundColor Cyan
        $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition | Where-Object { $_.DisplayName -in $rolesToInclude }
        Write-Host "Found $($roleDefinitions.Count) Teams role definitions" -ForegroundColor Green
        
        # Get ALL assignment types (regular + PIM eligible + PIM active)
        $allAssignments = @()
        
        # 1. Regular assignments
        Write-Host "Checking regular Teams assignments..." -ForegroundColor Cyan
        $regularAssignments = Get-MgRoleManagementDirectoryRoleAssignment | Where-Object { $_.RoleDefinitionId -in $roleDefinitions.Id }
        if ($regularAssignments) { $allAssignments += $regularAssignments }
        Write-Host "Found $($regularAssignments.Count) regular assignments" -ForegroundColor Gray
        
        # 2. PIM eligible assignments
        Write-Host "Checking PIM eligible Teams assignments..." -ForegroundColor Cyan
        $pimEligibleCount = 0
        try {
            foreach ($roleId in $roleDefinitions.Id) {
                $pimEligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -Filter "roleDefinitionId eq '$roleId'" -ErrorAction SilentlyContinue
                if ($pimEligible) {
                    $allAssignments += $pimEligible
                    $pimEligibleCount += $pimEligible.Count
                }
            }
        }
        catch {
            Write-Verbose "Could not retrieve PIM eligible assignments: $($_.Exception.Message)"
        }
        Write-Host "Found $pimEligibleCount PIM eligible assignments" -ForegroundColor Gray
        
        # 3. PIM active assignments
        Write-Host "Checking PIM active Teams assignments..." -ForegroundColor Cyan
        $pimActiveCount = 0
        try {
            foreach ($roleId in $roleDefinitions.Id) {
                $pimActive = Get-MgRoleManagementDirectoryRoleAssignmentSchedule -Filter "roleDefinitionId eq '$roleId'" -ErrorAction SilentlyContinue
                if ($pimActive) {
                    $allAssignments += $pimActive
                    $pimActiveCount += $pimActive.Count
                }
            }
        }
        catch {
            Write-Verbose "Could not retrieve PIM active assignments: $($_.Exception.Message)"
        }
        Write-Host "Found $pimActiveCount PIM active assignments" -ForegroundColor $(if($pimActiveCount -gt 0) {"Green"} else {"Gray"})
        
        Write-Host "Total Teams assignments to process: $($allAssignments.Count)" -ForegroundColor Green
        
        # Process all assignments
        foreach ($assignment in $allAssignments) {
            try {
                $role = $roleDefinitions | Where-Object { $_.Id -eq $assignment.RoleDefinitionId }
                
                # Determine assignment type
                $assignmentType = "Active"
                if ($assignment.PSObject.TypeNames -contains "Microsoft.Graph.PowerShell.Models.MicrosoftGraphUnifiedRoleEligibilitySchedule") {
                    $assignmentType = "Eligible (PIM)"
                } 
                elseif ($assignment.PSObject.TypeNames -contains "Microsoft.Graph.PowerShell.Models.MicrosoftGraphUnifiedRoleAssignmentSchedule") {
                    $assignmentType = "Active (PIM)"
                }
                
                # Resolve principal (users, groups, service principals)
                $principalInfo = @{
                    UserPrincipalName = "Unknown"
                    DisplayName = "Unknown"
                    UserId = $assignment.PrincipalId
                    UserEnabled = $null
                    LastSignIn = $null
                    PrincipalType = "Unknown"
                }
                
                # Try as user
                try {
                    $user = Get-MgUser -UserId $assignment.PrincipalId -ErrorAction SilentlyContinue
                    if ($user) {
                        $principalInfo.UserPrincipalName = $user.UserPrincipalName
                        $principalInfo.DisplayName = $user.DisplayName
                        $principalInfo.UserEnabled = $user.AccountEnabled
                        $principalInfo.LastSignIn = $user.SignInActivity.LastSignInDateTime
                        $principalInfo.PrincipalType = "User"
                    }
                }
                catch { }
                
                # Try as service principal if not user
                if ($principalInfo.PrincipalType -eq "Unknown") {
                    try {
                        $app = Get-MgServicePrincipal -ServicePrincipalId $assignment.PrincipalId -ErrorAction SilentlyContinue
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
                        $group = Get-MgGroup -GroupId $assignment.PrincipalId -ErrorAction SilentlyContinue
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
                    LastSignIn = $principalInfo.LastSignIn
                    Scope = $assignment.DirectoryScopeId
                    AssignmentId = $assignment.Id
                    AuthenticationType = "Certificate"
                    PrincipalType = $principalInfo.PrincipalType
                    PIMStartDateTime = $assignment.ScheduleInfo.StartDateTime
                    PIMEndDateTime = $assignment.ScheduleInfo.Expiration.EndDateTime
                }
                
            }
            catch {
                Write-Verbose "Error processing Teams assignment: $($_.Exception.Message)"
            }
        }
        
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
        
        Write-Host "Retrieving Defender-related Azure AD roles..." -ForegroundColor Cyan
        $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition | Where-Object { $_.DisplayName -in $rolesToInclude }
        Write-Host "Found $($roleDefinitions.Count) Defender role definitions" -ForegroundColor Green
        
        # Get ALL assignment types (regular + PIM eligible + PIM active)
        $allAssignments = @()
        
        # 1. Regular assignments
        Write-Host "Checking regular Defender assignments..." -ForegroundColor Cyan
        $regularAssignments = Get-MgRoleManagementDirectoryRoleAssignment | Where-Object { $_.RoleDefinitionId -in $roleDefinitions.Id }
        if ($regularAssignments) { $allAssignments += $regularAssignments }
        Write-Host "Found $($regularAssignments.Count) regular assignments" -ForegroundColor Gray
        
        # 2. PIM eligible assignments
        Write-Host "Checking PIM eligible Defender assignments..." -ForegroundColor Cyan
        $pimEligibleCount = 0
        try {
            foreach ($roleId in $roleDefinitions.Id) {
                $pimEligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -Filter "roleDefinitionId eq '$roleId'" -ErrorAction SilentlyContinue
                if ($pimEligible) {
                    $allAssignments += $pimEligible
                    $pimEligibleCount += $pimEligible.Count
                }
            }
        }
        catch {
            Write-Verbose "Could not retrieve PIM eligible assignments: $($_.Exception.Message)"
        }
        Write-Host "Found $pimEligibleCount PIM eligible assignments" -ForegroundColor Gray
        
        # 3. PIM active assignments
        Write-Host "Checking PIM active Defender assignments..." -ForegroundColor Cyan
        $pimActiveCount = 0
        try {
            foreach ($roleId in $roleDefinitions.Id) {
                $pimActive = Get-MgRoleManagementDirectoryRoleAssignmentSchedule -Filter "roleDefinitionId eq '$roleId'" -ErrorAction SilentlyContinue
                if ($pimActive) {
                    $allAssignments += $pimActive
                    $pimActiveCount += $pimActive.Count
                }
            }
        }
        catch {
            Write-Verbose "Could not retrieve PIM active assignments: $($_.Exception.Message)"
        }
        Write-Host "Found $pimActiveCount PIM active assignments" -ForegroundColor $(if($pimActiveCount -gt 0) {"Green"} else {"Gray"})
        
        Write-Host "Total Defender assignments to process: $($allAssignments.Count)" -ForegroundColor Green
        
        # Process all assignments
        foreach ($assignment in $allAssignments) {
            try {
                $role = $roleDefinitions | Where-Object { $_.Id -eq $assignment.RoleDefinitionId }
                
                # Determine assignment type
                $assignmentType = "Active"
                if ($assignment.PSObject.TypeNames -contains "Microsoft.Graph.PowerShell.Models.MicrosoftGraphUnifiedRoleEligibilitySchedule") {
                    $assignmentType = "Eligible (PIM)"
                } 
                elseif ($assignment.PSObject.TypeNames -contains "Microsoft.Graph.PowerShell.Models.MicrosoftGraphUnifiedRoleAssignmentSchedule") {
                    $assignmentType = "Active (PIM)"
                }
                
                # Resolve principal (users, groups, service principals)
                $principalInfo = @{
                    UserPrincipalName = "Unknown"
                    DisplayName = "Unknown"
                    UserId = $assignment.PrincipalId
                    UserEnabled = $null
                    LastSignIn = $null
                    PrincipalType = "Unknown"
                }
                
                # Try as user
                try {
                    $user = Get-MgUser -UserId $assignment.PrincipalId -ErrorAction SilentlyContinue
                    if ($user) {
                        $principalInfo.UserPrincipalName = $user.UserPrincipalName
                        $principalInfo.DisplayName = $user.DisplayName
                        $principalInfo.UserEnabled = $user.AccountEnabled
                        $principalInfo.LastSignIn = $user.SignInActivity.LastSignInDateTime
                        $principalInfo.PrincipalType = "User"
                    }
                }
                catch { }
                
                # Try as service principal if not user
                if ($principalInfo.PrincipalType -eq "Unknown") {
                    try {
                        $app = Get-MgServicePrincipal -ServicePrincipalId $assignment.PrincipalId -ErrorAction SilentlyContinue
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
                        $group = Get-MgGroup -GroupId $assignment.PrincipalId -ErrorAction SilentlyContinue
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
                    LastSignIn = $principalInfo.LastSignIn
                    Scope = $assignment.DirectoryScopeId
                    AssignmentId = $assignment.Id
                    AuthenticationType = "Certificate"
                    PrincipalType = $principalInfo.PrincipalType
                    PIMStartDateTime = $assignment.ScheduleInfo.StartDateTime
                    PIMEndDateTime = $assignment.ScheduleInfo.Expiration.EndDateTime
                }
                
            }
            catch {
                Write-Verbose "Error processing Defender assignment: $($_.Exception.Message)"
            }
        }
        
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
            "Dynamics 365 Administrator",
            "Power BI Administrator",
            "Power BI Service Administrator",  # Legacy name
            "CRM Service Administrator"       # Legacy name for Dynamics 365 Administrator
        )
        
        # Overarching roles that should only appear in Azure AD audit
        $overarchingRoles = @(
            "Global Administrator",
            "Security Administrator",
            "Security Reader",
            "Cloud Application Administrator",
            "Application Administrator",
            "Privileged Authentication Administrator",
            "Privileged Role Administrator",
            "Power Platform Administrator"  # This is overarching - covers all Power Platform services
        )
        
        # Determine which roles to include based on parameter
        $rolesToInclude = if ($IncludeAzureADRoles) {
            $powerPlatformSpecificRoles + $overarchingRoles
        } else {
            $powerPlatformSpecificRoles
        }
        
        Write-Host "Retrieving Power Platform-related Azure AD roles..." -ForegroundColor Cyan
        $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition | Where-Object { $_.DisplayName -in $rolesToInclude }
        Write-Host "Found $($roleDefinitions.Count) Power Platform role definitions" -ForegroundColor Green
        
        # Get ALL assignment types (regular + PIM eligible + PIM active)
        $allAssignments = @()
        
        # 1. Regular assignments (permanent)
        Write-Host "Checking regular Power Platform assignments..." -ForegroundColor Cyan
        $regularAssignments = Get-MgRoleManagementDirectoryRoleAssignment | Where-Object { $_.RoleDefinitionId -in $roleDefinitions.Id }
        if ($regularAssignments) {
            $allAssignments += $regularAssignments
        }
        Write-Host "Found $($regularAssignments.Count) regular assignments" -ForegroundColor Gray
        
        # 2. PIM eligible assignments (require activation)
        Write-Host "Checking PIM eligible Power Platform assignments..." -ForegroundColor Cyan
        $pimEligibleCount = 0
        try {
            foreach ($roleId in $roleDefinitions.Id) {
                $pimEligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -Filter "roleDefinitionId eq '$roleId'" -ErrorAction SilentlyContinue
                if ($pimEligible) {
                    $allAssignments += $pimEligible
                    $pimEligibleCount += $pimEligible.Count
                }
            }
        }
        catch {
            Write-Verbose "Could not retrieve PIM eligible assignments: $($_.Exception.Message)"
        }
        Write-Host "Found $pimEligibleCount PIM eligible assignments" -ForegroundColor Gray
        
        # 3. PIM active assignments (time-limited, currently activated)
        Write-Host "Checking PIM active Power Platform assignments..." -ForegroundColor Cyan
        $pimActiveCount = 0
        try {
            foreach ($roleId in $roleDefinitions.Id) {
                $pimActive = Get-MgRoleManagementDirectoryRoleAssignmentSchedule -Filter "roleDefinitionId eq '$roleId'" -ErrorAction SilentlyContinue
                if ($pimActive) {
                    $allAssignments += $pimActive
                    $pimActiveCount += $pimActive.Count
                }
            }
        }
        catch {
            Write-Verbose "Could not retrieve PIM active assignments: $($_.Exception.Message)"
        }
        Write-Host "Found $pimActiveCount PIM active assignments" -ForegroundColor $(if($pimActiveCount -gt 0) {"Green"} else {"Gray"})
        
        Write-Host "Total Power Platform assignments to process: $($allAssignments.Count)" -ForegroundColor Green
        
        # Process all assignments (regular + PIM eligible + PIM active)
        foreach ($assignment in $allAssignments) {
            try {
                $role = $roleDefinitions | Where-Object { $_.Id -eq $assignment.RoleDefinitionId }
                
                # Determine assignment type based on object type
                $assignmentType = "Active"
                if ($assignment.PSObject.TypeNames -contains "Microsoft.Graph.PowerShell.Models.MicrosoftGraphUnifiedRoleEligibilitySchedule") {
                    $assignmentType = "Eligible (PIM)"
                } 
                elseif ($assignment.PSObject.TypeNames -contains "Microsoft.Graph.PowerShell.Models.MicrosoftGraphUnifiedRoleAssignmentSchedule") {
                    $assignmentType = "Active (PIM)"
                }
                
                # Resolve principal (users, groups, service principals)
                $principalInfo = @{
                    UserPrincipalName = "Unknown"
                    DisplayName = "Unknown"
                    UserId = $assignment.PrincipalId
                    UserEnabled = $null
                    LastSignIn = $null
                    PrincipalType = "Unknown"
                }
                
                # Try as user first
                try {
                    $user = Get-MgUser -UserId $assignment.PrincipalId -ErrorAction SilentlyContinue
                    if ($user) {
                        $principalInfo.UserPrincipalName = $user.UserPrincipalName
                        $principalInfo.DisplayName = $user.DisplayName
                        $principalInfo.UserEnabled = $user.AccountEnabled
                        $principalInfo.LastSignIn = $user.SignInActivity.LastSignInDateTime
                        $principalInfo.PrincipalType = "User"
                    }
                }
                catch { }
                
                # Try as service principal if not user
                if ($principalInfo.PrincipalType -eq "Unknown") {
                    try {
                        $servicePrincipal = Get-MgServicePrincipal -ServicePrincipalId $assignment.PrincipalId -ErrorAction SilentlyContinue
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
                        $group = Get-MgGroup -GroupId $assignment.PrincipalId -ErrorAction SilentlyContinue
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
                    LastSignIn = $principalInfo.LastSignIn
                    Scope = $assignment.DirectoryScopeId
                    AssignmentId = $assignment.Id
                    AuthenticationType = "Certificate"
                    PrincipalType = $principalInfo.PrincipalType
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
                    LastSignIn = $null
                    Scope = $assignment.DirectoryScopeId
                    AssignmentId = $assignment.Id
                    AuthenticationType = "Certificate"
                    PrincipalType = "Error"
                    PIMStartDateTime = $null
                    PIMEndDateTime = $null
                }
            }
        }
        
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