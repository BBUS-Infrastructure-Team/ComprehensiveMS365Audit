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
            
            Connect-MgGraph -TenantId $script:AppConfig.TenantId -ClientId $script:AppConfig.ClientId -CertificateThumbprint $script:AppConfig.CertificateThumbprint
            
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

# Fixed Teams and Defender functions with PIM support
# Replace the existing functions in 03-AzureAD-Functions.ps1

# Teams Roles (Azure AD roles) - Certificate Authentication + PIM Support
function Get-TeamsRoleAudit {
    param(
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
        
        # Teams-specific Azure AD roles
        $teamsRoles = @(
            "Teams Administrator",
            "Teams Communications Administrator",
            "Teams Communications Support Engineer", 
            "Teams Communications Support Specialist",
            "Teams Devices Administrator",
            "Teams Telephony Administrator"
        )
        
        Write-Host "Retrieving Teams-related Azure AD roles..." -ForegroundColor Cyan
        $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition | Where-Object { $_.DisplayName -in $teamsRoles }
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
                
                $results += [PSCustomObject]@{
                    Service = "Microsoft Teams"
                    UserPrincipalName = $principalInfo.UserPrincipalName
                    DisplayName = $principalInfo.DisplayName
                    UserId = $principalInfo.UserId
                    RoleName = $role.DisplayName
                    RoleDefinitionId = $assignment.RoleDefinitionId
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
        
        # Show breakdown
        if ($results.Count -gt 0) {
            $typeSummary = $results | Group-Object PrincipalType
            $assignmentTypeSummary = $results | Group-Object AssignmentType
            
            Write-Host "Principal types:" -ForegroundColor Cyan
            foreach ($type in $typeSummary) {
                Write-Host "  $($type.Name): $($type.Count)" -ForegroundColor White
            }
            
            Write-Host "Assignment types:" -ForegroundColor Cyan
            foreach ($type in $assignmentTypeSummary) {
                Write-Host "  $($type.Name): $($type.Count)" -ForegroundColor White
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

# Microsoft Defender Roles - Certificate Authentication + PIM Support
function Get-DefenderRoleAudit {
    param(
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
        
        # Defender-related Azure AD roles
        $defenderRoles = @(
            "Security Administrator",
            "Security Operator", 
            "Security Reader",
            "Global Administrator",
            "Cloud Application Administrator",
            "Application Administrator"
        )
        
        Write-Host "Retrieving Defender-related Azure AD roles..." -ForegroundColor Cyan
        $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition | Where-Object { $_.DisplayName -in $defenderRoles }
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
                
                $results += [PSCustomObject]@{
                    Service = "Microsoft Defender"
                    UserPrincipalName = $principalInfo.UserPrincipalName
                    DisplayName = $principalInfo.DisplayName
                    UserId = $principalInfo.UserId
                    RoleName = $role.DisplayName
                    RoleDefinitionId = $assignment.RoleDefinitionId
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
        
        # Show breakdown
        if ($results.Count -gt 0) {
            $typeSummary = $results | Group-Object PrincipalType
            $assignmentTypeSummary = $results | Group-Object AssignmentType
            
            Write-Host "Principal types:" -ForegroundColor Cyan
            foreach ($type in $typeSummary) {
                Write-Host "  $($type.Name): $($type.Count)" -ForegroundColor White
            }
            
            Write-Host "Assignment types:" -ForegroundColor Cyan
            foreach ($type in $assignmentTypeSummary) {
                Write-Host "  $($type.Name): $($type.Count)" -ForegroundColor White
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

# Fixed Power Platform Azure AD Roles function
# Handles Users, Groups, AND Service Principals (Application Registrations)
# Check for PIM assignments - this is likely where your 3 assignments are hiding

# Final Power Platform Azure AD Roles function
# Includes Regular, PIM Eligible, AND PIM Active assignments
# This will find your 3 assignments!

function Get-PowerPlatformAzureADRoleAudit {
    param(
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
        
        # Power Platform-related Azure AD roles
        $powerPlatformRoles = @(
            "Power Platform Administrator",
            "Dynamics 365 Administrator",
            "Power BI Administrator",
            "Power BI Service Administrator",  # Legacy name
            "CRM Service Administrator",       # Legacy name for Dynamics 365 Administrator
            "Dynamics 365 Service Administrator"
        )
        
        Write-Host "Retrieving Power Platform-related Azure AD roles..." -ForegroundColor Cyan
        $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition | Where-Object { $_.DisplayName -in $powerPlatformRoles }
        Write-Host "Found $($roleDefinitions.Count) Power Platform role definitions" -ForegroundColor Green
        
        # Initialize assignment collection
        $allAssignments = @()
        
        # 1. Get regular assignments (permanent)
        Write-Host "Checking regular assignments..." -ForegroundColor Cyan
        $regularAssignments = Get-MgRoleManagementDirectoryRoleAssignment | Where-Object { $_.RoleDefinitionId -in $roleDefinitions.Id }
        if ($regularAssignments) {
            $allAssignments += $regularAssignments
        }
        Write-Host "Found $($regularAssignments.Count) regular assignments" -ForegroundColor Gray
        
        # 2. Get PIM eligible assignments (require activation)
        Write-Host "Checking PIM eligible assignments..." -ForegroundColor Cyan
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
        
        # 3. Get PIM active assignments (time-limited, currently activated) - THIS IS WHERE YOUR 3 ARE!
        Write-Host "Checking PIM active assignments..." -ForegroundColor Cyan
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
        
        Write-Host "Total assignments to process: $($allAssignments.Count)" -ForegroundColor Green
        
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
                
                # Initialize principal information
                $principalInfo = @{
                    UserPrincipalName = "Unknown"
                    DisplayName = "Unknown Principal"
                    UserId = $assignment.PrincipalId
                    UserEnabled = $null
                    LastSignIn = $null
                    PrincipalType = "Unknown"
                }
                
                # Try to resolve principal - check multiple types
                # Method 1: Try as User first
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
                catch {
                    Write-Verbose "Not a user: $($assignment.PrincipalId)"
                }
                
                # Method 2: Try as Service Principal (Application Registration) if user lookup failed
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
                    catch {
                        Write-Verbose "Not a service principal: $($assignment.PrincipalId)"
                    }
                }
                
                # Method 3: Try as Group if still unknown
                if ($principalInfo.PrincipalType -eq "Unknown") {
                    try {
                        $group = Get-MgGroup -GroupId $assignment.PrincipalId -ErrorAction SilentlyContinue
                        if ($group) {
                            $principalInfo.UserPrincipalName = $group.Mail
                            $principalInfo.DisplayName = "$($group.DisplayName) (Group)"
                            $principalInfo.PrincipalType = "Group"
                        }
                    }
                    catch {
                        Write-Verbose "Not a group: $($assignment.PrincipalId)"
                    }
                }
                
                # Method 4: If still unknown, try generic directory object
                if ($principalInfo.PrincipalType -eq "Unknown") {
                    try {
                        $directoryObject = Get-MgDirectoryObject -DirectoryObjectId $assignment.PrincipalId -ErrorAction SilentlyContinue
                        if ($directoryObject) {
                            $principalInfo.DisplayName = "$($directoryObject.DisplayName) ($($directoryObject.'@odata.type'))"
                            $principalInfo.PrincipalType = $directoryObject.'@odata.type'
                        }
                    }
                    catch {
                        Write-Verbose "Could not resolve directory object: $($assignment.PrincipalId)"
                    }
                }
                
                # Create result - INCLUDE ALL PRINCIPAL TYPES AND ASSIGNMENT TYPES
                $results += [PSCustomObject]@{
                    Service = "Power Platform"
                    UserPrincipalName = $principalInfo.UserPrincipalName
                    DisplayName = $principalInfo.DisplayName
                    UserId = $principalInfo.UserId
                    RoleName = $role.DisplayName
                    RoleDefinitionId = $assignment.RoleDefinitionId
                    AssignmentType = $assignmentType
                    AssignedDateTime = $assignment.CreatedDateTime
                    UserEnabled = $principalInfo.UserEnabled
                    LastSignIn = $principalInfo.LastSignIn
                    Scope = $assignment.DirectoryScopeId
                    AssignmentId = $assignment.Id
                    AuthenticationType = "Certificate"
                    PrincipalType = $principalInfo.PrincipalType
                    # PIM-specific fields
                    PIMStartDateTime = $assignment.ScheduleInfo.StartDateTime
                    PIMEndDateTime = $assignment.ScheduleInfo.Expiration.EndDateTime
                }
                
            }
            catch {
                Write-Verbose "Error processing Power Platform assignment: $($_.Exception.Message)"
                # Don't skip - still add with limited info
                $role = $roleDefinitions | Where-Object { $_.Id -eq $assignment.RoleDefinitionId }
                $results += [PSCustomObject]@{
                    Service = "Power Platform"
                    UserPrincipalName = "Error resolving principal"
                    DisplayName = "Principal ID: $($assignment.PrincipalId)"
                    UserId = $assignment.PrincipalId
                    RoleName = $role.DisplayName
                    RoleDefinitionId = $assignment.RoleDefinitionId
                    AssignmentType = "Error"
                    AssignedDateTime = $assignment.CreatedDateTime
                    UserEnabled = $null
                    LastSignIn = $null
                    Scope = $assignment.DirectoryScopeId
                    AssignmentId = $assignment.Id
                    AuthenticationType = "Certificate"
                    PrincipalType = "Error"
                }
            }
        }
        
        Write-Host "✓ Power Platform Azure AD role audit completed. Found $($results.Count) role assignments" -ForegroundColor Green
        
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