 # 09-Intune-Functions.ps1
# Microsoft Intune/Endpoint Manager role audit functions
# Certificate-based authentication ONLY - No interactive authentication
function Get-IntuneRoleAudit {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$CertificateThumbprint,
        [switch]$IncludePIM,
        [switch]$IncludeAnalysis
    )
    
    $results = @()
    
    try {
        # Certificate authentication is required for this function
        if ($TenantId -and $ClientId -and $CertificateThumbprint) {
            Set-M365AuditCertCredentials -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
        }
        
        # Verify certificate authentication is configured
        if (-not $script:AppConfig.UseAppAuth -or $script:AppConfig.AuthType -ne "Certificate") {
            throw "Certificate authentication is required for Intune role audit. Use Set-M365AuditCertCredentials first."
        }
        
        # Write-Host "Using certificate-based authentication for Intune audit:" -ForegroundColor Cyan
        # Write-Host "  Tenant ID: $($script:AppConfig.TenantId)" -ForegroundColor Gray
        # Write-Host "  Client ID: $($script:AppConfig.ClientId)" -ForegroundColor Gray
        # Write-Host "  Certificate Thumbprint: $($script:AppConfig.CertificateThumbprint)" -ForegroundColor Gray
        
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
            Write-Host "  App Name: $($context.AppName)" -ForegroundColor Gray
            Write-Host "  Authentication Type: $($context.AuthType)" -ForegroundColor Gray
        }
        else {
            Write-Host "✓ Already connected to Microsoft Graph with app-only authentication" -ForegroundColor Green
        }
        
        # Verify Graph connection and required permissions
        try {
            $null = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/deviceManagement/roleDefinitions" -Method GET -ErrorAction Stop
            Write-Host "✓ Intune Graph API access verified" -ForegroundColor Green
        }
        catch {
            Write-Warning "Intune Graph API access failed: $($_.Exception.Message)"
            Write-Host "Required permissions:" -ForegroundColor Yellow
            Write-Host "• DeviceManagementRBAC.Read.All" -ForegroundColor White
            Write-Host "• DeviceManagementConfiguration.Read.All" -ForegroundColor White
            Write-Host "• Directory.Read.All" -ForegroundColor White
            Write-Host "• RoleEligibilitySchedule.Read.Directory (for PIM)" -ForegroundColor White
            Write-Host "• RoleAssignmentSchedule.Read.Directory (for PIM)" -ForegroundColor White
            throw "Required permissions not granted or certificate authentication failed"
        }
        
        # Get Intune-specific Azure AD roles first
        Write-Host "Retrieving Intune-related Azure AD roles..." -ForegroundColor Cyan
        
        $intuneAzureRoles = @(
            "Intune Service Administrator",
            "Device Managers", 
            "Device Users",
            "Device Administrators",
            "Endpoint Privilege Manager Administrator"
        )
        
        $intuneRoleDefinitions = @()
        
        try {
            $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition
            $intuneRoleDefinitions = $roleDefinitions | Where-Object { $_.DisplayName -in $intuneAzureRoles }
            $assignments = Get-MgRoleManagementDirectoryRoleAssignment | Where-Object { $_.RoleDefinitionId -in $intuneRoleDefinitions.Id }
            
            Write-Host "Found $($assignments.Count) active Azure AD Intune role assignments" -ForegroundColor Green
            
            foreach ($assignment in $assignments) {
                try {
                    $user = Get-MgUser -UserId $assignment.PrincipalId -ErrorAction SilentlyContinue
                    if (-not $user) { continue }
                    
                    $role = $intuneRoleDefinitions | Where-Object { $_.Id -eq $assignment.RoleDefinitionId }
                    
                    $results += [PSCustomObject]@{
                        Service = "Microsoft Intune"
                        UserPrincipalName = $user.UserPrincipalName
                        DisplayName = $user.DisplayName
                        UserId = $user.Id
                        RoleName = $role.DisplayName
                        RoleDefinitionId = $assignment.RoleDefinitionId
                        AssignmentType = "Azure AD Role"
                        AssignedDateTime = $assignment.CreatedDateTime
                        UserEnabled = $user.AccountEnabled
                        LastSignIn = $user.SignInActivity.LastSignInDateTime
                        Scope = $assignment.DirectoryScopeId
                        AssignmentId = $assignment.Id
                        RoleType = "AzureAD"
                        AuthenticationType = "Certificate"
                    }
                }
                catch {
                    Write-Verbose "Error processing Azure AD Intune assignment: $($_.Exception.Message)"
                }
            }
        }
        catch {
            Write-Warning "Error retrieving Intune Azure AD roles: $($_.Exception.Message)"
        }
        
        # Get PIM eligible assignments for Intune Azure AD roles
        Write-Host "Retrieving PIM eligible assignments for Intune roles..." -ForegroundColor Cyan
        try {
            $intuneRoleIds = $intuneRoleDefinitions.Id
            $eligibleAssignments = Get-MgRoleManagementDirectoryRoleEligibilitySchedule | Where-Object { 
                $_.RoleDefinitionId -in $intuneRoleIds 
            }
            
            Write-Host "Found $($eligibleAssignments.Count) PIM eligible assignments for Intune roles" -ForegroundColor Green
            
            foreach ($assignment in $eligibleAssignments) {
                try {
                    $user = Get-MgUser -UserId $assignment.PrincipalId -ErrorAction SilentlyContinue
                    if (-not $user) { continue }
                    
                    $role = $intuneRoleDefinitions | Where-Object { $_.Id -eq $assignment.RoleDefinitionId }
                    
                    $results += [PSCustomObject]@{
                        Service = "Microsoft Intune"
                        UserPrincipalName = $user.UserPrincipalName
                        DisplayName = $user.DisplayName
                        UserId = $user.Id
                        RoleName = $role.DisplayName
                        RoleDefinitionId = $assignment.RoleDefinitionId
                        AssignmentType = "Eligible (PIM)"
                        AssignedDateTime = $assignment.CreatedDateTime
                        UserEnabled = $user.AccountEnabled
                        LastSignIn = $user.SignInActivity.LastSignInDateTime
                        Scope = $assignment.DirectoryScopeId
                        AssignmentId = $assignment.Id
                        RoleType = "AzureAD"
                        PIMEndDateTime = $assignment.ScheduleInfo.Expiration.EndDateTime
                        PIMStartDateTime = $assignment.ScheduleInfo.Expiration.StartDateTime
                        PIMDuration = $assignment.ScheduleInfo.Expiration.Duration
                        AuthenticationType = "Certificate"
                    }
                }
                catch {
                    Write-Verbose "Error processing PIM Intune assignment: $($_.Exception.Message)"
                }
            }
        }
        catch {
            Write-Verbose "Error retrieving PIM eligible assignments for Intune: $($_.Exception.Message)"
            Write-Host "Note: PIM may require additional permissions or licensing" -ForegroundColor Yellow
        }

        # Get PIM active assignments for Intune roles
        Write-Host "Retrieving PIM active assignments for Intune roles..." -ForegroundColor Cyan
        try {
            $activeAssignments = Get-MgRoleManagementDirectoryRoleAssignmentSchedule | Where-Object { 
                $_.RoleDefinitionId -in $intuneRoleIds -and $_.AssignmentType -eq "Activated"
            }
            
            Write-Host "Found $($activeAssignments.Count) PIM active assignments for Intune roles" -ForegroundColor Green
            
            foreach ($assignment in $activeAssignments) {
                try {
                    $user = Get-MgUser -UserId $assignment.PrincipalId -ErrorAction SilentlyContinue
                    if (-not $user) { continue }
                    
                    $role = $intuneRoleDefinitions | Where-Object { $_.Id -eq $assignment.RoleDefinitionId }
                    
                    $results += [PSCustomObject]@{
                        Service = "Microsoft Intune"
                        UserPrincipalName = $user.UserPrincipalName
                        DisplayName = $user.DisplayName
                        UserId = $user.Id
                        RoleName = $role.DisplayName
                        RoleDefinitionId = $assignment.RoleDefinitionId
                        AssignmentType = "Active (PIM Activated)"
                        AssignedDateTime = $assignment.CreatedDateTime
                        UserEnabled = $user.AccountEnabled
                        LastSignIn = $user.SignInActivity.LastSignInDateTime
                        Scope = $assignment.DirectoryScopeId
                        AssignmentId = $assignment.Id
                        RoleType = "AzureAD"
                        PIMActivatedDateTime = $assignment.ActivatedDateTime
                        PIMEndDateTime = $assignment.ScheduleInfo.Expiration.EndDateTime
                        AuthenticationType = "Certificate"
                    }
                }
                catch {
                    Write-Verbose "Error processing PIM active Intune assignment: $($_.Exception.Message)"
                }
            }
        }
        catch {
            Write-Verbose "Error retrieving PIM active assignments for Intune: $($_.Exception.Message)"
        }
        
        # Get Intune RBAC role definitions
        Write-Host "Retrieving Intune RBAC role definitions..." -ForegroundColor Cyan
        try {
            $intuneRBACRoleDefinitions = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/deviceManagement/roleDefinitions" -Method GET
            Write-Host "Found $($intuneRBACRoleDefinitions.value.Count) Intune role definitions" -ForegroundColor Green
        }
        catch {
            Write-Warning "Could not retrieve Intune role definitions: $($_.Exception.Message)"
            throw "Certificate authentication may lack required permissions"
        }
        
        # Get Intune RBAC role assignments
        Write-Host "Retrieving Intune RBAC role assignments..." -ForegroundColor Cyan
        try {
            $intuneRoleAssignments = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/deviceManagement/roleAssignments" -Method GET
            Write-Host "Found $($intuneRoleAssignments.value.Count) Intune role assignments" -ForegroundColor Green
            
            foreach ($assignment in $intuneRoleAssignments.value) {
                try {
                    # Get role definition details
                    $roleDefinition = $intuneRBACRoleDefinitions.value | Where-Object { $_.id -eq $assignment.roleDefinition.id }
                    if (-not $roleDefinition) {
                        # Fetch individual role definition if not found in bulk list
                        try {
                            $roleDefResponse = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/deviceManagement/roleDefinitions/$($assignment.roleDefinition.id)" -Method GET
                            $roleDefinition = $roleDefResponse
                        }
                        catch {
                            Write-Verbose "Could not retrieve role definition for ID: $($assignment.roleDefinition.id)"
                            continue
                        }
                    }
                    
                    # Process each member in the assignment
                    foreach ($member in $assignment.members) {
                        try {
                            # Try to get user details
                            $user = $null
                            $userDisplayName = "Unknown"
                            $userPrincipalName = "Unknown"
                            $userEnabled = $null
                            $lastSignIn = $null
                            
                            # Member could be a user, group, or service principal
                            try {
                                $user = Get-MgUser -UserId $member -ErrorAction SilentlyContinue
                                if ($user) {
                                    $userDisplayName = $user.DisplayName
                                    $userPrincipalName = $user.UserPrincipalName
                                    $userEnabled = $user.AccountEnabled
                                    $lastSignIn = $user.SignInActivity.LastSignInDateTime
                                }
                                else {
                                    # Try as group
                                    $group = Get-MgGroup -GroupId $member -ErrorAction SilentlyContinue
                                    if ($group) {
                                        $userDisplayName = $group.DisplayName + " (Group)"
                                        $userPrincipalName = $group.Mail
                                    }
                                    else {
                                        # Try as service principal
                                        $sp = Get-MgServicePrincipal -ServicePrincipalId $member -ErrorAction SilentlyContinue
                                        if ($sp) {
                                            $userDisplayName = $sp.DisplayName + " (Service Principal)"
                                            $userPrincipalName = $sp.AppId
                                        }
                                    }
                                }
                            }
                            catch {
                                Write-Verbose "Could not resolve member: $member"
                                $userDisplayName = $member
                                $userPrincipalName = $member
                            }
                            
                            # Get scope information
                            $scopeInfo = "Organization"
                            $scopeDetails = @()
                            
                            if ($assignment.resourceScopes -and $assignment.resourceScopes.Count -gt 0) {
                                foreach ($scopeId in $assignment.resourceScopes) {
                                    try {
                                        # Try to get scope details - could be group or other resource
                                        if ($scopeId -eq "/") {
                                            $scopeDetails += "Root"
                                        }
                                        else {
                                            # Try as group first
                                            $scopeGroup = Get-MgGroup -GroupId $scopeId -ErrorAction SilentlyContinue
                                            if ($scopeGroup) {
                                                $scopeDetails += $scopeGroup.DisplayName
                                            }
                                            else {
                                                $scopeDetails += $scopeId
                                            }
                                        }
                                    }
                                    catch {
                                        $scopeDetails += $scopeId
                                    }
                                }
                                $scopeInfo = $scopeDetails -join ", "
                            }
                            
                            # Check if this is a time-bound assignment
                            $assignmentType = "Intune RBAC"
                            $pimEndDateTime = $null
                            $pimStartDateTime = $null
                            
                            if ($assignment.scheduleInfo -and $assignment.scheduleInfo.expiration.endDateTime) {
                                $assignmentType = "Time-bound RBAC"
                                $pimEndDateTime = $assignment.scheduleInfo.expiration.endDateTime
                                $pimStartDateTime = $assignment.scheduleInfo.startDateTime
                            }
                            
                            $results += [PSCustomObject]@{
                                Service = "Microsoft Intune"
                                UserPrincipalName = $userPrincipalName
                                DisplayName = $userDisplayName
                                UserId = $member
                                RoleName = $roleDefinition.displayName
                                RoleDefinitionId = $roleDefinition.id
                                AssignmentType = $assignmentType
                                AssignedDateTime = $assignment.createdDateTime
                                UserEnabled = $userEnabled
                                LastSignIn = $lastSignIn
                                Scope = $scopeInfo
                                AssignmentId = $assignment.id
                                RoleType = "IntuneRBAC"
                                RoleDescription = $roleDefinition.description
                                IsBuiltIn = $roleDefinition.isBuiltIn
                                PIMEndDateTime = $pimEndDateTime
                                PIMStartDateTime = $pimStartDateTime
                                AuthenticationType = "Certificate"
                            }
                        }
                        catch {
                            Write-Verbose "Error processing assignment member: $($_.Exception.Message)"
                        }
                    }
                }
                catch {
                    Write-Verbose "Error processing Intune role assignment: $($_.Exception.Message)"
                    continue
                }
            }
        }
        catch {
            Write-Warning "Could not retrieve Intune role assignments: $($_.Exception.Message)"
        }
        
        # Get Device Configuration Managers (policy creators/owners)
        Write-Host "Retrieving Device Configuration policy information..." -ForegroundColor Cyan
        try {
            $deviceConfigs = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/deviceManagement/deviceConfigurations" -Method GET
            
            foreach ($config in $deviceConfigs.value) {
                if ($config.createdDateTime) {
                    $results += [PSCustomObject]@{
                        Service = "Microsoft Intune"
                        UserPrincipalName = "System Generated"
                        DisplayName = "Device Configuration Policy"
                        UserId = $null
                        RoleName = "Device Configuration Policy Owner"
                        RoleDefinitionId = $null
                        AssignmentType = "Policy Owner"
                        AssignedDateTime = $config.createdDateTime
                        UserEnabled = $null
                        LastSignIn = $null
                        Scope = $config.displayName
                        AssignmentId = $config.id
                        RoleType = "PolicyOwner"
                        PolicyType = "DeviceConfiguration"
                        AuthenticationType = "Certificate"
                    }
                }
            }
        }
        catch {
            Write-Verbose "Could not retrieve device configuration policies: $($_.Exception.Message)"
        }
        
        # Get Compliance Policies
        Write-Host "Retrieving Device Compliance policy information..." -ForegroundColor Cyan
        try {
            $compliancePolicies = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies" -Method GET
            
            foreach ($policy in $compliancePolicies.value) {
                if ($policy.createdDateTime) {
                    $results += [PSCustomObject]@{
                        Service = "Microsoft Intune"
                        UserPrincipalName = "System Generated"
                        DisplayName = "Device Compliance Policy"
                        UserId = $null
                        RoleName = "Device Compliance Policy Owner"
                        RoleDefinitionId = $null
                        AssignmentType = "Policy Owner"
                        AssignedDateTime = $policy.createdDateTime
                        UserEnabled = $null
                        LastSignIn = $null
                        Scope = $policy.displayName
                        AssignmentId = $policy.id
                        RoleType = "PolicyOwner"
                        PolicyType = "DeviceCompliance"
                        AuthenticationType = "Certificate"
                    }
                }
            }
        }
        catch {
            Write-Verbose "Could not retrieve device compliance policies: $($_.Exception.Message)"
        }
        
        # Get App Protection Policies
        Write-Host "Retrieving App Protection policy information..." -ForegroundColor Cyan
        try {
            # iOS App Protection Policies
            $iosAppPolicies = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/deviceAppManagement/iosManagedAppProtections" -Method GET
            foreach ($policy in $iosAppPolicies.value) {
                if ($policy.createdDateTime) {
                    $results += [PSCustomObject]@{
                        Service = "Microsoft Intune"
                        UserPrincipalName = "System Generated"
                        DisplayName = "iOS App Protection Policy"
                        UserId = $null
                        RoleName = "App Protection Policy Owner"
                        RoleDefinitionId = $null
                        AssignmentType = "Policy Owner"
                        AssignedDateTime = $policy.createdDateTime
                        UserEnabled = $null
                        LastSignIn = $null
                        Scope = $policy.displayName
                        AssignmentId = $policy.id
                        RoleType = "PolicyOwner"
                        PolicyType = "iOSAppProtection"
                        AuthenticationType = "Certificate"
                    }
                }
            }
            
            # Android App Protection Policies
            $androidAppPolicies = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/deviceAppManagement/androidManagedAppProtections" -Method GET
            foreach ($policy in $androidAppPolicies.value) {
                if ($policy.createdDateTime) {
                    $results += [PSCustomObject]@{
                        Service = "Microsoft Intune"
                        UserPrincipalName = "System Generated"
                        DisplayName = "Android App Protection Policy"
                        UserId = $null
                        RoleName = "App Protection Policy Owner"
                        RoleDefinitionId = $null
                        AssignmentType = "Policy Owner"
                        AssignedDateTime = $policy.createdDateTime
                        UserEnabled = $null
                        LastSignIn = $null
                        Scope = $policy.displayName
                        AssignmentId = $policy.id
                        RoleType = "PolicyOwner"
                        PolicyType = "AndroidAppProtection"
                        AuthenticationType = "Certificate"
                    }
                }
            }
        }
        catch {
            Write-Verbose "Could not retrieve app protection policies: $($_.Exception.Message)"
        }
        
        # Get Device Enrollment configurations
        Write-Host "Retrieving Device Enrollment configurations..." -ForegroundColor Cyan
        try {
            $enrollmentConfigs = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/deviceManagement/deviceEnrollmentConfigurations" -Method GET
            
            foreach ($config in $enrollmentConfigs.value) {
                if ($config.createdDateTime) {
                    $results += [PSCustomObject]@{
                        Service = "Microsoft Intune"
                        UserPrincipalName = "System Generated"
                        DisplayName = "Device Enrollment Configuration"
                        UserId = $null
                        RoleName = "Device Enrollment Configuration Owner"
                        RoleDefinitionId = $null
                        AssignmentType = "Configuration Owner"
                        AssignedDateTime = $config.createdDateTime
                        UserEnabled = $null
                        LastSignIn = $null
                        Scope = $config.displayName
                        AssignmentId = $config.id
                        RoleType = "ConfigurationOwner"
                        ConfigurationType = $config.deviceEnrollmentConfigurationType
                        AuthenticationType = "Certificate"
                    }
                }
            }
        }
        catch {
            Write-Verbose "Could not retrieve device enrollment configurations: $($_.Exception.Message)"
        }
        
        # PIM Security Analysis for Intune
        Write-Host ""
        Write-Host "=== Intune PIM Security Analysis ===" -ForegroundColor Cyan

        $intuneResults = $results | Where-Object { $_.Service -eq "Microsoft Intune" }
        $intuneEligibleAssignments = $intuneResults | Where-Object { $_.AssignmentType -like "*Eligible*" }
        $intuneActiveAssignments = $intuneResults | Where-Object { $_.AssignmentType -eq "Azure AD Role" -or $_.AssignmentType -eq "Intune RBAC" }
        $intuneServiceAdmins = $intuneResults | Where-Object { $_.RoleName -eq "Intune Service Administrator" }
        $intuneTimeBoundAssignments = $intuneResults | Where-Object { $_.AssignmentType -eq "Time-bound RBAC" }

        Write-Host "Intune PIM Statistics:" -ForegroundColor White
        Write-Host "• Total Intune Assignments: $($intuneResults.Count)" -ForegroundColor Gray
        Write-Host "• PIM Eligible Assignments: $($intuneEligibleAssignments.Count)" -ForegroundColor Gray
        Write-Host "• Time-bound RBAC Assignments: $($intuneTimeBoundAssignments.Count)" -ForegroundColor Gray
        Write-Host "• Permanent Active Assignments: $($intuneActiveAssignments.Count)" -ForegroundColor Gray
        Write-Host "• Intune Service Administrators: $($intuneServiceAdmins.Count)" -ForegroundColor Gray
        
        if ($IncludeAnalysis) {

            # PIM recommendations
            if ($intuneActiveAssignments.Count -gt $intuneEligibleAssignments.Count -and $intuneEligibleAssignments.Count -eq 0) {
                Write-Host "⚠️ RECOMMENDATION: Consider implementing PIM for Intune role assignments" -ForegroundColor Yellow
                Write-Host "   Benefits: Just-in-time access, audit trail, approval workflows" -ForegroundColor White
            }

            if ($intuneServiceAdmins.Count -gt 2) {
                Write-Host "⚠️ RECOMMENDATION: Review Intune Service Administrator count ($($intuneServiceAdmins.Count))" -ForegroundColor Yellow
                Write-Host "   Consider using Intune RBAC roles for more granular permissions" -ForegroundColor White
            }

            $intuneRBACAssignments = $intuneResults | Where-Object { $_.RoleType -eq "IntuneRBAC" }
            $intuneAzureADAssignments = $intuneResults | Where-Object { $_.RoleType -eq "AzureAD" }

            if ($intuneAzureADAssignments.Count -gt $intuneRBACAssignments.Count) {
                Write-Host "⚠️ RECOMMENDATION: Leverage Intune RBAC over Azure AD roles where possible" -ForegroundColor Yellow
                Write-Host "   Intune RBAC provides more granular, scope-specific permissions" -ForegroundColor White
            }

            if ($intuneEligibleAssignments.Count -gt 0) {
                Write-Host "✓ PIM is being used for some Intune role assignments" -ForegroundColor Green
            }

            if ($intuneTimeBoundAssignments.Count -gt 0) {
                Write-Host "✓ Time-bound RBAC assignments detected - good security practice" -ForegroundColor Green
            }

            # Security validation
            Write-Host ""
            Write-Host "=== Security Validation ===" -ForegroundColor Green
            Write-Host "✓ Certificate-based authentication enforced" -ForegroundColor Green
            Write-Host "✓ No interactive authentication required" -ForegroundColor Green
            Write-Host "✓ Suitable for automated/unattended execution" -ForegroundColor Green
            Write-Host "✓ Enhanced security posture maintained" -ForegroundColor Green

            # Additional security recommendations
            Write-Host ""
            Write-Host "Security Recommendations:" -ForegroundColor Cyan
            Write-Host "• Implement regular access reviews for Intune administrative roles" -ForegroundColor White
            Write-Host "• Use Intune RBAC for granular, scoped permissions instead of broad Azure AD roles" -ForegroundColor White
            Write-Host "• Enable PIM for Intune Service Administrator and other high-privilege roles" -ForegroundColor White
            Write-Host "• Consider time-bound assignments for temporary access needs" -ForegroundColor White
            Write-Host "• Monitor device policy changes through audit logs" -ForegroundColor White
            Write-Host "• Implement approval workflows for sensitive Intune operations" -ForegroundColor White
            Write-Host "• Regularly rotate certificates (recommended 12-24 months)" -ForegroundColor White

            Write-Host ""
            Write-Host "✓ Intune audit completed with PIM analysis" -ForegroundColor Green
            Write-Host "Found $($results.Count) Intune-related role assignments and configurations" -ForegroundColor Cyan
            Write-Host "Authentication Method: Certificate-based (Secure)" -ForegroundColor Green
        }
    } catch {
        Write-Error "Error during Intune role audit: $($_.Exception.Message)"
        Write-Error "Stack trace: $($_.ScriptStackTrace)"
        
        # Provide troubleshooting guidance based on error type
        if ($_.Exception.Message -like "*certificate*") {
            Write-Host ""
            Write-Host "Certificate troubleshooting:" -ForegroundColor Yellow
            Write-Host "• Verify certificate exists in Windows Certificate Store" -ForegroundColor White
            Write-Host "• Check certificate expiration date" -ForegroundColor White
            Write-Host "• Ensure certificate is uploaded to Azure AD app registration" -ForegroundColor White
            Write-Host "• Run: Get-M365AuditCurrentConfig to verify setup" -ForegroundColor White
            Write-Host "• Run: New-M365AuditCertificate to create new certificate" -ForegroundColor White
        }
        elseif ($_.Exception.Message -like "*permission*" -or $_.Exception.Message -like "*access*") {
            Write-Host ""
            Write-Host "Permission troubleshooting:" -ForegroundColor Yellow
            Write-Host "• Run: Get-M365AuditRequiredPermissions" -ForegroundColor White
            Write-Host "• Verify admin consent has been granted in Azure AD" -ForegroundColor White
            Write-Host "• Check if app registration has required API permissions" -ForegroundColor White
            Write-Host "• For PIM: Ensure RoleEligibilitySchedule and RoleAssignmentSchedule permissions" -ForegroundColor White
            Write-Host "• For Intune: Ensure DeviceManagementRBAC.Read.All permission is granted" -ForegroundColor White
        }
        
        throw
    }

    
    return $results
}