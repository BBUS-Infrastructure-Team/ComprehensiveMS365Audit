# 05-Exchange-Functions.ps1 - REWRITTEN VERSION
# Updated Get-ExchangeRoleAudit function with proper Azure AD role filtering
function Get-ExchangeRoleAudit {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Organization,

        [string]$TenantId,
        
        [string]$ClientId,
        
        [string]$CertificateThumbprint,

        [switch]$Summary,

        [switch]$IncludeAzureADRoles  # New parameter to control inclusion of overarching roles
    )
    
    $results = @()
    
    try {
        # Set app credentials if provided, otherwise use existing script variables
        if ($TenantId -and $ClientId -and $CertificateThumbprint) {
            Set-M365AuditCertCredentials -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
        }
        
        # Verify certificate authentication is configured
        if (-not $script:AppConfig.UseAppAuth -or $script:AppConfig.AuthType -ne "Certificate") {
            Write-Warning "Certificate authentication is required for Exchange role audit"
            Write-Host "Please configure certificate authentication first:" -ForegroundColor Yellow
            Write-Host "• Run: Set-M365AuditCertCredentials -TenantId <id> -ClientId <id> -CertificateThumbprint <thumbprint>" -ForegroundColor White
            return $results
        }
        
        Write-Host "Using configured certificate credentials for Exchange audit:" -ForegroundColor Cyan
        Write-Host "  Tenant ID: $($script:AppConfig.TenantId)" -ForegroundColor Gray
        Write-Host "  Client ID: $($script:AppConfig.ClientId)" -ForegroundColor Gray
        Write-Host "  Certificate Thumbprint: $($script:AppConfig.CertificateThumbprint)" -ForegroundColor Gray
        
        # === STEP 1: Get Exchange-related Azure AD roles via Graph ===
        Write-Host "Retrieving Exchange-related Azure AD roles via Graph..." -ForegroundColor Cyan
        
        # Connect to Microsoft Graph if not already connected
        $context = Get-MgContext
        if (-not $context -or $context.AuthType -ne "AppOnly") {
            Write-Host "Connecting to Microsoft Graph for Exchange roles..." -ForegroundColor Yellow
            
            $null = Connect-MgGraph -TenantId $script:AppConfig.TenantId -ClientId $script:AppConfig.ClientId -CertificateThumbprint $script:AppConfig.CertificateThumbprint -NoWelcome
            
            # Verify app-only authentication
            $context = Get-MgContext
            if ($context.AuthType -ne "AppOnly") {
                throw "Expected app-only authentication but got: $($context.AuthType). Check certificate configuration."
            }
            
            Write-Host "✓ Connected with certificate authentication" -ForegroundColor Green
        }
        
        # Exchange-specific Azure AD roles (NOT overarching roles)
        $exchangeSpecificRoles = @(
            "Exchange Service Administrator"  # Legacy name for Exchange Administrator
        )
        
        # Overarching roles that should only appear in Azure AD audit
        $overarchingRoles = @(
            "Exchange Administrator",
            "Global Administrator",
            "Security Administrator",
            "Security Reader",
            "Cloud Application Administrator",
            "Application Administrator",
            "Privileged Authentication Administrator",
            "Privileged Role Administrator",
            "Compliance Administrator",
            "Compliance Data Administrator"
        )
        
        # Determine which roles to include based on parameter
        $rolesToInclude = if ($IncludeAzureADRoles) {
            $exchangeSpecificRoles + $overarchingRoles
        } else {
            $exchangeSpecificRoles
        }
        
        $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition | Where-Object { $_.DisplayName -in $rolesToInclude }
        Write-Host "Found $($roleDefinitions.Count) Exchange-related role definitions in Azure AD" -ForegroundColor Green
        
        # Get ALL assignment types (regular + PIM eligible + PIM active)
        $allAssignments = @()
        
        # Regular assignments
        $regularAssignments = Get-MgRoleManagementDirectoryRoleAssignment | Where-Object { $_.RoleDefinitionId -in $roleDefinitions.Id }
        if ($regularAssignments) { $allAssignments += $regularAssignments }
        Write-Host "Found $($regularAssignments.Count) regular Exchange role assignments" -ForegroundColor Gray
        
        # PIM eligible assignments
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
        Write-Host "Found $pimEligibleCount PIM eligible Exchange assignments" -ForegroundColor Gray
        
        # PIM active assignments
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
        Write-Host "Found $pimActiveCount PIM active Exchange assignments" -ForegroundColor $(if($pimActiveCount -gt 0) {"Green"} else {"Gray"})
        
        # Process Azure AD assignments
        foreach ($assignment in $allAssignments) {
            try {
                $role = $roleDefinitions | Where-Object { $_.Id -eq $assignment.RoleDefinitionId }
                
                # Determine assignment type
                $assignmentType = "Azure AD Role"
                if ($assignment.PSObject.TypeNames -contains "Microsoft.Graph.PowerShell.Models.MicrosoftGraphUnifiedRoleEligibilitySchedule") {
                    $assignmentType = "Eligible (PIM)"
                } 
                elseif ($assignment.PSObject.TypeNames -contains "Microsoft.Graph.PowerShell.Models.MicrosoftGraphUnifiedRoleAssignmentSchedule") {
                    $assignmentType = "Active (PIM)"
                }
                
                # Resolve principal (users AND groups for hybrid environments)
                $principalInfo = @{
                    UserPrincipalName = "Unknown"
                    DisplayName = "Unknown"
                    UserId = $assignment.PrincipalId
                    UserEnabled = $null
                    LastSignIn = $null
                    PrincipalType = "Unknown"
                    OnPremisesSyncEnabled = $null
                    GroupMemberCount = $null
                }
                
                # Try as user first
                try {
                    $user = Get-MgUser -UserId $assignment.PrincipalId -Property "UserPrincipalName,DisplayName,AccountEnabled,SignInActivity,OnPremisesSyncEnabled" -ErrorAction SilentlyContinue
                    if ($user) {
                        $principalInfo.UserPrincipalName = $user.UserPrincipalName
                        $principalInfo.DisplayName = $user.DisplayName
                        $principalInfo.UserEnabled = $user.AccountEnabled
                        $principalInfo.LastSignIn = $user.SignInActivity.LastSignInDateTime
                        $principalInfo.PrincipalType = "User"
                        $principalInfo.OnPremisesSyncEnabled = $user.OnPremisesSyncEnabled
                    }
                }
                catch { }

                # Try as group if not user (important for hybrid environments)
                if ($principalInfo.PrincipalType -eq "Unknown") {
                    try {
                        $group = Get-MgGroup -GroupId $assignment.PrincipalId -Property "Mail,DisplayName,OnPremisesSyncEnabled" -ErrorAction SilentlyContinue
                        if ($group) {
                            $principalInfo.UserPrincipalName = $group.Mail
                            $principalInfo.DisplayName = $group.DisplayName
                            $principalInfo.PrincipalType = "Group"
                            $principalInfo.OnPremisesSyncEnabled = $group.OnPremisesSyncEnabled
                        }
                    }
                    catch { }
                }
                
                # Try as service principal if still unknown
                if ($principalInfo.PrincipalType -eq "Unknown") {
                    try {
                        $servicePrincipal = Get-MgServicePrincipal -ServicePrincipalId $assignment.PrincipalId -ErrorAction SilentlyContinue
                        if ($servicePrincipal) {
                            $principalInfo.UserPrincipalName = $servicePrincipal.AppId
                            $principalInfo.DisplayName = "$($servicePrincipal.DisplayName) (Application)"
                            $principalInfo.PrincipalType = "ServicePrincipal"
                        }
                    }
                    catch { }
                }
                
                # Determine role scope for enhanced deduplication
                $roleScope = if ($role.DisplayName -in $overarchingRoles) { "Overarching" } else { "Service-Specific" }
                
                # Include both users and groups (critical for hybrid environments)
                if ($principalInfo.PrincipalType -eq "User" -or $principalInfo.PrincipalType -eq "Group" -or $principalInfo.PrincipalType -eq "ServicePrincipal") {
                    $results += [PSCustomObject]@{
                        Service = "Exchange Online"
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
                        RoleSource = "AzureAD"
                        PIMStartDateTime = $assignment.ScheduleInfo.StartDateTime
                        PIMEndDateTime = $assignment.ScheduleInfo.Expiration.EndDateTime
                        OnPremisesSyncEnabled = $principalInfo.OnPremisesSyncEnabled
                        GroupMemberCount = $principalInfo.GroupMemberCount
                        # Additional fields for consistency
                        RoleGroupDescription = $null
                        OrganizationalUnit = $null
                        ManagementScope = $null
                        RecipientType = $null
                    }
                }
            }
            catch {
                Write-Verbose "Error processing Azure AD Exchange assignment: $($_.Exception.Message)"
            }
        }
        
        # === STEP 2: Get Exchange role groups directly ===
        Write-Host "Retrieving Exchange role groups..." -ForegroundColor Cyan
        
        # Check if connected to Exchange Online
        
        $EXOSession = Get-ConnectionInformation | Where-Object { $_.connectionUrl -like "*outlook*" -and $_.State -eq 'Connected' }

        if (-not $EXOSession) {
            Write-Host "Connecting to Exchange Online with certificate authentication..." -ForegroundColor Yellow
            
            try {
                # Use script variables for connection
                if ($IsWindows) {
                    $null = Connect-ExchangeOnline `
                        -AppId $script:AppConfig.ClientId `
                        -CertificateThumbprint $script:AppConfig.CertificateThumbprint `
                        -Organization $Organization `
                        -ShowBanner:$false
                    Write-Host "✓ Connected to Exchange Online successfully" -ForegroundColor Green
                    Write-Host "Authentication Type: Certificate" -ForegroundColor Cyan
                } elseif ($IsLinux -or $IsMacOS) {
                    $null = Connect-ExchangeOnline `
                        -AppId $script:AppConfig.ClientId `
                        -Certificate $script:AppConfig.Certificate `
                        -Organization $Organization `
                        -ShowBanner:$false
                }
            }
            catch {
                Write-Error "Exchange Online certificate authentication failed: $($_.Exception.Message)"
                Write-Host "Troubleshooting steps:" -ForegroundColor Yellow
                Write-Host "• Ensure certificate is uploaded to Azure AD app registration" -ForegroundColor White
                Write-Host "• Verify app has Exchange.ManageAsApp permission" -ForegroundColor White
                Write-Host "• Check certificate expiration and validity" -ForegroundColor White
                Write-Host "• Verify Organization parameter matches your tenant" -ForegroundColor White
                return $results
            }
        }
        else {
            Write-Host "✓ Already connected to Exchange Online" -ForegroundColor Green
        }
        
        # Verify Exchange connection functionality
        try {
            $null = Get-OrganizationConfig -ErrorAction Stop | Select-Object -First 1
        }
        catch {
            Write-Warning "Exchange Online connection verification failed: $($_.Exception.Message)"
            Write-Host "Note: Some Exchange features may not be available" -ForegroundColor Yellow
            # Continue anyway as some commands might still work
        }
        
        # Focus on privileged Exchange role groups only
        $privilegedRoleGroups = @(
            "Organization Management",
            "Recipient Management", 
            "Server Management",
            "Exchange Administrators",
            "Security Administrator",
            "Security Reader",
            "Compliance Management",
            "Records Management",
            "Discovery Management",
            "Hygiene Management",
            "Help Desk"
        )
        
        # Get all role groups first, then filter
        Write-Host "Retrieving Exchange role groups..." -ForegroundColor Cyan
        $allRoleGroups = Get-RoleGroup -ErrorAction SilentlyContinue
        $targetRoleGroups = $allRoleGroups | Where-Object { $_.Name -in $privilegedRoleGroups }
        
        Write-Host "Found $($targetRoleGroups.Count) privileged role groups to process" -ForegroundColor Green
        
        foreach ($group in $targetRoleGroups) {
            try {
                $members = Get-RoleGroupMember -Identity $group.Identity -ErrorAction SilentlyContinue
                foreach ($member in $members) {
                    # Include users AND groups (critical for hybrid environments)
                    $isUser = $member.PrimarySmtpAddress -and $member.RecipientType -eq "UserMailbox"
                    $isGroup = $member.RecipientType -in @("MailUniversalSecurityGroup", "UniversalSecurityGroup", "MailUniversalDistributionGroup")
                    
                    if ($isUser -or $isGroup) {
                        $principalType = if ($isUser) { "User" } else { "Group" }
                        $userPrincipalName = if ($isUser) { $member.PrimarySmtpAddress } else { $member.Name }
                        
                        # Try to get additional user info from Graph for consistency
                        $userEnabled = $null
                        $lastSignIn = $null
                        $onPremisesSyncEnabled = $null

                        if ($isUser -and $member.ExternalDirectoryObjectId) {
                            try {
                                $graphUser = Get-MgUser -UserId $member.ExternalDirectoryObjectId -Property "AccountEnabled,SignInActivity,OnPremisesSyncEnabled,OnPremisesDistinguishedName" -ErrorAction SilentlyContinue
                                if ($graphUser) {
                                    $userEnabled = $graphUser.AccountEnabled
                                    $lastSignIn = $graphUser.SignInActivity.LastSignInDateTime
                                    $onPremisesSyncEnabled = $graphUser.OnPremisesSyncEnabled
                                    
                                    # Debug output to verify hybrid detection
                                    Write-Verbose "User $($member.PrimarySmtpAddress): OnPremSync = $onPremisesSyncEnabled"
                                }
                            }
                            catch {
                                Write-Verbose "Could not retrieve Graph data for user $($member.PrimarySmtpAddress): $($_.Exception.Message)"
                            }
                        }
                        
                        $results += [PSCustomObject]@{
                            Service = "Exchange Online"
                            UserPrincipalName = $userPrincipalName
                            DisplayName = $member.DisplayName
                            UserId = $member.ExternalDirectoryObjectId
                            RoleName = $group.Name
                            RoleDefinitionId = $group.Guid
                            RoleScope = "Service-Specific"  # Exchange role groups are service-specific
                            AssignmentType = "Role Group Member"
                            AssignedDateTime = $null
                            UserEnabled = $userEnabled
                            LastSignIn = $lastSignIn
                            Scope = "Organization"
                            AssignmentId = $group.Identity
                            AuthenticationType = "Certificate"
                            PrincipalType = $principalType
                            RoleSource = "ExchangeOnline"
                            # Exchange-specific fields for consistency
                            RoleGroupDescription = $group.Description
                            RecipientType = $member.RecipientType
                            OrganizationalUnit = $null
                            ManagementScope = $null
                            OnPremisesSyncEnabled = $onPremisesSyncEnabled
                            GroupMemberCount = if ($isGroup) { "See Group Details" } else { $null }
                            # Additional fields for consistency
                            PIMStartDateTime = $null
                            PIMEndDateTime = $null
                        }
                    }
                }
            }
            catch {
                Write-Warning "Could not get members for role group $($group.Name): $($_.Exception.Message)"
            }
        }
        
        # Get direct role assignments (privileged users with direct assignments)
        try {
            Write-Host "Retrieving direct role assignments..." -ForegroundColor Cyan
            $directAssignments = Get-ManagementRoleAssignment -ErrorAction SilentlyContinue | Where-Object { 
                $_.AssignmentMethod -eq "Direct" -and $_.Role -in $privilegedRoleGroups
            }
            
            foreach ($assignment in $directAssignments) {
                # Include both user and group direct assignments
                $isUserAssignment = $assignment.User -like "*@*" -and $assignment.User -notlike "*DiscoverySearchMailbox*"
                $isGroupAssignment = $assignment.User -notlike "*@*" -or $assignment.User -like "*Group*"
                
                if ($isUserAssignment -or $isGroupAssignment) {
                    $principalType = if ($isUserAssignment) { "User" } else { "Group" }
                    
                    $results += [PSCustomObject]@{
                        Service = "Exchange Online"
                        UserPrincipalName = $assignment.User
                        DisplayName = $assignment.User
                        UserId = $null
                        RoleName = $assignment.Role
                        RoleDefinitionId = $null
                        RoleScope = "Service-Specific"  # Direct assignments are service-specific
                        AssignmentType = "Direct Assignment"
                        AssignedDateTime = $assignment.WhenCreated
                        UserEnabled = $null
                        LastSignIn = $null
                        Scope = $assignment.RecipientOrganizationUnitScope
                        AssignmentId = $assignment.Identity
                        AuthenticationType = "Certificate"
                        PrincipalType = $principalType
                        RoleSource = "ExchangeOnline"
                        # Exchange-specific fields for consistency
                        RoleGroupDescription = $null
                        OrganizationalUnit = $assignment.RecipientOrganizationUnitScope
                        ManagementScope = $assignment.CustomRecipientWriteScope
                        RecipientType = "DirectAssignment"
                        OnPremisesSyncEnabled = $null
                        GroupMemberCount = $null
                        # Additional fields for consistency
                        PIMStartDateTime = $null
                        PIMEndDateTime = $null
                    }
                }
            }
            Write-Host "Found $($directAssignments.Count) direct role assignments" -ForegroundColor Green
        }
        catch {
            Write-Verbose "Could not retrieve direct role assignments: $($_.Exception.Message)"
        }
        
        Write-Host "✓ Exchange Online privileged role audit completed" -ForegroundColor Green
        
        # Provide feedback about role filtering
        if (-not $IncludeAzureADRoles) {
            Write-Host "  (Excluding overarching Azure AD roles - use -IncludeAzureADRoles to include)" -ForegroundColor Yellow
        }
        
        if ($Summary) {
            # Summary of results
            $azureADRoles = $results | Where-Object { $_.RoleSource -eq "AzureAD" }
            $exchangeRoleGroups = $results | Where-Object { $_.RoleSource -eq "ExchangeOnline" }
            $userAssignments = $results | Where-Object { $_.PrincipalType -eq "User" }
            $groupAssignments = $results | Where-Object { $_.PrincipalType -eq "Group" }
            $onPremSyncedObjects = $results | Where-Object { $_.OnPremisesSyncEnabled -eq $true }
            
            Write-Host ""
            Write-Host "=== EXCHANGE AUDIT SUMMARY ===" -ForegroundColor Cyan
            Write-Host "Azure AD Exchange Roles: $($azureADRoles.Count)" -ForegroundColor White
            Write-Host "Exchange Role Groups: $($exchangeRoleGroups.Count)" -ForegroundColor White
            Write-Host "User Assignments: $($userAssignments.Count)" -ForegroundColor White
            Write-Host "Group Assignments: $($groupAssignments.Count)" -ForegroundColor Green
            Write-Host "On-Premises Synced Objects: $($onPremSyncedObjects.Count)" -ForegroundColor Yellow
            Write-Host "Total Exchange Privileged Assignments: $($results.Count)" -ForegroundColor Green
            Write-Host "Authentication Method: Certificate-based" -ForegroundColor Cyan
            
            # Show breakdown by role type
            if ($results.Count -gt 0) {
                Write-Host ""
                Write-Host "Assignment Type Breakdown:" -ForegroundColor Cyan
                $assignmentTypes = $results | Group-Object AssignmentType
                foreach ($type in $assignmentTypes) {
                    Write-Host "  $($type.Name): $($type.Count)" -ForegroundColor White
                }
                
                Write-Host ""
                Write-Host "Role Scope Breakdown:" -ForegroundColor Cyan
                $scopeTypes = $results | Group-Object RoleScope
                foreach ($scope in $scopeTypes) {
                    Write-Host "  $($scope.Name): $($scope.Count)" -ForegroundColor White
                }
                
                # Show top roles and highlight group assignments
                Write-Host ""
                Write-Host "Top Exchange Roles:" -ForegroundColor Cyan
                $topRoles = $results | Group-Object RoleName | Sort-Object Count -Descending | Select-Object -First 5
                foreach ($role in $topRoles) {
                    $roleUsers = $role.Group | Where-Object { $_.PrincipalType -eq "User" }
                    $roleGroups = $role.Group | Where-Object { $_.PrincipalType -eq "Group" }
                    Write-Host "  $($role.Name): $($roleUsers.Count) users, $($roleGroups.Count) groups" -ForegroundColor White
                }
                
                # Highlight hybrid environment groups
                if ($Summary) {
                    if ($groupAssignments.Count -gt 0) {
                        Write-Host ""
                        Write-Host "=== GROUP ASSIGNMENTS (HYBRID ENVIRONMENT) ===" -ForegroundColor Yellow
                        $hybridGroups = $groupAssignments | Where-Object { $_.OnPremisesSyncEnabled -eq $true }
                        $cloudGroups = $groupAssignments | Where-Object { $_.OnPremisesSyncEnabled -ne $true }
                        
                        Write-Host "On-Premises Synced Groups: $($hybridGroups.Count)" -ForegroundColor Green
                        Write-Host "Cloud-Only Groups: $($cloudGroups.Count)" -ForegroundColor White
                        
                        # Show some example group assignments
                        Write-Host ""
                        Write-Host "Sample Group Assignments:" -ForegroundColor Cyan
                        $sampleGroups = $groupAssignments | Select-Object -First 5
                        foreach ($group in $sampleGroups) {
                            $syncStatus = if ($group.OnPremisesSyncEnabled) { "(On-Prem Synced)" } else { "(Cloud-Only)" }
                            Write-Host "  Group: $($group.DisplayName) $syncStatus" -ForegroundColor White
                            Write-Host "    Role: $($group.RoleName)" -ForegroundColor Gray
                            Write-Host "    Source: $($group.RoleSource)" -ForegroundColor Gray
                        }
                    }
                }
            }
        }
    }
    catch {
        Write-Error "Error auditing Exchange roles: $($_.Exception.Message)"
        Write-Host ""
        Write-Host "Troubleshooting Steps:" -ForegroundColor Yellow
        Write-Host "• Verify certificate is uploaded to Azure AD app registration" -ForegroundColor White
        Write-Host "• Ensure app has Exchange.ManageAsApp permission with admin consent" -ForegroundColor White
        Write-Host "• Check certificate expiration and validity" -ForegroundColor White
        Write-Host "• Verify Organization parameter matches your tenant" -ForegroundColor White
        Write-Host "• Run: Get-M365AuditCurrentConfig to verify setup" -ForegroundColor White
    }
    finally {
        # Clean up Exchange connection if we created it
        try {
            if (Get-PSSession | Where-Object { $_.ComputerName -like "*outlook.office365.com*" }) {
                $null = Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
                Write-Host "Disconnected from Exchange Online" -ForegroundColor Gray
            }
        }
        catch {
            # Ignore cleanup errors
        }
    }
    
    return $results
}