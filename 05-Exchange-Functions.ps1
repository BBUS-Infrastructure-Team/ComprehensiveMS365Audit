 # 05-Exchange-Functions.ps1 - SIMPLIFIED VERSION
# Exchange Online privileged role audit - Certificate Authentication Only
# Focused on privileged access roles only, no mail-enabled security groups

function Get-ExchangeRoleAudit {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Organization,

        [string]$TenantId,
        
        [string]$ClientId,
        
        [string]$CertificateThumbprint
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
        
        # === STEP 1: Get Exchange-related Azure AD roles via Graph (no conflicts) ===
        Write-Host "Retrieving Exchange-related Azure AD roles via Graph..." -ForegroundColor Cyan
        
        # Connect to Microsoft Graph if not already connected
        $context = Get-MgContext
        if (-not $context -or $context.AuthType -ne "AppOnly") {
            Write-Host "Connecting to Microsoft Graph for Exchange roles..." -ForegroundColor Yellow
            
            Connect-MgGraph -TenantId $script:AppConfig.TenantId -ClientId $script:AppConfig.ClientId -CertificateThumbprint $script:AppConfig.CertificateThumbprint -NoWelcome
            
            # Verify app-only authentication
            $context = Get-MgContext
            if ($context.AuthType -ne "AppOnly") {
                throw "Expected app-only authentication but got: $($context.AuthType). Check certificate configuration."
            }
            
            Write-Host "✓ Connected with certificate authentication" -ForegroundColor Green
        }
        
        # Exchange privileged roles in Azure AD
        $exchangePrivilegedRoles = @(
            "Exchange Administrator",
            "Exchange Recipient Administrator", 
            "Global Administrator"  # Include Global Admin as it has Exchange rights
        )
        
        $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition | Where-Object { $_.DisplayName -in $exchangePrivilegedRoles }
        Write-Host "Found $($roleDefinitions.Count) Exchange privileged role definitions in Azure AD" -ForegroundColor Green
        
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
                    $user = Get-MgUser -UserId $assignment.PrincipalId -ErrorAction SilentlyContinue
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
                        $group = Get-MgGroup -GroupId $assignment.PrincipalId -ErrorAction SilentlyContinue
                        if ($group) {
                            $principalInfo.UserPrincipalName = $group.Mail
                            $principalInfo.DisplayName = $group.DisplayName
                            $principalInfo.PrincipalType = "Group"
                            $principalInfo.OnPremisesSyncEnabled = $group.OnPremisesSyncEnabled
                            
                            # Get group member count for context
                            try {
                                $members = Get-MgGroupMember -GroupId $group.Id -All -ErrorAction SilentlyContinue
                                $principalInfo.GroupMemberCount = $members.Count
                            }
                            catch {
                                $principalInfo.GroupMemberCount = "Unknown"
                            }
                        }
                    }
                    catch { }
                }
                
                # Include both users and groups (critical for hybrid environments)
                if ($principalInfo.PrincipalType -eq "User" -or $principalInfo.PrincipalType -eq "Group") {
                    $results += [PSCustomObject]@{
                        Service = "Exchange Online"
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
                        RoleSource = "AzureAD"
                        PIMStartDateTime = $assignment.ScheduleInfo.StartDateTime
                        PIMEndDateTime = $assignment.ScheduleInfo.Expiration.EndDateTime
                        OnPremisesSyncEnabled = $principalInfo.OnPremisesSyncEnabled
                        GroupMemberCount = $principalInfo.GroupMemberCount
                        # Additional fields for consistency with existing data
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
        
        # === STEP 2: Get Exchange role groups via isolated runspace (avoids assembly conflicts) ===
        Write-Host "Retrieving Exchange role groups via isolated runspace..." -ForegroundColor Cyan
        
        # Temporarily disconnect from Graph to avoid assembly conflicts
        try {
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            Write-Host "Temporarily disconnected from Graph" -ForegroundColor Gray
        }
        catch {
            Write-Verbose "Graph disconnect not needed: $($_.Exception.Message)"
        }
        
        # Execute Exchange commands in separate runspace
        $exchangeScriptBlock = {
            param($TenantId, $ClientId, $CertificateThumbprint, $Organization)
            
            try {
                # Import Exchange module in isolated runspace
                Import-Module ExchangeOnlineManagement -Force -DisableNameChecking
                
                # Connect to Exchange Online
                Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -Organization $Organization -ShowBanner:$false
                
                $exchangeResults = @()
                
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
                $allRoleGroups = Get-RoleGroup -ErrorAction Stop
                $targetRoleGroups = $allRoleGroups | Where-Object { $_.Name -in $privilegedRoleGroups }
                
                Write-Output "Found $($targetRoleGroups.Count) privileged role groups to process"
                
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
                                        # Note: This will be called from the runspace, so Graph connection may not be available
                                        # We'll populate these fields in post-processing if needed
                                    }
                                    catch {
                                        # Expected in runspace context
                                    }
                                }
                                
                                $exchangeResults += [PSCustomObject]@{
                                    Service = "Exchange Online"
                                    UserPrincipalName = $userPrincipalName
                                    DisplayName = $member.DisplayName
                                    UserId = $member.ExternalDirectoryObjectId
                                    RoleName = $group.Name
                                    RoleDefinitionId = $group.Guid
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
                    $directAssignments = Get-RoleAssignment -ErrorAction SilentlyContinue | Where-Object { 
                        $_.AssignmentMethod -eq "Direct" -and $_.Role -in $privilegedRoleGroups
                    }
                    
                    foreach ($assignment in $directAssignments) {
                        # Include both user and group direct assignments
                        $isUserAssignment = $assignment.User -like "*@*" -and $assignment.User -notlike "*DiscoverySearchMailbox*"
                        $isGroupAssignment = $assignment.User -notlike "*@*" -or $assignment.User -like "*Group*"
                        
                        if ($isUserAssignment -or $isGroupAssignment) {
                            $principalType = if ($isUserAssignment) { "User" } else { "Group" }
                            
                            $exchangeResults += [PSCustomObject]@{
                                Service = "Exchange Online"
                                UserPrincipalName = $assignment.User
                                DisplayName = $assignment.User
                                UserId = $null
                                RoleName = $assignment.Role
                                RoleDefinitionId = $null
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
                }
                catch {
                    Write-Verbose "Could not retrieve direct role assignments: $($_.Exception.Message)"
                }
                
                # Disconnect Exchange
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
                
                return $exchangeResults
            }
            catch {
                Write-Error "Exchange Online runspace error: $($_.Exception.Message)"
                return @()
            }
        }
        
        # Create and execute runspace
        Write-Host "Creating isolated runspace for Exchange operations..." -ForegroundColor Gray
        
        $runspace = [runspacefactory]::CreateRunspace()
        $runspace.Open()
        
        $powershell = [powershell]::Create()
        $powershell.Runspace = $runspace
        $powershell.AddScript($exchangeScriptBlock).AddArgument($script:AppConfig.TenantId).AddArgument($script:AppConfig.ClientId).AddArgument($script:AppConfig.CertificateThumbprint).AddArgument($Organization)
        
        # Execute with timeout
        $asyncResult = $powershell.BeginInvoke()
        
        # Wait for completion with timeout (5 minutes)
        $timeout = 300
        $completed = $asyncResult.AsyncWaitHandle.WaitOne($timeout * 1000)
        
        if ($completed) {
            $exchangeResults = $powershell.EndInvoke($asyncResult)
            
            if ($powershell.Streams.Error.Count -gt 0) {
                foreach ($objerror in $powershell.Streams.Error) {
                    Write-Verbose "Exchange runspace error: $($objerror.Exception.Message)"
                }
            }
            
            if ($exchangeResults -and $exchangeResults.Count -gt 0) {
                Write-Host "✓ Retrieved $($exchangeResults.Count) Exchange role group assignments" -ForegroundColor Green
                $results += $exchangeResults
            }
            else {
                Write-Warning "No Exchange role group assignments retrieved"
            }
        }
        else {
            Write-Warning "Exchange runspace timed out after $timeout seconds"
            $powershell.Stop()
        }
        
        $powershell.Dispose()
        $runspace.Close()
        
        # Reconnect to Graph for post-processing and consistency
        Write-Host "Reconnecting to Microsoft Graph for data enrichment..." -ForegroundColor Gray
        Connect-MgGraph -TenantId $script:AppConfig.TenantId -ClientId $script:AppConfig.ClientId -CertificateThumbprint $script:AppConfig.CertificateThumbprint -NoWelcome
        
        # Post-process Exchange results to enrich with Graph data for consistency
        if ($exchangeResults -and $exchangeResults.Count -gt 0) {
            Write-Host "Enriching Exchange results with Graph data for consistency..." -ForegroundColor Gray
            
            $enrichedResults = @()
            foreach ($result in $exchangeResults) {
                $enrichedResult = $result.PSObject.Copy()
                
                # Try to enrich user data from Graph if UserId is available
                if ($result.UserId -and $result.PrincipalType -eq "User") {
                    try {
                        $graphUser = Get-MgUser -UserId $result.UserId -ErrorAction SilentlyContinue
                        if ($graphUser) {
                            $enrichedResult.UserEnabled = $graphUser.AccountEnabled
                            $enrichedResult.LastSignIn = $graphUser.SignInActivity.LastSignInDateTime
                            $enrichedResult.OnPremisesSyncEnabled = $graphUser.OnPremisesSyncEnabled
                            # Ensure UserPrincipalName is populated
                            if (-not $enrichedResult.UserPrincipalName -or $enrichedResult.UserPrincipalName -eq "") {
                                $enrichedResult.UserPrincipalName = $graphUser.UserPrincipalName
                            }
                        }
                    }
                    catch {
                        Write-Verbose "Could not enrich user data for $($result.UserId): $($_.Exception.Message)"
                    }
                }
                
                # Try to enrich group data from Graph if it's a group assignment
                if ($result.PrincipalType -eq "Group" -and $result.UserId) {
                    try {
                        $graphGroup = Get-MgGroup -GroupId $result.UserId -ErrorAction SilentlyContinue
                        if ($graphGroup) {
                            $enrichedResult.OnPremisesSyncEnabled = $graphGroup.OnPremisesSyncEnabled
                            # Get actual group member count
                            try {
                                $members = Get-MgGroupMember -GroupId $graphGroup.Id -All -ErrorAction SilentlyContinue
                                $enrichedResult.GroupMemberCount = $members.Count
                            }
                            catch {
                                $enrichedResult.GroupMemberCount = "Unknown"
                            }
                        }
                    }
                    catch {
                        Write-Verbose "Could not enrich group data for $($result.UserId): $($_.Exception.Message)"
                    }
                }
                
                $enrichedResults += $enrichedResult
            }
            
            $results += $enrichedResults
            Write-Host "✓ Enhanced $($enrichedResults.Count) Exchange role assignments with Graph data" -ForegroundColor Green
        }
        
        Write-Host "✓ Exchange Online privileged role audit completed" -ForegroundColor Green
        
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
    catch {
        Write-Error "Error auditing Exchange roles: $($_.Exception.Message)"
        Write-Host ""
        Write-Host "Troubleshooting Steps:" -ForegroundColor Yellow
        Write-Host "• Verify certificate is uploaded to Azure AD app registration" -ForegroundColor White
        Write-Host "• Ensure app has Exchange.ManageAsApp permission with admin consent" -ForegroundColor White
        Write-Host "• Check certificate expiration and validity" -ForegroundColor White
        Write-Host "• Verify Organization parameter matches your tenant" -ForegroundColor White
        Write-Host "• Run: Test-M365ModuleConflicts to check for assembly issues" -ForegroundColor White
        
        # Reconnect to Graph even if Exchange failed
        try {
            Connect-MgGraph -TenantId $script:AppConfig.TenantId -ClientId $script:AppConfig.ClientId -CertificateThumbprint $script:AppConfig.CertificateThumbprint -NoWelcome
        }
        catch {
            Write-Verbose "Could not reconnect to Graph: $($_.Exception.Message)"
        }
    }
    
    return $results
}

# Helper function to test Exchange connection without conflicts
function Test-ExchangeConnectionIsolated {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Organization,
        [string]$TenantId,
        [string]$ClientId,
        [string]$CertificateThumbprint
    )
    
    Write-Host "Testing Exchange connection in isolated runspace..." -ForegroundColor Cyan
    
    $testScript = {
        param($TenantId, $ClientId, $CertificateThumbprint, $Organization)
        
        try {
            Import-Module ExchangeOnlineManagement -Force
            Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -Organization $Organization -ShowBanner:$false
            
            # Test basic commands
            $orgConfig = Get-OrganizationConfig | Select-Object Name, Identity
            $roleGroupCount = (Get-RoleGroup).Count
            
            Disconnect-ExchangeOnline -Confirm:$false
            
            return @{
                Success = $true
                OrganizationName = $orgConfig.Name
                RoleGroupCount = $roleGroupCount
                Message = "Connection successful"
            }
        }
        catch {
            return @{
                Success = $false
                OrganizationName = $null
                RoleGroupCount = 0
                Message = $_.Exception.Message
            }
        }
    }
    
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()
    
    $powershell = [powershell]::Create()
    $powershell.Runspace = $runspace
    $powershell.AddScript($testScript).AddArgument($TenantId).AddArgument($ClientId).AddArgument($CertificateThumbprint).AddArgument($Organization)
    
    $result = $powershell.Invoke()[0]
    
    $powershell.Dispose()
    $runspace.Close()
    
    if ($result.Success) {
        Write-Host "✓ Exchange connection test successful" -ForegroundColor Green
        Write-Host "  Organization: $($result.OrganizationName)" -ForegroundColor Gray
        Write-Host "  Role Groups Available: $($result.RoleGroupCount)" -ForegroundColor Gray
    }
    else {
        Write-Host "❌ Exchange connection test failed" -ForegroundColor Red
        Write-Host "  Error: $($result.Message)" -ForegroundColor Yellow
    }
    
    return $result
}