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
        <#
        Write-Host "Using configured certificate credentials for Exchange audit:" -ForegroundColor Cyan
        Write-Host "  Tenant ID: $($script:AppConfig.TenantId)" -ForegroundColor Gray
        Write-Host "  Client ID: $($script:AppConfig.ClientId)" -ForegroundColor Gray
        Write-Host "  Certificate Thumbprint: $($script:AppConfig.CertificateThumbprint)" -ForegroundColor Gray
        #>
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
            "Exchange Recipient Administrator"  # Legacy name for Exchange Administrator
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

        try {
        
            $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition -All | Where-Object { $_.DisplayName -in $rolesToInclude }
            Write-Host "Found $($roleDefinitions.Count) Exchange-related role definitions in Azure AD" -ForegroundColor Green
            
            $allAssignments = Get-RoleAssignmentsForService -RoleDefinitions $roleDefinitions -ServiceName "Exchange" -IncludePIM

            $convertParams = @{
                Assignments = $allAssignments
                RoleDefinitions = $roleDefinitions
                ServiceName = "Exchange Online"
                OverarchingRoles = $overarchingRoles
            }

            $results += ConvertTo-ServiceAssignmentResults @convertParams
    
            # Process Azure AD assignments
    <#         foreach ($assignment in $allAssignments) {
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
                        #LastSignIn = $null
                        PrincipalType = "Unknown"
                        OnPremisesSyncEnabled = $null
                        GroupMemberCount = $null
                    }
                    
                    # Try as user first
                    try {
                        $user = Get-MgUser -UserId $assignment.PrincipalId -Property "UserPrincipalName,DisplayName,AccountEnabled,OnPremisesSyncEnabled" -ErrorAction SilentlyContinue
                        if ($user) {
                            $principalInfo.UserPrincipalName = $user.UserPrincipalName
                            $principalInfo.DisplayName = $user.DisplayName
                            $principalInfo.UserEnabled = $user.AccountEnabled
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
                            #LastSignIn = $principalInfo.LastSignIn
                            Scope = $assignment.DirectoryScopeId
                            AssignmentId = $assignment.Id
                            AuthenticationType = "Certificate"
                            PrincipalType = $principalInfo.PrincipalType
                            #RoleSource = "AzureAD"
                            OnPremisesSyncEnabled = $principalInfo.OnPremisesSyncEnabled
                            PIMStartDateTime = $assignment.ScheduleInfo.StartDateTime
                            PIMEndDateTime = $assignment.ScheduleInfo.Expiration.EndDateTime
                            
                            # GroupMemberCount = $principalInfo.GroupMemberCount
                            # Additional fields for consistency
                            # RoleGroupDescription = $null
                            # OrganizationalUnit = $null
                            # ManagementScope = $null
                            # RecipientType = $null

                        }
                    }
                }
                catch {
                    Write-Verbose "Error processing Azure AD Exchange assignment: $($_.Exception.Message)"
                }
            }
            #>
        }
        catch {
            Write-Warning "Error retrieving Exchange Azure AD administrative roles: $($_.Exception.Message)"
        }

        # === STEP 2: Get Exchange role groups directly ===
        try {
            Write-Host "Retrieving Exchange role groups..." -ForegroundColor Cyan
            
            # Check if connected to Exchange Online
            
            $EXOSession = Get-ConnectionInformation | Where-Object { $_.connectionUri -like "*outlook.office365.com*" -and $_.State -eq 'Connected' }

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
                Write-Host "  ✓ Already connected to Exchange Online" -ForegroundColor Green
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
            
            foreach ($roleGroup in $targetRoleGroups) {
                try {
                    $members = Get-RoleGroupMember -Identity $roleGroup.Name -ErrorAction SilentlyContinue
                    Write-Host "  Role Group: $($roleGroup.Name) has $($members.count) members." -ForegroundColor Yellow
                    foreach ($member in $members) {
                        $results += Get-RoleGroupMemberResult -Member $member -Service 'Exchange Online' -RoleGroup $roleGroup                      
                    }
                }
                catch {
                    Write-Warning "Could not get members for role group $($group.Name): $($_.Exception.Message)"
                }
            }

        } catch {
            Write-Host "Error Retrieving Exchange Role Groups assignments: $($_.Exception.Message)"
        }
        
        Write-Host "  ✓ Exchange Online privileged role audit completed" -ForegroundColor Green
        
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

    <#
    .Description
    Audits privileged Exchange Online roles by retrieving role assignments from both Azure AD and Exchange role groups.
    Provides detailed assignment information including user/group details, role names, assignment types, and synchronization status for hybrid environments.
    .PARAMETER Organization
    The Exchange Online organization (usually your tenant domain).
    .PARAMETER TenantId
    (Optional) The Azure AD tenant ID for certificate authentication.
    .PARAMETER ClientId
    (Optional) The Azure AD application (client) ID for certificate authentication.
    .PARAMETER CertificateThumbprint
    (Optional) The thumbprint of the certificate used for authentication.
    .PARAMETER Summary
    If specified, outputs a summary of the audit results including counts and breakdowns.
    .PARAMETER IncludeAzureADRoles
    If specified, includes overarching Azure AD roles like "Exchange Administrator" in addition to Exchange-specific roles.
    .EXAMPLE
    Get-ExchangeRoleAudit -Organization "contoso.com" -Summary
    Audits Exchange roles for the contoso.com organization and provides a summary of results.
    .EXAMPLE
    Get-ExchangeRoleAudit -Organization "contoso.com" -TenantId "<id>" -ClientId "<id>" -CertificateThumbprint "<thumbprint>" -IncludeAzureADRoles
    .NOTES
    Optional you can use Set-M365AuditCredentials to set credentials globally instead of passing each time.
    #>
}