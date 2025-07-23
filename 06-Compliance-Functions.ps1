# 06-Compliance-Functions.ps1
# Focused Microsoft Purview Administrative Role Audit Function
# Updated to properly separate Azure AD roles from Compliance Center role groups

# 06-Compliance-Functions.ps1 - OPTIMIZED VERSION
# Enhanced Microsoft Purview Role Audit Function with proper Azure AD role filtering
# Applies the same optimization pattern as Teams, Defender, and other services

function Get-PurviewRoleAudit {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Organization,

        [string]$TenantId,
        [string]$ClientId,
        [string]$CertificateThumbprint,
        [switch]$IncludeAzureADRoles,  # New parameter to control inclusion of overarching roles
        [switch]$IncludePIM,           # Enhanced PIM support
        [switch]$IncludeComplianceCenter  # Option to include Compliance Center role groups (if accessible)
    )
    
    $results = @()
    
    try {
        # Certificate authentication is required for this function
        if ($TenantId -and $ClientId -and $CertificateThumbprint) {
            Set-M365AuditCertCredentials -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
        }
        
        # Verify certificate authentication is configured
        if (-not $script:AppConfig.UseAppAuth -or $script:AppConfig.AuthType -ne "Certificate") {
            throw "Certificate authentication is required for Purview role audit. Use Set-M365AuditCertCredentials first."
        }
        
        # Connect to Microsoft Graph if not already connected with certificate auth
        $context = Get-MgContext
        if (-not $context -or $context.AuthType -ne "AppOnly") {
            Write-Host "Connecting to Microsoft Graph for Purview roles..." -ForegroundColor Yellow
            
            Connect-MgGraph -TenantId $script:AppConfig.TenantId -ClientId $script:AppConfig.ClientId -CertificateThumbprint $script:AppConfig.CertificateThumbprint -NoWelcome
            
            # Verify app-only authentication
            $context = Get-MgContext
            if ($context.AuthType -ne "AppOnly") {
                throw "Expected app-only authentication but got: $($context.AuthType). Check certificate configuration."
            }
            
            Write-Host "✓ Connected with certificate authentication" -ForegroundColor Green
        }
        
        # === ENHANCED AZURE AD ROLE FILTERING ===
        # Purview-specific Azure AD administrative roles (NOT overarching roles)
        $purviewSpecificRoles = @(
            "Compliance Administrator",        # Purview-focused
            "Compliance Data Administrator",   # Purview-focused  
            "eDiscovery Administrator",
            "eDiscovery Manager", 
            "Information Protection Administrator",
            "Information Protection Analyst",
            "Information Protection Investigator",
            "Information Protection Reader",
            "Supervisory Review Administrator", # Clearly administrative
            "Data Loss Prevention Administrator",
            "Records Management Administrator",
            "Retention Administrator"
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
            "Privacy Management Administrator",
            "Information Protection Analyst"
        )
        
        # Determine which roles to include based on parameter
        $rolesToInclude = if ($IncludeAzureADRoles) {
            $purviewSpecificRoles + $overarchingRoles
        } else {
            $purviewSpecificRoles
        }
        
        Write-Host "Retrieving Purview-related Azure AD administrative roles..." -ForegroundColor Cyan
        $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition -All | Where-Object { $_.DisplayName -in $rolesToInclude }
        Write-Host "Found $($roleDefinitions.Count) Purview role definitions" -ForegroundColor Green

        $allAssignments = Get-RoleAssignmentsForService -RoleDefinitions $roleDefinitions -ServiceName "Purview" -IncludePIM:$IncludePIM
        
<#         # Get ALL assignment types (regular + PIM eligible + PIM active)
        $allAssignments = @()
        
        # 1. Regular assignments
        Write-Host "Checking regular Purview assignments..." -ForegroundColor Cyan
        $regularAssignments = Get-MgRoleManagementDirectoryRoleAssignment -All | Where-Object { $_.RoleDefinitionId -in $roleDefinitions.Id }
        if ($regularAssignments) { $allAssignments += $regularAssignments }
        Write-Host "Found $($regularAssignments.Count) regular assignments" -ForegroundColor Gray
        
        # 2. PIM eligible assignments
        if ($IncludePIM) {
            Write-Host "Checking PIM eligible Purview assignments..." -ForegroundColor Cyan
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
            Write-Host "Checking PIM active Purview assignments..." -ForegroundColor Cyan
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
        }
  #>       
        Write-Host "Total Purview assignments to process: $($allAssignments.Count)" -ForegroundColor Green
        
        # Process all assignments
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
                
                # Resolve principal (users, groups, service principals)
                $principalInfo = @{
                    UserPrincipalName = "Unknown"
                    DisplayName = "Unknown"
                    UserId = $assignment.PrincipalId
                    UserEnabled = $null
                    #LastSignIn = $null
                    PrincipalType = "Unknown"
                    OnPremisesSyncEnabled = $null
                }
                
                # Try as user first
                try {
                    $user = Get-MgUser -UserId $assignment.PrincipalId -Property "UserPrincipalName,DisplayName,AccountEnabled,OnPremisesSyncEnabled" -ErrorAction SilentlyContinue
                    if ($user) {
                        $principalInfo.UserPrincipalName = $user.UserPrincipalName
                        $principalInfo.DisplayName = $user.DisplayName
                        $principalInfo.UserEnabled = $user.AccountEnabled
                        #$principalInfo.LastSignIn = $user.SignInActivity.LastSignInDateTime
                        $principalInfo.PrincipalType = "User"
                        $principalInfo.OnPremisesSyncEnabled = $user.OnPremisesSyncEnabled
                    }
                }
                catch { }
                
                # Try as group if not user
                if ($principalInfo.PrincipalType -eq "Unknown") {
                    try {
                        $group = Get-MgGroup -GroupId $assignment.PrincipalId -Property "Mail,DisplayName,OnPremisesSyncEnabled" -ErrorAction SilentlyContinue
                        if ($group) {
                            $principalInfo.UserPrincipalName = $group.Mail
                            $principalInfo.DisplayName = "$($group.DisplayName) (Group)"
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
                
                $results += [PSCustomObject]@{
                    Service = "Microsoft Purview"
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
                    RoleSource = "AzureAD"
                    OnPremisesSyncEnabled = $principalInfo.OnPremisesSyncEnabled
                    RoleGroupDescription = $role.Description
                    PIMStartDateTime = $assignment.ScheduleInfo.StartDateTime
                    PIMEndDateTime = $assignment.ScheduleInfo.Expiration.EndDateTime
                }
                
            }
            catch {
                Write-Verbose "Error processing Purview assignment: $($_.Exception.Message)"
            }
        }
        
        # === COMPLIANCE CENTER ROLE GROUPS (OPTIONAL) ===
        if ($IncludeComplianceCenter) {
            Write-Host "Attempting to retrieve Compliance Center role groups..." -ForegroundColor Cyan
            
            try {
                # Check if connected to Exchange/Compliance PowerShell
                $complianceSession = Get-PSSession | Where-Object { 
                    $_.ComputerName -like "*compliance*" -or $_.ComputerName -like "*protection*" 
                }
                
                if (-not $complianceSession) {
                    Write-Host "Connecting to Security & Compliance Center..." -ForegroundColor Yellow
                    
                    # Attempt connection with certificate authentication
                    try {
                        if ($IsWindows) {
                            Connect-IPPSSession -AppId $script:AppConfig.ClientId -CertificateThumbprint $script:AppConfig.CertificateThumbprint -Organization $Organization -ShowBanner:$false
                        } elseif ($IsLinux -or $IsMacOS) {
                            Connect-IPPSSession -AppId $script:AppConfig.ClientId -Certificate $script:AppConfig.Certificate -Organization $Organization -ShowBanner:$false
                        }
                        Write-Host "✓ Connected to Security & Compliance Center" -ForegroundColor Green
                    }
                    catch {
                        Write-Warning "Could not connect to Compliance Center: $($_.Exception.Message)"
                        Write-Host "Note: Compliance Center has limited certificate authentication support" -ForegroundColor Yellow
                        # Continue without Compliance Center data
                    }
                }
                
                # If connected, try to get compliance role groups
                if (Get-PSSession | Where-Object { $_.ComputerName -like "*compliance*" -or $_.ComputerName -like "*protection*" }) {
                    Write-Host "Retrieving Compliance Center role groups..." -ForegroundColor Cyan
                    
                    try {
                        # Get compliance-specific role groups
                        $complianceRoleGroups = @(
                            "Compliance Administrator",
                            "Compliance Data Administrator", 
                            "eDiscovery Manager",
                            "Organization Management",
                            "Records Management",
                            "Reviewer",
                            "Supervisory Review"
                        )
                        
                        foreach ($roleGroupName in $complianceRoleGroups) {
                            try {
                                $roleGroup = Get-RoleGroup -Identity $roleGroupName -ErrorAction SilentlyContinue
                                if ($roleGroup) {
                                    $members = Get-RoleGroupMember -Identity $roleGroup.Identity -ErrorAction SilentlyContinue
                                    
                                    foreach ($member in $members) {
                                        # Include users AND groups
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
                                                    $graphUser = Get-MgUser -UserId $member.ExternalDirectoryObjectId -Property "AccountEnabled,SignInActivity,OnPremisesSyncEnabled" -ErrorAction SilentlyContinue
                                                    if ($graphUser) {
                                                        $userEnabled = $graphUser.AccountEnabled
                                                        $lastSignIn = $graphUser.SignInActivity.LastSignInDateTime
                                                        $onPremisesSyncEnabled = $graphUser.OnPremisesSyncEnabled
                                                    }
                                                }
                                                catch {
                                                    Write-Verbose "Could not retrieve Graph data for compliance user $($member.PrimarySmtpAddress): $($_.Exception.Message)"
                                                }
                                            }
                                            
                                            $results += [PSCustomObject]@{
                                                Service = "Microsoft Purview"
                                                UserPrincipalName = $userPrincipalName
                                                DisplayName = $member.DisplayName
                                                UserId = $member.ExternalDirectoryObjectId
                                                RoleName = $roleGroup.Name
                                                RoleDefinitionId = $roleGroup.Guid
                                                RoleScope = "Service-Specific"  # Compliance role groups are service-specific
                                                AssignmentType = "Role Group Member"
                                                AssignedDateTime = $null
                                                UserEnabled = $userEnabled
                                                LastSignIn = $lastSignIn
                                                Scope = "Organization"
                                                AssignmentId = $roleGroup.Identity
                                                AuthenticationType = "Certificate"
                                                PrincipalType = $principalType
                                                RoleSource = "ComplianceCenter"
                                                OnPremisesSyncEnabled = $onPremisesSyncEnabled
                                                #RoleGroupDescription = $roleGroup.Description
                                                RecipientType = $member.RecipientType
                                                # Additional fields for consistency
                                                PIMStartDateTime = $null
                                                PIMEndDateTime = $null
                                            }
                                        }
                                    }
                                }
                            }
                            catch {
                                Write-Verbose "Could not retrieve role group $roleGroupName`: $($_.Exception.Message)"
                            }
                        }
                    }
                    catch {
                        Write-Warning "Error retrieving Compliance Center role groups: $($_.Exception.Message)"
                    }
                }
            }
            catch {
                Write-Warning "Could not access Compliance Center: $($_.Exception.Message)"
            }
        }
        
        Write-Host "✓ Purview administrative role audit completed. Found $($results.Count) administrative role assignments" -ForegroundColor Green
        
        # Provide feedback about role filtering
        if (-not $IncludeAzureADRoles) {
            Write-Host "  (Excluding overarching Azure AD roles - use -IncludeAzureADRoles to include)" -ForegroundColor Yellow
        }
        
        if ($IncludeComplianceCenter) {
            $complianceCenterResults = $results | Where-Object { $_.RoleSource -eq "ComplianceCenter" }
            Write-Host "  (Compliance Center role groups: $($complianceCenterResults.Count))" -ForegroundColor Cyan
        }
        
        # Show detailed breakdown
        if ($results.Count -gt 0) {
            Write-Host ""
            Write-Host "Administrative role breakdown:" -ForegroundColor Cyan
            
            $sourceSummary = $results | Group-Object RoleSource
            Write-Host "Role sources:" -ForegroundColor Yellow
            foreach ($source in $sourceSummary) {
                Write-Host "  $($source.Name): $($source.Count)" -ForegroundColor White
            }
            
            $typeSummary = $results | Group-Object PrincipalType
            Write-Host "Principal types:" -ForegroundColor Yellow
            foreach ($type in $typeSummary) {
                Write-Host "  $($type.Name): $($type.Count)" -ForegroundColor White
            }
            
            $assignmentTypeSummary = $results | Group-Object AssignmentType
            Write-Host "Assignment types:" -ForegroundColor Yellow
            foreach ($type in $assignmentTypeSummary) {
                Write-Host "  $($type.Name): $($type.Count)" -ForegroundColor White
            }
            
            $scopeSummary = $results | Group-Object RoleScope
            Write-Host "Role scope:" -ForegroundColor Yellow
            foreach ($scope in $scopeSummary) {
                Write-Host "  $($scope.Name): $($scope.Count)" -ForegroundColor White
            }
            
            # Show top roles
            $roleSummary = $results | Group-Object RoleName | Sort-Object Count -Descending | Select-Object -First 5
            Write-Host "Top Purview roles:" -ForegroundColor Yellow
            foreach ($role in $roleSummary) {
                Write-Host "  $($role.Name): $($role.Count) assignments" -ForegroundColor White
            }
        }
        
        Write-Host ""
        Write-Host "=== SCOPE CLARIFICATION ===" -ForegroundColor Green
        Write-Host "✓ Focused on Purview/Compliance Azure AD administrative roles only" -ForegroundColor Green
        Write-Host "✓ Included: Compliance Administrator, eDiscovery roles, Information Protection roles" -ForegroundColor Green
        Write-Host "✓ Included: Records Management, DLP Administrator, Retention Administrator" -ForegroundColor Green
        if ($IncludeComplianceCenter) {
            Write-Host "✓ Included: Compliance Center role groups (where accessible)" -ForegroundColor Green
        } else {
            Write-Host "✓ Excluded: Compliance Center role groups (use -IncludeComplianceCenter to include)" -ForegroundColor Green
        }
        Write-Host "✓ Enhanced filtering prevents duplication with Azure AD audit" -ForegroundColor Green
        
    }
    catch {
        Write-Error "Error auditing Purview administrative roles: $($_.Exception.Message)"
        
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
            Write-Host "• For Compliance Center: Exchange.ManageAsApp" -ForegroundColor White
            Write-Host "Run: Get-M365AuditRequiredPermissions for complete list" -ForegroundColor White
        }
        
        throw
    }
    finally {
        # Clean up Compliance Center connection if we created it
        if ($IncludeComplianceCenter) {
            try {
                if (Get-PSSession | Where-Object { $_.ComputerName -like "*compliance*" -or $_.ComputerName -like "*protection*" }) {
                    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
                    Write-Host "Disconnected from Security & Compliance Center" -ForegroundColor Gray
                }
            }
            catch {
                # Ignore cleanup errors
            }
        }
    }
    
    return $results
}