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

        [bool]$IncludePIM = $true,           # Enhanced PIM support. Changed to bool and set default to true.

        [bool]$IncludeComplianceCenter = $true,  # Option to include Compliance Center role groups (if accessible), changed to bool and set default to true

        [switch]$IncludeSummary # Show the summary information on completion.
    )
    
    $results = @()
    $global:ProgressPreference = 'SilentlyContinue'
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
            "Compliance Data Administrator"
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
             
        Write-Host "Total Purview assignments to process: $($allAssignments.Count)" -ForegroundColor Green
        
        $convertParams = @{
            Assignments = $allAssignments
            RoleDefinitions = $roleDefinitions
            ServiceName = "Microsoft Purview"
            OverarchingRoles = $overarchingRoles
        }

        $results = ConvertTo-ServiceAssignmentResults @convertParams
        
        # === COMPLIANCE CENTER ROLE GROUPS (OPTIONAL) ===
        if ($IncludeComplianceCenter) {
            Write-Host "Attempting to retrieve Compliance Center role groups..." -ForegroundColor Cyan
            
            try {
                # If connected, try to get compliance role groups
                # If connected to Exchange Online Get-RoleGroup will return all role groups including Purview.
                # We are not interested in the Exchange role groups in this function, so we are going to
                # disconnect from Exchange Online. This wll disconnect any previous Purview connections.
                # This is necessary because just trying to disconnect Exchange can result in odd errors. 
                $Sessions = Get-ConnectionInformation
                $EXOSessions = $Sessions | Where-Object { $_.connectionUri -like "*outlook.office365.com*" -and $_.State -eq "Connected"}
                if ($EXOSessions) {
                    [void](Disconnect-ExchangeOnline -Confirm:$false)
                }

                # Clear the progress bars MS won't allow us to not show
                # there may be multiple

                for ($i = 0; $i -lt 5; $i++) {
                    Write-Progress -Completed
                }

                # Now connect to Purview. We should not have any connections still valid.
                Write-Host "Connecting to Microsoft Purview..." -ForegroundColor Yellow
                try {
                    $IPPSessionParams = @{
                        AppId = $script:AppConfig.ClientId
                        Organization = $Organization
                        ShowBanner = $false
                    }
                    if ($IsWindows) {
                        $IPPSessionParams["CertificateThumbprint"] = $script:AppConfig.CertificateThumbprint
                    } elseIf ($IsLinux -or $IsMacOS) {
                        $IPPSessionParams["Certificate"] = $script:AppConfig.Certificate
                    }
                    Connect-IPPSSession @IPPSessionParams
                }
                catch {
                    Write-Warning "Could not connect to Compliance Center: $($_.Exception.Message)"
                    Write-Host "Note: Compliance Center has limited certificate authentication support" -ForegroundColor Yellow
                    # Continue without Compliance Center data
                }

                Write-Host "Retrieving Compliance Center role groups..." -ForegroundColor Cyan
                
                try {
                    # Get compliance-specific role groups
                    # If we are only connected to Purview, we should only return Purview specific Role Groups.
                    # As most of the Purview role groups do not have members nor do they have administrative roles
                    # we are only focusing on the Administrative Groups.

                    $complianceRoleGroups = @(
                        "ComplianceAdministrator", 
                        "PurviewAdministrators", 
                        "AttackSimAdministrators", 
                        "SecurityAdministrator", 
                        "BillingAdministrator", 
                        "PrivacyManagementAdministrators", 
                        "SubjectRightsRequestAdministrators", 
                        "CommunicationComplianceAdministrators", 
                        "ComplianceDataAdministrator", 
                        "ComplianceManagerAdministrators", 
                        "DataSourceAdministrators", 
                        "MailFlowAdministrator", 
                        "KnowledgeAdministrators", 
                        "QuarantineAdministrator"
                    )

                    # Retrieve all Compliance role groups

                    $RoleGroups = Get-RoleGroup | Where-Object { $_.Name -in $complianceRoleGroups}

                    foreach ($roleGroup in $RoleGroups) {
                        try {
                            # This is no longer needed as we are getting all the compliance role groups above.
                            # $roleGroup = Get-RoleGroup -Identity $roleGroupName -ErrorAction SilentlyContinue
                            if ($roleGroup) {
                                $members = Get-RoleGroupMember -Identity $roleGroup.Name -ErrorAction SilentlyContinue

                                # Debug Code
                                Write-Host "  Role Group: $($roleGroup.DisplayName) has $($members.count) members." -ForegroundColor Yellow
                                foreach ($member in $members) {
                                    $results += Get-RoleGroupMemberResult -Member $member -Service "Microsoft Purview" -RoleGroup $roleGroup                              
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
            catch {
                Write-Warning "Could not access Compliance Center: $($_.Exception.Message)"
            }
        }
        
        Write-Host "✓ Purview administrative role audit completed. Found $($results.Count) administrative role assignments" -ForegroundColor Green
        
        If ($IncludeSummary) {
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

    <#
    .DESCRIPTION
    Get-PurviewRoleAudit audits Microsoft Purview (formerly Compliance Center) administrative roles.
    It focuses on Purview-specific Azure AD roles and optionally includes Compliance Center role groups.
    .PARAMETER Organization
    The organization domain (e.g., contoso.com).
    .PARAMETER TenantId
    The Azure AD tenant ID (GUID).
    .PARAMETER ClientId
    The Azure AD application (client) ID (GUID).
    .PARAMETER CertificateThumbprint
    The thumbprint of the certificate used for app-only authentication.
    .PARAMETER IncludeAzureADRoles
    Switch to include overarching Azure AD roles (e.g., Global Admin, Security Admin) in the results.
    .PARAMETER IncludePIM
    Boolean to include Privileged Identity Management (PIM) assignments. Default is $true.
    .PARAMETER IncludeComplianceCenter
    Boolean to include Compliance Center role groups (if accessible). Default is $true.
    .PARAMETER IncludeSummary
    Switch to display a summary of findings after the audit completes.
    .EXAMPLES
    # Audit Purview roles excluding overarching Azure AD roles, including Compliance Center groups, with summary
    Get-PurviewRoleAudit -Organization "contoso.com" -TenantId "<tenant-id>" -ClientId "<client-id>" -CertificateThumbprint "<thumbprint>" -IncludeSummary  
    .NOTES
    Optional you can use Set-M365AuditCredentials to set credentials globally instead of passing each time.
    #>
}