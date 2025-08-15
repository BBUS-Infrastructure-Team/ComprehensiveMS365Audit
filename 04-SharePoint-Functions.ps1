# 04-SharePoint-Functions.ps1
# Updated Get-SharePointRoleAudit function focused ONLY on administrative roles
# Removed: Site-level permissions, Search Center admins, Term Store access verification

function Get-SharePointRoleAudit {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantUrl,
        
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
            throw "Certificate authentication is required for SharePoint role audit. Use Set-M365AuditCertCredentials first."
        }
        
        # Import and check PnP PowerShell
        if (-not (Get-Module -ListAvailable -Name "PnP.PowerShell")) {
            Write-Warning "PnP.PowerShell module not found. Installing..."
            Install-Module -Name PnP.PowerShell -Force -AllowClobber -Scope CurrentUser
        }
        
        Import-Module PnP.PowerShell -Force
        
        # Connect to SharePoint admin center with certificate authentication
        Write-Host "Connecting to SharePoint admin center with certificate authentication..." -ForegroundColor Yellow
        
        try {
            # Primary connection attempt using certificate
            Write-Host "Using certificate authentication..." -ForegroundColor Cyan
            Connect-PnPOnline -Url $TenantUrl -ClientId $script:AppConfig.ClientId -Thumbprint $script:AppConfig.CertificateThumbprint -Tenant $script:AppConfig.TenantId
            
            # Verify connection
            $connection = Get-PnPConnection
            if (-not $connection) {
                throw "Failed to establish connection to SharePoint admin center"
            }
            
            Write-Host "✓ Connected to SharePoint admin center successfully with certificate authentication" -ForegroundColor Green
            Write-Host "Authentication Type: Certificate" -ForegroundColor Cyan
        }
        catch {
            Write-Error "SharePoint certificate authentication failed: $($_.Exception.Message)"
            Write-Host "Troubleshooting steps:" -ForegroundColor Yellow
            Write-Host "• Verify certificate is uploaded to the app registration in Azure AD" -ForegroundColor White
            Write-Host "• Ensure app registration has SharePoint Sites.FullControl.All permission" -ForegroundColor White
            Write-Host "• Verify certificate exists in Windows Certificate Store" -ForegroundColor White
            Write-Host "• Check that admin consent has been granted for SharePoint permissions" -ForegroundColor White
            throw "SharePoint connection failed with certificate authentication"
        }
        
        # === ENHANCED AZURE AD ROLE FILTERING ===
        Write-Host "Retrieving SharePoint-related Azure AD administrative roles..." -ForegroundColor Cyan
        
        # Connect to Microsoft Graph if not already connected with certificate auth
        $context = Get-MgContext
        if (-not $context -or $context.AuthType -ne "AppOnly") {
            Write-Host "Connecting to Microsoft Graph for SharePoint administrative roles..." -ForegroundColor Yellow
            
            Connect-MgGraph -TenantId $script:AppConfig.TenantId -ClientId $script:AppConfig.ClientId -CertificateThumbprint $script:AppConfig.CertificateThumbprint -NoWelcome
            
            # Verify app-only authentication
            $context = Get-MgContext
            if ($context.AuthType -ne "AppOnly") {
                throw "Expected app-only authentication but got: $($context.AuthType). Check certificate configuration."
            }
            
            Write-Host "✓ Connected to Microsoft Graph with certificate authentication" -ForegroundColor Green
        }
        
        # SharePoint-specific Azure AD roles (NOT overarching roles)
        $sharePointSpecificRoles = @(
            "SharePoint Service Administrator",  # Legacy name for SharePoint Administrator
            "SharePoint Administrator"
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
            $sharePointSpecificRoles + $overarchingRoles
        } else {
            $sharePointSpecificRoles
        }
        
        # FIX 1: Add -All parameter to get ALL role definitions
        try {
            $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition -All | Where-Object { $_.DisplayName -in $rolesToInclude }
            Write-Host "Found $($roleDefinitions.Count) SharePoint-related administrative role definitions" -ForegroundColor Green

            $allAssignments = Get-RoleAssignmentsForService -RoleDefinitions $roleDefinitions -ServiceName "SharePoint" -IncludePIM 

            $convertParams = @{
                Assignments = $allAssignments
                RoleDefinitions = $roleDefinitions
                ServiceName = "SharePoint Online"
                OverarchingRoles = $overarchingRoles
            }            

            $results = ConvertTo-ServiceAssignmentResults @convertParams
            
            # Process all assignments
<#             foreach ($assignment in $allAssignments) {
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
                        Service = "SharePoint Online"
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
                        RoleType = "AzureAD"
                        OnPremisesSyncEnabled = $principalInfo.OnPremisesSyncEnabled
                        PIMStartDateTime = $assignment.ScheduleInfo.StartDateTime
                        PIMEndDateTime = $assignment.ScheduleInfo.Expiration.EndDateTime
                    }
                    
                }
                catch {
                    Write-Verbose "Error processing SharePoint assignment: $($_.Exception.Message)"
                }
            } #>
        }
        catch {
            Write-Warning "Error retrieving SharePoint Azure AD administrative roles: $($_.Exception.Message)"
        }
        
        # === SHAREPOINT TENANT-LEVEL ADMINISTRATIVE ROLES ONLY ===
        Write-Host "Verifying SharePoint tenant administrative access..." -ForegroundColor Cyan
        try {
            # Verify admin center connection and access
            $adminCenterUrl = $TenantUrl
            Connect-PnPOnline -Url $adminCenterUrl -ClientId $script:AppConfig.ClientId -Thumbprint $script:AppConfig.CertificateThumbprint -Tenant $script:AppConfig.TenantId
            
            # Get tenant properties to verify administrative access
            $tenantProperties = Get-PnPTenant -ErrorAction SilentlyContinue
            if ($tenantProperties) {
                Write-Host "✓ Successfully accessed tenant administrative properties" -ForegroundColor Green
            }
            
        }
        catch {
            Write-Verbose "Could not access tenant properties: $($_.Exception.Message)"
        }
        
        # Get SharePoint App Catalog administrators (TENANT-LEVEL ONLY)
        Write-Host "Checking SharePoint Tenant App Catalog administrators..." -ForegroundColor Cyan
        try {
            # Check if tenant app catalog exists
            $appCatalog = Get-PnPTenantAppCatalogUrl -ErrorAction SilentlyContinue
            if ($appCatalog) {
                Write-Host "Tenant App Catalog found: $appCatalog" -ForegroundColor Gray
                
                # Connect to app catalog
                Connect-PnPOnline -Url $appCatalog -ClientId $script:AppConfig.ClientId -Thumbprint $script:AppConfig.CertificateThumbprint -Tenant $script:AppConfig.TenantId
                
                # Get app catalog administrators (TENANT-LEVEL ADMINISTRATIVE ROLE)
                $appCatalogAdmins = Get-PnPSiteCollectionAdmin -ErrorAction SilentlyContinue
                foreach ($admin in $appCatalogAdmins) {
                    $cleanLoginName = $admin.LoginName -replace "i:0#\.f\|membership\|", "" -replace "i:0#\.w\|", ""
                    
                    $results += [PSCustomObject]@{
                        Service = "SharePoint Online"
                        UserPrincipalName = $cleanLoginName
                        DisplayName = $admin.Title
                        UserId = $null
                        RoleName = "Tenant App Catalog Administrator"
                        RoleDefinitionId = $null
                        RoleScope = "Service-Specific"  # New property
                        AssignmentType = "Active"
                        AssignedDateTime = $null
                        UserEnabled = $null
                        #LastSignIn = $null
                        Scope = "Tenant App Catalog"
                        AssignmentId = $null
                        #AuthenticationType = "Certificate"
                        PrincipalType = "User"
                        RoleType = "SharePointSpecific"
                        PIMStartDateTime = $null
                        PIMEndDateTime = $null
                    }
                }
                Write-Host "Found $($appCatalogAdmins.Count) Tenant App Catalog administrators" -ForegroundColor Gray
            }
            else {
                Write-Host "No tenant app catalog found" -ForegroundColor Gray
            }
        }
        catch {
            Write-Verbose "Could not access Tenant App Catalog: $($_.Exception.Message)"
        }
        
        Write-Host "✓ SharePoint administrative role audit completed. Found $($results.Count) administrative role assignments" -ForegroundColor Green
        
        # Provide feedback about role filtering
        if (-not $IncludeAzureADRoles) {
            Write-Host "  (Excluding overarching Azure AD roles - use -IncludeAzureADRoles to include)" -ForegroundColor Yellow
        }
        
        # Show breakdown
        if ($results.Count -gt 0) {
            $roleSummary = $results | Group-Object RoleName
            $typeSummary = $results | Group-Object PrincipalType
            $assignmentTypeSummary = $results | Group-Object AssignmentType
            $scopeSummary = $results | Group-Object RoleScope
            
            Write-Host ""
            Write-Host "Administrative role breakdown:" -ForegroundColor Cyan
            foreach ($role in $roleSummary) {
                Write-Host "  $($role.Name): $($role.Count)" -ForegroundColor White
            }
            
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
            
            Write-Host ""
            Write-Host "=== SCOPE CLARIFICATION ===" -ForegroundColor Green
            Write-Host "✓ Focused on tenant-level administrative roles only" -ForegroundColor Green
            Write-Host "✓ Excluded: Site-level permissions and individual site administrators" -ForegroundColor Green
            Write-Host "✓ Excluded: Search Center administrators for individual sites" -ForegroundColor Green
            Write-Host "✓ Excluded: Term Store access verification" -ForegroundColor Green
            Write-Host "✓ Included: Tenant App Catalog administrators (tenant-level administrative role)" -ForegroundColor Green
        }
        
    }
    catch {
        Write-Error "Error in SharePoint administrative role audit: $($_.Exception.Message)"
        
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
            Write-Host "• Sites.FullControl.All" -ForegroundColor White
            Write-Host "• Sites.Read.All" -ForegroundColor White
            Write-Host "• Directory.Read.All (for Azure AD roles)" -ForegroundColor White
            Write-Host "• RoleManagement.Read.All (for PIM)" -ForegroundColor White
            Write-Host "Run: Get-M365AuditRequiredPermissions for complete list" -ForegroundColor White
        }
        
        throw
    }
    finally {
        # Clean up connection
        try {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
        }
        catch {
            # Ignore cleanup errors
        }
    }
    
    return $results

    <#
    .DESCRIPTION
    Retrieves SharePoint Online administrative roles and assignments, focusing on tenant-level roles only.
    Excludes site-level permissions, Search Center admins, and Term Store access verification.
    Optionally includes overarching Azure AD roles with -IncludeAzureADRoles switch.
    .PARAMETER TenantUrl
    The URL of the SharePoint admin center (e.g., https://contoso-admin.sharepoint.com).
    .PARAMETER TenantId
    The Azure AD tenant ID for certificate authentication.
    .PARAMETER ClientId
    The client ID of the Azure AD app registration for certificate authentication.
    .PARAMETER CertificateThumbprint
    The thumbprint of the certificate used for authentication.
    .PARAMETER IncludeAzureADRoles
    Switch to include overarching Azure AD roles (e.g., Global Admin) in the results.
    .EXAMPLES
    $roles = Get-SharePointRoleAudit -TenantUrl "https://contoso-admin.sharepoint.com" -TenantId "<tenant-id>" -ClientId "<client-id>" -CertificateThumbprint "<thumbprint>"
    Retrieves SharePoint tenant-level administrative roles only.    
    $roles = Get-SharePointRoleAudit -TenantUrl "https://contoso-admin.sharepoint.com" -TenantId "<tenant-id>" -ClientId "<client-id>" -CertificateThumbprint "<thumbprint>" -IncludeAzureADRoles
    Retrieves SharePoint tenant-level administrative roles including overarching Azure AD roles.
    .NOTES
    Requires PnP.PowerShell and Microsoft.Graph modules.
    Optionally you can use Set-M365AUditCredentials to set app credentials for subsequent calls.
    #>
}