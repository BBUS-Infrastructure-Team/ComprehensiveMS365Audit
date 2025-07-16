# 04-SharePoint-Functions.ps1
# SharePoint Online ADMINISTRATIVE role audit functions - Certificate Authentication Only
# Fixed to focus on administrative roles only, not individual site permissions

function Get-SharePointRoleAudit {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantUrl,
        
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
            # Write-Host "Client ID: $($script:AppConfig.ClientId)" -ForegroundColor Gray
            # Write-Host "Certificate Thumbprint: $($script:AppConfig.CertificateThumbprint)" -ForegroundColor Gray
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
        
        # === FOCUS ON ADMINISTRATIVE ROLES ONLY ===
        Write-Host "Retrieving SharePoint administrative roles..." -ForegroundColor Cyan
        
        # 1. Verify SharePoint tenant access
        Write-Host "Verifying SharePoint tenant access..." -ForegroundColor Cyan
        try {
            # Verify admin center connection and access
            $adminCenterUrl = $TenantUrl
            Connect-PnPOnline -Url $adminCenterUrl -ClientId $script:AppConfig.ClientId -Thumbprint $script:AppConfig.CertificateThumbprint -Tenant $script:AppConfig.TenantId
            
            # Get tenant properties to verify administrative access
            $tenantProperties = Get-PnPTenant -ErrorAction SilentlyContinue
            if ($tenantProperties) {
                Write-Host "✓ Successfully accessed tenant properties" -ForegroundColor Green
            }
            
        }
        catch {
            Write-Verbose "Could not access tenant properties: $($_.Exception.Message)"
        }
        
        # 2. Get SharePoint-related Azure AD roles using Microsoft Graph
        Write-Host "Retrieving SharePoint-related Azure AD roles..." -ForegroundColor Cyan
        try {
            # Connect to Microsoft Graph if not already connected
            $context = Get-MgContext
            if (-not $context -or $context.AuthType -ne "AppOnly") {
                Connect-MgGraph -TenantId $script:AppConfig.TenantId -ClientId $script:AppConfig.ClientId -CertificateThumbprint $script:AppConfig.CertificateThumbprint -NoWelcome
            }
            
            # SharePoint-related Azure AD roles
            $sharePointRoles = @(
                "SharePoint Administrator",
                "SharePoint Service Administrator",  # Legacy name
                "Global Administrator",
                "Application Administrator",
                "Cloud Application Administrator"
            )
            
            $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition | Where-Object { $_.DisplayName -in $sharePointRoles }
            Write-Host "Found $($roleDefinitions.Count) SharePoint-related role definitions" -ForegroundColor Green
            
            # Get ALL assignment types (regular + PIM eligible + PIM active)
            $allAssignments = @()
            
            # Regular assignments
            $regularAssignments = Get-MgRoleManagementDirectoryRoleAssignment | Where-Object { $_.RoleDefinitionId -in $roleDefinitions.Id }
            if ($regularAssignments) { $allAssignments += $regularAssignments }
            Write-Host "Found $($regularAssignments.Count) regular SharePoint role assignments" -ForegroundColor Gray
            
            # PIM eligible assignments
            try {
                foreach ($roleId in $roleDefinitions.Id) {
                    $pimEligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -Filter "roleDefinitionId eq '$roleId'" -ErrorAction SilentlyContinue
                    if ($pimEligible) {
                        $allAssignments += $pimEligible
                    }
                }
                $pimEligibleCount = ($allAssignments | Where-Object { $_.PSObject.TypeNames -contains "Microsoft.Graph.PowerShell.Models.MicrosoftGraphUnifiedRoleEligibilitySchedule" }).Count
                Write-Host "Found $pimEligibleCount PIM eligible SharePoint assignments" -ForegroundColor Gray
            }
            catch {
                Write-Verbose "Could not retrieve PIM eligible assignments: $($_.Exception.Message)"
            }
            
            # PIM active assignments
            try {
                foreach ($roleId in $roleDefinitions.Id) {
                    $pimActive = Get-MgRoleManagementDirectoryRoleAssignmentSchedule -Filter "roleDefinitionId eq '$roleId'" -ErrorAction SilentlyContinue
                    if ($pimActive) {
                        $allAssignments += $pimActive
                    }
                }
                $pimActiveCount = ($allAssignments | Where-Object { $_.PSObject.TypeNames -contains "Microsoft.Graph.PowerShell.Models.MicrosoftGraphUnifiedRoleAssignmentSchedule" }).Count
                Write-Host "Found $pimActiveCount PIM active SharePoint assignments" -ForegroundColor Gray
            }
            catch {
                Write-Verbose "Could not retrieve PIM active assignments: $($_.Exception.Message)"
            }
            
            Write-Host "Total SharePoint administrative assignments to process: $($allAssignments.Count)" -ForegroundColor Green
            
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
                        Service = "SharePoint Online"
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
                        RoleType = "AzureAD"
                        PIMStartDateTime = $assignment.ScheduleInfo.StartDateTime
                        PIMEndDateTime = $assignment.ScheduleInfo.Expiration.EndDateTime
                    }
                    
                }
                catch {
                    Write-Verbose "Error processing SharePoint assignment: $($_.Exception.Message)"
                }
            }
        }
        catch {
            Write-Warning "Error retrieving SharePoint Azure AD roles: $($_.Exception.Message)"
        }
        
        # 3. Get SharePoint App Catalog administrators (if app catalog exists)
        Write-Host "Checking SharePoint App Catalog administrators..." -ForegroundColor Cyan
        try {
            # Check if tenant app catalog exists
            $appCatalog = Get-PnPTenantAppCatalogUrl -ErrorAction SilentlyContinue
            if ($appCatalog) {
                Write-Host "App Catalog found: $appCatalog" -ForegroundColor Gray
                
                # Connect to app catalog
                Connect-PnPOnline -Url $appCatalog -ClientId $script:AppConfig.ClientId -Thumbprint $script:AppConfig.CertificateThumbprint -Tenant $script:AppConfig.TenantId
                
                # Get app catalog administrators
                $appCatalogAdmins = Get-PnPSiteCollectionAdmin -ErrorAction SilentlyContinue
                foreach ($admin in $appCatalogAdmins) {
                    $cleanLoginName = $admin.LoginName -replace "i:0#\.f\|membership\|", "" -replace "i:0#\.w\|", ""
                    
                    $results += [PSCustomObject]@{
                        Service = "SharePoint Online"
                        UserPrincipalName = $cleanLoginName
                        DisplayName = $admin.Title
                        UserId = $null
                        RoleName = "App Catalog Administrator"
                        RoleDefinitionId = $null
                        AssignmentType = "Active"
                        AssignedDateTime = $null
                        UserEnabled = $null
                        LastSignIn = $null
                        Scope = $appCatalog
                        AssignmentId = $null
                        AuthenticationType = "Certificate"
                        PrincipalType = "User"
                        RoleType = "SharePointSpecific"
                        SiteTitle = "Tenant App Catalog"
                        Template = "APPCATALOG#0"
                    }
                }
                Write-Host "Found $($appCatalogAdmins.Count) App Catalog administrators" -ForegroundColor Gray
            }
            else {
                Write-Host "No tenant app catalog found" -ForegroundColor Gray
            }
        }
        catch {
            Write-Verbose "Could not access App Catalog: $($_.Exception.Message)"
        }
        
        # 4. Get Search Service Application administrators
        Write-Host "Checking Search Center administrators..." -ForegroundColor Cyan
        try {
            # Look for search center sites
            $searchSites = Get-PnPTenantSite -Template "SRCHCEN#0" -ErrorAction SilentlyContinue
            foreach ($searchSite in $searchSites) {
                try {
                    Connect-PnPOnline -Url $searchSite.Url -ClientId $script:AppConfig.ClientId -Thumbprint $script:AppConfig.CertificateThumbprint -Tenant $script:AppConfig.TenantId
                    
                    $searchAdmins = Get-PnPSiteCollectionAdmin -ErrorAction SilentlyContinue
                    foreach ($admin in $searchAdmins) {
                        $cleanLoginName = $admin.LoginName -replace "i:0#\.f\|membership\|", "" -replace "i:0#\.w\|", ""
                        
                        $results += [PSCustomObject]@{
                            Service = "SharePoint Online"
                            UserPrincipalName = $cleanLoginName
                            DisplayName = $admin.Title
                            UserId = $null
                            RoleName = "Search Center Administrator"
                            RoleDefinitionId = $null
                            AssignmentType = "Active"
                            AssignedDateTime = $null
                            UserEnabled = $null
                            LastSignIn = $null
                            Scope = $searchSite.Url
                            AssignmentId = $null
                            AuthenticationType = "Certificate"
                            PrincipalType = "User"
                            RoleType = "SharePointSpecific"
                            SiteTitle = $searchSite.Title
                            Template = $searchSite.Template
                        }
                    }
                }
                catch {
                    Write-Verbose "Could not access search site $($searchSite.Url): $($_.Exception.Message)"
                }
            }
            Write-Host "Processed $($searchSites.Count) Search Center sites" -ForegroundColor Gray
        }
        catch {
            Write-Verbose "Could not retrieve search sites: $($_.Exception.Message)"
        }
        
        # 5. Check Term Store access (if accessible)
        Write-Host "Checking Term Store access..." -ForegroundColor Cyan
        try {
            # Reconnect to admin center
            Connect-PnPOnline -Url $TenantUrl -ClientId $script:AppConfig.ClientId -Thumbprint $script:AppConfig.CertificateThumbprint -Tenant $script:AppConfig.TenantId
            
            # Try to get term store information
            $termStore = Get-PnPTermStore -ErrorAction SilentlyContinue
            if ($termStore) {
                Write-Host "✓ Term Store access verified" -ForegroundColor Gray
                
                # Add term store access verification to results
                $results += [PSCustomObject]@{
                    Service = "SharePoint Online"
                    UserPrincipalName = "System Configuration"
                    DisplayName = "Term Store Configuration"
                    UserId = $null
                    RoleName = "Term Store Access Verified"
                    RoleDefinitionId = $null
                    AssignmentType = "System"
                    AssignedDateTime = (Get-Date)
                    UserEnabled = $null
                    LastSignIn = $null
                    Scope = "Term Store"
                    AssignmentId = $null
                    AuthenticationType = "Certificate"
                    PrincipalType = "System"
                    RoleType = "SharePointSpecific"
                }
            }
            else {
                Write-Host "No Term Store access available" -ForegroundColor Gray
            }
        }
        catch {
            Write-Verbose "Could not access Term Store: $($_.Exception.Message)"
        }
        
        Write-Host "✓ SharePoint administrative role audit completed. Found $($results.Count) administrative role assignments" -ForegroundColor Green
        
        # Show breakdown
        if ($results.Count -gt 0) {
            $roleSummary = $results | Group-Object RoleName
            $typeSummary = $results | Group-Object PrincipalType
            $assignmentTypeSummary = $results | Group-Object AssignmentType
            
            Write-Host ""
            Write-Host "Role breakdown:" -ForegroundColor Cyan
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
        }
        
    }
    catch {
        Write-Error "Error in SharePoint role audit: $($_.Exception.Message)"
        
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
            Write-Host "• TermStore.ReadWrite.All" -ForegroundColor White
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
}