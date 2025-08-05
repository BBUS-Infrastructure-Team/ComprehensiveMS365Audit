# 01-CoreFunctions.ps1
# Core authentication and configuration functions for M365 Role Audit

# Configuration object for app registration credentials
$script:AppConfig = @{
    TenantId = $null
    ClientId = $null
    ClientSecret = $null
    CertificateThumbprint = $null
    UseAppAuth = $false
    AuthType = "Interactive" # Interactive, ClientSecret, Certificate
}

function Set-M365AuditCredentials {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantId,
        
        [Parameter(Mandatory = $true)]
        [string]$ClientId,
        
        [Parameter(Mandatory = $true)]
        [string]$CertificateThumbprint,

        [switch]$Quiet
    )
    
    Write-Host "Setting certificate-based authentication credentials..." -ForegroundColor Cyan
    
    # Verify certificate exists
    if ($IsWindows) {
        $cert = Get-ChildItem -Path "Cert:\CurrentUser\My" | Where-Object { $_.Thumbprint -eq $CertificateThumbprint }
        if (-not $cert) {
            $cert = Get-ChildItem -Path "Cert:\LocalMachine\My" | Where-Object { $_.Thumbprint -eq $CertificateThumbprint }
        }
    } elseIf ($IsLinux -or $IsMacOS) {
        $cert = Get-X509Certificate -Thumbprint $CertificateThumbprint
        if (-not $cert) {
            $cert = Get-X509Certificate -Thumbprint $CertificateThumbprint -Scope LocalMachine
        }
    }
    
    if (-not $cert) {
        throw "Certificate with thumbprint '$CertificateThumbprint' not found in certificate store"
    } else {
        # validate certificate expiration
        If ($cert.NotAfter -lt (Get-Date)) {
            Write-Host "Certificate has expired: $($Cert.NotAfter)" -ForegroundColor Red
            exit
        }
    }
    
    # Check certificate validity
    $isValid = (Get-Date) -ge $cert.NotBefore -and (Get-Date) -le $cert.NotAfter
    if (-not $isValid) {
        throw "Certificate is expired or not yet valid. Valid from $($cert.NotBefore) to $($cert.NotAfter)"
    }
    
    $script:AppConfig.TenantId = $TenantId
    $script:AppConfig.ClientId = $ClientId
    $script:AppConfig.CertificateThumbprint = $CertificateThumbprint
    $script:AppConfig.ClientSecret = $null
    $script:AppConfig.UseAppAuth = $true
    $script:AppConfig.AuthType = "Certificate"
    $script:AppConfig.Certificate = $cert
    
    Write-Host "✓ Certificate-based authentication configured successfully" -ForegroundColor Green
    if (-not $Quiet) {
        Write-Host "Tenant ID: $TenantId" -ForegroundColor Gray
        Write-Host "Client ID: $ClientId" -ForegroundColor Gray
        Write-Host "Certificate Thumbprint: $CertificateThumbprint" -ForegroundColor Gray
        Write-Host "Certificate Subject: $($cert.Subject)" -ForegroundColor Gray
        Write-Host "Certificate Expires: $($cert.NotAfter)" -ForegroundColor Gray
    } 
}

function Clear-M365AuditAppCredentials {
    Write-Host "Clearing application registration credentials..." -ForegroundColor Yellow
    
    $script:AppConfig.TenantId = $null
    $script:AppConfig.ClientId = $null
    $script:AppConfig.ClientSecret = $null
    $script:AppConfig.CertificateThumbprint = $null
    $script:AppConfig.UseAppAuth = $false
    $script:AppConfig.AuthType = "Interactive"
    
    Write-Host "✓ Application credentials cleared" -ForegroundColor Green
}

function Get-M365AuditCurrentConfig {
    Write-Host "=== Current M365 Audit Configuration ===" -ForegroundColor Green
    Write-Host "Authentication Type: $($script:AppConfig.AuthType)" -ForegroundColor Cyan
    Write-Host "Use App Auth: $($script:AppConfig.UseAppAuth)" -ForegroundColor Cyan
    
    if ($script:AppConfig.UseAppAuth) {
        Write-Host "Tenant ID: $($script:AppConfig.TenantId)" -ForegroundColor Gray
        Write-Host "Client ID: $($script:AppConfig.ClientId)" -ForegroundColor Gray
        
        if ($script:AppConfig.AuthType -eq "Certificate") {
            Write-Host "Certificate Thumbprint: $($script:AppConfig.CertificateThumbprint)" -ForegroundColor Gray
            
            # Verify certificate still exists and is valid
            $cert = Get-ChildItem -Path "Cert:\CurrentUser\My" | Where-Object { $_.Thumbprint -eq $script:AppConfig.CertificateThumbprint }
            if (-not $cert) {
                $cert = Get-ChildItem -Path "Cert:\LocalMachine\My" | Where-Object { $_.Thumbprint -eq $script:AppConfig.CertificateThumbprint }
            }
            
            if ($cert) {
                $isValid = (Get-Date) -ge $cert.NotBefore -and (Get-Date) -le $cert.NotAfter
                Write-Host "Certificate Status: $(if ($isValid) { 'Valid' } else { 'Expired/Invalid' })" -ForegroundColor $(if ($isValid) { 'Green' } else { 'Red' })
                Write-Host "Certificate Expires: $($cert.NotAfter)" -ForegroundColor Gray
            }
            else {
                Write-Host "Certificate Status: Not Found" -ForegroundColor Red
            }
        }
        elseif ($script:AppConfig.AuthType -eq "ClientSecret") {
            Write-Host "Client Secret: $('*' * $script:AppConfig.ClientSecret.Length)" -ForegroundColor Gray
        }
    }
    else {
        Write-Host "No application credentials configured - using interactive authentication" -ForegroundColor Yellow
    }
}

function Connect-M365ServiceWithAuth {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("Graph", "SharePoint", "Exchange", "Compliance", "PowerPlatform")]
        [string]$Service,
        
        [string]$SharePointUrl,
        [string]$AuthMethod
    )
    
    # Determine auth method if not specified
    if (-not $AuthMethod) {
        if ($script:AppConfig.UseAppAuth) {
            $AuthMethod = "Application"
        }
        else {
            $AuthMethod = "Interactive"
        }
    }
    
    try {
        switch ($Service) {
            "Graph" {
                if ($AuthMethod -eq "Application" -and $script:AppConfig.UseAppAuth) {
                    if ($script:AppConfig.AuthType -eq "Certificate") {
                        Connect-MgGraph -TenantId $script:AppConfig.TenantId -ClientId $script:AppConfig.ClientId -CertificateThumbprint $script:AppConfig.CertificateThumbprint -NoWelcome
                    }
                    else {
                        $secureSecret = ConvertTo-SecureString $script:AppConfig.ClientSecret -AsPlainText -Force
                        $credential = New-Object System.Management.Automation.PSCredential($script:AppConfig.ClientId, $secureSecret)
                        Connect-MgGraph -TenantId $script:AppConfig.TenantId -ClientSecretCredential $credential -NoWelcome
                    }
                }
                elseif ($AuthMethod -eq "DeviceCode") {
                    Connect-MgGraph -Scopes "RoleManagement.Read.All", "Directory.Read.All" -UseDeviceAuthentication
                }
                else {
                    Connect-MgGraph -Scopes "RoleManagement.Read.All", "Directory.Read.All"
                }
            }
            
            "SharePoint" {
                if ($AuthMethod -eq "Application" -and $script:AppConfig.UseAppAuth) {
                    if ($script:AppConfig.AuthType -eq "Certificate") {
                        Connect-PnPOnline -Url $SharePointUrl -ClientId $script:AppConfig.ClientId -CertificateThumbprint $script:AppConfig.CertificateThumbprint -Tenant $script:AppConfig.TenantId
                    }
                    else {
                        Connect-PnPOnline -Url $SharePointUrl -ClientId $script:AppConfig.ClientId -ClientSecret $script:AppConfig.ClientSecret -Tenant $script:AppConfig.TenantId
                    }
                }
                elseif ($AuthMethod -eq "DeviceCode") {
                    Connect-PnPOnline -Url $SharePointUrl -DeviceLogin
                }
                else {
                    Connect-PnPOnline -Url $SharePointUrl -Interactive
                }
            }
            
            "Exchange" {
                if ($AuthMethod -eq "Application" -and $script:AppConfig.UseAppAuth) {
                    if ($script:AppConfig.AuthType -eq "Certificate") {
                        Connect-ExchangeOnline -AppId $script:AppConfig.ClientId -CertificateThumbprint $script:AppConfig.CertificateThumbprint -Organization $script:AppConfig.TenantId -ShowBanner:$false
                    }
                    else {
                        Connect-ExchangeOnline -AppId $script:AppConfig.ClientId -ClientSecret (ConvertTo-SecureString $script:AppConfig.ClientSecret -AsPlainText -Force) -Organization $script:AppConfig.TenantId -ShowBanner:$false
                    }
                }
                else {
                    Connect-ExchangeOnline -ShowBanner:$false
                }
            }
            
            "Compliance" {
                if ($AuthMethod -eq "Application" -and $script:AppConfig.UseAppAuth) {
                    if ($script:AppConfig.AuthType -eq "Certificate") {
                        Connect-IPPSSession -AppId $script:AppConfig.ClientId -CertificateThumbprint $script:AppConfig.CertificateThumbprint -Organization $script:AppConfig.TenantId -ShowBanner:$false
                    }
                    else {
                        Connect-IPPSSession -AppId $script:AppConfig.ClientId -ClientSecret (ConvertTo-SecureString $script:AppConfig.ClientSecret -AsPlainText -Force) -Organization $script:AppConfig.TenantId -ShowBanner:$false
                    }
                }
                else {
                    Connect-IPPSSession -ShowBanner:$false
                }
            }
            
            "PowerPlatform" {
                # Power Platform has limited app-only support and may require interactive auth
                if ($AuthMethod -eq "Application" -and $script:AppConfig.UseAppAuth) {
                    Add-PowerAppsAccount -TenantID $script:AppConfig.TenantId
                }
                else {
                    Add-PowerAppsAccount
                }
            }
        }
        
        return $true
    }
    catch {
        Write-Warning "Failed to connect to $Service with $AuthMethod authentication: $($_.Exception.Message)"
        return $false
    }
}

function Get-M365AuditRequiredPermissions {
    Write-Host "=== Required API Permissions for M365 Role Audit App Registration ===" -ForegroundColor Green
    Write-Host ""
    
    Write-Host "Microsoft Graph API Permissions (Application):" -ForegroundColor Cyan
    Write-Host "• Directory.Read.All - Read directory data" -ForegroundColor White
    Write-Host "• RoleManagement.Read.All - Read role management data" -ForegroundColor White
    Write-Host "• User.Read.All - Read all users' full profiles" -ForegroundColor White
    Write-Host "• Sites.Read.All - Read items in all site collections (for SharePoint)" -ForegroundColor White
    Write-Host ""
    
    Write-Host "Office 365 Exchange Online API Permissions (Application):" -ForegroundColor Cyan
    Write-Host "• Exchange.ManageAsApp - Manage Exchange as application" -ForegroundColor White
    Write-Host ""
    
    Write-Host "SharePoint API Permissions (Application):" -ForegroundColor Cyan
    Write-Host "• Sites.FullControl.All - Have full control of all site collections" -ForegroundColor White
    Write-Host "• TermStore.ReadWrite.All - Read and write managed metadata" -ForegroundColor White
    Write-Host ""
    
    Write-Host "Setup Instructions:" -ForegroundColor Yellow
    Write-Host "1. Go to Azure Portal > Azure Active Directory > App registrations" -ForegroundColor White
    Write-Host "2. Select your app registration" -ForegroundColor White
    Write-Host "3. Go to 'API permissions'" -ForegroundColor White
    Write-Host "4. Click 'Add a permission'" -ForegroundColor White
    Write-Host "5. Add the permissions listed above" -ForegroundColor White
    Write-Host "6. Click 'Grant admin consent' for your organization" -ForegroundColor White
    Write-Host "7. Go to 'Certificates & secrets'" -ForegroundColor White
    Write-Host "8. Upload your certificate (.cer file) OR create a client secret" -ForegroundColor White
    Write-Host ""
    
    Write-Host "Authentication Options:" -ForegroundColor Yellow
    Write-Host "Certificate-based (Recommended):" -ForegroundColor Green
    Write-Host "• More secure than client secrets" -ForegroundColor White
    Write-Host "• No password required after setup" -ForegroundColor White
    Write-Host "• Non-exportable private key" -ForegroundColor White
    Write-Host "• Use: Set-M365AuditCertCredentials" -ForegroundColor White
    Write-Host ""
    Write-Host "Client Secret (Legacy):" -ForegroundColor Yellow
    Write-Host "• Requires secret management" -ForegroundColor White
    Write-Host "• Secrets expire and need rotation" -ForegroundColor White
    Write-Host "• Use: Set-M365AuditAppCredentials" -ForegroundColor White
    Write-Host ""
    
    Write-Host "Security Recommendations:" -ForegroundColor Yellow
    Write-Host "• Use certificate authentication instead of client secret for production" -ForegroundColor White
    Write-Host "• Limit app registration to specific IP ranges if possible" -ForegroundColor White
    Write-Host "• Regularly rotate certificates (recommended 12-24 months)" -ForegroundColor White
    Write-Host "• Monitor app usage through Azure AD audit logs" -ForegroundColor White
    Write-Host "• Use non-exportable certificates stored in Windows Certificate Store" -ForegroundColor White
}

function Initialize-M365AuditEnvironment {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$CertificateThumbprint
    )
    
    Write-Host "Setting up Microsoft 365 audit environment..." -ForegroundColor Green
    
    # Set app credentials if provided
    if ($TenantId -and $ClientId) {
        if ($CertificateThumbprint) {
            Set-M365AuditCertCredentials -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
            Write-Host "Certificate-based authentication configured" -ForegroundColor Green
        }
        elseif ($ClientSecret) {
            Set-M365AuditAppCredentials -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
            Write-Host "Client secret authentication configured" -ForegroundColor Green
        }
    }
    
    $modules = @(
        @{
            Name = 'Microsoft.Graph.Authentication'
            Version = '2.29.0'
        },
        @{
            Name = 'Microsoft.Graph.Users'
            Version = '2.29.0'
        },
        @{
            Name = 'Microsoft.Graph.Identity.DirectoryManagement'
            Version = '2.29.0'
        },
        @{
            Name = 'PnP.PowerShell'
            Version = '3.1.0'
        },
        @{
            Name = 'ExchangeOnlineManagement'
            Version = '3.8.0'
        }
    )
    
   
    foreach ($module in $modules) {
        $InstalledModule = Get-InstalledModule -Name $module.Name
        if (-not $InstalledModule) {            
            Write-Host "Installing $($Module.Name)..." -ForegroundColor Yellow
            try {
                Install-Module -Name $($module.Name) -Force -AllowClobber -Scope CurrentUser
                Write-Host "✓ $($module.Name) installed successfully" -ForegroundColor Green
            }
            catch {
                Write-Warning "Failed to install $($module.Name): $($_.Exception.Message)"
            }
        }
        else {
            if ($InstalledModule.version -le $Module.Version) {
                Write-Host "Module $($Module.Name) is not at the correct version level of $($Module.Version) or greater" -ForegroundColor Yellow
                Write-Host "Installing module $($MOdule.Name) version $)$MOdule.Version)..." -ForegroundColor Yellow
                try {
                    Install-Module -Name $Module.Name -Force -AllowClobber -Scope CurrentUser
                    Write-Host "✓ $($module.Name) installed successfully" -ForegroundColor Green
                } catch {
                    Write-Host "Failed to Install $($Module.Name): $($_.Exception.Message)"
                }
            }
            Write-Host "✓ $($module.Name) already installed" -ForegroundColor Green
        }
    }
    
    Write-Host ""
    Write-Host "Environment setup complete!" -ForegroundColor Green
    if ($script:AppConfig.UseAppAuth) {
        Write-Host "Application authentication configured and ready to use." -ForegroundColor Green
        Write-Host "Authentication Type: $($script:AppConfig.AuthType)" -ForegroundColor Cyan
    }
    else {
        Write-Host "You can now run comprehensive audits." -ForegroundColor Cyan
        Write-Host "Use Set-M365AuditCertCredentials (recommended) or Set-M365AuditAppCredentials to configure app authentication." -ForegroundColor Cyan
    }
}


function Remove-M365AuditDuplicates {
    param(
        [Parameter(Mandatory = $true)]
        [array]$AuditResults,
        
        [ValidateSet("Strict", "Loose", "ServicePreference")]
        [string]$DeduplicationMode = "ServicePreference",
        
        [switch]$ShowDuplicatesRemoved,
        [switch]$PreferAzureADSource
    )
    
    Write-Host "Deduplicating M365 audit results..." -ForegroundColor Cyan
    Write-Host "Mode: $DeduplicationMode" -ForegroundColor Gray
    
    if ($AuditResults.Count -eq 0) {
        Write-Warning "No audit results provided"
        return @()
    }
    
    $duplicatesFound = @()
    $uniqueResults = @()
    
    switch ($DeduplicationMode) {
        "Strict" {
            # Strict mode: Remove exact duplicates based on User + Role + Assignment Type
            Write-Verbose "Using strict deduplication (User + Role + Assignment Type)..."
            
            $processedKeys = @{}
            
            foreach ($result in $AuditResults) {
                # Create unique key for strict matching
                $key = "$($result.UserPrincipalName)|$($result.RoleName)|$($result.AssignmentType)"
                
                if (-not $processedKeys.ContainsKey($key)) {
                    $processedKeys[$key] = $result
                    $uniqueResults += $result
                }
                else {
                    # Found duplicate - keep track for reporting
                    $duplicatesFound += [PSCustomObject]@{
                        OriginalService = $processedKeys[$key].Service
                        DuplicateService = $result.Service
                        UserPrincipalName = $result.UserPrincipalName
                        RoleName = $result.RoleName
                        AssignmentType = $result.AssignmentType
                        Reason = "Exact duplicate"
                    }
                }
            }
        }
        
        "Loose" {
            # Loose mode: Remove duplicates based only on User + Role (ignore assignment type differences)
            Write-Verbose "Using loose deduplication (User + Role only)..."
            
            $processedKeys = @{}
            
            foreach ($result in $AuditResults) {
                # Create unique key for loose matching (ignore assignment type)
                $key = "$($result.UserPrincipalName)|$($result.RoleName)"
                
                if (-not $processedKeys.ContainsKey($key)) {
                    $processedKeys[$key] = $result
                    $uniqueResults += $result
                }
                else {
                    # Found duplicate - prefer PIM assignments over regular assignments
                    $existing = $processedKeys[$key]
                    $current = $result
                    
                    $shouldReplace = $false
                    
                    # Prefer PIM eligible over regular assignments
                    if ($current.AssignmentType -like "*Eligible*" -and $existing.AssignmentType -notlike "*Eligible*") {
                        $shouldReplace = $true
                    }
                    # Prefer PIM active over regular assignments
                    elseif ($current.AssignmentType -like "*Active (PIM*" -and $existing.AssignmentType -eq "Active") {
                        $shouldReplace = $true
                    }
                    
                    if ($shouldReplace) {
                        # Replace existing with current (better assignment type)
                        $processedKeys[$key] = $current
                        $uniqueResults = $uniqueResults | Where-Object { 
                            -not ($_.UserPrincipalName -eq $existing.UserPrincipalName -and $_.RoleName -eq $existing.RoleName) 
                        }
                        $uniqueResults += $current
                        
                        $duplicatesFound += [PSCustomObject]@{
                            OriginalService = $existing.Service
                            DuplicateService = $current.Service
                            UserPrincipalName = $current.UserPrincipalName
                            RoleName = $current.RoleName
                            AssignmentType = "$($existing.AssignmentType) -> $($current.AssignmentType)"
                            Reason = "Replaced with better assignment type"
                        }
                    }
                    else {
                        $duplicatesFound += [PSCustomObject]@{
                            OriginalService = $existing.Service
                            DuplicateService = $current.Service
                            UserPrincipalName = $current.UserPrincipalName
                            RoleName = $current.RoleName
                            AssignmentType = $current.AssignmentType
                            Reason = "Kept original assignment type"
                        }
                    }
                }
            }
        }
        
        "ServicePreference" {
            # Service preference mode: Remove duplicates with service hierarchy preference
            Write-Verbose "Using service preference deduplication..."
            
            # Define service preference hierarchy (higher number = higher preference)
            $servicePreference = @{
                "Azure AD/Entra ID" = 10          # Primary source for AD roles
                "Microsoft Intune" = 9            # Primary source for Intune roles
                "SharePoint Online" = 8           # Primary source for SharePoint roles
                "Exchange Online" = 7             # Primary source for Exchange roles
                "Microsoft Purview" = 6           # Primary source for Compliance roles
                "Microsoft Teams" = 5             # Teams roles (but often Azure AD roles)
                "Microsoft Defender" = 4          # Security roles (but often Azure AD roles)
                "Power Platform" = 3              # Power Platform roles (but often Azure AD roles)
            }
            
            $processedKeys = @{}
            
            foreach ($result in $AuditResults) {
                # Create unique key
                $key = "$($result.UserPrincipalName)|$($result.RoleName)|$($result.AssignmentType)"
                
                if (-not $processedKeys.ContainsKey($key)) {
                    $processedKeys[$key] = $result
                    $uniqueResults += $result
                }
                else {
                    # Found duplicate - check service preference
                    $existing = $processedKeys[$key]
                    $current = $result
                    
                    $existingPreference = if ($servicePreference.ContainsKey($existing.Service)) { $servicePreference[$existing.Service] } else { 1 }
                    $currentPreference = if ($servicePreference.ContainsKey($current.Service)) { $servicePreference[$current.Service] } else { 1 }
                    
                    # If current service has higher preference, replace
                    if ($currentPreference -gt $existingPreference) {
                        $processedKeys[$key] = $current
                        $uniqueResults = $uniqueResults | Where-Object { 
                            -not ($_.UserPrincipalName -eq $existing.UserPrincipalName -and 
                                  $_.RoleName -eq $existing.RoleName -and 
                                  $_.AssignmentType -eq $existing.AssignmentType) 
                        }
                        $uniqueResults += $current
                        
                        $duplicatesFound += [PSCustomObject]@{
                            OriginalService = $existing.Service
                            DuplicateService = $current.Service
                            UserPrincipalName = $current.UserPrincipalName
                            RoleName = $current.RoleName
                            AssignmentType = $current.AssignmentType
                            Reason = "Preferred service source ($($current.Service) over $($existing.Service))"
                        }
                    }
                    else {
                        $duplicatesFound += [PSCustomObject]@{
                            OriginalService = $existing.Service
                            DuplicateService = $current.Service
                            UserPrincipalName = $current.UserPrincipalName
                            RoleName = $current.RoleName
                            AssignmentType = $current.AssignmentType
                            Reason = "Kept preferred service source ($($existing.Service) over $($current.Service))"
                        }
                    }
                }
            }
        }
    }
    
    # Special handling for Azure AD preference if requested
    if ($PreferAzureADSource) {
        Write-Verbose "Applying Azure AD source preference..."
        
        # Group by user+role and prefer Azure AD source
        $groupedResults = $uniqueResults | Group-Object { "$($_.UserPrincipalName)|$($_.RoleName)" }
        $finalResults = @()
        
        foreach ($group in $groupedResults) {
            $azureADVersion = $group.Group | Where-Object { $_.Service -eq "Azure AD/Entra ID" } | Select-Object -First 1
            
            if ($azureADVersion) {
                # Use Azure AD version
                $finalResults += $azureADVersion
                
                # Track what we're removing
                $otherVersions = $group.Group | Where-Object { $_.Service -ne "Azure AD/Entra ID" }
                foreach ($other in $otherVersions) {
                    $duplicatesFound += [PSCustomObject]@{
                        OriginalService = "Azure AD/Entra ID"
                        DuplicateService = $other.Service
                        UserPrincipalName = $other.UserPrincipalName
                        RoleName = $other.RoleName
                        AssignmentType = $other.AssignmentType
                        Reason = "Preferred Azure AD as authoritative source"
                    }
                }
            }
            else {
                # No Azure AD version, keep all in group
                $finalResults += $group.Group
            }
        }
        
        $uniqueResults = $finalResults
    }
    
    if ($ShowDuplicatesRemoved -and $duplicatesFound.Count -gt 0) {
        Write-Host ""
        Write-Host "=== DUPLICATES REMOVED ===" -ForegroundColor Yellow
        $duplicatesFound | Select-Object -First 15 | Format-Table -AutoSize
        
        if ($duplicatesFound.Count -gt 15) {
            Write-Host "... and $($duplicatesFound.Count - 15) more duplicates removed" -ForegroundColor Gray
        }
        
        # Summary by service
        Write-Host "Duplicates by service:" -ForegroundColor Cyan
        $duplicatesByService = $duplicatesFound | Group-Object DuplicateService | Sort-Object Count -Descending
        foreach ($service in $duplicatesByService) {
            Write-Host "  $($service.Name): $($service.Count) duplicates removed" -ForegroundColor White
        }
    }
    
    return $uniqueResults
}

# Helper function to get report metadata
function Get-ReportMetadata {
    param([string]$OrganizationName, [hashtable]$Stats, [array]$AuditResults)
    
    return @{
        organizationName = $OrganizationName
        generatedDate = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
        auditVersion = "2.1"
        reportType = "Comprehensive"
        totalAssignments = $Stats.totalAssignments
        uniqueUsers = $Stats.uniqueUsers
        servicesAudited = $Stats.servicesAudited
        certificateAuthUsed = ($Stats.authTypes | Where-Object { $_.Name -eq "Certificate" }).Count -gt 0
        pimEnabled = $Stats.pimEligible.Count -gt 0
        hybridEnvironmentDetected = $Stats.onPremSynced.Count -gt 0
        exchangeDataEnhanced = $Stats.exchangeResults.Count -gt 0
    }
}

# Helper function to get report summary
function Get-ReportSummary {
    param([array]$AuditResults, [hashtable]$Stats)
    
    return @{
        serviceBreakdown = @($AuditResults | Group-Object Service | ForEach-Object {
            @{
                service = $_.Name
                count = $_.Count
                percentage = [math]::Round(($_.Count / $Stats.totalAssignments) * 100, 2)
            }
        })
        
        topRoles = @($AuditResults | Group-Object RoleName | Sort-Object Count -Descending | Select-Object -First 15 | ForEach-Object {
            @{
                roleName = $_.Name
                assignmentCount = $_.Count
                riskLevel = Get-RoleRiskLevel -RoleName $_.Name
                services = @($_.Group | Group-Object Service | ForEach-Object { $_.Name })
            }
        })
        
        usersWithMostRoles = @($AuditResults | Group-Object UserPrincipalName | Sort-Object Count -Descending | Select-Object -First 15 | ForEach-Object {
            if ($_.Name) {
                $userInfo = $_.Group | Select-Object -First 1
                @{
                    userPrincipalName = $_.Name
                    displayName = $userInfo.DisplayName
                    roleCount = $_.Count
                    isEnabled = $userInfo.UserEnabled
                    lastSignIn = $userInfo.LastSignIn
                    services = @($_.Group | Group-Object Service | ForEach-Object { $_.Name })
                    onPremisesSynced = $userInfo.OnPremisesSyncEnabled -eq $true
                }
            }
        } | Where-Object { $_ })
        
        assignmentTypes = @($AuditResults | Group-Object AssignmentType | ForEach-Object {
            @{
                type = $_.Name
                count = $_.Count
                percentage = [math]::Round(($_.Count / $Stats.totalAssignments) * 100, 2)
            }
        })
    }
}

# Helper function for service analysis
function Get-ServiceAnalysis {
    param([array]$AuditResults, [switch]$IncludeExchangeAnalysis)
    
    $serviceStats = @{}
    
    foreach ($service in ($AuditResults | Group-Object Service)) {
        $serviceData = $service.Group
        $serviceStats[$service.Name] = @{
            totalAssignments = $service.Count
            uniqueUsers = ($serviceData | Where-Object { $_.UserPrincipalName } | Select-Object -Unique UserPrincipalName).Count
            topRole = ($serviceData | Group-Object RoleName | Sort-Object Count -Descending | Select-Object -First 1).Name
            authMethods = @($serviceData | Group-Object AuthenticationType | ForEach-Object {
                @{ method = $_.Name; count = $_.Count }
            })
        }
        
        # Add service-specific analysis
        if ($service.Name -eq "Exchange Online" -and $IncludeExchangeAnalysis) {
            $serviceStats[$service.Name].exchangeAnalysis = Get-ExchangeSpecificAnalysis -ExchangeData $serviceData
        }
        elseif ($service.Name -eq "Microsoft Intune") {
            $serviceStats[$service.Name].intuneAnalysis = Get-IntuneSpecificAnalysis -IntuneData $serviceData
        }
        elseif ($service.Name -eq "SharePoint Online") {
            $serviceStats[$service.Name].sharePointAnalysis = Get-SharePointSpecificAnalysis -SharePointData $serviceData
        }
    }
    
    return $serviceStats
}

# Helper function for Exchange-specific analysis
function Get-ExchangeSpecificAnalysis {
    param([array]$ExchangeData)
    
    $roleGroups = $ExchangeData | Where-Object { $_.AssignmentType -eq "Role Group Member" }
    $azureADRoles = $ExchangeData | Where-Object { $_.RoleSource -eq "AzureAD" }
    $onPremSynced = $ExchangeData | Where-Object { $_.OnPremisesSyncEnabled -eq $true }
    $groups = $ExchangeData | Where-Object { $_.PrincipalType -eq "Group" }
    
    return @{
        roleGroupAssignments = $roleGroups.Count
        azureADRoleAssignments = $azureADRoles.Count
        onPremisesSyncedUsers = $onPremSynced.Count
        groupAssignments = $groups.Count
        hybridEnvironment = $onPremSynced.Count -gt 0
        orgManagementMembers = ($ExchangeData | Where-Object { $_.RoleName -eq "Organization Management" }).Count
        securityAdministrators = ($ExchangeData | Where-Object { $_.RoleName -eq "Security Administrator" }).Count
        crossServiceSyncedGroups = ($ExchangeData | Where-Object { 
            $_.RoleGroupDescription -like "*synchronized across services*" 
        }).Count
    }
}

# Helper function for Intune-specific analysis
function Get-IntuneSpecificAnalysis {
    param([array]$IntuneData)
    
    return @{
        rbacAssignments = ($IntuneData | Where-Object { $_.RoleType -eq "IntuneRBAC" }).Count
        azureADAssignments = ($IntuneData | Where-Object { $_.RoleType -eq "AzureAD" }).Count
        builtInRoles = ($IntuneData | Where-Object { $_.IsBuiltIn -eq $true }).Count
        customRoles = ($IntuneData | Where-Object { $_.IsBuiltIn -eq $false }).Count
        serviceAdministrators = ($IntuneData | Where-Object { $_.RoleName -eq "Intune Service Administrator" }).Count
    }
}

# Helper function for SharePoint-specific analysis
function Get-SharePointSpecificAnalysis {
    param([array]$SharePointData)
    
    return @{
        uniqueSites = ($SharePointData | Where-Object { $_.SiteTitle } | Select-Object -Unique SiteTitle).Count
        totalStorageMB = ($SharePointData | Where-Object { $_.StorageUsedMB } | Measure-Object StorageUsedMB -Sum).Sum
        siteCollectionAdmins = ($SharePointData | Where-Object { $_.RoleName -like "*Site*Administrator*" }).Count
        appCatalogAdmins = ($SharePointData | Where-Object { $_.RoleName -eq "App Catalog Administrator" }).Count
    }
}

# Helper function for PIM analysis
function Get-PIMAnalysis {
    param([array]$AuditResults, [switch]$IncludeDetailedAnalysis)
    
    $pimEligible = $AuditResults | Where-Object { $_.AssignmentType -like "*Eligible*" }
    $pimActive = $AuditResults | Where-Object { $_.AssignmentType -like "*Active (PIM*" }
    
    $analysis = @{
        enabled = $pimEligible.Count -gt 0
        totalEligible = $pimEligible.Count
        totalActive = $pimActive.Count
    }
    
    if ($IncludeDetailedAnalysis) {
        $analysis.detailed = @{
            eligible = @{
                total = $pimEligible.Count
                byService = @($pimEligible | Group-Object Service | ForEach-Object {
                    @{ service = $_.Name; count = $_.Count }
                })
                expiringWithin30Days = @($pimEligible | Where-Object { 
                    $_.PIMEndDateTime -and [datetime]$_.PIMEndDateTime -lt (Get-Date).AddDays(30) 
                }).Count
            }
            active = @{
                total = $pimActive.Count
                byService = @($pimActive | Group-Object Service | ForEach-Object {
                    @{ service = $_.Name; count = $_.Count }
                })
            }
        }
    }
    
    return $analysis
}

# Helper function for principal analysis
function Get-PrincipalAnalysis {
    param([array]$AuditResults)
    
    return @{
        users = ($AuditResults | Where-Object { $_.PrincipalType -eq "User" -or (!$_.PrincipalType -and $_.UserPrincipalName -like "*@*") } | Select-Object -Unique UserPrincipalName).Count
        servicePrincipals = ($AuditResults | Where-Object { $_.PrincipalType -eq "ServicePrincipal" }).Count
        groups = ($AuditResults | Where-Object { $_.PrincipalType -eq "Group" }).Count
        onPremisesSyncedUsers = ($AuditResults | Where-Object { $_.OnPremisesSyncEnabled -eq $true }).Count
        totalGroupMembers = ($AuditResults | Where-Object { 
            $_.PrincipalType -eq "Group" -and $_.GroupMemberCount -and $_.GroupMemberCount -ne "Unknown" 
        } | ForEach-Object { 
            try { [int]$_.GroupMemberCount } catch { 0 } 
        } | Measure-Object -Sum).Sum
    }
}

# Helper function for cross-service analysis
function Get-CrossServiceAnalysis {
    param([array]$AuditResults)
    
    $multiServiceUsers = $AuditResults | Where-Object { 
        $_.UserPrincipalName -and $_.UserPrincipalName -ne "Unknown" 
    } | Group-Object UserPrincipalName | Where-Object {
        ($_.Group | Group-Object Service).Count -gt 1
    }
    
    return @{
        usersWithMultipleServices = $multiServiceUsers.Count
        exchangeAzureADCombinations = ($multiServiceUsers | Where-Object {
            $userServices = ($_.Group | Group-Object Service).Name
            $userServices -contains "Exchange Online" -and $userServices -contains "Azure AD/Entra ID"
        }).Count
        exchangePurviewCombinations = ($multiServiceUsers | Where-Object {
            $userServices = ($_.Group | Group-Object Service).Name
            $userServices -contains "Exchange Online" -and $userServices -contains "Microsoft Purview"
        }).Count
    }
}

# Helper function for security alerts
function Get-SecurityAlerts {
    param([array]$AuditResults, [hashtable]$Stats)
    
    $alerts = @{
        critical = @()
        high = @()
        medium = @()
        low = @()
        globalAdminCount = $Stats.globalAdmins.Count
        disabledUsersWithRoles = $Stats.disabledUsers.Count
        certificateBasedAuth = ($Stats.authTypes | Where-Object { $_.Name -eq "Certificate" }).Count
        clientSecretAuth = ($Stats.authTypes | Where-Object { $_.Name -eq "ClientSecret" }).Count
    }
    
    # Generate alerts based on findings
    if ($Stats.globalAdmins.Count -gt 5) {
        $alerts.critical += "Excessive Global Administrators: $($Stats.globalAdmins.Count) accounts (recommended: ≤5)"
    }
    
    if ($Stats.disabledUsers.Count -gt 0) {
        $alerts.high += "Disabled users with active roles: $($Stats.disabledUsers.Count) accounts need review"
    }
    
    if ($alerts.clientSecretAuth -gt 0) {
        $alerts.medium += "Client secret authentication detected - migrate to certificate-based"
    }
    
    if ($Stats.pimEligible.Count -eq 0 -and $Stats.totalAssignments -gt 0) {
        $alerts.medium += "No PIM eligible assignments - consider implementing privileged access management"
    }
    
    # Exchange-specific alerts
    if ($Stats.exchangeResults.Count -gt 0) {
        $exchangeAlerts = Get-ExchangeSecurityAlerts -ExchangeData $Stats.exchangeResults
        $alerts.exchangeSecurityAlerts = $exchangeAlerts
        
        if ($exchangeAlerts.orgManagementMembers -gt 10) {
            $alerts.medium += "High number of Organization Management members: $($exchangeAlerts.orgManagementMembers)"
        }
    }
    
    return $alerts
}

# Helper function for recommendations
function Get-SecurityRecommendations {
    param([array]$AuditResults, [hashtable]$Stats)
    
    $recommendations = @{
        immediate = @()
        shortTerm = @()
        longTerm = @()
    }
    
    # Immediate recommendations
    if ($Stats.globalAdmins.Count -gt 5) {
        $recommendations.immediate += "Reduce Global Administrator count to 5 or fewer"
    }
    
    if ($Stats.disabledUsers.Count -gt 0) {
        $recommendations.immediate += "Remove role assignments from $($Stats.disabledUsers.Count) disabled user accounts"
    }
    
    # Short-term recommendations
    if ($Stats.pimEligible.Count -eq 0) {
        $recommendations.shortTerm += "Implement Privileged Identity Management (PIM) for eligible assignments"
    }
    
    if (($Stats.authTypes | Where-Object { $_.Name -eq "ClientSecret" }).Count -gt 0) {
        $recommendations.shortTerm += "Migrate from client secret to certificate-based authentication"
    }
    
    # Long-term recommendations
    $recommendations.longTerm += "Implement regular access reviews for privileged roles"
    $recommendations.longTerm += "Monitor privileged role assignments with automated alerts"
    $recommendations.longTerm += "Establish break-glass emergency access procedures"
    
    # Exchange-specific recommendations
    if ($Stats.exchangeResults.Count -gt 0) {
        $recommendations.longTerm += "Implement Exchange role group governance and regular membership reviews"
        
        if ($Stats.onPremSynced.Count -gt 0) {
            $recommendations.longTerm += "Ensure hybrid identity governance for on-premises synced users"
        }
    }
    
    return $recommendations
}

# Helper function for compliance analysis
function Get-ComplianceAnalysis {
    param([array]$AuditResults, [hashtable]$Stats)
    
    return @{
        privilegedAccessCompliance = @{
            globalAdminLimit = @{
                compliant = $Stats.globalAdmins.Count -le 5
                current = $Stats.globalAdmins.Count
                recommended = 5
            }
            pimImplementation = @{
                compliant = $Stats.pimEligible.Count -gt 0
                eligibleCount = $Stats.pimEligible.Count
            }
            disabledAccountCleanup = @{
                compliant = $Stats.disabledUsers.Count -eq 0
                violationCount = $Stats.disabledUsers.Count
            }
        }
        authenticationCompliance = @{
            certificateBasedAuth = @{
                compliant = ($Stats.authTypes | Where-Object { $_.Name -eq "Certificate" }).Count -gt 0
                percentage = [math]::Round((($Stats.authTypes | Where-Object { $_.Name -eq "Certificate" }).Count / $Stats.totalAssignments) * 100, 2)
            }
        }
    }
}

# Helper function to format assignments
function Get-FormattedAssignments {
    param([array]$AuditResults)
    
    return @($AuditResults | ForEach-Object {
        $assignment = @{
            service = $_.Service
            userPrincipalName = $_.UserPrincipalName
            displayName = $_.DisplayName
            roleName = $_.RoleName
            assignmentType = $_.AssignmentType
            assignedDateTime = $_.AssignedDateTime
            userEnabled = $_.UserEnabled
            authenticationType = $_.AuthenticationType
        }
        
        # Add optional fields if present
        @('PIMEndDateTime', 'RoleSource', 'PrincipalType', 'OnPremisesSyncEnabled', 'GroupMemberCount', 
          'RoleGroupDescription', 'OrganizationalUnit', 'ManagementScope', 'RecipientType', 'SiteTitle', 
          'Template', 'RoleType', 'PolicyType', 'IsBuiltIn') | ForEach-Object {
            if ($null -ne $_.$_) { $assignment.$_ = $_.$_ }
        }
        
        $assignment
    })
}

# Helper function for role risk assessment
<# function Get-RoleRiskLevel {
    param([string]$RoleName)
    
    switch -Regex ($RoleName) {
        "Global Administrator|Company Administrator" { return "CRITICAL" }
        "Security Administrator|Exchange Administrator|SharePoint Administrator|Intune Service Administrator" { return "HIGH" }
        ".*Administrator.*|.*Admin.*" { return "MEDIUM" }
        default { return "LOW" }
    }
} #>

# Helper function to show report summary
function Show-ReportSummary {
    param([hashtable]$Report, [string]$OutputPath)
    
    $fileSize = [math]::Round((Get-Item $OutputPath).Length / 1KB, 2)
    Write-Host "File size: $fileSize KB" -ForegroundColor Gray
    Write-Host "Total assignments: $($Report.metadata.totalAssignments)" -ForegroundColor Gray
    Write-Host "Services included: $($Report.metadata.servicesAudited)" -ForegroundColor Gray
    Write-Host "Security alerts: $($Report.securityAlerts.critical.Count) critical, $($Report.securityAlerts.high.Count) high" -ForegroundColor Gray
    
    if ($Report.metadata.hybridEnvironmentDetected) {
        Write-Host "✓ Hybrid environment detected and analyzed" -ForegroundColor Green
    }
    
    if ($Report.metadata.pimEnabled) {
        Write-Host "✓ PIM assignments detected" -ForegroundColor Green
    }
}

# Consolidated Get-AuditStatistics function for 01-CoreFunctions.ps1
function Get-AuditStatistics {
    param(
        [Parameter(Mandatory = $true)]
        [array]$AuditResults,
        [switch]$IncludeDetailedAnalysis
    )
    
    if ($AuditResults.Count -eq 0) {
        return @{
            totalAssignments = 0
            uniqueUsers = 0
            servicesAudited = 0
            authTypes = @()
            globalAdmins = @()
            disabledUsers = @()
            pimEligible = @()
            pimActive = @()
            permanentActive = @()
            onPremSynced = @()
            exchangeResults = @()
        }
    }
    
    # Basic statistics
    $totalAssignments = $AuditResults.Count
    $uniqueUsers = ($AuditResults | Where-Object { $_.UserPrincipalName -and $_.UserPrincipalName -ne "Unknown" } |
                   Select-Object -Unique UserPrincipalName).Count
    $servicesAudited = ($AuditResults | Group-Object Service).Count
    $authenticationTypes = $AuditResults | Group-Object AuthenticationType
    
    # Security analysis
    $globalAdmins = $AuditResults | Where-Object { $_.RoleName -eq "Global Administrator" }
    $disabledUsers = $AuditResults | Where-Object { $_.UserEnabled -eq $false }
    $privilegedRoles = $AuditResults | Where-Object {
        $_.RoleName -match "Administrator|Admin" -and
        $_.RoleName -ne "Global Administrator" -and
        $_.RoleName -notlike "*Policy*" -and
        $_.RoleName -notlike "*Configuration*"
    }
    $securityRoles = $AuditResults | Where-Object { $_.RoleName -match "Security" }
    $complianceRoles = $AuditResults | Where-Object { $_.RoleName -match "Compliance|eDiscovery|DLP" }
    
    # PIM analysis
    $pimEligible = $AuditResults | Where-Object { $_.AssignmentType -like "*Eligible*" }
    $pimActive = $AuditResults | Where-Object { $_.AssignmentType -like "*Active (PIM*" }
    $permanentActive = $AuditResults | Where-Object { 
        $_.AssignmentType -eq "Active" -or 
        $_.AssignmentType -eq "Azure AD Role" -or
        $_.AssignmentType -eq "Intune RBAC" -or
        $_.AssignmentType -eq "Role Group Member"
    }
    $timeBoundAssignments = $AuditResults | Where-Object { $_.AssignmentType -eq "Time-bound RBAC" }
    
    # User analysis
    $usersWithoutRecentSignIn = $AuditResults | Where-Object {
        $_.LastSignIn -and $_.LastSignIn -lt (Get-Date).AddDays(-90)
    }
    $systemGeneratedAccounts = $AuditResults | Where-Object {
        $_.UserPrincipalName -eq "System Generated" -or $_.DisplayName -like "*System*" -or $_.DisplayName -like "*Policy*"
    }
    $onPremSynced = $AuditResults | Where-Object { $_.OnPremisesSyncEnabled -eq $true }
    
    # Service-specific results
    $exchangeResults = $AuditResults | Where-Object { $_.Service -eq "Exchange Online" }
    $sharePointResults = $AuditResults | Where-Object { $_.Service -eq "SharePoint Online" }
    $intuneResults = $AuditResults | Where-Object { $_.Service -eq "Microsoft Intune" }
    $teamsResults = $AuditResults | Where-Object { $_.Service -eq "Microsoft Teams" }
    $purviewResults = $AuditResults | Where-Object { $_.Service -eq "Microsoft Purview" }
    $defenderResults = $AuditResults | Where-Object { $_.Service -eq "Microsoft Defender" }
    $powerPlatformResults = $AuditResults | Where-Object { $_.Service -eq "Power Platform" }
    $azureADResults = $AuditResults | Where-Object { $_.Service -eq "Azure AD/Entra ID" }
    
    # Build base statistics object - maintains backward compatibility with all existing versions
    $stats = @{
        # Basic metrics (compatible with all versions)
        totalAssignments = $totalAssignments
        uniqueUsers = $uniqueUsers
        servicesAudited = $servicesAudited
        authTypes = $authenticationTypes
        
        # Security metrics (compatible with existing version)
        globalAdmins = $globalAdmins
        disabledUsers = $disabledUsers
        privilegedRoles = $privilegedRoles
        securityRoles = $securityRoles
        complianceRoles = $complianceRoles
        
        # PIM metrics (compatible with existing version)
        pimEligible = $pimEligible
        pimActive = $pimActive
        permanentActive = $permanentActive
        timeBoundAssignments = $timeBoundAssignments
        
        # User metrics (compatible with existing version)
        usersWithoutRecentSignIn = $usersWithoutRecentSignIn
        systemGeneratedAccounts = $systemGeneratedAccounts
        onPremSynced = $onPremSynced
        
        # Service-specific results (compatible with existing version)
        exchangeResults = $exchangeResults
        
        # Enhanced service results (new, additive only)
        sharePointResults = $sharePointResults
        intuneResults = $intuneResults
        teamsResults = $teamsResults
        purviewResults = $purviewResults
        defenderResults = $defenderResults
        powerPlatformResults = $powerPlatformResults
        azureADResults = $azureADResults
    }
    
    # Add detailed analysis if requested
    if ($IncludeDetailedAnalysis) {
        # Principal type analysis
        $principalTypes = $AuditResults | Where-Object { $_.PrincipalType } | Group-Object PrincipalType
        $stats.principalTypes = $principalTypes
        
        # Role distribution analysis
        $topRoles = $AuditResults | Group-Object RoleName | Sort-Object Count -Descending | Select-Object -First 15
        $stats.topRoles = $topRoles
        
        # User role distribution
        $usersWithMostRoles = $AuditResults | Where-Object { $_.UserPrincipalName -and $_.UserPrincipalName -ne "Unknown" } |
                             Group-Object UserPrincipalName | Sort-Object Count -Descending | Select-Object -First 15
        $stats.usersWithMostRoles = $usersWithMostRoles
        
        # Service distribution
        $serviceDistribution = $AuditResults | Group-Object Service | Sort-Object Count -Descending
        $stats.serviceDistribution = $serviceDistribution
        
        # Assignment type distribution
        $assignmentTypes = $AuditResults | Group-Object AssignmentType
        $stats.assignmentTypes = $assignmentTypes
        
        # Cross-service user analysis
        $crossServiceUsers = $AuditResults | Where-Object { 
            $_.UserPrincipalName -and $_.UserPrincipalName -ne "Unknown" -and $_.UserPrincipalName -ne "System Generated"
        } | Group-Object UserPrincipalName | Where-Object {
            ($_.Group | Group-Object Service).Count -gt 1
        }
        $stats.crossServiceUsers = $crossServiceUsers
        
        # Hybrid environment indicators
        $stats.hybridEnvironmentDetected = $onPremSynced.Count -gt 0
        $stats.groupAssignments = ($AuditResults | Where-Object { $_.PrincipalType -eq "Group" }).Count
        $stats.servicePrincipalAssignments = ($AuditResults | Where-Object { $_.PrincipalType -eq "ServicePrincipal" }).Count
        
        # PIM adoption analysis by service
        $pimByService = @{}
        foreach ($service in $serviceDistribution) {
            $serviceData = $service.Group
            $servicePIMEligible = $serviceData | Where-Object { $_.AssignmentType -like "*Eligible*" }
            $servicePermanent = $serviceData | Where-Object { 
                $_.AssignmentType -eq "Active" -or $_.AssignmentType -eq "Azure AD Role" -or $_.AssignmentType -eq "Role Group Member"
            }
            
            $adoptionRate = if (($servicePIMEligible.Count + $servicePermanent.Count) -gt 0) {
                [math]::Round(($servicePIMEligible.Count / ($servicePIMEligible.Count + $servicePermanent.Count)) * 100, 2)
            } else { 0 }
            
            $pimByService[$service.Name] = @{
                eligible = $servicePIMEligible.Count
                permanent = $servicePermanent.Count
                adoptionRate = $adoptionRate
            }
        }
        $stats.pimByService = $pimByService
        
        # Certificate usage analysis
        $certificateAuth = $authenticationTypes | Where-Object { $_.Name -eq "Certificate" }
        $clientSecretAuth = $authenticationTypes | Where-Object { $_.Name -eq "ClientSecret" }
        $stats.certificateAuthUsage = if ($certificateAuth) { $certificateAuth.Count } else { 0 }
        $stats.clientSecretAuthUsage = if ($clientSecretAuth) { $clientSecretAuth.Count } else { 0 }
    }
    
    return $stats
}



# Helper function to create compliance gap objects
function New-ComplianceGap {
    param(
        [string]$Category,
        [string]$Issue,
        [string]$Details,
        [ValidateSet("Critical", "High", "Medium", "Low")]
        [string]$Severity,
        [string]$Recommendation,
        [object]$AffectedUsers,
        [string]$ComplianceFramework,
        [array]$RemediationSteps
    )
    
    # Convert AffectedUsers to string if it's an array
    $affectedUsersString = if ($AffectedUsers -is [array]) {
        $AffectedUsers -join "; "
    } else {
        $AffectedUsers
    }
    
    return [PSCustomObject]@{
        Category = $Category
        Issue = $Issue
        Details = $Details
        Severity = $Severity
        Recommendation = $Recommendation
        AffectedUsers = $affectedUsersString
        ComplianceFramework = $ComplianceFramework
        RemediationSteps = $RemediationSteps
    }
}

# Note: This function is no longer needed as we now use the consolidated Get-AuditStatistics 
# from 01-CoreFunctions.ps1 which provides all the same data and more

# Helper function for Intune-specific compliance gaps
function Get-IntuneComplianceGaps {
    param([array]$AuditResults)
    
    $gaps = @()
    $intuneResults = $AuditResults | Where-Object { $_.Service -eq "Microsoft Intune" }
    
    if ($intuneResults.Count -gt 0) {
        # Check Intune Service Administrator count
        $intuneServiceAdmins = $intuneResults | Where-Object { $_.RoleName -eq "Intune Service Administrator" }
        if ($intuneServiceAdmins.Count -gt 3) {
            $gaps += New-ComplianceGap -Category "Device Management" `
                                      -Issue "Excessive Intune Service Administrators" `
                                      -Details "$($intuneServiceAdmins.Count) Intune Service Administrators (recommended: ≤3)" `
                                      -Severity "Medium" `
                                      -Recommendation "Use Intune RBAC roles for granular permissions instead of broad service administrator role" `
                                      -AffectedUsers ($intuneServiceAdmins | Select-Object -ExpandProperty UserPrincipalName) `
                                      -ComplianceFramework "Device Security, NIST" `
                                      -RemediationSteps @(
                                          "1. Review Intune administrative requirements",
                                          "2. Implement Intune RBAC roles for specific functions",
                                          "3. Remove unnecessary Service Administrator roles",
                                          "4. Train admins on scoped permissions"
                                      )
        }
        
        # Check RBAC vs Azure AD role usage
        $intuneRBACAssignments = $intuneResults | Where-Object { $_.RoleType -eq "IntuneRBAC" }
        $intuneAzureADAssignments = $intuneResults | Where-Object { $_.RoleType -eq "AzureAD" }
        
        if ($intuneAzureADAssignments.Count -gt $intuneRBACAssignments.Count -and $intuneResults.Count -gt 10) {
            $gaps += New-ComplianceGap -Category "Device Management" `
                                      -Issue "Underutilized Intune RBAC" `
                                      -Details "$($intuneAzureADAssignments.Count) Azure AD roles vs $($intuneRBACAssignments.Count) Intune RBAC roles" `
                                      -Severity "Low" `
                                      -Recommendation "Leverage Intune RBAC for more granular, scope-specific permissions" `
                                      -AffectedUsers "Intune administrators" `
                                      -ComplianceFramework "Least Privilege" `
                                      -RemediationSteps @(
                                          "1. Map current Azure AD roles to Intune RBAC equivalents",
                                          "2. Create custom Intune roles for specific needs",
                                          "3. Migrate to Intune RBAC where appropriate",
                                          "4. Implement scope-based assignments"
                                      )
        }
        
        # Check for policy ownership and management
        $intunePolicyOwners = $intuneResults | Where-Object { $_.RoleType -eq "PolicyOwner" }
        if ($intunePolicyOwners.Count -eq 0) {
            $gaps += New-ComplianceGap -Category "Device Management" `
                                      -Issue "No Policy Ownership Tracking" `
                                      -Details "No Intune policy ownership information found" `
                                      -Severity "Low" `
                                      -Recommendation "Implement policy ownership tracking and governance" `
                                      -AffectedUsers "Policy administrators" `
                                      -ComplianceFramework "Change Management" `
                                      -RemediationSteps @(
                                          "1. Document policy ownership and responsibilities",
                                          "2. Implement policy change approval process",
                                          "3. Regular review of policy configurations",
                                          "4. Track policy creation and modification"
                                      )
        }
    }
    
    return $gaps
}

# Helper function for Power Platform-specific compliance gaps
function Get-PowerPlatformComplianceGaps {
    param([array]$AuditResults)
    
    $gaps = @()
    $powerPlatformResults = $AuditResults | Where-Object { $_.Service -eq "Power Platform" }
    
    if ($powerPlatformResults.Count -gt 0) {
        # Check for service principals with Power Platform access
        $servicePrincipals = $powerPlatformResults | Where-Object { $_.PrincipalType -eq "ServicePrincipal" }
        if ($servicePrincipals.Count -gt 0) {
            $spNames = ($servicePrincipals | Select-Object -ExpandProperty DisplayName -Unique) -join "; "
            $gaps += New-ComplianceGap -Category "Application Security" `
                                      -Issue "Service Principals with Power Platform Access" `
                                      -Details "$($servicePrincipals.Count) service principals have Power Platform administrative access" `
                                      -Severity "Medium" `
                                      -Recommendation "Review and validate service principal access to Power Platform resources" `
                                      -AffectedUsers $spNames `
                                      -ComplianceFramework "Application Security" `
                                      -RemediationSteps @(
                                          "1. Review each service principal's business justification",
                                          "2. Validate minimum required permissions",
                                          "3. Implement managed identities where possible",
                                          "4. Regular audit of application permissions"
                                      )
        }
        
        # Check Power Platform administrator count
        $powerPlatformAdmins = $powerPlatformResults | Where-Object { $_.RoleName -eq "Power Platform Administrator" }
        if ($powerPlatformAdmins.Count -gt 5) {
            $gaps += New-ComplianceGap -Category "Power Platform Governance" `
                                      -Issue "Excessive Power Platform Administrators" `
                                      -Details "$($powerPlatformAdmins.Count) Power Platform Administrators (consider environment-specific roles)" `
                                      -Severity "Medium" `
                                      -Recommendation "Use environment-specific admin roles instead of tenant-wide Power Platform Administrator" `
                                      -AffectedUsers ($powerPlatformAdmins | Select-Object -ExpandProperty UserPrincipalName) `
                                      -ComplianceFramework "Least Privilege" `
                                      -RemediationSteps @(
                                          "1. Review Power Platform administrative requirements",
                                          "2. Implement environment-specific roles",
                                          "3. Use DLP policies for governance",
                                          "4. Regular review of platform usage"
                                      )
        }
    }
    
    return $gaps
}

# Helper function to show compliance gap summary
function Show-ComplianceGapSummary {
    param(
        [array]$Gaps,
        [array]$CriticalGaps,
        [array]$HighGaps,
        [array]$MediumGaps,
        [array]$LowGaps
    )
    
    Write-Host "Total Gaps Found: $($Gaps.Count)" -ForegroundColor White
    Write-Host "  Critical: $($CriticalGaps.Count)" -ForegroundColor Red
    Write-Host "  High: $($HighGaps.Count)" -ForegroundColor Red
    Write-Host "  Medium: $($MediumGaps.Count)" -ForegroundColor Yellow
    Write-Host "  Low: $($LowGaps.Count)" -ForegroundColor Cyan
}

# Helper function to show detailed compliance gaps
function Show-DetailedComplianceGaps {
    param(
        [array]$CriticalGaps,
        [array]$HighGaps,
        [array]$MediumGaps,
        [array]$LowGaps
    )
    
    Write-Host ""
    Write-Host "=== DETAILED GAP ANALYSIS ===" -ForegroundColor Cyan
    
    if ($CriticalGaps.Count -gt 0) {
        Write-Host "CRITICAL GAPS:" -ForegroundColor Red
        foreach ($gap in $CriticalGaps) {
            Write-Host "  ⚠️ $($gap.Issue): $($gap.Details)" -ForegroundColor White
            Write-Host "     Recommendation: $($gap.Recommendation)" -ForegroundColor Gray
        }
    }
    
    if ($HighGaps.Count -gt 0) {
        Write-Host "HIGH PRIORITY GAPS:" -ForegroundColor Red
        foreach ($gap in $HighGaps) {
            Write-Host "  ⚠️ $($gap.Issue): $($gap.Details)" -ForegroundColor White
            Write-Host "     Recommendation: $($gap.Recommendation)" -ForegroundColor Gray
        }
    }
    
    if ($MediumGaps.Count -gt 0) {
        Write-Host "MEDIUM PRIORITY GAPS:" -ForegroundColor Yellow
        foreach ($gap in $MediumGaps) {
            Write-Host "  • $($gap.Issue): $($gap.Details)" -ForegroundColor White
            Write-Host "    Recommendation: $($gap.Recommendation)" -ForegroundColor Gray
        }
    }
    
    if ($LowGaps.Count -gt 0) {
        Write-Host "LOW PRIORITY GAPS:" -ForegroundColor Cyan
        foreach ($gap in $LowGaps) {
            Write-Host "  • $($gap.Issue): $($gap.Details)" -ForegroundColor White
            Write-Host "    Recommendation: $($gap.Recommendation)" -ForegroundColor Gray
        }
    }
}

# Helper function to show compliance framework impact
function Show-ComplianceFrameworkImpact {
    param([array]$Gaps)
    
    Write-Host ""
    Write-Host "=== COMPLIANCE FRAMEWORK IMPACT ===" -ForegroundColor Cyan
    $frameworkImpact = $Gaps | ForEach-Object { $_.ComplianceFramework -split ", " } | 
                      Group-Object | Sort-Object Count -Descending
    
    foreach ($framework in $frameworkImpact) {
        Write-Host "  $($framework.Name): $($framework.Count) gaps" -ForegroundColor White
    }
}

# Helper function to show priority recommendations
function Show-PriorityRecommendations {
    param(
        [array]$CriticalGaps,
        [array]$HighGaps
    )
    
    Write-Host ""
    Write-Host "Priority Remediation Recommendations:" -ForegroundColor Yellow
    $priorityRecommendations = @()
    $priorityRecommendations += $CriticalGaps | ForEach-Object { $_.Recommendation }
    $priorityRecommendations += $HighGaps | ForEach-Object { $_.Recommendation }
    
    $priorityRecommendations | Select-Object -Unique | ForEach-Object {
        Write-Host "• $_" -ForegroundColor White
    }
}

# Helper function for Exchange-specific security alerts
function Get-ExchangeSecurityAlerts {
    param([array]$ExchangeData)
    
    return @{
        orgManagementMembers = ($ExchangeData | Where-Object { $_.RoleName -eq "Organization Management" }).Count
        securityAdministrators = ($ExchangeData | Where-Object { $_.RoleName -eq "Security Administrator" }).Count
        disabledUsersWithExchangeRoles = ($ExchangeData | Where-Object { $_.UserEnabled -eq $false }).Count
        groupAssignments = ($ExchangeData | Where-Object { $_.PrincipalType -eq "Group" }).Count
        onPremisesSyncedUsers = ($ExchangeData | Where-Object { $_.OnPremisesSyncEnabled -eq $true }).Count
    }
}

# Helper function for SharePoint-specific analysis
function Get-SharePointSpecificAnalysis {
    param([array]$SharePointData)
    
    return @{
        uniqueSites = ($SharePointData | Where-Object { $_.SiteTitle } | Select-Object -Unique SiteTitle).Count
        totalStorageMB = ($SharePointData | Where-Object { $_.StorageUsedMB } | Measure-Object StorageUsedMB -Sum).Sum
        siteCollectionAdmins = ($SharePointData | Where-Object { $_.RoleName -like "*Site*Administrator*" }).Count
        appCatalogAdmins = ($SharePointData | Where-Object { $_.RoleName -eq "App Catalog Administrator" }).Count
        termStoreAccess = ($SharePointData | Where-Object { $_.RoleName -eq "Term Store Access Verified" }).Count
        searchCenterAdmins = ($SharePointData | Where-Object { $_.RoleName -eq "Search Center Administrator" }).Count
    }
}

# Enhanced helper function for role risk assessment
function Get-RoleRiskLevel {
    param([string]$RoleName)
    
    switch -Regex ($RoleName) {
        "Global Administrator|Company Administrator" { return "CRITICAL" }
        "Security Administrator|Exchange Administrator|SharePoint Administrator|Intune Service Administrator|Power Platform Administrator" { return "HIGH" }
        ".*Administrator.*|.*Admin.*" { return "MEDIUM" }
        ".*Reader.*|.*Viewer.*" { return "LOW" }
        default { return "LOW" }
    }
}

# Helper function to get role category
function Get-RoleCategory {
    param([string]$RoleName, [string]$Service)
    
    # Service-specific categorization
    switch ($Service) {
        "Azure AD/Entra ID" {
            switch -Regex ($RoleName) {
                "Global Administrator" { return "Global" }
                ".*Security.*" { return "Security" }
                ".*User.*|.*Guest.*" { return "Identity" }
                ".*Application.*|.*App.*" { return "Application" }
                default { return "Administrative" }
            }
        }
        "Microsoft Intune" {
            switch -Regex ($RoleName) {
                ".*Policy.*" { return "PolicyManagement" }
                ".*Device.*" { return "DeviceManagement" }
                ".*App.*" { return "ApplicationManagement" }
                default { return "DeviceAdministration" }
            }
        }
        "Exchange Online" {
            switch -Regex ($RoleName) {
                "Organization Management" { return "FullExchangeAdmin" }
                ".*Security.*" { return "SecurityCompliance" }
                ".*Recipient.*" { return "RecipientManagement" }
                ".*Transport.*|.*Mail.*" { return "TransportManagement" }
                default { return "ExchangeAdmin" }
            }
        }
        "SharePoint Online" {
            switch -Regex ($RoleName) {
                ".*Site.*" { return "SiteManagement" }
                ".*App.*" { return "AppManagement" }
                ".*Search.*" { return "SearchManagement" }
                default { return "SharePointAdmin" }
            }
        }
        default { return "ServiceSpecific" }
    }
}

# Helper function for advanced PIM analysis
function Get-AdvancedPIMAnalysis {
    param([array]$AuditResults)
    
    $pimEligible = $AuditResults | Where-Object { $_.AssignmentType -like "*Eligible*" }
    $pimActive = $AuditResults | Where-Object { $_.AssignmentType -like "*Active (PIM*" }
    $timeBound = $AuditResults | Where-Object { $_.AssignmentType -eq "Time-bound RBAC" }
    
    # Analyze PIM by service
    $pimByService = @{}
    foreach ($service in ($AuditResults | Group-Object Service)) {
        $serviceEligible = $service.Group | Where-Object { $_.AssignmentType -like "*Eligible*" }
        $serviceActive = $service.Group | Where-Object { $_.AssignmentType -like "*Active (PIM*" }
        $servicePermanent = $service.Group | Where-Object { 
            $_.AssignmentType -eq "Active" -or $_.AssignmentType -eq "Azure AD Role" -or $_.AssignmentType -eq "Role Group Member"
        }
        
        $pimByService[$service.Name] = @{
            eligible = $serviceEligible.Count
            active = $serviceActive.Count
            permanent = $servicePermanent.Count
            adoptionRate = if (($serviceEligible.Count + $servicePermanent.Count) -gt 0) {
                [math]::Round(($serviceEligible.Count / ($serviceEligible.Count + $servicePermanent.Count)) * 100, 2)
            } else { 0 }
        }
    }
    
    # Check for expiring assignments
    $expiringPIM = $AuditResults | Where-Object { 
        $_.PIMEndDateTime -and [DateTime]$_.PIMEndDateTime -lt (Get-Date).AddDays(30)
    }
    
    return @{
        totalEligible = $pimEligible.Count
        totalActive = $pimActive.Count
        totalTimeBound = $timeBound.Count
        byService = $pimByService
        expiringWithin30Days = $expiringPIM.Count
        overallAdoptionRate = if (($pimEligible.Count + ($AuditResults | Where-Object { $_.AssignmentType -eq "Active" }).Count) -gt 0) {
            [math]::Round(($pimEligible.Count / ($pimEligible.Count + ($AuditResults | Where-Object { $_.AssignmentType -eq "Active" }).Count)) * 100, 2)
        } else { 0 }
    }
}

function Get-RoleUsers {
    param(
        [array]$AuditResults,
        [string]$RoleName
    )
    
    # Get all users assigned to a specific role, return with DisplayName
    $roleAssignments = $AuditResults | Where-Object { $_.RoleName -eq $RoleName }
    
    $users = @()
    foreach ($assignment in $roleAssignments) {
        $displayName = if ($assignment.DisplayName -and $assignment.DisplayName -ne "Unknown") {
            $assignment.DisplayName
        } elseif ($assignment.UserPrincipalName -and $assignment.UserPrincipalName -ne "Unknown") {
            $assignment.UserPrincipalName
        } else {
            "Unknown User"
        }
        
        $users += @{
            displayName = $displayName
            userPrincipalName = $assignment.UserPrincipalName
            assignmentType = $assignment.AssignmentType
        }
    }
    
    # Return unique users only (avoid duplicates)
    return $users | Sort-Object displayName | Select-Object -Unique displayName, userPrincipalName, assignmentType
}

function Get-UserRoles {
    param(
        [array]$AuditResults,
        [string]$UserPrincipalName
    )
    
    # Get all roles assigned to a specific user
    $userAssignments = $AuditResults | Where-Object { 
        $_.UserPrincipalName -eq $UserPrincipalName 
    }
    
    $roles = @()
    foreach ($assignment in $userAssignments) {
        $roles += @{
            Service = $assignment.Service
            RoleName = $assignment.RoleName
            AssignmentType = $assignment.AssignmentType
        }
    }
    
    return $roles | Sort-Object Service, RoleName
}

function Get-ServicePIMCounts {
    param(
        [array]$AuditResults,
        [string]$ServiceName
    )
    
    # Calculate PIM Active and PIM Eligible counts for a specific service
    $serviceResults = $AuditResults | Where-Object { $_.Service -eq $ServiceName }
    
    $pimActive = ($serviceResults | Where-Object { 
        $_.AssignmentType -like "*Active (PIM*" 
    }).Count
    
    $pimEligible = ($serviceResults | Where-Object { 
        $_.AssignmentType -like "*Eligible*" 
    }).Count
    
    return @{
        pimActive = $pimActive
        pimEligible = $pimEligible
    }
}

function Get-RoleAssignmentsForService {
    param(
        [Parameter(Mandatory = $true)]
        [array]$RoleDefinitions,
        
        [Parameter(Mandatory = $true)]
        [string]$ServiceName,
        
        [switch]$IncludePIM,
        [switch]$Quiet
    )
    
    if (-not $Quiet) {
        Write-Host "Retrieving all $ServiceName assignment types..." -ForegroundColor Cyan
    }
    
    # Initialize collections
    $activeAssignments = @()
    $pimEligibleAssignments = @()
    $pimActiveAssignments = @()
    
    # Get active assignments (permanent assignments)
    if (-not $Quiet) {
        Write-Host "Getting active $ServiceName assignments..." -ForegroundColor Gray
    }

    $allAssignments = Get-MgRoleManagementDirectoryRoleAssignment -All
    $ActiveAssignments = $allAssignments | Where-Object { $_.RoleDefinitionId -in $RoleDefinitions.id}
    
<#     foreach ($roleId in $RoleDefinitions.Id) {
        try {
            $assignments = Get-MgRoleManagementDirectoryRoleAssignment -Filter "roleDefinitionId eq '$roleId'" -ErrorAction SilentlyContinue
            if ($assignments) {
                $activeAssignments += $assignments
            }
        }
        catch {
            Write-Verbose "Error getting active assignments for role $roleId`: $($_.Exception.Message)"
        }
    } #>
    
    if (-not $Quiet) {
        Write-Host "Found $($activeAssignments.Count) active assignments" -ForegroundColor Green
    }
    
    # Get PIM assignments if requested
    if ($IncludePIM) {
        # Get PIM eligible assignments
        try {

            $allPimAssignments = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -All
            $pimEligibleAssignments = $allPimAssignments | Where-Object {$_.RoleDefinitionId -in $RoleDefinitions.Id}

            if (-not $Quiet) {
                Write-Host "Getting PIM eligible $ServiceName assignments..." -ForegroundColor Gray
            }
            
<#             foreach ($roleId in $RoleDefinitions.Id) {
                try {
                    $pimEligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -Filter "roleDefinitionId eq '$roleId'" -ErrorAction SilentlyContinue
                    if ($pimEligible) {
                        $pimEligibleAssignments += $pimEligible
                    }
                }
                catch {
                    Write-Verbose "Error getting PIM eligible assignments for role $roleId`: $($_.Exception.Message)"
                }
            } #>
            
            if (-not $Quiet) {
                Write-Host "Found $($pimEligibleAssignments.Count) PIM eligible assignments" -ForegroundColor Green
            }
        }
        catch {
            if (-not $Quiet) {
                Write-Host "Could not retrieve PIM eligible assignments (may not be licensed)" -ForegroundColor Yellow
            }
        }
        
        # Get PIM active assignments
        try {
            if (-not $Quiet) {
                Write-Host "Getting PIM active $ServiceName assignments..." -ForegroundColor Gray
            }

            $allPimActiveAssignments = Get-MgRoleManagementDirectoryRoleAssignmentSchedule -All

            $pimActiveAssignments = $allPimActiveAssignments | Where-Object{ $_.RoleDefinitionId -in $RoleDefinitions.Id}
            
<#             foreach ($roleId in $RoleDefinitions.Id) {
                try {
                    $pimActive = Get-MgRoleManagementDirectoryRoleAssignmentSchedule -Filter "roleDefinitionId eq '$roleId'" -ErrorAction SilentlyContinue
                    if ($pimActive) {
                        $pimActiveAssignments += $pimActive
                    }
                }
                catch {
                    Write-Verbose "Error getting PIM active assignments for role $roleId`: $($_.Exception.Message)"
                }
            }
             #>
            if (-not $Quiet) {
                Write-Host "Found $($pimActiveAssignments.Count) PIM active assignments" -ForegroundColor Green
            }
        }
        catch {
            if (-not $Quiet) {
                Write-Host "Could not retrieve PIM active assignments (may not be licensed)" -ForegroundColor Yellow
            }
        }
    }
    
    # ========= CRITICAL FIX: PROPER DEDUPLICATION LOGIC =========
    # The issue is that PIM active assignments can appear in both regular active assignments 
    # AND PIM active assignments. We need to deduplicate properly.
    
    # Create a hashtable to track unique assignments by a composite key
    $uniqueAssignments = @{}
    $duplicateCount = 0
    
    # Add active assignments first
    foreach ($assignment in $activeAssignments) {
        $key = "$($assignment.PrincipalId)|$($assignment.RoleDefinitionId)|$($assignment.DirectoryScopeId)"
        
        if (-not $uniqueAssignments.ContainsKey($key)) {
            $assignment | Add-Member -NotePropertyName "AssignmentSource" -NotePropertyValue "Active" -Force
            $uniqueAssignments[$key] = $assignment
        } else {
            $duplicateCount++
            Write-Verbose "Duplicate active assignment found: $key"
        }
    }
    
    # Add PIM eligible assignments (these should be unique from active)
    foreach ($assignment in $pimEligibleAssignments) {
        $key = "$($assignment.PrincipalId)|$($assignment.RoleDefinitionId)|$($assignment.DirectoryScopeId)"
        
        if (-not $uniqueAssignments.ContainsKey($key)) {
            $assignment | Add-Member -NotePropertyName "AssignmentSource" -NotePropertyValue "PIMEligible" -Force
            $uniqueAssignments[$key] = $assignment
        } else {
            $duplicateCount++
            Write-Verbose "Duplicate PIM eligible assignment found: $key"
        }
    }
    
    # Add PIM active assignments - these might overlap with regular active assignments
    foreach ($assignment in $pimActiveAssignments) {
        $key = "$($assignment.PrincipalId)|$($assignment.RoleDefinitionId)|$($assignment.DirectoryScopeId)"
        
        if (-not $uniqueAssignments.ContainsKey($key)) {
            # New assignment - add as PIM active
            $assignment | Add-Member -NotePropertyName "AssignmentSource" -NotePropertyValue "PIMActive" -Force
            $uniqueAssignments[$key] = $assignment
        } else {
            # Duplicate found - prefer PIM active over regular active
            $existing = $uniqueAssignments[$key]
            if ($existing.AssignmentSource -eq "Active") {
                # Replace regular active with PIM active (more specific)
                $assignment | Add-Member -NotePropertyName "AssignmentSource" -NotePropertyValue "PIMActive" -Force
                $uniqueAssignments[$key] = $assignment
                $duplicateCount++
                Write-Verbose "Replaced active assignment with PIM active: $key"
            } else {
                # Keep existing (eligible or already PIM active)
                $duplicateCount++
                Write-Verbose "Duplicate PIM active assignment found: $key"
            }
        }
    }
    
    # Convert hashtable values back to array
    $allAssignments = @($uniqueAssignments.Values)
    
    if (-not $Quiet) {
        Write-Host "Total $ServiceName assignments after deduplication: $($allAssignments.Count)" -ForegroundColor Green
        if ($duplicateCount -gt 0) {
            Write-Host "  Removed $duplicateCount duplicate assignments during processing" -ForegroundColor Yellow
        }
        
        # Show breakdown by assignment source
        $sourceBreakdown = $allAssignments | Group-Object AssignmentSource | Sort-Object Name
        Write-Host "  Assignment source breakdown:" -ForegroundColor Cyan
        foreach ($source in $sourceBreakdown) {
            Write-Host "    $($source.Name): $($source.Count)" -ForegroundColor White
        }
    }
    
    return $allAssignments
}

function ConvertTo-ServiceAssignmentResults {
    param(
        [Parameter(Mandatory = $true)]
        [array]$Assignments,
        
        [Parameter(Mandatory = $true)]
        [array]$RoleDefinitions,
        
        [Parameter(Mandatory = $true)]
        [string]$ServiceName,
        
        [Parameter(Mandatory = $true)]
        [array]$OverarchingRoles,
        
        [string]$DefaultAssignmentType = "Undefined",
        [bool]$IncludeAllPrincipalTypes = $true,
        [switch]$IncludeUnknownPrincipals,
        [switch]$Quiet
    )
    
    $results = @()
    
    if ($Assignments.Count -eq 0) {
        if (-not $Quiet) { Write-Verbose "No assignments provided for $ServiceName" }
        return $results
    }
    
    if (-not $Quiet) {
        Write-Host "Processing $($Assignments.Count) $ServiceName assignments with optimized batch filtering..." -ForegroundColor Cyan
    }
    
    # Create role definition lookup hashtable for performance
    $roleDefinitionHash = @{}
    foreach ($roleDef in $RoleDefinitions) {
        $roleDefinitionHash[$roleDef.Id] = $roleDef
    }
    
    # Get unique principal IDs and resolve them using your optimized batch approach
    $uniquePrincipalIds = @($Assignments | Select-Object -ExpandProperty PrincipalId -Unique)
    if (-not $Quiet) {
        Write-Host "  Resolving $($uniquePrincipalIds.Count) unique principals using your optimized batch filtering..." -ForegroundColor Gray
    }
    
    # Initialize principal cache and counters
    $principalCache = @{}
    $userResolvedCount = 0
    $groupResolvedCount = 0
    $servicePrincipalResolvedCount = 0
    $unknownPrincipalCount = 0
    
    $batchSize = 15
    
    # ========= USER RESOLUTION USING YOUR EXACT PATTERN =========
    if (-not $Quiet) {
        Write-Host "    Batch resolving users..." -ForegroundColor Gray
    }
    
    $allUsers = for ($i = 0; $i -lt $uniquePrincipalIds.Count; $i += $batchSize) {
        $batch = $uniquePrincipalIds[$i..([Math]::Min($i + $batchSize - 1, $uniquePrincipalIds.Count - 1))]
        $filter = "id in ('" + ($batch -join "','") + "')"
        Get-MgUser -Filter $filter -Property "Id,UserPrincipalName,DisplayName,AccountEnabled,OnPremisesSyncEnabled" -ErrorAction SilentlyContinue
    }
    
    # Cache all resolved users
    foreach ($user in $allUsers) {
        if ($user) {
            $principalCache[$user.Id] = @{
                Type = "User"
                UserPrincipalName = $user.UserPrincipalName
                DisplayName = $user.DisplayName
                AccountEnabled = $user.AccountEnabled
                OnPremisesSyncEnabled = $user.OnPremisesSyncEnabled
            }
            $userResolvedCount++
        }
    }
    
    # ========= GROUP RESOLUTION USING YOUR EXACT PATTERN =========
    if ($IncludeAllPrincipalTypes) {
        $unresolvedPrincipals = $uniquePrincipalIds | Where-Object { -not $principalCache.ContainsKey($_) }
        
        if ($unresolvedPrincipals.Count -gt 0) {
            if (-not $Quiet) {
                Write-Host "    Batch resolving $($unresolvedPrincipals.Count) remaining principals as groups..." -ForegroundColor Gray
            }
            
            $allGroups = for ($i = 0; $i -lt $unresolvedPrincipals.Count; $i += $batchSize) {
                $batch = $unresolvedPrincipals[$i..([Math]::Min($i + $batchSize - 1, $unresolvedPrincipals.Count - 1))]
                $filter = "id in ('" + ($batch -join "','") + "')"
                Get-MgGroup -Filter $filter -Property "Id,Mail,DisplayName,OnPremisesSyncEnabled" -ErrorAction SilentlyContinue
            }
            
            # Cache all resolved groups
            foreach ($group in $allGroups) {
                if ($group) {
                    $principalCache[$group.Id] = @{
                        Type = "Group"
                        UserPrincipalName = if ($group.Mail) { $group.Mail } else { $group.DisplayName }
                        DisplayName = "$($group.DisplayName) (Group)"
                        AccountEnabled = $null
                        OnPremisesSyncEnabled = $group.OnPremisesSyncEnabled
                    }
                    $groupResolvedCount++
                }
            }
        }
        
        # ========= SERVICE PRINCIPAL RESOLUTION USING YOUR EXACT PATTERN =========
        $unresolvedPrincipals = $uniquePrincipalIds | Where-Object { -not $principalCache.ContainsKey($_) }
        
        if ($unresolvedPrincipals.Count -gt 0) {
            if (-not $Quiet) {
                Write-Host "    Batch resolving $($unresolvedPrincipals.Count) remaining principals as service principals..." -ForegroundColor Gray
            }
            
            $allServicePrincipals = for ($i = 0; $i -lt $unresolvedPrincipals.Count; $i += $batchSize) {
                $batch = $unresolvedPrincipals[$i..([Math]::Min($i + $batchSize - 1, $unresolvedPrincipals.Count - 1))]
                $filter = "id in ('" + ($batch -join "','") + "')"
                Get-MgServicePrincipal -Filter $filter -Property "Id,AppId,DisplayName,AccountEnabled" -ErrorAction SilentlyContinue
            }
            
            # Cache all resolved service principals
            foreach ($sp in $allServicePrincipals) {
                if ($sp) {
                    $principalCache[$sp.Id] = @{
                        Type = "ServicePrincipal"
                        UserPrincipalName = $sp.AppId
                        DisplayName = "$($sp.DisplayName) (Application)"
                        AccountEnabled = $sp.AccountEnabled
                        OnPremisesSyncEnabled = $false
                    }
                    $servicePrincipalResolvedCount++
                }
            }
        }
    }
    
    # Mark remaining principals as unknown
    $unresolvedPrincipals = $uniquePrincipalIds | Where-Object { -not $principalCache.ContainsKey($_) }
    foreach ($principalId in $unresolvedPrincipals) {
        $principalCache[$principalId] = @{
            Type = "Unknown"
            UserPrincipalName = "Unknown-$principalId"
            DisplayName = "Unknown Principal"
            AccountEnabled = $null
            OnPremisesSyncEnabled = $null
        }
        $unknownPrincipalCount++
    }
    
    # ========= OPTIMIZE: FILTER PRINCIPAL CACHE UPFRONT INSTEAD OF IN LOOP =========
    $originalCacheSize = $principalCache.Count
    $unknownPrincipalFilteredCount = 0
    $nonUserPrincipalFilteredCount = 0
    
    # Filter unknown principals if not including them
    if (-not $IncludeUnknownPrincipals) {
        $unknownPrincipals = $principalCache.GetEnumerator() | Where-Object { $_.Value.Type -eq "Unknown" }
        $unknownPrincipalFilteredCount = $unknownPrincipals.Count
        
        foreach ($unknownPrincipal in $unknownPrincipals) {
            $principalCache.Remove($unknownPrincipal.Key)
        }
    }
    
    # Filter non-users if not including all principal types
    if (-not $IncludeAllPrincipalTypes) {
        $nonUserPrincipals = $principalCache.GetEnumerator() | Where-Object { $_.Value.Type -ne "User" }
        $nonUserPrincipalFilteredCount = $nonUserPrincipals.Count
        
        foreach ($nonUserPrincipal in $nonUserPrincipals) {
            $principalCache.Remove($nonUserPrincipal.Key)
        }
    }
    
    if (-not $Quiet) {
        Write-Host "  ✓ Principal resolution completed - Users: $userResolvedCount, Groups: $groupResolvedCount, SPs: $servicePrincipalResolvedCount, Unknown: $unknownPrincipalCount" -ForegroundColor Green
        Write-Host "  ✓ Cache optimized - Original: $originalCacheSize, Filtered: $($principalCache.Count) (Removed: Unknown=$unknownPrincipalFilteredCount, NonUser=$nonUserPrincipalFilteredCount)" -ForegroundColor Cyan
    }
    
    # ========= FAST ASSIGNMENT PROCESSING USING FILTERED CACHED PRINCIPALS =========
    $processedCount = 0
    
    foreach ($assignment in $Assignments) {
        try {
            $processedCount++
            
            # Progress indicator (less frequent for performance)
            if (-not $Quiet -and $processedCount % 50 -eq 0) {
                Write-Host "  Processed $processedCount of $($Assignments.Count) assignments..." -ForegroundColor Gray
            }
            
            # Get role definition (cached lookup)
            $role = $roleDefinitionHash[$assignment.RoleDefinitionId]
            if (-not $role) {
                Write-Verbose "Unknown role definition: $($assignment.RoleDefinitionId)"
                continue
            }
            
            # Get principal info from filtered cache (no API calls needed)
            $principalInfo = $principalCache[$assignment.PrincipalId]
            if (-not $principalInfo) {
                # Principal was filtered out or doesn't exist - skip silently
                continue
            }
            
            # Determine assignment type
            $assignmentType = switch ($assignment.AssignmentSource) {
                "Active" { "Active" }
                "PIMEligible" { "Eligible (PIM)" }
                "PIMActive" { "Active (PIM)" }
                default { $DefaultAssignmentType }
            }
            
            # Determine role scope
            $roleScope = if ($role.DisplayName -in $OverarchingRoles) { "Overarching" } else { "Service-Specific" }
            
            # Create result object
            $results += [PSCustomObject]@{
                Service = $ServiceName
                UserPrincipalName = $principalInfo.UserPrincipalName
                DisplayName = $principalInfo.DisplayName
                UserId = $assignment.PrincipalId
                RoleName = $role.DisplayName
                RoleDefinitionId = $assignment.RoleDefinitionId
                RoleScope = $roleScope
                AssignmentType = $assignmentType
                AssignedDateTime = $assignment.CreatedDateTime
                UserEnabled = $principalInfo.AccountEnabled
                Scope = $assignment.DirectoryScopeId
                AssignmentId = $assignment.Id
                PrincipalType = $principalInfo.Type
                OnPremisesSyncEnabled = $principalInfo.OnPremisesSyncEnabled
                PIMStartDateTime = if ($assignment.ScheduleInfo) { $assignment.ScheduleInfo.StartDateTime } else { $null }
                PIMEndDateTime = if ($assignment.ScheduleInfo -and $assignment.ScheduleInfo.Expiration) { $assignment.ScheduleInfo.Expiration.EndDateTime } else { $null }
            }
            
        }
        catch {
            Write-Warning "Error processing $ServiceName assignment $($assignment.Id): $($_.Exception.Message)"
            continue
        }
    }
    
    # Final summary
    if (-not $Quiet) {
        Write-Host "✓ $ServiceName assignment processing completed with your optimized batch filtering" -ForegroundColor Green
        Write-Host "  Results: $($results.Count) assignments processed" -ForegroundColor White
        Write-Host "  Performance: Used batch principal resolution for $($uniquePrincipalIds.Count) unique principals" -ForegroundColor Cyan
        
        if ($unknownPrincipalFilteredCount -gt 0) {
            Write-Host "  Filtered out: $unknownPrincipalFilteredCount unknown principals" -ForegroundColor Green
        }
        
        if ($results.Count -gt 0) {
            $assignmentTypeBreakdown = $results | Group-Object AssignmentType
            Write-Host "  Assignment types:" -ForegroundColor Cyan
            foreach ($type in $assignmentTypeBreakdown) {
                Write-Host "    $($type.Name): $($type.Count)" -ForegroundColor White
            }
        }
        
        if ($unknownPrincipalFilteredCount -gt 10) {
            Write-Host ""
            Write-Host "💡 RECOMMENDATION: Consider cleaning up $unknownPrincipalFilteredCount orphaned role assignments" -ForegroundColor Yellow
            Write-Host "   Use -IncludeUnknownPrincipals to see the full list for cleanup" -ForegroundColor White
        }
    }
    
    return $results
}

function Get-RoleGroupMemberResult () {
    Param (
        [psObject]$Member,
        [string]$Service,
        [psObject]$RoleGroup
    )

    $recipientType = $member.RecipientType
    $isUser = $recipientType -in @("UserMailbox", "MailUser", "User")
    $isGroup = $recipientType -in @("MailUniversalSecurityGroup", "UniversalSecurityGroup", "MailUniversalDistributionGroup", "Group")
    
    if ($isUser -or $isGroup) {
        $principalType = if ($isUser) { "User" } else { "Group" } 
        
        # Try to get additional user info from Graph for consistency
        $userEnabled = $null
        #$lastSignIn = $null
        $onPremisesSyncEnabled = $null
        
        if ($isUser) {
            try {
                # First we have to get the user ID because we are using a filter
                $userId = (Get-MgUser -Filter "DisplayName eq '$($member.DisplayName)'").Id
                # Now we can get the graph user and the properties we need using the user ID
                if ($userId) {
                    $graphUser = Get-MgUser -UserId $UserId -Property Id, UserPrincipalName, AccountEnabled, OnPremisesSyncEnabled -ErrorAction SilentlyContinue
                }
                if ($graphUser) {
                    $userEnabled = $graphUser.AccountEnabled
                    $userPrincipalName = $graphUser.UserPrincipalName
                    #$lastSignIn = $graphUser.SignInActivity.LastSignInDateTime
                    $onPremisesSyncEnabled = $graphUser.OnPremisesSyncEnabled
                } else {
                    $userPrincipalName = "Unable to retrieve Entra ID user"
                }
            }
            catch {
                Write-Verbose "Could not retrieve Graph data for compliance user $($member.PrimarySmtpAddress): $($_.Exception.Message)"
            }
        } 
        
        $result = [PSCustomObject]@{
            Service = $Service
            UserPrincipalName = $userPrincipalName
            DisplayName = $member.DisplayName
            UserId = $UserId
            RoleName = $roleGroup.Name
            RoleDefinitionId = $roleGroup.Guid
            RoleScope = "Service-Specific"  # Compliance role groups are service-specific
            AssignmentType = "Role Group Member"
            AssignedDateTime = $null
            UserEnabled = $userEnabled
            Scope = "Organization"
            AssignmentId = $roleGroup.Identity
            PrincipalType = $principalType
            OnPremisesSyncEnabled = $onPremisesSyncEnabled
            PIMStartDateTime = $null
            PIMEndDateTime = $null
        }
    }
    if ($null -eq $result) {
            Throw "Cannot identify Principal! Name: $($member.Name), RecipientType: $($member.RecipientType)"
    }
    return $result
}