# 08-Testing-Troubleshooting-Functions.ps1
# Testing and troubleshooting functions for M365 Role Audit (Certificate Authentication Only)
# Fixed version - removed unused variables

function Test-M365AuditConnections {
    param(
        [string]$SharePointTenantUrl = "https://balfourbeattyus-admin.sharepoint.com",
        [string]$TenantId,
        [string]$ClientId,
        [string]$CertificateThumbprint
    )
    
    # Set app credentials if provided
    if ($TenantId -and $ClientId -and $CertificateThumbprint) {
        Set-M365AuditCertCredentials -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
    }
    
    Write-Host "=== Testing M365 Service Connections ===" -ForegroundColor Green
    if ($script:AppConfig.UseAppAuth) {
        Write-Host "Using Certificate-Based Authentication" -ForegroundColor Cyan
        Write-Host "Tenant ID: $($script:AppConfig.TenantId)" -ForegroundColor Gray
        Write-Host "Client ID: $($script:AppConfig.ClientId)" -ForegroundColor Gray
        Write-Host "Certificate Thumbprint: $($script:AppConfig.CertificateThumbprint)" -ForegroundColor Gray
        
        # Validate certificate
        $cert = Get-ChildItem -Path "Cert:\CurrentUser\My" | Where-Object { $_.Thumbprint -eq $script:AppConfig.CertificateThumbprint }
        if (-not $cert) {
            $cert = Get-ChildItem -Path "Cert:\LocalMachine\My" | Where-Object { $_.Thumbprint -eq $script:AppConfig.CertificateThumbprint }
        }
        
        if ($cert) {
            $isValid = (Get-Date) -ge $cert.NotBefore -and (Get-Date) -le $cert.NotAfter
            $status = if ($isValid) { "Valid" } else { "Expired/Invalid" }
            $color = if ($isValid) { "Green" } else { "Red" }
            Write-Host "Certificate Status: $status" -ForegroundColor $color
            Write-Host "Certificate Expires: $($cert.NotAfter)" -ForegroundColor Gray
        }
        else {
            Write-Host "Certificate Status: Not Found" -ForegroundColor Red
        }
    }
    else {
        Write-Host "Using Interactive Authentication" -ForegroundColor Cyan
    }
    Write-Host ""
    
    $testResults = @{}
    $authMethod = if ($script:AppConfig.UseAppAuth) { "Application" } else { "Interactive" }
    
    # Test Microsoft Graph
    Write-Host "Testing Microsoft Graph connection..." -ForegroundColor Yellow
    try {
        $connectionSuccess = Connect-M365ServiceWithAuth -Service "Graph" -AuthMethod $authMethod
        
        if ($connectionSuccess) {
            $testUser = Get-MgUser -Top 1
            if ($testUser) {
                Write-Host "✓ Microsoft Graph: Connected and working" -ForegroundColor Green
                $testResults.Graph = $true
                
                # Test context details
                $context = Get-MgContext
                Write-Host "  Auth Type: $($context.AuthType)" -ForegroundColor Gray
                Write-Host "  Scopes: $($context.Scopes -join ', ')" -ForegroundColor Gray
            }
        }
        else {
            throw "Connection failed"
        }
    }
    catch {
        Write-Host "✗ Microsoft Graph: Failed - $($_.Exception.Message)" -ForegroundColor Red
        $testResults.Graph = $false
        
        if ($script:AppConfig.UseAppAuth) {
            Write-Host "  Check: Certificate uploaded to app registration" -ForegroundColor Yellow
            Write-Host "  Check: Graph API permissions granted" -ForegroundColor Yellow
        }
    }
    
    # Test SharePoint
    Write-Host "Testing SharePoint connection..." -ForegroundColor Yellow
    try {
        $connectionSuccess = Connect-M365ServiceWithAuth -Service "SharePoint" -SharePointUrl $SharePointTenantUrl -AuthMethod $authMethod
        
        if ($connectionSuccess) {
            $sites = Get-PnPTenantSite -ErrorAction Stop | Select-Object -First 1
            if ($sites) {
                Write-Host "✓ SharePoint: Connected and working" -ForegroundColor Green
                $testResults.SharePoint = $true
                
                # Test connection details
                $connection = Get-PnPConnection
                Write-Host "  Connection Type: $($connection.ConnectionType)" -ForegroundColor Gray
                Write-Host "  URL: $($connection.Url)" -ForegroundColor Gray
            }
        }
        else {
            throw "Connection failed"
        }
        
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
    }
    catch {
        Write-Host "✗ SharePoint: Failed - $($_.Exception.Message)" -ForegroundColor Red
        $testResults.SharePoint = $false
        
        if ($script:AppConfig.UseAppAuth) {
            Write-Host "  Check: Certificate registered with SharePoint app permissions" -ForegroundColor Yellow
        }
    }
    
    # Test Exchange Online
    Write-Host "Testing Exchange Online connection..." -ForegroundColor Yellow
    try {
        $connectionSuccess = Connect-M365ServiceWithAuth -Service "Exchange" -AuthMethod $authMethod
        
        if ($connectionSuccess) {
            $org = Get-OrganizationConfig -ErrorAction Stop
            if ($org) {
                Write-Host "✓ Exchange Online: Connected and working" -ForegroundColor Green
                $testResults.Exchange = $true
                Write-Host "  Organization: $($org.DisplayName)" -ForegroundColor Gray
            }
        }
        else {
            throw "Connection failed"
        }
        
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }
    catch {
        Write-Host "✗ Exchange Online: Failed - $($_.Exception.Message)" -ForegroundColor Red
        $testResults.Exchange = $false
        
        if ($script:AppConfig.UseAppAuth) {
            Write-Host "  Check: Certificate has Exchange.ManageAsApp permission" -ForegroundColor Yellow
        }
    }
    
    # Test Security & Compliance Center
    Write-Host "Testing Security & Compliance Center connection..." -ForegroundColor Yellow
    try {
        $connectionSuccess = Connect-M365ServiceWithAuth -Service "Compliance" -AuthMethod $authMethod
        
        if ($connectionSuccess) {
            # Fixed: Use returned value from Get-RoleGroup to validate connection
            $complianceRoles = Get-RoleGroup -ErrorAction Stop | Select-Object -First 1
            if ($complianceRoles) {
                Write-Host "✓ Security & Compliance Center: Connected and working" -ForegroundColor Green
                $testResults.Compliance = $true
            }
        }
        else {
            throw "Connection failed"
        }
    }
    catch {
        Write-Host "✗ Security & Compliance Center: Failed - $($_.Exception.Message)" -ForegroundColor Red
        $testResults.Compliance = $false
        
        if ($script:AppConfig.UseAppAuth) {
            Write-Host "  Check: Certificate has compliance permissions" -ForegroundColor Yellow
        }
    }
    
    # Test Power Platform (Windows PowerShell 5.x only)
    if ($PSVersionTable.PSVersion.Major -eq 5) {
        Write-Host "Testing Power Platform connection..." -ForegroundColor Yellow
        try {
            $connectionSuccess = Connect-M365ServiceWithAuth -Service "PowerPlatform" -AuthMethod $authMethod
            
            if ($connectionSuccess) {
                $powerPlatformEnvs = Get-AdminPowerAppEnvironment -ErrorAction Stop | Select-Object -First 1
                if ($powerPlatformEnvs) {
                    Write-Host "✓ Power Platform: Connected and working" -ForegroundColor Green
                    $testResults.PowerPlatform = $true
                    
                    if ($script:AppConfig.UseAppAuth) {
                        Write-Host "  Note: Power Platform has limited certificate support" -ForegroundColor Yellow
                    }
                }
            }
            else {
                throw "Connection failed"
            }
        }
        catch {
            Write-Host "✗ Power Platform: Failed - $($_.Exception.Message)" -ForegroundColor Red
            $testResults.PowerPlatform = $false
            Write-Host "  Note: Power Platform may require interactive auth" -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "⚠ Power Platform: Skipped (requires Windows PowerShell 5.x)" -ForegroundColor Yellow
        $testResults.PowerPlatform = $null
        Write-Host "  Current: PowerShell $($PSVersionTable.PSVersion)" -ForegroundColor Gray
    }
    
    Write-Host ""
    Write-Host "=== Connection Test Summary ===" -ForegroundColor Cyan
    foreach ($service in $testResults.Keys) {
        $status = switch ($testResults[$service]) {
            $true { "✓ Working" }
            $false { "✗ Failed" }
            $null { "⚠ Skipped" }
        }
        $color = switch ($testResults[$service]) {
            $true { "Green" }
            $false { "Red" }
            $null { "Yellow" }
        }
        Write-Host "  $service`: $status" -ForegroundColor $color
    }
    
    # Provide recommendations based on results
    $failedServices = $testResults.Keys | Where-Object { $testResults[$_] -eq $false }
    if ($failedServices.Count -gt 0) {
        Write-Host ""
        Write-Host "=== Troubleshooting Recommendations ===" -ForegroundColor Yellow
        Write-Host "For failed services, try:" -ForegroundColor White
        Write-Host "• Run: Get-M365AuditRequiredPermissions" -ForegroundColor White
        Write-Host "• Verify admin consent granted in Azure AD" -ForegroundColor White
        Write-Host "• Check app registration permissions and certificates" -ForegroundColor White
        Write-Host "• Run: Get-M365AuditTroubleshooting for detailed help" -ForegroundColor White
        
        if ($script:AppConfig.UseAppAuth) {
            Write-Host "• Verify certificate is uploaded to app registration (.cer file)" -ForegroundColor White
            Write-Host "• Check certificate expiration date" -ForegroundColor White
            Write-Host "• Ensure certificate has private key access" -ForegroundColor White
        }
    }
    
    return $testResults
}

function Test-M365AppRegistrationSetup {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantId,
        
        [Parameter(Mandatory = $true)]
        [string]$ClientId,
        
        [Parameter(Mandatory = $true)]
        [string]$CertificateThumbprint
    )
    
    Write-Host "=== Testing App Registration Setup (Certificate-Based) ===" -ForegroundColor Green
    Write-Host "Tenant ID: $TenantId" -ForegroundColor Gray
    Write-Host "Client ID: $ClientId" -ForegroundColor Gray
    Write-Host "Certificate Thumbprint: $CertificateThumbprint" -ForegroundColor Gray
    Write-Host "Authentication Type: Certificate-based" -ForegroundColor Cyan
    Write-Host ""
    
    $testResults = @{}
    
    # Test certificate availability
    Write-Host "Testing certificate availability..." -ForegroundColor Yellow
    try {
        $cert = Get-ChildItem -Path "Cert:\CurrentUser\My" | Where-Object { $_.Thumbprint -eq $CertificateThumbprint }
        if (-not $cert) {
            $cert = Get-ChildItem -Path "Cert:\LocalMachine\My" | Where-Object { $_.Thumbprint -eq $CertificateThumbprint }
        }
        
        if ($cert) {
            $isValid = (Get-Date) -ge $cert.NotBefore -and (Get-Date) -le $cert.NotAfter
            if ($isValid) {
                Write-Host "✓ Certificate found and valid" -ForegroundColor Green
                Write-Host "  Subject: $($cert.Subject)" -ForegroundColor Gray
                Write-Host "  Expires: $($cert.NotAfter)" -ForegroundColor Gray
                Write-Host "  Has Private Key: $($cert.HasPrivateKey)" -ForegroundColor Gray
                Write-Host "  Store Location: $(if ($cert.PSPath -like '*CurrentUser*') { 'CurrentUser' } else { 'LocalMachine' })" -ForegroundColor Gray
                $testResults.CertificateValid = $true
            }
            else {
                Write-Host "✗ Certificate expired or not yet valid" -ForegroundColor Red
                Write-Host "  Valid from: $($cert.NotBefore) to $($cert.NotAfter)" -ForegroundColor Red
                $testResults.CertificateValid = $false
            }
        }
        else {
            Write-Host "✗ Certificate not found in certificate store" -ForegroundColor Red
            Write-Host "  Searched in both CurrentUser\My and LocalMachine\My" -ForegroundColor Yellow
            $testResults.CertificateValid = $false
        }
    }
    catch {
        Write-Host "✗ Certificate test failed: $($_.Exception.Message)" -ForegroundColor Red
        $testResults.CertificateValid = $false
    }
    
    # Test basic Graph connection
    Write-Host "Testing Microsoft Graph authentication..." -ForegroundColor Yellow
    try {
        # Import Microsoft Graph module if not loaded
        if (-not (Get-Module -Name "Microsoft.Graph.Authentication" -ListAvailable)) {
            Write-Warning "Microsoft.Graph.Authentication module not found. Installing..."
            Install-Module -Name Microsoft.Graph.Authentication -Force -AllowClobber -Scope CurrentUser
        }
        
        Import-Module Microsoft.Graph.Authentication -Force
        
        # Test connection
        Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome
        
        # Verify connection
        $context = Get-MgContext
        if ($context -and $context.AuthType -eq "AppOnly") {
            Write-Host "✓ Graph authentication successful (App-Only)" -ForegroundColor Green
            Write-Host "  App Name: $($context.AppName)" -ForegroundColor Gray
            Write-Host "  Account: $($context.Account)" -ForegroundColor Gray
            Write-Host "  Auth Type: $($context.AuthType)" -ForegroundColor Gray
            $testResults.GraphAuth = $true
            
            # Test specific permissions
            try {
                # Fixed: Use the returned value to validate Directory.Read.All permission
                $testUsers = Get-MgUser -Top 1 -ErrorAction Stop
                if ($testUsers) {
                    Write-Host "✓ Directory.Read.All permission working" -ForegroundColor Green
                    $testResults.DirectoryRead = $true
                }
            }
            catch {
                Write-Host "✗ Directory.Read.All permission failed: $($_.Exception.Message)" -ForegroundColor Red
                $testResults.DirectoryRead = $false
            }
            
            try {
                # Fixed: Use the returned value to validate RoleManagement.Read.All permission
                $testRoles = Get-MgRoleManagementDirectoryRoleDefinition -Top 1 -ErrorAction Stop
                if ($testRoles) {
                    Write-Host "✓ RoleManagement.Read.All permission working" -ForegroundColor Green
                    $testResults.RoleManagementRead = $true
                }
            }
            catch {
                Write-Host "✗ RoleManagement.Read.All permission failed: $($_.Exception.Message)" -ForegroundColor Red
                $testResults.RoleManagementRead = $false
            }
            
            Disconnect-MgGraph -ErrorAction SilentlyContinue
        }
        else {
            Write-Host "✗ Graph authentication failed or not app-only" -ForegroundColor Red
            $testResults.GraphAuth = $false
        }
    }
    catch {
        Write-Host "✗ Graph authentication failed: $($_.Exception.Message)" -ForegroundColor Red
        $testResults.GraphAuth = $false
    }
    
    # Test Exchange Online authentication
    Write-Host "Testing Exchange Online authentication..." -ForegroundColor Yellow
    try {
        Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -Organization $TenantId -ShowBanner:$false -ErrorAction Stop
        
        $org = Get-OrganizationConfig -ErrorAction Stop
        if ($org) {
            Write-Host "✓ Exchange Online authentication successful" -ForegroundColor Green
            Write-Host "  Organization: $($org.DisplayName)" -ForegroundColor Gray
            $testResults.ExchangeAuth = $true
        }
        
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }
    catch {
        Write-Host "✗ Exchange Online authentication failed: $($_.Exception.Message)" -ForegroundColor Red
        $testResults.ExchangeAuth = $false
    }
    
    # Test SharePoint Online authentication
    Write-Host "Testing SharePoint Online authentication..." -ForegroundColor Yellow
    try {
        # Use a generic SharePoint admin URL for testing
        $testUrl = "https://tenant-admin.sharepoint.com"
        if ($SharePointTenantUrl) {
            $testUrl = $SharePointTenantUrl
        }
        
        Connect-PnPOnline -Url $testUrl -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -Tenant $TenantId -ErrorAction Stop
        
        # Test basic SharePoint operation
        $connection = Get-PnPConnection
        if ($connection) {
            Write-Host "✓ SharePoint Online authentication successful" -ForegroundColor Green
            Write-Host "  Connection Type: $($connection.ConnectionType)" -ForegroundColor Gray
            $testResults.SharePointAuth = $true
        }
        
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
    }
    catch {
        Write-Host "✗ SharePoint Online authentication failed: $($_.Exception.Message)" -ForegroundColor Red
        $testResults.SharePointAuth = $false
    }
    
    Write-Host ""
    Write-Host "=== App Registration Test Summary ===" -ForegroundColor Cyan
    foreach ($test in $testResults.Keys) {
        $status = if ($testResults[$test]) { "✓ Working" } else { "✗ Failed" }
        $color = if ($testResults[$test]) { "Green" } else { "Red" }
        Write-Host "  $test`: $status" -ForegroundColor $color
    }
    
    # Provide specific recommendations
    Write-Host ""
    $failedTests = $testResults.Keys | Where-Object { $testResults[$_] -eq $false }
    if ($failedTests.Count -gt 0) {
        Write-Host "=== Recommendations ===" -ForegroundColor Yellow
        
        if ("CertificateValid" -in $failedTests) {
            Write-Host "Certificate Issues:" -ForegroundColor Red
            Write-Host "• Verify certificate thumbprint is correct" -ForegroundColor White
            Write-Host "• Check certificate expiration date" -ForegroundColor White
            Write-Host "• Ensure certificate is in CurrentUser\My or LocalMachine\My store" -ForegroundColor White
            Write-Host "• Run: New-M365AuditCertificate to create a new certificate" -ForegroundColor White
        }
        
        if ("GraphAuth" -in $failedTests) {
            Write-Host "Authentication Issues:" -ForegroundColor Red
            Write-Host "• Verify TenantId and ClientId are correct" -ForegroundColor White
            Write-Host "• Check certificate is uploaded to app registration (.cer file)" -ForegroundColor White
            Write-Host "• Ensure app registration is not disabled" -ForegroundColor White
            Write-Host "• Verify certificate thumbprint matches uploaded certificate" -ForegroundColor White
        }
        
        if ("DirectoryRead" -in $failedTests -or "RoleManagementRead" -in $failedTests) {
            Write-Host "Permission Issues:" -ForegroundColor Red
            Write-Host "• Run: Get-M365AuditRequiredPermissions" -ForegroundColor White
            Write-Host "• Grant admin consent for API permissions" -ForegroundColor White
            Write-Host "• Verify app has required Graph permissions" -ForegroundColor White
        }
        
        if ("ExchangeAuth" -in $failedTests) {
            Write-Host "Exchange Issues:" -ForegroundColor Red
            Write-Host "• Ensure Exchange.ManageAsApp permission is granted" -ForegroundColor White
            Write-Host "• Verify admin consent for Exchange permissions" -ForegroundColor White
        }
        
        if ("SharePointAuth" -in $failedTests) {
            Write-Host "SharePoint Issues:" -ForegroundColor Red
            Write-Host "• Ensure Sites.FullControl.All permission is granted" -ForegroundColor White
            Write-Host "• Check SharePoint admin center app permissions" -ForegroundColor White
            Write-Host "• Verify SharePoint tenant URL is correct" -ForegroundColor White
        }
    }
    else {
        Write-Host "✓ All tests passed! App registration is ready for use." -ForegroundColor Green
        Write-Host "You can now use certificate-based authentication for M365 role audits." -ForegroundColor Cyan
    }
    
    return $testResults
}

function Get-M365AuditTroubleshooting {
    Write-Host "=== M365 Role Audit Troubleshooting Guide (Certificate-Based) ===" -ForegroundColor Green
    Write-Host ""
    
    Write-Host "Certificate-Based Authentication Issues:" -ForegroundColor Cyan
    Write-Host ""
    
    Write-Host "1. Certificate Not Found:" -ForegroundColor Yellow
    Write-Host "   Problem: 'Certificate with thumbprint 'xxx' not found'"
    Write-Host "   Solution: Create or import certificate"
    Write-Host "   Steps:"
    Write-Host "   - Run: New-M365AuditCertificate to create new certificate"
    Write-Host "   - Run: Get-M365AuditCertificate to find existing certificates"
    Write-Host "   - Verify certificate is in CurrentUser\My or LocalMachine\My store"
    Write-Host ""
    
    Write-Host "2. Certificate Expired:" -ForegroundColor Yellow
    Write-Host "   Problem: 'Certificate is expired or not yet valid'"
    Write-Host "   Solution: Create new certificate and update app registration"
    Write-Host "   Steps:"
    Write-Host "   - Run: New-M365AuditCertificate -Force"
    Write-Host "   - Upload new .cer file to Azure AD app registration"
    Write-Host "   - Remove old certificate from app registration"
    Write-Host ""
    
    Write-Host "3. Certificate Not Uploaded to Azure AD:" -ForegroundColor Yellow
    Write-Host "   Problem: Authentication fails despite valid local certificate"
    Write-Host "   Solution: Upload certificate to app registration"
    Write-Host "   Steps:"
    Write-Host "   - Go to Azure Portal > App registrations > Your App"
    Write-Host "   - Navigate to 'Certificates & secrets'"
    Write-Host "   - Click 'Upload certificate'"
    Write-Host "   - Upload the .cer file (not .pfx)"
    Write-Host ""
    
    Write-Host "4. Certificate Thumbprint Mismatch:" -ForegroundColor Yellow
    Write-Host "   Problem: 'Forbidden' or 'Invalid client' errors"
    Write-Host "   Solution: Verify thumbprint matches exactly"
    Write-Host "   Steps:"
    Write-Host "   - Run: Get-M365AuditCertificate to get correct thumbprint"
    Write-Host "   - Compare with thumbprint in Azure AD app registration"
    Write-Host "   - Ensure no extra spaces or characters in thumbprint"
    Write-Host ""
    
    Write-Host "App Registration Setup Issues:" -ForegroundColor Cyan
    Write-Host ""
    
    Write-Host "5. Insufficient Permissions:" -ForegroundColor Yellow
    Write-Host "   Problem: 'Insufficient privileges to complete the operation'"
    Write-Host "   Solution: Grant required API permissions and admin consent"
    Write-Host "   Run: Get-M365AuditRequiredPermissions"
    Write-Host "   Then grant admin consent in Azure Portal"
    Write-Host ""
    
    Write-Host "6. SharePoint Certificate Authentication:" -ForegroundColor Yellow
    Write-Host "   Problem: SharePoint connection fails with certificate"
    Write-Host "   Solution: Ensure certificate has SharePoint permissions"
    Write-Host "   - Verify Sites.FullControl.All permission is granted"
    Write-Host "   - Check SharePoint admin center app permissions"
    Write-Host "   - Ensure certificate is uploaded to app registration"
    Write-Host ""
    
    Write-Host "7. Exchange Certificate Authentication:" -ForegroundColor Yellow
    Write-Host "   Problem: 'Application does not have permission to perform this operation'"
    Write-Host "   Solution: Grant Exchange.ManageAsApp permission"
    Write-Host "   - Add Exchange.ManageAsApp in app registration"
    Write-Host "   - Grant admin consent"
    Write-Host "   - Verify certificate is properly configured"
    Write-Host ""
    
    Write-Host "Common Issues and Solutions:" -ForegroundColor Cyan
    Write-Host ""
    
    Write-Host "8. Mixed Authentication Errors:" -ForegroundColor Yellow
    Write-Host "   Problem: Some services work, others don't"
    Write-Host "   Solution: Verify permissions for each service"
    Write-Host "   - Each M365 service requires specific API permissions"
    Write-Host "   - Run service-specific connection tests"
    Write-Host "   - Check admin consent status for each permission"
    Write-Host ""
    
    Write-Host "9. PowerShell Version Issues:" -ForegroundColor Yellow
    Write-Host "   Problem: 'Power Platform modules require Windows PowerShell 5.x'"
    Write-Host "   Solution: Use appropriate PowerShell version"
    Write-Host "   - Power Platform: Windows PowerShell 5.x only"
    Write-Host "   - Other services: PowerShell 5.x or 7.x"
    Write-Host "   - Check: `$PSVersionTable.PSVersion"
    Write-Host ""
    
    Write-Host "10. Certificate Store Access:" -ForegroundColor Yellow
    Write-Host "    Problem: Certificate exists but can't be accessed"
    Write-Host "    Solution: Check certificate store permissions"
    Write-Host "    - Ensure user has access to certificate private key"
    Write-Host "    - Try running PowerShell as administrator"
    Write-Host "    - Check certificate was created with proper key usage"
    Write-Host ""
    
    Write-Host "11. Conditional Access Policies:" -ForegroundColor Yellow
    Write-Host "    Problem: App blocked by conditional access"
    Write-Host "    Solution: Configure conditional access exemptions"
    Write-Host "    - Check Azure AD > Security > Conditional Access"
    Write-Host "    - Create app exemption if needed"
    Write-Host "    - Consider trusted network locations"
    Write-Host ""
    
    Write-Host "Quick Diagnostic Commands:" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "# Check current configuration"
    Write-Host "Get-M365AuditCurrentConfig"
    Write-Host ""
    Write-Host "# Test certificate-based setup"
    Write-Host "Test-M365AppRegistrationSetup -TenantId 'your-tenant' -ClientId 'your-client' -CertificateThumbprint 'thumbprint'"
    Write-Host ""
    Write-Host "# Test individual service connections"
    Write-Host "Test-M365AuditConnections -TenantId 'your-tenant' -ClientId 'your-client' -CertificateThumbprint 'thumbprint'"
    Write-Host ""
    Write-Host "# Find certificates"
    Write-Host "Get-M365AuditCertificate"
    Write-Host ""
    Write-Host "# Create new certificate"
    Write-Host "New-M365AuditCertificate -ValidityMonths 24"
    Write-Host ""
    Write-Host "# Get required permissions list"
    Write-Host "Get-M365AuditRequiredPermissions"
    Write-Host ""
    Write-Host "# Clear all sessions and start fresh"
    Write-Host "Clear-M365AuditAppCredentials; Disconnect-MgGraph; Disconnect-ExchangeOnline -Confirm:`$false"
    Write-Host ""
    
    Write-Host "Security Best Practices (Certificate-Based):" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "1. Certificate Management:" -ForegroundColor Green
    Write-Host "   • Store certificates in Windows Certificate Store with non-exportable keys" -ForegroundColor White
    Write-Host "   • Set appropriate certificate validity period (12-24 months)" -ForegroundColor White
    Write-Host "   • Implement certificate rotation policies" -ForegroundColor White
    Write-Host "   • Use strong key lengths (2048-bit minimum)" -ForegroundColor White
    Write-Host ""
    Write-Host "2. Access Control:" -ForegroundColor Green
    Write-Host "   • Use least privilege principle for API permissions" -ForegroundColor White
    Write-Host "   • Restrict app registration to specific IP ranges if possible" -ForegroundColor White
    Write-Host "   • Monitor app usage through Azure AD audit logs" -ForegroundColor White
    Write-Host "   • Regularly review and audit app registrations" -ForegroundColor White
    Write-Host ""
    Write-Host "3. Certificate Lifecycle:" -ForegroundColor Green
    Write-Host "   • Monitor certificate expiration dates" -ForegroundColor White
    Write-Host "   • Plan certificate renewal 30 days before expiration" -ForegroundColor White
    Write-Host "   • Test new certificates before deployment" -ForegroundColor White
    Write-Host "   • Remove old certificates from Azure AD after successful rotation" -ForegroundColor White
    Write-Host ""
    
    Write-Host "Certificate Management Workflow:" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Certificate Lifecycle:" -ForegroundColor White
    Write-Host "• Creation: New-M365AuditCertificate" -ForegroundColor Gray
    Write-Host "• Validation: Test-M365AuditCertificate" -ForegroundColor Gray
    Write-Host "• Monitoring: Get-M365AuditCertificate" -ForegroundColor Gray
    Write-Host "• Renewal: New-M365AuditCertificate -Force (when near expiry)" -ForegroundColor Gray
    Write-Host "• Cleanup: Remove-M365AuditCertificate" -ForegroundColor Gray
    Write-Host ""
    
    Write-Host "Why Certificate-Based Authentication:" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Benefits over Client Secrets:" -ForegroundColor Green
    Write-Host "• No password/secret to manage or store" -ForegroundColor White
    Write-Host "• Private key cannot be extracted from certificate store" -ForegroundColor White
    Write-Host "• Better compliance with security frameworks" -ForegroundColor White
    Write-Host "• Reduced risk of credential exposure" -ForegroundColor White
    Write-Host "• Supports hardware security modules (HSMs)" -ForegroundColor White
    Write-Host "• Automatic certificate-based authentication" -ForegroundColor White
    Write-Host "• No expiration management for secrets (only certificate renewal)" -ForegroundColor White
    Write-Host "• Better audit trail and non-repudiation" -ForegroundColor White
    Write-Host ""
    
    Write-Host "Certificate Requirements:" -ForegroundColor Green
    Write-Host "• Must be stored in Windows Certificate Store" -ForegroundColor White
    Write-Host "• Private key must be available and accessible" -ForegroundColor White
    Write-Host "• Certificate (.cer) file must be uploaded to Azure AD app registration" -ForegroundColor White
    Write-Host "• Key usage must include Digital Signature" -ForegroundColor White
    Write-Host "• Recommended key length: 2048 bits or higher" -ForegroundColor White
    Write-Host ""
}

# Quick test to see if Teams and Defender roles are also PIM-based in your environment

function Test-TeamsDefenderPIM {
    try {
        Write-Host "=== Testing Teams and Defender Roles for PIM ===" -ForegroundColor Green
        Write-Host ""
        
        # Teams-specific Azure AD roles
        $teamsRoles = @(
            "Teams Administrator",
            "Teams Communications Administrator",
            "Teams Communications Support Engineer", 
            "Teams Communications Support Specialist",
            "Teams Devices Administrator",
            "Teams Telephony Administrator"
        )
        
        # Defender-related Azure AD roles
        $defenderRoles = @(
            "Security Administrator",
            "Security Operator", 
            "Security Reader",
            "Global Administrator",
            "Cloud Application Administrator",
            "Application Administrator"
        )
        
        Write-Host "1. Testing Teams Roles..." -ForegroundColor Cyan
        $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition | Where-Object { $_.DisplayName -in $teamsRoles }
        Write-Host "Found $($roleDefinitions.Count) Teams role definitions" -ForegroundColor Gray
        
        # Check regular assignments
        $regularAssignments = Get-MgRoleManagementDirectoryRoleAssignment | Where-Object { $_.RoleDefinitionId -in $roleDefinitions.Id }
        Write-Host "Regular Teams assignments: $($regularAssignments.Count)" -ForegroundColor $(if($regularAssignments.Count -eq 0) {"Red"} else {"Green"})
        
        # Check PIM active assignments
        $pimActiveCount = 0
        try {
            foreach ($roleId in $roleDefinitions.Id) {
                $pimActive = Get-MgRoleManagementDirectoryRoleAssignmentSchedule -Filter "roleDefinitionId eq '$roleId'" -ErrorAction SilentlyContinue
                if ($pimActive) {
                    $pimActiveCount += $pimActive.Count
                }
            }
        }
        catch {
            Write-Verbose "PIM active check failed for Teams: $($_.Exception.Message)"
        }
        Write-Host "PIM Active Teams assignments: $pimActiveCount" -ForegroundColor $(if($pimActiveCount -gt 0) {"Green"} else {"Gray"})
        
        # Check PIM eligible assignments
        $pimEligibleCount = 0
        try {
            foreach ($roleId in $roleDefinitions.Id) {
                $pimEligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -Filter "roleDefinitionId eq '$roleId'" -ErrorAction SilentlyContinue
                if ($pimEligible) {
                    $pimEligibleCount += $pimEligible.Count
                }
            }
        }
        catch {
            Write-Verbose "PIM eligible check failed for Teams: $($_.Exception.Message)"
        }
        Write-Host "PIM Eligible Teams assignments: $pimEligibleCount" -ForegroundColor $(if($pimEligibleCount -gt 0) {"Green"} else {"Gray"})
        
        Write-Host ""
        Write-Host "2. Testing Defender/Security Roles..." -ForegroundColor Cyan
        $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition | Where-Object { $_.DisplayName -in $defenderRoles }
        Write-Host "Found $($roleDefinitions.Count) Defender/Security role definitions" -ForegroundColor Gray
        
        # Check regular assignments
        $regularAssignments = Get-MgRoleManagementDirectoryRoleAssignment | Where-Object { $_.RoleDefinitionId -in $roleDefinitions.Id }
        Write-Host "Regular Defender/Security assignments: $($regularAssignments.Count)" -ForegroundColor $(if($regularAssignments.Count -eq 0) {"Red"} else {"Green"})
        
        # Check PIM active assignments
        $pimActiveCount = 0
        try {
            foreach ($roleId in $roleDefinitions.Id) {
                $pimActive = Get-MgRoleManagementDirectoryRoleAssignmentSchedule -Filter "roleDefinitionId eq '$roleId'" -ErrorAction SilentlyContinue
                if ($pimActive) {
                    $pimActiveCount += $pimActive.Count
                }
            }
        }
        catch {
            Write-Verbose "PIM active check failed for Defender: $($_.Exception.Message)"
        }
        Write-Host "PIM Active Defender/Security assignments: $pimActiveCount" -ForegroundColor $(if($pimActiveCount -gt 0) {"Green"} else {"Gray"})
        
        # Check PIM eligible assignments
        $pimEligibleCount = 0
        try {
            foreach ($roleId in $roleDefinitions.Id) {
                $pimEligible = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -Filter "roleDefinitionId eq '$roleId'" -ErrorAction SilentlyContinue
                if ($pimEligible) {
                    $pimEligibleCount += $pimEligible.Count
                }
            }
        }
        catch {
            Write-Verbose "PIM eligible check failed for Defender: $($_.Exception.Message)"
        }
        Write-Host "PIM Eligible Defender/Security assignments: $pimEligibleCount" -ForegroundColor $(if($pimEligibleCount -gt 0) {"Green"} else {"Gray"})
        
        Write-Host ""
        Write-Host "=== SUMMARY ===" -ForegroundColor Green
        Write-Host "If you see 0 regular assignments but >0 PIM assignments," -ForegroundColor Yellow
        Write-Host "then your Teams and Defender functions need the same PIM fix!" -ForegroundColor Yellow
        
    }
    catch {
        Write-Error "Test failed: $($_.Exception.Message)"
    }
}

# Function to quickly update Teams function with PIM support
function Get-TeamsRoleAudit-WithPIM {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$CertificateThumbprint
    )
    
    $results = @()
    
    try {
        # Certificate authentication setup (same as before)
        if ($TenantId -and $ClientId -and $CertificateThumbprint) {
            Set-M365AuditCertCredentials -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
        }
        
        if (-not $script:AppConfig.UseAppAuth -or $script:AppConfig.AuthType -ne "Certificate") {
            throw "Certificate authentication is required for Teams role audit. Use Set-M365AuditCertCredentials first."
        }
        
        $context = Get-MgContext
        if (-not $context -or $context.AuthType -ne "AppOnly") {
            Connect-MgGraph -TenantId $script:AppConfig.TenantId -ClientId $script:AppConfig.ClientId -CertificateThumbprint $script:AppConfig.CertificateThumbprint -NoWelcome
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
        
        # Get ALL assignment types (regular + PIM)
        $allAssignments = @()
        
        # 1. Regular assignments
        $regularAssignments = Get-MgRoleManagementDirectoryRoleAssignment | Where-Object { $_.RoleDefinitionId -in $roleDefinitions.Id }
        if ($regularAssignments) { $allAssignments += $regularAssignments }
        Write-Host "Found $($regularAssignments.Count) regular Teams assignments" -ForegroundColor Gray
        
        # 2. PIM eligible assignments
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
        catch { }
        Write-Host "Found $pimEligibleCount PIM eligible Teams assignments" -ForegroundColor Gray
        
        # 3. PIM active assignments
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
        catch { }
        Write-Host "Found $pimActiveCount PIM active Teams assignments" -ForegroundColor $(if($pimActiveCount -gt 0) {"Green"} else {"Gray"})
        
        # Process all assignments (same logic as Power Platform function)
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
                }
                
            }
            catch {
                Write-Verbose "Error processing Teams assignment: $($_.Exception.Message)"
            }
        }
        
        Write-Host "✓ Teams role audit completed. Found $($results.Count) role assignments (including PIM)" -ForegroundColor Green
        
    }
    catch {
        Write-Error "Error auditing Teams roles: $($_.Exception.Message)"
        throw
    }
    
    return $results
}

# Write-Host "Run this to test if Teams and Defender need PIM support:"
# Write-Host "Test-TeamsDefenderPIM" -ForegroundColor Cyan
# Write-Host ""
# Write-Host "Then test the PIM-aware Teams function:"
# Write-Host "Get-TeamsRoleAudit-WithPIM" -ForegroundColor Cyan