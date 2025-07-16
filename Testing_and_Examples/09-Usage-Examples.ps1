# 09-Usage-Examples.ps1
# Usage examples and documentation for M365 Role Audit with Certificate-based App Registration

<#
=== MICROSOFT 365 ROLE AUDIT WITH CERTIFICATE-BASED APP REGISTRATION ===

This script collection provides comprehensive role auditing across Microsoft 365 services
with support for both interactive and certificate-based application registration authentication.

FILE STRUCTURE:
01-CoreFunctions.ps1                  - Core authentication and configuration functions
02-AzureAD-Functions.ps1              - Azure AD/Entra ID role functions  
03-SharePoint-Functions.ps1           - SharePoint Online role functions
04-Exchange-Functions.ps1             - Exchange Online role functions
05-Compliance-Functions.ps1           - Microsoft Purview/Compliance role functions
06-PowerPlatform-Functions.ps1        - Power Platform role functions (Windows PS 5.x only)
07-Main-Audit-Function.ps1            - Main comprehensive audit function
08-Testing-Troubleshooting-Functions.ps1 - Testing and troubleshooting functions
09-Usage-Examples.ps1                 - This file with usage examples
10-Enhanced-Reporting-Functions.ps1   - Enhanced reporting and analysis functions
11-Intune-Functions.ps1               - Microsoft Intune/Endpoint Manager role functions
New-M365AuditCertificate.ps1          - Certificate creation and management script

SETUP INSTRUCTIONS:
1. Load all files in order, or combine them into a single script
2. Set up your Azure AD App Registration with required permissions
3. Create and configure certificate-based authentication
4. Run audits across M365 services

=== BASIC USAGE EXAMPLES ===
#>

# Example 1: Quick setup with certificate-based app registration
function Example-BasicCertificateSetup {
    # Step 1: Create certificate for authentication
    Write-Host "Creating certificate for M365 audit..." -ForegroundColor Green
    $certResult = New-M365AuditCertificate -CertificateName "M365-RoleAudit-Prod" -ValidityMonths 24
    
    Write-Host "Certificate created successfully!" -ForegroundColor Green
    Write-Host "Thumbprint: $($certResult.Thumbprint)" -ForegroundColor Cyan
    Write-Host "Export file: $($certResult.ExportPath)" -ForegroundColor Cyan
    
    # Step 2: Manual Azure AD configuration required
    Write-Host ""
    Write-Host "=== NEXT STEPS ===" -ForegroundColor Yellow
    Write-Host "1. Upload $($certResult.ExportPath) to your Azure AD app registration" -ForegroundColor White
    Write-Host "2. Grant required API permissions (run Get-M365AuditRequiredPermissions)" -ForegroundColor White
    Write-Host "3. Grant admin consent for the permissions" -ForegroundColor White
    Write-Host "4. Test with Example-CertificateBasedAudit" -ForegroundColor White
    
    # Step 3: Set credentials for future use
    $tenantId = Read-Host "Enter your Tenant ID"
    $clientId = Read-Host "Enter your Client (App) ID"
    
    Set-M365AuditCertCredentials -TenantId $tenantId -ClientId $clientId -CertificateThumbprint $certResult.Thumbprint
    
    # Step 4: Initialize environment and install required modules
    Initialize-M365AuditEnvironment
    
    Write-Host "Setup complete! Ready to run certificate-based audits." -ForegroundColor Green
}

# Example 2: Comprehensive audit with certificate authentication
function Example-CertificateBasedAudit {
    # Ensure certificate credentials are configured
    if (-not $script:AppConfig.UseAppAuth -or $script:AppConfig.AuthType -ne "Certificate") {
        Write-Warning "Certificate authentication not configured. Run Example-BasicCertificateSetup first."
        return
    }
    
    # Test connections first
    Write-Host "Testing all service connections..." -ForegroundColor Yellow
    $testResults = Test-M365AuditConnections -SharePointTenantUrl "https://yourtenant-admin.sharepoint.com"
    
    $failedServices = $testResults.Keys | Where-Object { $testResults[$_] -eq $false }
    if ($failedServices.Count -gt 0) {
        Write-Warning "Some services failed connection tests. Continuing with available services..."
        Write-Host "Failed: $($failedServices -join ', ')" -ForegroundColor Red
    }
    
    # Run comprehensive audit of all services
    $results = Get-ComprehensiveM365RoleAuditPnP -IncludeAll -SharePointTenantUrl "https://yourtenant-admin.sharepoint.com"
    
    # Results are automatically exported to CSV and displayed on screen
    Write-Host "Certificate-based audit completed. Found $($results.Count) role assignments." -ForegroundColor Green
    return $results
}

# Example 3: Individual service audits with certificate authentication
function Example-IndividualCertificateAudits {
    # Replace with your actual details
    $tenantId = "your-tenant-id"
    $clientId = "your-client-id"
    $thumbprint = "your-certificate-thumbprint"
    
    # Set certificate credentials
    Set-M365AuditCertCredentials -TenantId $tenantId -ClientId $clientId -CertificateThumbprint $thumbprint
    
    # Azure AD roles with PIM
    Write-Host "Auditing Azure AD roles..." -ForegroundColor Cyan
    $azureADRoles = Get-AzureADRoleAudit -IncludePIM
    Write-Host "Azure AD: $($azureADRoles.Count) assignments found" -ForegroundColor Green
    
    # SharePoint roles including OneDrive
    Write-Host "Auditing SharePoint roles..." -ForegroundColor Cyan
    $sharePointRoles = Get-SharePointRoleAudit -TenantUrl "https://yourtenant-admin.sharepoint.com" -IncludeOneDrive
    Write-Host "SharePoint: $($sharePointRoles.Count) assignments found" -ForegroundColor Green
    
    # Exchange roles
    Write-Host "Auditing Exchange roles..." -ForegroundColor Cyan
    $exchangeRoles = Get-ExchangeRoleAudit
    Write-Host "Exchange: $($exchangeRoles.Count) assignments found" -ForegroundColor Green
    
    # Compliance/Purview roles
    Write-Host "Auditing Purview roles..." -ForegroundColor Cyan
    $purviewRoles = Get-PurviewRoleAudit
    Write-Host "Purview: $($purviewRoles.Count) assignments found" -ForegroundColor Green
    
    # Intune roles
    Write-Host "Auditing Intune roles..." -ForegroundColor Cyan
    $intuneRoles = Get-IntuneRoleAudit
    Write-Host "Intune: $($intuneRoles.Count) assignments found" -ForegroundColor Green
    
    Write-Host "Individual certificate-based audits completed." -ForegroundColor Green
}

# Example 4: Using parameters for certificate-based authentication
function Example-ParameterBasedCertificateAuth {
    $tenantId = "your-tenant-id"
    $clientId = "your-client-id"
    $thumbprint = "your-certificate-thumbprint"
    
    # Pass certificate credentials directly to main function
    $results = Get-ComprehensiveM365RoleAuditPnP `
        -IncludeAll `
        -SharePointTenantUrl "https://yourtenant-admin.sharepoint.com" `
        -TenantId $tenantId `
        -ClientId $clientId `
        -CertificateThumbprint $thumbprint
    
    Write-Host "Parameter-based certificate audit completed." -ForegroundColor Green
    return $results
}

# Example 5: Interactive authentication (for development/testing)
function Example-InteractiveAuth {
    # Clear any stored app credentials to force interactive auth
    Clear-M365AuditAppCredentials
    
    # Run audit with interactive authentication
    $results = Get-ComprehensiveM365RoleAuditPnP -IncludeAll -AuthMethod "Interactive"
    
    Write-Host "Interactive audit completed." -ForegroundColor Green
    return $results
}

# Example 6: Device code authentication (for headless scenarios)
function Example-DeviceCodeAuth {
    # Use device code authentication (useful for headless systems)
    $results = Get-ComprehensiveM365RoleAuditPnP -IncludeAll -AuthMethod "DeviceCode"
    
    Write-Host "Device code audit completed." -ForegroundColor Green
    return $results
}

# Example 7: Selective service auditing with certificate
function Example-SelectiveCertificateAudit {
    # Only audit specific services with certificate authentication
    $results = Get-ComprehensiveM365RoleAuditPnP `
        -IncludeExchange `
        -IncludeSharePoint `
        -IncludePurview `
        -IncludeIntune `
        -SharePointTenantUrl "https://yourtenant-admin.sharepoint.com"
    
    Write-Host "Selective certificate audit completed." -ForegroundColor Green
    return $results
}

# Example 8: Certificate management and testing
function Example-CertificateManagement {
    # Create new certificate
    Write-Host "Creating new certificate..." -ForegroundColor Cyan
    $newCert = New-M365AuditCertificate -CertificateName "M365-Audit-New" -ValidityMonths 12
    
    # Find existing certificates
    Write-Host "Finding existing certificates..." -ForegroundColor Cyan
    $existingCerts = Get-M365AuditCertificate
    
    # Test certificate with app registration
    $tenantId = Read-Host "Enter Tenant ID for testing"
    $clientId = Read-Host "Enter Client ID for testing"
    
    Write-Host "Testing certificate authentication..." -ForegroundColor Cyan
    $testResult = Test-M365AuditCertificate -TenantId $tenantId -ClientId $clientId -CertificateThumbprint $newCert.Thumbprint
    
    if ($testResult) {
        Write-Host "✓ Certificate test passed!" -ForegroundColor Green
    } else {
        Write-Host "✗ Certificate test failed. Check Azure AD app registration." -ForegroundColor Red
    }
    
    # Get troubleshooting guide
    Get-M365AuditTroubleshooting
    
    # Get required permissions list
    Get-M365AuditRequiredPermissions
}

# Example 9: Automated script for scheduled execution (certificate-based)
function Example-ScheduledCertificateAudit {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantId,
        
        [Parameter(Mandatory = $true)]
        [string]$ClientId,
        
        [Parameter(Mandatory = $true)]
        [string]$CertificateThumbprint,
        
        [Parameter(Mandatory = $true)]
        [string]$SharePointTenantUrl,
        
        [string]$OutputDirectory = "C:\M365Audits"
    )
    
    try {
        # Ensure output directory exists
        if (-not (Test-Path $OutputDirectory)) {
            New-Item -ItemType Directory -Path $OutputDirectory -Force
        }
        
        # Set custom export path with timestamp
        $exportPath = Join-Path $OutputDirectory "M365_CertAudit_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $logPath = Join-Path $OutputDirectory "audit_log.txt"
        
        # Set certificate-based credentials
        Set-M365AuditCertCredentials -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
        
        # Validate certificate before running audit
        $certValid = Test-M365AuditCertificate -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
        if (-not $certValid) {
            throw "Certificate validation failed"
        }
        
        # Run comprehensive audit
        $results = Get-ComprehensiveM365RoleAuditPnP `
            -IncludeAll `
            -SharePointTenantUrl $SharePointTenantUrl `
            -ExportPath $exportPath
        
        # Generate enhanced reports
        $htmlReport = Export-M365AuditHtmlReport -AuditResults $results -OutputPath $exportPath.Replace('.csv', '.html')
        $jsonReport = Export-M365AuditJsonReport -AuditResults $results -OutputPath $exportPath.Replace('.csv', '.json')
        
        # Log results
        $logEntry = "$(Get-Date): Certificate-based audit completed. Found $($results.Count) role assignments. Exports: CSV=$exportPath, HTML=$htmlReport, JSON=$jsonReport"
        Add-Content -Path $logPath -Value $logEntry
        
        Write-Host "Scheduled certificate audit completed successfully." -ForegroundColor Green
        return $results
    }
    catch {
        $errorEntry = "$(Get-Date): Certificate audit failed. Error: $($_.Exception.Message)"
        Add-Content -Path $logPath -Value $errorEntry
        Write-Error "Scheduled certificate audit failed: $($_.Exception.Message)"
        throw
    }
}

# Example 10: Custom analysis of certificate-based audit results
function Example-CustomCertificateAnalysis {
    # Run audit and capture results using certificate authentication
    $results = Get-ComprehensiveM365RoleAuditPnP -IncludeAll
    
    if ($results.Count -eq 0) {
        Write-Warning "No audit results found. Ensure certificate authentication is properly configured."
        return
    }
    
    # Custom analysis examples
    Write-Host "=== CERTIFICATE-BASED AUDIT ANALYSIS ===" -ForegroundColor Green
    
    # Verify all connections used certificate authentication
    $certAuthResults = $results | Where-Object { $_.AuthenticationType -eq "Certificate" }
    $otherAuthResults = $results | Where-Object { $_.AuthenticationType -ne "Certificate" }
    
    Write-Host "Certificate Authentication: $($certAuthResults.Count) assignments" -ForegroundColor Green
    if ($otherAuthResults.Count -gt 0) {
        Write-Host "Other Authentication: $($otherAuthResults.Count) assignments" -ForegroundColor Yellow
        Write-Host "Note: Some services may fall back to interactive authentication" -ForegroundColor Gray
    }
    
    # Find users with Global Administrator roles
    $globalAdmins = $results | Where-Object { $_.RoleName -eq "Global Administrator" }
    Write-Host "Global Administrators: $($globalAdmins.Count)" -ForegroundColor $(if ($globalAdmins.Count -gt 5) { "Red" } else { "Green" })
    
    # Find disabled users with active roles
    $disabledUsersWithRoles = $results | Where-Object { $_.UserEnabled -eq $false }
    Write-Host "Disabled users with active roles: $($disabledUsersWithRoles.Count)" -ForegroundColor $(if ($disabledUsersWithRoles.Count -gt 0) { "Red" } else { "Green" })
    
    # Find users who haven't signed in recently (if data available)
    $staleUsers = $results | Where-Object { 
        $_.LastSignIn -and 
        $_.LastSignIn -lt (Get-Date).AddDays(-90) 
    }
    Write-Host "Users with roles who haven't signed in for 90+ days: $($staleUsers.Count)" -ForegroundColor $(if ($staleUsers.Count -gt 0) { "Orange" } else { "Green" })
    
    # Analyze Intune-specific findings
    $intuneResults = $results | Where-Object { $_.Service -eq "Microsoft Intune" }
    if ($intuneResults.Count -gt 0) {
        Write-Host ""
        Write-Host "=== INTUNE ANALYSIS ===" -ForegroundColor Cyan
        
        $intuneServiceAdmins = $intuneResults | Where-Object { $_.RoleName -eq "Intune Service Administrator" }
        $intuneRBACRoles = $intuneResults | Where-Object { $_.RoleType -eq "IntuneRBAC" }
        $intuneAzureADRoles = $intuneResults | Where-Object { $_.RoleType -eq "AzureAD" }
        
        Write-Host "Intune Service Administrators: $($intuneServiceAdmins.Count)" -ForegroundColor White
        Write-Host "Intune RBAC Role Assignments: $($intuneRBACRoles.Count)" -ForegroundColor White
        Write-Host "Azure AD Intune Role Assignments: $($intuneAzureADRoles.Count)" -ForegroundColor White
        
        if ($intuneAzureADRoles.Count -gt $intuneRBACRoles.Count) {
            Write-Host "⚠️ Recommendation: Consider using Intune RBAC roles for better granular control" -ForegroundColor Yellow
        }
    }
    
    # Find users with multiple high-privilege roles
    $highPrivRoles = @("Global Administrator", "Security Administrator", "Exchange Administrator", "SharePoint Administrator", "Intune Service Administrator")
    $multiRoleUsers = $results | 
        Where-Object { $_.RoleName -in $highPrivRoles } |
        Group-Object UserPrincipalName |
        Where-Object { $_.Count -gt 1 }
    
    Write-Host "Users with multiple high-privilege roles: $($multiRoleUsers.Count)" -ForegroundColor $(if ($multiRoleUsers.Count -gt 0) { "Yellow" } else { "Green" })
    
    # Service coverage analysis
    Write-Host ""
    Write-Host "=== SERVICE COVERAGE ===" -ForegroundColor Cyan
    $serviceCoverage = $results | Group-Object Service | Sort-Object Count -Descending
    foreach ($service in $serviceCoverage) {
        Write-Host "  $($service.Name): $($service.Count) assignments" -ForegroundColor White
    }
    
    # Export custom analysis
    $analysisPath = ".\M365_CertificateAudit_Analysis_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $customAnalysis = @{
        GlobalAdmins = $globalAdmins
        DisabledUsersWithRoles = $disabledUsersWithRoles  
        StaleUsers = $staleUsers
        MultiRoleUsers = $multiRoleUsers
        IntuneAnalysis = $intuneResults
    }
    
    # Save analysis results
    foreach ($analysisType in $customAnalysis.Keys) {
        $analysisData = $customAnalysis[$analysisType]
        if ($analysisData -and $analysisData.Count -gt 0) {
            $analysisData | Export-Csv -Path ".\$analysisType`_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv" -NoTypeInformation
        }
    }
    
    # Generate compliance gap analysis
    $complianceGaps = Get-M365ComplianceGaps -AuditResults $results
    
    Write-Host "Certificate-based audit analysis completed and exported." -ForegroundColor Green
    return @{
        Results = $results
        Analysis = $customAnalysis
        ComplianceGaps = $complianceGaps
    }
}

<#
=== COMMON CERTIFICATE-BASED WORKFLOWS ===

WORKFLOW 1: First-time certificate setup
1. Example-BasicCertificateSetup
2. Upload certificate to Azure AD app registration
3. Get-M365AuditRequiredPermissions (review and configure in Azure Portal)
4. Test-M365AuditCertificate (validate setup)
5. Example-CertificateBasedAudit

WORKFLOW 2: Regular scheduled certificate audits
1. Example-ScheduledCertificateAudit (configure as scheduled task)
2. Example-CustomCertificateAnalysis (for detailed reporting)

WORKFLOW 3: Certificate troubleshooting
1. Get-M365AuditTroubleshooting
2. Test-M365AuditConnections
3. Test-M365AuditCertificate

WORKFLOW 4: Interactive development/testing
1. Example-InteractiveAuth (for development)
2. Example-IndividualCertificateAudits (for production)
3. Example-CustomCertificateAnalysis

=== CERTIFICATE SECURITY CONSIDERATIONS ===

1. Certificate Management:
   - Use certificate-based authentication for production environments
   - Store certificates in Windows Certificate Store with non-exportable keys
   - Set appropriate certificate validity period (12-24 months)
   - Implement certificate rotation policies before expiration

2. App Registration Security:
   - Use least privilege permissions for API access
   - Implement conditional access policies for app registrations
   - Monitor app usage through Azure AD audit logs
   - Restrict app registration to specific IP ranges if possible

3. Data Protection:
   - Encrypt exported CSV/JSON files containing sensitive role data
   - Store audit results in secure locations with appropriate access controls
   - Implement data retention policies for audit exports
   - Restrict access to audit results based on business need

4. Certificate Lifecycle:
   - Monitor certificate expiration dates
   - Test certificate authentication before expiration
   - Plan certificate renewal in advance
   - Remove old certificates from Azure AD after rotation

=== CERTIFICATE TROUBLESHOOTING QUICK REFERENCE ===

Common Certificate Issues:
- Certificate not found: Run Get-M365AuditCertificate to verify installation
- Certificate expired: Create new certificate with New-M365AuditCertificate -Force
- Azure AD upload: Upload .cer file (not .pfx) to app registration
- Permission errors: Run Get-M365AuditRequiredPermissions and grant admin consent
- SharePoint authentication: Verify Sites.FullControl.All permission granted
- Exchange authentication: Verify Exchange.ManageAsApp permission granted

Certificate Support Functions:
- New-M365AuditCertificate: Create new certificate for authentication
- Get-M365AuditCertificate: Find existing certificates
- Test-M365AuditCertificate: Validate certificate authentication
- Remove-M365AuditCertificate: Clean up old certificates
- Get-M365AuditTroubleshooting: Comprehensive troubleshooting guide

Authentication Hierarchy (Recommended Order):
1. Certificate-based (Production) - Most secure, no password required
2. Interactive (Development) - User-friendly for testing and development
3. Device Code (Headless) - For servers or automated systems without UI

=== END OF CERTIFICATE-BASED EXAMPLES ===
#>