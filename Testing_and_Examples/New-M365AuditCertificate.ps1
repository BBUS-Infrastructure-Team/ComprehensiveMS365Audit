# New-M365AuditCertificate.ps1
# Certificate creation script for M365 Role Audit App Registration

function New-M365AuditCertificate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CertificateName,
        
        [Parameter(Mandatory = $false)]
        [string]$Subject,
        
        [Parameter(Mandatory = $false)]
        [int]$ValidityMonths = 24,
        
        [Parameter(Mandatory = $false)]
        [string]$ExportPath = ".\M365-Audit-Certificate.cer",
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("CurrentUser", "LocalMachine")]
        [string]$StoreLocation = "CurrentUser",
        
        [Parameter(Mandatory = $false)]
        [switch]$Force
    )
    
    Write-Host "=== Creating M365 Audit Certificate ===" -ForegroundColor Green
    Write-Host "Certificate Name: $CertificateName" -ForegroundColor Cyan
    Write-Host "Subject: $Subject" -ForegroundColor Cyan
    Write-Host "Validity: $ValidityMonths months" -ForegroundColor Cyan
    Write-Host "Store Location: $StoreLocation" -ForegroundColor Cyan
    Write-Host ""
    
    try {
        # Check if certificate already exists
        $existingCert = Get-ChildItem -Path "Cert:\$StoreLocation\My" | Where-Object { $_.Subject -eq $Subject }
        
        if ($existingCert -and -not $Force) {
            Write-Warning "Certificate with subject '$Subject' already exists."
            Write-Host "Existing certificate details:" -ForegroundColor Yellow
            Write-Host "  Thumbprint: $($existingCert.Thumbprint)" -ForegroundColor Gray
            Write-Host "  NotBefore: $($existingCert.NotBefore)" -ForegroundColor Gray
            Write-Host "  NotAfter: $($existingCert.NotAfter)" -ForegroundColor Gray
            Write-Host ""
            Write-Host "Use -Force to create a new certificate or use the existing one." -ForegroundColor Yellow
            
            # Export existing certificate for Azure AD registration
            $certBytes = $existingCert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
            Set-Content -Path $ExportPath -Value $certBytes -Encoding Byte
            Write-Host "✓ Existing certificate exported to: $ExportPath" -ForegroundColor Green
            
            return @{
                Certificate = $existingCert
                Thumbprint = $existingCert.Thumbprint
                ExportPath = $ExportPath
                IsNew = $false
            }
        }
        
        # Remove existing certificate if Force is specified
        if ($existingCert -and $Force) {
            Write-Host "Removing existing certificate..." -ForegroundColor Yellow
            $existingCert | Remove-Item -Force
        }
        
        # Calculate expiration date
        $notAfter = (Get-Date).AddMonths($ValidityMonths)
        
        # Create new self-signed certificate
        Write-Host "Creating new self-signed certificate..." -ForegroundColor Yellow
        
        $cert = New-SelfSignedCertificate `
            -Subject $Subject `
            -CertStoreLocation "Cert:\$StoreLocation\My" `
            -KeyExportPolicy NonExportable `
            -KeySpec Signature `
            -KeyLength 2048 `
            -KeyAlgorithm RSA `
            -HashAlgorithm SHA256 `
            -NotAfter $notAfter `
            -KeyUsage DigitalSignature `
            -Type Custom `
            -FriendlyName $CertificateName
        
        if (-not $cert) {
            throw "Failed to create certificate"
        }
        
        Write-Host "✓ Certificate created successfully" -ForegroundColor Green
        Write-Host "  Thumbprint: $($cert.Thumbprint)" -ForegroundColor Gray
        Write-Host "  Subject: $($cert.Subject)" -ForegroundColor Gray
        Write-Host "  NotBefore: $($cert.NotBefore)" -ForegroundColor Gray
        Write-Host "  NotAfter: $($cert.NotAfter)" -ForegroundColor Gray
        Write-Host "  Store Location: Cert:\$StoreLocation\My" -ForegroundColor Gray
        
        # Export certificate for Azure AD app registration
        Write-Host "Exporting certificate for Azure AD registration..." -ForegroundColor Yellow
        
        $certBytes = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
        Set-Content -Path $ExportPath -Value $certBytes -Encoding Byte
        
        Write-Host "✓ Certificate exported to: $ExportPath" -ForegroundColor Green
        Write-Host ""
        
        # Display Azure AD registration instructions
        Write-Host "=== Azure AD App Registration Instructions ===" -ForegroundColor Cyan
        Write-Host "1. Go to Azure Portal > Azure Active Directory > App registrations" -ForegroundColor White
        Write-Host "2. Select your M365 Role Audit app registration" -ForegroundColor White
        Write-Host "3. Go to 'Certificates & secrets'" -ForegroundColor White
        Write-Host "4. Click 'Upload certificate'" -ForegroundColor White
        Write-Host "5. Upload the file: $ExportPath" -ForegroundColor White
        Write-Host "6. Click 'Add'" -ForegroundColor White
        Write-Host ""
        Write-Host "Certificate Thumbprint for PowerShell scripts:" -ForegroundColor Yellow
        Write-Host $cert.Thumbprint -ForegroundColor Green
        Write-Host ""
        
        # Display PowerShell usage example
        Write-Host "=== PowerShell Usage Example ===" -ForegroundColor Cyan
        Write-Host "# Set certificate-based authentication" -ForegroundColor Green
        Write-Host "`$tenantId = 'your-tenant-id'" -ForegroundColor White
        Write-Host "`$clientId = 'your-client-id'" -ForegroundColor White
        Write-Host "`$thumbprint = '$($cert.Thumbprint)'" -ForegroundColor White
        Write-Host ""
        Write-Host "Set-M365AuditCertCredentials -TenantId `$tenantId -ClientId `$clientId -CertificateThumbprint `$thumbprint" -ForegroundColor White
        Write-Host "Get-ComprehensiveM365RoleAuditPnP -IncludeAll" -ForegroundColor White
        Write-Host ""
        
        return @{
            Certificate = $cert
            Thumbprint = $cert.Thumbprint
            ExportPath = $ExportPath
            IsNew = $true
        }
    }
    catch {
        Write-Error "Failed to create certificate: $($_.Exception.Message)"
        throw
    }
}

function Get-M365AuditCertificate {
    param(
        [Parameter(Mandatory = $false)]
        [string]$Subject = "CN=M365-RoleAudit-Certificate",
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("CurrentUser", "LocalMachine")]
        [string]$StoreLocation = "CurrentUser",
        
        [Parameter(Mandatory = $false)]
        [string]$Thumbprint
    )
    
    Write-Host "Searching for M365 Audit certificates..." -ForegroundColor Yellow
    
    try {
        if ($Thumbprint) {
            $certificates = Get-ChildItem -Path "Cert:\$StoreLocation\My" | Where-Object { $_.Thumbprint -eq $Thumbprint }
        }
        else {
            $certificates = Get-ChildItem -Path "Cert:\$StoreLocation\My" | Where-Object { $_.Subject -eq $Subject }
        }
        
        if ($certificates) {
            Write-Host "Found $($certificates.Count) matching certificate(s):" -ForegroundColor Green
            
            foreach ($cert in $certificates) {
                Write-Host ""
                Write-Host "Certificate Details:" -ForegroundColor Cyan
                Write-Host "  Thumbprint: $($cert.Thumbprint)" -ForegroundColor White
                Write-Host "  Subject: $($cert.Subject)" -ForegroundColor White
                Write-Host "  Issuer: $($cert.Issuer)" -ForegroundColor White
                Write-Host "  NotBefore: $($cert.NotBefore)" -ForegroundColor White
                Write-Host "  NotAfter: $($cert.NotAfter)" -ForegroundColor White
                Write-Host "  HasPrivateKey: $($cert.HasPrivateKey)" -ForegroundColor White
                Write-Host "  FriendlyName: $($cert.FriendlyName)" -ForegroundColor White
                
                # Check if certificate is valid
                $isValid = (Get-Date) -ge $cert.NotBefore -and (Get-Date) -le $cert.NotAfter
                if ($isValid) {
                    Write-Host "  Status: Valid" -ForegroundColor Green
                }
                else {
                    Write-Host "  Status: Expired or Not Yet Valid" -ForegroundColor Red
                }
            }
            
            return $certificates
        }
        else {
            Write-Host "No matching certificates found." -ForegroundColor Yellow
            Write-Host "Search criteria:" -ForegroundColor Gray
            if ($Thumbprint) {
                Write-Host "  Thumbprint: $Thumbprint" -ForegroundColor Gray
            }
            else {
                Write-Host "  Subject: $Subject" -ForegroundColor Gray
            }
            Write-Host "  Store Location: Cert:\$StoreLocation\My" -ForegroundColor Gray
            
            return $null
        }
    }
    catch {
        Write-Error "Error searching for certificates: $($_.Exception.Message)"
        return $null
    }
}

function Remove-M365AuditCertificate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Thumbprint,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("CurrentUser", "LocalMachine")]
        [string]$StoreLocation = "CurrentUser",
        
        [Parameter(Mandatory = $false)]
        [switch]$Force
    )
    
    try {
        $cert = Get-ChildItem -Path "Cert:\$StoreLocation\My" | Where-Object { $_.Thumbprint -eq $Thumbprint }
        
        if (-not $cert) {
            Write-Warning "Certificate with thumbprint '$Thumbprint' not found in Cert:\$StoreLocation\My"
            return
        }
        
        Write-Host "Found certificate to remove:" -ForegroundColor Yellow
        Write-Host "  Thumbprint: $($cert.Thumbprint)" -ForegroundColor Gray
        Write-Host "  Subject: $($cert.Subject)" -ForegroundColor Gray
        Write-Host "  NotAfter: $($cert.NotAfter)" -ForegroundColor Gray
        
        if (-not $Force) {
            $confirmation = Read-Host "Are you sure you want to remove this certificate? (y/N)"
            if ($confirmation -ne 'y' -and $confirmation -ne 'Y') {
                Write-Host "Certificate removal cancelled." -ForegroundColor Yellow
                return
            }
        }
        
        $cert | Remove-Item -Force
        Write-Host "✓ Certificate removed successfully" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to remove certificate: $($_.Exception.Message)"
    }
}

function Test-M365AuditCertificate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantId,
        
        [Parameter(Mandatory = $true)]
        [string]$ClientId,
        
        [Parameter(Mandatory = $true)]
        [string]$CertificateThumbprint,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("CurrentUser", "LocalMachine")]
        [string]$StoreLocation = "CurrentUser"
    )
    
    Write-Host "=== Testing M365 Audit Certificate Authentication ===" -ForegroundColor Green
    Write-Host "Tenant ID: $TenantId" -ForegroundColor Gray
    Write-Host "Client ID: $ClientId" -ForegroundColor Gray
    Write-Host "Certificate Thumbprint: $CertificateThumbprint" -ForegroundColor Gray
    Write-Host "Store Location: $StoreLocation" -ForegroundColor Gray
    Write-Host ""
    
    try {
        # Check if certificate exists and is valid
        $cert = Get-ChildItem -Path "Cert:\$StoreLocation\My" | Where-Object { $_.Thumbprint -eq $CertificateThumbprint }
        
        if (-not $cert) {
            Write-Error "Certificate with thumbprint '$CertificateThumbprint' not found in Cert:\$StoreLocation\My"
            return $false
        }
        
        # Check certificate validity
        $isValid = (Get-Date) -ge $cert.NotBefore -and (Get-Date) -le $cert.NotAfter
        if (-not $isValid) {
            Write-Error "Certificate is expired or not yet valid"
            Write-Host "Certificate validity: $($cert.NotBefore) to $($cert.NotAfter)" -ForegroundColor Red
            return $false
        }
        
        # Check private key
        if (-not $cert.HasPrivateKey) {
            Write-Error "Certificate does not have a private key"
            return $false
        }
        
        Write-Host "✓ Certificate found and valid" -ForegroundColor Green
        
        # Test Microsoft Graph authentication
        Write-Host "Testing Microsoft Graph authentication..." -ForegroundColor Yellow
        
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
            Write-Host "✓ Microsoft Graph authentication successful" -ForegroundColor Green
            
            # Test a simple Graph call
            try {
                [void](Get-MgUser -Top 1 -ErrorAction Stop)
                Write-Host "✓ Graph API call successful" -ForegroundColor Green
            }
            catch {
                Write-Warning "Graph API call failed: $($_.Exception.Message)"
                Write-Host "This may indicate insufficient permissions" -ForegroundColor Yellow
            }
            
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            return $true
        }
        else {
            Write-Error "Microsoft Graph authentication failed"
            return $false
        }
    }
    catch {
        Write-Error "Certificate test failed: $($_.Exception.Message)"
        return $false
    }
}
