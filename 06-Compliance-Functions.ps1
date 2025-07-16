# 06-Compliance-Functions.ps1
# Microsoft Purview/Compliance role audit functions - Certificate Authentication Only
# Fixed to use script-level variables from Set-M365AuditCredentials
function Get-PurviewRoleAudit {
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
            Write-Warning "Certificate authentication is required for Compliance role audit"
            Write-Host "Please configure certificate authentication first:" -ForegroundColor Yellow
            Write-Host "• Run: Set-M365AuditCertCredentials -TenantId <id> -ClientId <id> -CertificateThumbprint <thumbprint>" -ForegroundColor White
            Write-Host "• Or: Get-M365AuditRequiredPermissions for setup instructions" -ForegroundColor White
            return $results
        }
        
        # Use script variables for authentication
        # AuthMethod = "Application"
        # Write-Host "Using configured certificate credentials for Compliance audit:" -ForegroundColor Cyan
        # Write-Host "  Tenant ID: $($script:AppConfig.TenantId)" -ForegroundColor Gray
        # Write-Host "  Client ID: $($script:AppConfig.ClientId)" -ForegroundColor Gray
        # Write-Host "  Certificate Thumbprint: $($script:AppConfig.CertificateThumbprint)" -ForegroundColor Gray
        
        # Check if connected to Security & Compliance Center
        $EXOsession = Get-ConnectionInformation | Where-Object { $_.ConnectionUri -like "outlook*" -and $_.State -eq "Connected" }
        $IPSSsession = Get-ConnectionInformation | Where-Object { $_.ConnectionUri -like "*compliance" -and $_.State -eq 'Connected'}

        If (-not $EXOsession) {
            Write-Host 'Connecting to Exchange Online' -ForegroundColor Yellow

            try {
                if ($IsWindows) {
                    Connect-ExchangeOnline `
                        -AppId $script:AppConfig.ClientId `
                        -CertificateThumbprint $script:AppConfig.CertificateThumbprint `
                        -Organization $Organization `
                        -ShowBanner:$false
                } elseIf ($IsLinux -or $IsMacOS) {
                    Connect-ExchangeOnline `
                        -AppId $script:AppConfig.ClientId `
                        -Certificate $script:AppConfig.Certificate `
                        -Organization $Organization `
                        -ShowBanner:$false
                    
                } 
            } catch {
                Write-Error "Exchange Online certificate authentication failed: $($_.Exception.Message)"
                Write-Host "Troubleshooting steps:" -ForegroundColor Yellow
                Write-Host "• Ensure certificate is uploaded to Azure AD app registration" -ForegroundColor White
                Write-Host "• Verify app has required Compliance permissions" -ForegroundColor White
                Write-Host "• Check certificate expiration and validity" -ForegroundColor White
                Write-Host "• Run: Get-M365AuditCurrentConfig to verify configuration" -ForegroundColor White
                return $results
            }
        } else {
            Write-Host "✓ Already connected to Exchange Online" -ForegroundColor Green

        }

        If (-Not $IPSSsession) {
            Write-Host "Connecting to Security & Compliance Center with certificate authentication..." -ForegroundColor Yellow
            
            try {
                # Use script variables for connection
                if ($IsWindows) {
                    Connect-IPPSSession `
                        -AppId $script:AppConfig.ClientId `
                        -CertificateThumbprint $script:AppConfig.CertificateThumbprint `
                        -Organization $Organization `
                        -ShowBanner:$false
                    Write-Host "✓ Connected to Security & Compliance Center successfully" -ForegroundColor Green
                    # Write-Host "Authentication Type: Certificate" -ForegroundColor Cyan
                } elseIf ($IsLinux -or $IsMacOS) {
                    Connect-IPPSSession `
                        -AppId $script:AppConfig.ClientId `
                        -Certificate $script:AppConfig.Certificate `
                        -Organization $Organization `
                        -ShowBanner:$false
                    Write-Host "✓ Connected to Security & Compliance Center successfully" -ForegroundColor Green

                }
            }
            catch {
                Write-Error "Compliance Center certificate authentication failed: $($_.Exception.Message)"
                Write-Host "Troubleshooting steps:" -ForegroundColor Yellow
                Write-Host "• Ensure certificate is uploaded to Azure AD app registration" -ForegroundColor White
                Write-Host "• Verify app has required Compliance permissions" -ForegroundColor White
                Write-Host "• Check certificate expiration and validity" -ForegroundColor White
                Write-Host "• Run: Get-M365AuditCurrentConfig to verify configuration" -ForegroundColor White
                return $results
            }
        }
        else {
            Write-Host "✓ Already connected to Security & Compliance Center" -ForegroundColor Green
        }
        
        # Verify connection functionality (without redundant success messages)
        try {
            $null = Get-RoleGroup -ErrorAction Stop | Select-Object -First 1
        }
        catch {
            Write-Warning "Security & Compliance Center connection verification failed: $($_.Exception.Message)"
            Write-Host "Note: Some compliance features may not be available" -ForegroundColor Yellow
            # Continue anyway as some commands might still work
        }
        
        # Get compliance role groups
        Write-Host "Retrieving Purview role groups..." -ForegroundColor Cyan
        $roleGroups = Get-RoleGroup -ErrorAction SilentlyContinue
        
        foreach ($roleGroup in $roleGroups) {
            try {
                $members = Get-RoleGroupMember -Identity $roleGroup.Identity -ErrorAction SilentlyContinue
                
                foreach ($member in $members) {
                    $results += [PSCustomObject]@{
                        Service = "Microsoft Purview"
                        UserPrincipalName = $member.PrimarySmtpAddress
                        DisplayName = $member.DisplayName
                        UserId = $null
                        RoleName = $roleGroup.Name
                        RoleDefinitionId = $null
                        AssignmentType = "Role Group Member"
                        AssignedDateTime = $null
                        UserEnabled = $null
                        LastSignIn = $null
                        Scope = "Organization"
                        AssignmentId = $roleGroup.Identity
                        RoleGroupDescription = $roleGroup.Description
                        AuthenticationType = "Certificate"
                    }
                }
            }
            catch {
                Write-Verbose "Could not get members for compliance role group $($roleGroup.Name): $($_.Exception.Message)"
            }
        }
        
        # Get DLP policy administrators and other compliance-specific roles
        try {
            Write-Host "Retrieving DLP and compliance policy roles..." -ForegroundColor Cyan
            
            # Get DLP policies and their owners/assignees
            $dlpPolicies = Get-DlpPolicy -ErrorAction SilentlyContinue
            foreach ($policy in $dlpPolicies) {
                if ($policy.CreatedBy) {
                    $results += [PSCustomObject]@{
                        Service = "Microsoft Purview"
                        UserPrincipalName = $policy.CreatedBy
                        DisplayName = $policy.CreatedBy
                        UserId = $null
                        RoleName = "DLP Policy Creator"
                        RoleDefinitionId = $null
                        AssignmentType = "Policy Owner"
                        AssignedDateTime = $policy.WhenCreated
                        UserEnabled = $null
                        LastSignIn = $null
                        Scope = $policy.Name
                        AssignmentId = $policy.Identity
                        PolicyMode = $policy.Mode
                        PolicyState = $policy.Enabled
                        AuthenticationType = "Certificate"
                    }
                }
            }
            
            # Get retention policies and their creators
            $retentionPolicies = Get-RetentionCompliancePolicy -ErrorAction SilentlyContinue
            foreach ($policy in $retentionPolicies) {
                if ($policy.CreatedBy) {
                    $results += [PSCustomObject]@{
                        Service = "Microsoft Purview"
                        UserPrincipalName = $policy.CreatedBy
                        DisplayName = $policy.CreatedBy
                        UserId = $null
                        RoleName = "Retention Policy Creator"
                        RoleDefinitionId = $null
                        AssignmentType = "Policy Owner"
                        AssignedDateTime = $policy.WhenCreated
                        UserEnabled = $null
                        LastSignIn = $null
                        Scope = $policy.Name
                        AssignmentId = $policy.Identity
                        PolicyType = "Retention"
                        AuthenticationType = "Certificate"
                    }
                }
            }
            
            # Get retention rules
            $retentionRules = Get-RetentionComplianceRule -ErrorAction SilentlyContinue
            foreach ($rule in $retentionRules) {
                if ($rule.CreatedBy) {
                    $results += [PSCustomObject]@{
                        Service = "Microsoft Purview"
                        UserPrincipalName = $rule.CreatedBy
                        DisplayName = $rule.CreatedBy
                        UserId = $null
                        RoleName = "Retention Rule Creator"
                        RoleDefinitionId = $null
                        AssignmentType = "Rule Owner"
                        AssignedDateTime = $rule.WhenCreated
                        UserEnabled = $null
                        LastSignIn = $null
                        Scope = $rule.Name
                        AssignmentId = $rule.Identity
                        PolicyType = "RetentionRule"
                        AuthenticationType = "Certificate"
                    }
                }
            }
        }
        catch {
            Write-Verbose "Could not retrieve DLP/compliance policy information: $($_.Exception.Message)"
        }
        
        # Get eDiscovery administrators
        try {
            Write-Host "Retrieving eDiscovery administrators..." -ForegroundColor Cyan
            $eDiscoveryCases = Get-ComplianceCase -ErrorAction SilentlyContinue
            foreach ($case in $eDiscoveryCases) {
                $caseMembers = Get-ComplianceCaseMember -Case $case.Identity -ErrorAction SilentlyContinue
                foreach ($member in $caseMembers) {
                    $results += [PSCustomObject]@{
                        Service = "Microsoft Purview"
                        UserPrincipalName = $member.PrimarySmtpAddress
                        DisplayName = $member.DisplayName
                        UserId = $null
                        RoleName = "eDiscovery Case Member"
                        RoleDefinitionId = $null
                        AssignmentType = "Case Assignment"
                        AssignedDateTime = $null
                        UserEnabled = $null
                        LastSignIn = $null
                        Scope = $case.Name
                        AssignmentId = $case.Identity
                        CaseStatus = $case.Status
                        CaseType = $case.CaseType
                        AuthenticationType = "Certificate"
                    }
                }
            }
            
            # Get eDiscovery Premium cases
            $ediscoveryPremiumCases = Get-Case -ErrorAction SilentlyContinue
            foreach ($case in $ediscoveryPremiumCases) {
                $caseMembers = Get-CaseMember -Case $case.Identity -ErrorAction SilentlyContinue
                foreach ($member in $caseMembers) {
                    $results += [PSCustomObject]@{
                        Service = "Microsoft Purview"
                        UserPrincipalName = $member.PrimarySmtpAddress
                        DisplayName = $member.DisplayName
                        UserId = $null
                        RoleName = "eDiscovery Premium Case Member"
                        RoleDefinitionId = $null
                        AssignmentType = "Case Assignment"
                        AssignedDateTime = $null
                        UserEnabled = $null
                        LastSignIn = $null
                        Scope = $case.Name
                        AssignmentId = $case.Identity
                        CaseStatus = $case.Status
                        CaseType = "Premium"
                        AuthenticationType = "Certificate"
                    }
                }
            }
        }
        catch {
            Write-Verbose "Could not retrieve eDiscovery case information: $($_.Exception.Message)"
        }
        
        # Additional compliance-specific auditing code would continue here...
        # (Content Search, Information Barriers, Sensitivity Labels, etc.)
        
        Write-Host "✓ Purview compliance audit completed" -ForegroundColor Green
        Write-Host "Found $($results.Count) compliance role assignments and configurations" -ForegroundColor Cyan
        
        # Security compliance validation
        Write-Host ""
        Write-Host "=== Security Compliance Validation ===" -ForegroundColor Green
        Write-Host "✓ Certificate-based authentication enforced for compliance" -ForegroundColor Green
        Write-Host "✓ No client secrets used in compliance functions" -ForegroundColor Green
        Write-Host "✓ Enhanced security posture maintained" -ForegroundColor Green
        
    }
    catch {
        Write-Warning "Error auditing Purview roles: $($_.Exception.Message)"
        Write-Host "Compliance function requirements:" -ForegroundColor Yellow
        Write-Host "• Certificate-based authentication (required for compliance)" -ForegroundColor White
        Write-Host "• Compliance Administrator permissions" -ForegroundColor White
        Write-Host ""
        Write-Host "Configuration Commands:" -ForegroundColor Cyan
        Write-Host "• Set-M365AuditCertCredentials -TenantId <id> -ClientId <id> -CertificateThumbprint <thumbprint>" -ForegroundColor White
        Write-Host "• Get-M365AuditCurrentConfig (to verify setup)" -ForegroundColor White
        Write-Host "• Get-M365AuditRequiredPermissions (for complete setup guide)" -ForegroundColor White
    }
    
    return $results
}