function Get-ComprehensiveM365RoleAudit {
    param(
        [string]$ExportPath = ".\M365_Comprehensive_RoleAudit_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",

        [Parameter(Mandatory = $true)]
        [string]$SharePointTenantUrl, # Example  = "https://balfourbeattyus-admin.sharepoint.com"

        [Parameter(Mandatory = $False)]
        [string]$Organization,  # Required for Exchange and Compliance

        [switch]$IncludePIM,
        [switch]$IncludeExchange,
        [switch]$IncludeSharePoint,
        [switch]$IncludeTeams,
        [switch]$IncludePurview,
        [switch]$IncludeDefender,
        [switch]$IncludeIntune,
        [switch]$IncludePowerPlatform,
        [switch]$IncludeAll,

        # Can gbe set with Set-M365AuditCredentials
        [string]$TenantId,
        [string]$ClientId,
        [string]$CertificateThumbprint,
        
        # Deduplication Parameters
        [ValidateSet("Strict", "Loose", "ServicePreference", "None")]
        [string]$DeduplicationMode = "ServicePreference",
        [switch]$ShowDuplicatesRemoved,
        [switch]$PreferAzureADSource
    )
    
    $allResults = @()
    #$installationNeeded = @()
    
    try {
        Write-Host "=== Microsoft 365 Comprehensive Role Audit ===" -ForegroundColor Green
        Write-Host "Using Certificate-Based Authentication Only" -ForegroundColor Cyan
        if ($DeduplicationMode -ne "None") {
            Write-Host "Deduplication Mode: $DeduplicationMode" -ForegroundColor Cyan
        }
        Write-Host ""
        
        # Determine Organization from TenantId if not provided
        If ($IncludeExchange -or $IncludePurview -or $IncludeAll) {
            If (-not $Organization) {
                Write-Host "Organization is required for Exchange and Purview auditing!"
                exit
            }
        }
        # Set app credentials if provided
        if ($TenantId -and $ClientId -and $CertificateThumbprint) {
            Set-M365AuditCertCredentials -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
            Write-Host "Using certificate-based authentication" -ForegroundColor Green
        }
        else {
            Write-Host "Using previously configured certificate credentials" -ForegroundColor Cyan
            Write-Host "Authentication Type: $($script:AppConfig.AuthType)" -ForegroundColor Cyan
        }
        
        Write-Host "Authentication Method: Certificate-based Application Authentication" -ForegroundColor Cyan
        Write-Host "Target SharePoint tenant: $SharePointTenantUrl" -ForegroundColor Cyan
        Write-Host "Exchange Organization: $Organization" -ForegroundColor Cyan
        Write-Host ""
        
        # Display current configuration
        # Get-M365AuditCurrentConfig
        # Write-Host ""  
        
        # 1. Azure AD/Entra ID Roles (always included)
        Write-Host "Auditing Azure AD/Entra ID roles..." -ForegroundColor Cyan
        $azureADRoles = Get-AzureADRoleAudit -IncludePIM:$IncludePIM -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
        $allResults += $azureADRoles
        Write-Host "✓ Found $($azureADRoles.Count) Azure AD role assignments" -ForegroundColor Green
        
        # 2. SharePoint Online Roles
        if ($IncludeSharePoint -or $IncludeAll) {
            Write-Host "Auditing SharePoint Online roles..." -ForegroundColor Cyan
            $sharePointRoles = Get-SharePointRoleAudit -TenantUrl $SharePointTenantUrl -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
            $allResults += $sharePointRoles
            Write-Host "✓ Found $($sharePointRoles.Count) SharePoint role assignments" -ForegroundColor Green
        }
        
        # 3. Exchange Online Roles
        if ($IncludeExchange -or $IncludeAll) {
            if ($Organization) {
                Write-Host "Auditing Exchange Online roles..." -ForegroundColor Cyan
                $exchangeRoles = Get-ExchangeRoleAudit -Organization $Organization -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
                $allResults += $exchangeRoles
                Write-Host "✓ Found $($exchangeRoles.Count) Exchange role assignments" -ForegroundColor Green
            }
            else {
                Write-Warning "Skipping Exchange audit - Organization parameter required"
            }
        }
        
        # 4. Microsoft Purview/Compliance Roles
        if ($IncludePurview -or $IncludeAll) {
            if ($Organization) {
                Write-Host "Auditing Microsoft Purview/Compliance roles..." -ForegroundColor Cyan
                $purviewRoles = Get-PurviewRoleAudit -Organization $Organization -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
                $allResults += $purviewRoles
                Write-Host "✓ Found $($purviewRoles.Count) Purview role assignments" -ForegroundColor Green
            }
            else {
                Write-Warning "Skipping Purview audit - Organization parameter required"
            }
        }
        
        # 5. Teams Roles
        if ($IncludeTeams -or $IncludeAll) {
            Write-Host "Auditing Microsoft Teams roles..." -ForegroundColor Cyan
            $teamsRoles = Get-TeamsRoleAudit -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
            $allResults += $teamsRoles
            Write-Host "✓ Found $($teamsRoles.Count) Teams role assignments" -ForegroundColor Green
        }
        
        # 6. Microsoft Defender Roles
        if ($IncludeDefender -or $IncludeAll) {
            Write-Host "Auditing Microsoft Defender roles..." -ForegroundColor Cyan
            $defenderRoles = Get-DefenderRoleAudit -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
            $allResults += $defenderRoles
            Write-Host "✓ Found $($defenderRoles.Count) Defender role assignments" -ForegroundColor Green
        }
        
        # 7. Microsoft Intune/Endpoint Manager Roles
        if ($IncludeIntune -or $IncludeAll) {
            Write-Host "Auditing Microsoft Intune/Endpoint Manager roles..." -ForegroundColor Cyan
            $intuneRoles = Get-IntuneRoleAudit -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -IncludePIM:$IncludePIM
            $allResults += $intuneRoles
            Write-Host "✓ Found $($intuneRoles.Count) Intune role assignments" -ForegroundColor Green
        }
        
        # 8. Power Platform Roles (Windows PowerShell 5.x only)
        if ($IncludePowerPlatform -or $IncludeAll) {
            Write-Host "Auditing Power Platform roles..." -ForegroundColor Cyan
            try {
                # Use the Azure AD Power Platform roles function since native Power Platform has limited cert auth
                $powerPlatformRoles = Get-PowerPlatformAzureADRoleAudit -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
                $allResults += $powerPlatformRoles
                Write-Host "✓ Found $($powerPlatformRoles.Count) Power Platform role assignments" -ForegroundColor Green
            }
            catch {
                Write-Warning "Power Platform audit failed: $($_.Exception.Message)"
                Write-Host "Note: Power Platform has limited certificate authentication support" -ForegroundColor Yellow
            }
        }
        
        # Apply deduplication if requested
        if ($DeduplicationMode -ne "None" -and $allResults.Count -gt 0) {
            Write-Host ""
            Write-Host "=== APPLYING DEDUPLICATION ===" -ForegroundColor Cyan
            
            $originalCount = $allResults.Count
            
            $deduplicationParams = @{
                AuditResults = $allResults
                DeduplicationMode = $DeduplicationMode
                ShowDuplicatesRemoved = $ShowDuplicatesRemoved
                PreferAzureADSource = $PreferAzureADSource
            }
            
            $allResults = Remove-M365AuditDuplicates @deduplicationParams
            
            $deduplicatedCount = $allResults.Count
            $removedCount = $originalCount - $deduplicatedCount
            
            Write-Host ""
            Write-Host "Deduplication Summary:" -ForegroundColor Green
            Write-Host "  Original: $originalCount assignments" -ForegroundColor White
            Write-Host "  Final: $deduplicatedCount assignments" -ForegroundColor White
            Write-Host "  Removed: $removedCount duplicates" -ForegroundColor Yellow
            Write-Host "  Efficiency: $([math]::Round(($removedCount / $originalCount) * 100, 1))%" -ForegroundColor Cyan
        }
        
        # Export comprehensive results
        if ($allResults.Count -gt 0) {
            $allResults | Export-Csv -Path $ExportPath -NoTypeInformation
            Write-Host ""
            Write-Host "=== AUDIT COMPLETED SUCCESSFULLY ===" -ForegroundColor Green
            Write-Host "Total role assignments found: $($allResults.Count)" -ForegroundColor Green
            
            if ($DeduplicationMode -ne "None") {
                Write-Host "Results deduplicated using: $DeduplicationMode mode" -ForegroundColor Cyan
            }
            
            Write-Host "Results exported to: $ExportPath" -ForegroundColor Green
            Write-Host "Authentication used: Certificate-based" -ForegroundColor Cyan
            
            # Display detailed summary
            Write-Host ""
            Write-Host "=== SUMMARY BY SERVICE ===" -ForegroundColor Cyan
            $serviceSummary = $allResults | Group-Object Service | Sort-Object Count -Descending
            foreach ($service in $serviceSummary) {
                Write-Host "  $($service.Name): $($service.Count) assignments" -ForegroundColor White
            }
            
            Write-Host ""
            Write-Host "=== AUTHENTICATION TYPE BREAKDOWN ===" -ForegroundColor Cyan
            $authSummary = $allResults | Group-Object AuthenticationType | Sort-Object Count -Descending
            foreach ($authType in $authSummary) {
                Write-Host "  $($authType.Name): $($authType.Count) assignments" -ForegroundColor White
            }
            
            Write-Host ""
            Write-Host "=== ASSIGNMENT TYPE BREAKDOWN ===" -ForegroundColor Cyan
            $assignmentTypeSummary = $allResults | Group-Object AssignmentType | Sort-Object Count -Descending
            foreach ($assignmentType in $assignmentTypeSummary) {
                Write-Host "  $($assignmentType.Name): $($assignmentType.Count) assignments" -ForegroundColor White
            }
            
            Write-Host ""
            Write-Host "=== TOP ROLES ACROSS ALL SERVICES ===" -ForegroundColor Cyan
            $roleSummary = $allResults | Group-Object RoleName | Sort-Object Count -Descending | Select-Object -First 10
            foreach ($role in $roleSummary) {
                Write-Host "  $($role.Name): $($role.Count) assignments" -ForegroundColor White
            }
            
            Write-Host ""
            Write-Host "=== USERS WITH MOST ROLES ===" -ForegroundColor Cyan
            $userSummary = $allResults | Group-Object UserPrincipalName | Sort-Object Count -Descending | Select-Object -First 10
            foreach ($user in $userSummary) {
                if ($user.Name) {
                    Write-Host "  $($user.Name): $($user.Count) roles" -ForegroundColor White
                }
            }
            
            # PIM Analysis
            Write-Host ""
            Write-Host "=== PIM ANALYSIS ===" -ForegroundColor Cyan
            $pimEligible = $allResults | Where-Object { $_.AssignmentType -like "*Eligible*" }
            $pimActive = $allResults | Where-Object { $_.AssignmentType -like "*Active (PIM*" }
            $permanentActive = $allResults | Where-Object { $_.AssignmentType -eq "Active" -or $_.AssignmentType -eq "Azure AD Role" }
            
            Write-Host "  PIM Eligible Assignments: $($pimEligible.Count)" -ForegroundColor $(if($pimEligible.Count -gt 0) {"Green"} else {"Yellow"})
            Write-Host "  PIM Active Assignments: $($pimActive.Count)" -ForegroundColor White
            Write-Host "  Permanent Active Assignments: $($permanentActive.Count)" -ForegroundColor White
            
            # Security recommendations based on findings
            Write-Host ""
            Write-Host "=== SECURITY RECOMMENDATIONS ===" -ForegroundColor Yellow
            
            $globalAdmins = $allResults | Where-Object { $_.RoleName -eq "Global Administrator" }
            if ($globalAdmins.Count -gt 5) {
                Write-Host "⚠ Consider reducing Global Administrator count ($($globalAdmins.Count) found)" -ForegroundColor Yellow
            }
            
            $disabledUsers = $allResults | Where-Object { $_.UserEnabled -eq $false }
            if ($disabledUsers.Count -gt 0) {
                Write-Host "⚠ Found $($disabledUsers.Count) disabled users with active roles - review needed" -ForegroundColor Yellow
            }
            
            # PIM recommendations
            if ($pimEligible.Count -eq 0 -and $permanentActive.Count -gt 0) {
                Write-Host "⚠ Consider implementing PIM for eligible assignments" -ForegroundColor Yellow
            }
            elseif ($pimEligible.Count -gt 0) {
                Write-Host "✓ PIM eligible assignments detected - good security practice" -ForegroundColor Green
            }
            
            # Check for Intune-specific recommendations
            $intuneResults = $allResults | Where-Object { $_.Service -eq "Microsoft Intune" }
            if ($intuneResults.Count -gt 0) {
                $intuneServiceAdmins = $intuneResults | Where-Object { $_.RoleName -eq "Intune Service Administrator" }
                if ($intuneServiceAdmins.Count -gt 3) {
                    Write-Host "⚠ Consider reducing Intune Service Administrator count ($($intuneServiceAdmins.Count) found)" -ForegroundColor Yellow
                }
                
                $intuneRBACAssignments = $intuneResults | Where-Object { $_.RoleType -eq "IntuneRBAC" }
                $intuneAzureADAssignments = $intuneResults | Where-Object { $_.RoleType -eq "AzureAD" }
                
                if ($intuneAzureADAssignments.Count -gt $intuneRBACAssignments.Count) {
                    Write-Host "⚠ Consider using Intune RBAC roles instead of Azure AD roles for better granularity" -ForegroundColor Yellow
                }
            }
            
            Write-Host "✓ Using secure certificate-based authentication" -ForegroundColor Green
            
            # Additional security recommendations
            Write-Host ""
            Write-Host "General Security Recommendations:" -ForegroundColor Cyan
            Write-Host "• Implement regular access reviews for privileged roles" -ForegroundColor White
            Write-Host "• Enable Privileged Identity Management (PIM) for eligible assignments" -ForegroundColor White
            Write-Host "• Monitor privileged role assignments with alerts" -ForegroundColor White
            Write-Host "• Implement break-glass emergency access accounts" -ForegroundColor White
            Write-Host "• Regularly rotate certificates (recommended 12-24 months)" -ForegroundColor White
            
            if ($intuneResults.Count -gt 0) {
                Write-Host "• Review Intune policy ownership and scope assignments" -ForegroundColor White
                Write-Host "• Ensure device compliance policies are properly managed" -ForegroundColor White
            }
            
            # Offer enhanced reporting
            Write-Host ""
            Write-Host "Enhanced Reporting Options:" -ForegroundColor Cyan
            Write-Host "• Export-M365AuditHtmlReport -AuditResults `$results - Generate interactive HTML dashboard" -ForegroundColor White
            Write-Host "• Export-M365AuditJsonReport -AuditResults `$results - Generate structured JSON for automation" -ForegroundColor White
            Write-Host "• Get-M365RoleAnalysis -AuditResults `$results - Perform detailed role sprawl analysis" -ForegroundColor White
            Write-Host "• Get-M365ComplianceGaps -AuditResults `$results - Identify compliance gaps and risks" -ForegroundColor White
        }
        else {
            Write-Host "No role assignments found." -ForegroundColor Yellow
        }
        
        return $allResults
    }
    catch {
        Write-Error "Error during comprehensive audit: $($_.Exception.Message)"
        Write-Error "Stack trace: $($_.ScriptStackTrace)"
        
        # Provide troubleshooting guidance based on error type
        if ($_.Exception.Message -like "*certificate*") {
            Write-Host ""
            Write-Host "Certificate troubleshooting:" -ForegroundColor Yellow
            Write-Host "• Verify certificate exists in Windows Certificate Store" -ForegroundColor White
            Write-Host "• Check certificate expiration date" -ForegroundColor White
            Write-Host "• Ensure certificate is uploaded to Azure AD app registration" -ForegroundColor White
            Write-Host "• Run: Get-M365AuditCurrentConfig to verify setup" -ForegroundColor White
            Write-Host "• Run: New-M365AuditCertificate to create new certificate" -ForegroundColor White
        }
        elseif ($_.Exception.Message -like "*permission*" -or $_.Exception.Message -like "*access*") {
            Write-Host ""
            Write-Host "Permission troubleshooting:" -ForegroundColor Yellow
            Write-Host "• Run: Get-M365AuditRequiredPermissions" -ForegroundColor White
            Write-Host "• Verify admin consent has been granted in Azure AD" -ForegroundColor White
            Write-Host "• Check if app registration has required API permissions" -ForegroundColor White
            Write-Host "• For Intune: Ensure DeviceManagementRBAC.Read.All permission is granted" -ForegroundColor White
        }
        elseif ($_.Exception.Message -like "*Client secret*") {
            Write-Host ""
            Write-Host "Authentication migration required:" -ForegroundColor Yellow
            Write-Host "• Client secret authentication is no longer supported" -ForegroundColor White
            Write-Host "• Run: New-M365AuditCertificate to create certificate" -ForegroundColor White
            Write-Host "• Upload certificate to Azure AD app registration" -ForegroundColor White
            Write-Host "• Run: Set-M365AuditCertCredentials to configure" -ForegroundColor White
        }
        
        return @()
    }
}

