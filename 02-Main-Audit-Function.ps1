# 02-Main-Audit-Function.ps1 - UPDATED VERSION
# Updated Get-ComprehensiveM365RoleAudit function with proper role filtering to eliminate deduplication needs

function Get-ComprehensiveM365RoleAudit {
    [CmdletBinding(DefaultParameterSetName = 'default')]
    param(
        # Common parameters available to all parameter sets
        [string]$ExportPath,
        [bool]$IncludePIM = $true,
        [switch]$IncludeAnalysis,
        [string]$TenantId,
        [string]$ClientId,
        [string]$CertificateThumbprint,
        [switch]$IncludeOverarchingRolesInServices,
        [ValidateSet("Strict", "Loose", "ServicePreference", "RoleScoped", "None")]
        [string]$DeduplicationMode = "None",
        [switch]$ShowDuplicatesRemoved,
        [switch]$PreferAzureADSource,

        # Individual service switches - only available when NOT using IncludeAll
        [Parameter(ParameterSetName = 'Exchange')]
        [Parameter(ParameterSetName = 'SharePoint')]
        [Parameter(ParameterSetName = 'Individual')]
        [switch]$IncludeTeams,

        [Parameter(ParameterSetName = 'Exchange')]
        [Parameter(ParameterSetName = 'SharePoint')]
        [Parameter(ParameterSetName = 'Individual')]
        [switch]$IncludePurview,

        [Parameter(ParameterSetName = 'Exchange')]
        [Parameter(ParameterSetName = 'SharePoint')]
        [Parameter(ParameterSetName = 'Individual')]
        [switch]$IncludeDefender,
        
        [Parameter(ParameterSetName = 'Exchange')]
        [Parameter(ParameterSetName = 'SharePoint')]
        [Parameter(ParameterSetName = 'Individual')]
        [switch]$IncludePowerPlatform,

        [Parameter(ParameterSetName = 'Exchange')]
        [Parameter(ParameterSetName = 'SharePoint')]
        [Parameter(ParameterSetName = 'Individual')]
        [switch]$IncludeExchange,

        [Parameter(ParameterSetName = 'Exchange')]
        [Parameter(ParameterSetName = 'SharePoint')]
        [Parameter(ParameterSetName = 'Individual')]
        [switch]$IncludeSharepoint,

        # Parameter set specific parameters
        [Parameter(
            Mandatory = $true,
            ParameterSetName = 'IncludeAll'
        )]
        [Parameter(
            Mandatory = $true,
            ParameterSetName = 'Exchange'
        )]
        [string]$Organization,

        [Parameter(
            Mandatory = $true,
            ParameterSetName = 'IncludeAll'
        )]
        [Parameter(
            Mandatory = $true,
            ParameterSetName = 'SharePoint'
        )]
        [string]$SharePointTenantUrl,

        [Parameter(
            Mandatory = $true,
            ParameterSetName = 'IncludeAll'
        )]
        [switch]$IncludeAll
    )

    $allResults = @()
    
    try {
        Write-Host "=== Microsoft 365 Comprehensive Role Audit ===" -ForegroundColor Green
        Write-Host "Using Enhanced Role Filtering (No Deduplication Required)" -ForegroundColor Cyan
        if ($IncludeOverarchingRolesInServices) {
            Write-Host "Including overarching roles in service audits (may cause duplicates)" -ForegroundColor Yellow
        } else {
            Write-Host "Service audits will exclude overarching roles (clean separation)" -ForegroundColor Green
        }
        
        
        # Validate required parameters
        if ($IncludeExchange -or $IncludePurview -or $IncludeAll) {
            if (-not $Organization) {
                Write-Host "Organization parameter is required for Exchange and Purview auditing!" -ForegroundColor Red
                return @()
            }
        }
        
        # Set app credentials if provided
        if ($TenantId -and $ClientId -and $CertificateThumbprint) {
            Set-M365AuditCertCredentials -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
            Write-Host "Using certificate-based authentication" -ForegroundColor Green
        }
        else {
            Write-Host "Using previously configured certificate credentials" -ForegroundColor Cyan
            if ($script:AppConfig.AuthType) {
                Write-Host "Authentication Type: $($script:AppConfig.AuthType)" -ForegroundColor Cyan
            }
        }
        
        Write-Host "Authentication Method: Certificate-based Application Authentication" -ForegroundColor Cyan
        Write-Host "Target SharePoint tenant: $SharePointTenantUrl" -ForegroundColor Cyan
        if ($Organization) {
            Write-Host "Exchange Organization: $Organization" -ForegroundColor Cyan
        }
        
        
        # === 1. AZURE AD/ENTRA ID ROLES (ALWAYS INCLUDED WITH ALL ROLES) ===
        Write-Host "Auditing Azure AD/Entra ID roles (including all overarching roles)..." -ForegroundColor Cyan
        $azureADRoles = Get-AzureADRoleAudit -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint 
        $allResults += $azureADRoles
        Write-Host "✓ Found $($azureADRoles.Count) Azure AD role assignments" -ForegroundColor Green
       
        # === 2. SERVICE-SPECIFIC AUDITS (EXCLUDE OVERARCHING ROLES BY DEFAULT) ===
        
        # SharePoint Online Roles
        if ($IncludeSharePoint -or $IncludeAll) {
            Write-Host "Auditing SharePoint Online roles..." -ForegroundColor Cyan
            $sharePointRoles = Get-SharePointRoleAudit -TenantUrl $SharePointTenantUrl -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -IncludeAzureADRoles:$IncludeOverarchingRolesInServices
            $allResults += $sharePointRoles
            Write-Host "✓ Found $($sharePointRoles.Count) SharePoint role assignments" -ForegroundColor Green
            if (-not $IncludeOverarchingRolesInServices) {
                Write-Host "  (Excluding overarching Azure AD roles to prevent duplicates)" -ForegroundColor Gray
            }
        }
        
        # Exchange Online Roles
        if ($IncludeExchange -or $IncludeAll) {
            if ($Organization) {
                Write-Host "Auditing Exchange Online roles..." -ForegroundColor Cyan
                $exchangeRoles = Get-ExchangeRoleAudit -Organization $Organization -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -IncludeAzureADRoles:$IncludeOverarchingRolesInServices
                $allResults += $exchangeRoles
                Write-Host "✓ Found $($exchangeRoles.Count) Exchange role assignments" -ForegroundColor Green
                if (-not $IncludeOverarchingRolesInServices) {
                    Write-Host "  (Excluding overarching Azure AD roles to prevent duplicates)" -ForegroundColor Gray
                }
            }
            else {
                Write-Warning "Skipping Exchange audit - Organization parameter required"
            }
        }
        
        # Microsoft Purview/Compliance Roles
        if ($IncludePurview -or $IncludeAll) {
            if ($Organization) {
                Write-Host "Auditing Microsoft Purview/Compliance roles..." -ForegroundColor Cyan
                $purviewRoles = Get-PurviewRoleAudit -Organization $Organization -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -IncludeAzureADRoles:$IncludeOverarchingRolesInServices
                $allResults += $purviewRoles
                Write-Host "✓ Found $($purviewRoles.Count) Purview role assignments" -ForegroundColor Green
                if (-not $IncludeOverarchingRolesInServices) {
                    Write-Host "  (Excluding overarching Azure AD roles to prevent duplicates)" -ForegroundColor Gray
                }
            }
            else {
                Write-Warning "Skipping Purview audit - Organization parameter required"
            }
        }
        
        # Teams Roles
        if ($IncludeTeams -or $IncludeAll) {
            Write-Host "Auditing Microsoft Teams roles..." -ForegroundColor Cyan
            $teamsRoles = Get-TeamsRoleAudit -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -IncludeAzureADRoles:$IncludeOverarchingRolesInServices
            $allResults += $teamsRoles
            Write-Host "✓ Found $($teamsRoles.Count) Teams role assignments" -ForegroundColor Green
            if (-not $IncludeOverarchingRolesInServices) {
                Write-Host "  (Excluding overarching Azure AD roles to prevent duplicates)" -ForegroundColor Gray
            }
        }
        
        # Microsoft Defender Roles
        if ($IncludeDefender -or $IncludeAll) {
            Write-Host "Auditing Microsoft Defender roles..." -ForegroundColor Cyan
            $defenderRoles = Get-DefenderRoleAudit -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -IncludeAzureADRoles:$IncludeOverarchingRolesInServices
            $allResults += $defenderRoles
            Write-Host "✓ Found $($defenderRoles.Count) Defender role assignments" -ForegroundColor Green
            if (-not $IncludeOverarchingRolesInServices) {
                Write-Host "  (Excluding overarching Azure AD roles to prevent duplicates)" -ForegroundColor Gray
            }
        }
        
        # Microsoft Intune/Endpoint Manager Roles
        if ($IncludeIntune -or $IncludeAll) {
            Write-Host "Auditing Microsoft Intune/Endpoint Manager roles..." -ForegroundColor Cyan
            $intuneRoles = Get-IntuneRoleAudit -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -IncludePIM:$IncludePIM -IncludeAzureADRoles:$IncludeOverarchingRolesInServices
            $allResults += $intuneRoles
            Write-Host "✓ Found $($intuneRoles.Count) Intune role assignments" -ForegroundColor Green
            if (-not $IncludeOverarchingRolesInServices) {
                Write-Host "  (Excluding overarching Azure AD roles to prevent duplicates)" -ForegroundColor Gray
            }
        }
        
        # Power Platform Roles
        if ($IncludePowerPlatform -or $IncludeAll) {
            Write-Host "Auditing Power Platform roles..." -ForegroundColor Cyan
            try {
                # Use the Azure AD Power Platform roles function with proper filtering
                $powerPlatformRoles = Get-PowerPlatformAzureADRoleAudit -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -IncludeAzureADRoles:$IncludeOverarchingRolesInServices
                $allResults += $powerPlatformRoles
                Write-Host "✓ Found $($powerPlatformRoles.Count) Power Platform role assignments" -ForegroundColor Green
                if (-not $IncludeOverarchingRolesInServices) {
                    Write-Host "  (Excluding overarching Azure AD roles to prevent duplicates)" -ForegroundColor Gray
                }
            }
            catch {
                Write-Warning "Power Platform audit failed: $($_.Exception.Message)"
                Write-Host "Note: Power Platform has limited certificate authentication support" -ForegroundColor Yellow
            }
        }
        
        # === 3. OPTIONAL DEDUPLICATION (NOT RECOMMENDED WITH PROPER FILTERING) ===
        # if ($DeduplicationMode -ne "None" -and $allResults.Count -gt 0) {
        #     
        #     Write-Host "=== APPLYING OPTIONAL DEDUPLICATION ===" -ForegroundColor Yellow
        #     Write-Host "WARNING: Deduplication should not be needed with proper role filtering!" -ForegroundColor Yellow
        #     Write-Host "Consider using -IncludeOverarchingRolesInServices:$false for cleaner results" -ForegroundColor Yellow
            
        #     $originalCount = $allResults.Count
            
        #     $deduplicationParams = @{
        #         AuditResults = $allResults
        #         DeduplicationMode = $DeduplicationMode
        #         ShowDuplicatesRemoved = $ShowDuplicatesRemoved
        #         PreferAzureADSource = $PreferAzureADSource
        #     }
            
        #     $allResults = Remove-M365AuditDuplicates @deduplicationParams
            
        #     $deduplicatedCount = $allResults.Count
        #     $removedCount = $originalCount - $deduplicatedCount
            
        #     
        #     Write-Host "Deduplication Summary:" -ForegroundColor Yellow
        #     Write-Host "  Original: $originalCount assignments" -ForegroundColor White
        #     Write-Host "  Final: $deduplicatedCount assignments" -ForegroundColor White
        #     Write-Host "  Removed: $removedCount duplicates" -ForegroundColor Yellow
            
        #     if ($removedCount -eq 0) {
        #         Write-Host "  ✓ No duplicates found - role filtering is working correctly!" -ForegroundColor Green
        #     } else {
        #         Write-Host "  ⚠️ $removedCount duplicates removed - consider review role filtering settings" -ForegroundColor Yellow
        #     }
        # }
        
        # === 4. EXPORT AND DISPLAY RESULTS ===
        if ($allResults.Count -gt 0) {
            if ($ExportPath) {
                # $allResults | Export-Csv -Path $ExportPath -NoTypeInformation
            }            
            
            Write-Host "=== AUDIT COMPLETED SUCCESSFULLY ===" -ForegroundColor Green
            Write-Host "Total role assignments found: $($allResults.Count)" -ForegroundColor Green
            
            # if ($DeduplicationMode -ne "None") {
            #     Write-Host "Deduplication applied: $DeduplicationMode mode" -ForegroundColor Cyan
            # } else {
            #     Write-Host "No deduplication applied (recommended with role filtering)" -ForegroundColor Green
            # }
            
            if ($IncludeAnalysis) {
                Write-Host "Results exported to: $ExportPath" -ForegroundColor Green
                Write-Host "Authentication used: Certificate-based" -ForegroundColor Cyan
                
                # === 5. ENHANCED SUMMARY ANALYSIS ===
                
                Write-Host "=== COMPREHENSIVE SUMMARY ANALYSIS ===" -ForegroundColor Cyan
                
                # Service distribution
                $serviceSummary = $allResults | Group-Object Service | Sort-Object Count -Descending
                
                Write-Host "Service Distribution:" -ForegroundColor Yellow
                foreach ($service in $serviceSummary) {
                    $percentage = [math]::Round(($service.Count / $allResults.Count) * 100, 1)
                    Write-Host "  $($service.Name): $($service.Count) assignments ($percentage%)" -ForegroundColor White
                }
                
                # Authentication analysis
                
                Write-Host "Authentication Type Analysis:" -ForegroundColor Yellow
                $authSummary = $allResults | Group-Object AuthenticationType | Sort-Object Count -Descending
                foreach ($authType in $authSummary) {
                    $percentage = [math]::Round(($authType.Count / $allResults.Count) * 100, 1)
                    $color = if ($authType.Name -eq "Certificate") { "Green" } elseif ($authType.Name -eq "ClientSecret") { "Yellow" } else { "White" }
                    Write-Host "  $($authType.Name): $($authType.Count) assignments ($percentage%)" -ForegroundColor $color
                }
                
                # Assignment type analysis
                
                Write-Host "Assignment Type Distribution:" -ForegroundColor Yellow
                $assignmentTypeSummary = $allResults | Group-Object AssignmentType | Sort-Object Count -Descending
                foreach ($assignmentType in $assignmentTypeSummary) {
                    $percentage = [math]::Round(($assignmentType.Count / $allResults.Count) * 100, 1)
                    $color = if ($assignmentType.Name -like "*Eligible*") { "Green" } elseif ($assignmentType.Name -like "*PIM*") { "Cyan" } else { "White" }
                    Write-Host "  $($assignmentType.Name): $($assignmentType.Count) assignments ($percentage%)" -ForegroundColor $color
                }
                
                # Top roles across all services
                
                Write-Host "Top Roles Across All Services:" -ForegroundColor Yellow
                $roleSummary = $allResults | Group-Object RoleName | Sort-Object Count -Descending | Select-Object -First 15
                foreach ($role in $roleSummary) {
                    $services = ($allResults | Where-Object { $_.RoleName -eq $role.Name } | Group-Object Service).Name -join ", "
                    Write-Host "  $($role.Name): $($role.Count) assignments" -ForegroundColor White
                    Write-Host "    Services: $services" -ForegroundColor Gray
                }
                
                # User analysis
                
                Write-Host "Users with Most Role Assignments:" -ForegroundColor Yellow
                $userSummary = $allResults | Where-Object { $_.UserPrincipalName -and $_.UserPrincipalName -ne "Unknown" } | 
                            Group-Object UserPrincipalName | Sort-Object Count -Descending | Select-Object -First 10
                foreach ($user in $userSummary) {
                    $userServices = ($allResults | Where-Object { $_.UserPrincipalName -eq $user.Name } | Group-Object Service).Name -join ", "
                    Write-Host "  $($user.Name): $($user.Count) roles across [$userServices]" -ForegroundColor White
                }
                
                # === 6. SECURITY AND PIM ANALYSIS ===
                
                Write-Host "=== SECURITY ANALYSIS ===" -ForegroundColor Cyan
                
                # Global administrators
                $globalAdmins = $allResults | Where-Object { $_.RoleName -eq "Global Administrator" }
                $globalAdminColor = if ($globalAdmins.Count -le 5) { "Green" } else { "Red" }
                Write-Host "Global Administrators: $($globalAdmins.Count)" -ForegroundColor $globalAdminColor
                if ($globalAdmins.Count -gt 5) {
                    Write-Host "  ⚠️ Consider reducing to 5 or fewer" -ForegroundColor Yellow
                }
                
                # Disabled users with roles
                $disabledUsers = $allResults | Where-Object { $_.UserEnabled -eq $false }
                $disabledColor = if ($disabledUsers.Count -eq 0) { "Green" } else { "Red" }
                Write-Host "Disabled Users with Active Roles: $($disabledUsers.Count)" -ForegroundColor $disabledColor
                if ($disabledUsers.Count -gt 0) {
                    Write-Host "  ⚠️ Review and remove role assignments from disabled accounts" -ForegroundColor Yellow
                }
                
                # PIM analysis
                $pimEligible = $allResults | Where-Object { $_.AssignmentType -like "*Eligible*" }
                $pimActive = $allResults | Where-Object { $_.AssignmentType -like "*Active (PIM*" }
                $permanentActive = $allResults | Where-Object { 
                    $_.AssignmentType -eq "Active" -or 
                    $_.AssignmentType -eq "Azure AD Role" -or
                    $_.AssignmentType -eq "Role Group Member" -or
                    $_.AssignmentType -eq "Intune RBAC"
                }
                
                
                Write-Host "Privileged Identity Management (PIM) Analysis:" -ForegroundColor Yellow
                Write-Host "  PIM Eligible Assignments: $($pimEligible.Count)" -ForegroundColor $(if($pimEligible.Count -gt 0) {"Green"} else {"Yellow"})
                Write-Host "  PIM Active Assignments: $($pimActive.Count)" -ForegroundColor White
                Write-Host "  Permanent Active Assignments: $($permanentActive.Count)" -ForegroundColor White
                
                # PIM adoption rate
                $totalEligibleAndPermanent = $pimEligible.Count + $permanentActive.Count
                if ($totalEligibleAndPermanent -gt 0) {
                    $pimAdoptionRate = [math]::Round(($pimEligible.Count / $totalEligibleAndPermanent) * 100, 1)
                    Write-Host "  PIM Adoption Rate: $pimAdoptionRate%" -ForegroundColor $(if($pimAdoptionRate -gt 30) {"Green"} elseif($pimAdoptionRate -gt 0) {"Yellow"} else {"Red"})
                }
                
                # === 7. SERVICE-SPECIFIC INSIGHTS ===
                
                Write-Host "=== SERVICE-SPECIFIC INSIGHTS ===" -ForegroundColor Cyan
                
                # Intune analysis
                $intuneResults = $allResults | Where-Object { $_.Service -eq "Microsoft Intune" }
                if ($intuneResults.Count -gt 0) {
                    
                    Write-Host "Microsoft Intune Analysis:" -ForegroundColor Yellow
                    $intuneServiceAdmins = $intuneResults | Where-Object { $_.RoleName -eq "Intune Service Administrator" }
                    $intuneRBACAssignments = $intuneResults | Where-Object { $_.RoleType -eq "IntuneRBAC" }
                    $intuneAzureADAssignments = $intuneResults | Where-Object { $_.RoleType -eq "AzureAD" }
                    
                    Write-Host "  Service Administrators: $($intuneServiceAdmins.Count)" -ForegroundColor $(if($intuneServiceAdmins.Count -le 3) {"Green"} else {"Yellow"})
                    Write-Host "  Intune RBAC Assignments: $($intuneRBACAssignments.Count)" -ForegroundColor White
                    Write-Host "  Azure AD Role Assignments: $($intuneAzureADAssignments.Count)" -ForegroundColor White
                    
                    if ($intuneServiceAdmins.Count -gt 3) {
                        Write-Host "  ⚠️ Consider reducing Intune Service Administrator count" -ForegroundColor Yellow
                    }
                    if ($intuneAzureADAssignments.Count -gt $intuneRBACAssignments.Count) {
                        Write-Host "  ⚠️ Consider using more Intune RBAC roles for granular permissions" -ForegroundColor Yellow
                    }
                }
                
                # Exchange analysis
                $exchangeResults = $allResults | Where-Object { $_.Service -eq "Exchange Online" }
                if ($exchangeResults.Count -gt 0) {
                    
                    Write-Host "Exchange Online Analysis:" -ForegroundColor Yellow
                    $orgManagement = $exchangeResults | Where-Object { $_.RoleName -eq "Organization Management" }
                    $roleGroups = $exchangeResults | Where-Object { $_.AssignmentType -eq "Role Group Member" }
                    $azureADRoles = $exchangeResults | Where-Object { $_.RoleSource -eq "AzureAD" }
                    $onPremSynced = $exchangeResults | Where-Object { $_.OnPremisesSyncEnabled -eq $true }
                    
                    Write-Host "  Organization Management Members: $($orgManagement.Count)" -ForegroundColor White
                    Write-Host "  Role Group Assignments: $($roleGroups.Count)" -ForegroundColor White
                    Write-Host "  Azure AD Role Assignments: $($azureADRoles.Count)" -ForegroundColor White
                    Write-Host "  On-Premises Synced Objects: $($onPremSynced.Count)" -ForegroundColor $(if($onPremSynced.Count -gt 0) {"Cyan"} else {"White"})
                    
                    if ($onPremSynced.Count -gt 0) {
                        Write-Host "  ✓ Hybrid environment detected" -ForegroundColor Green
                    }
                }
                
                # SharePoint analysis
                $sharePointResults = $allResults | Where-Object { $_.Service -eq "SharePoint Online" }
                if ($sharePointResults.Count -gt 0) {
                    
                    Write-Host "SharePoint Online Analysis:" -ForegroundColor Yellow
                    $siteAdmins = $sharePointResults | Where-Object { $_.RoleName -like "*Site*Administrator*" }
                    $appCatalogAdmins = $sharePointResults | Where-Object { $_.RoleName -eq "App Catalog Administrator" }
                    $uniqueSites = ($sharePointResults | Where-Object { $_.SiteTitle } | Select-Object -Unique SiteTitle).Count
                    
                    Write-Host "  Site Collection Administrators: $($siteAdmins.Count)" -ForegroundColor White
                    Write-Host "  App Catalog Administrators: $($appCatalogAdmins.Count)" -ForegroundColor White
                    Write-Host "  Unique Sites with Assignments: $uniqueSites" -ForegroundColor White
                }
                
                # === 8. RECOMMENDATIONS ===
                
                Write-Host "=== SECURITY RECOMMENDATIONS ===" -ForegroundColor Green
                
                $recommendations = @()
                
                # Critical recommendations
                if ($globalAdmins.Count -gt 5) {
                    $recommendations += "🔴 CRITICAL: Reduce Global Administrator count from $($globalAdmins.Count) to 5 or fewer"
                }
                if ($disabledUsers.Count -gt 0) {
                    $recommendations += "🔴 HIGH: Remove role assignments from $($disabledUsers.Count) disabled user accounts"
                }
                
                # Medium priority recommendations
                if ($pimEligible.Count -eq 0 -and $permanentActive.Count -gt 0) {
                    $recommendations += "🟡 MEDIUM: Consider implementing PIM for eligible assignments"
                }
                
                $clientSecretAuth = $allResults | Where-Object { $_.AuthenticationType -eq "ClientSecret" }
                if ($clientSecretAuth.Count -gt 0) {
                    $recommendations += "🟡 MEDIUM: Migrate $($clientSecretAuth.Count) client secret authentications to certificate-based"
                }
                
                # Service-specific recommendations
                if ($intuneResults.Count -gt 0 -and $intuneServiceAdmins.Count -gt 3) {
                    $recommendations += "🟡 MEDIUM: Consider reducing Intune Service Administrator count"
                }
                
                # Positive findings
                $positiveFindings = @()
                if ($globalAdmins.Count -le 5) {
                    $positiveFindings += "✅ Global Administrator count is within recommended limits"
                }
                if ($disabledUsers.Count -eq 0) {
                    $positiveFindings += "✅ No disabled users with active role assignments"
                }
                if ($pimEligible.Count -gt 0) {
                    $positiveFindings += "✅ PIM eligible assignments detected"
                }
                $certificateAuth = $allResults | Where-Object { $_.AuthenticationType -eq "Certificate" }
                if ($certificateAuth.Count -eq $allResults.Count) {
                    $positiveFindings += "✅ All authentications use secure certificate-based method"
                }
                
                # Display recommendations
                if ($recommendations.Count -gt 0) {
                    Write-Host "Action Required:" -ForegroundColor Red
                    foreach ($rec in $recommendations) {
                        Write-Host "  $rec" -ForegroundColor White
                    }
                }
                
                if ($positiveFindings.Count -gt 0) {
                    
                    Write-Host "Positive Security Findings:" -ForegroundColor Green
                    foreach ($finding in $positiveFindings) {
                        Write-Host "  $finding" -ForegroundColor White
                    }
                }
                
                # General recommendations
                
                Write-Host "General Security Best Practices:" -ForegroundColor Cyan
                Write-Host "• Implement regular access reviews for privileged roles" -ForegroundColor White
                Write-Host "• Enable PIM for eligible assignments to reduce standing privileges" -ForegroundColor White
                Write-Host "• Monitor privileged role assignments with automated alerts" -ForegroundColor White
                Write-Host "• Implement break-glass emergency access procedures" -ForegroundColor White
                Write-Host "• Regularly rotate certificates (recommended 12-24 months)" -ForegroundColor White
                Write-Host "• Use service-specific RBAC roles instead of broad administrative roles" -ForegroundColor White
                
                # === 9. ENHANCED REPORTING OPTIONS ===
                
                Write-Host "=== ENHANCED REPORTING OPTIONS ===" -ForegroundColor Cyan
                Write-Host "PowerShell Commands for Additional Analysis:" -ForegroundColor Yellow
                Write-Host "• Export-M365AuditHtmlReport -AuditResults `$results" -ForegroundColor White
                Write-Host "• Export-M365AuditJsonReport -AuditResults `$results" -ForegroundColor White
                Write-Host "• Export-M365AuditExcelReport -AuditResults `$results" -ForegroundColor White
                Write-Host "• Get-M365RoleAnalysis -AuditResults `$results" -ForegroundColor White
                Write-Host "• Get-M365ComplianceGaps -AuditResults `$results" -ForegroundColor White
                
                
                Write-Host "✓ Enhanced comprehensive audit completed successfully!" -ForegroundColor Green
                Write-Host "No deduplication required with proper role filtering architecture" -ForegroundColor Green
            }
        }
        else {
                Write-Host "No role assignments found." -ForegroundColor Yellow
        }
        
        
        return $allResults
    }
    catch {
        Write-Error "Error during comprehensive audit: $($_.Exception.Message)"
        Write-Error "Stack trace: $($_.ScriptStackTrace)"
        
        # Enhanced troubleshooting guidance
        if ($_.Exception.Message -like "*certificate*") {
            Write-Host "================================================================="
            Write-Host "Certificate troubleshooting:" -ForegroundColor Yellow
            Write-Host "• Verify certificate exists in Windows Certificate Store" -ForegroundColor White
            Write-Host "• Check certificate expiration date" -ForegroundColor White
            Write-Host "• Ensure certificate is uploaded to Azure AD app registration" -ForegroundColor White
            Write-Host "• Run: Get-M365AuditCurrentConfig to verify setup" -ForegroundColor White
        }
        elseif ($_.Exception.Message -like "*permission*" -or $_.Exception.Message -like "*access*") {
            
            Write-Host "Permission troubleshooting:" -ForegroundColor Yellow
            Write-Host "• Run: Get-M365AuditRequiredPermissions" -ForegroundColor White
            Write-Host "• Verify admin consent has been granted in Azure AD" -ForegroundColor White
            Write-Host "• Check if app registration has required API permissions" -ForegroundColor White
        }
        
        return @()
    }
}
