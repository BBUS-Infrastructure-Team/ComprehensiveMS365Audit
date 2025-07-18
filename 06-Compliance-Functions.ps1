# 06-Compliance-Functions.ps1
# Focused Microsoft Purview Administrative Role Audit Function
# Updated to properly separate Azure AD roles from Compliance Center role groups

function Get-PurviewRoleAudit {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Organization,

        [string]$TenantId,
        [string]$ClientId,
        [string]$CertificateThumbprint,
        [switch]$IncludeAzureADRoles  # Include overarching Azure AD Purview roles
    )
    
    $results = @()
    
    try {
        # Set app credentials if provided
        if ($TenantId -and $ClientId -and $CertificateThumbprint) {
            Set-M365AuditCertCredentials -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint
        }
        
        # Verify certificate authentication
        if (-not $script:AppConfig.UseAppAuth -or $script:AppConfig.AuthType -ne "Certificate") {
            Write-Warning "Certificate authentication is required for Compliance role audit"
            return $results
        }
        
        # === ENHANCED AZURE AD ROLE FILTERING ===
        # Purview-specific Azure AD administrative roles (NOT overarching roles)
        $purviewSpecificRoles = @(
            "Compliance Administrator",        # Purview-focused
            "Compliance Data Administrator",   # Purview-focused  
            "eDiscovery Administrator",
            "eDiscovery Manager", 
            "Information Protection Administrator",
            "Information Protection Analyst",
            "Information Protection Investigator",
            "Information Protection Reader",
            "Role Management",                 # Clearly administrative
            "Organization Configuration",      # Clearly administrative  
            "Supervisory Review Administrator" # Clearly administrative
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
            $purviewSpecificRoles + $overarchingRoles
        } else {
            $purviewSpecificRoles
        }
        
        # === GET PURVIEW-RELATED AZURE AD ADMINISTRATIVE ROLES ===
        if ($rolesToInclude.Count -gt 0) {
            Write-Host "Retrieving Purview-related Azure AD administrative roles..." -ForegroundColor Cyan
            
            # Connect to Microsoft Graph
            $context = Get-MgContext
            if (-not $context -or $context.AuthType -ne "AppOnly") {
                $null = Connect-MgGraph -TenantId $script:AppConfig.TenantId -ClientId $script:AppConfig.ClientId -CertificateThumbprint $script:AppConfig.CertificateThumbprint -NoWelcome
            }
            
            $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition | Where-Object { $_.DisplayName -in $rolesToInclude }
            $assignments = Get-MgRoleManagementDirectoryRoleAssignment | Where-Object { $_.RoleDefinitionId -in $roleDefinitions.Id }
            
            Write-Host "Found $($assignments.Count) Azure AD Purview administrative role assignments" -ForegroundColor Green
            
            foreach ($assignment in $assignments) {
                try {
                    $roleDefinition = $roleDefinitions | Where-Object { $_.Id -eq $assignment.RoleDefinitionId }
                    $user = Get-MgUser -UserId $assignment.PrincipalId -Property "UserPrincipalName,DisplayName,AccountEnabled,SignInActivity" -ErrorAction SilentlyContinue
                    
                    if ($user) {
                        # Determine role scope for enhanced deduplication
                        $roleScope = if ($roleDefinition.DisplayName -in $overarchingRoles) { "Overarching" } else { "Service-Specific" }
                        
                        $results += [PSCustomObject]@{
                            Service = "Microsoft Purview"
                            UserPrincipalName = $user.UserPrincipalName
                            DisplayName = $user.DisplayName
                            UserId = $assignment.PrincipalId
                            RoleName = $roleDefinition.DisplayName
                            RoleDefinitionId = $assignment.RoleDefinitionId
                            RoleScope = $roleScope  # New property for enhanced deduplication
                            AssignmentType = "Azure AD Role"
                            AssignedDateTime = $assignment.CreatedDateTime
                            UserEnabled = $user.AccountEnabled
                            LastSignIn = $user.SignInActivity.LastSignInDateTime
                            Scope = "Organization"
                            AssignmentId = $assignment.Id
                            AuthenticationType = "Certificate"
                            PrincipalType = "User"
                            RoleSource = "AzureAD"
                            RoleGroupDescription = $roleDefinition.Description
                        }
                    }
                }
                catch {
                    Write-Verbose "Error processing Azure AD Purview assignment: $($_.Exception.Message)"
                }
            }
        }
        
        Write-Host "✓ Purview administrative role audit completed. Found $($results.Count) administrative role assignments" -ForegroundColor Green
        
        # Provide feedback about role filtering
        if (-not $IncludeAzureADRoles) {
            Write-Host "  (Excluding overarching Azure AD roles - use -IncludeAzureADRoles to include)" -ForegroundColor Yellow
        }
        
        # Display focused results
        Write-Host ""
        Write-Host "=== Purview Administrative Role Audit Summary ===" -ForegroundColor Green
        Write-Host "Total administrative assignments found: $($results.Count)" -ForegroundColor White
        
        if ($results.Count -gt 0) {
            $roleSummary = $results | Group-Object RoleName | Sort-Object Count -Descending
            #$sourceSummary = $results | Group-Object RoleSource
            $scopeSummary = $results | Group-Object RoleScope
            
            Write-Host "Top administrative roles:" -ForegroundColor Cyan
            foreach ($role in $roleSummary | Select-Object -First 10) {
                Write-Host "  $($role.Name): $($role.Count) members" -ForegroundColor White
            }
            
            Write-Host "Role scope:" -ForegroundColor Cyan
            foreach ($scope in $scopeSummary) {
                Write-Host "  $($scope.Name): $($scope.Count)" -ForegroundColor White
            }
        }
        
        Write-Host ""
        Write-Host "=== SCOPE CLARIFICATION ===" -ForegroundColor Green
        Write-Host "✓ Focused on Purview/Compliance Azure AD administrative roles only" -ForegroundColor Green
        Write-Host "✓ Included: Compliance Administrator, eDiscovery roles, Information Protection roles" -ForegroundColor Green
        Write-Host "✓ Included: Role Management, Organization Configuration, Supervisory Review Administrator" -ForegroundColor Green
        Write-Host "✓ Excluded: Operational roles (Content Search, Export, Preview, etc.)" -ForegroundColor Green
        Write-Host "✓ Excluded: Read-only roles (View-Only Audit Logs, View-Only Configuration, etc.)" -ForegroundColor Green
        Write-Host "✓ No Exchange PowerShell connection required - pure Azure AD role audit" -ForegroundColor Green
        
        Write-Host ""
        Write-Host "=== Performance Summary ===" -ForegroundColor Green
        Write-Host "Authentication: Certificate-based Azure AD only (Secure)" -ForegroundColor White
        Write-Host "Role System: Azure AD administrative roles only" -ForegroundColor White
        Write-Host "Scope: Administrative functions only" -ForegroundColor White
        
    }
    catch {
        Write-Warning "Error auditing Purview administrative roles: $($_.Exception.Message)"
        throw
    }
    
    return $results
}