# 09-Intune-Functions.ps1
# Enhanced Microsoft Intune Role Audit Function focused ONLY on administrative roles
# Removed: Policy ownership tracking, configuration ownership tracking
# Updated Get-IntuneRoleAudit function for 09-Intune-Functions.ps1

function Get-IntuneRoleAudit {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$CertificateThumbprint,
        [switch]$IncludePIM,
        [switch]$IncludeAnalysis,
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
            throw "Certificate authentication is required for Intune role audit. Use Set-M365AuditCertCredentials first."
        }
        
        # Connect to Microsoft Graph with certificate authentication
        $context = Get-MgContext
        if (-not $context -or $context.AuthType -ne "AppOnly") {
            Write-Host "Connecting to Microsoft Graph with certificate authentication..." -ForegroundColor Yellow
            
            Connect-MgGraph -TenantId $script:AppConfig.TenantId -ClientId $script:AppConfig.ClientId -CertificateThumbprint $script:AppConfig.CertificateThumbprint -NoWelcome
            
            # Verify app-only authentication
            $context = Get-MgContext
            if ($context.AuthType -ne "AppOnly") {
                throw "Expected app-only authentication but got: $($context.AuthType). Check certificate configuration."
            }
            
            Write-Host "✓ Connected with certificate authentication" -ForegroundColor Green
            Write-Host "  App Name: $($context.AppName)" -ForegroundColor Gray
        }
        else {
            Write-Host "✓ Already connected to Microsoft Graph with app-only authentication" -ForegroundColor Green
        }
        
        # Verify Graph connection and required permissions
        try {
            $null = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/deviceManagement/roleDefinitions" -Method GET -ErrorAction Stop
            Write-Host "✓ Intune Graph API access verified" -ForegroundColor Green
        }
        catch {
            Write-Warning "Intune Graph API access failed: $($_.Exception.Message)"
            Write-Host "Required permissions:" -ForegroundColor Yellow
            Write-Host "• DeviceManagementRBAC.Read.All" -ForegroundColor White
            Write-Host "• DeviceManagementConfiguration.Read.All" -ForegroundColor White
            Write-Host "• Directory.Read.All" -ForegroundColor White
            Write-Host "• RoleEligibilitySchedule.Read.Directory (for PIM)" -ForegroundColor White
            Write-Host "• RoleAssignmentSchedule.Read.Directory (for PIM)" -ForegroundColor White
            throw "Required permissions not granted or certificate authentication failed"
        }
        
        # === ENHANCED AZURE AD ROLE FILTERING ===
        # Intune-specific Azure AD roles (NOT overarching roles)
        $intuneSpecificRoles = @(
            "Device Managers", 
            "Device Users",
            "Device Administrators",
            "Endpoint Privilege Manager Administrator"
        )
        
        # Overarching roles that should only appear in Azure AD audit
        $overarchingRoles = @(
            "Global Administrator",
            "Intune Service Administrator",
            "Security Administrator",
            "Security Reader",
            "Cloud Application Administrator",
            "Application Administrator",
            "Privileged Authentication Administrator",
            "Privileged Role Administrator"
        )
        
        # Determine which roles to include based on parameter
        $rolesToInclude = if ($IncludeAzureADRoles) {
            $intuneSpecificRoles + $overarchingRoles
        } else {
            $intuneSpecificRoles
        }
        
        # Get Intune-related Azure AD administrative roles first
        Write-Host "Retrieving Intune-related Azure AD administrative roles..." -ForegroundColor Cyan
        
        $intuneRoleDefinitions = @()
        
        try {
            $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition
            $intuneRoleDefinitions = $roleDefinitions | Where-Object { $_.DisplayName -in $rolesToInclude }
            $assignments = Get-MgRoleManagementDirectoryRoleAssignment | Where-Object { $_.RoleDefinitionId -in $intuneRoleDefinitions.Id }
            
            Write-Host "Found $($assignments.Count) active Azure AD Intune administrative role assignments" -ForegroundColor Green
            
            foreach ($assignment in $assignments) {
                try {
                    $user = Get-MgUser -UserId $assignment.PrincipalId -ErrorAction SilentlyContinue
                    if (-not $user) { continue }
                    
                    $role = $intuneRoleDefinitions | Where-Object { $_.Id -eq $assignment.RoleDefinitionId }
                    
                    # Determine role scope for enhanced deduplication
                    $roleScope = if ($role.DisplayName -in $overarchingRoles) { "Overarching" } else { "Service-Specific" }
                    
                    $results += [PSCustomObject]@{
                        Service = "Microsoft Intune"
                        UserPrincipalName = $user.UserPrincipalName
                        DisplayName = $user.DisplayName
                        UserId = $user.Id
                        RoleName = $role.DisplayName
                        RoleDefinitionId = $assignment.RoleDefinitionId
                        RoleScope = $roleScope  # New property for enhanced deduplication
                        AssignmentType = "Azure AD Role"
                        AssignedDateTime = $assignment.CreatedDateTime
                        UserEnabled = $user.AccountEnabled
                        LastSignIn = $user.SignInActivity.LastSignInDateTime
                        Scope = $assignment.DirectoryScopeId
                        AssignmentId = $assignment.Id
                        RoleType = "AzureAD"
                        AuthenticationType = "Certificate"
                        PrincipalType = "User"
                    }
                }
                catch {
                    Write-Verbose "Error processing Azure AD Intune assignment: $($_.Exception.Message)"
                }
            }
        }
        catch {
            Write-Warning "Error retrieving Intune Azure AD administrative roles: $($_.Exception.Message)"
        }
        
        # Get PIM eligible assignments for Intune Azure AD roles
        if ($IncludePIM) {
            Write-Host "Retrieving PIM eligible assignments for Intune administrative roles..." -ForegroundColor Cyan
            try {
                $intuneRoleIds = $intuneRoleDefinitions.Id
                $eligibleAssignments = Get-MgRoleManagementDirectoryRoleEligibilitySchedule | Where-Object { 
                    $_.RoleDefinitionId -in $intuneRoleIds 
                }
                
                Write-Host "Found $($eligibleAssignments.Count) PIM eligible assignments for Intune roles" -ForegroundColor Green
                
                foreach ($assignment in $eligibleAssignments) {
                    try {
                        $user = Get-MgUser -UserId $assignment.PrincipalId -ErrorAction SilentlyContinue
                        if (-not $user) { continue }
                        
                        $role = $intuneRoleDefinitions | Where-Object { $_.Id -eq $assignment.RoleDefinitionId }
                        
                        # Determine role scope for enhanced deduplication
                        $roleScope = if ($role.DisplayName -in $overarchingRoles) { "Overarching" } else { "Service-Specific" }
                        
                        $results += [PSCustomObject]@{
                            Service = "Microsoft Intune"
                            UserPrincipalName = $user.UserPrincipalName
                            DisplayName = $user.DisplayName
                            UserId = $user.Id
                            RoleName = $role.DisplayName
                            RoleDefinitionId = $assignment.RoleDefinitionId
                            RoleScope = $roleScope  # New property for enhanced deduplication
                            AssignmentType = "Eligible (PIM)"
                            AssignedDateTime = $assignment.CreatedDateTime
                            UserEnabled = $user.AccountEnabled
                            LastSignIn = $user.SignInActivity.LastSignInDateTime
                            Scope = $assignment.DirectoryScopeId
                            AssignmentId = $assignment.Id
                            RoleType = "AzureAD"
                            PIMEndDateTime = $assignment.ScheduleInfo.Expiration.EndDateTime
                            PIMStartDateTime = $assignment.ScheduleInfo.Expiration.StartDateTime
                            PIMDuration = $assignment.ScheduleInfo.Expiration.Duration
                            AuthenticationType = "Certificate"
                            PrincipalType = "User"
                        }
                    }
                    catch {
                        Write-Verbose "Error processing PIM Intune assignment: $($_.Exception.Message)"
                    }
                }
            }
            catch {
                Write-Verbose "Error retrieving PIM eligible assignments for Intune: $($_.Exception.Message)"
                Write-Host "Note: PIM may require additional permissions or licensing" -ForegroundColor Yellow
            }

            # Get PIM active assignments for Intune roles
            Write-Host "Retrieving PIM active assignments for Intune administrative roles..." -ForegroundColor Cyan
            try {
                $activeAssignments = Get-MgRoleManagementDirectoryRoleAssignmentSchedule | Where-Object { 
                    $_.RoleDefinitionId -in $intuneRoleIds -and $_.AssignmentType -eq "Activated"
                }
                
                Write-Host "Found $($activeAssignments.Count) PIM active assignments for Intune roles" -ForegroundColor Green
                
                foreach ($assignment in $activeAssignments) {
                    try {
                        $user = Get-MgUser -UserId $assignment.PrincipalId -ErrorAction SilentlyContinue
                        if (-not $user) { continue }
                        
                        $role = $intuneRoleDefinitions | Where-Object { $_.Id -eq $assignment.RoleDefinitionId }
                        
                        # Determine role scope for enhanced deduplication
                        $roleScope = if ($role.DisplayName -in $overarchingRoles) { "Overarching" } else { "Service-Specific" }
                        
                        $results += [PSCustomObject]@{
                            Service = "Microsoft Intune"
                            UserPrincipalName = $user.UserPrincipalName
                            DisplayName = $user.DisplayName
                            UserId = $user.Id
                            RoleName = $role.DisplayName
                            RoleDefinitionId = $assignment.RoleDefinitionId
                            RoleScope = $roleScope  # New property for enhanced deduplication
                            AssignmentType = "Active (PIM Activated)"
                            AssignedDateTime = $assignment.CreatedDateTime
                            UserEnabled = $user.AccountEnabled
                            LastSignIn = $user.SignInActivity.LastSignInDateTime
                            Scope = $assignment.DirectoryScopeId
                            AssignmentId = $assignment.Id
                            RoleType = "AzureAD"
                            PIMActivatedDateTime = $assignment.ActivatedDateTime
                            PIMEndDateTime = $assignment.ScheduleInfo.Expiration.EndDateTime
                            AuthenticationType = "Certificate"
                            PrincipalType = "User"
                        }
                    }
                    catch {
                        Write-Verbose "Error processing PIM active Intune assignment: $($_.Exception.Message)"
                    }
                }
            }
            catch {
                Write-Verbose "Error retrieving PIM active assignments for Intune: $($_.Exception.Message)"
            }
        }
        
        # Enhanced Get-IntuneRoleAudit function with workaround for missing roleDefinition references
        # Replace the Intune RBAC section with this improved version

        # Get Intune RBAC role definitions (ADMINISTRATIVE ROLES ONLY)
        Write-Host "Retrieving Intune RBAC administrative role definitions..." -ForegroundColor Cyan
        try {
            $intuneRBACRoleDefinitions = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/deviceManagement/roleDefinitions" -Method GET
            Write-Host "Found $($intuneRBACRoleDefinitions.value.Count) Intune role definitions" -ForegroundColor Green
            
            # Create lookup hashtables for faster role definition lookups
            $roleDefinitionLookup = @{}
            $roleNameToIdLookup = @{}
            
            foreach ($roleDef in $intuneRBACRoleDefinitions.value) {
                $roleDefinitionLookup[$roleDef.id] = $roleDef
                $roleNameToIdLookup[$roleDef.displayName] = $roleDef.id
                Write-Verbose "Added role to lookup: $($roleDef.displayName) (ID: $($roleDef.id))"
            }
            
            Write-Host "Role definitions loaded into lookup tables" -ForegroundColor Green
        }
        catch {
            Write-Warning "Could not retrieve Intune role definitions: $($_.Exception.Message)"
            throw "Certificate authentication may lack required permissions"
        }

        # Get Intune RBAC role assignments (ADMINISTRATIVE ASSIGNMENTS ONLY)
        Write-Host "Retrieving Intune RBAC administrative role assignments..." -ForegroundColor Cyan
        try {
            $intuneRoleAssignments = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/deviceManagement/roleAssignments" -Method GET
            Write-Host "Found $($intuneRoleAssignments.value.Count) Intune administrative role assignments" -ForegroundColor Green
            
            foreach ($assignment in $intuneRoleAssignments.value) {
                try {
                    Write-Verbose "Processing assignment ID: $($assignment.id)"
                    Write-Verbose "Assignment displayName: $($assignment.displayName)"
                    
                    # ENHANCED ROLE DEFINITION RESOLUTION with multiple fallback strategies
                    $roleDefinition = $null
                    $roleResolutionMethod = "Unknown"
                    
                    # Method 1: Try standard roleDefinition.id property
                    if ($assignment.roleDefinition -and $assignment.roleDefinition.id) {
                        $roleDefinitionId = $assignment.roleDefinition.id
                        $roleDefinition = $roleDefinitionLookup[$roleDefinitionId]
                        $roleResolutionMethod = "Standard roleDefinition.id"
                        Write-Verbose "Method 1 - Found role via roleDefinition.id: $roleDefinitionId"
                    }
                    
                    # Method 2: Try direct roleDefinitionId property
                    elseif ($assignment.roleDefinitionId) {
                        $roleDefinitionId = $assignment.roleDefinitionId
                        $roleDefinition = $roleDefinitionLookup[$roleDefinitionId]
                        $roleResolutionMethod = "Direct roleDefinitionId"
                        Write-Verbose "Method 2 - Found role via roleDefinitionId: $roleDefinitionId"
                    }
                    
                    # Method 3: Try @odata.id extraction
                    elseif ($assignment.roleDefinition -and $assignment.roleDefinition.'@odata.id') {
                        $odataId = $assignment.roleDefinition.'@odata.id'
                        if ($odataId -match "/roleDefinitions/([^/]+)") {
                            $roleDefinitionId = $matches[1]
                            $roleDefinition = $roleDefinitionLookup[$roleDefinitionId]
                            $roleResolutionMethod = "@odata.id extraction"
                            Write-Verbose "Method 3 - Found role via @odata.id: $roleDefinitionId"
                        }
                    }
                    
                    # Method 4: WORKAROUND - Try to match by assignment displayName
                    elseif ($assignment.displayName) {
                        Write-Verbose "Method 4 - Attempting displayName matching for: $($assignment.displayName)"
                        
                        # Try exact match first
                        if ($roleNameToIdLookup.ContainsKey($assignment.displayName)) {
                            $roleDefinitionId = $roleNameToIdLookup[$assignment.displayName]
                            $roleDefinition = $roleDefinitionLookup[$roleDefinitionId]
                            $roleResolutionMethod = "Exact displayName match"
                            Write-Verbose "Method 4a - Exact match found: $($assignment.displayName) -> $roleDefinitionId"
                        }
                        else {
                            # Try to find partial matches (remove common prefixes/suffixes)
                            $cleanDisplayName = $assignment.displayName -replace "^Intune\s+", "" -replace "\s+\(Group\)$", ""
                            
                            if ($roleNameToIdLookup.ContainsKey($cleanDisplayName)) {
                                $roleDefinitionId = $roleNameToIdLookup[$cleanDisplayName]
                                $roleDefinition = $roleDefinitionLookup[$roleDefinitionId]
                                $roleResolutionMethod = "Cleaned displayName match ($cleanDisplayName)"
                                Write-Verbose "Method 4b - Cleaned match found: $cleanDisplayName -> $roleDefinitionId"
                            }
                            else {
                                # Try fuzzy matching by looking for roles that contain key words
                                $possibleMatches = $roleNameToIdLookup.Keys | Where-Object { 
                                    $_ -like "*$($cleanDisplayName.Split(' ')[0])*" -or
                                    $cleanDisplayName -like "*$($_.Split(' ')[0])*"
                                }
                                
                                if ($possibleMatches.Count -eq 1) {
                                    $matchedRoleName = $possibleMatches[0]
                                    $roleDefinitionId = $roleNameToIdLookup[$matchedRoleName]
                                    $roleDefinition = $roleDefinitionLookup[$roleDefinitionId]
                                    $roleResolutionMethod = "Fuzzy match: $matchedRoleName"
                                    Write-Verbose "Method 4c - Fuzzy match found: $matchedRoleName -> $roleDefinitionId"
                                }
                                elseif ($possibleMatches.Count -gt 1) {
                                    Write-Verbose "Method 4c - Multiple fuzzy matches found: $($possibleMatches -join ', ')"
                                }
                            }
                        }
                    }
                    
                    # Method 5: Try individual API fetch (last resort)
                    if (-not $roleDefinition -and $assignment.roleDefinition) {
                        try {
                            Write-Verbose "Method 5 - Attempting individual role definition fetch..."
                            if ($assignment.roleDefinition.id) {
                                $roleDefResponse = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/deviceManagement/roleDefinitions/$($assignment.roleDefinition.id)" -Method GET -ErrorAction SilentlyContinue
                                $roleDefinition = $roleDefResponse
                                $roleResolutionMethod = "Individual API fetch"
                                Write-Verbose "Method 5 - Successfully fetched: $($roleDefinition.displayName)"
                            }
                        }
                        catch {
                            Write-Verbose "Method 5 failed: $($_.Exception.Message)"
                        }
                    }
                    
                    # If we still can't resolve the role definition, create a warning but continue
                    if (-not $roleDefinition) {
                        Write-Warning "Could not resolve role definition for assignment ID: $($assignment.id)"
                        Write-Verbose "Assignment displayName: $($assignment.displayName)"
                        Write-Verbose "Available methods tried: Standard, Direct, @odata, DisplayName matching, Individual fetch"
                        Write-Verbose "Available role names: $($roleNameToIdLookup.Keys -join ', ')"
                        
                        # Create a placeholder role definition so we don't lose the assignment data
                        $roleDefinition = [PSCustomObject]@{
                            id = "Unknown"
                            displayName = if ($assignment.displayName) { 
                                "$($assignment.displayName) (Unresolved)" 
                            } else { 
                                "Unknown Role (Assignment: $($assignment.id))" 
                            }
                            description = "Role definition could not be resolved - assignment reference may be missing"
                            isBuiltIn = $null
                        }
                        $roleResolutionMethod = "Placeholder (unresolved)"
                    }
                    
                    Write-Verbose "Role resolution result: $($roleDefinition.displayName) via $roleResolutionMethod"
                    
                    # Process each member in the assignment
                    if (-not $assignment.members -or $assignment.members.Count -eq 0) {
                        Write-Verbose "Assignment $($assignment.id) has no members"
                        continue
                    }
                    
                    foreach ($member in $assignment.members) {
                        try {
                            # Resolve member details (user, group, or service principal)
                            $user = $null
                            $userDisplayName = "Unknown"
                            $userPrincipalName = "Unknown"
                            $userEnabled = $null
                            $lastSignIn = $null
                            $principalType = "Unknown"
                            
                            # Try as user first
                            try {
                                $user = Get-MgUser -UserId $member -ErrorAction SilentlyContinue
                                if ($user) {
                                    $userDisplayName = $user.DisplayName
                                    $userPrincipalName = $user.UserPrincipalName
                                    $userEnabled = $user.AccountEnabled
                                    $lastSignIn = $user.SignInActivity.LastSignInDateTime
                                    $principalType = "User"
                                    Write-Verbose "Resolved as user: $userPrincipalName"
                                }
                            }
                            catch {
                                Write-Verbose "Not a user: $($_.Exception.Message)"
                            }
                            
                            # Try as group if not a user
                            if (-not $user) {
                                try {
                                    $group = Get-MgGroup -GroupId $member -ErrorAction SilentlyContinue
                                    if ($group) {
                                        $userDisplayName = $group.DisplayName + " (Group)"
                                        $userPrincipalName = $group.Mail
                                        $principalType = "Group"
                                        Write-Verbose "Resolved as group: $($group.DisplayName)"
                                    }
                                }
                                catch {
                                    Write-Verbose "Not a group: $($_.Exception.Message)"
                                }
                            }
                            
                            # Try as service principal if still not resolved
                            if (-not $user -and $principalType -eq "Unknown") {
                                try {
                                    $sp = Get-MgServicePrincipal -ServicePrincipalId $member -ErrorAction SilentlyContinue
                                    if ($sp) {
                                        $userDisplayName = $sp.DisplayName + " (Service Principal)"
                                        $userPrincipalName = $sp.AppId
                                        $principalType = "ServicePrincipal"
                                        Write-Verbose "Resolved as service principal: $($sp.DisplayName)"
                                    }
                                }
                                catch {
                                    Write-Verbose "Not a service principal: $($_.Exception.Message)"
                                }
                            }
                            
                            # If still not resolved, use the member ID
                            if ($principalType -eq "Unknown") {
                                $userDisplayName = "Unresolved Principal"
                                $userPrincipalName = $member
                                $principalType = "Unknown"
                                Write-Verbose "Could not resolve principal: $member"
                            }
                            
                            # Get scope information
                            $scopeInfo = "Organization"
                            $scopeDetails = @()
                            
                            if ($assignment.resourceScopes -and $assignment.resourceScopes.Count -gt 0) {
                                foreach ($scopeId in $assignment.resourceScopes) {
                                    try {
                                        if ($scopeId -eq "/" -or $scopeId -eq "") {
                                            $scopeDetails += "Root"
                                        }
                                        else {
                                            # Try as group first
                                            $scopeGroup = Get-MgGroup -GroupId $scopeId -ErrorAction SilentlyContinue
                                            if ($scopeGroup) {
                                                $scopeDetails += $scopeGroup.DisplayName
                                            }
                                            else {
                                                $scopeDetails += $scopeId
                                            }
                                        }
                                    }
                                    catch {
                                        $scopeDetails += $scopeId
                                    }
                                }
                                $scopeInfo = $scopeDetails -join ", "
                            }
                            
                            # Check for time-bound assignments
                            $assignmentType = "Intune RBAC"
                            $pimEndDateTime = $null
                            $pimStartDateTime = $null
                            
                            if ($assignment.scheduleInfo -and $assignment.scheduleInfo.expiration.endDateTime) {
                                $assignmentType = "Time-bound RBAC"
                                $pimEndDateTime = $assignment.scheduleInfo.expiration.endDateTime
                                $pimStartDateTime = $assignment.scheduleInfo.startDateTime
                            }
                            
                            # Create the result object with enhanced information
                            $results += [PSCustomObject]@{
                                Service = "Microsoft Intune"
                                UserPrincipalName = $userPrincipalName
                                DisplayName = $userDisplayName
                                UserId = $member
                                RoleName = $roleDefinition.displayName
                                RoleDefinitionId = $roleDefinition.id
                                RoleScope = "Service-Specific"
                                AssignmentType = $assignmentType
                                AssignedDateTime = $assignment.createdDateTime
                                UserEnabled = $userEnabled
                                LastSignIn = $lastSignIn
                                Scope = $scopeInfo
                                AssignmentId = $assignment.id
                                RoleType = "IntuneRBAC"
                                RoleDescription = $roleDefinition.description
                                IsBuiltIn = $roleDefinition.isBuiltIn
                                PIMEndDateTime = $pimEndDateTime
                                PIMStartDateTime = $pimStartDateTime
                                AuthenticationType = "Certificate"
                                PrincipalType = $principalType
                                RoleResolutionMethod = $roleResolutionMethod  # New field for debugging
                                AssignmentDisplayName = $assignment.displayName  # New field for reference
                            }
                            
                            Write-Verbose "Successfully processed member $userDisplayName for role $($roleDefinition.displayName) via $roleResolutionMethod"
                        }
                        catch {
                            Write-Warning "Error processing assignment member $member`: $($_.Exception.Message)"
                        }
                    }
                }
                catch {
                    Write-Warning "Error processing Intune role assignment $($assignment.id)`: $($_.Exception.Message)"
                    continue
                }
            }
        }
        catch {
            Write-Warning "Could not retrieve Intune administrative role assignments: $($_.Exception.Message)"
        }

        Write-Host "✓ Intune administrative role audit completed. Found $($results.Count) administrative role assignments (including PIM)" -ForegroundColor Green
              
        # Provide feedback about role filtering
        if (-not $IncludeAzureADRoles) {
            Write-Host "  (Excluding overarching Azure AD roles - use -IncludeAzureADRoles to include)" -ForegroundColor Yellow
        }
        
        # Enhanced administrative role analysis for Intune
        if ($IncludeAnalysis) {
            Write-Host ""
            Write-Host "=== Intune Administrative Role Analysis ===" -ForegroundColor Cyan

            $intuneResults = $results | Where-Object { $_.Service -eq "Microsoft Intune" }
            $intuneEligibleAssignments = $intuneResults | Where-Object { $_.AssignmentType -like "*Eligible*" }
            $intuneActiveAssignments = $intuneResults | Where-Object { $_.AssignmentType -eq "Azure AD Role" -or $_.AssignmentType -eq "Intune RBAC" }
            $intuneServiceAdmins = $intuneResults | Where-Object { $_.RoleName -eq "Intune Service Administrator" }
            $intuneTimeBoundAssignments = $intuneResults | Where-Object { $_.AssignmentType -eq "Time-bound RBAC" }
            $intuneRBACAssignments = $intuneResults | Where-Object { $_.RoleType -eq "IntuneRBAC" }
            $intuneAzureADAssignments = $intuneResults | Where-Object { $_.RoleType -eq "AzureAD" }

            Write-Host "Intune Administrative Role Statistics:" -ForegroundColor White
            Write-Host "• Total Administrative Assignments: $($intuneResults.Count)" -ForegroundColor Gray
            Write-Host "• PIM Eligible Assignments: $($intuneEligibleAssignments.Count)" -ForegroundColor Gray
            Write-Host "• Time-bound RBAC Assignments: $($intuneTimeBoundAssignments.Count)" -ForegroundColor Gray
            Write-Host "• Permanent Active Assignments: $($intuneActiveAssignments.Count)" -ForegroundColor Gray
            Write-Host "• Intune Service Administrators: $($intuneServiceAdmins.Count)" -ForegroundColor Gray
            Write-Host "• Intune RBAC Assignments: $($intuneRBACAssignments.Count)" -ForegroundColor Gray
            Write-Host "• Azure AD Role Assignments: $($intuneAzureADAssignments.Count)" -ForegroundColor Gray
            
            # Administrative role recommendations
            if ($intuneActiveAssignments.Count -gt $intuneEligibleAssignments.Count -and $intuneEligibleAssignments.Count -eq 0) {
                Write-Host "⚠️ RECOMMENDATION: Consider implementing PIM for Intune administrative role assignments" -ForegroundColor Yellow
                Write-Host "   Benefits: Just-in-time access, audit trail, approval workflows" -ForegroundColor White
            }

            if ($intuneServiceAdmins.Count -gt 2) {
                Write-Host "⚠️ RECOMMENDATION: Review Intune Service Administrator count ($($intuneServiceAdmins.Count))" -ForegroundColor Yellow
                Write-Host "   Consider using Intune RBAC roles for more granular administrative permissions" -ForegroundColor White
            }

            if ($intuneAzureADAssignments.Count -gt $intuneRBACAssignments.Count) {
                Write-Host "⚠️ RECOMMENDATION: Leverage Intune RBAC over Azure AD roles where possible" -ForegroundColor Yellow
                Write-Host "   Intune RBAC provides more granular, scope-specific administrative permissions" -ForegroundColor White
            }

            if ($intuneEligibleAssignments.Count -gt 0) {
                Write-Host "✓ PIM is being used for some Intune administrative role assignments" -ForegroundColor Green
            }

            if ($intuneTimeBoundAssignments.Count -gt 0) {
                Write-Host "✓ Time-bound RBAC assignments detected - good security practice" -ForegroundColor Green
            }

            # Security validation
            Write-Host ""
            Write-Host "=== Security Validation ===" -ForegroundColor Green
            Write-Host "✓ Certificate-based authentication enforced" -ForegroundColor Green
            Write-Host "✓ No interactive authentication required" -ForegroundColor Green
            Write-Host "✓ Suitable for automated/unattended execution" -ForegroundColor Green
            Write-Host "✓ Enhanced security posture maintained" -ForegroundColor Green

            # Administrative role security recommendations
            Write-Host ""
            Write-Host "Administrative Role Security Recommendations:" -ForegroundColor Cyan
            Write-Host "• Implement regular access reviews for Intune administrative roles" -ForegroundColor White
            Write-Host "• Use Intune RBAC for granular, scoped permissions instead of broad Azure AD roles" -ForegroundColor White
            Write-Host "• Enable PIM for Intune Service Administrator and other high-privilege roles" -ForegroundColor White
            Write-Host "• Consider time-bound assignments for temporary administrative access needs" -ForegroundColor White
            Write-Host "• Monitor administrative role changes through audit logs" -ForegroundColor White
            Write-Host "• Implement approval workflows for sensitive Intune administrative operations" -ForegroundColor White
            Write-Host "• Regularly rotate certificates (recommended 12-24 months)" -ForegroundColor White

            Write-Host ""
            Write-Host "=== SCOPE CLARIFICATION ===" -ForegroundColor Green
            Write-Host "✓ Focused on administrative role assignments only" -ForegroundColor Green
            Write-Host "✓ Excluded: Device Configuration policy ownership tracking" -ForegroundColor Green
            Write-Host "✓ Excluded: Device Compliance policy ownership tracking" -ForegroundColor Green
            Write-Host "✓ Excluded: App Protection policy ownership tracking" -ForegroundColor Green
            Write-Host "✓ Excluded: Device Enrollment configuration ownership tracking" -ForegroundColor Green
            Write-Host "✓ Included: Azure AD Intune administrative roles and Intune RBAC administrative assignments only" -ForegroundColor Green

            Write-Host ""
            Write-Host "✓ Intune administrative role audit completed with focused analysis" -ForegroundColor Green
            Write-Host "Found $($results.Count) Intune administrative role assignments" -ForegroundColor Cyan
            Write-Host "Authentication Method: Certificate-based (Secure)" -ForegroundColor Green
        }
        
        # Show breakdown
        if ($results.Count -gt 0) {
            $typeSummary = $results | Group-Object PrincipalType
            $assignmentTypeSummary = $results | Group-Object AssignmentType
            $scopeSummary = $results | Group-Object RoleScope
            $roleTypeSummary = $results | Group-Object RoleType
            
            Write-Host ""
            Write-Host "Administrative assignment breakdown:" -ForegroundColor Cyan
            
            Write-Host "Principal types:" -ForegroundColor Yellow
            foreach ($type in $typeSummary) {
                Write-Host "  $($type.Name): $($type.Count)" -ForegroundColor White
            }
            
            Write-Host "Assignment types:" -ForegroundColor Yellow
            foreach ($type in $assignmentTypeSummary) {
                Write-Host "  $($type.Name): $($type.Count)" -ForegroundColor White
            }
            
            Write-Host "Role scope:" -ForegroundColor Yellow
            foreach ($scope in $scopeSummary) {
                Write-Host "  $($scope.Name): $($scope.Count)" -ForegroundColor White
            }
            
            Write-Host "Role types:" -ForegroundColor Yellow
            foreach ($roleType in $roleTypeSummary) {
                Write-Host "  $($roleType.Name): $($roleType.Count)" -ForegroundColor White
            }
        }
        
    } catch {
        Write-Error "Error during Intune administrative role audit: $($_.Exception.Message)"
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
            Write-Host "• For PIM: Ensure RoleEligibilitySchedule and RoleAssignmentSchedule permissions" -ForegroundColor White
            Write-Host "• For Intune: Ensure DeviceManagementRBAC.Read.All permission is granted" -ForegroundColor White
        }
        
        throw
    }
    
    return $results
}