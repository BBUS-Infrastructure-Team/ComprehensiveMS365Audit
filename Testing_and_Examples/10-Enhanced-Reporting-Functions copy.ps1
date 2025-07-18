# 10-Enhanced-Reporting-Functions.ps1
# Enhanced HTML reporting function updated for comprehensive M365 audit module
#


function Export-M365AuditHtmlReport {
    param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipeline
        )]
        [array]$AuditResults,
        
        [string]$OutputPath = ".\M365_Audit_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').html",
        [string]$OrganizationName = "Organization",
        [switch]$IncludeCharts,
        [bool]$IncludePIMAnalysis = $true,
        [bool]$IncludeComplianceGaps = $true
    )
    
    Write-Host "Generating enhanced HTML audit report..." -ForegroundColor Cyan
    
    if ($AuditResults.Count -eq 0) {
        Write-Warning "No audit results provided"
        return
    }
    
    # === USE HELPER FUNCTIONS FOR CALCULATIONS ===
    
    # Calculate comprehensive statistics using helper function
    $stats = Get-AuditStatistics -AuditResults $AuditResults
    
    # Get report metadata
    $metadata = Get-ReportMetadata -OrganizationName $OrganizationName -Stats $stats -AuditResults $AuditResults
    
    # Get report summary
    $summary = Get-ReportSummary -AuditResults $AuditResults -Stats $stats
    
    # Get service analysis
    $serviceAnalysis = Get-ServiceAnalysis -AuditResults $AuditResults -IncludeExchangeAnalysis
    
    # Get PIM analysis
    $pimAnalysis = Get-PIMAnalysis -AuditResults $AuditResults -IncludeDetailedAnalysis:$IncludePIMAnalysis
    
    # Get principal analysis
    $principalAnalysis = Get-PrincipalAnalysis -AuditResults $AuditResults
    
    # Get cross-service analysis
    $crossServiceAnalysis = Get-CrossServiceAnalysis -AuditResults $AuditResults
    
    # Get security alerts
    $securityAlerts = Get-SecurityAlerts -AuditResults $AuditResults -Stats $stats
    
    # Get recommendations
    $recommendations = Get-SecurityRecommendations -AuditResults $AuditResults -Stats $stats
    
    # Get compliance analysis if requested
    $complianceAnalysis = if ($IncludeComplianceGaps) {
        Get-ComplianceAnalysis -AuditResults $AuditResults -Stats $stats
    } else { $null }
    
    # === BUILD HTML USING HELPER FUNCTION DATA ===
    
    # Build enhanced HTML content
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>M365 Comprehensive Role Audit Report - $OrganizationName</title>
    <style>
        :root {
            --primary-color: #0078d4;
            --secondary-color: #106ebe;
            --success-color: #107c10;
            --warning-color: #ff8c00;
            --danger-color: #d13438;
            --info-color: #00bcf2;
            --dark-color: #323130;
            --light-color: #f3f2f1;
        }
        
        * { box-sizing: border-box; }
        
        body { 
            font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif; 
            margin: 0; 
            padding: 20px; 
            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
            line-height: 1.6;
        }
        
        .container { 
            max-width: 1400px; 
            margin: 0 auto; 
            background: white; 
            padding: 40px; 
            border-radius: 12px; 
            box-shadow: 0 8px 32px rgba(0,0,0,0.1);
            backdrop-filter: blur(10px);
        }
        
        .header { 
            text-align: center; 
            margin-bottom: 40px; 
            padding-bottom: 30px; 
            border-bottom: 3px solid var(--primary-color);
            background: linear-gradient(135deg, var(--primary-color), var(--secondary-color));
            color: white;
            margin: -40px -40px 40px -40px;
            padding: 40px;
            border-radius: 12px 12px 0 0;
        }
        
        .header h1 { 
            margin: 0; 
            font-size: 2.8em; 
            font-weight: 300;
            text-shadow: 0 2px 4px rgba(0,0,0,0.3);
        }
        
        .header p { 
            margin: 15px 0 0 0; 
            font-size: 1.2em; 
            opacity: 0.9;
        }
        
        .metadata {
            background: var(--info-color);
            color: white;
            padding: 15px;
            border-radius: 8px;
            margin: 20px 0;
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
        }
        
        .metadata-item {
            text-align: center;
        }
        
        .metadata-item strong {
            display: block;
            font-size: 0.9em;
            opacity: 0.8;
        }
        
        .metadata-item span {
            display: block;
            font-size: 1.1em;
            font-weight: bold;
        }
        
        .summary-grid { 
            display: grid; 
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); 
            gap: 25px; 
            margin-bottom: 40px; 
        }
        
        .summary-card { 
            background: linear-gradient(135deg, var(--primary-color), var(--secondary-color)); 
            color: white; 
            padding: 30px; 
            border-radius: 12px; 
            text-align: center;
            position: relative;
            overflow: hidden;
            transition: transform 0.3s ease;
        }
        
        .summary-card:hover {
            transform: translateY(-5px);
        }
        
        .summary-card::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: linear-gradient(45deg, transparent 30%, rgba(255,255,255,0.1) 50%, transparent 70%);
            transform: translateX(-100%);
            transition: transform 0.6s;
        }
        
        .summary-card:hover::before {
            transform: translateX(100%);
        }
        
        .summary-card h3 { 
            margin: 0 0 15px 0; 
            font-size: 1.3em; 
            font-weight: 400;
        }
        
        .summary-card .number { 
            font-size: 3em; 
            font-weight: bold; 
            margin: 15px 0;
            text-shadow: 0 2px 4px rgba(0,0,0,0.3);
        }
        
        .summary-card .subtitle {
            font-size: 0.9em;
            opacity: 0.8;
        }
        
        .section { 
            margin-bottom: 50px; 
            position: relative;
        }
        
        .section h2 { 
            color: var(--primary-color); 
            border-bottom: 3px solid var(--primary-color); 
            padding-bottom: 15px; 
            font-size: 2em;
            font-weight: 300;
            margin-bottom: 25px;
            position: relative;
        }
        
        .section h2::after {
            content: '';
            position: absolute;
            bottom: -3px;
            left: 0;
            width: 60px;
            height: 3px;
            background: var(--warning-color);
        }
        
        .alert-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 20px;
            margin: 25px 0;
        }
        
        .security-alerts { 
            background: linear-gradient(135deg, #fff3cd, #ffeaa7); 
            border: 1px solid var(--warning-color); 
            border-left: 5px solid var(--warning-color);
            border-radius: 8px; 
            padding: 25px; 
            margin: 25px 0; 
        }
        
        .security-alerts h3 {
            margin-top: 0;
            color: #856404;
            font-size: 1.2em;
        }
        
        .alert-item { 
            margin: 15px 0; 
            padding: 10px;
            border-radius: 5px;
            background: rgba(255,255,255,0.7);
        }
        
        .alert-warning { 
            color: #856404; 
            border-left: 4px solid var(--warning-color);
            padding-left: 15px;
        }
        
        .alert-critical { 
            color: #721c24; 
            font-weight: bold; 
            border-left: 4px solid var(--danger-color);
            padding-left: 15px;
            background: rgba(209, 52, 56, 0.1);
        }
        
        .alert-success {
            color: var(--success-color);
            border-left: 4px solid var(--success-color);
            padding-left: 15px;
            background: rgba(16, 124, 16, 0.1);
        }
        
        table { 
            width: 100%; 
            border-collapse: collapse; 
            margin: 25px 0; 
            background: white;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 4px 16px rgba(0,0,0,0.1);
        }
        
        th, td { 
            padding: 15px 12px; 
            text-align: left; 
            border-bottom: 1px solid #e1dfdd; 
        }
        
        th { 
            background: linear-gradient(135deg, var(--primary-color), var(--secondary-color)); 
            color: white; 
            font-weight: 600;
            text-transform: uppercase;
            font-size: 0.9em;
            letter-spacing: 0.5px;
        }
        
        tr:hover { 
            background-color: #f9f9f9; 
            transition: background-color 0.2s ease;
        }
        
        tr:nth-child(even) {
            background-color: #fafafa;
        }
        
        .service-badge {
            padding: 4px 12px;
            border-radius: 20px;
            color: white;
            font-size: 0.85em;
            font-weight: 600;
            text-align: center;
            display: inline-block;
            min-width: 120px;
        }
        
        .service-azure { background: linear-gradient(135deg, #0078d4, #106ebe); }
        .service-sharepoint { background: linear-gradient(135deg, #0b6623, #0e7629); }
        .service-exchange { background: linear-gradient(135deg, #d13438, #b02a37); }
        .service-teams { background: linear-gradient(135deg, #464775, #5b5d8a); }
        .service-purview { background: linear-gradient(135deg, #8b4789, #9e5a9c); }
        .service-intune { background: linear-gradient(135deg, #00bcf2, #0078d4); }
        .service-defender { background: linear-gradient(135deg, #ff8c00, #e67e22); }
        .service-powerplatform { background: linear-gradient(135deg, #742774, #8b4789); }
        
        .privilege-high { 
            background: linear-gradient(135deg, #ffebee, #ffcdd2); 
            border-left: 5px solid var(--danger-color); 
        }
        
        .privilege-medium { 
            background: linear-gradient(135deg, #fff3e0, #ffe0b2); 
            border-left: 5px solid var(--warning-color); 
        }
        
        .privilege-low { 
            background: linear-gradient(135deg, #e8f5e8, #c8e6c9); 
            border-left: 5px solid var(--success-color); 
        }
        
        .pim-analysis {
            background: linear-gradient(135deg, #e3f2fd, #bbdefb);
            border: 1px solid var(--info-color);
            border-radius: 8px;
            padding: 25px;
            margin: 25px 0;
        }
        
        .pim-analysis h3 {
            color: var(--info-color);
            margin-top: 0;
        }
        
        .progress-bar {
            background: #e0e0e0;
            border-radius: 10px;
            height: 20px;
            margin: 10px 0;
            overflow: hidden;
        }
        
        .progress-fill {
            height: 100%;
            border-radius: 10px;
            transition: width 0.5s ease;
        }
        
        .progress-pim { background: linear-gradient(135deg, var(--success-color), #16a085); }
        .progress-active { background: linear-gradient(135deg, var(--warning-color), #e67e22); }
        .progress-permanent { background: linear-gradient(135deg, var(--danger-color), #c0392b); }
        
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin: 20px 0;
        }
        
        .stat-item {
            background: white;
            padding: 20px;
            border-radius: 8px;
            text-align: center;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            border-top: 4px solid var(--primary-color);
        }
        
        .stat-number {
            font-size: 2em;
            font-weight: bold;
            color: var(--primary-color);
        }
        
        .stat-label {
            font-size: 0.9em;
            color: #666;
            margin-top: 5px;
        }
        
        .footer { 
            margin-top: 60px; 
            padding-top: 30px; 
            border-top: 2px solid #ddd; 
            text-align: center; 
            color: #666; 
            background: #f9f9f9;
            margin-left: -40px;
            margin-right: -40px;
            padding-left: 40px;
            padding-right: 40px;
            border-radius: 0 0 12px 12px;
        }
        
        .expandable {
            cursor: pointer;
            user-select: none;
        }
        
        .expandable:hover {
            background: #f0f0f0;
        }
        
        .expandable-content {
            display: none;
            padding: 15px;
            background: #f9f9f9;
            border-radius: 5px;
            margin-top: 10px;
        }
        
        .tag {
            display: inline-block;
            background: var(--primary-color);
            color: white;
            padding: 2px 8px;
            border-radius: 12px;
            font-size: 0.8em;
            margin: 2px;
        }
        
        .tag.pim { background: var(--success-color); }
        .tag.permanent { background: var(--warning-color); }
        .tag.disabled { background: var(--danger-color); }
        
        .scroll-to-top {
            position: fixed;
            bottom: 20px;
            right: 20px;
            background: var(--primary-color);
            color: white;
            border: none;
            border-radius: 50%;
            width: 50px;
            height: 50px;
            cursor: pointer;
            box-shadow: 0 4px 16px rgba(0,0,0,0.2);
            transition: all 0.3s ease;
        }
        
        .scroll-to-top:hover {
            background: var(--secondary-color);
            transform: translateY(-2px);
        }

        /* Role Analysis Expandable Rows */
        .expandable-role-row {
            cursor: pointer;
            transition: background-color 0.2s ease;
        }
        
        .expandable-role-row:hover {
            background-color: #f0f8ff !important;
        }
        
        .expand-indicator {
            float: right;
            font-size: 0.8em;
            color: var(--primary-color);
            transition: transform 0.3s ease;
        }
        
        .expand-indicator.expanded {
            transform: rotate(90deg);
        }
        
        .role-assignments-row {
            background: #f9f9f9 !important;
        }
        
        .assignments-container {
            padding: 20px;
            background: #ffffff;
            border-radius: 8px;
            margin: 10px 0;
            box-shadow: inset 0 2px 4px rgba(0,0,0,0.1);
        }
        
        .assignments-title {
            font-size: 1.1em;
            font-weight: bold;
            color: var(--primary-color);
            margin-bottom: 15px;
            padding-bottom: 8px;
            border-bottom: 2px solid var(--primary-color);
        }
        
        .assignments-horizontal {
            display: flex;
            flex-wrap: wrap;
            gap: 12px;
            padding: 8px 0;
        }
        
        .assignment-chip {
            background: linear-gradient(135deg, #e3f2fd, #bbdefb);
            border: 1px solid var(--info-color);
            border-radius: 20px;
            padding: 8px 16px;
            font-size: 0.9em;
            font-weight: 500;
            color: #1565c0;
            white-space: nowrap;
            transition: all 0.2s ease;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        
        .assignment-chip:hover {
            background: linear-gradient(135deg, #bbdefb, #90caf9);
            transform: translateY(-1px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.15);
        }
        
        .no-assignments {
            text-align: center;
            padding: 20px;
            color: #666;
            font-style: italic;
            background: #f9f9f9;
            border-radius: 8px;
        }

        /* User Analysis Expandable Rows */
        .expandable-user-row {
            cursor: pointer;
            transition: background-color 0.2s ease;
        }
        
        .expandable-user-row:hover {
            background-color: #f0f8ff !important;
        }
        
        .user-roles-row {
            background: #f9f9f9 !important;
        }
        
        .user-roles-cell {
            padding: 20px;
            background: #ffffff;
            border-radius: 8px;
            margin: 10px 0;
            box-shadow: inset 0 2px 4px rgba(0,0,0,0.1);
        }
        
        .user-roles-title {
            font-size: 1.1em;
            font-weight: bold;
            color: var(--primary-color);
            margin-bottom: 15px;
            padding-bottom: 8px;
            border-bottom: 2px solid var(--primary-color);
        }
        
        .user-roles-container {
            display: flex;
            flex-direction: column;
            gap: 15px;
            /* Removed max-height and overflow to allow full expansion */
        }
        
        .service-group {
            border: 1px solid #e1dfdd;
            border-radius: 8px;
            overflow: hidden;
            background: #ffffff;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            margin-bottom: 10px;
        }
        
        .service-group-header {
            padding: 12px 16px;
            font-weight: bold;
            color: white;
            font-size: 0.9em;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        /* Service Group Header Colors */
        .service-group-header.service-azure { background: linear-gradient(135deg, #0078d4, #106ebe); }
        .service-group-header.service-sharepoint { background: linear-gradient(135deg, #0b6623, #0e7629); }
        .service-group-header.service-exchange { background: linear-gradient(135deg, #d13438, #b02a37); }
        .service-group-header.service-teams { background: linear-gradient(135deg, #464775, #5b5d8a); }
        .service-group-header.service-purview { background: linear-gradient(135deg, #8b4789, #9e5a9c); }
        .service-group-header.service-intune { background: linear-gradient(135deg, #00bcf2, #0078d4); }
        .service-group-header.service-defender { background: linear-gradient(135deg, #ff8c00, #e67e22); }
        .service-group-header.service-powerplatform { background: linear-gradient(135deg, #742774, #8b4789); }
        
        .service-roles {
            padding: 15px;
            background: #fafafa;
            min-height: 50px;
        }
        
        .role-item {
            padding: 8px 12px;
            margin: 5px 0;
            background: #ffffff;
            border: 1px solid #e1dfdd;
            border-radius: 4px;
            font-size: 0.9em;
            color: #333;
            transition: all 0.2s ease;
            display: block;
            width: 100%;
        }
        
        .role-item:hover {
            background: #f0f8ff;
            border-color: var(--primary-color);
            transform: translateX(5px);
        }
        
        .no-roles {
            text-align: center;
            padding: 20px;
            color: #666;
            font-style: italic;
            background: #f9f9f9;
            border-radius: 8px;
        }
        
        /* Mobile Responsive */
        @media (max-width: 768px) {
            .container { padding: 20px; }
            .summary-grid { grid-template-columns: 1fr; }
            .header h1 { font-size: 2em; }
            table { font-size: 0.9em; }
            th, td { padding: 10px 8px; }
            
            .assignment-chip {
                font-size: 0.8em;
                padding: 6px 12px;
            }
            
            .assignments-horizontal {
                gap: 8px;
            }
            
            .service-group-header {
                font-size: 0.8em;
                padding: 10px 12px;
            }
            
            .role-item {
                font-size: 0.8em;
                padding: 6px 10px;
            }
        }

        .assignments-section {
            display: flex;
            flex-direction: column;
            gap: 12px;
        }
        
        .assignments-legend {
            display: flex;
            gap: 20px;
            padding: 8px 12px;
            background: #f8f9fa;
            border-radius: 6px;
            border: 1px solid #e9ecef;
            font-size: 0.85em;
            justify-content: center;
        }
        
        .legend-item {
            display: flex;
            align-items: center;
            gap: 6px;
            font-weight: 500;
        }
        
        .legend-color {
            width: 12px;
            height: 12px;
            border-radius: 50%;
            border: 1px solid rgba(0,0,0,0.2);
        }
        
        .legend-color.active {
            background: linear-gradient(135deg, #e3f2fd, #bbdefb);
        }
        
        .legend-color.eligible {
            background: linear-gradient(135deg, #fff3e0, #ffcc80);
        }
        
        .assignment-chip.active {
            background: linear-gradient(135deg, #e3f2fd, #bbdefb);
            border: 1px solid var(--info-color);
            color: #1565c0;
        }
        
        .assignment-chip.eligible {
            background: linear-gradient(135deg, #fff3e0, #ffcc80);
            border: 1px solid var(--warning-color);
            color: #ef6c00;
        }
        
        .assignment-chip.active:hover {
            background: linear-gradient(135deg, #bbdefb, #90caf9);
        }
        
        .assignment-chip.eligible:hover {
            background: linear-gradient(135deg, #ffcc80, #ffb74d);
        }
    </style>
    <script>
        function toggleSection(id) {
            const content = document.getElementById(id);
            if (content.style.display === 'none' || content.style.display === '') {
                content.style.display = 'block';
            } else {
                content.style.display = 'none';
            }
        }
        
        function scrollToTop() {
            window.scrollTo({ top: 0, behavior: 'smooth' });
        }

        function toggleRoleAssignments(roleId) {
            const assignmentsRow = document.getElementById('roleAssignments_' + roleId);
            const expandIndicator = event.currentTarget.querySelector('.expand-indicator');
            
            if (assignmentsRow.style.display === 'none' || assignmentsRow.style.display === '') {
                assignmentsRow.style.display = 'table-row';
                expandIndicator.textContent = '▼';
                expandIndicator.classList.add('expanded');
            } else {
                assignmentsRow.style.display = 'none';
                expandIndicator.textContent = '▶';
                expandIndicator.classList.remove('expanded');
            }
        }

        function toggleUserRoles(userId) {
            const rolesRow = document.getElementById('userRoles_' + userId);
            const expandIndicator = event.currentTarget.querySelector('.expand-indicator');
            
            if (rolesRow.style.display === 'none' || rolesRow.style.display === '') {
                rolesRow.style.display = 'table-row';
                expandIndicator.textContent = '▼';
                expandIndicator.classList.add('expanded');
            } else {
                rolesRow.style.display = 'none';
                expandIndicator.textContent = '▶';
                expandIndicator.classList.remove('expanded');
            }
        }
        
        document.addEventListener('DOMContentLoaded', function() {
            window.addEventListener('scroll', function() {
                const scrollBtn = document.querySelector('.scroll-to-top');
                if (window.pageYOffset > 300) {
                    scrollBtn.style.display = 'block';
                } else {
                    scrollBtn.style.display = 'none';
                }
            });
        });
    </script>
</head>
<body>
    <button class="scroll-to-top" onclick="scrollToTop()" style="display: none;">↑</button>
    
    <div class="container">
        <div class="header">
            <h1>Microsoft 365 Comprehensive Role Audit</h1>
            <p>$OrganizationName | Generated on $(Get-Date -Format 'MMMM dd, yyyy at HH:mm')</p>
        </div>
        
        <div class="metadata">
            <div class="metadata-item">
                <strong>Audit Scope</strong>
                <span>$($metadata.servicesAudited) Services</span>
            </div>
            <div class="metadata-item">
                <strong>Authentication</strong>
                <span>$(if($metadata.certificateAuthUsed) { "Certificate-Based" } else { "Mixed" })</span>
            </div>
            <div class="metadata-item">
                <strong>PIM Analysis</strong>
                <span>$(if($metadata.pimEnabled) { "Included" } else { "N/A" })</span>
            </div>
            <div class="metadata-item">
                <strong>Hybrid Environment</strong>
                <span>$(if($metadata.hybridEnvironmentDetected) { "Detected" } else { "Cloud-Only" })</span>
            </div>
        </div>

        <div class="summary-grid">
            <div class="summary-card">
                <h3>Total Role Assignments</h3>
                <div class="number">$($metadata.totalAssignments)</div>
                <div class="subtitle">Across all services</div>
            </div>
            <div class="summary-card">
                <h3>Unique Users</h3>
                <div class="number">$($metadata.uniqueUsers)</div>
                <div class="subtitle">With role assignments</div>
            </div>
            <div class="summary-card">
                <h3>Services Audited</h3>
                <div class="number">$($metadata.servicesAudited)</div>
                <div class="subtitle">Microsoft 365 services</div>
            </div>
            <div class="summary-card">
                <h3>Global Administrators</h3>
                <div class="number">$($stats.globalAdmins.Count)</div>
                <div class="subtitle">Highest privilege level</div>
            </div>
"@

    # Add PIM summary cards if PIM data exists
    if ($pimAnalysis.enabled) {
        $html += @"
            <div class="summary-card">
                <h3>PIM Eligible</h3>
                <div class="number">$($pimAnalysis.totalEligible)</div>
                <div class="subtitle">Require activation</div>
            </div>
            <div class="summary-card">
                <h3>PIM Active</h3>
                <div class="number">$($pimAnalysis.totalActive)</div>
                <div class="subtitle">Currently activated</div>
            </div>
"@
    }

    $html += @"
        </div>

        <div class="section">
            <h2>🔐 Security Analysis</h2>
"@

    # Enhanced security alerts using helper function data
    $html += '<div class="alert-grid">'

    # Critical alerts
    if ($securityAlerts.critical.Count -gt 0) {
        $html += @"
            <div class="security-alerts" style="background: linear-gradient(135deg, #ffebee, #ffcdd2); border-color: var(--danger-color);">
                <h3 style="color: var(--danger-color);">Critical Issues</h3>
"@
        foreach ($alert in $securityAlerts.critical) {
            $html += "<div class='alert-item alert-critical'>⚠️ $alert</div>"
        }
        $html += "</div>"
    }

    # High priority alerts
    if ($securityAlerts.high.Count -gt 0) {
        $html += @"
            <div class="security-alerts">
                <h3>High Priority Issues</h3>
"@
        foreach ($alert in $securityAlerts.high) {
            $html += "<div class='alert-item alert-warning'>⚠️ $alert</div>"
        }
        $html += "</div>"
    }

    # Medium priority alerts
    if ($securityAlerts.medium.Count -gt 0) {
        $html += @"
            <div class="security-alerts">
                <h3>Medium Priority Issues</h3>
"@
        foreach ($alert in $securityAlerts.medium) {
            $html += "<div class='alert-item alert-warning'>⚠️ $alert</div>"
        }
        $html += "</div>"
    }

    # Low priority alerts
    if ($securityAlerts.low.Count -gt 0) {
        $html += @"
            <div class="security-alerts" style="background: linear-gradient(135deg, #e8f5e8, #c8e6c9); border-color: var(--success-color);">
                <h3 style="color: var(--success-color);">Low Priority Items</h3>
"@
        foreach ($alert in $securityAlerts.low) {
            $html += "<div class='alert-item alert-success'>ℹ️ $alert</div>"
        }
        $html += "</div>"
    }

    $html += "</div>" # Close alert-grid

    # Add PIM Analysis section if applicable
    if ($IncludePIMAnalysis -and $pimAnalysis.enabled) {
        $html += @"
        </div>

        <div class="section">
            <h2>🔑 Privileged Identity Management Analysis</h2>
            <div class="pim-analysis">
                <h3>Assignment Type Distribution</h3>
                <div class="stats-grid">
                    <div class="stat-item">
                        <div class="stat-number" style="color: var(--success-color);">$($pimAnalysis.totalEligible)</div>
                        <div class="stat-label">PIM Eligible</div>
                    </div>
                    <div class="stat-item">
                        <div class="stat-number" style="color: var(--info-color);">$($pimAnalysis.totalActive)</div>
                        <div class="stat-label">PIM Active</div>
                    </div>
                </div>
            </div>
"@
    }

    $html += @"
        </div>

        <div class="section">
            <h2>📊 Service Distribution</h2>
            <table>
                <tr><th>Service</th><th>Assignments</th><th>Percentage</th></tr>
"@

    # Service breakdown using helper function data
    foreach ($service in $summary.serviceBreakdown) {
        $serviceClass = switch ($service.service) {
            "Azure AD/Entra ID" { "service-azure" }
            "SharePoint Online" { "service-sharepoint" }
            "Exchange Online" { "service-exchange" }
            "Microsoft Teams" { "service-teams" }
            "Microsoft Purview" { "service-purview" }
            "Microsoft Intune" { "service-intune" }
            "Microsoft Defender" { "service-defender" }
            "Power Platform" { "service-powerplatform" }
            default { "service-azure" }
        }
        
        $html += @"
<tr>
    <td><span class='service-badge $serviceClass'>$($service.service)</span></td>
    <td><strong>$($service.count)</strong></td>
    <td>$($service.percentage)%</td>
</tr>
"@
    }

    $html += @"
            </table>
        </div>

        <div class="section">
            <h2>👑 High-Privilege Role Analysis</h2>
            <table>
                <tr><th>Role Name</th><th>Users Assigned</th><th>Risk Level</th><th>Services</th></tr>
"@

        # Top roles using helper function data
    foreach ($role in $summary.topRoles) {
        $riskClass = switch ($role.riskLevel) {
            "CRITICAL" { "privilege-high" }
            "HIGH" { "privilege-medium" }
            "MEDIUM" { "privilege-medium" }
            default { "privilege-low" }
        }
        
        $servicesDisplay = ($role.services | Select-Object -First 3) -join ", "
        if ($role.services.Count -gt 3) { $servicesDisplay += "..." }
        
        # Get role users using the helper function
        $roleUsers = Get-RoleUsers -AuditResults $AuditResults -RoleName $role.roleName
        $roleId = ($role.roleName -replace '[^a-zA-Z0-9]', '')
        
# Build simple assignments list with PIM differentiation
        $assignmentsGrid = ""
        if ($roleUsers.Count -gt 0) {
            # Count different assignment types for legend
            $activeCount = ($roleUsers | Where-Object { $_.assignmentType -notlike "*Eligible*" }).Count
            $eligibleCount = ($roleUsers | Where-Object { $_.assignmentType -like "*Eligible*" }).Count
            
            $assignmentsGrid = "<div class='assignments-section'>"
            $assignmentsGrid += "<div class='assignments-horizontal'>"
            
            foreach ($user in $roleUsers) {
                $chipClass = if ($user.assignmentType -like "*Eligible*") { "assignment-chip eligible" } else { "assignment-chip active" }
                $assignmentsGrid += "<div class='$chipClass'>$($user.displayName)</div>"
            }
            $assignmentsGrid += "</div>"
            
            # Add legend at the bottom if there are both types
            if ($activeCount -gt 0 -and $eligibleCount -gt 0) {
                $assignmentsGrid += "<div class='assignments-legend'>"
                $assignmentsGrid += "<span class='legend-item'><span class='legend-color active'></span> Active Assignments ($activeCount)</span>"
                $assignmentsGrid += "<span class='legend-item'><span class='legend-color eligible'></span> PIM Eligible ($eligibleCount)</span>"
                $assignmentsGrid += "</div>"
            } elseif ($eligibleCount -gt 0) {
                $assignmentsGrid += "<div class='assignments-legend'>"
                $assignmentsGrid += "<span class='legend-item'><span class='legend-color eligible'></span> PIM Eligible ($eligibleCount)</span>"
                $assignmentsGrid += "</div>"
            } elseif ($activeCount -gt 0) {
                $assignmentsGrid += "<div class='assignments-legend'>"
                $assignmentsGrid += "<span class='legend-item'><span class='legend-color active'></span> Active Assignments ($activeCount)</span>"
                $assignmentsGrid += "</div>"
            }
            
            $assignmentsGrid += "</div>"
        } else {
            $assignmentsGrid = "<div class='no-assignments'>No assignments found</div>"
        }
        
        $html += @"
<tr class='$riskClass expandable-role-row' onclick="toggleRoleAssignments('$roleId')">
    <td><strong>$($role.roleName)</strong> <span class='expand-indicator'>▶</span></td>
    <td>$($role.assignmentCount)</td>
    <td><span class='tag $(if($role.riskLevel -eq "CRITICAL") {"disabled"} elseif($role.riskLevel -eq "HIGH") {"permanent"} else {"pim"})'>$($role.riskLevel)</span></td>
    <td><small>$servicesDisplay</small></td>
</tr>
<tr id="roleAssignments_$roleId" class="role-assignments-row" style="display: none;">
    <td colspan="4" class="assignments-container">
        <div class="assignments-title">Role Assignments for $($role.roleName)</div>
        $assignmentsGrid
    </td>
</tr>
"@
    }

    $html += @"
            </table>
        </div>

        <div class="section">
            <h2>👥 User Analysis</h2>
            <table>
                <tr><th>User</th><th>Role Count</th><th>Status</th><th>Services</th></tr>
"@

    # User analysis using helper function data
    foreach ($user in $summary.usersWithMostRoles) {
        $statusColor = if ($user.isEnabled -eq $false) { "style='color: var(--danger-color); font-weight: bold;'" } else { "" }
        $status = if ($user.isEnabled -eq $false) { "DISABLED" } else { "Active" }
        
        $servicesDisplay = ($user.services | Select-Object -First 3) -join ", "
        if ($user.services.Count -gt 3) { $servicesDisplay += " +$($user.services.Count - 3) more" }
        
        # Get user roles grouped by service using helper function
        $userRoles = Get-UserRoles -AuditResults $AuditResults -UserPrincipalName $user.userPrincipalName
        $userId = ($user.userPrincipalName -replace '[^a-zA-Z0-9]', '')
        
        # Build service-grouped roles
        $serviceGroupedRoles = ""
        if ($userRoles.Count -gt 0) {
            $rolesByService = $userRoles | Group-Object Service | Sort-Object Name
            $serviceGroupedRoles = "<div class='user-roles-container'>"
            
            foreach ($serviceGroup in $rolesByService) {
                $serviceClass = switch ($serviceGroup.Name) {
                    "Azure AD/Entra ID" { "service-azure" }
                    "SharePoint Online" { "service-sharepoint" }
                    "Exchange Online" { "service-exchange" }
                    "Microsoft Teams" { "service-teams" }
                    "Microsoft Purview" { "service-purview" }
                    "Microsoft Intune" { "service-intune" }
                    "Microsoft Defender" { "service-defender" }
                    "Power Platform" { "service-powerplatform" }
                    default { "service-azure" }
                }
                
                $serviceGroupedRoles += "<div class='service-group'>"
                $serviceGroupedRoles += "<div class='service-group-header $serviceClass'>$($serviceGroup.Name)</div>"
                $serviceGroupedRoles += "<div class='service-roles'>"
                
                foreach ($role in $serviceGroup.Group) {
                    $serviceGroupedRoles += "<div class='role-item'>$($role.RoleName)</div>"
                }
                
                $serviceGroupedRoles += "</div></div>"
            }
            $serviceGroupedRoles += "</div>"
        } else {
            $serviceGroupedRoles = "<div class='no-roles'>No roles found</div>"
        }
        
        $html += @"
<tr class='expandable-user-row' onclick="toggleUserRoles('$userId')">
    <td $statusColor><strong>$($user.displayName)</strong><br><small style='color: #666;'>$($user.userPrincipalName)</small> <span class='expand-indicator'>▶</span></td>
    <td><strong>$($user.roleCount)</strong></td>
    <td><span class='tag $(if($status -eq "DISABLED") {"disabled"} else {"pim"})' $statusColor>$status</span></td>
    <td><small>$servicesDisplay</small></td>
</tr>
<tr id="userRoles_$userId" class="user-roles-row" style="display: none;">
    <td colspan="4" class="user-roles-cell">
        <div class="user-roles-title">Roles assigned to: $($user.displayName)</div>
        $serviceGroupedRoles
    </td>
</tr>
"@
    }

    $html += @"
            </table>
        </div>

        <div class="section">
            <h2>🔍 Principal Type Analysis</h2>
            <div class="stats-grid">
                <div class="stat-item">
                    <div class="stat-number">$($principalAnalysis.users)</div>
                    <div class="stat-label">Users</div>
                </div>
                <div class="stat-item">
                    <div class="stat-number">$($principalAnalysis.groups)</div>
                    <div class="stat-label">Groups</div>
                </div>
                <div class="stat-item">
                    <div class="stat-number">$($principalAnalysis.servicePrincipals)</div>
                    <div class="stat-label">Service Principals</div>
                </div>
                <div class="stat-item">
                    <div class="stat-number">$($principalAnalysis.onPremisesSyncedUsers)</div>
                    <div class="stat-label">On-Premises Synced</div>
                </div>
            </div>
        </div>

        <div class="section">
            <h2>🌐 Cross-Service Analysis</h2>
            <div class="stats-grid">
                <div class="stat-item">
                    <div class="stat-number">$($crossServiceAnalysis.usersWithMultipleServices)</div>
                    <div class="stat-label">Multi-Service Users</div>
                </div>
                <div class="stat-item">
                    <div class="stat-number">$($crossServiceAnalysis.exchangeAzureADCombinations)</div>
                    <div class="stat-label">Exchange + Azure AD</div>
                </div>
                <div class="stat-item">
                    <div class="stat-number">$($crossServiceAnalysis.exchangePurviewCombinations)</div>
                    <div class="stat-label">Exchange + Purview</div>
                </div>
            </div>
        </div>

        <div class="section">
            <h2>📋 Service-Specific Insights</h2>
            <div class="alert-grid">
"@

# Service-specific insights using helper function data
    foreach ($serviceName in $serviceAnalysis.Keys) {
        $serviceData = $serviceAnalysis[$serviceName]
        
        # Get PIM counts for this service using helper function
        $pimCounts = Get-ServicePIMCounts -AuditResults $AuditResults -ServiceName $serviceName
        
        # Determine service-specific color and icon
        $serviceInfo = switch ($serviceName) {
            "Azure AD/Entra ID" { @{ color = "#0078d4"; icon = "🔷" } }
            "SharePoint Online" { @{ color = "#0b6623"; icon = "📊" } }
            "Exchange Online" { @{ color = "#d13438"; icon = "📧" } }
            "Microsoft Teams" { @{ color = "#464775"; icon = "👥" } }
            "Microsoft Purview" { @{ color = "#8b4789"; icon = "🛡️" } }
            "Microsoft Intune" { @{ color = "#00bcf2"; icon = "📱" } }
            "Microsoft Defender" { @{ color = "#ff8c00"; icon = "🔐" } }
            "Power Platform" { @{ color = "#742774"; icon = "⚡" } }
            default { @{ color = "#0078d4"; icon = "🔒" } }
        }
        
        $html += @"
                <div class="security-alerts" style="border-color: $($serviceInfo.color);">
                    <h3 style="color: $($serviceInfo.color);">$($serviceInfo.icon) $serviceName Analysis</h3>
                    <div class="alert-item">📊 Total Assignments: $($serviceData.totalAssignments)</div>
                    <div class="alert-item">👥 Unique Users: $($serviceData.uniqueUsers)</div>
                    <div class="alert-item">🔴 PIM Active: $($pimCounts.pimActive)</div>
                    <div class="alert-item">🟢 PIM Eligible: $($pimCounts.pimEligible)</div>
                    <div class="alert-item">🎯 Top Role: $($serviceData.topRole)</div>
"@
        
        # Add service-specific analysis if available
        if ($serviceData.ContainsKey('exchangeAnalysis')) {
            $exData = $serviceData.exchangeAnalysis
            $html += @"
                    <div class="alert-item">🔄 On-Premises Synced: $($exData.onPremisesSyncedUsers)</div>
                    <div class="alert-item">👥 Group Assignments: $($exData.groupAssignments)</div>
                    <div class="alert-item">🏢 Hybrid Environment: $(if($exData.hybridEnvironment) {"Yes"} else {"No"})</div>
"@
        }
        
        if ($serviceData.ContainsKey('intuneAnalysis')) {
            $intData = $serviceData.intuneAnalysis
            $html += @"
                    <div class="alert-item">🔧 RBAC Assignments: $($intData.rbacAssignments)</div>
                    <div class="alert-item">🔷 Azure AD Assignments: $($intData.azureADAssignments)</div>
                    <div class="alert-item">⚙️ Service Administrators: $($intData.serviceAdministrators)</div>
"@
        }
        
        if ($serviceData.ContainsKey('sharePointAnalysis')) {
            $spData = $serviceData.sharePointAnalysis
            $html += @"
                    <div class="alert-item">🌐 Unique Sites: $($spData.uniqueSites)</div>
                    <div class="alert-item">👑 Site Collection Admins: $($spData.siteCollectionAdmins)</div>
                    <div class="alert-item">📱 App Catalog Admins: $($spData.appCatalogAdmins)</div>
"@
        }
        
        $html += "</div>"
    }

    $html += @"
            </div>
        </div>

        <div class="section">
            <h2>📋 Detailed Assignment Summary</h2>
            <div class="expandable" onclick="toggleSection('detailedAssignments')">
                <h3 style="cursor: pointer; color: var(--primary-color);">▶ View Detailed Assignments (Click to expand)</h3>
            </div>
            <div id="detailedAssignments" class="expandable-content">
                <p><em>Showing first 150 assignments. Export to CSV for complete data.</em></p>
                <table>
                    <tr>
                        <th>Service</th>
                        <th>User</th>
                        <th>Role</th>
                        <th>Assignment Type</th>
                        <th>Status</th>
                        <th>Scope</th>
                    </tr>
"@

    # Show first 150 assignments using formatted assignments from helper function
    $formattedAssignments = Get-FormattedAssignments -AuditResults $AuditResults
    $displayResults = $formattedAssignments | Select-Object -First 150
    
    foreach ($assignment in $displayResults) {
        $serviceClass = switch ($assignment.service) {
            "Azure AD/Entra ID" { "service-azure" }
            "SharePoint Online" { "service-sharepoint" }
            "Exchange Online" { "service-exchange" }
            "Microsoft Teams" { "service-teams" }
            "Microsoft Purview" { "service-purview" }
            "Microsoft Intune" { "service-intune" }
            "Microsoft Defender" { "service-defender" }
            "Power Platform" { "service-powerplatform" }
            default { "service-azure" }
        }
        
        $userDisplay = if ($assignment.userPrincipalName -and $assignment.userPrincipalName -ne "Unknown") { 
            $assignment.userPrincipalName 
        } else { 
            $assignment.displayName 
        }
        
        $scopeDisplay = if ($assignment.scope -and $assignment.scope.Length -gt 50) { 
            $assignment.scope.Substring(0, 47) + "..." 
        } else { 
            $assignment.scope 
        }
        
        # User status
        $userStatus = if ($assignment.userEnabled -eq $false) {
            "<span class='tag disabled'>DISABLED</span>"
        } else {
            "<span class='tag pim'>ACTIVE</span>"
        }
        
        # Assignment type styling
        $assignmentTypeDisplay = if ($assignment.assignmentType -like "*Eligible*") {
            "<span class='tag pim'>$($assignment.assignmentType)</span>"
        } elseif ($assignment.assignmentType -like "*PIM*") {
            "<span class='tag'>$($assignment.assignmentType)</span>"
        } else {
            "<span class='tag permanent'>$($assignment.assignmentType)</span>"
        }
        
        $html += @"
<tr>
    <td><span class='service-badge $serviceClass' style='font-size: 0.7em; padding: 2px 8px;'>$($assignment.service)</span></td>
    <td><small>$userDisplay</small></td>
    <td><strong>$($assignment.roleName)</strong></td>
    <td>$assignmentTypeDisplay</td>
    <td>$userStatus</td>
    <td><small>$scopeDisplay</small></td>
</tr>
"@
    }

    if ($AuditResults.Count -gt 150) {
        $html += @"
<tr>
    <td colspan='6' style='text-align: center; font-style: italic; color: #666; background: #f9f9f9;'>
        ... and $($AuditResults.Count - 150) more assignments<br>
        <small>Export to CSV for complete data set</small>
    </td>
</tr>
"@
    }

    $html += @"
                </table>
            </div>
        </div>

        <div class="section">
            <h2>💡 Security Recommendations</h2>
            <div class="security-alerts" style="background: linear-gradient(135deg, #e8f5e8, #c8e6c9); border-color: var(--success-color);">
                <h3 style="color: var(--success-color);">Actionable Recommendations</h3>
"@

    # Generate recommendations using helper function
    foreach ($category in @('immediate', 'shortTerm', 'longTerm')) {
        if ($recommendations.ContainsKey($category) -and $recommendations[$category].Count -gt 0) {
            $priorityLabel = switch ($category) {
                'immediate' { 'Immediate Action Required' }
                'shortTerm' { 'Short-term Improvements' }
                'longTerm' { 'Long-term Strategy' }
            }
            
            $html += "<div class='alert-item'><strong>${priorityLabel}:</strong></div>"
            foreach ($rec in $recommendations[$category]) {
                $html += "<div class='alert-item'>• $rec</div>"
            }
        }
    }

    $html += @"
            </div>
        </div>

        <div class="section">
            <h2>📈 Enhanced Reporting Options</h2>
            <div class="pim-analysis">
                <h3>Additional Analysis Available</h3>
                <div class="stats-grid">
                    <div class="stat-item">
                        <div class="stat-label"><strong>PowerShell Commands</strong></div>
                        <small>Export-M365AuditJsonReport</small><br>
                        <small>Get-M365RoleAnalysis</small><br>
                        <small>Get-M365ComplianceGaps</small>
                    </div>
                    <div class="stat-item">
                        <div class="stat-label"><strong>Automation</strong></div>
                        <small>New-M365AuditScheduledTask</small><br>
                        <small>Certificate-based auth</small><br>
                        <small>Unattended execution</small>
                    </div>
                    <div class="stat-item">
                        <div class="stat-label"><strong>Integration</strong></div>
                        <small>JSON export for SIEM</small><br>
                        <small>CSV for spreadsheet analysis</small><br>
                        <small>API-ready data format</small>
                    </div>
                </div>
            </div>
        </div>
"@

    # Add compliance analysis if requested and available
    if ($IncludeComplianceGaps -and $complianceAnalysis) {
        $html += @"
        <div class="section">
            <h2>⚖️ Compliance Analysis</h2>
            <div class="stats-grid">
                <div class="stat-item">
                    <div class="stat-number" style="color: $(if($complianceAnalysis.privilegedAccessCompliance.globalAdminLimit.compliant) {'var(--success-color)'} else {'var(--danger-color)'});">
                        $(if($complianceAnalysis.privilegedAccessCompliance.globalAdminLimit.compliant) {'✓'} else {'✗'})
                    </div>
                    <div class="stat-label">Global Admin Compliance</div>
                    <small>$($complianceAnalysis.privilegedAccessCompliance.globalAdminLimit.current)/$($complianceAnalysis.privilegedAccessCompliance.globalAdminLimit.recommended) admins</small>
                </div>
                <div class="stat-item">
                    <div class="stat-number" style="color: $(if($complianceAnalysis.privilegedAccessCompliance.pimImplementation.compliant) {'var(--success-color)'} else {'var(--warning-color)'});">
                        $(if($complianceAnalysis.privilegedAccessCompliance.pimImplementation.compliant) {'✓'} else {'△'})
                    </div>
                    <div class="stat-label">PIM Implementation</div>
                    <small>$($complianceAnalysis.privilegedAccessCompliance.pimImplementation.eligibleCount) eligible assignments</small>
                </div>
                <div class="stat-item">
                    <div class="stat-number" style="color: $(if($complianceAnalysis.privilegedAccessCompliance.disabledAccountCleanup.compliant) {'var(--success-color)'} else {'var(--danger-color)'});">
                        $(if($complianceAnalysis.privilegedAccessCompliance.disabledAccountCleanup.compliant) {'✓'} else {'✗'})
                    </div>
                    <div class="stat-label">Account Cleanup</div>
                    <small>$($complianceAnalysis.privilegedAccessCompliance.disabledAccountCleanup.violationCount) violations</small>
                </div>
                <div class="stat-item">
                    <div class="stat-number" style="color: $(if($complianceAnalysis.authenticationCompliance.certificateBasedAuth.compliant) {'var(--success-color)'} else {'var(--warning-color)'});">
                        $($complianceAnalysis.authenticationCompliance.certificateBasedAuth.percentage)%
                    </div>
                    <div class="stat-label">Certificate Auth Usage</div>
                    <small>Secure authentication percentage</small>
                </div>
            </div>
        </div>
"@
    }

    $html += @"
        <div class="footer">
            <p><strong>Microsoft 365 Comprehensive Role Audit Report</strong></p>
            <p>Generated by Enhanced PowerShell Audit Module v2.0</p>
            <p>Authentication: Certificate-based (Secure) | Report Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
            <p>For complete data analysis, export results to CSV format</p>
            <p style="margin-top: 15px; font-size: 0.9em; color: #888;">
                This report provides comprehensive role assignment analysis across Microsoft 365 services.<br>
                Regular auditing helps maintain security posture and compliance requirements.
            </p>
        </div>
    </div>
</body>
</html>
"@

    try {
        $html | Out-File -FilePath $OutputPath -Encoding UTF8
        Write-Host "✓ Enhanced HTML report generated using helper functions: $OutputPath" -ForegroundColor Green
        
        # Show report summary using helper function
        Show-ReportSummary -Report @{
            metadata = $metadata
            securityAlerts = $securityAlerts
        } -OutputPath $OutputPath
        
        # Open report if on Windows
        if ($IsWindows -ne $false -and (Test-Path $OutputPath)) {
            $openReport = Read-Host "Open enhanced report in browser? (y/N)"
            if ($openReport -eq "y" -or $openReport -eq "Y") {
                Start-Process $OutputPath
            }
        }
        
        return $OutputPath
    }
    catch {
        Write-Error "Failed to generate enhanced HTML report: $($_.Exception.Message)"
        return $null
    }
}


function Export-M365AuditJsonReport {
    param(
        [Parameter(Mandatory = $true)]
        [array]$AuditResults,
        
        [string]$OutputPath = ".\M365_Audit_Data_$(Get-Date -Format 'yyyyMMdd_HHmmss').json",
        [string]$OrganizationName = "Organization",
        [switch]$IncludeComplianceAnalysis,
        [switch]$IncludePIMAnalysis,
        [switch]$IncludeExchangeAnalysis
    )
    
    Write-Host "Generating comprehensive JSON audit report..." -ForegroundColor Cyan
    
    if ($AuditResults.Count -eq 0) {
        Write-Warning "No audit results provided"
        return
    }
    
    # Calculate basic statistics
    $stats = Get-AuditStatistics -AuditResults $AuditResults
    
    # Build comprehensive JSON report
    $report = @{
        metadata = Get-ReportMetadata -OrganizationName $OrganizationName -Stats $stats -AuditResults $AuditResults
        summary = Get-ReportSummary -AuditResults $AuditResults -Stats $stats
        serviceAnalysis = Get-ServiceAnalysis -AuditResults $AuditResults -IncludeExchangeAnalysis:$IncludeExchangeAnalysis
        pimAnalysis = Get-PIMAnalysis -AuditResults $AuditResults -IncludeDetailedAnalysis:$IncludePIMAnalysis
        principalAnalysis = Get-PrincipalAnalysis -AuditResults $AuditResults
        crossServiceAnalysis = Get-CrossServiceAnalysis -AuditResults $AuditResults
        securityAlerts = Get-SecurityAlerts -AuditResults $AuditResults -Stats $stats
        recommendations = Get-SecurityRecommendations -AuditResults $AuditResults -Stats $stats
        assignments = Get-FormattedAssignments -AuditResults $AuditResults
    }
    
    # Add compliance analysis if requested
    if ($IncludeComplianceAnalysis) {
        $report.complianceAnalysis = Get-ComplianceAnalysis -AuditResults $AuditResults -Stats $stats
    }
    
    try {
        $jsonOutput = $report | ConvertTo-Json -Depth 15
        $jsonOutput | Out-File -FilePath $OutputPath -Encoding UTF8
        
        Write-Host "✓ Enhanced JSON report generated: $OutputPath" -ForegroundColor Green
        Show-ReportSummary -Report $report -OutputPath $OutputPath
        
        return $OutputPath
    }
    catch {
        Write-Error "Failed to generate enhanced JSON report: $($_.Exception.Message)"
        return $null
    }
}

function Get-M365RoleAnalysis {
    param(
        [Parameter(Mandatory = $true)]
        [array]$AuditResults,
        [switch]$IncludePIMAnalysis,
        [switch]$IncludeIntuneAnalysis,
        [switch]$IncludePowerPlatformAnalysis,
        [switch]$IncludeExchangeAnalysis,
        [switch]$ShowDetailedRecommendations,
        [switch]$ReturnStructuredData
    )
    
    Write-Host "=== M365 Comprehensive Role Analysis ===" -ForegroundColor Green
    
    if ($AuditResults.Count -eq 0) {
        Write-Warning "No audit results provided"
        return
    }
    
    # Use centralized statistics calculation from core functions
    Write-Host ""
    Write-Host "Calculating comprehensive statistics..." -ForegroundColor Cyan
    $stats = Get-AuditStatistics -AuditResults $AuditResults
    
    # Initialize analysis results object
    $analysisResults = @{
        Summary = $stats
        RoleSprawl = @{}
        OrphanedAccounts = @{}
        PrivilegeEscalation = @{}
        ServiceSpecific = @{}
        CrossService = @{}
        PIM = @{}
        Intune = @{}
        PowerPlatform = @{}
        Exchange = @{}
        Security = @{}
        Recommendations = @()
        ComplianceAnalysis = @{}
        SecurityAlerts = @{}
    }
    
    # === BASIC SUMMARY ANALYSIS ===
    Write-Host ""
    Write-Host "Basic Summary Analysis:" -ForegroundColor Cyan
    Write-Host "Total Assignments: $($stats.totalAssignments)" -ForegroundColor White
    Write-Host "Unique Users: $($stats.uniqueUsers)" -ForegroundColor White
    Write-Host "Services Audited: $($stats.servicesAudited)" -ForegroundColor White
    Write-Host "Authentication Methods:" -ForegroundColor White
    foreach ($authType in $stats.authTypes) {
        $color = if ($authType.Name -eq "Certificate") { "Green" } elseif ($authType.Name -eq "ClientSecret") { "Yellow" } else { "White" }
        Write-Host "  $($authType.Name): $($authType.Count)" -ForegroundColor $color
    }
    
    # === ROLE SPRAWL ANALYSIS ===
    Write-Host ""
    Write-Host "Role Sprawl Analysis:" -ForegroundColor Cyan
    $usersWithMultipleRoles = $AuditResults | Group-Object UserPrincipalName | Where-Object { 
        $_.Name -and $_.Name -ne "Unknown" -and $_.Count -gt 5 
    } | Sort-Object Count -Descending
    
    $analysisResults.RoleSprawl = @{
        UsersWithExcessiveRoles = $usersWithMultipleRoles
        ThresholdExceeded = $usersWithMultipleRoles.Count
    }
    
    if ($usersWithMultipleRoles.Count -gt 0) {
        Write-Host "⚠️ Users with excessive role assignments (>5 roles): $($usersWithMultipleRoles.Count)" -ForegroundColor Yellow
        foreach ($user in $usersWithMultipleRoles | Select-Object -First 10) {
            Write-Host "  $($user.Name): $($user.Count) roles" -ForegroundColor White
            
            # Show role distribution for top users
            $userRoles = $AuditResults | Where-Object { $_.UserPrincipalName -eq $user.Name }
            $rolesByService = $userRoles | Group-Object Service
            Write-Host "    Services: $($rolesByService.Name -join ', ')" -ForegroundColor Gray
        }
        
        $analysisResults.Recommendations += "Review users with excessive role assignments for principle of least privilege"
    }
    else {
        Write-Host "✓ No users found with excessive role assignments" -ForegroundColor Green
    }
    
    # === ORPHANED ACCOUNTS ANALYSIS ===
    Write-Host ""
    Write-Host "Orphaned Accounts Analysis:" -ForegroundColor Cyan
    
    $analysisResults.OrphanedAccounts = @{
        DisabledUsers = $stats.disabledUsers
        InactiveUsers = $stats.usersWithoutRecentSignIn
        SystemAccounts = $stats.systemGeneratedAccounts
    }
    
    if ($stats.disabledUsers.Count -gt 0) {
        Write-Host "⚠️ Disabled users with active roles: $($stats.disabledUsers.Count)" -ForegroundColor Red
        $disabledUsersGrouped = $stats.disabledUsers | Group-Object UserPrincipalName
        foreach ($user in $disabledUsersGrouped | Select-Object -First 5) {
            if ($user.Name -and $user.Name -ne "Unknown") {
                Write-Host "  $($user.Name): $($user.Count) active role(s)" -ForegroundColor Yellow
            }
        }
        $analysisResults.Recommendations += "Remove role assignments from disabled user accounts"
    }
    else {
        Write-Host "✓ No disabled users with active roles found" -ForegroundColor Green
    }
    
    if ($stats.usersWithoutRecentSignIn.Count -gt 0) {
        Write-Host "⚠️ Users with roles but no sign-in in 90+ days: $($stats.usersWithoutRecentSignIn.Count)" -ForegroundColor Yellow
        $analysisResults.Recommendations += "Review role assignments for users without recent sign-in activity"
    }
    
    if ($stats.systemGeneratedAccounts.Count -gt 0) {
        Write-Host "ℹ️ System-generated accounts/policies: $($stats.systemGeneratedAccounts.Count)" -ForegroundColor Cyan
    }
    
    # === PRIVILEGE ESCALATION ANALYSIS ===
    Write-Host ""
    Write-Host "Privilege Escalation Analysis:" -ForegroundColor Cyan
    
    $analysisResults.PrivilegeEscalation = @{
        GlobalAdmins = $stats.globalAdmins
        PrivilegedRoles = $stats.privilegedRoles
        SecurityRoles = $stats.securityRoles
        ComplianceRoles = $stats.complianceRoles
    }
    
    Write-Host "Global Administrators: $($stats.globalAdmins.Count)" -ForegroundColor $(if ($stats.globalAdmins.Count -gt 5) { "Red" } else { "Green" })
    Write-Host "Other Administrative Roles: $($stats.privilegedRoles.Count)" -ForegroundColor Gray
    Write-Host "Security-related Roles: $($stats.securityRoles.Count)" -ForegroundColor Gray
    Write-Host "Compliance-related Roles: $($stats.complianceRoles.Count)" -ForegroundColor Gray
    
    if ($stats.globalAdmins.Count -gt 5) {
        $analysisResults.Recommendations += "Reduce Global Administrator count to 5 or fewer (currently $($stats.globalAdmins.Count))"
    }
    
    # === SERVICE-SPECIFIC ANALYSIS ===
    Write-Host ""
    Write-Host "Service-Specific Analysis:" -ForegroundColor Cyan
    
    # Use the enhanced service analysis function from core
    $analysisResults.ServiceSpecific = Get-ServiceAnalysis -AuditResults $AuditResults -IncludeExchangeAnalysis:$IncludeExchangeAnalysis
    
    # Display service analysis summary
    foreach ($serviceName in $analysisResults.ServiceSpecific.Keys) {
        $serviceData = $analysisResults.ServiceSpecific[$serviceName]
        Write-Host "$serviceName`: $($serviceData.totalAssignments) assignments, $($serviceData.uniqueUsers) users" -ForegroundColor White
        Write-Host "  Top Role: $($serviceData.topRole)" -ForegroundColor Gray
        
        # Display service-specific insights
        if ($serviceData.ContainsKey('exchangeAnalysis')) {
            Write-Host "  Exchange: $($serviceData.exchangeAnalysis.roleGroupAssignments) role groups, $($serviceData.exchangeAnalysis.onPremisesSyncedUsers) synced users" -ForegroundColor Gray
        }
        if ($serviceData.ContainsKey('intuneAnalysis')) {
            Write-Host "  Intune: $($serviceData.intuneAnalysis.rbacAssignments) RBAC, $($serviceData.intuneAnalysis.serviceAdministrators) service admins" -ForegroundColor Gray
        }
        if ($serviceData.ContainsKey('sharePointAnalysis')) {
            Write-Host "  SharePoint: $($serviceData.sharePointAnalysis.uniqueSites) sites, $($serviceData.sharePointAnalysis.siteCollectionAdmins) site admins" -ForegroundColor Gray
        }
    }
    
    # === CROSS-SERVICE PRIVILEGE ANALYSIS ===
    Write-Host ""
    Write-Host "Cross-Service Privilege Analysis:" -ForegroundColor Cyan
    
    # Use the cross-service analysis function from core
    $analysisResults.CrossService = Get-CrossServiceAnalysis -AuditResults $AuditResults
    
    Write-Host "Users with roles across multiple services: $($analysisResults.CrossService.usersWithMultipleServices)" -ForegroundColor White
    Write-Host "Exchange + Azure AD combinations: $($analysisResults.CrossService.exchangeAzureADCombinations)" -ForegroundColor Gray
    Write-Host "Exchange + Purview combinations: $($analysisResults.CrossService.exchangePurviewCombinations)" -ForegroundColor Gray
    
    if ($analysisResults.CrossService.usersWithMultipleServices -gt 0) {
        $analysisResults.Recommendations += "Review users with high-risk cross-service role combinations"
    }
    
    # === PIM ANALYSIS ===
    if ($IncludePIMAnalysis) {
        Write-Host ""
        Write-Host "PIM (Privileged Identity Management) Analysis:" -ForegroundColor Cyan
        
        # Use the enhanced PIM analysis function from core
        $analysisResults.PIM = Get-PIMAnalysis -AuditResults $AuditResults -IncludeDetailedAnalysis:$true
        
        Write-Host "PIM Enabled: $($analysisResults.PIM.enabled)" -ForegroundColor $(if($analysisResults.PIM.enabled) {"Green"} else {"Yellow"})
        Write-Host "PIM Eligible Assignments: $($analysisResults.PIM.totalEligible)" -ForegroundColor $(if($analysisResults.PIM.totalEligible -gt 0) {"Green"} else {"Yellow"})
        Write-Host "PIM Active Assignments: $($analysisResults.PIM.totalActive)" -ForegroundColor White
        
        if ($analysisResults.PIM.detailed) {
            Write-Host "Expiring within 30 days: $($analysisResults.PIM.detailed.eligible.expiringWithin30Days)" -ForegroundColor $(if($analysisResults.PIM.detailed.eligible.expiringWithin30Days -gt 0) {"Yellow"} else {"Green"})
        }
        
        # PIM recommendations
        if (-not $analysisResults.PIM.enabled) {
            $analysisResults.Recommendations += "Consider implementing PIM for eligible assignments to reduce standing privileges"
        }
        
        if ($analysisResults.PIM.detailed.eligible.expiringWithin30Days -gt 0) {
            $analysisResults.Recommendations += "Review and renew expiring PIM assignments"
        }
    }
    
    # === INTUNE-SPECIFIC ANALYSIS ===
    if ($IncludeIntuneAnalysis) {
        Write-Host ""
        Write-Host "Intune-Specific Analysis:" -ForegroundColor Cyan
        
        $intuneResults = $AuditResults | Where-Object { $_.Service -eq "Microsoft Intune" }
        if ($intuneResults.Count -gt 0) {
            # Use the Intune analysis function from core if available, otherwise use existing logic
            if ($analysisResults.ServiceSpecific.ContainsKey('Microsoft Intune') -and 
                $analysisResults.ServiceSpecific['Microsoft Intune'].ContainsKey('intuneAnalysis')) {
                
                $intuneAnalysis = $analysisResults.ServiceSpecific['Microsoft Intune'].intuneAnalysis
                $analysisResults.Intune = $intuneAnalysis
                
                Write-Host "Total Intune Assignments: $($intuneResults.Count)" -ForegroundColor White
                Write-Host "RBAC Assignments: $($intuneAnalysis.rbacAssignments)" -ForegroundColor White
                Write-Host "Azure AD Assignments: $($intuneAnalysis.azureADAssignments)" -ForegroundColor White
                Write-Host "Service Administrators: $($intuneAnalysis.serviceAdministrators)" -ForegroundColor $(if($intuneAnalysis.serviceAdministrators -le 3) {"Green"} else {"Yellow"})
                Write-Host "Built-in Roles: $($intuneAnalysis.builtInRoles)" -ForegroundColor White
                Write-Host "Custom Roles: $($intuneAnalysis.customRoles)" -ForegroundColor White
                
                # Intune-specific recommendations
                if ($intuneAnalysis.serviceAdministrators -gt 3) {
                    $analysisResults.Recommendations += "Consider reducing Intune Service Administrator count (currently $($intuneAnalysis.serviceAdministrators))"
                }
                
                if ($intuneAnalysis.azureADAssignments -gt $intuneAnalysis.rbacAssignments) {
                    $analysisResults.Recommendations += "Consider using Intune RBAC roles instead of Azure AD roles for better granularity"
                }
                
                if ($intuneAnalysis.customRoles -eq 0 -and $intuneResults.Count -gt 20) {
                    $analysisResults.Recommendations += "Consider creating custom Intune roles for specific administrative needs"
                }
            }
        }
        else {
            Write-Host "No Intune assignments found in audit results" -ForegroundColor Yellow
        }
    }
    
    # === POWER PLATFORM ANALYSIS ===
    if ($IncludePowerPlatformAnalysis) {
        Write-Host ""
        Write-Host "Power Platform Analysis:" -ForegroundColor Cyan
        
        $powerPlatformResults = $AuditResults | Where-Object { $_.Service -eq "Power Platform" }
        if ($powerPlatformResults.Count -gt 0) {
            $powerPlatformAdmins = $powerPlatformResults | Where-Object { $_.RoleName -eq "Power Platform Administrator" }
            $dynamicsAdmins = $powerPlatformResults | Where-Object { $_.RoleName -like "*Dynamics*" }
            $powerBIAdmins = $powerPlatformResults | Where-Object { $_.RoleName -like "*Power BI*" }
            $principalTypes = $powerPlatformResults | Group-Object PrincipalType
            
            $analysisResults.PowerPlatform = @{
                TotalAssignments = $powerPlatformResults.Count
                PowerPlatformAdmins = $powerPlatformAdmins
                DynamicsAdmins = $dynamicsAdmins
                PowerBIAdmins = $powerBIAdmins
                PrincipalTypes = $principalTypes
            }
            
            Write-Host "Total Power Platform Assignments: $($powerPlatformResults.Count)" -ForegroundColor White
            Write-Host "Power Platform Administrators: $($powerPlatformAdmins.Count)" -ForegroundColor White
            Write-Host "Dynamics 365 Administrators: $($dynamicsAdmins.Count)" -ForegroundColor White
            Write-Host "Power BI Administrators: $($powerBIAdmins.Count)" -ForegroundColor White
            
            Write-Host "Principal Types:" -ForegroundColor Cyan
            foreach ($principalType in $principalTypes) {
                Write-Host "  $($principalType.Name): $($principalType.Count)" -ForegroundColor Gray
            }
            
            # Check for service principals with Power Platform access
            $servicePrincipals = $powerPlatformResults | Where-Object { $_.PrincipalType -eq "ServicePrincipal" }
            if ($servicePrincipals.Count -gt 0) {
                Write-Host "⚠️ Service Principals with Power Platform access: $($servicePrincipals.Count)" -ForegroundColor Yellow
                $analysisResults.Recommendations += "Review service principal access to Power Platform resources"
            }
        }
        else {
            Write-Host "No Power Platform assignments found in audit results" -ForegroundColor Yellow
        }
    }
    
    # === SECURITY ANALYSIS ===
    Write-Host ""
    Write-Host "Security Analysis:" -ForegroundColor Cyan
    
    # Use the security alerts function from core
    $analysisResults.SecurityAlerts = Get-SecurityAlerts -AuditResults $AuditResults -Stats $stats
    $analysisResults.Security = @{
        CertificateAuthCount = ($stats.authTypes | Where-Object { $_.Name -eq "Certificate" }).Count
        ClientSecretAuthCount = ($stats.authTypes | Where-Object { $_.Name -eq "ClientSecret" }).Count
        SecurityScore = 0
        GlobalAdminCount = $analysisResults.SecurityAlerts.globalAdminCount
        DisabledUsersWithRoles = $analysisResults.SecurityAlerts.disabledUsersWithRoles
        CertificateBasedAuthEnabled = $analysisResults.SecurityAlerts.certificateBasedAuth
        ClientSecretAuthUsage = $analysisResults.SecurityAlerts.clientSecretAuth
    }
    
    # Calculate basic security score
    $securityScore = 0
    if ($analysisResults.Security.CertificateAuthCount -gt 0 -and $analysisResults.Security.ClientSecretAuthCount -eq 0) { $securityScore += 25 }
    if ($stats.globalAdmins.Count -le 5) { $securityScore += 25 }
    if ($stats.disabledUsers.Count -eq 0) { $securityScore += 25 }
    if ($stats.pimEligible.Count -gt 0) { $securityScore += 25 }
    
    $analysisResults.Security.SecurityScore = $securityScore
    
    Write-Host "Certificate Authentication: $($analysisResults.Security.CertificateAuthCount)" -ForegroundColor Green
    Write-Host "Client Secret Authentication: $($analysisResults.Security.ClientSecretAuthCount)" -ForegroundColor $(if($analysisResults.Security.ClientSecretAuthCount -eq 0) {"Green"} else {"Yellow"})
    Write-Host "Security Score: $securityScore/100" -ForegroundColor $(if($securityScore -ge 75) {"Green"} elseif($securityScore -ge 50) {"Yellow"} else {"Red"})
    
    # Display security alerts
    if ($analysisResults.SecurityAlerts.critical.Count -gt 0) {
        Write-Host "CRITICAL ALERTS:" -ForegroundColor Red
        foreach ($alert in $analysisResults.SecurityAlerts.critical) {
            Write-Host "  • $alert" -ForegroundColor Red
        }
    }
    
    if ($analysisResults.SecurityAlerts.high.Count -gt 0) {
        Write-Host "HIGH ALERTS:" -ForegroundColor Yellow
        foreach ($alert in $analysisResults.SecurityAlerts.high) {
            Write-Host "  • $alert" -ForegroundColor Yellow
        }
    }
    
    # === COMPLIANCE ANALYSIS ===
    Write-Host ""
    Write-Host "Compliance Analysis:" -ForegroundColor Cyan
    
    # Use the compliance analysis function from core
    $analysisResults.ComplianceAnalysis = Get-ComplianceAnalysis -AuditResults $AuditResults -Stats $stats
    
    $privAccess = $analysisResults.ComplianceAnalysis.privilegedAccessCompliance
    Write-Host "Global Admin Compliance: $(if($privAccess.globalAdminLimit.compliant) {"✓ Compliant"} else {"⚠ Non-compliant"})" -ForegroundColor $(if($privAccess.globalAdminLimit.compliant) {"Green"} else {"Red"})
    Write-Host "PIM Implementation: $(if($privAccess.pimImplementation.compliant) {"✓ Implemented"} else {"⚠ Not implemented"})" -ForegroundColor $(if($privAccess.pimImplementation.compliant) {"Green"} else {"Yellow"})
    Write-Host "Disabled Account Cleanup: $(if($privAccess.disabledAccountCleanup.compliant) {"✓ Clean"} else {"⚠ Needs cleanup"})" -ForegroundColor $(if($privAccess.disabledAccountCleanup.compliant) {"Green"} else {"Red"})
    
    $authCompliance = $analysisResults.ComplianceAnalysis.authenticationCompliance
    Write-Host "Certificate Auth Usage: $($authCompliance.certificateBasedAuth.percentage)%" -ForegroundColor $(if($authCompliance.certificateBasedAuth.compliant) {"Green"} else {"Yellow"})
    
    # === DETAILED RECOMMENDATIONS ===
    if ($ShowDetailedRecommendations) {
        Write-Host ""
        Write-Host "=== DETAILED RECOMMENDATIONS ===" -ForegroundColor Yellow
        
        # Use the recommendations function from core
        $detailedRecommendations = Get-SecurityRecommendations -AuditResults $AuditResults -Stats $stats
        
        if ($detailedRecommendations.immediate.Count -gt 0) {
            Write-Host "IMMEDIATE ACTION REQUIRED:" -ForegroundColor Red
            foreach ($rec in $detailedRecommendations.immediate) {
                Write-Host "  • $rec" -ForegroundColor White
            }
        }
        
        if ($detailedRecommendations.shortTerm.Count -gt 0) {
            Write-Host "SHORT-TERM (1-3 months):" -ForegroundColor Yellow
            foreach ($rec in $detailedRecommendations.shortTerm) {
                Write-Host "  • $rec" -ForegroundColor White
            }
        }
        
        if ($detailedRecommendations.longTerm.Count -gt 0) {
            Write-Host "LONG-TERM (3+ months):" -ForegroundColor Cyan
            foreach ($rec in $detailedRecommendations.longTerm) {
                Write-Host "  • $rec" -ForegroundColor White
            }
        }
        
        # Merge recommendations from detailed analysis
        $analysisResults.Recommendations += $detailedRecommendations.immediate
        $analysisResults.Recommendations += $detailedRecommendations.shortTerm
        $analysisResults.Recommendations += $detailedRecommendations.longTerm
    }
    
    Write-Host ""
    Write-Host "✓ Role analysis completed successfully" -ForegroundColor Green
    Write-Host "Analysis covered $($AuditResults.Count) role assignments across $($stats.servicesAudited) services" -ForegroundColor Cyan
    
    # Return structured data if requested, otherwise return analysis results
    if ($ReturnStructuredData) {
        return $analysisResults
    }
    
    return $analysisResults
}

# Helper function to get audit statistics (if not already in core functions)


function Get-M365ComplianceGaps {
    param(
        [Parameter(Mandatory = $true)]
        [array]$AuditResults,
        [switch]$IncludeDetailedAnalysis,
        [switch]$IncludePIMGaps,
        [switch]$IncludeIntuneGaps,
        [switch]$IncludePowerPlatformGaps
    )
    
    Write-Host "=== M365 Comprehensive Compliance Gap Analysis ===" -ForegroundColor Green
    
    if ($AuditResults.Count -eq 0) {
        Write-Warning "No audit results provided"
        return @()
    }
    
    $gaps = @()
    $criticalGaps = @()
    $highGaps = @()
    $mediumGaps = @()
    $lowGaps = @()
    
    # === IDENTITY GOVERNANCE GAPS ===
    Write-Host "Analyzing Identity Governance..." -ForegroundColor Cyan
    
    # Check Global Admin count
    $globalAdmins = $AuditResults | Where-Object { $_.RoleName -eq "Global Administrator" }
    if ($globalAdmins.Count -gt 5) {
        $gap = [PSCustomObject]@{
            Category = "Identity Governance"
            Issue = "Excessive Global Administrators"
            Details = "$($globalAdmins.Count) Global Admin accounts (recommended: ≤5)"
            Severity = "Critical"
            Recommendation = "Reduce Global Admin count using principle of least privilege and role-specific admin roles"
            AffectedUsers = ($globalAdmins | Select-Object -ExpandProperty UserPrincipalName) -join "; "
            ComplianceFramework = "ISO 27001, NIST, CIS Controls"
            RemediationSteps = @(
                "1. Review each Global Admin's actual responsibilities",
                "2. Assign appropriate role-specific admin roles",
                "3. Remove Global Admin role where not required",
                "4. Implement break-glass accounts for emergency access"
            )
        }
        $criticalGaps += $gap
        $gaps += $gap
    }
    
    # Check for disabled users with roles
    $disabledWithRoles = $AuditResults | Where-Object { $_.UserEnabled -eq $false }
    if ($disabledWithRoles.Count -gt 0) {
        $affectedUsers = ($disabledWithRoles | Group-Object UserPrincipalName | Select-Object -ExpandProperty Name) -join "; "
        $gap = [PSCustomObject]@{
            Category = "Access Management"
            Issue = "Disabled Users with Active Roles"
            Details = "$($disabledWithRoles.Count) disabled users still have role assignments"
            Severity = "High"
            Recommendation = "Implement automated role removal process for disabled accounts"
            AffectedUsers = $affectedUsers
            ComplianceFramework = "SOX, PCI DSS, GDPR"
            RemediationSteps = @(
                "1. Identify all disabled users with role assignments",
                "2. Remove role assignments from disabled accounts",
                "3. Implement automated workflow to remove roles when accounts are disabled",
                "4. Regular review of disabled account permissions"
            )
        }
        $highGaps += $gap
        $gaps += $gap
    }
    
    # === AUTHENTICATION SECURITY GAPS ===
    Write-Host "Analyzing Authentication Security..." -ForegroundColor Cyan
    
    # Check authentication methods
    $clientSecretAuth = $AuditResults | Where-Object { $_.AuthenticationType -eq "ClientSecret" }
    if ($clientSecretAuth.Count -gt 0) {
        $gap = [PSCustomObject]@{
            Category = "Authentication Security"
            Issue = "Client Secret Authentication Usage"
            Details = "$($clientSecretAuth.Count) connections use client secret instead of certificate authentication"
            Severity = "Medium"
            Recommendation = "Migrate to certificate-based authentication for enhanced security"
            AffectedUsers = "Application Authentication"
            ComplianceFramework = "NIST Cybersecurity Framework, Zero Trust"
            RemediationSteps = @(
                "1. Generate X.509 certificates for app registrations",
                "2. Upload certificates to Azure AD app registrations",
                "3. Update automation scripts to use certificate authentication",
                "4. Remove client secrets after migration"
            )
        }
        $mediumGaps += $gap
        $gaps += $gap
    }
    
    # === PIM AND PRIVILEGED ACCESS GAPS ===
    if ($IncludePIMGaps) {
        Write-Host "Analyzing Privileged Identity Management..." -ForegroundColor Cyan
        
        # Check for PIM usage
        $eligibleAssignments = $AuditResults | Where-Object { $_.AssignmentType -like "*Eligible*" }
        $activeAssignments = $AuditResults | Where-Object { 
            $_.AssignmentType -eq "Active" -or 
            $_.AssignmentType -eq "Azure AD Role" -or
            $_.AssignmentType -eq "Intune RBAC"
        }
        
        if ($eligibleAssignments.Count -eq 0 -and $activeAssignments.Count -gt 0) {
            $gap = [PSCustomObject]@{
                Category = "Privileged Access Management"
                Issue = "No PIM Eligible Assignments"
                Details = "All $($activeAssignments.Count) privileged roles are permanently assigned (no PIM eligible assignments found)"
                Severity = "High"
                Recommendation = "Implement Privileged Identity Management (PIM) for just-in-time access to privileged roles"
                AffectedUsers = "All privileged users"
                ComplianceFramework = "NIST 800-53, ISO 27001, Zero Trust"
                RemediationSteps = @(
                    "1. Enable Azure AD PIM licensing",
                    "2. Identify roles suitable for PIM eligible assignments",
                    "3. Convert permanent assignments to eligible assignments",
                    "4. Configure approval workflows and access reviews"
                )
            }
            $highGaps += $gap
            $gaps += $gap
        }
        
        # Check PIM adoption rate
        $totalPrivilegedAssignments = $eligibleAssignments.Count + $activeAssignments.Count
        if ($totalPrivilegedAssignments -gt 0) {
            $pimAdoptionRate = [math]::Round(($eligibleAssignments.Count / $totalPrivilegedAssignments) * 100, 2)
            if ($pimAdoptionRate -lt 30 -and $eligibleAssignments.Count -gt 0) {
                $gap = [PSCustomObject]@{
                    Category = "Privileged Access Management"
                    Issue = "Low PIM Adoption Rate"
                    Details = "PIM adoption rate is only $pimAdoptionRate% ($($eligibleAssignments.Count) eligible vs $($activeAssignments.Count) permanent)"
                    Severity = "Medium"
                    Recommendation = "Expand PIM usage to cover more privileged roles"
                    AffectedUsers = "Privileged role users"
                    ComplianceFramework = "Zero Trust, NIST"
                    RemediationSteps = @(
                        "1. Review current permanently assigned roles",
                        "2. Evaluate which roles can be converted to eligible",
                        "3. Implement phased PIM rollout",
                        "4. Train users on PIM activation process"
                    )
                }
                $mediumGaps += $gap
                $gaps += $gap
            }
        }
        
        # Check for expiring PIM assignments
        $expiringPIMAssignments = $AuditResults | Where-Object { 
            $_.PIMEndDateTime -and 
            [DateTime]$_.PIMEndDateTime -lt (Get-Date).AddDays(30)
        }
        
        if ($expiringPIMAssignments.Count -gt 0) {
            $gap = [PSCustomObject]@{
                Category = "Privileged Access Management"
                Issue = "Expiring PIM Assignments"
                Details = "$($expiringPIMAssignments.Count) PIM assignments expire within 30 days"
                Severity = "Medium"
                Recommendation = "Review and renew expiring PIM assignments to prevent access disruption"
                AffectedUsers = ($expiringPIMAssignments | Select-Object -ExpandProperty UserPrincipalName -Unique) -join "; "
                ComplianceFramework = "Access Management"
                RemediationSteps = @(
                    "1. Review expiring PIM assignments",
                    "2. Validate continued business need",
                    "3. Renew or remove assignments as appropriate",
                    "4. Implement automated renewal notifications"
                )
            }
            $mediumGaps += $gap
            $gaps += $gap
        }
    }
    
    # === ACCOUNT MANAGEMENT GAPS ===
    Write-Host "Analyzing Account Management..." -ForegroundColor Cyan
    
    # Check for service accounts or shared accounts
    $potentialServiceAccounts = $AuditResults | Where-Object { 
        $_.UserPrincipalName -like "*service*" -or 
        $_.UserPrincipalName -like "*admin*" -or
        $_.UserPrincipalName -like "*shared*" -or
        $_.UserPrincipalName -like "*system*"
    } | Where-Object { $_.UserPrincipalName -ne "System Generated" }
    
    if ($potentialServiceAccounts.Count -gt 0) {
        $serviceAccountUsers = ($potentialServiceAccounts | Group-Object UserPrincipalName | Select-Object -ExpandProperty Name) -join "; "
        $gap = [PSCustomObject]@{
            Category = "Account Management"
            Issue = "Potential Service/Shared Accounts"
            Details = "Found $($potentialServiceAccounts.Count) accounts that may be service/shared accounts with privileged access"
            Severity = "Medium"
            Recommendation = "Review account naming conventions, implement managed identities where possible"
            AffectedUsers = $serviceAccountUsers
            ComplianceFramework = "CIS Controls, NIST"
            RemediationSteps = @(
                "1. Review account usage patterns and ownership",
                "2. Convert to managed identities where applicable",
                "3. Implement proper service account governance",
                "4. Regular review of service account permissions"
            )
        }
        $mediumGaps += $gap
        $gaps += $gap
    }
    
    # Check for users without recent sign-in
    $usersWithoutRecentSignIn = $AuditResults | Where-Object { 
        $_.LastSignIn -and $_.LastSignIn -lt (Get-Date).AddDays(-90) -and $_.UserEnabled -eq $true
    }
    
    if ($usersWithoutRecentSignIn.Count -gt 0) {
        $inactiveUsers = ($usersWithoutRecentSignIn | Group-Object UserPrincipalName | Select-Object -ExpandProperty Name) -join "; "
        $gap = [PSCustomObject]@{
            Category = "Access Management"
            Issue = "Inactive Users with Roles"
            Details = "$($usersWithoutRecentSignIn.Count) enabled users with roles haven't signed in for 90+ days"
            Severity = "Medium"
            Recommendation = "Review and remove role assignments for inactive users"
            AffectedUsers = $inactiveUsers
            ComplianceFramework = "SOX, Access Management"
            RemediationSteps = @(
                "1. Contact users to verify continued need",
                "2. Remove role assignments for confirmed inactive users",
                "3. Implement regular inactive user reviews",
                "4. Consider conditional access policies"
            )
        }
        $mediumGaps += $gap
        $gaps += $gap
    }
    
    # === LEAST PRIVILEGE GAPS ===
    Write-Host "Analyzing Least Privilege Compliance..." -ForegroundColor Cyan
    
    # Check for excessive scope assignments
    $organizationWideRoles = $AuditResults | Where-Object { 
        $_.Scope -eq "Organization" -or $_.Scope -eq "/" -or [string]::IsNullOrEmpty($_.Scope)
    }
    $scopedRoles = $AuditResults | Where-Object { 
        ![string]::IsNullOrEmpty($_.Scope) -and $_.Scope -ne "Organization" -and $_.Scope -ne "/"
    }
    
    if ($organizationWideRoles.Count -gt ($scopedRoles.Count * 2) -and $organizationWideRoles.Count -gt 20) {
        $gap = [PSCustomObject]@{
            Category = "Least Privilege"
            Issue = "Excessive Organization-Wide Role Assignments"
            Details = "$($organizationWideRoles.Count) organization-wide roles vs $($scopedRoles.Count) scoped roles"
            Severity = "Low"
            Recommendation = "Consider implementing scoped role assignments where appropriate to limit access"
            AffectedUsers = "Multiple users"
            ComplianceFramework = "Principle of Least Privilege"
            RemediationSteps = @(
                "1. Review organization-wide role assignments",
                "2. Identify opportunities for scope-specific assignments",
                "3. Implement resource-scoped roles where possible",
                "4. Regular review of role scope requirements"
            )
        }
        $lowGaps += $gap
        $gaps += $gap
    }
    
    # Check for role sprawl
    $usersWithMultipleRoles = $AuditResults | Where-Object { 
        $_.UserPrincipalName -and $_.UserPrincipalName -ne "Unknown" -and $_.UserPrincipalName -ne "System Generated"
    } | Group-Object UserPrincipalName | Where-Object { $_.Count -gt 5 }
    
    if ($usersWithMultipleRoles.Count -gt 0) {
        $sprawlUsers = ($usersWithMultipleRoles | Select-Object -First 5 | ForEach-Object { "$($_.Name) ($($_.Count) roles)" }) -join "; "
        $gap = [PSCustomObject]@{
            Category = "Role Management"
            Issue = "Role Sprawl Detected"
            Details = "$($usersWithMultipleRoles.Count) users have more than 5 role assignments"
            Severity = "Medium"
            Recommendation = "Review users with excessive role assignments for consolidation opportunities"
            AffectedUsers = $sprawlUsers
            ComplianceFramework = "Least Privilege, Role-Based Access Control"
            RemediationSteps = @(
                "1. Analyze role combinations for each user",
                "2. Identify overlapping or redundant permissions",
                "3. Consolidate roles where possible",
                "4. Create custom roles for specific needs"
            )
        }
        $mediumGaps += $gap
        $gaps += $gap
    }
    
    # === INTUNE-SPECIFIC COMPLIANCE GAPS ===
    if ($IncludeIntuneGaps) {
        Write-Host "Analyzing Intune Compliance..." -ForegroundColor Cyan
        
        $intuneResults = $AuditResults | Where-Object { $_.Service -eq "Microsoft Intune" }
        if ($intuneResults.Count -gt 0) {
            # Check Intune Service Administrator count
            $intuneServiceAdmins = $intuneResults | Where-Object { $_.RoleName -eq "Intune Service Administrator" }
            if ($intuneServiceAdmins.Count -gt 3) {
                $gap = [PSCustomObject]@{
                    Category = "Device Management"
                    Issue = "Excessive Intune Service Administrators"
                    Details = "$($intuneServiceAdmins.Count) Intune Service Administrators (recommended: ≤3)"
                    Severity = "Medium"
                    Recommendation = "Use Intune RBAC roles for granular permissions instead of broad service administrator role"
                    AffectedUsers = ($intuneServiceAdmins | Select-Object -ExpandProperty UserPrincipalName) -join "; "
                    ComplianceFramework = "Device Security, NIST"
                    RemediationSteps = @(
                        "1. Review Intune administrative requirements",
                        "2. Implement Intune RBAC roles for specific functions",
                        "3. Remove unnecessary Service Administrator roles",
                        "4. Train admins on scoped permissions"
                    )
                }
                $mediumGaps += $gap
                $gaps += $gap
            }
            
            # Check RBAC vs Azure AD role usage
            $intuneRBACAssignments = $intuneResults | Where-Object { $_.RoleType -eq "IntuneRBAC" }
            $intuneAzureADAssignments = $intuneResults | Where-Object { $_.RoleType -eq "AzureAD" }
            
            if ($intuneAzureADAssignments.Count -gt $intuneRBACAssignments.Count -and $intuneResults.Count -gt 10) {
                $gap = [PSCustomObject]@{
                    Category = "Device Management"
                    Issue = "Underutilized Intune RBAC"
                    Details = "$($intuneAzureADAssignments.Count) Azure AD roles vs $($intuneRBACAssignments.Count) Intune RBAC roles"
                    Severity = "Low"
                    Recommendation = "Leverage Intune RBAC for more granular, scope-specific permissions"
                    AffectedUsers = "Intune administrators"
                    ComplianceFramework = "Least Privilege"
                    RemediationSteps = @(
                        "1. Map current Azure AD roles to Intune RBAC equivalents",
                        "2. Create custom Intune roles for specific needs",
                        "3. Migrate to Intune RBAC where appropriate",
                        "4. Implement scope-based assignments"
                    )
                }
                $lowGaps += $gap
                $gaps += $gap
            }
            
            # Check for policy ownership and management
            $intunePolicyOwners = $intuneResults | Where-Object { $_.RoleType -eq "PolicyOwner" }
            if ($intunePolicyOwners.Count -eq 0) {
                $gap = [PSCustomObject]@{
                    Category = "Device Management"
                    Issue = "No Policy Ownership Tracking"
                    Details = "No Intune policy ownership information found"
                    Severity = "Low"
                    Recommendation = "Implement policy ownership tracking and governance"
                    AffectedUsers = "Policy administrators"
                    ComplianceFramework = "Change Management"
                    RemediationSteps = @(
                        "1. Document policy ownership and responsibilities",
                        "2. Implement policy change approval process",
                        "3. Regular review of policy configurations",
                        "4. Track policy creation and modification"
                    )
                }
                $lowGaps += $gap
                $gaps += $gap
            }
        }
    }
    
    # === POWER PLATFORM COMPLIANCE GAPS ===
    if ($IncludePowerPlatformGaps) {
        Write-Host "Analyzing Power Platform Compliance..." -ForegroundColor Cyan
        
        $powerPlatformResults = $AuditResults | Where-Object { $_.Service -eq "Power Platform" }
        if ($powerPlatformResults.Count -gt 0) {
            # Check for service principals with Power Platform access
            $servicePrincipals = $powerPlatformResults | Where-Object { $_.PrincipalType -eq "ServicePrincipal" }
            if ($servicePrincipals.Count -gt 0) {
                $spNames = ($servicePrincipals | Select-Object -ExpandProperty DisplayName -Unique) -join "; "
                $gap = [PSCustomObject]@{
                    Category = "Application Security"
                    Issue = "Service Principals with Power Platform Access"
                    Details = "$($servicePrincipals.Count) service principals have Power Platform administrative access"
                    Severity = "Medium"
                    Recommendation = "Review and validate service principal access to Power Platform resources"
                    AffectedUsers = $spNames
                    ComplianceFramework = "Application Security"
                    RemediationSteps = @(
                        "1. Review each service principal's business justification",
                        "2. Validate minimum required permissions",
                        "3. Implement managed identities where possible",
                        "4. Regular audit of application permissions"
                    )
                }
                $mediumGaps += $gap
                $gaps += $gap
            }
            
            # Check Power Platform administrator count
            $powerPlatformAdmins = $powerPlatformResults | Where-Object { $_.RoleName -eq "Power Platform Administrator" }
            if ($powerPlatformAdmins.Count -gt 5) {
                $gap = [PSCustomObject]@{
                    Category = "Power Platform Governance"
                    Issue = "Excessive Power Platform Administrators"
                    Details = "$($powerPlatformAdmins.Count) Power Platform Administrators (consider environment-specific roles)"
                    Severity = "Medium"
                    Recommendation = "Use environment-specific admin roles instead of tenant-wide Power Platform Administrator"
                    AffectedUsers = ($powerPlatformAdmins | Select-Object -ExpandProperty UserPrincipalName) -join "; "
                    ComplianceFramework = "Least Privilege"
                    RemediationSteps = @(
                        "1. Review Power Platform administrative requirements",
                        "2. Implement environment-specific roles",
                        "3. Use DLP policies for governance",
                        "4. Regular review of platform usage"
                    )
                }
                $mediumGaps += $gap
                $gaps += $gap
            }
        }
    }
    
    # === MULTI-SERVICE AND CROSS-PLATFORM GAPS ===
    Write-Host "Analyzing Cross-Service Security..." -ForegroundColor Cyan
    
    # Check for high-risk cross-service combinations
    $usersWithCrossServiceRoles = $AuditResults | Where-Object { 
        $_.UserPrincipalName -and $_.UserPrincipalName -ne "Unknown" -and $_.UserPrincipalName -ne "System Generated"
    } | Group-Object UserPrincipalName | Where-Object {
        ($_.Group | Group-Object Service).Count -gt 1
    }
    
    $highRiskCombinations = $usersWithCrossServiceRoles | Where-Object {
        $userServices = ($_.Group | Group-Object Service).Name
        ($userServices -contains "Microsoft Purview" -and $userServices -contains "Exchange Online") -or
        ($userServices -contains "Azure AD/Entra ID" -and $userServices -contains "SharePoint Online" -and $userServices -contains "Exchange Online") -or
        ($userServices -contains "Microsoft Intune" -and $userServices -contains "Azure AD/Entra ID" -and $_.Count -gt 8)
    }
    
    if ($highRiskCombinations.Count -gt 0) {
        $riskUsers = ($highRiskCombinations | Select-Object -First 3 | ForEach-Object { 
            $services = ($_.Group | Group-Object Service).Name -join ","
            "$($_.Name) [$services]"
        }) -join "; "
        
        $gap = [PSCustomObject]@{
            Category = "Cross-Service Security"
            Issue = "High-Risk Cross-Service Role Combinations"
            Details = "$($highRiskCombinations.Count) users have high-risk combinations of roles across multiple services"
            Severity = "High"
            Recommendation = "Review and segregate duties for users with extensive cross-service privileges"
            AffectedUsers = $riskUsers
            ComplianceFramework = "Segregation of Duties, SOX"
            RemediationSteps = @(
                "1. Review business justification for cross-service access",
                "2. Implement segregation of duties where possible",
                "3. Use separate accounts for different administrative functions",
                "4. Enhanced monitoring for high-privilege users"
            )
        }
        $highGaps += $gap
        $gaps += $gap
    }
    
    # === COMPLIANCE REPORTING AND SUMMARY ===
    Write-Host ""
    Write-Host "=== COMPLIANCE GAP SUMMARY ===" -ForegroundColor Yellow
    
    if ($gaps.Count -eq 0) {
        Write-Host "✓ No significant compliance gaps identified!" -ForegroundColor Green
        return @()
    }
    
    Write-Host "Total Gaps Found: $($gaps.Count)" -ForegroundColor White
    Write-Host "  Critical: $($criticalGaps.Count)" -ForegroundColor Red
    Write-Host "  High: $($highGaps.Count)" -ForegroundColor Red
    Write-Host "  Medium: $($mediumGaps.Count)" -ForegroundColor Yellow
    Write-Host "  Low: $($lowGaps.Count)" -ForegroundColor Cyan
    
    if ($IncludeDetailedAnalysis) {
        Write-Host ""
        Write-Host "=== DETAILED GAP ANALYSIS ===" -ForegroundColor Cyan
        
        if ($criticalGaps.Count -gt 0) {
            Write-Host "CRITICAL GAPS:" -ForegroundColor Red
            foreach ($gap in $criticalGaps) {
                Write-Host "  ⚠️ $($gap.Issue): $($gap.Details)" -ForegroundColor White
                Write-Host "     Recommendation: $($gap.Recommendation)" -ForegroundColor Gray
            }
        }
        
        if ($highGaps.Count -gt 0) {
            Write-Host "HIGH PRIORITY GAPS:" -ForegroundColor Red
            foreach ($gap in $highGaps) {
                Write-Host "  ⚠️ $($gap.Issue): $($gap.Details)" -ForegroundColor White
                Write-Host "     Recommendation: $($gap.Recommendation)" -ForegroundColor Gray
            }
        }
        
        if ($mediumGaps.Count -gt 0) {
            Write-Host "MEDIUM PRIORITY GAPS:" -ForegroundColor Yellow
            foreach ($gap in $mediumGaps) {
                Write-Host "  • $($gap.Issue): $($gap.Details)" -ForegroundColor White
                Write-Host "    Recommendation: $($gap.Recommendation)" -ForegroundColor Gray
            }
        }
        
        if ($lowGaps.Count -gt 0) {
            Write-Host "LOW PRIORITY GAPS:" -ForegroundColor Cyan
            foreach ($gap in $lowGaps) {
                Write-Host "  • $($gap.Issue): $($gap.Details)" -ForegroundColor White
                Write-Host "    Recommendation: $($gap.Recommendation)" -ForegroundColor Gray
            }
        }
    }
    
    # Compliance framework mapping
    Write-Host ""
    Write-Host "=== COMPLIANCE FRAMEWORK IMPACT ===" -ForegroundColor Cyan
    $frameworkImpact = $gaps | ForEach-Object { $_.ComplianceFramework -split ", " } | 
                      Group-Object | Sort-Object Count -Descending
    
    foreach ($framework in $frameworkImpact) {
        Write-Host "  $($framework.Name): $($framework.Count) gaps" -ForegroundColor White
    }
    
    Write-Host ""
    Write-Host "Priority Remediation Recommendations:" -ForegroundColor Yellow
    $priorityRecommendations = @()
    $priorityRecommendations += $criticalGaps | ForEach-Object { $_.Recommendation }
    $priorityRecommendations += $highGaps | ForEach-Object { $_.Recommendation }
    
    $priorityRecommendations | Select-Object -Unique | ForEach-Object {
        Write-Host "• $_" -ForegroundColor White
    }
    
    return $gaps
}

function Export-M365ServiceAuditHtmlReport {
    param(
        [Parameter(Mandatory = $true)]
        [array]$AuditResults,
        
        [string]$OutputPath,
        [string]$OrganizationName = "Organization",
        [switch]$IncludeCharts,
        [switch]$IncludePIMAnalysis,
        [switch]$IncludeServiceSpecificAnalysis,
        [switch]$Quite
    )
    
    if ($AuditResults.Count -eq 0) {
        Write-Warning "No audit results provided"
        return
    }
    
    # Auto-detect service(s) and set appropriate defaults
    $services = $AuditResults | Group-Object Service | Sort-Object Count -Descending
    $primaryService = $services[0].Name
    $isMultiService = $services.Count -gt 1
    
    # Generate appropriate output path if not specified
    if (-not $OutputPath) {
        $serviceSlug = if ($isMultiService) { "MultiService" } else { $primaryService -replace "[^a-zA-Z0-9]", "" }
        $OutputPath = ".\M365_${serviceSlug}_Audit_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    }
    
    # Service-specific customizations
    $reportTitle = if ($isMultiService) {
        "Microsoft 365 Multi-Service Role Audit"
    } else {
        "$primaryService Role Audit Report"
    }
    
    $reportIcon = switch -Regex ($primaryService) {
        "Azure AD|Entra ID" { "🔷" }
        "SharePoint" { "📊" }
        "Exchange" { "📧" }
        "Teams" { "👥" }
        "Intune" { "📱" }
        "Purview|Compliance" { "🛡️" }
        "Defender" { "🔐" }
        "Power Platform" { "⚡" }
        default { "🔒" }
    }
    
    # Calculate service-specific statistics
    $totalAssignments = $AuditResults.Count
    $uniqueUsers = ($AuditResults | Where-Object { $_.UserPrincipalName -and $_.UserPrincipalName -ne "Unknown" } | 
                   Select-Object -Unique UserPrincipalName).Count
    
    # Service-specific analysis
    $serviceAnalysis = @{}
    
    switch -Regex ($primaryService) {
        "Azure AD|Entra ID" {
            $globalAdmins = $AuditResults | Where-Object { $_.RoleName -eq "Global Administrator" }
            $pimEligible = $AuditResults | Where-Object { $_.AssignmentType -like "*Eligible*" }
            $serviceAnalysis = @{
                GlobalAdmins = $globalAdmins.Count
                PIMEligible = $pimEligible.Count
                HasPIM = $pimEligible.Count -gt 0
                KeyMetric = "Global Administrator Count"
                KeyMetricValue = $globalAdmins.Count
                KeyMetricStatus = if ($globalAdmins.Count -le 5) { "✓ Good" } else { "⚠ Review" }
            }
        }
        "SharePoint" {
            $siteAdmins = $AuditResults | Where-Object { $_.RoleName -eq "Site Collection Administrator" }
            $sites = $AuditResults | Where-Object { $_.SiteTitle } | Select-Object -Unique SiteTitle
            $totalStorage = ($AuditResults | Where-Object { $_.StorageUsedMB } | Measure-Object StorageUsedMB -Sum).Sum
            $serviceAnalysis = @{
                SiteAdmins = $siteAdmins.Count
                UniqueSites = $sites.Count
                TotalStorageMB = $totalStorage
                KeyMetric = "Sites Audited"
                KeyMetricValue = $sites.Count
                KeyMetricStatus = "ℹ Info"
            }
        }
        "Exchange" {
            $roleGroups = $AuditResults | Where-Object { $_.AssignmentType -eq "Role Group Member" }
            $orgAdmins = $AuditResults | Where-Object { $_.RoleName -eq "Organization Management" }
            $serviceAnalysis = @{
                RoleGroupMembers = $roleGroups.Count
                OrgAdmins = $orgAdmins.Count
                KeyMetric = "Role Group Assignments"
                KeyMetricValue = $roleGroups.Count
                KeyMetricStatus = "ℹ Info"
            }
        }
        "Intune" {
            $serviceAdmins = $AuditResults | Where-Object { $_.RoleName -eq "Intune Service Administrator" }
            $rbacAssignments = $AuditResults | Where-Object { $_.RoleType -eq "IntuneRBAC" }
            $azureADAssignments = $AuditResults | Where-Object { $_.RoleType -eq "AzureAD" }
            $serviceAnalysis = @{
                ServiceAdmins = $serviceAdmins.Count
                RBACAssignments = $rbacAssignments.Count
                AzureADAssignments = $azureADAssignments.Count
                KeyMetric = "RBAC vs Azure AD Roles"
                KeyMetricValue = "$($rbacAssignments.Count) vs $($azureADAssignments.Count)"
                KeyMetricStatus = if ($rbacAssignments.Count -ge $azureADAssignments.Count) { "✓ Good" } else { "⚠ Consider RBAC" }
            }
        }
        "Teams" {
            $teamsAdmins = $AuditResults | Where-Object { $_.RoleName -like "*Teams*Administrator*" }
            $principalTypes = $AuditResults | Group-Object PrincipalType
            $serviceAnalysis = @{
                TeamsAdmins = $teamsAdmins.Count
                PrincipalTypes = $principalTypes.Count
                KeyMetric = "Teams Administrators"
                KeyMetricValue = $teamsAdmins.Count
                KeyMetricStatus = "ℹ Info"
            }
        }
        "Purview|Compliance" {
            $roleGroups = $AuditResults | Where-Object { $_.AssignmentType -eq "Role Group Member" }
            $policyOwners = $AuditResults | Where-Object { $_.AssignmentType -eq "Policy Owner" }
            $serviceAnalysis = @{
                RoleGroups = $roleGroups.Count
                PolicyOwners = $policyOwners.Count
                KeyMetric = "Compliance Role Groups"
                KeyMetricValue = $roleGroups.Count
                KeyMetricStatus = "ℹ Info"
            }
        }
        "Defender" {
            $securityAdmins = $AuditResults | Where-Object { $_.RoleName -like "*Security*Administrator*" }
            $securityReaders = $AuditResults | Where-Object { $_.RoleName -like "*Security*Reader*" }
            $serviceAnalysis = @{
                SecurityAdmins = $securityAdmins.Count
                SecurityReaders = $securityReaders.Count
                KeyMetric = "Security Administrators"
                KeyMetricValue = $securityAdmins.Count
                KeyMetricStatus = if ($securityAdmins.Count -le 3) { "✓ Good" } else { "⚠ Review" }
            }
        }
        "Power Platform" {
            $powerPlatformAdmins = $AuditResults | Where-Object { $_.RoleName -eq "Power Platform Administrator" }
            $principalTypes = $AuditResults | Group-Object PrincipalType
            $servicePrincipals = $AuditResults | Where-Object { $_.PrincipalType -eq "ServicePrincipal" }
            $serviceAnalysis = @{
                PowerPlatformAdmins = $powerPlatformAdmins.Count
                ServicePrincipals = $servicePrincipals.Count
                KeyMetric = "Power Platform Admins"
                KeyMetricValue = $powerPlatformAdmins.Count
                KeyMetricStatus = if ($powerPlatformAdmins.Count -le 5) { "✓ Good" } else { "⚠ Review" }
            }
        }
    }
    
    # Build HTML with enhanced styling and functionality
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$reportTitle - $OrganizationName</title>
    <style>
        :root {
            --primary-color: #0078d4;
            --secondary-color: #106ebe;
            --success-color: #107c10;
            --warning-color: #ff8c00;
            --danger-color: #d13438;
            --info-color: #00bcf2;
            --dark-color: #323130;
            --light-color: #f3f2f1;
        }
        
        * { box-sizing: border-box; }
        
        body { 
            font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, sans-serif; 
            margin: 0; 
            padding: 20px; 
            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
            line-height: 1.6;
        }
        
        .container { 
            max-width: 1400px; 
            margin: 0 auto; 
            background: white; 
            padding: 40px; 
            border-radius: 12px; 
            box-shadow: 0 8px 32px rgba(0,0,0,0.1);
            backdrop-filter: blur(10px);
        }
        
        .header { 
            text-align: center; 
            margin-bottom: 40px; 
            padding-bottom: 30px; 
            border-bottom: 3px solid var(--primary-color);
            background: linear-gradient(135deg, var(--primary-color), var(--secondary-color));
            color: white;
            margin: -40px -40px 40px -40px;
            padding: 40px;
            border-radius: 12px 12px 0 0;
        }
        
        .header h1 { 
            margin: 0; 
            font-size: 2.8em; 
            font-weight: 300;
            text-shadow: 0 2px 4px rgba(0,0,0,0.3);
        }
        
        .header p { 
            margin: 15px 0 0 0; 
            font-size: 1.2em; 
            opacity: 0.9;
        }
        
        .summary-grid { 
            display: grid; 
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); 
            gap: 25px; 
            margin-bottom: 40px; 
        }
        
        .summary-card { 
            background: linear-gradient(135deg, var(--primary-color), var(--secondary-color)); 
            color: white; 
            padding: 30px; 
            border-radius: 12px; 
            text-align: center;
            position: relative;
            overflow: hidden;
            transition: transform 0.3s ease;
        }
        
        .summary-card:hover {
            transform: translateY(-5px);
        }
        
        .summary-card::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: linear-gradient(45deg, transparent 30%, rgba(255,255,255,0.1) 50%, transparent 70%);
            transform: translateX(-100%);
            transition: transform 0.6s;
        }
        
        .summary-card:hover::before {
            transform: translateX(100%);
        }
        
        .summary-card h3 { 
            margin: 0 0 15px 0; 
            font-size: 1.3em; 
            font-weight: 400;
        }
        
        .summary-card .number { 
            font-size: 3em; 
            font-weight: bold; 
            margin: 15px 0;
            text-shadow: 0 2px 4px rgba(0,0,0,0.3);
        }
        
        .summary-card .subtitle {
            font-size: 0.9em;
            opacity: 0.8;
        }
        
        .section { 
            margin-bottom: 50px; 
            position: relative;
        }
        
        .section h2 { 
            color: var(--primary-color); 
            border-bottom: 3px solid var(--primary-color); 
            padding-bottom: 15px; 
            font-size: 2em;
            font-weight: 300;
            margin-bottom: 25px;
            position: relative;
        }
        
        .section h2::after {
            content: '';
            position: absolute;
            bottom: -3px;
            left: 0;
            width: 60px;
            height: 3px;
            background: var(--warning-color);
        }
        
        .security-alerts { 
            background: linear-gradient(135deg, #fff3cd, #ffeaa7); 
            border: 1px solid var(--warning-color); 
            border-left: 5px solid var(--warning-color);
            border-radius: 8px; 
            padding: 25px; 
            margin: 25px 0; 
        }
        
        .security-alerts h3 {
            margin-top: 0;
            color: #856404;
            font-size: 1.2em;
        }
        
        .alert-item { 
            margin: 15px 0; 
            padding: 10px;
            border-radius: 5px;
            background: rgba(255,255,255,0.7);
        }
        
        .alert-warning { 
            color: #856404; 
            border-left: 4px solid var(--warning-color);
            padding-left: 15px;
        }
        
        .alert-success {
            color: var(--success-color);
            border-left: 4px solid var(--success-color);
            padding-left: 15px;
            background: rgba(16, 124, 16, 0.1);
        }
        
        table { 
            width: 100%; 
            border-collapse: collapse; 
            margin: 25px 0; 
            background: white;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 4px 16px rgba(0,0,0,0.1);
        }
        
        th, td { 
            padding: 15px 12px; 
            text-align: left; 
            border-bottom: 1px solid #e1dfdd; 
        }
        
        th { 
            background: linear-gradient(135deg, var(--primary-color), var(--secondary-color)); 
            color: white; 
            font-weight: 600;
            text-transform: uppercase;
            font-size: 0.9em;
            letter-spacing: 0.5px;
        }
        
        tr:hover { 
            background-color: #f9f9f9; 
            transition: background-color 0.2s ease;
        }
        
        tr:nth-child(even) {
            background-color: #fafafa;
        }
        
        .service-badge {
            padding: 4px 12px;
            border-radius: 20px;
            color: white;
            font-size: 0.85em;
            font-weight: 600;
            text-align: center;
            display: inline-block;
            min-width: 120px;
        }
        
        .service-azure { background: linear-gradient(135deg, #0078d4, #106ebe); }
        .service-sharepoint { background: linear-gradient(135deg, #0b6623, #0e7629); }
        .service-exchange { background: linear-gradient(135deg, #d13438, #b02a37); }
        .service-teams { background: linear-gradient(135deg, #464775, #5b5d8a); }
        .service-purview { background: linear-gradient(135deg, #8b4789, #9e5a9c); }
        .service-intune { background: linear-gradient(135deg, #00bcf2, #0078d4); }
        .service-defender { background: linear-gradient(135deg, #ff8c00, #e67e22); }
        .service-powerplatform { background: linear-gradient(135deg, #742774, #8b4789); }
        
        .tag {
            display: inline-block;
            background: var(--primary-color);
            color: white;
            padding: 3px 8px;
            border-radius: 12px;
            font-size: 0.8em;
            margin: 2px;
            font-weight: 500;
        }
        
        .tag.pim { background: var(--success-color); }
        .tag.permanent { background: var(--warning-color); }
        .tag.disabled { background: var(--danger-color); }
        
        .stat-number {
            font-size: 1.5em;
            font-weight: bold;
            color: var(--primary-color);
        }
        
        .user-details {
            display: none;
            margin-top: 15px;
            padding: 15px;
            background: #f9f9f9;
            border-radius: 6px;
            border-left: 4px solid var(--primary-color);
            max-height: 250px;
            overflow-y: auto;
        }
        
        .user-details.show {
            display: block;
            animation: slideDown 0.3s ease-out;
        }
        
        @keyframes slideDown {
            from { opacity: 0; max-height: 0; }
            to { opacity: 1; max-height: 250px; }
        }
        
        .role-badge-container {
            max-width: 300px;
            line-height: 1.4;
        }
        
        .action-button {
            background: var(--primary-color);
            color: white;
            border: none;
            padding: 6px 12px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 0.85em;
            transition: all 0.2s ease;
        }
        
        .action-button:hover {
            background: var(--secondary-color);
            transform: translateY(-1px);
        }
        
        .action-button.active {
            background: var(--warning-color);
        }
        
        .footer { 
            margin-top: 60px; 
            padding-top: 30px; 
            border-top: 2px solid #ddd; 
            text-align: center; 
            color: #666; 
            background: #f9f9f9;
            margin-left: -40px;
            margin-right: -40px;
            padding-left: 40px;
            padding-right: 40px;
            border-radius: 0 0 12px 12px;
        }
        
        @media (max-width: 768px) {
            .container { padding: 20px; }
            .summary-grid { grid-template-columns: 1fr; }
            .header h1 { font-size: 2em; }
            table { font-size: 0.9em; }
            th, td { padding: 10px 8px; }
            .role-badge-container { max-width: 200px; }
        }
    </style>
    <script>
        function toggleUserDetails(userId) {
            const details = document.getElementById('userDetails_' + userId);
            const button = event.target;
            
            if (details.classList.contains('show')) {
                details.classList.remove('show');
                setTimeout(() => { details.style.display = 'none'; }, 300);
                button.textContent = 'View All Roles';
                button.classList.remove('active');
            } else {
                details.style.display = 'block';
                setTimeout(() => { details.classList.add('show'); }, 10);
                button.textContent = 'Hide Roles';
                button.classList.add('active');
            }
        }
        
        function scrollToTop() {
            window.scrollTo({ top: 0, behavior: 'smooth' });
        }
        
        document.addEventListener('DOMContentLoaded', function() {
            // Add scroll functionality if needed
            window.addEventListener('scroll', function() {
                // Could add scroll-based interactions here
            });
        });
    </script>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>$reportIcon $reportTitle</h1>
            <p>$OrganizationName | Generated on $(Get-Date -Format 'MMMM dd, yyyy at HH:mm')</p>
        </div>
        
        <!-- Service-Specific Dashboard -->
        <div class="summary-grid">
            <div class="summary-card">
                <h3>Total Assignments</h3>
                <div class="number">$totalAssignments</div>
                <div class="subtitle">Role assignments found</div>
            </div>
            <div class="summary-card">
                <h3>Unique Users</h3>
                <div class="number">$uniqueUsers</div>
                <div class="subtitle">With role assignments</div>
            </div>
            <div class="summary-card">
                <h3>$($serviceAnalysis.KeyMetric)</h3>
                <div class="number">$($serviceAnalysis.KeyMetricValue)</div>
                <div class="subtitle">$($serviceAnalysis.KeyMetricStatus)</div>
            </div>
"@
    
    # Add service-specific cards
    if ($primaryService -match "Azure AD|Entra ID" -and $serviceAnalysis.HasPIM) {
        $html += @"
            <div class="summary-card">
                <h3>PIM Eligible</h3>
                <div class="number">$($serviceAnalysis.PIMEligible)</div>
                <div class="subtitle">Require activation</div>
            </div>
"@
    } elseif ($primaryService -eq "SharePoint Online" -and $serviceAnalysis.TotalStorageMB -gt 0) {
        $storageGB = [math]::Round($serviceAnalysis.TotalStorageMB / 1024, 1)
        $html += @"
            <div class="summary-card">
                <h3>Storage Used</h3>
                <div class="number">$storageGB</div>
                <div class="subtitle">GB across sites</div>
            </div>
"@
    } elseif ($primaryService -eq "Microsoft Intune") {
        $html += @"
            <div class="summary-card">
                <h3>RBAC Usage</h3>
                <div class="number">$($serviceAnalysis.RBACAssignments)</div>
                <div class="subtitle">Intune RBAC roles</div>
            </div>
"@
    }
    
    $html += @"
        </div>
        
        <!-- Service-Specific Analysis Section -->
        <div class="section">
            <h2>📊 $primaryService Analysis</h2>
"@
    
    # Add service-specific insights
    switch -Regex ($primaryService) {
        "Azure AD|Entra ID" {
            if ($serviceAnalysis.GlobalAdmins -gt 5) {
                $html += @"
            <div class="security-alerts">
                <h3>Security Recommendations</h3>
                <div class="alert-item alert-warning">⚠️ Consider reducing Global Administrator count from $($serviceAnalysis.GlobalAdmins) to 5 or fewer</div>
"@
                if ($serviceAnalysis.PIMEligible -eq 0) {
                    $html += @"
                <div class="alert-item alert-warning">⚠️ Consider implementing PIM for privileged role assignments</div>
"@
                }
                $html += "</div>"
            } else {
                $html += @"
            <div class="security-alerts" style="background: linear-gradient(135deg, #e8f5e8, #c8e6c9); border-color: var(--success-color);">
                <h3 style="color: var(--success-color);">Security Status</h3>
                <div class="alert-item alert-success">✓ Global Administrator count is within recommended limits</div>
"@
                if ($serviceAnalysis.PIMEligible -gt 0) {
                    $html += @"
                <div class="alert-item alert-success">✓ PIM eligible assignments detected</div>
"@
                }
                $html += "</div>"
            }
        }
        "Microsoft Intune" {
            $html += @"
            <div class="security-alerts">
                <h3>Intune Configuration Analysis</h3>
                <div class="alert-item">📊 Service Administrators: $($serviceAnalysis.ServiceAdmins)</div>
                <div class="alert-item">🔧 RBAC Assignments: $($serviceAnalysis.RBACAssignments)</div>
                <div class="alert-item">🔷 Azure AD Assignments: $($serviceAnalysis.AzureADAssignments)</div>
"@
            if ($serviceAnalysis.RBACAssignments -lt $serviceAnalysis.AzureADAssignments) {
                $html += @"
                <div class="alert-item alert-warning">⚠️ Consider using more Intune RBAC roles for granular permissions</div>
"@
            } else {
                $html += @"
                <div class="alert-item alert-success">✓ Good use of Intune RBAC roles</div>
"@
            }
            $html += "</div>"
        }
        # Add other service-specific analysis blocks as needed...
    }
    
    $html += @"
        </div>
        
        <!-- Role Distribution Table -->
        <div class="section">
            <h2>👑 Role Distribution</h2>
            <table>
                <tr><th>Role Name</th><th>Users Assigned</th><th>Assignment Type</th><th>Details</th></tr>
"@
    
    # Group by role and show distribution
    $roleDistribution = $AuditResults | Group-Object RoleName | Sort-Object Count -Descending
    foreach ($role in $roleDistribution) {
        $assignmentTypes = ($role.Group | Group-Object AssignmentType | ForEach-Object { "$($_.Name): $($_.Count)" }) -join "; "
        $roleDetails = switch -Regex ($role.Name) {
            "Global Administrator" { "Highest privilege - monitor closely" }
            ".*Administrator.*" { "Administrative role" }
            ".*Reader.*" { "Read-only access" }
            default { "Standard role" }
        }
        
        $html += @"
<tr>
    <td><strong>$($role.Name)</strong></td>
    <td>$($role.Count)</td>
    <td><small>$assignmentTypes</small></td>
    <td><small>$roleDetails</small></td>
</tr>
"@
    }
    
    $html += @"
            </table>
        </div>
        
        <!-- Enhanced User Analysis -->
        <div class="section">
            <h2>👥 User Analysis</h2>
            <table>
                <tr><th>User</th><th>Role Count</th><th>Primary Roles</th><th>Status</th><th>Last Sign-In</th><th>Actions</th></tr>
"@
    
    # Enhanced User Analysis with expandable details
    $userAnalysis = $AuditResults | Where-Object { $_.UserPrincipalName -and $_.UserPrincipalName -ne "Unknown" } |
                   Group-Object UserPrincipalName | Sort-Object Count -Descending | Select-Object -First 20
    
    foreach ($user in $userAnalysis) {
        $userInfo = $user.Group[0]
        
        # Get top 3 most important roles for display
        $allRoles = $user.Group | Select-Object -ExpandProperty RoleName | Sort-Object
        $topRoles = @()
        
        # Prioritize certain critical roles for display
        $criticalRoles = $allRoles | Where-Object { 
            $_ -match "Global Administrator|Security Administrator|Exchange Administrator|SharePoint Administrator|Intune Service Administrator" 
        }
        
        if ($criticalRoles.Count -gt 0) {
            $topRoles += $criticalRoles | Select-Object -First 2
        }
        
        # Add other roles up to 3 total
        $remainingSlots = 3 - $topRoles.Count
        if ($remainingSlots -gt 0) {
            $otherRoles = $allRoles | Where-Object { $_ -notin $topRoles } | Select-Object -First $remainingSlots
            $topRoles += $otherRoles
        }
        
        # Format role display
        $roleDisplay = ""
        foreach ($role in $topRoles) {
            $roleClass = switch -Regex ($role) {
                "Global Administrator" { "tag disabled" }
                ".*Security.*Administrator.*" { "tag permanent" }
                ".*Administrator.*" { "tag" }
                default { "tag pim" }
            }
            $roleDisplay += "<span class='$roleClass'>$role</span> "
        }
        
        # Add "more roles" indicator if user has additional roles
        $additionalRoles = $allRoles.Count - $topRoles.Count
        if ($additionalRoles -gt 0) {
            $roleDisplay += "<br><small style='color: #666; font-style: italic;'>+$additionalRoles more role$(if($additionalRoles -gt 1){'s'})</small>"
        }
        
        $statusColor = if ($userInfo.UserEnabled -eq $false) { "style='color: red; font-weight: bold;'" } else { "" }
        $status = if ($userInfo.UserEnabled -eq $false) { "DISABLED" } else { "Active" }
        $lastSignIn = if ($userInfo.LastSignIn) { 
            ([DateTime]$userInfo.LastSignIn).ToString("yyyy-MM-dd") 
        } else { 
            "Unknown" 
        }
        
        # Create expandable details
        $userId = $user.Name -replace '[^a-zA-Z0-9]', ''
        
        # Build detailed role list grouped by service
        $rolesByService = $user.Group | Group-Object Service | Sort-Object Name
        $allRolesList = ""
        foreach ($serviceGroup in $rolesByService) {
            $serviceClass = switch ($serviceGroup.Name) {
                "Azure AD/Entra ID" { "service-azure" }
                "SharePoint Online" { "service-sharepoint" }
                "Exchange Online" { "service-exchange" }
                "Microsoft Teams" { "service-teams" }
                "Microsoft Purview" { "service-purview" }
                "Microsoft Intune" { "service-intune" }
                "Microsoft Defender" { "service-defender" }
                "Power Platform" { "service-powerplatform" }
                default { "service-azure" }
            }
            
            $allRolesList += "<div style='margin-bottom: 12px;'>"
            $allRolesList += "<span class='service-badge $serviceClass' style='font-size: 0.75em; margin-bottom: 5px; display: inline-block;'>$($serviceGroup.Name)</span><br>"
            
            $serviceRoles = $serviceGroup.Group | Sort-Object RoleName
            foreach ($roleAssignment in $serviceRoles) {
                $assignmentBadge = if ($roleAssignment.AssignmentType -like "*Eligible*") {
                    "<span class='tag pim' style='font-size: 0.7em; margin-left: 8px;'>PIM Eligible</span>"
                } elseif ($roleAssignment.AssignmentType -like "*PIM*") {
                    "<span class='tag' style='font-size: 0.7em; margin-left: 8px;'>PIM Active</span>"
                } else {
                    "<span class='tag permanent' style='font-size: 0.7em; margin-left: 8px;'>Permanent</span>"
                }
                
                $allRolesList += "<div style='margin: 4px 0 4px 16px; font-size: 0.9em;'>• $($roleAssignment.RoleName) $assignmentBadge</div>"
            }
            $allRolesList += "</div>"
        }
        
        $html += @"
<tr>
    <td $statusColor>
        <strong>$($user.Name)</strong><br>
        <small style='color: #666;'>$($userInfo.DisplayName)</small>
    </td>
    <td style='text-align: center;'>
        <span class='stat-number'>$($user.Count)</span>
    </td>
    <td class='role-badge-container'>
        $roleDisplay
    </td>
    <td><span class='tag $(if($status -eq "DISABLED") {"disabled"} else {"pim"})' $statusColor>$status</span></td>
    <td>$lastSignIn</td>
    <td>
        <button onclick="toggleUserDetails('$userId')" class="action-button">
            View All Roles
        </button>
        <div id="userDetails_$userId" class="user-details">
            $allRolesList
        </div>
    </td>
</tr>
"@
    }
    
    $html += @"
            </table>
        </div>
        
        <!-- Recommendations -->
        <div class="section">
            <h2>💡 Service-Specific Recommendations</h2>
            <div class="security-alerts" style="background: linear-gradient(135deg, #e8f5e8, #c8e6c9); border-color: var(--success-color);">
                <h3 style="color: var(--success-color);">Best Practices for $primaryService</h3>
"@
    
    # Add service-specific recommendations
    switch -Regex ($primaryService) {
        "Azure AD|Entra ID" {
            $html += @"
                <div class="alert-item">• Implement regular access reviews for privileged roles</div>
                <div class="alert-item">• Enable PIM for eligible assignments to reduce standing privileges</div>
                <div class="alert-item">• Monitor Global Administrator usage and reduce count if possible</div>
                <div class="alert-item">• Implement conditional access policies for administrative accounts</div>
"@
        }
        "SharePoint" {
            $html += @"
                <div class="alert-item">• Regularly review site collection administrator assignments</div>
                <div class="alert-item">• Implement SharePoint governance policies</div>
                <div class="alert-item">• Monitor site creation and ownership</div>
                <div class="alert-item">• Use SharePoint groups for easier permission management</div>
"@
        }
        "Exchange" {
            $html += @"
                <div class="alert-item">• Review Exchange role group memberships regularly</div>
                <div class="alert-item">• Implement management scopes for limited administrative access</div>
                <div class="alert-item">• Monitor mailbox permissions and delegation</div>
                <div class="alert-item">• Use Exchange RBAC for granular permissions</div>
"@
        }
        "Microsoft Intune" {
            $html += @"
                <div class="alert-item">• Use Intune RBAC roles instead of broad Azure AD roles</div>
                <div class="alert-item">• Implement scope-based assignments for administrative roles</div>
                <div class="alert-item">• Regularly review device policy ownership</div>
                <div class="alert-item">• Monitor app and device enrollment permissions</div>
"@
        }
        "Teams" {
            $html += @"
                <div class="alert-item">• Review Teams administrative role assignments regularly</div>
                <div class="alert-item">• Monitor Teams policy changes and ownership</div>
                <div class="alert-item">• Implement governance for Teams creation and management</div>
                <div class="alert-item">• Use Teams-specific roles for granular permissions</div>
"@
        }
        "Purview|Compliance" {
            $html += @"
                <div class="alert-item">• Review compliance role group memberships regularly</div>
                <div class="alert-item">• Monitor eDiscovery case access and permissions</div>
                <div class="alert-item">• Implement data classification governance</div>
                <div class="alert-item">• Regular audit of DLP and retention policies</div>
"@
        }
        "Defender" {
            $html += @"
                <div class="alert-item">• Review security administrator assignments regularly</div>
                <div class="alert-item">• Monitor security policy changes and configurations</div>
                <div class="alert-item">• Implement least privilege for security operations</div>
                <div class="alert-item">• Regular review of incident response permissions</div>
"@
        }
        "Power Platform" {
            $html += @"
                <div class="alert-item">• Review Power Platform administrator assignments regularly</div>
                <div class="alert-item">• Monitor environment creation and DLP policies</div>
                <div class="alert-item">• Implement governance for app and flow creation</div>
                <div class="alert-item">• Regular audit of connector and data source access</div>
"@
        }
    }
    
    $html += @"
            </div>
        </div>
        
        <div class="footer">
            <p><strong>$primaryService Role Audit Report</strong></p>
            <p>Generated by M365 Role Audit PowerShell Module | $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
            <p>Total assignments analyzed: $totalAssignments | Authentication: Certificate-based</p>
            <p style="margin-top: 15px; font-size: 0.9em; color: #888;">
                Enhanced user analysis with expandable role details for improved readability.<br>
                Click "View All Roles" to see complete role assignments grouped by service.
            </p>
        </div>
    </div>
</body>
</html>
"@
    
    try {
        $html | Out-File -FilePath $OutputPath -Encoding UTF8
        Write-Host "✓ Enhanced $primaryService HTML report generated: $OutputPath" -ForegroundColor Green

        if (-not $Quite) {
        
            $fileSize = [math]::Round((Get-Item $OutputPath).Length / 1KB, 2)
            Write-Host "Report size: $fileSize KB" -ForegroundColor Gray
            Write-Host "Service: $primaryService" -ForegroundColor Gray
            Write-Host "Total assignments: $totalAssignments" -ForegroundColor Gray
            Write-Host "Enhanced features:" -ForegroundColor Cyan
            Write-Host "  • Expandable user role details" -ForegroundColor White
            Write-Host "  • Prioritized role display" -ForegroundColor White
            Write-Host "  • Service-grouped role listings" -ForegroundColor White
            Write-Host "  • Interactive role exploration" -ForegroundColor White
            
            # Open report if on Windows
            if ($IsWindows -ne $false -and (Test-Path $OutputPath)) {
                $openReport = Read-Host "Open enhanced report in browser? (y/N)"
                if ($openReport -eq "y" -or $openReport -eq "Y") {
                    Start-Process $OutputPath
                }
            }
        }

        return $OutputPath
    }
    catch {
        Write-Error "Failed to generate enhanced service-specific HTML report: $($_.Exception.Message)"
        return $null
    }
}

function Export-M365ServiceAuditJsonReport {
    param(
        [Parameter(Mandatory = $true)]
        [array]$AuditResults,
        
        [string]$OutputPath,
        [string]$OrganizationName = "Organization",
        [switch]$IncludeDetailedAnalysis
    )
    
    if ($AuditResults.Count -eq 0) {
        Write-Warning "No audit results provided"
        return
    }
    
    # Auto-detect service and generate appropriate filename
    $primaryService = ($AuditResults | Group-Object Service | Sort-Object Count -Descending)[0].Name
    if (-not $OutputPath) {
        $serviceSlug = $primaryService -replace "[^a-zA-Z0-9]", ""
        $OutputPath = ".\M365_${serviceSlug}_Audit_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    }
    
    # Build service-specific JSON report
    $report = @{
        metadata = @{
            organizationName = $OrganizationName
            primaryService = $primaryService
            generatedDate = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
            reportType = "Service-Specific"
            totalAssignments = $AuditResults.Count
            uniqueUsers = ($AuditResults | Where-Object { $_.UserPrincipalName } | Select-Object -Unique UserPrincipalName).Count
            authenticationUsed = ($AuditResults | Group-Object AuthenticationType | ForEach-Object { @{ method = $_.Name; count = $_.Count } })
        }
        
        serviceAnalysis = @{
            serviceName = $primaryService
            totalAssignments = $AuditResults.Count
            roleDistribution = @($AuditResults | Group-Object RoleName | Sort-Object Count -Descending | ForEach-Object {
                @{
                    roleName = $_.Name
                    userCount = $_.Count
                    assignmentTypes = @($_.Group | Group-Object AssignmentType | ForEach-Object { 
                        @{ type = $_.Name; count = $_.Count } 
                    })
                }
            })
            userAnalysis = @($AuditResults | Where-Object { $_.UserPrincipalName } | Group-Object UserPrincipalName | 
                           Sort-Object Count -Descending | Select-Object -First 20 | ForEach-Object {
                $userInfo = $_.Group[0]
                @{
                    userPrincipalName = $_.Name
                    displayName = $userInfo.DisplayName
                    roleCount = $_.Count
                    roles = @($_.Group | Select-Object -ExpandProperty RoleName)
                    isEnabled = $userInfo.UserEnabled
                    lastSignIn = $userInfo.LastSignIn
                }
            })
        }
        
        assignments = @($AuditResults | ForEach-Object {
            $assignment = @{
                service = $_.Service
                userPrincipalName = $_.UserPrincipalName
                displayName = $_.DisplayName
                roleName = $_.RoleName
                assignmentType = $_.AssignmentType
                assignedDateTime = $_.AssignedDateTime
                userEnabled = $_.UserEnabled
                lastSignIn = $_.LastSignIn
                scope = $_.Scope
                authenticationType = $_.AuthenticationType
            }
            
            # Add service-specific properties
            if ($_.PrincipalType) { $assignment.principalType = $_.PrincipalType }
            if ($_.SiteTitle) { $assignment.siteTitle = $_.SiteTitle }
            if ($_.RoleType) { $assignment.roleType = $_.RoleType }
            if ($_.PolicyType) { $assignment.policyType = $_.PolicyType }
            # Add other service-specific fields as needed...
            
            $assignment
        })
    }
    
    # Add service-specific analysis if requested
    if ($IncludeDetailedAnalysis) {
        switch -Regex ($primaryService) {
            "Azure AD|Entra ID" {
                $pimEligible = $AuditResults | Where-Object { $_.AssignmentType -like "*Eligible*" }
                $globalAdmins = $AuditResults | Where-Object { $_.RoleName -eq "Global Administrator" }
                
                $report.serviceSpecificAnalysis = @{
                    azureADAnalysis = @{
                        globalAdministrators = $globalAdmins.Count
                        pimEligibleAssignments = $pimEligible.Count
                        privilegedRoles = ($AuditResults | Where-Object { $_.RoleName -match "Administrator" }).Count
                        complianceStatus = @{
                            globalAdminLimit = @{
                                compliant = $globalAdmins.Count -le 5
                                current = $globalAdmins.Count
                                recommended = 5
                            }
                            pimImplementation = @{
                                implemented = $pimEligible.Count -gt 0
                                eligibleCount = $pimEligible.Count
                            }
                        }
                    }
                }
            }
            "Microsoft Intune" {
                $serviceAdmins = $AuditResults | Where-Object { $_.RoleName -eq "Intune Service Administrator" }
                $rbacRoles = $AuditResults | Where-Object { $_.RoleType -eq "IntuneRBAC" }
                $azureADRoles = $AuditResults | Where-Object { $_.RoleType -eq "AzureAD" }
                
                $report.serviceSpecificAnalysis = @{
                    intuneAnalysis = @{
                        serviceAdministrators = $serviceAdmins.Count
                        rbacAssignments = $rbacRoles.Count
                        azureADAssignments = $azureADRoles.Count
                        builtInRoles = ($AuditResults | Where-Object { $_.IsBuiltIn -eq $true }).Count
                        customRoles = ($AuditResults | Where-Object { $_.IsBuiltIn -eq $false }).Count
                        policyOwnership = ($AuditResults | Where-Object { $_.RoleType -eq "PolicyOwner" }).Count
                    }
                }
            }
            # Add analysis for other services...
        }
    }
    
    try {
        $jsonOutput = $report | ConvertTo-Json -Depth 10
        $jsonOutput | Out-File -FilePath $OutputPath -Encoding UTF8
        
        Write-Host "✓ $primaryService JSON report generated: $OutputPath" -ForegroundColor Green
        
        $fileSize = [math]::Round((Get-Item $OutputPath).Length / 1KB, 2)
        Write-Host "File size: $fileSize KB" -ForegroundColor Gray
        Write-Host "Service: $primaryService" -ForegroundColor Gray
        Write-Host "Total assignments: $($AuditResults.Count)" -ForegroundColor Gray
        
        return $OutputPath
    }
    catch {
        Write-Error "Failed to generate service-specific JSON report: $($_.Exception.Message)"
        return $null
    }
}

# Convenience function that automatically detects service and calls appropriate export
function Export-M365ServiceAuditReport {
    param(
        [Parameter(Mandatory = $true)]
        [array]$AuditResults,
        
        [ValidateSet("HTML", "JSON", "Both")]
        [string]$Format = "HTML",
        
        [string]$OutputPath,
        [string]$OrganizationName = "Organization",
        [switch]$IncludeCharts,
        [switch]$IncludeDetailedAnalysis
    )
    
    if ($AuditResults.Count -eq 0) {
        Write-Warning "No audit results provided"
        return
    }
    
    $results = @()
    
    if ($Format -eq "HTML" -or $Format -eq "Both") {
        $htmlPath = Export-M365ServiceAuditHtmlReport -AuditResults $AuditResults -OutputPath $OutputPath -OrganizationName $OrganizationName -IncludeCharts:$IncludeCharts
        if ($htmlPath) { $results += $htmlPath }
    }
    
    if ($Format -eq "JSON" -or $Format -eq "Both") {
        $jsonPath = if ($OutputPath -and $Format -eq "Both") {
            $OutputPath -replace "\.html$", ".json"
        } else { 
            $OutputPath 
        }
        $jsonPath = Export-M365ServiceAuditJsonReport -AuditResults $AuditResults -OutputPath $jsonPath -OrganizationName $OrganizationName -IncludeDetailedAnalysis:$IncludeDetailedAnalysis
        if ($jsonPath) { $results += $jsonPath }
    }
    
    return $results
}

function Export-M365AuditExcelReport {
    param(
        [Parameter(Mandatory = $true)]
        [array]$AuditResults,
        
        [string]$OutputPath = ".\M365_Audit_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx",
        [string]$OrganizationName = "Organization",
        [switch]$AutoOpen
    )
    
    Write-Host "Generating Excel audit report..." -ForegroundColor Cyan
    
    if ($AuditResults.Count -eq 0) {
        Write-Warning "No audit results provided"
        return
    }
    
    # Check if ImportExcel module is available
    if (-not (Get-Module -ListAvailable -Name "ImportExcel")) {
        Write-Host "ImportExcel module not found. Installing..." -ForegroundColor Yellow
        try {
            Install-Module -Name ImportExcel -Force -AllowClobber -Scope CurrentUser
            Write-Host "✓ ImportExcel module installed successfully" -ForegroundColor Green
        }
        catch {
            Write-Error "Failed to install ImportExcel module: $($_.Exception.Message)"
            return
        }
    }
    
    Import-Module ImportExcel -Force
    
    try {
        # Remove existing file if it exists
        if (Test-Path $OutputPath) {
            Remove-Item $OutputPath -Force
        }
        
        # Calculate summary statistics
        $totalAssignments = $AuditResults.Count
        $uniqueUsers = ($AuditResults | Where-Object { $_.UserPrincipalName } | Select-Object -Unique UserPrincipalName).Count
        $services = $AuditResults | Group-Object Service | Sort-Object Count -Descending
        $topRoles = $AuditResults | Group-Object RoleName | Sort-Object Count -Descending | Select-Object -First 10
        $usersWithMostRoles = $AuditResults | Group-Object UserPrincipalName | Sort-Object Count -Descending | Select-Object -First 10
        $globalAdmins = $AuditResults | Where-Object { $_.RoleName -eq "Global Administrator" }
        $disabledUsers = $AuditResults | Where-Object { $_.UserEnabled -eq $false }
        $pimEligible = $AuditResults | Where-Object { $_.AssignmentType -like "*Eligible*" }
        $pimActive = $AuditResults | Where-Object { $_.AssignmentType -like "*Active (PIM*" }
        $authTypes = $AuditResults | Group-Object AuthenticationType
        
        Write-Host "Creating Summary sheet..." -ForegroundColor Cyan
        
        # Create Summary Sheet
        $summaryData = @()
        
        # Organization Overview
        $summaryData += [PSCustomObject]@{
            Category = "Organization Overview"
            Metric = "Organization Name"
            Value = $OrganizationName
            Details = ""
            Status = "Info"
        }
        $summaryData += [PSCustomObject]@{
            Category = "Organization Overview"
            Metric = "Report Generated"
            Value = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            Details = ""
            Status = "Info"
        }
        $summaryData += [PSCustomObject]@{
            Category = "Organization Overview"
            Metric = "Total Role Assignments"
            Value = $totalAssignments
            Details = "Across all audited services"
            Status = "Info"
        }
        $summaryData += [PSCustomObject]@{
            Category = "Organization Overview"
            Metric = "Unique Users with Roles"
            Value = $uniqueUsers
            Details = "Distinct users with role assignments"
            Status = "Info"
        }
        $summaryData += [PSCustomObject]@{
            Category = "Organization Overview"
            Metric = "Services Audited"
            Value = $services.Count
            Details = ($services.Name -join ", ")
            Status = "Info"
        }
        
        # Security Metrics
        $summaryData += [PSCustomObject]@{
            Category = "Security Metrics"
            Metric = "Global Administrators"
            Value = $globalAdmins.Count
            Details = if ($globalAdmins.Count -gt 5) { "⚠️ Exceeds recommended limit of 5" } else { "✓ Within recommended limit" }
            Status = if ($globalAdmins.Count -gt 5) { "Warning" } else { "Good" }
        }
        $summaryData += [PSCustomObject]@{
            Category = "Security Metrics"
            Metric = "Disabled Users with Roles"
            Value = $disabledUsers.Count
            Details = if ($disabledUsers.Count -gt 0) { "⚠️ Review required" } else { "✓ No issues found" }
            Status = if ($disabledUsers.Count -gt 0) { "Warning" } else { "Good" }
        }
        $summaryData += [PSCustomObject]@{
            Category = "Security Metrics"
            Metric = "PIM Eligible Assignments"
            Value = $pimEligible.Count
            Details = if ($pimEligible.Count -eq 0) { "⚠️ Consider implementing PIM" } else { "✓ PIM in use" }
            Status = if ($pimEligible.Count -eq 0) { "Warning" } else { "Good" }
        }
        $summaryData += [PSCustomObject]@{
            Category = "Security Metrics"
            Metric = "PIM Active Assignments"
            Value = $pimActive.Count
            Details = "Currently activated PIM roles"
            Status = "Info"
        }
        
        # Authentication Overview
        foreach ($authType in $authTypes) {
            $summaryData += [PSCustomObject]@{
                Category = "Authentication Methods"
                Metric = "$($authType.Name) Authentication"
                Value = $authType.Count
                Details = "Assignments using this auth method"
                Status = if ($authType.Name -eq "Certificate") { "Good" } elseif ($authType.Name -eq "ClientSecret") { "Warning" } else { "Info" }
            }
        }
        
        # Service Breakdown
        foreach ($service in $services) {
            $percentage = [math]::Round(($service.Count / $totalAssignments) * 100, 1)
            $summaryData += [PSCustomObject]@{
                Category = "Service Breakdown"
                Metric = $service.Name
                Value = $service.Count
                Details = "$percentage% of total assignments"
                Status = "Info"
            }
        }
        
        # Export Summary Sheet with formatting
        $summaryData | Export-Excel -Path $OutputPath -WorksheetName "Summary" -TableStyle Medium2 -AutoSize -FreezeTopRow
        
        # Add conditional formatting to Summary sheet
        $excel = Open-ExcelPackage -Path $OutputPath
        $summaryWorksheet = $excel.Workbook.Worksheets["Summary"]
        
        # Format Status column with colors
        $statusColumn = 5 # Column E (Status)
        $dataRange = $summaryWorksheet.Dimension
        
        for ($row = 2; $row -le $dataRange.End.Row; $row++) {
            $status = $summaryWorksheet.Cells[$row, $statusColumn].Text
            switch ($status) {
                "Good" { 
                    $summaryWorksheet.Cells[$row, $statusColumn].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $summaryWorksheet.Cells[$row, $statusColumn].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGreen)
                }
                "Warning" { 
                    $summaryWorksheet.Cells[$row, $statusColumn].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $summaryWorksheet.Cells[$row, $statusColumn].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Yellow)
                }
                "Info" { 
                    $summaryWorksheet.Cells[$row, $statusColumn].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $summaryWorksheet.Cells[$row, $statusColumn].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightBlue)
                }
            }
        }
        
        Close-ExcelPackage $excel
        
        # Create service-specific sheets grouped by role
        Write-Host "Creating service sheets grouped by role..." -ForegroundColor Cyan
        foreach ($service in $services) {
            $serviceData = $AuditResults | Where-Object { $_.Service -eq $service.Name }
            $serviceRoleGroups = $serviceData | Group-Object RoleName | Sort-Object Count -Descending
            
            $roleGroupData = @()
            foreach ($roleGroup in $serviceRoleGroups) {
                foreach ($assignment in $roleGroup.Group) {
                    $roleGroupData += [PSCustomObject]@{
                        RoleName = $assignment.RoleName
                        UserPrincipalName = $assignment.UserPrincipalName
                        DisplayName = $assignment.DisplayName
                        AssignmentType = $assignment.AssignmentType
                        AssignedDateTime = $assignment.AssignedDateTime
                        UserEnabled = $assignment.UserEnabled
                        LastSignIn = $assignment.LastSignIn
                        Scope = $assignment.Scope
                        AuthenticationType = $assignment.AuthenticationType
                        PIMEndDateTime = $assignment.PIMEndDateTime
                        RoleCount = $roleGroup.Count
                    }
                }
            }
            
            $safeSheetName = ($service.Name -replace '[^\w\s-]', '').Substring(0, [Math]::Min(25, $service.Name.Length)) + " by Role"
            $roleGroupData | Export-Excel -Path $OutputPath -WorksheetName $safeSheetName -TableStyle Medium6 -AutoSize -FreezeTopRow
        }
        
        # Create service-specific sheets grouped by assignee
        Write-Host "Creating service sheets grouped by assignee..." -ForegroundColor Cyan
        foreach ($service in $services) {
            $serviceData = $AuditResults | Where-Object { $_.Service -eq $service.Name }
            $serviceUserGroups = $serviceData | Group-Object UserPrincipalName | Sort-Object Count -Descending
            
            $userGroupData = @()
            foreach ($userGroup in $serviceUserGroups) {
                if ($userGroup.Name) {
                    $userRoles = $userGroup.Group | Sort-Object RoleName
                    $userInfo = $userRoles | Select-Object -First 1
                    
                    foreach ($assignment in $userRoles) {
                        $userGroupData += [PSCustomObject]@{
                            UserPrincipalName = $assignment.UserPrincipalName
                            DisplayName = $assignment.DisplayName
                            UserEnabled = $assignment.UserEnabled
                            LastSignIn = $assignment.LastSignIn
                            RoleName = $assignment.RoleName
                            AssignmentType = $assignment.AssignmentType
                            AssignedDateTime = $assignment.AssignedDateTime
                            Scope = $assignment.Scope
                            AuthenticationType = $assignment.AuthenticationType
                            PIMEndDateTime = $assignment.PIMEndDateTime
                            TotalRolesInService = $userGroup.Count
                        }
                    }
                }
            }
            
            $safeSheetName = ($service.Name -replace '[^\w\s-]', '').Substring(0, [Math]::Min(22, $service.Name.Length)) + " by User"
            $userGroupData | Export-Excel -Path $OutputPath -WorksheetName $safeSheetName -TableStyle Medium10 -AutoSize -FreezeTopRow
        }
        
        # Create detailed assignments sheet
        Write-Host "Creating detailed assignments sheet..." -ForegroundColor Cyan
        $detailedData = $AuditResults | Select-Object Service, UserPrincipalName, DisplayName, RoleName, AssignmentType, 
                                                   AssignedDateTime, UserEnabled, LastSignIn, Scope, AuthenticationType,
                                                   PIMEndDateTime, RoleDefinitionId, AssignmentId | Sort-Object Service, RoleName, UserPrincipalName
        
        $detailedData | Export-Excel -Path $OutputPath -WorksheetName "All Assignments" -TableStyle Medium14 -AutoSize -FreezeTopRow
        
        # Create top roles analysis sheet
        Write-Host "Creating top roles analysis sheet..." -ForegroundColor Cyan
        $topRolesData = @()
        foreach ($role in $topRoles | Select-Object -First 15) {
            $roleAssignments = $AuditResults | Where-Object { $_.RoleName -eq $role.Name }
            $services = ($roleAssignments | Group-Object Service).Name -join ", "
            $users = ($roleAssignments | Group-Object UserPrincipalName).Count
            $disabledUsersInRole = ($roleAssignments | Where-Object { $_.UserEnabled -eq $false }).Count
            $pimEligibleInRole = ($roleAssignments | Where-Object { $_.AssignmentType -like "*Eligible*" }).Count
            $pimActiveInRole = ($roleAssignments | Where-Object { $_.AssignmentType -like "*Active (PIM*" }).Count
            
            # Determine risk level based on role name
            $riskLevel = switch -Regex ($role.Name) {
                "Global Administrator|Company Administrator" { "CRITICAL" }
                "Security Administrator|Exchange Administrator|SharePoint Administrator|Intune Service Administrator" { "HIGH" }
                ".*Administrator.*|.*Admin.*" { "MEDIUM" }
                default { "LOW" }
            }
            
            $topRolesData += [PSCustomObject]@{
                RoleName = $role.Name
                TotalAssignments = $role.Count
                UniqueUsers = $users
                RiskLevel = $riskLevel
                ServicesUsedIn = $services
                DisabledUsersWithRole = $disabledUsersInRole
                PIMEligibleAssignments = $pimEligibleInRole
                PIMActiveAssignments = $pimActiveInRole
                PercentageOfTotal = [math]::Round(($role.Count / $totalAssignments) * 100, 1)
                Status = if ($disabledUsersInRole -gt 0) { "NEEDS REVIEW" } elseif ($pimEligibleInRole -gt 0) { "PIM ENABLED" } else { "ACTIVE" }
            }
        }
        
        $topRolesData | Export-Excel -Path $OutputPath -WorksheetName "Top Roles Analysis" -TableStyle Medium16 -AutoSize -FreezeTopRow
        
        # Create top users analysis sheet
        Write-Host "Creating top users analysis sheet..." -ForegroundColor Cyan
        $topUsersData = @()
        foreach ($user in $usersWithMostRoles | Select-Object -First 20) {
            if ($user.Name) {
                $userRoles = $AuditResults | Where-Object { $_.UserPrincipalName -eq $user.Name }
                $userInfo = $userRoles | Select-Object -First 1
                $services = ($userRoles | Group-Object Service).Name -join ", "
                $roles = ($userRoles | Group-Object RoleName).Name -join ", "
                
                $topUsersData += [PSCustomObject]@{
                    UserPrincipalName = $user.Name
                    DisplayName = $userInfo.DisplayName
                    TotalRoles = $user.Count
                    UserEnabled = $userInfo.UserEnabled
                    LastSignIn = $userInfo.LastSignIn
                    ServicesWithRoles = $services
                    RoleNames = $roles
                    Status = if ($userInfo.UserEnabled -eq $false) { "DISABLED" } else { "Active" }
                }
            }
        }
        
        $topUsersData | Export-Excel -Path $OutputPath -WorksheetName "Top Users Analysis" -TableStyle Medium18 -AutoSize -FreezeTopRow
        
        # Create security alerts sheet
        Write-Host "Creating security alerts sheet..." -ForegroundColor Cyan
        $alertsData = @()
        
        if ($globalAdmins.Count -gt 5) {
            $alertsData += [PSCustomObject]@{
                AlertType = "Excessive Global Admins"
                Severity = "High"
                Count = $globalAdmins.Count
                Recommendation = "Reduce Global Administrator count to 5 or fewer"
                Details = ($globalAdmins.UserPrincipalName -join "; ")
            }
        }
        
        if ($disabledUsers.Count -gt 0) {
            $disabledUsersList = ($disabledUsers | Select-Object -First 10).UserPrincipalName -join "; "
            if ($disabledUsers.Count -gt 10) { $disabledUsersList += "... and $($disabledUsers.Count - 10) more" }
            
            $alertsData += [PSCustomObject]@{
                AlertType = "Disabled Users with Roles"
                Severity = "Medium"
                Count = $disabledUsers.Count
                Recommendation = "Remove role assignments from disabled accounts"
                Details = $disabledUsersList
            }
        }
        
        $clientSecretAuth = $authTypes | Where-Object { $_.Name -eq "ClientSecret" }
        if ($clientSecretAuth) {
            $alertsData += [PSCustomObject]@{
                AlertType = "Client Secret Authentication"
                Severity = "Medium"
                Count = $clientSecretAuth.Count
                Recommendation = "Migrate to certificate-based authentication"
                Details = "Some connections use less secure client secret authentication"
            }
        }
        
        if ($pimEligible.Count -eq 0 -and $totalAssignments -gt 0) {
            $alertsData += [PSCustomObject]@{
                AlertType = "No PIM Eligible Assignments"
                Severity = "Medium"
                Count = 0
                Recommendation = "Implement Privileged Identity Management (PIM)"
                Details = "All privileged roles appear to be permanently assigned"
            }
        }
        
        # Check for users with excessive roles (>10)
        $excessiveRoleUsers = $usersWithMostRoles | Where-Object { $_.Count -gt 10 }
        if ($excessiveRoleUsers.Count -gt 0) {
            $excessiveUsersList = ($excessiveRoleUsers | Select-Object -First 5).Name -join "; "
            if ($excessiveRoleUsers.Count -gt 5) { $excessiveUsersList += "... and $($excessiveRoleUsers.Count - 5) more" }
            
            $alertsData += [PSCustomObject]@{
                AlertType = "Role Sprawl"
                Severity = "Low"
                Count = $excessiveRoleUsers.Count
                Recommendation = "Review role assignments for principle of least privilege"
                Details = "Users with >10 roles: $excessiveUsersList"
            }
        }
        
        if ($alertsData.Count -eq 0) {
            $alertsData += [PSCustomObject]@{
                AlertType = "No Issues Found"
                Severity = "Good"
                Count = 0
                Recommendation = "Continue monitoring and regular access reviews"
                Details = "No significant security alerts identified"
            }
        }
        
        $alertsData | Export-Excel -Path $OutputPath -WorksheetName "Security Alerts" -TableStyle Medium22 -AutoSize -FreezeTopRow
        
        # Add conditional formatting to Security Alerts sheet
        $excel = Open-ExcelPackage -Path $OutputPath
        $alertsWorksheet = $excel.Workbook.Worksheets["Security Alerts"]
        
        # Format Severity column with colors
        $severityColumn = 2 # Column B (Severity)
        $alertsRange = $alertsWorksheet.Dimension
        
        for ($row = 2; $row -le $alertsRange.End.Row; $row++) {
            $severity = $alertsWorksheet.Cells[$row, $severityColumn].Text
            switch ($severity) {
                "High" { 
                    $alertsWorksheet.Cells[$row, $severityColumn].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $alertsWorksheet.Cells[$row, $severityColumn].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Red)
                    $alertsWorksheet.Cells[$row, $severityColumn].Style.Font.Color.SetColor([System.Drawing.Color]::White)
                }
                "Medium" { 
                    $alertsWorksheet.Cells[$row, $severityColumn].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $alertsWorksheet.Cells[$row, $severityColumn].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Orange)
                }
                "Low" { 
                    $alertsWorksheet.Cells[$row, $severityColumn].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $alertsWorksheet.Cells[$row, $severityColumn].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::Yellow)
                }
                "Good" { 
                    $alertsWorksheet.Cells[$row, $severityColumn].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $alertsWorksheet.Cells[$row, $severityColumn].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightGreen)
                }
            }
        }
        
        Close-ExcelPackage $excel
        
        Write-Host "✓ Excel report generated successfully: $OutputPath" -ForegroundColor Green
        Write-Host "File size: $([math]::Round((Get-Item $OutputPath).Length / 1MB, 2)) MB" -ForegroundColor Gray
        
        # Display sheet summary
        $excel = Open-ExcelPackage -Path $OutputPath
        Write-Host ""
        Write-Host "=== Excel Report Contents ===" -ForegroundColor Cyan
        foreach ($worksheet in $excel.Workbook.Worksheets) {
            $rowCount = if ($worksheet.Dimension) { $worksheet.Dimension.End.Row - 1 } else { 0 }
            Write-Host "  $($worksheet.Name): $rowCount rows" -ForegroundColor White
        }
        Close-ExcelPackage $excel
        
        # Auto-open if requested
        if ($AutoOpen -and (Test-Path $OutputPath)) {
            if ($IsWindows -ne $false) {
                Write-Host "Opening Excel report..." -ForegroundColor Green
                Start-Process $OutputPath
            }
            else {
                Write-Host "Auto-open not supported on this platform" -ForegroundColor Yellow
            }
        }
        
        return $OutputPath
    }
    catch {
        Write-Error "Failed to generate Excel report: $($_.Exception.Message)"
        Write-Error "Stack trace: $($_.ScriptStackTrace)"
        return $null
    }
}