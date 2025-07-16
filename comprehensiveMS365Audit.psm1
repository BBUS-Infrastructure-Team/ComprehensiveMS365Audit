# Load files in order
Get-ChildItem -Path "$PSScriptRoot" -Filter "*.ps1" | ForEach-Object {
    . "$($_.FullName)"
}

Export-ModuleMember -Function @(
    'Set-M365AuditCredentials',
    'Set-M365AuditAppCredentials',
    'Clear-M365AuditAppCredentials',
    'Get-M365AuditCurrentConfig',
    'Connect-M365ServiceWithAuth',
    'Get-M365AuditRequiredPermissions',
    'Initialize-M365AuditEnvironment',
    'Get-ComprehensiveM365RoleAudit',
    'Get-AzureADRoleAudit',
    'Get-TeamsRoleAudit',
    'Get-DefenderRoleAudit'
    'Get-SharePointRoleAudit',
    'Get-ExchangeRoleAudit',
    'Get-PurviewRoleAudit',
    'Get-PowerPlatformRoleAudit',
    'Get-IntuneRoleAudit',
    'Get-PowerPlatformAzureADRoleAudit',
    'Test-YourCurrentFunction',
    'Export-M365AuditHtmlReport',
    'Export-M365AuditJsonReport',
    'Get-M365RoleAnalysis',
    'Get-M365ComplianceGaps',
    'Export-M365ComplianceGapsHtmlReport',
    'Export-M365AuditExcelReport',
    'Export-M365ServiceAuditHtmlReport'
    'Export-M365ServiceAuditJsonReport'
)