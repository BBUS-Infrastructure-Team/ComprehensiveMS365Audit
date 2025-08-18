# Comprehensive Microsoft 365 Audit Module.

This module allows you to do a comprehensive audit of all administrative roles in Office 365. Roles can be audited in Entra ID, Exchange Online, Microsoft Purview, Microsoft SharePoint, Microsoft Teams, Microsoft Defender, amd Microsoft Power Platform.

This module is a script module that has not been published to the PowerShell Gallery.

# Installation
Clone the repository with Git.

```bash
# Change into the folder where you want to repository to reside.
git clone https://github.com/BBUS-Infrastructure-Team/ComprehensiveMS365Audit.git
```
This will create the folder comprehensiveM365RoleAudit.

To import the module:

```powershell
Import-Module [path to]\comprehensiveM365Audit
```

If you place the module into a folder that is in the PSModulePath you don't have to use the full path.

## Authentication.

### Retrieve the Application Keys.

The application keys are kept in AWS Secrets Manager. To retrieve the keys you must have access to the secret. App Full AWS Administrators have access

To retrieve the keys use the following PowerShell commands:

```powershell
$BBUSAutit = (Get-SECSecretValue -SecretId 'BBUS_Audit').SecretString | ConvertFrom-Json
```
!!!NOTE
    The secret name is case sensitive.
    

This will put the keys into the object $BBUSAudit.

From this object you can use the following properties:

- .ClientID: This is the application ID.
- .TenantID: This is the tenant ID.
- .CertificateThumbprint: This is the certificate Thumbprint for the certificate.

You will need this information to use the module.

### Install the application certificate.

The application certificates are available in the Infrastructure Team in Microsoft Teams. The certificates are located in:
[Private Documentation/Application Certificates](https://balfourbeattyus.sharepoint.com/:f:/r/sites/ITInfrastructure-PrivateDocumentation/Shared%20Documents/Private%20Documentation/Security/Application%20Certificates?csf=1&web=1&e=IFgNkj)

There are 3 versions of the certificate.

- M365_Audit_Certificate.cer: Standard certificate format.
- M365_Audit_Security_pfx: PFX export with privater key. This certificate is secured to the IT Infrastructure Active Directory group.
- M365_Audit_Security_Linux-pfx: PFX export to install in the .Net certificate store on linux systems. The passphrase for this certificate is in the Infrastructure Private Password.xlsx file in Private Documentation / Security.

#### Windows

Import the certificate M365_Audit_Certificate.pfx into your personal certificate store with the certificate MMC. If you are a member of the IT Infrastructure AD Group no passphrase is required.

If you require this certificate on a server for a scheduled task, you must install into the Local Machine store.

#### Linux

You will need to retrieve the application keys to complete this step.

Install the PowerShell module MOA_Module.

Run the following commands:

```powershell
$PassPhrase = ConvertTo-SecureString -String '[passphase]' -AsPlainText -Force

# for personal use
Install-x509Certificate -CertificatePath {Path to .pfx file} -Scope CurrentUser -Thumbprint $BBUSAUdit.CertificateThumbprint

# for global use.
# You must open powershell with elevated permissions using 'sudo pwsh'
Install-x509Certificate -CertificatePath {Path to .pfx file} -Scope LocalMachine -Thumbprint $BBUSAUdit.CertificateThumbprint
```

## Generating Audit Data

Before you can create reports you must first generate the audit data.

Ensure you have the proper certificate installed on your workstation.
Retrieve the application keys into a variable using the Get-SECSecretValue. e.g. $BBUSAudit.

There are 2 ways you can use the application keys.

1. Pass the ClientID, TenantId, and CertificateThumbprint with each audit command.
2. Set the application credentials globally.

### Setting the application authentication globally.

To set authentication globally use the following function:

```powershell
Set-M365AuditCredentials -ClientId $BBUSAudit.ClientId -TenantId $BBUSAudit.TenantId -CertificateThumbprint $BBUSAudit.CertificateThumbprint
```

This will set the credentials for as long as the module is imported.

### Audit Data Functions

The main audit data function is Get-ComprehensiveM365RoleAudit.
This function can generate data for all or part for the audited services.

You must either preset the application authentication or provide the application keys on the command line.

This command can produce a full set of audit data for all services or selected services. See the help text for more details.

There are also individual functions to audit selective services.

- Get-AzureADRoleAudit
- Get-DefenderRoleAudit
- Get-ExchangeRoleAudit
- Get-IntuneRoleAudit
- Get-PowerPlatformAzureADRoleAudit
- Get-PurviewRoleAudit
- Get-SharePointRoleAudit
- Get-TeamsRoleAudit

There are some services that require specific parameters.

- When auditing Exchange or Purview either through Get-ComprehensiveM365RoleAudit or Get-ExchangeRoleAudit you must provide the -Organization parameter. This is 'balfourbeattyus.onmicrosoft.com' for Balfour Beatty US.
- For SharePoint you must provide the SharePoint admin URL. For Balfour Beatty US this is https://balfourbeattyus-admin.sharepoint.com

## Generating Reports

You can generate both HTML and Excel reports from collected audit data.

Run one of the auditing functions and store the retrieved data into a variable. For example.

```powershell
# This assume application authentication is set globally
$AuditResults = Get-ComprehensiveM365RoleAudit -IncludeAll -Organization 'balfourbeattyus.onmicrosoft.com' -SharePointURL 'https://balfourbeattyus-admin.sharepoint.com'
```
### Generate an HTML Report
```powershell
Export-M365AuditHtmlReport -AuditResults $AuditResults -OrganizationName 'Balfour Beatty US'
```

### Generate an HTML Gap Analysis Report.
``` powershell
$Gaps = Get-M365ComplianceGaps -AuditResults $AUditResults 
Export-M365ComplianceGapsHtmlReport -ComplianceGaps $Gaps -OrganizationName 'Balfour Beatty US'
```

### Generate an Excel Report
```powershell
Export-M365AuditExcelReport -AuditResults $AUditResults -OrganizationName 'Balfour Beatty US'
```

### Generate an Excel Report with a Gap analysis sheet.
```powershell
Export-M365AuditExcelReport -AuditResults $AUditResults -OrganizationName 'Balfour Beatty US' -IncludeGapAnalysis
```

!!!Note
    Organization Name in the report functions is just a company name for the report. It is not related to -Organization in Exchange or Purview.