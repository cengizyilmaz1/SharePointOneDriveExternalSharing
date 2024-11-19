# SharePointOneDriveExternalSharing
# External Sharing Audit Script for Microsoft 365

PowerShell script to audit external sharing events in SharePoint Online and OneDrive for Business.

## Features

- Audits external sharing events across SharePoint Online and OneDrive
- Generates reports in multiple formats (CSV, HTML, JSON)
- Supports Certificate-Based Authentication (CBA)
- Configurable warning thresholds
- Detailed logging and progress tracking

## Prerequisites

- PowerShell 5.1 or later
- ExchangeOnlineManagement Module v3.0.0 or later
- Microsoft 365 Admin account or Certificate-Based Authentication setup

## Quick Start

1. **Basic Usage**
```powershell
.\ExternalSharingAudit.ps1
```

2. **SharePoint Only**
```powershell
.\ExternalSharingAudit.ps1 -SharePointOnline
```

3. **OneDrive Only**
```powershell
.\ExternalSharingAudit.ps1 -OneDrive
```

4. **Custom Date Range**
```powershell
.\ExternalSharingAudit.ps1 -StartDate (Get-Date).AddDays(-30) -EndDate (Get-Date)
```

5. **Specific Report Format**
```powershell
.\ExternalSharingAudit.ps1 -ReportFormat HTML
```

6. **Using Certificate-Based Authentication**
```powershell
.\ExternalSharingAudit.ps1 -Organization "contoso.onmicrosoft.com" `
                          -ClientId "YOUR_APP_ID" `
                          -CertificateThumbprint "YOUR_CERT_THUMBPRINT"
```

## Parameters

| Parameter | Description | Required |
|-----------|-------------|----------|
| StartDate | Start date for audit search | No (Default: 5 days ago) |
| EndDate | End date for audit search | No (Default: Current date) |
| SharePointOnline | Audit SharePoint only | No |
| OneDrive | Audit OneDrive only | No |
| ReportFormat | Output format (CSV/HTML/JSON/ALL) | No (Default: ALL) |
| Organization | Microsoft 365 tenant name | For CBA only |
| ClientId | Azure AD App ID | For CBA only |
| CertificateThumbprint | Certificate thumbprint | For CBA only |

## Report Types

- **CSV**: Detailed spreadsheet format
- **HTML**: Interactive web report with filtering
- **JSON**: Machine-readable format for automation

## Warning Thresholds

- **Warning**: 100+ sharing events
- **Critical**: 500+ sharing events

## Required Permissions

### Basic Authentication
- Global Administrator or
- Security Administrator or
- Compliance Administrator

### Certificate-Based Authentication (CBA)
- Exchange.ManageAsApp
- ActivityFeed.Read
- ActivityFeed.ReadDlp
- ServiceHealth.Read

## Output Location

Reports are saved in the `Reports` folder with timestamp:
- `SharingReport_yyyy-MM-dd_HH-mm.csv`
- `SharingReport_yyyy-MM-dd_HH-mm.html`
- `SharingReport_yyyy-MM-dd_HH-mm.json`

## Author

Created by Cengiz YILMAZ - Microsoft MVP
