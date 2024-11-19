<#
.SYNOPSIS
    Audits external sharing events in SharePoint Online and OneDrive for Business.

.DESCRIPTION
    This script retrieves and reports external sharing events from SharePoint Online and/or OneDrive for Business.
    It generates detailed reports in various formats (CSV, HTML, JSON) showing who shared what with whom.

.PARAMETER StartDate
    The start date for the audit search. Defaults to 5 days ago.

.PARAMETER EndDate
    The end date for the audit search. Defaults to current date.

.PARAMETER SharePointOnline
    Switch to audit SharePoint Online sharing events only.

.PARAMETER OneDrive
    Switch to audit OneDrive sharing events only.

.PARAMETER ReportFormat
    Output format for the report (CSV, HTML, JSON, or ALL). Defaults to ALL.

.NOTES
    Author: Cengiz YILMAZ
    Version: 1.0
    Created: 2024-01-10
    Modified: 2024-01-10

.EXAMPLE
    .\ExternalSharingAudit.ps1
    Runs the script for both SharePoint Online and OneDrive

.EXAMPLE
    .\ExternalSharingAudit.ps1 -SharePointOnline
    Runs the script for SharePoint Online only

.EXAMPLE
    .\ExternalSharingAudit.ps1 -OneDrive
    Runs the script for OneDrive only
#>

#Requires -Version 5.1
#Requires -Modules @{ ModuleName="ExchangeOnlineManagement"; ModuleVersion="3.0.0" }

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [ValidateScript({
        if ($_ -gt (Get-Date)) {
            throw "StartDate cannot be in the future"
        }
        return $true
    })]
    [DateTime]$StartDate = ((Get-Date).AddDays(-5)).Date,

    [Parameter(Mandatory = $false)]
    [ValidateScript({
        if ($_ -gt (Get-Date)) {
            throw "EndDate cannot be in the future"
        }
        return $true
    })]
    [DateTime]$EndDate = (Get-Date),

    [Parameter(Mandatory = $false)]
    [switch]$SharePointOnline,
    
    [Parameter(Mandatory = $false)]
    [switch]$OneDrive,
    
    [ValidateSet("CSV", "HTML", "JSON", "ALL")]
    [string]$ReportFormat = "ALL",

    [string]$Organization,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [string]$AdminName,
    [SecureString]$Password
)

# Configuration
$script:Config = @{
    ReportPath = Join-Path (Get-Location) "Reports"
    TimeStamp = Get-Date -Format "yyyy-MM-dd_HH-mm"
    TimeZone = [System.TimeZoneInfo]::Local
    Culture = [System.Globalization.CultureInfo]::GetCultureInfo('en-US')
    MaxStartDate = ((Get-Date).AddDays(-179)).Date
    BatchSize = 5000
    IntervalMinutes = 1440
    Thresholds = @{
        Warning = 100
        Critical = 500
    }
}

# Progress tracking
$script:Progress = @{
    Activity = "Processing External Sharing Audit"
    Status = "Initializing..."
    PercentComplete = 0
    CurrentOperation = ""
}

# Logging functions
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("Info", "Warning", "Error")]
        [string]$Level = "Info"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    switch ($Level) {
        "Info" { Write-Host $logMessage -ForegroundColor Cyan }
        "Warning" { Write-Host $logMessage -ForegroundColor Yellow }
        "Error" { 
            Write-Host $logMessage -ForegroundColor Red
            Add-Content -Path "$($Config.ReportPath)\error.log" -Value $logMessage
        }
    }
}

function Update-ProgressStatus {
    param(
        [string]$Status,
        [string]$CurrentOperation,
        [int]$PercentComplete
    )
    
    $Progress.Status = $Status
    $Progress.CurrentOperation = $CurrentOperation
    $Progress.PercentComplete = $PercentComplete
    Write-Progress @Progress
}

# Environment setup
function Initialize-Environment {
    try {
        Update-ProgressStatus -Status "Initializing environment" -CurrentOperation "Creating directories" -PercentComplete 10
        
        if (-not (Test-Path $Config.ReportPath)) {
            New-Item -ItemType Directory -Path $Config.ReportPath -Force | Out-Null
        }

        Update-ProgressStatus -Status "Checking dependencies" -CurrentOperation "Verifying ExchangeOnlineManagement module" -PercentComplete 30
        
        if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            Write-Log "Installing ExchangeOnlineManagement module..." -Level Warning
            Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
        }

        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Update-ProgressStatus -Status "Environment ready" -CurrentOperation "Complete" -PercentComplete 100
        return $true
    }
    catch {
        Write-Log "Failed to initialize environment: $_" -Level Error
        return $false
    }
}

# Authentication
function Connect-ToService {
    try {
        Update-ProgressStatus -Status "Connecting to Exchange Online" -CurrentOperation "Authenticating..." -PercentComplete 0
        
        if ($Organization -and $ClientId -and $CertificateThumbprint) {
            Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -Organization $Organization -ShowBanner:$false -ErrorAction Stop
        }
        elseif ($AdminName -and $Password) {
            $credential = New-Object System.Management.Automation.PSCredential($AdminName, $Password)
            Connect-ExchangeOnline -Credential $credential -ShowBanner:$false -ErrorAction Stop
        }
        else {
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        }
        
        Write-Log "Successfully connected to Exchange Online"
        return $true
    }
    catch {
        Write-Log "Failed to connect to Exchange Online: $_" -Level Error
        return $false
    }
}

# Data collection
function Get-SharingEvents {
    param (
        [DateTime]$Start,
        [DateTime]$End
    )

    try {
        $events = Search-UnifiedAuditLog -StartDate $Start -EndDate $End `
            -Operations "Sharinginvitationcreated", "AnonymousLinkcreated", "AddedToSecureLink" `
            -ResultSize $Config.BatchSize -ErrorAction Stop

        $processedEvents = @()
        foreach ($event in $events) {
            $auditData = $event.AuditData | ConvertFrom-Json

            # Filter events based on parameters
            if ($SharePointOnline -and -not $OneDrive -and $auditData.Workload -eq "OneDrive") {
                continue
            }
            if ($OneDrive -and -not $SharePointOnline -and $auditData.Workload -eq "SharePoint") {
                continue
            }

            $sharedWith = if ($event.Operations -ne "AnonymousLinkcreated") {
                if ($auditData.TargetUserOrGroupType -ne "Guest") { continue }
                $auditData.TargetUserOrGroupName
            }
            else {
                "Anyone with the link"
            }

            $localTime = [System.TimeZoneInfo]::ConvertTimeFromUtc(
                [DateTime]::Parse($auditData.CreationTime),
                $Config.TimeZone
            )

            $processedEvents += [PSCustomObject]@{
                'Sharing Time' = $localTime.ToString("g", $Config.Culture)
                'Shared By' = $auditData.UserId
                'Shared With' = $sharedWith
                'Resource Type' = $auditData.ItemType
                'Resource' = $auditData.ObjectId
                'Site URL' = $auditData.SiteUrl
                'Sharing Type' = $event.Operations
                'System' = $auditData.Workload
                'More Info' = $event.AuditData
            }
        }

        return $processedEvents
    }
    catch {
        Write-Log "Failed to retrieve sharing events: $_" -Level Error
        return @()
    }
}

# Report generation
function Export-Results {
    param (
        [Array]$Data,
        [string]$Format
    )

    $reportFiles = @{
        CSV = Join-Path $Config.ReportPath "SharingReport_$($Config.TimeStamp).csv"
        HTML = Join-Path $Config.ReportPath "SharingReport_$($Config.TimeStamp).html"
        JSON = Join-Path $Config.ReportPath "SharingReport_$($Config.TimeStamp).json"
    }

    try {
        switch ($Format) {
            "CSV" {
                $Data | Select-Object 'Sharing Time','Shared By','Shared With','Resource Type','Resource','Site URL','Sharing Type','System','More Info' | 
                Export-Csv -Path $reportFiles.CSV -NoTypeInformation -Encoding UTF8
                Write-Log "CSV report saved: $($reportFiles.CSV)"
            }
            "HTML" {
                $htmlHead = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>External Sharing Audit Report</title>
    <style>
        :root {
            --primary-color: #2563eb;
            --background: #f8fafc;
            --surface: #ffffff;
            --text: #1e293b;
            --text-light: #64748b;
            --border: #e2e8f0;
            --danger: #ef4444;
            --warning: #f59e0b;
        }

        body {
            font-family: system-ui, -apple-system, sans-serif;
            line-height: 1.6;
            margin: 0;
            padding: 2rem;
            background: var(--background);
            color: var(--text);
        }

        .container {
            max-width: 100%;
            margin: 0 auto;
            background: var(--surface);
            padding: 2rem;
            border-radius: 1rem;
            box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1);
        }

        h1 {
            color: var(--primary-color);
            font-size: 1.875rem;
            font-weight: 700;
            margin-bottom: 2rem;
            border-bottom: 2px solid var(--border);
            padding-bottom: 1rem;
        }

        .summary {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 1.5rem;
            margin-bottom: 2rem;
            padding: 1.5rem;
            background: var(--background);
            border-radius: 0.75rem;
            border-left: 4px solid var(--primary-color);
        }

        .table-wrapper {
            overflow-x: auto;
            margin: 2rem 0;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 0.875rem;
            table-layout: fixed;
        }

        th {
            background: var(--primary-color);
            color: white;
            font-weight: 600;
            text-transform: uppercase;
            font-size: 0.75rem;
            letter-spacing: 0.05em;
            padding: 1rem;
            text-align: left;
            position: sticky;
            top: 0;
        }

        td {
            padding: 1rem;
            border-bottom: 1px solid var(--border);
            vertical-align: top;
            word-wrap: break-word;
            max-width: 300px;
        }

        .details-content {
            white-space: pre-wrap;
            word-break: break-word;
            font-size: 0.8rem;
            background: var(--background);
            padding: 0.5rem;
            border-radius: 0.25rem;
            margin: 0;
            max-height: 200px;
            overflow-y: auto;
        }

        tr:nth-child(even) {
            background: var(--background);
        }

        .warning-banner {
            margin: 1rem 0;
            padding: 1rem;
            border-radius: 0.5rem;
            font-weight: 500;
            display: flex;
            align-items: center;
            gap: 0.5rem;
        }

        .warning-banner.critical {
            background: #fef2f2;
            color: var(--danger);
            border: 1px solid #fee2e2;
        }

        .warning-banner.warning {
            background: #fffbeb;
            color: var(--warning);
            border: 1px solid #fef3c7;
        }

        .footer {
            margin-top: 2rem;
            padding-top: 1rem;
            border-top: 1px solid var(--border);
            text-align: center;
            color: var(--text-light);
        }

        .footer a {
            color: var(--primary-color);
            text-decoration: none;
        }

        .footer a:hover {
            text-decoration: underline;
        }

        @media screen and (max-width: 1024px) {
            body { padding: 1rem; }
            .container { padding: 1rem; }
            td, th { padding: 0.75rem; }
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>External Sharing Audit Report</h1>
        <div class="summary">
            <div>
                <strong>Generated:</strong> $(Get-Date -Format "g")
            </div>
            <div>
                <strong>Total Records:</strong> $($Data.Count)
            </div>
            <div>
                <strong>Date Range:</strong> $StartDate to $EndDate
            </div>
        </div>
"@
                $warning = if ($Data.Count -ge $Config.Thresholds.Critical) {
                    "<div class='warning-banner critical'>⚠️ Critical: Sharing events ($($Data.Count)) exceed critical threshold ($($Config.Thresholds.Critical))</div>"
                } 
                elseif ($Data.Count -ge $Config.Thresholds.Warning) {
                    "<div class='warning-banner warning'>⚠️ Warning: Sharing events ($($Data.Count)) exceed warning threshold ($($Config.Thresholds.Warning))</div>"
                } 
                else { "" }

                $footer = @"
        <div class="footer">
            <p>Created by <a href="https://yilmazcengiz.tr" target="_blank">Cengiz YILMAZ</a> - Microsoft MVP</p>
        </div>
"@

                $htmlTable = "<div class='table-wrapper'><table>"
                $htmlTable += "<tr>"
                @('Sharing Time','Shared By','Shared With','Resource Type','Resource','Site URL','Sharing Type','System','More Info') | ForEach-Object {
                    $htmlTable += "<th>$_</th>"
                }
                $htmlTable += "</tr>"

                $index = 0
                foreach ($event in $Data) {
                    $htmlTable += "<tr>"
                    $htmlTable += "<td>$($event.'Sharing Time')</td>"
                    $htmlTable += "<td>$($event.'Shared By')</td>"
                    $htmlTable += "<td>$($event.'Shared With')</td>"
                    $htmlTable += "<td>$($event.'Resource Type')</td>"
                    $htmlTable += "<td>$($event.Resource)</td>"
                    $htmlTable += "<td>$($event.'Site URL')</td>"
                    $htmlTable += "<td>$($event.'Sharing Type')</td>"
                    $htmlTable += "<td>$($event.System)</td>"
                    $htmlTable += "<td><pre class='details-content'>$($event.'More Info')</pre></td>"
                    $htmlTable += "</tr>"
                    $index++
                }
                $htmlTable += "</table></div>"

                $htmlReport = $htmlHead + $warning + $htmlTable + $footer + "</div></body></html>"
                $htmlReport | Out-File -FilePath $reportFiles.HTML -Encoding UTF8
                Write-Log "HTML report saved: $($reportFiles.HTML)"
            }
            "JSON" {
                $Data | ConvertTo-Json -Depth 10 | Out-File -FilePath $reportFiles.JSON -Encoding UTF8
                Write-Log "JSON report saved: $($reportFiles.JSON)"
            }
        }
        return $true
    }
    catch {
        Write-Log "Failed to export results in $Format format: $_" -Level Error
        return $false
    }
}

# Main execution
try {
    if (-not (Initialize-Environment)) {
        throw "Environment initialization failed"
    }

    if (-not (Connect-ToService)) {
        throw "Connection to Exchange Online failed"
    }

    Write-Log "`nStarting data collection: $StartDate - $EndDate"
    $allEvents = @()
    $current = $StartDate
    $totalMinutes = ($EndDate - $StartDate).TotalMinutes
    $processedMinutes = 0

    while ($current -le $EndDate) {
        $batchEnd = $current.AddMinutes($Config.IntervalMinutes)
        if ($batchEnd -gt $EndDate) { $batchEnd = $EndDate }

        $processedMinutes = ($current - $StartDate).TotalMinutes
        $percentComplete = [math]::Min(100, [math]::Round(($processedMinutes / $totalMinutes) * 100))
        
        Update-ProgressStatus -Status "Collecting sharing events" `
                            -CurrentOperation "Processing: $current to $batchEnd" `
                            -PercentComplete $percentComplete

        $events = Get-SharingEvents -Start $current -End $batchEnd
        $allEvents += $events

        if ($batchEnd -eq $EndDate) { break }
        $current = $batchEnd
    }

    Update-ProgressStatus -Status "Exporting results" -CurrentOperation "Generating reports" -PercentComplete 90

    switch ($ReportFormat.ToUpper()) {
        "CSV" { Export-Results -Data $allEvents -Format "CSV" }
        "HTML" { Export-Results -Data $allEvents -Format "HTML" }
        "JSON" { Export-Results -Data $allEvents -Format "JSON" }
        "ALL" {
            Export-Results -Data $allEvents -Format "CSV"
            Export-Results -Data $allEvents -Format "HTML"
            Export-Results -Data $allEvents -Format "JSON"
        }
    }

    Write-Host "`nReport Summary:" -ForegroundColor Cyan
    Write-Host "==============="
    Write-Host "Total Records: $($allEvents.Count)"
    Write-Host "Date Range: $StartDate to $EndDate"
    Write-Host "Report Location: $($Config.ReportPath)"

    if ($allEvents.Count -ge $Config.Thresholds.Critical) {
        Write-Log "Sharing events ($($allEvents.Count)) exceed critical threshold ($($Config.Thresholds.Critical))!" -Level Error
    }
    elseif ($allEvents.Count -ge $Config.Thresholds.Warning) {
        Write-Log "Sharing events ($($allEvents.Count)) exceed warning threshold ($($Config.Thresholds.Warning))!" -Level Warning
    }

    <#
    if ($allEvents.Count -gt 0) {
        $prompt = New-Object -ComObject wscript.shell
        $userInput = $prompt.popup("Do you want to open the HTML report?", 0, "Open Report", 4)
        if ($userInput -eq 6) {
            Invoke-Item (Join-Path $Config.ReportPath "SharingReport_$($Config.TimeStamp).html")
        }
    }
    #>
}
catch {
    Write-Log $_.Exception.Message -Level Error
    throw
}
finally {
    Write-Progress -Activity $Progress.Activity -Completed
    Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
}