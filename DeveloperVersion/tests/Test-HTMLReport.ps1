<#
.SYNOPSIS
Check HTML report metadata, controls and colors using synthetic IME logs.
.EXAMPLE
.\tests\Test-HTMLReport.ps1 -WorkPath "C:\Temp\IME-HTML-Regression"
.NOTES
Use a new writable WorkPath. No device logs, Graph connection or browser are required.
#>
[CmdletBinding()]
param(
    [string]$ScriptPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'Get-IntuneManagementExtensionDiagnostics.ps1'),
    [Parameter(Mandatory=$true)]
    [string]$WorkPath
)

$ErrorActionPreference = 'Stop'
$ScriptPath = (Get-Item -LiteralPath $ScriptPath).FullName
$null = New-Item -ItemType Directory -Path $WorkPath
$WorkPath = (Get-Item -LiteralPath $WorkPath).FullName
$script:assertions = 0

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if(-not $Condition) { throw "FAIL: $Message" }
    $script:assertions++
}

$tokens = $null
$parseErrors = $null
$null = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$parseErrors)
Assert-True ($parseErrors.Count -eq 0) 'PowerShell syntax is valid'

$logPath = Join-Path $WorkPath 'IntuneManagementExtension.log'
$records = @(
    '<![LOG[EMS Agent Started]LOG]!><time="01:00:00.0000000" date="10-12-2026" component="IntuneManagementExtension" context="" type="1" thread="4" file="">',
    '<![LOG[Get policies = [{"Id":"11111111-1111-1111-1111-111111111111","Name":"Synthetic application","Version":1,"Intent":3,"TargetType":2,"InstallContext":1}]]LOG]!><time="01:00:00.0010000" date="10-12-2026" component="IntuneManagementExtension" context="" type="1" thread="4" file="">',
    '<![LOG[[APv2] Device is in APv2 mode: True.]LOG]!><time="01:00:01.0000000" date="10-12-2026" component="IntuneManagementExtension" context="" type="1" thread="4" file="">',
    '<![LOG[[APv2] Successfully disabled APv2 mode.]LOG]!><time="01:00:02.0000000" date="10-12-2026" component="IntuneManagementExtension" context="" type="1" thread="4" file="">'
)
[System.IO.File]::WriteAllLines($logPath, [string[]]$records, [System.Text.UTF8Encoding]::new($false))
& $ScriptPath -LogFile $logPath -AllLogEntries -AllLogFiles -ExportHTMLReportPath $WorkPath -DoNotOpenReportAutomatically *> (Join-Path $WorkPath 'console.txt')
$reports = @(Get-ChildItem -LiteralPath $WorkPath -Filter '*.html')
Assert-True ($reports.Count -eq 1) 'Synthetic logs produce one HTML report'
$html = [System.IO.File]::ReadAllText($reports[0].FullName)
Assert-True ([regex]::Matches($html, '<title>').Count -eq 1 -and $html.Contains('<title>Get-IntuneManagementExtensionDiagnostics Observed Timeline Report</title>')) 'The report has one meaningful browser-tab title'
Assert-True ($html.Contains('<meta name="viewport" content="width=device-width, initial-scale=1" />')) 'The report declares its viewport'
foreach($filterIndex in 1..3) {
    Assert-True ($html.Contains("<label for=`"filterDropdown$filterIndex`">")) "Filter $filterIndex has its associated label"
}
Assert-True ($html.Contains('id="searchInput" aria-label="Search timeline"')) 'Timeline search has an accessible name'
Assert-True ($html.Contains('<button type="button" id="clearSearch" aria-label="Clear search" title="Clear search"')) 'Clearing the search has an accessible name and no submit behavior'
Assert-True ($html.Contains('<button type="button" id="resetFilter"')) 'Resetting filters has no submit behavior'
Assert-True ([regex]::IsMatch($html, '(?s)\.filter-row\s*\{[^}]*flex-wrap:\s*wrap;[^}]*row-gap:\s*12px;')) 'Filter groups wrap with row spacing'
Assert-True ([regex]::IsMatch($html, '(?s)table tr:hover\s*\{[^}]*background-color:\s*#0c3565;')) 'Hover uses a dark background'
Assert-True ($html.Contains("'Red': '#ff9e9e'")) 'Error text uses the readable light red'
Assert-True ($html.Contains('rgba(1, 36, 86, 1)') -and $html.Contains('rgba(12, 53, 101, 1)')) 'Grouped rows use both dark backgrounds'
Assert-True ($html.Contains('Device Preparation mode (APv2) detected') -and $html.Contains('Device Preparation mode (APv2) disabled')) 'Synthetic timeline events survive HTML generation'

[PSCustomObject]@{
    Result = 'PASS'
    Assertions = $script:assertions
    PowerShell = $PSVersionTable.PSVersion.ToString()
    ReportPath = $reports[0].FullName
}
