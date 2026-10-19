[CmdletBinding()]
param(
    [string]$ScriptPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'Get-IntuneManagementExtensionDiagnostics.ps1'),
    [string]$SampleZip,
    [Parameter(Mandatory=$true)]
    [string]$WorkPath
)

$ErrorActionPreference = 'Stop'
$ScriptPath = (Get-Item -LiteralPath $ScriptPath).FullName
$null = New-Item -ItemType Directory -Path $WorkPath
$WorkPath = (Get-Item -LiteralPath $WorkPath).FullName
if($SampleZip) {
    $SampleZip = (Get-Item -LiteralPath $SampleZip).FullName
    $sampleCopy = Join-Path $WorkPath 'sample.zip'
    Copy-Item -LiteralPath $SampleZip -Destination $sampleCopy
    $SampleZip = $sampleCopy
}
$script:assertions = 0

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if(-not $Condition) { throw "FAIL: $Message" }
    $script:assertions++
}

function Invoke-Report {
    param([string]$Name, [hashtable]$Arguments)
    $output = Join-Path $WorkPath $Name
    $null = New-Item -ItemType Directory -Path $output
    & $ScriptPath @Arguments -ExportHTMLReportPath $output -DoNotOpenReportAutomatically *> (Join-Path $output 'console.txt')
    $jsonFiles = @(Get-ChildItem -LiteralPath $output -Filter '*.diagnostics.json')
    $htmlFiles = @(Get-ChildItem -LiteralPath $output -Filter '*.html')
    Assert-True ($jsonFiles.Count -eq 1 -and $htmlFiles.Count -eq 1) "$Name produces HTML and diagnostics"
    $result = Get-Content -LiteralPath $jsonFiles[0].FullName -Raw | ConvertFrom-Json
    foreach($file in $result.Files | Where-Object { $_.Status -notin @('Ignored', 'Not selected') }) {
        Assert-True ($file.Lines -eq $file.ParsedLines + $file.RejectedLines + $file.BlankLines) "$Name accounts for every physical line in $($file.File)"
    }
    $html = Get-Content -LiteralPath $htmlFiles[0].FullName -Raw
    Assert-True ($html.Contains('Input and analysis coverage')) "$Name includes input coverage in HTML"
    $result
}

function Assert-Failure {
    param([string]$Name, [hashtable]$Arguments, [string]$Pattern)
    $failedAsExpected = $false
    try {
        & $ScriptPath @Arguments -DoNotOpenReportAutomatically *> (Join-Path $WorkPath "$Name.txt")
    } catch {
        $failedAsExpected = $_.Exception.Message -match $Pattern
        if(-not $failedAsExpected) { throw }
    }
    Assert-True $failedAsExpected "$Name terminates with an explicit reason"
}

function New-CMTraceRecord {
    param(
        [string]$Message = 'Fixture record',
        [string]$Time = '01:02:03.1234567',
        [string]$Date = '10-12-2026',
        [string]$Component = 'Fixture',
        [string]$Thread = '1'
    )
    '<![LOG[' + $Message + ']LOG]!><time="' + $Time + '" date="' + $Date + '" component="' + $Component + '" context="" type="1" thread="' + $Thread + '" file="">'
}

function New-RemediationEntry {
    param(
        [string]$File = 'IntuneManagementExtension.log',
        [string]$Time = '01:00:00.0000000',
        [string]$Thread = '4',
        [string]$Action = 'detect',
        [string]$Message
    )
    if(-not $Message) {
        $Message = '"C:\Program Files (x86)\Microsoft Intune Management Extension\agentexecutor.exe"  -remediationScript  ""C:\Windows\IMECache\HealthScripts\11111111-1111-1111-1111-111111111111_1\' + $Action + '.ps1""'
    }
    [PSCustomObject]@{ File = $File; Time = $Time; Thread = $Thread; Message = $Message }
}

function Invoke-RemediationScenario {
    param([string]$Name, [object[]]$Entries)

    $folder = Join-Path $WorkPath $Name
    $null = New-Item -ItemType Directory -Path $folder
    foreach($group in $Entries | Group-Object File) {
        $records = foreach($entry in $group.Group) {
            New-CMTraceRecord -Message $entry.Message -Time $entry.Time -Thread $entry.Thread -Component ([System.IO.Path]::GetFileNameWithoutExtension($entry.File))
        }
        [System.IO.File]::WriteAllLines((Join-Path $folder $group.Name), [string[]]$records, [System.Text.UTF8Encoding]::new($false))
    }
    & {
        . $ScriptPath -InputPath $folder -ExportHTMLReportPath $folder -DoNotOpenReportAutomatically *> (Join-Path $folder 'console.txt')
        $htmlPath = @(Get-ChildItem -LiteralPath $folder -Filter '*.html')[0].FullName
        [PSCustomObject]@{
            Timeline = @($observedTimeline | Where-Object Type -in @('Remediation', 'Custom Compliance'))
            Logs = @($LogEntryList)
            Html = [System.IO.File]::ReadAllText($htmlPath)
        }
    }
}

function Invoke-StartupScenario {
    param([string]$Name, [string]$UserName, [string]$MockFolder, [string]$ExternalInput)

    $oldUserName = $env:USERNAME
    $oldComputerName = $env:COMPUTERNAME
    $output = Join-Path $WorkPath $Name
    $null = New-Item -ItemType Directory -Path $output
    $script:localProbeCount = 0
    try {
        $env:USERNAME = $UserName
        $env:COMPUTERNAME = 'ESP-FIXTURE'
        & {
            # Route legacy local paths to test data; never inspect this workstation's IME logs.
            function Test-Path {
                [CmdletBinding()]
                param([string]$Path, [string]$LiteralPath, [string]$PathType)
                $target = if($LiteralPath) { $LiteralPath } else { $Path }
                if($target -like 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs*') {
                    $script:localProbeCount++
                    return $true
                }
                Microsoft.PowerShell.Management\Test-Path @PSBoundParameters
            }
            function Get-Item {
                [CmdletBinding()]
                param([string]$Path, [string]$LiteralPath)
                $target = if($LiteralPath) { $LiteralPath } else { $Path }
                if($target -eq 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs') {
                    $script:localProbeCount++
                    return Microsoft.PowerShell.Management\Get-Item -LiteralPath $MockFolder
                }
                Microsoft.PowerShell.Management\Get-Item @PSBoundParameters
            }
            $arguments = @{ ExportHTMLReportPath = $output }
            if($ExternalInput) { $arguments.InputPath = $ExternalInput }
            if($UserName -ne 'defaultuser0' -or $ExternalInput) { $arguments.DoNotOpenReportAutomatically = $true }
            . $ScriptPath @arguments *> (Join-Path $output 'console.txt')
            [PSCustomObject]@{
                Local = $UsingLocalLogs
                LocalProbes = $script:localProbeCount
                Source = $InputSourcePath
                NoBrowser = [bool]$DoNotOpenReportAutomatically
                NoCleartextRemediations = [bool]$DoNotDownloadClearTextRemediationScriptsToReport
                ShowAllEvents = [bool]$ShowAllTimelineEvents
                AllFiles = [bool]$AllLogFiles
                AllEntries = [bool]$AllLogEntries
                Timeline = @($observedTimeline)
            }
        }
    } finally {
        $env:USERNAME = $oldUserName
        $env:COMPUTERNAME = $oldComputerName
    }
}

$tokens = $null
$parseErrors = $null
$scriptAst = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$parseErrors)
Assert-True ($parseErrors.Count -eq 0) 'PowerShell syntax is valid'
$invalidVerboseCalls = $scriptAst.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.CommandAst] -and
    $node.GetCommandName() -eq 'Write-Verbose' -and
    @($node.CommandElements | Where-Object {
        $_ -is [System.Management.Automation.Language.CommandParameterAst] -and $_.ParameterName -eq 'ForegroundColor'
    }).Count -gt 0
}, $true)
Assert-True ($invalidVerboseCalls.Count -eq 0) 'Issue #13: active Write-Verbose calls do not use ForegroundColor'
foreach($functionName in @(
    'Get-IntunePowershellScriptContentInCleartext',
    'Get-IntuneRemediationDetectionScriptContentInCleartext',
    'Get-IntuneRemediationRemediateScriptContentInCleartext',
    'Get-IntuneCustomComplianceScriptContentInCleartext'
)) {
    $functionAst = $scriptAst.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $functionName
    }, $true)[0]
    $result = & {
        . ([scriptblock]::Create($functionAst.Extent.Text))
        function Invoke-MgGraphRequestGetAllPages { param($Uri) [PSCustomObject]@{ id = 'fixture' } }
        $Online = $true
        $DoNotDownloadClearTextRemediationScriptsToReport = $false
        $IdHashtable = @{ fixture = [PSCustomObject]@{ id = 'fixture' } }
        & $functionName 'fixture' -Verbose 4>&1
    }
    Assert-True (@($result | Where-Object { $_ -is [System.Management.Automation.VerboseRecord] }).Count -eq 1) "Issue #13: $functionName emits its verbose diagnostic without a parameter error"
    Assert-True ($result -contains 'N/A') "Issue #13: $functionName preserves its missing-content result"
}
$command = Get-Command -Name $ScriptPath
Assert-True ($command.Parameters.ContainsKey('LogFile')) 'LogFile remains an original parameter, not an alias for the new option'
Assert-True ($command.Parameters['LogFile'].Aliases -contains 'FullName') 'The original FullName pipeline alias is preserved'
Assert-True (-not [bool]($command.Parameters['InputPath'].Attributes | Where-Object { $_ -is [System.Management.Automation.ParameterAttribute] -and $_.Mandatory })) 'InputPath is strictly optional'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zipResult = $null
if($SampleZip) {
    $sampleFolder = Join-Path $WorkPath 'sample'
    [System.IO.Compression.ZipFile]::ExtractToDirectory($SampleZip, $sampleFolder)
    $zipResult = Invoke-Report 'sample-zip' @{ InputPath = $SampleZip }
    $folderResult = Invoke-Report 'sample-folder' @{ LogFilesFolder = $sampleFolder }
    Assert-True ($zipResult.Files.Count -eq 11) 'All eleven supplied log files are inventoried'
    Assert-True ($zipResult.ParsedRecords -eq 38891) 'The sample contains 38,891 complete CMTrace records'
    Assert-True ($zipResult.ParsedLines -eq 42436 -and $zipResult.RejectedLines -eq 3) 'The sample has 42,436 parsed lines and three explicit rejections'
    Assert-True ($zipResult.Files.Status -notcontains 'Ignored') 'No supplied .log file is silently excluded'
    Assert-True ($zipResult.TimelineEvents -eq 34) 'The original 34 sample events are preserved without duplicate-remediation warnings'
    $zipInventory = $zipResult.Files | ConvertTo-Json -Depth 5 -Compress
    $folderInventory = $folderResult.Files | ConvertTo-Json -Depth 5 -Compress
    Assert-True ($zipInventory -ceq $folderInventory) 'Folder and ZIP have identical per-file diagnostics'
    Assert-True (($zipResult.Issues | ConvertTo-Json -Compress) -ceq ($folderResult.Issues | ConvertTo-Json -Compress)) 'Folder and ZIP have identical rejection reasons'
    Assert-True ($zipResult.TimelineEvents -eq $folderResult.TimelineEvents) 'Folder and ZIP produce the same number of timeline events'
    $zipHtmlPath = @(Get-ChildItem -LiteralPath (Join-Path $WorkPath 'sample-zip') -Filter '*.html')[0].FullName
    $folderHtmlPath = @(Get-ChildItem -LiteralPath (Join-Path $WorkPath 'sample-folder') -Filter '*.html')[0].FullName
    $timelinePattern = '(?s)<table id="ObservedTimeline">.*?</table>'
    $zipTimeline = [regex]::Match([System.IO.File]::ReadAllText($zipHtmlPath), $timelinePattern).Value
    $folderTimeline = [regex]::Match([System.IO.File]::ReadAllText($folderHtmlPath), $timelinePattern).Value
    Assert-True ($zipTimeline.Length -gt 0 -and $zipTimeline -ceq $folderTimeline) 'Folder and ZIP produce identical timeline tables'
    $lock = [System.IO.File]::Open($SampleZip, 'Open', 'Read', 'None')
    $lock.Dispose()
    Assert-True $true 'The ZIP handle is released after analysis'
}

foreach($offset in @('00.0000000', '00.0010000', '01.0000000')) {
    $caseName = 'remediation-mirror-' + $offset.Replace('.', '-')
    $case = Invoke-RemediationScenario $caseName @(
        (New-RemediationEntry),
        (New-RemediationEntry -File '_HealthScripts-20261012-010000.log' -Time "01:00:$offset"),
        (New-RemediationEntry -File '_HealthScripts-20261012-010000.log' -Time '01:00:05.0000000' -Message 'Powershell execution is done, exitCode = 0')
    )
    Assert-True ($case.Timeline.Count -eq 1 -and $case.Timeline[0].Status -eq 'Success') "Issue #14: mirrored start at $offset produces one result without a false warning"
    Assert-True ($case.Timeline[0].Date -eq '2026-10-12 01:00:05.0000000' -and $case.Timeline[0].Seconds -eq 5) "Issue #14: $offset uses the end timestamp and original duration"
    Assert-True ($case.Timeline[0].LogEntry -eq 'Line 2') "Issue #14: $offset references the completion line"
    $timelineHtml = [regex]::Match($case.Html, '(?s)<table id="ObservedTimeline">.*?</table>').Value
    Assert-True ($timelineHtml.Contains('2026-10-12 01:00:05.0000000') -and -not $timelineHtml.Contains('did not detect process end')) "Issue #14: $offset renders the corrected result in HTML"
}

$restart = Invoke-RemediationScenario 'remediation-same-source-restart' @(
    (New-RemediationEntry),
    (New-RemediationEntry -Time '01:00:00.0010000'),
    (New-RemediationEntry -Time '01:00:05.0000000' -Message 'Powershell execution is done, exitCode = 0')
)
Assert-True (@($restart.Timeline | Where-Object Status -eq 'Warning').Count -eq 1) 'Issue #14: a same-source restart is not collapsed into a duplicate'
Assert-True (@($restart.Timeline | Where-Object Status -eq 'Success').Count -eq 1) 'Issue #14: the restarted execution still completes'

$agentEntries = @(
    (New-RemediationEntry -File 'AgentExecutor.log' -Thread '1' -Time '01:00:00.1000000' -Message 'cmd line for running powershell is -NoProfile -executionPolicy bypass -file  "C:\Windows\IMECache\HealthScripts\11111111-1111-1111-1111-111111111111_1\detect.ps1"'),
    (New-RemediationEntry -File 'AgentExecutor.log' -Thread '1' -Time '01:00:00.2000000' -Message 'Powershell exit code is 0'),
    (New-RemediationEntry -File 'AgentExecutor.log' -Thread '1' -Time '01:00:00.3000000' -Message 'Powershell script is successfully executed.')
)
foreach($output in @(
    @{ Name = 'missing'; Text = $null },
    @{ Name = 'single-line'; Text = 'write output done. output = fixture output' },
    @{ Name = 'multiline'; Text = "write output done. output = fixture output`nsecond output line`n" }
)) {
    $entries = @((New-RemediationEntry)) + $agentEntries
    if($null -ne $output.Text) {
        $entries += New-RemediationEntry -File 'AgentExecutor.log' -Thread '1' -Time '01:00:00.4000000' -Message $output.Text
    }
    $restartAfterOutput = Invoke-RemediationScenario "remediation-output-$($output.Name)-restart" ($entries + @(
        (New-RemediationEntry -Time '01:00:01.0000000'),
        (New-RemediationEntry -Time '01:00:06.0000000' -Message 'Powershell execution is done, exitCode = 0')
    ))
    $successes = @($restartAfterOutput.Timeline | Where-Object Status -eq 'Success')
    Assert-True (@($restartAfterOutput.Timeline | Where-Object Status -eq 'Warning').Count -eq 1) "Issue #14: $($output.Name) output cannot hide an interrupted execution"
    Assert-True ($successes.Count -eq 1) "Issue #14: $($output.Name) output cannot reuse one completion for two executions"
    Assert-True ($successes[0].Date -eq '2026-10-12 01:00:06.0000000' -and $successes[0].Seconds -eq 5 -and $successes[0].LogEntry -eq 'Line 3') "Issue #14: $($output.Name) output preserves the restarted execution's completion and duration"
    Assert-True ($restartAfterOutput.Html.Contains('Remediation started but we did not detect process end')) "Issue #14: $($output.Name) output retains the missing-end warning in HTML"

    $completedOutput = Invoke-RemediationScenario "remediation-output-$($output.Name)-completion" ($entries + @(
        (New-RemediationEntry -File 'AgentExecutor.log' -Thread '1' -Time '01:00:00.5000000' -Message 'Unrelated AgentExecutor activity'),
        (New-RemediationEntry -Time '01:00:05.0000000' -Message 'Powershell execution is done, exitCode = 0')
    ))
    Assert-True ($completedOutput.Timeline.Count -eq 1 -and $completedOutput.Timeline[0].Status -eq 'Success') "Issue #14: $($output.Name) output preserves a valid completion without a restart"
    Assert-True ($completedOutput.Timeline[0].Date -eq '2026-10-12 01:00:05.0000000' -and $completedOutput.Timeline[0].Seconds -eq 5) "Issue #14: $($output.Name) output preserves the actual end timestamp and duration"
    if($null -ne $output.Text) {
        Assert-True ($completedOutput.Timeline[0].DetailToolTip.Contains('fixture output')) "Issue #14: $($output.Name) output is retained in the tooltip"
        Assert-True (-not $completedOutput.Timeline[0].DetailToolTip.Contains('Unrelated AgentExecutor activity')) "Issue #14: $($output.Name) output does not consume another record"
    }
    if($output.Name -eq 'multiline') {
        Assert-True ($completedOutput.Timeline[0].DetailToolTip.Contains('second output line')) 'Issue #14: multiline output retains its continuation lines'
    }
}
$outputAtEof = Invoke-RemediationScenario 'remediation-agent-output-eof' (@((New-RemediationEntry)) + $agentEntries)
Assert-True (@($outputAtEof.Timeline | Where-Object Status -eq 'Warning').Count -eq 1 -and @($outputAtEof.Timeline | Where-Object Status -eq 'Success').Count -eq 0) 'Issue #14: AgentExecutor success at EOF cannot invent an IME completion'

$outsideWindow = Invoke-RemediationScenario 'remediation-outside-mirror-window' @(
    (New-RemediationEntry),
    (New-RemediationEntry -File 'HealthScripts.log' -Time '01:00:01.0010000'),
    (New-RemediationEntry -File 'HealthScripts.log' -Time '01:00:05.0000000' -Message 'Powershell execution is done, exitCode = 0')
)
Assert-True (@($outsideWindow.Timeline | Where-Object Status -eq 'Warning').Count -eq 1) 'Issue #14: starts over one second apart are not assumed to be mirrors'

$repeated = Invoke-RemediationScenario 'remediation-repeated-runs' @(
    (New-RemediationEntry),
    (New-RemediationEntry -Time '01:00:05.0000000' -Message 'Powershell execution is done, exitCode = 0'),
    (New-RemediationEntry -Time '01:00:06.0000000'),
    (New-RemediationEntry -Time '01:00:11.0000000' -Message 'Powershell execution is done, exitCode = 0')
)
Assert-True ($repeated.Timeline.Count -eq 2 -and @($repeated.Timeline | Where-Object Status -eq 'Success').Count -eq 2) 'Issue #14: successive completed runs stay distinct'

$parallel = Invoke-RemediationScenario 'remediation-parallel-threads' @(
    (New-RemediationEntry),
    (New-RemediationEntry -File 'HealthScripts.log' -Time '01:00:00.0010000' -Thread '7'),
    (New-RemediationEntry -Time '01:00:05.0000000' -Message 'Powershell execution is done, exitCode = 0'),
    (New-RemediationEntry -File 'HealthScripts.log' -Time '01:00:06.0000000' -Thread '7' -Message 'Powershell execution is done, exitCode = 0')
)
Assert-True ($parallel.Timeline.Count -eq 2 -and @($parallel.Timeline | Where-Object Status -eq 'Success').Count -eq 2) 'Issue #14: parallel threads are neither deduplicated nor mistaken for interrupted runs'

$differentCommand = Invoke-RemediationScenario 'remediation-different-command' @(
    (New-RemediationEntry),
    (New-RemediationEntry -File 'HealthScripts.log' -Time '01:00:00.0010000' -Action 'remediate'),
    (New-RemediationEntry -File 'HealthScripts.log' -Time '01:00:05.0000000' -Message 'Powershell execution is done, exitCode = 1')
)
Assert-True (@($differentCommand.Timeline | Where-Object Status -eq 'Warning').Count -eq 1) 'Issue #14: detect and remediate commands are not merged'
Assert-True (@($differentCommand.Timeline | Where-Object { $_.Status -eq 'Failed' -and $_.Intent -eq 'Remediate' -and $_.Date -eq '2026-10-12 01:00:05.0000000' }).Count -eq 1) 'Issue #14: a failed remediation is reported at completion as Failed'

$notDetected = Invoke-RemediationScenario 'remediation-not-detected' @(
    (New-RemediationEntry),
    (New-RemediationEntry -Time '01:01:05.0000000' -Message 'Powershell execution is done, exitCode = 1')
)
Assert-True ($notDetected.Timeline.Count -eq 1 -and $notDetected.Timeline[0].Status -eq 'Not Detected' -and $notDetected.Timeline[0].Date -eq '2026-10-12 01:01:05.0000000' -and $notDetected.Timeline[0].Seconds -eq 65) 'Issue #14: negative detection retains its status, completion time and duration over one minute'
Assert-True ([bool]($notDetected.Logs.ProcessRunTime -match '65 seconds')) 'Issue #14: the log viewer also shows the full duration'

$unfinished = Invoke-RemediationScenario 'remediation-missing-end' @(
    (New-RemediationEntry),
    (New-RemediationEntry -File 'HealthScripts.log' -Time '01:00:00.0010000')
)
Assert-True ($unfinished.Timeline.Count -eq 2 -and @($unfinished.Timeline | Where-Object Status -eq 'Warning').Count -eq 1) 'Issue #14: a genuinely missing end at EOF still produces one warning'

$fixtures = Join-Path $WorkPath '[fixtures & logs]'
$deep = Join-Path $fixtures 'a\b\c\d\e'
$null = New-Item -ItemType Directory -Path $deep
$utf8 = [System.Text.UTF8Encoding]::new($false)
$logPath = Join-Path $deep '_IntuneManagementExtension.log'
[System.IO.File]::WriteAllLines($logPath, [string[]]@(
    (New-CMTraceRecord -Time '01:02:03.1'),
    (New-CMTraceRecord -Time '01:02:03.12'),
    (New-CMTraceRecord -Time '01:02:03.1234567'),
    (New-CMTraceRecord -Time '01:02:04'),
    (New-CMTraceRecord -Time '01:02:05.500-120')
), $utf8)
$malformed = @(
    (New-CMTraceRecord),
    'unframed text',
    '',
    '<![LOG[Multiline start',
    'middle',
    ('end]LOG]!><time="01:03:00.12" date="10-12-2026" component="Fixture" context="" type="1" thread="1" file="">'),
    (New-CMTraceRecord -Date '2-30-2026'),
    '<![LOG[Interrupted record',
    (New-CMTraceRecord -Time '01:04:00'),
    '<![LOG[Unfinished at EOF'
)
[System.IO.File]::WriteAllLines((Join-Path $fixtures 'malformed.log'), [string[]]$malformed, $utf8)
[System.IO.File]::WriteAllText((Join-Path $fixtures 'agentexecutor.log'), ((New-CMTraceRecord) + "`n" + ']LOG]!><time="23:59:59.123" date="10-12-2026">'), $utf8)
[System.IO.File]::WriteAllText((Join-Path $fixtures 'healthscripts.log'), (New-CMTraceRecord), [System.Text.Encoding]::Unicode)
[System.IO.File]::WriteAllText((Join-Path $fixtures 'appactionprocessor.lo_'), '<![LOG[Reordered attributes]LOG]!><date="10-12-2026" thread="1" component="Fixture" time="01:02:03.12" extra="metadata">', $utf8)
[System.IO.File]::WriteAllText((Join-Path $fixtures 'empty.log'), '', $utf8)
[System.IO.File]::WriteAllText((Join-Path $fixtures 'notes.txt'), 'Not an IME log.', $utf8)
$fixtureResult = Invoke-Report 'fixtures' @{ InputPath = $fixtures }
$fixtureZip = Join-Path $WorkPath 'fixtures.zip'
[System.IO.Compression.ZipFile]::CreateFromDirectory($fixtures, $fixtureZip)
$fixtureZipResult = Invoke-Report 'fixtures-zip' @{ InputPath = $fixtureZip }
Assert-True (($fixtureResult.Files | ConvertTo-Json -Depth 5 -Compress) -ceq ($fixtureZipResult.Files | ConvertTo-Json -Depth 5 -Compress)) 'Synthetic folder and ZIP have identical per-file diagnostics'
Assert-True ($fixtureResult.TimelineEvents -eq $fixtureZipResult.TimelineEvents) 'Synthetic folder and ZIP have identical timeline-event counts'
$malformedResult = @($fixtureResult.Files | Where-Object File -eq 'malformed.log')[0]
Assert-True ($malformedResult.Records -eq 3 -and $malformedResult.ParsedLines -eq 5) 'Multiline records preserve all physical message lines'
Assert-True ($malformedResult.RejectedLines -eq 4 -and $malformedResult.BlankLines -eq 1) 'Invalid dates, broken framing and blank lines are distinguished'
Assert-True (@($fixtureResult.Issues | Where-Object File -eq 'malformed.log').Count -eq 4) 'Every malformed record has a specific reason'
$deepResult = @($fixtureResult.Files | Where-Object File -like '*_IntuneManagementExtension.log')[0]
Assert-True ($deepResult.Records -eq 5) 'Recursion has no two-level limit; timestamp fractions and CMTrace offsets are supported'
Assert-True ($deepResult.Coverage -like 'IME timeline rules*') 'Underscore-prefixed rotated IME logs have timeline coverage'
Assert-True (@($fixtureResult.Files | Where-Object File -eq 'healthscripts.log')[0].Records -eq 1) 'UTF-16 BOM is detected'
Assert-True (@($fixtureResult.Files | Where-Object File -eq 'appactionprocessor.lo_')[0].Records -eq 1) 'Rotated .lo_ logs and reordered attributes are supported'
Assert-True (@($fixtureResult.Files | Where-Object File -eq 'notes.txt')[0].Status -eq 'Ignored') 'Non-log files have an explicit ignored status'
Assert-True (@($fixtureResult.Files | Where-Object File -eq 'empty.log')[0].Status -eq 'No usable records') 'Empty logs have an explicit status'
$orphanResult = @($fixtureResult.Files | Where-Object File -eq 'agentexecutor.log')[0]
Assert-True ($orphanResult.Records -eq 1 -and $orphanResult.RejectedLines -eq 1) 'Orphan closing markers are rejected rather than attached to earlier records'
Assert-True (([datetime]$orphanResult.LastTimestamp).Hour -eq 1) 'An orphan closing marker cannot overwrite a valid timestamp'
$fixtureHtml = [System.IO.File]::ReadAllText(@(Get-ChildItem -LiteralPath (Join-Path $WorkPath 'fixtures') -Filter '*.html')[0].FullName)
Assert-True ($fixtureHtml.Contains('[fixtures &amp; logs]')) 'Input paths are HTML-encoded'

$singleResult = Invoke-Report 'single-log' @{ LogFile = $logPath }
Assert-True ($singleResult.Files.Count -eq 1 -and $singleResult.ParsedRecords -eq 5) 'The original LogFile parameter still works with literal paths'
$start = [datetime]::new(2026, 10, 12, 1, 2, 3)
$end = [datetime]::new(2026, 10, 12, 1, 2, 4)
$rangeResult = Invoke-Report 'time-range' @{ InputPath = $logPath; LogStartDateTime = $start; LogEndDateTime = $end }
Assert-True ($rangeResult.LinesInTimeRange -eq 3) 'Date filtering is inclusive at the start and exclusive at the end'
$todayResult = Invoke-Report 'today' @{ InputPath = $logPath; Today = $true }
Assert-True (([datetime]$todayResult.TimeRangeStart).Date -eq (Get-Date).Date) 'Today is not overridden by AllLogEntries'

$espFolder = Join-Path $WorkPath 'esp-fixture'
$null = New-Item -ItemType Directory -Path $espFolder
$espLog = Join-Path $espFolder 'IntuneManagementExtension.log'
$espMessages = @(
    'EMS Agent Started',
    '[Win32App] CheckDeviceOnlyFromFirstSyncReg set skipDeviceStatusPage: False',
    '[Win32App] CheckDeviceOnlyFromFirstSyncReg set skipUserStatusPage: False',
    '[Win32App] In EspPhase: DeviceSetup',
    '[Win32App] In EspPhase: AccountSetup. All Win32App has been added. Set SidecarReady: True',
    '[Win32App] The EspPhase: NotInEsp in session',
    '[APv2] Device is in APv2 mode: True.',
    '[APv2] Successfully disabled APv2 mode.'
)
$espRecords = for($i = 0; $i -lt $espMessages.Count; $i++) {
    New-CMTraceRecord -Message $espMessages[$i] -Time ('01:{0:00}:00.123' -f $i)
}
[System.IO.File]::WriteAllLines($espLog, [string[]]$espRecords, $utf8)
$espStartup = Invoke-StartupScenario 'esp-startup' 'defaultuser0' $espFolder
Assert-True ($espStartup.Local -and $espStartup.LocalProbes -gt 0) 'ESP startup without arguments retains the legacy local source'
Assert-True ($espStartup.NoBrowser -and $espStartup.NoCleartextRemediations) 'ESP startup disables browser opening and cleartext remediation downloads'
Assert-True ($espStartup.ShowAllEvents -and $espStartup.AllFiles -and $espStartup.AllEntries) 'ESP startup retains all-files, all-entries and detailed-timeline settings'
Assert-True ([bool]($espStartup.Timeline.Detail -match 'Device ESP phase')) 'Device ESP detection is preserved'
Assert-True ([bool]($espStartup.Timeline.Detail -match 'User ESP finished')) 'User ESP completion detection is preserved'
Assert-True ([bool]($espStartup.Timeline.Detail -match 'Not in ESP detected')) 'Leaving ESP is still detected'
Assert-True ([bool]($espStartup.Timeline.Detail -match 'Device Preparation mode \(APv2\) detected')) 'Autopilot device preparation detection is preserved'
Assert-True ([bool]($espStartup.Timeline.Detail -match 'Device Preparation mode \(APv2\) disabled')) 'Autopilot device preparation completion is preserved'
$espZip = Join-Path $WorkPath 'esp-fixture.zip'
[System.IO.Compression.ZipFile]::CreateFromDirectory($espFolder, $espZip)
$externalEsp = Invoke-StartupScenario 'esp-external' 'defaultuser0' $espFolder $espZip
Assert-True (-not $externalEsp.Local -and $externalEsp.LocalProbes -eq 0) 'Explicit external input is never overridden in ESP'
Assert-True ($externalEsp.Source -eq $espZip) 'ESP offline analysis retains the supplied ZIP source'
$localStartup = Invoke-StartupScenario 'local-startup' 'FixtureUser' $espFolder
Assert-True ($localStartup.Local -and $localStartup.LocalProbes -gt 0) 'Legacy local startup outside ESP remains available'

$unsupported = Join-Path $WorkPath 'unsupported'
$null = New-Item -ItemType Directory -Path $unsupported
[System.IO.File]::WriteAllText((Join-Path $unsupported 'notes.txt'), 'Text.', $utf8)
$badZip = Join-Path $WorkPath 'corrupt.zip'
[System.IO.File]::WriteAllText($badZip, 'Not a ZIP archive.', $utf8)
Assert-Failure 'conflicting-inputs' @{ InputPath = $fixtures; LogFilesFolder = $fixtures } 'not both'
Assert-Failure 'conflicting-file-input' @{ InputPath = $fixtures; LogFile = $logPath } 'not both'
Assert-Failure 'empty-explicit-input' @{ InputPath = '' } 'InputPath'
Assert-Failure 'missing-path' @{ InputPath = (Join-Path $WorkPath 'missing') } 'does not exist|Cannot find'
Assert-Failure 'unsupported-file' @{ InputPath = (Join-Path $unsupported 'notes.txt') } 'Unsupported input'
Assert-Failure 'no-logs' @{ InputPath = $unsupported } 'No .log or .lo_'
Assert-Failure 'empty-log' @{ InputPath = (Join-Path $fixtures 'empty.log') } 'No usable CMTrace'
Assert-Failure 'invalid-zip' @{ InputPath = $badZip } 'central directory|archive|ZIP'
Assert-Failure 'invalid-output' @{ InputPath = $fixtures; ExportHTMLReportPath = (Join-Path $WorkPath 'missing-output') } 'Output directory does not exist'
Assert-Failure 'invalid-start' @{ InputPath = $fixtures; LogStartDateTime = 'not-a-date' } 'not a valid date'
Assert-Failure 'reversed-range' @{ InputPath = $fixtures; LogStartDateTime = $end; LogEndDateTime = $start } 'earlier than'
$lock = [System.IO.File]::Open($badZip, 'Open', 'Read', 'None')
$lock.Dispose()
Assert-True $true 'A failed archive open does not leave the ZIP locked'

[PSCustomObject]@{
    Result = 'PASS'
    Assertions = $script:assertions
    PowerShell = $PSVersionTable.PSVersion.ToString()
    WorkPath = $WorkPath
    SampleRecords = $zipResult.ParsedRecords
    SampleRejectedLines = $zipResult.RejectedLines
    SampleTimelineEvents = $zipResult.TimelineEvents
}
