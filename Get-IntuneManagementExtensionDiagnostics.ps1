<#PSScriptInfo

.VERSION 1.1

.GUID ab5a8b63-97d5-4b1a-a4ab-6dcedbab1eb7

.AUTHOR Petri.Paavola@yodamiitti.fi

.COMPANYNAME Yodamiitti Oy

.COPYRIGHT Petri.Paavola@yodamiitti.fi

.TAGS Intune Windows Autopilot

.LICENSEURI

.PROJECTURI https://github.com/petripaavola/Get-IntuneManagementExtensionDiagnostics

.ICONURI

.EXTERNALMODULEDEPENDENCIES

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES
Version 1.0:  Original published version
Version 1.1:  Win32App and WinGetApp Required/Available and Install/Uninstall intent is detected right
              Win32App Supersedence should be recognized (first uninstall and then install)
              Win32App failed (un)install process is detected
              Win32App Download Statistics table added
			  Added export to text files
#>

<#
.Synopsis
   This script analyzes Microsoft Intune Management Extension (IME) log(s) and creates timeline report from found actions.

.DESCRIPTION
   This script analyzes Microsoft Intune Management Extension (IME) log(s) and creates timeline report from found actions.
   
   Timeline report includes information about Intune Win32App, WinGetApp, Powershell scripts, Proactive Remedation scripts and custom Compliance Policy scripts events. Windows Autopilot ESP phases are also shown on timeline.
   
   Script also includes really capable Log Viewer UI if scripts is started with parameter -ShowLogViewerUI

   LogViewerUI (Out-GridView) looks a lot like cmtrace.exe tool but it is better because all found log actions are added to log for easier debugging.
   
   LogViewerUI has good search and filtering capabilities. Try to filter known log entries in Timeline: Add criteria -> ProcessRunTime -> is not empty.
   
   Selecting last line (RELOAD) and OK will reload log file.
   
   Script can merge multiple log files so especially in LogViewerUI you can see Powershell command outputs from AgentExecutor.log
   
   Powershell command outputs and errors can be also shown in Timeline view with parameters -ShowStdOutInTimeline and -ShowErrorsInTimeline
   This shows instantly what is possible problem in Powershell scripts.


   Possible Microsoft 365 App and MSI Line-of-Business Apps (maybe change to Win32App ;) installations are not seen by this report because they are not installed with Intune Management Agent.


   Author:
   Petri.Paavola@yodamiitti.fi
   Senior Modern Management Principal
   Microsoft MVP - Windows and Devices for IT
   
   2023-03-12

   https://github.com/petripaavola/Get-IntuneManagementExtensionDiagnostics

.PARAMETER Online
Download Powershell, Proactive Remediation and custom Compliance policy scripts to get displayName to Timeline report

.PARAMETER LogFile
Specify log file fullpath

.PARAMETER LogFilesFolder
Specify folder where to check log files. Will show UI where you can select what logs to process

.PARAMETER LogStartDateTime
Speficy date and time to start log entries. For example -

.PARAMETER LogEndDateTime
Speficy date and time to stop log entries

.PARAMETER ShowLogViewerUI
Shows graphical LogViewerUI where all log entries are easily browsed, searched and filtered in graphical UI

.PARAMETER $LogViewerUI
Shows graphical LogViewerUI where all log entries are easily browsed, searched and filtered in graphical UI

.PARAMETER AllLogEntries
Process all found log entries.
Selecting this parameter will disable UI which asks date/time/hour selection for logs (use for silent commands or scripts)

.PARAMETER AllLogFiles
Process all found supported log file(s) automatically. This includes *AgentExecutor*.log and *IntuneManagementExtension*.log
Selecting this parameter will disable UI which asks which log files to process (use for silent commands or scripts)

.PARAMETER Today
Show log entries from today (from midnight)

.PARAMETER ShowAllTimelineEntries
Shows more entries in Timeline. This option will show starting messages for events which are not shown by default

.PARAMETER ShowStdOutInTimeline
Show script StdOut in events. This shows for example what Proactive Remediations will return back to Intune

.PARAMETER ShowErrorsInTimeline
This will show found error messages from Powershell scripts. Note that Powershell script may succeed and still have errors shown here.

.PARAMETER ShowErrorsSummary
Show separate all errors summary after Timeline.

.PARAMETER ConvertAllKnownGuidsToClearText
This parameter replaces all known GUIDs to cleartext in LogViewerUI. Known GUIDs are Win32Apps and WinGetApps by default.
With -Online option also Powershell scripts, Proactive Remediation scripts and custom Compliance script will get name shown in UI.
Often this parameter helps a lot debugging log entries in LogViewerUI

.PARAMETER LongRunningPowershellNotifyThreshold
Threshold (seconds) after Timeline report will show warning message for long running Powershell scripts. Default value is 180 seconds.

.PARAMETER ExportTextFileName
Export Timeline information and possible Powershell script error to text file.
This expects either text filename or fullpath to textfile.

.EXAMPLE
   .\Get-IntuneManagementExtensionDiagnostics.ps1
.EXAMPLE
   .\Get-IntuneManagementExtensionDiagnostics.ps1 -Online
.EXAMPLE
   .\Get-IntuneManagementExtensionDiagnostics.ps1 -Online -AllLogEntries -AllLogFiles
.EXAMPLE
   .\Get-IntuneManagementExtensionDiagnostics.ps1 -Online -AllLogEntries -ShowAllTimelineEntries
.EXAMPLE
   .\Get-IntuneManagementExtensionDiagnostics.ps1 -Online -ShowLogViewerUI
.EXAMPLE
   .\Get-IntuneManagementExtensionDiagnostics.ps1 -Online -ShowLogViewerUI -ConvertAllKnownGuidsToClearText
.EXAMPLE
   .\Get-IntuneManagementExtensionDiagnostics.ps1 -Online -Today
.EXAMPLE
   .\Get-IntuneManagementExtensionDiagnostics.ps1 -Online -LogStartDateTime "10.3.2023 5.00:00" -LogEndDateTime "11.3.2023 23.00:00"
.EXAMPLE
   .\Get-IntuneManagementExtensionDiagnostics.ps1 "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log"
.EXAMPLE
    .\Get-IntuneManagementExtensionDiagnostics.ps1 -LogFile "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log"
.EXAMPLE
    .\Get-IntuneManagementExtensionDiagnostics.ps1 -LogFilesFolder "C:\temp\MDMDiagReport"
.EXAMPLE
   .\Get-IntuneManagementExtensionDiagnostics.ps1 -Online -ShowStdOutInTimeline
.EXAMPLE
   .\Get-IntuneManagementExtensionDiagnostics.ps1 -Online -ShowErrorsInTimeline
.EXAMPLE
   .\Get-IntuneManagementExtensionDiagnostics.ps1 -Online -ShowErrorsSummary
.EXAMPLE
   .\Get-IntuneManagementExtensionDiagnostics.ps1 -Online -ExportTextFileName ExportTextFile.txt
.EXAMPLE
   .\Get-IntuneManagementExtensionDiagnostics.ps1 -Online -ExportTextFileName C:\temp\ExportTextFile.txt
.EXAMPLE
   .\Get-IntuneManagementExtensionDiagnostics.ps1 -AllLogEntries -AllLogFiles -ExportTextFileName C:\temp\ExportTextFile.txt
.EXAMPLE
   Get-ChildItem "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log" | .\Get-IntuneManagementExtensionDiagnostics.ps1 -AllLogEntries -Online
.INPUTS
   Script accepts File object as input. This would be same than specifying Parameter -LogFile
.OUTPUTS
   None
.NOTES
   You can download current version of this script from PowershellGallery with command

   Save-Script Get-IntuneManagementExtensionDiagnostics -Path ./
.LINK
   https://github.com/petripaavola/Get-IntuneManagementExtensionDiagnostics
#>

[CmdletBinding()]
Param(
	[Parameter(Mandatory=$false,
				HelpMessage = 'Enter Intune IME log file fullpath',
                ValueFromPipeline=$true,
                ValueFromPipelineByPropertyName=$true)]
	[Alias("FullName")]
    [String]$LogFile = $null,
	[Parameter(Mandatory=$false,
				HelpMessage = 'Enter Intune IME log files folder path',
                ValueFromPipeline=$false,
                ValueFromPipelineByPropertyName=$false)]
    [String]$LogFilesFolder = $null,
	[Parameter(Mandatory=$false)]
    [Switch]$Online,
	[Parameter(Mandatory=$false,
				HelpMessage = 'Enter Start DateTime for log entries (for example "10.3.2023 5:00:00")',
                ValueFromPipeline=$false,
                ValueFromPipelineByPropertyName=$false)]
    $LogStartDateTime = $null,
	[Parameter(Mandatory=$false,
				HelpMessage = 'Enter End DateTime for log entries (for example "11.3.2023 23.00:00")',
                ValueFromPipeline=$false,
                ValueFromPipelineByPropertyName=$false)]
    $LogEndDateTime = $null,
	[Parameter(Mandatory=$false)]
    [Switch]$ShowLogViewerUI,
	[Parameter(Mandatory=$false)]
    [Switch]$LogViewerUI,
	[Parameter(Mandatory=$false)]
    [Switch]$AllLogEntries,
	[Parameter(Mandatory=$false)]
    [Switch]$AllLogFiles,
	[Parameter(Mandatory=$false)]
    [Switch]$Today,
	[Parameter(Mandatory=$false)]
    [Switch]$ShowAllTimelineEntries,
	[Parameter(Mandatory=$false)]
	[Switch]$ShowStdOutInTimeline,
	[Parameter(Mandatory=$false)]
    [Switch]$ShowErrorsInTimeline,
	[Parameter(Mandatory=$false)]
    [Switch]$ShowErrorsSummary,
	[Parameter(Mandatory=$false)]
    [Switch]$ConvertAllKnownGuidsToClearText,
	[Parameter(Mandatory=$false,
				HelpMessage = 'Threshold seconds to highlight long running Powershell scripts',
                ValueFromPipeline=$true,
                ValueFromPipelineByPropertyName=$true)]
	[int]$LongRunningPowershellNotifyThreshold = 180,
	[Parameter(Mandatory=$false,
				HelpMessage = 'Enter text (.txt) filename to export info to',
                ValueFromPipeline=$false,
                ValueFromPipelineByPropertyName=$false)]
    [String]$ExportTextFileName=$null
)


Write-Host "Get-IntuneManagementExtensionDiagnostics.ps1 v1.1" -ForegroundColor Cyan
Write-Host "Author: Petri.Paavola@yodamiitti.fi / Microsoft MVP - Windows and Devices for IT"
Write-Host ""



# Set variables if we are in Windows Autopilot ESP (Enrollment Status Page)
# Idea is that user can just run the script without checking Parameters first
if($env:UserName -eq 'defaultUser0') {

	Write-Host "Detected running in Windows Autopilot Enrollment Status Page (ESP)" -ForegroundColor Yellow

	#$LOGFile='C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log'

	if((-not $LogFilesFolder) -or (-not $LOGFile)) {
		Write-Host "Automatically configuring parameters"
		Write-Host

		$LogFilesFolder = 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs'

		if(-not (Test-Path $LogFilesFolder)) {
			Write-Host "Log folder does not exist yet: $LogFilesFolder"
			Write-Host "Try again in a moment..."  -ForegroundColor Yellow
			Write-Host ""
			Exit 0
		}
	}

	# Process all found supported log files
	# This will not show file selection UI
	$AllLogFiles=$True

	# Process all log entries
	# This will not show time selection UI
	$AllLogEntries=$True

	# Not sure if this is good or bad by default so not configured by default for now
	#$ShowAllTimelineEntries=$True
	
	if(-not (Test-Path 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log')) {
		Write-Host "Log file does not exist yet: C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log"  -ForegroundColor Yellow
		Write-Host "Try again in a moment..."  -ForegroundColor Yellow
		Write-Host ""
		Exit 0
	}
}


# Work in Progress
<#
	[Parameter(Mandatory=$false)]
    [Switch]$ExportJson
#>


# Hashtable with ids and names
$IdHashtable = @{}

# Save timeline objects to this List
$observedTimeline = [System.Collections.Generic.List[PSObject]]@()

# Save application download statistics to this list
$ApplicationDownloadStatistics = [System.Collections.Generic.List[PSObject]]@()

################ Functions ################

# This is aligned with Michael Niehaus's Get-AutopilotDiagnostics script just in case
# region Functions
    Function RecordStatusToTimeline {
        param
        (
            [Parameter(Mandatory=$true)] [String] $date,
			[Parameter(Mandatory=$true)] [String] $status,
			[Parameter(Mandatory=$false)] [String] $detail,
			[Parameter(Mandatory=$false)] $seconds,
            [Parameter(Mandatory=$false)] [String] $logEntry,
            [Parameter(Mandatory=$false)] [String] $color
            
        )

		$observedTimeline.add([PSCustomObject]@{
			'Date' = $date
			'Status' = $status
			'Detail' = $detail
			'Seconds' = $seconds
			'LogEntry' = $logEntry
			'Color' = $color
			})
	}
# endregion Functions

	Function Get-AppIntent {
		Param(
			$AppId
			)

		$intent = 'Unknown Intent'

		if($AppId) {
			if($IdHashtable.ContainsKey($AppId)) {
				$AppPolicy=$IdHashtable[$AppId]

				if($AppPolicy.Intent) {
					Switch ($AppPolicy.Intent)
					{
						0	{ $intent = 'Not Targeted' }
						1	{ $intent = 'Available Install' }
						3	{ $intent = 'Required Install' }
						4	{ $intent = 'Required Uninstall' }
						default { $intent = 'Unknown Intent' }
					}
				}				
			}
		}
		
		return $intent
	}

	Function Get-AppIntentNameForNumber {
		Param(
			$IntentNumber
			)

		Switch ($IntentNumber)
		{
			0	{ $intentNumber = 'Not Targeted' }
			1	{ $intentNumber = 'Available Install' }
			3	{ $intentNumber = 'Required Install' }
			4	{ $intentNumber = 'Required Uninstall' }
			default { $intentNumber = 'Unknown Intent' }
		}
		
		return $intentNumber
	}

	Function Get-AppName {
		Param(
			$AppId
			)

		$AppName = $null

		if($AppId) {
			if($IdHashtable.ContainsKey($AppId)) {
				$AppPolicy=$IdHashtable[$AppId]

				if($AppPolicy.Name) {
					$AppName = $AppPolicy.Name
				}				
			}
		}
		
		return $AppName
	}



################ Functions ################

Write-Host "Starting Get-IntuneManagementExtensionDiagnostics`n"

# If LogFilePath is not specified then show log files in Out-GridView
# from folder C:\ProgramData\Microsoft\intunemanagementextension\Logs
if($LOGFile) {

	if(-not (Test-Path $LOGFile)) {
		Write-Host "Log file does not exist: $LOGFile" -ForegroundColor Yellow
		Write-Host "Script will exit" -ForegroundColor Yellow
		Write-Host ""
		Exit 0
	}
	$SelectedLogFiles = Get-ChildItem -Path $LOGFile
	
} else {

	if($LogFilesFolder) {

		if(-not (Test-Path $LogFilesFolder)) {
			Write-Host "LogFilesFolder: $LogFilesFolder does not exist" -ForegroundColor Yellow
			Write-Host "Script will exit" -ForegroundColor Yellow
			exit 0
		}

		# Sort files: new files first and IntuneManagementExtension before AgentExecutor
		$LogFiles = Get-ChildItem -Path $LogFilesFolder -Filter *.log | Where-Object { ($_.Name -like '*IntuneManagementExtension*.log') -or ($_.Name -like '*AgentExecutor*.log') } | Sort-Object -Property Name -Descending | Sort-Object -Property LastWriteTime -Descending

	} else {

		# Sort files: new files first and IntuneManagementExtension before AgentExecutor
		$LogFiles = Get-ChildItem -Path 'C:\ProgramData\Microsoft\intunemanagementextension\Logs' -Filter *.log | Where-Object { ($_.Name -like '*IntuneManagementExtension*.log') -or ($_.Name -like '*AgentExecutor*.log') } | Sort-Object -Property Name -Descending | Sort-Object -Property LastWriteTime -Descending

	}
	
	# Show log files in Out-GridView
	# This variable is automatically configured in ESP
	if(-not $SelectedLogFiles) {
		if($AllLogFiles) {
			$SelectedLogFiles = $LogFiles
		} else {
			$SelectedLogFiles = $LogFiles | Out-GridView -Title 'Select log file to show in Out-GridView from path C:\ProgramData\Microsoft\intunemanagementextension\Logs' -OutputMode Multiple
		}
	}
	
	if(-not $SelectedLogFiles) {
		Write-Host "No log file(s) selected. Script will exit!`n" -ForegroundColor Yellow
		Exit 0
	}
}


# Check whether we show time selection on Out-GridView or not
if(((-not $LogStartDateTime) -and (-not $LogEndDateTime)) -and (-not $AllLogEntries) -and (-not $Today)) {
	
	$LogStartEndTimeOutGridviewEntries = [System.Collections.Generic.List[PSObject]]@()
	
	$LogStartEndTimeOutGridviewEntries.add([PSCustomObject]@{
		'Log Start and End time' = 'All log entries';
		'LogStartDateTimeObject' = Get-Date 1/1/1900;
		'LogEndDateTimeObject' = (Get-Date).AddYears(1000);
		})

	$LogStartEndTimeOutGridviewEntries.add([PSCustomObject]@{
		'Log Start and End time' = 'Current day from midnight';
		'LogStartDateTimeObject' = Get-Date -Hour 0 -Minute 0 -Second 0;
		'LogEndDateTimeObject' = (Get-Date).AddYears(1000);
		})
		
	$LogStartEndTimeOutGridviewEntries.add([PSCustomObject]@{
		'Log Start and End time' = 'Last 5 minutes';
		'LogStartDateTimeObject' = (Get-Date).AddMinutes(-5);
		'LogEndDateTimeObject' = (Get-Date).AddYears(1000);
		})

	$LogStartEndTimeOutGridviewEntries.add([PSCustomObject]@{
		'Log Start and End time' = 'Last 15 minutes';
		'LogStartDateTimeObject' = (Get-Date).AddMinutes(-15);
		'LogEndDateTimeObject' = (Get-Date).AddYears(1000);
		})

	$LogStartEndTimeOutGridviewEntries.add([PSCustomObject]@{
		'Log Start and End time' = 'Last 30 minutes';
		'LogStartDateTimeObject' = (Get-Date).AddMinutes(-30);
		'LogEndDateTimeObject' = (Get-Date).AddYears(1000);
		})

	$LogStartEndTimeOutGridviewEntries.add([PSCustomObject]@{
		'Log Start and End time' = 'Last 1 hour';
		'LogStartDateTimeObject' = (Get-Date).AddHours(-1);
		'LogEndDateTimeObject' = (Get-Date).AddYears(1000);
		})

	$LogStartEndTimeOutGridviewEntries.add([PSCustomObject]@{
		'Log Start and End time' = 'Last 2 hours';
		'LogStartDateTimeObject' = (Get-Date).AddHours(-2);
		'LogEndDateTimeObject' = (Get-Date).AddYears(1000);
		})

	$LogStartEndTimeOutGridviewEntries.add([PSCustomObject]@{
		'Log Start and End time' = 'Last 6 hours';
		'LogStartDateTimeObject' = (Get-Date).AddHours(-6);
		'LogEndDateTimeObject' = (Get-Date).AddYears(1000);
		})

	$LogStartEndTimeOutGridviewEntries.add([PSCustomObject]@{
		'Log Start and End time' = 'Last 24 hours';
		'LogStartDateTimeObject' = (Get-Date).AddHours(-24);
		'LogEndDateTimeObject' = (Get-Date).AddYears(1000);
		})

	$LogStartEndTimeOutGridviewEntries.add([PSCustomObject]@{
		'Log Start and End time' = 'First 30 minutes';
		'LogStartDateTimeObject' = Get-Date 1/1/1900;
		'LogEndDateTimeObject' = 'Analyzing DateTime later';
		})

	$LogStartEndTimeOutGridviewEntries.add([PSCustomObject]@{
		'Log Start and End time' = 'First 1 hour';
		'LogStartDateTimeObject' = Get-Date 1/1/1900;
		'LogEndDateTimeObject' = 'Analyzing DateTime later';
		})

	$LogStartEndTimeOutGridviewEntries.add([PSCustomObject]@{
		'Log Start and End time' = 'First 2 hours';
		'LogStartDateTimeObject' = Get-Date 1/1/1900;
		'LogEndDateTimeObject' = 'Analyzing DateTime later';
		})

	$LogStartEndTimeOutGridviewEntries.add([PSCustomObject]@{
		'Log Start and End time' = 'First 6 hours';
		'LogStartDateTimeObject' = Get-Date 1/1/1900;
		'LogEndDateTimeObject' = 'Analyzing DateTime later';
		})

	$LogStartEndTimeOutGridviewEntries.add([PSCustomObject]@{
		'Log Start and End time' = 'First 24 hours';
		'LogStartDateTimeObject' = Get-Date 1/1/1900;
		'LogEndDateTimeObject' = 'Analyzing DateTime later';
		})

	# Show predefined values in Out-GridView
	$SelectedTimeFrameObject = $LogStartEndTimeOutGridviewEntries | Out-GridView -Title 'Select timeframe to show log entries' -OutputMode Multiple
	
	if($SelectedTimeFrameObject) {
	$LogStartDateTimeObject = $SelectedTimeFrameObject.LogStartDateTimeObject
	$LogEndDateTimeObject = $SelectedTimeFrameObject.LogEndDateTimeObject
		
	} else {
		Write-Host "No timeframe selected." -ForegroundColor Red
		Write-Host "Script will exit"
		Exit 0		
	}
	

} else {

	# Set default time which can be changed depenging on -Parameters

	# Use StartTime from year 1900
	$LogStartDateTimeObject = Get-Date 1/1/1900

	# Use EndTime +1000 years from today
	$LogEndDateTimeObject = (Get-Date).AddYears(1000)

	if($Today) {
		$LogStartDateTimeObject = Get-Date -Hour 0 -Minute 0 -Second 0
		$LogEndDateTimeObject = (Get-Date).AddYears(1000)
	}

	if($AllLogEntries) {
		# Use StartTime from year 1900
		$LogStartDateTimeObject = Get-Date 1/1/1900
		
		# Use EndTime +1000 years from today
		$LogEndDateTimeObject = (Get-Date).AddYears(1000)
	}

	# Check $LogStartDateTime can be converted to Powershell DateTime object
	if($LogStartDateTime) {
		# Try converting given parameter to Powershell DateTime object
		
		Try {
			$LogStartDateTimeObject = Get-Date $LogStartDateTime
		} catch {
			Write-Host "Given parameter LogStartDateTime is not in valid DateTime format." -ForegroundColor Red
			Write-Host "Script will exit"
			Exit 0
		}
	}

	# Check $LogEndDateTime can be converted to Powershell DateTime object
	if($LogEndDateTime) {
		# Try converting given parameter to Powershell DateTime object
		
		Try {
			$LogEndDateTimeObject = Get-Date $LogEndDateTime
		} catch {
			Write-Host "Given parameter LogEndDateTime is not in valid DateTime format." -ForegroundColor Red
			Write-Host "Script will exit"
			Exit 0
		}
	}
}


# Download Powershell scripts, Proactive Remediation scripts and custom Compliance Policy scripts
if ($Online) {

	# Test if we are in Enrollment Status Page (ESP) phase
	# Detect defaultuser0 loggedon
	if($env:UserName -eq 'defaultUser0') {

		# Make sure we can connect
		$module = Import-Module Microsoft.Graph.Intune -PassThru -ErrorAction Ignore
		if (-not $module) {
			Write-Host "Installing module Microsoft.Graph.Intune"
			Install-Module Microsoft.Graph.Intune -Force
		}
		Import-Module Microsoft.Graph.Intune

		$MSGraphEnvironment = Connect-MSGraph
		Write-Host "Connected to tenant $($graph.TenantId)"

	
	} else {
		Write-Host "Connecting to Intune using Powershell Intune-Module"
		Update-MSGraphEnvironment -SchemaVersion 'beta'
		$Success = $?

		if (-not $Success) {
			Write-Host "Failed to update MSGraph Environment schema to Beta!`n" -ForegroundColor Red
			Write-Host "Make sure you have installed Intune Powershell module"
			Write-Host "`nYou can install Intune module to your user account with command:"
			Write-Host "Install-Module -Name Microsoft.Graph.Intune -Scope CurrentUser`n" -ForegroundColor Yellow
			Write-Host "`nor you can install machine-wide Intune module with command:`nInstall-Module -Name Microsoft.Graph.Intune"
			Write-Host "More information: https://github.com/microsoft/Intune-PowerShell-SDK"
			Exit 1
		}

		$MSGraphEnvironment = Connect-MSGraph
		$Success = $?

		if ($Success -and $MSGraphEnvironment) {
			$TenantId = $MSGraphEnvironment.tenantId
			$AdminUserUPN = $MSGraphEnvironment.upn

			Write-Host "Connected to tenant $TenantId as user $AdminUserUPN"
			
		} else {
			Write-Host "Could not connect to MSGraph!" -ForegroundColor Red
			Write-Host "Will not download information from Microsoft Intune" -ForegroundColor Yellow
			Write-Host "Script will exit."
			Exit 0
		}
	}

	if($MSGraphEnvironment) {
		Write-Host "Download Intune Powershell scripts"
		# Get PowerShell Scripts
		$url = 'https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts'
		$MSGraphRequest = Invoke-MSGraphRequest -Url $url -HttpMethod 'GET'
		$Success = $?

		if (-not ($Success)) {
			Write-Error "Error downloading Intune Powershell scripts"
			#return $null
		} else {
			$AllIntunePowershellScripts = Get-MSGraphAllPages -SearchResult $MSGraphRequest
			$Success = $?
			if($Success) {
				Write-Host "Done" -ForegroundColor Green
				
				# Add Name Property to object
				$AllIntunePowershellScripts | Foreach-Object { $_ | Add-Member -MemberType noteProperty -Name name -Value $_.displayName }
				
				# Add all PowershellScripts to Hashtable
				$AllIntunePowershellScripts | Foreach-Object { $id = $_.id; $value=$_; $IdHashtable["$id"] = $value }
			} else {
				Write-Error "Error downloading All Intune Powershell scripts"
			}
		}

		Start-Sleep 1
		
		Write-Host "Download Intune Proactive Remediations Scripts"
		# Get Proactive Remediations Scripts
		$url = 'https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts'
		$MSGraphRequest = Invoke-MSGraphRequest -Url $url -HttpMethod 'GET'
		$Success = $?

		if (-not ($Success)) {
			Write-Error "Error downloading Intune Proactive Remediations scripts"
			#return $null
		} else {
			$AllIntuneProactiveRemediationsScripts = Get-MSGraphAllPages -SearchResult $MSGraphRequest
			$Success = $?
			if($Success) {
				Write-Host "Done" -ForegroundColor Green
				
				# Add Name Property to object
				$AllIntuneProactiveRemediationsScripts | Foreach-Object { $_ | Add-Member -MemberType noteProperty -Name name -Value $_.displayName }
				
				# Add all PowershellScripts to Hashtable
				$AllIntuneProactiveRemediationsScripts | Foreach-Object { $id = $_.id; $value=$_; $IdHashtable["$id"] = $value }
			} else {
				Write-Error "Error downloading All Intune Proactive Remediations scripts"
			}
		}

		Start-Sleep 1

		Write-Host "Download Intune Windows Device Compliance custom Scripts"
		# Get Windows Device Compliance custom Scripts
		$url = 'https://graph.microsoft.com/beta/deviceManagement/deviceComplianceScripts'
		$MSGraphRequest = Invoke-MSGraphRequest -Url $url -HttpMethod 'GET'
		$Success = $?

		if (-not ($Success)) {
			Write-Error "Error downloading Intune Windows custom Compliance scripts"
			#return $null
		} else {
			$AllIntuneCustomComplianceScripts = Get-MSGraphAllPages -SearchResult $MSGraphRequest
			$Success = $?
			if($Success) {
				Write-Host "Done" -ForegroundColor Green
				
				# Add Name Property to object
				$AllIntuneCustomComplianceScripts | Foreach-Object { $_ | Add-Member -MemberType noteProperty -Name name -Value $_.displayName }
				
				# Add all PowershellScripts to Hashtable
				$AllIntuneCustomComplianceScripts | Foreach-Object { $id = $_.id; $value=$_; $IdHashtable["$id"] = $value }
			} else {
				Write-Error "Error downloading All Intune Windows custom Compliance scripts"
			}
		}
	} else {
		Write-Host "Not connected to Microsoft Intune, skip downloading script names...." -ForegroundColor Yellow
	}
}
Write-Host ""


# Run Report and/or Out-GridView in loop as long as user selects last line which will reload log file
# Otherwise Out-GridView will exit if something else is selected than last line (reload line)
do {

	# Create Generic list where log entry custom objects are added
	$LogEntryList = [System.Collections.Generic.List[PSObject]]@()
	
	# Go through all selected log file(s)
	Foreach ($SelectedLogFile in $SelectedLogFiles) {

		$LogFilePath = $SelectedLogFile.FullName
		$LogFileName = $SelectedLogFile.Name

		Write-Host "Processing file $LogFileName"

		# Initialize variables
		$LineNumber=1
		$MultilineLogEntryStartsArrayIndex=0
		$MultilineLogEntryStartFound=$False

		$Log = Get-Content -Path $LogFilePath

		# This matches for cmtrace type logs
		# Test with https://regex101.com
		# String: <![LOG[ExecutorLog AgentExecutor gets invoked]LOG]!><time="21:38:06.3814532" date="2-14-2023" component="AgentExecutor" context="" type="1" thread="1" file="">

		# This matches single line full log entry
		$SingleLineRegex = '^\<\!\[LOG\[(.*)]LOG\].*\<time="([0-9]{1,2}):([0-9]{1,2}):([0-9]{1,2}).([0-9]{1,})".*date="([0-9]{1,2})-([0-9]{1,2})-([0-9]{4})" component="(.*?)" context="(.*?)" type="(.*?)" thread="(.*?)" file="(.*?)">$'

		# Start of multiline log entry
		$FirstLineOfMultiLineLogRegex = '^\<\!\[LOG\[(.*)$'

		# End of multiline log entry
		$LastLineOfMultiLineLogRegex = '^(.*)\]LOG\]\!>\<time="([0-9]{1,2}):([0-9]{1,2}):([0-9]{1,2}).([0-9]{1,})".*date="([0-9]{1,2})-([0-9]{1,2})-([0-9]{4})" component="(.*?)" context="(.*?)" type="(.*?)" thread="(.*?)" file="(.*?)">$'


		# Process each log file line one by one
		Foreach ($CurrentLogEntry in $Log) {

			# Get data from CurrentLogEntry
			if($CurrentLogEntry -Match $SingleLineRegex) {
				# This matches single line log entry
				
				# Regex found match
				$LogMessage = $Matches[1].Trim()

				$Hour = $Matches[2]
				$Minute = $Matches[3]
				$Second = $Matches[4]
				
				$MilliSecondFull = $Matches[5]
				# Cut milliseconds to 0-999
				# Time unit is so small that we don't even bother to round the value
				$MilliSecond = $MilliSecondFull.Substring(0,3)

				$Month = $Matches[6]
				$Day = $Matches[7]
				$Year = $Matches[8]
				
				$Component = $Matches[9]
				$Context = $Matches[10]
				$Type = $Matches[11]
				$Thread = $Matches[12]
				$File = $Matches[13]

				$Param = @{
					Hour=$Hour
					Minute=$Minute
					Second=$Second
					MilliSecond=$MilliSecond
					Year=$Year
					Month=$Month
					Day=$Day
				}

				$LogEntryDateTime = Get-Date @Param
				#Write-Host "DEBUG `$LogEntryDateTime: $LogEntryDateTime" -ForegroundColor Yellow
				
				# This works for humans but does not sort
				#$DateTimeToLogFile = "$($Hour):$($Minute):$($Second).$MilliSecondFull $Day/$Month/$Year"

				# Add leading 0 so sorting works right
				if($Month -like "?") { $Month = "0$Month" }

				# Add leading 0 so sorting works right
				if($Day -like "?") { $Day = "0$Day" }

				# This does sorting right way
				$DateTimeToLogFile = "$Year-$Month-$Day $($Hour):$($Minute):$($Second).$MilliSecondFull"

				# Create Powershell custom object and add it to list
				$LogEntryList.add([PSCustomObject]@{
					'Index' = $null;
					'LogEntryDateTime' = $LogEntryDateTime;
					'FileName' = $LogFileName;
					'Line' = $LineNumber;
					'DateTime' = $DateTimeToLogFile;
					'Multiline' = '';
					'ProcessRunTime' = $ProcessRunTime;
					'Message' = $LogMessage;
					'Component' = $Component;
					'Context' = $Context;
					'Type' = $Type;
					'Thread' = $Thread;
					'File' = $File
					})

			} elseif ($CurrentLogEntry -Match $FirstLineOfMultiLineLogRegex) {
				# This is start of multiline log entry
				# Single line regex did not get results so we are dealing multiline case separately here

				#Write-Host "DEBUG Start of multiline regex: $CurrentLogEntry" -ForegroundColor Yellow

				$MultilineLogEntryStartFound=$True

				# Regex found match
				$LogMessage = $Matches[1].Trim()
				
				$LogEntryDateTime = ''
				$DateTimeToLogFile = ''
				$Component = ''
				$Context = ''
				$Type = ''
				$Thread = ''
				$File = ''

				# Create Powershell custom object and add it to list
				$LogEntryList.add([PSCustomObject]@{
					'Index' = $null;
					'LogEntryDateTime' = $LogEntryDateTime;
					'FileName' = $LogFileName;
					'Line' = $LineNumber;
					'DateTime' = $DateTimeToLogFile;
					'Multiline' = '';
					'ProcessRunTime' = '';
					'Message' = $LogMessage;
					'Component' = $Component;
					'Context' = $Context;
					'Type' = $Type;
					'Thread' = $Thread;
					'File' = $File
					})

				$MultilineLogEntryStartsArrayIndex = $LogEntryList.Count - 1

			} elseif ($CurrentLogEntry -Match $LastLineOfMultiLineLogRegex) {
				# This is end of multiline log entry
				# Single line regex did not get results so we are dealing multiline case separately here

				# Regex found match
				$LogMessage = $Matches[1].Trim()

				$Hour = $Matches[2]
				$Minute = $Matches[3]
				$Second = $Matches[4]
				
				$MilliSecondFull = $Matches[5]
				# Cut milliseconds to 0-999
				# Time unit is so small that we don't even bother to round the value
				$MilliSecond = $MilliSecondFull.Substring(0,3)
				
				$Month = $Matches[6]
				$Day = $Matches[7]
				$Year = $Matches[8]
				
				$Component = $Matches[9]
				$Context = $Matches[10]
				$Type = $Matches[11]
				$Thread = $Matches[12]
				$File = $Matches[13]

				$Param = @{
					Hour=$Hour
					Minute=$Minute
					Second=$Second
					MilliSecond=$MilliSecond
					Year=$Year
					Month=$Month
					Day=$Day
				}

				$LogEntryDateTime = Get-Date @Param
				#Write-Host "DEBUG `$LogEntryDateTime: $LogEntryDateTime" -ForegroundColor Yellow

				# This works for humans but does not sort
				#$DateTimeToLogFile = "$($Hour):$($Minute):$($Second).$MilliSecondFull $Day/$Month/$Year"

				# Add leading 0 so sorting works right
				if($Month -like "?") { $Month = "0$Month" }

				# Add leading 0 so sorting works right
				if($Day -like "?") { $Day = "0$Day" }

				# This does sorting right way
				$DateTimeToLogFile = "$Year-$Month-$Day $($Hour):$($Minute):$($Second).$MilliSecondFull"

				if($MultilineLogEntryStartFound) {
					$Multiline = '---->'
				} else {
					$Multiline = ''
				}

				# Create Powershell custom object and add it to list
				$LogEntryList.add([PSCustomObject]@{
					'Index' = $null;
					'LogEntryDateTime' = $LogEntryDateTime;
					'FileName' = $LogFileName;
					'Line' = $LineNumber;
					'DateTime' = $DateTimeToLogFile;
					'Multiline' = $Multiline;
					'ProcessRunTime' = '';
					'Message' = $LogMessage;
					'Component' = $Component;
					'Context' = $Context;
					'Type' = $Type;
					'Thread' = $Thread;
					'File' = $File
					})

				# Process previous lines and add DateTime, Multiline, Component, Context, Type, Thread and File
				# There can be corrupted log entries so there can be multiline entries without proper multiline start log entry
				# So we will add information until we find log entry which has DateTime value

				$MultilineLogEntryEndsArrayIndex = $LogEntryList.Count - 1
				
				# Add data to starting multiline log entry object
				# DateTime, Component, Context, Type, Thread, File information
				$LogEntryList[$MultilineLogEntryStartsArrayIndex].LogEntryDateTime = $LogEntryDateTime
				$LogEntryList[$MultilineLogEntryStartsArrayIndex].DateTime = $DateTimeToLogFile
				$LogEntryList[$MultilineLogEntryStartsArrayIndex].Multiline = '<----'
				$LogEntryList[$MultilineLogEntryStartsArrayIndex].Component = $Component
				$LogEntryList[$MultilineLogEntryStartsArrayIndex].Context = $Context
				$LogEntryList[$MultilineLogEntryStartsArrayIndex].Type = $Type
				$LogEntryList[$MultilineLogEntryStartsArrayIndex].Thread = $Thread
				$LogEntryList[$MultilineLogEntryStartsArrayIndex].File = $File
				
				# Add data to multiline entries which are not end or start entries
				#if($MultilineLogEntryEndsArrayIndex - $MultilineLogEntryStartsArrayIndex -gt 1) {
					# Multiline entry is more than 2 lines

					# Add DateTime, Component, Context, Type, Thread, File information
					# to Multiline log entry from last to first
					#for($i=$MultilineLogEntryEndsArrayIndex - 1; $i -gt $MultilineLogEntryStartsArrayIndex ; $i--) {
					
					# Add information to previous log entries if they don't have DateTime-value
					# Stop when we find DateTime-value
					$i = $MultilineLogEntryEndsArrayIndex -1
					While($LogEntryList[$i].DateTime -eq $null) {
						#Write-Host "DEBUG: Multiline add data: `$MultilineLogEntryStartsArrayIndex=$MultilineLogEntryStartsArrayIndex `$MultilineLogEntryEndsArrayIndex=$MultilineLogEntryEndsArrayIndex `$i=$i `$Message=$($LogEntryList[$i].Message)"
						
						$LogEntryList[$i].LogEntryDateTime = $LogEntryDateTime
						$LogEntryList[$i].DateTime = $DateTimeToLogFile
						$LogEntryList[$i].Multiline = ' --- '
						$LogEntryList[$i].Component = $Component
						$LogEntryList[$i].Context = $Context
						$LogEntryList[$i].Type = $Type
						$LogEntryList[$i].Thread = $Thread
						$LogEntryList[$i].File = $File
						
						$i--
					}
					
			} else {
				# We didn't catch log entry with our regex
				# This should be multiline log entry but not first or last line in that log entry
				# This can also be some line that should be matched with (other) regex
				
				#Write-Host "DEBUG: $CurrentLogEntry"  -ForegroundColor Yellow
				
				# Create Powershell custom object and add it to list
				$LogEntryList.add([PSCustomObject]@{
					'Index' = $null;
					'LogEntryDateTime' = $null;
					'FileName' = $LogFileName;
					'Line' = $LineNumber;
					'DateTime' = $null;
					'Multiline' = '';
					'ProcessRunTime' = $null;
					'Message' = $CurrentLogEntry;
					'Component' = '';
					'Context' = '';
					'Type' = '';
					'Thread' = '';
					'File' = ''
					})
			}

			$LineNumber++

		} # Foreach end single log file line by line foreach
	
	} # Foreach end all log files
	Write-Host ""

	# Check if there was log entries/lines which were not detected
	$DroppedLines = $LogEntryList | Where-Object DateTime -eq $null
	if($DroppedLines) {
		Write-Host "These lines were detected not to have right syntax so they were dropped" -ForegroundColor Yellow
		Foreach($Line in $DroppedLines) {
			Write-Host "$($Line.FileName) Line $($Line.Line) $($Line.Message)"
			
			# FIXME Copy timestamp from previous log entry
			
		}
		Write-Host ""
	}


	# Take last $LogEntryDateTime and add 1 seconds to it
	$ReloadDate = (Get-Date).AddHours(1)
	$ReloadDateTime = "$($ReloadDate.Year)-$($ReloadDate.Month)-$($ReloadDate.Day) $($ReloadDate.TimeOfDay)"

	# Add last entry which will reload log file if selected
	$LogEntryList.add([PSCustomObject]@{
				'Index' = $null;
				'LogEntryDateTime' = $ReloadDate;
				'FileName' = '';
				'Line' = '';
				'DateTime' = $ReloadDateTime;
				'Multiline' = '';
				'ProcessRunTime' = '';
				'Message' = 'RELOAD LOG FILE(S).      Select this line and OK button from right botton corner';
				'Component' = '';
				'Context' = '';
				'Type' = '';
				'Thread' = '';
				'File' = ''
				})



	# Check if First x minutes/hours was selected
	if($LogEndDateTimeObject -eq 'Analyzing DateTime later') {
		# Sort log entries by DateTime and secondary by Line number
		# because with multiple files we are not organized by default
		# Find first log entry with LogEntryDateTime value
		# There is some rare case there could be log entry with empty value
		Foreach ($LogEntryObject in $LogEntryList | Sort-Object -Property DateTime, Line) {
			Try {
				$FirstLogEntryTimeDate = Get-Date $LogEntryObject.LogEntryDateTime
				Break
			} catch {
				
			}
		}

		# And figure out LogEndDateTime based on first log entry timestamp
		if($SelectedTimeFrameObject.'Log Start and End time' -eq 'First 30 minutes') {
			$LogEndDateTimeObject = ($FirstLogEntryTimeDate).AddMinutes(30)
		}

		if($SelectedTimeFrameObject.'Log Start and End time' -eq 'First 1 hour') {
			$LogEndDateTimeObject = ($FirstLogEntryTimeDate).AddHours(1)
		}

		if($SelectedTimeFrameObject.'Log Start and End time' -eq 'First 2 hours') {
			$LogEndDateTimeObject = ($FirstLogEntryTimeDate).AddHours(2)
		}

		if($SelectedTimeFrameObject.'Log Start and End time' -eq 'First 6 hours') {
			$LogEndDateTimeObject = ($FirstLogEntryTimeDate).AddHours(6)
		}
		
		if($SelectedTimeFrameObject.'Log Start and End time' -eq 'First 24 hours') {
			$LogEndDateTimeObject = ($FirstLogEntryTimeDate).AddHours(24)
		}

		Write-Host "Log entries End time:   $LogEndDateTimeObject`n"
	}

	Write-Host "Creating report"
	Write-Host "   Start time: $LogStartDateTimeObject"
	Write-Host "   End time:   $LogEndDateTimeObject`n"


	# Filter and Sort objects
	# We need to sort with DateTime AND Line
	# because there can are multiple log entries with same timestamp
	#$LogEntryList = $LogEntryList | Sort-Object -Property DateTime, Line
	$LogEntryList = $LogEntryList | Where-Object { ($_.LogEntryDateTime -gt $LogStartDateTimeObject) -and ($_.LogEntryDateTime -lt $LogEndDateTimeObject) } | Sort-Object -Property DateTime, Line

	# Add Index property value
	# This is used for sorting when doing analyzing in Out-GridView
	$index = 1
	Foreach($LogEntry in $LogEntryList) {
		
		# Use this to find out if we did catch all lines
		# Look for the first object index and line numbers that if they match
		#Write-host "$index $($LogEntry.Line) $($LogEntry.Message)"

		# DateTime is empty if there was corrupted log entries
		# We should set last known timestamp to those but this is just in case we still missed something
		if($LogEntry.DateTime) {
			$LogEntry.Index = $index
			$index++
		}
	}


	# region analyze log entries
	##########################
	# Analyze log entries
	#
	# Try to find end and start log entries for Powershell, Win32App and WinGetApp installations
	# We go backwards because if there is process end log entry
	# then there has to be start log entry (going upwards) for the same process (same thread in this case)
	#
	# Start process could be "orphan" if process is stopped without end log entry
	# so it is hard to detect these cases (going from up to down)

	$DetectedInEspPhase=$False
	$skipDeviceStatusPage='NotFoundYet'
	$skipUserStatusPage='NotFoundYet'
	$skipDeviceStatusPageIndex=$null
	$skipUserStatusPageIndex=$null
	$NotInEspInSessionFound=$false
	$UserESPFinished=$false
	$UserESPFinishedNotFoundWarningMessageShown=$false
	
	$PreviousLoggedOnUser='NotInitializedYet'
	$LoggedOnUserDefaultUser0Shown=$false

	$PowershellScriptStartLogEntryFound=$false
	$PowershellScriptStartLogEntryIndex=$null

	# Array
	$PowershellScriptErrorLogs = @()
	$IntuneAppPolicies = $null
	
	# This is index of colon character in timeline
	# This uses -f formatting syntax
	$ColonIndentInTimeLine = -22
	
	$i = 0
	While ($i -lt $LogEntryList.Count) {
		# Clear variables in every loop round
		$LoggedOnUser=$false
		
		#Write-Host "`$i=$i"
		#Write-Host "$($LogEntryList[$i].Message)"


		# region Find Username and UserSID
		if($LogEntryList[$i].Message -Like '`[Win32App`] Expected usersid for session *') {
			$Matches=$null
			if($LogEntryList[$i].Message -Match '^\[Win32App\] Expected usersid for session (.*) with name (.*) is (.*)$') {
				$UserSession = $Matches[1]
				$UserName = $Matches[2]
				$UserSID = $Matches[3]

				#Write-Host "DEBUG: $($LogEntryList[$i].Message)"

				# Check if we have UserSID object in Hashtable UserId object
				if($IdHashtable.ContainsKey($UserSId)) {
					$IdHashtable[$UserSId].userSID = $UserSID
					$IdHashtable[$UserSId].name = $UserName
					$IdHashtable[$UserSId].displayName = $UserName
					$IdHashtable[$UserSId].session = $UserSession

					$userId = $IdHashtable[$UserSId].userId
					if($userId) {
						# Check if we have UserID object in Hashtable
						if($IdHashtable.ContainsKey($UserId)) {
								# Add Name-value. We don't even check existing value
								# because we are already here.
								$IdHashtable[$UserId].userSID = $UserSID
								$IdHashtable[$UserId].name = $UserName
								$IdHashtable[$UserId].displayName = $UserName
								$IdHashtable[$UserId].session = $UserSession
						}
					}
				} else {
					# UserSID object does not exist in Hashtable
					# Create object which we will add to Hashtable
					$UserSIDCustomObjectProperties = @{
						id = $null
						userId = $null
						userSID = $UserSID
						name = $UserName
						displayName = $UserName
						session = $UserSession
					}
					$UserSIDCustomObject = New-Object -TypeName PSObject -Prop $UserSIDCustomObjectProperties

					# Create new UserId hashtable entry
					$IdHashtable.Add($UserSId , $UserSIDCustomObject)
				}
				
				#Write-Host "DEBUG: We should have name Property now`n$($IdHashtable | ConvertTo-Json)"
			}

			# Cleanup variables
			$UserSession = $null
			$UserName = $null
			$UserSID= $null
			$UserSIDCustomObject = $null
		}
		# endregion Find UserName and UserSID


<#		# This code works, but there was log entries with mismatched UserID and UserSID (from different users)
		# so took another approach to map UserID, UserSID and UserName
	

		# region Find UserId and UserSID
		if($LogEntryList[$i].Message -Like '`[Win32App`] EspPreparation starts for userId: *') {
			$Matches=$null
			if($LogEntryList[$i].Message -Match '^\[Win32App\] EspPreparation starts for userId: (.*) userSID: (.*)$') {
				$UserId = $Matches[1]
				$UserSID = $Matches[2]

				Write-Host "DEBUG: $($LogEntryList[$i].Message)"

				# Add information to $IdHashtable UserId key
				if($IdHashtable.ContainsKey($UserId)) {

					# Add UserID and UserSID -values. We don't even check existing value
					# because we are already here.
					$IdHashtable[$UserId].Id = $UserId
					$IdHashtable[$UserId].userId = $UserId
					$IdHashtable[$UserId].userSID = $UserSID

				} else {
					# Hashtable does NOT have UserId object

					if($IdHashtable.ContainsKey($UserSId)) {
						$UserName = $IdHashtable[$UserSId].name
						$Session = $IdHashtable[$UserSId].session
					} else {
						$UserName = $UserId
						$Session = $null
					}

					# Create object which we will add to Hashtable
					$UserIdCustomObjectProperties = @{
						id = $UserId
						userId = $UserId
						userSID = $UserSID
						name = $UserName
						session = $Session
					}

					$UserIdCustomObject = New-Object -TypeName PSObject -Prop $UserIdCustomObjectProperties

					# Create new UserId hashtable entry
					$IdHashtable.Add($UserId , $UserIdCustomObject)
				}

				# Add information to $IdHashtable UserSID key
				if($IdHashtable.ContainsKey($UserSID)) {

					if($IdHashtable.ContainsKey($UserId)) {
						$UserName = $IdHashtable[$UserId].name
						$Session = $IdHashtable[$UserId].session
					} else {
						$UserName = $UserId
						$Session = $null
					}

					# Add UserId-value. We don't even check existing value
					# because we are already here.
					$IdHashtable[$UserSId].id = $UserId
					$IdHashtable[$UserSId].userId = $UserId
					$IdHashtable[$UserSId].userSId = $UserSId
					$IdHashtable[$UserSId].name = $UserName
					$IdHashtable[$UserSId].session = $Session
					
				} else {
					# UserSID object does not exist in Hashtable
					# Create object which we will add to Hashtable
					$UserSIDCustomObjectProperties = @{
						id = $UserId
						userId = $UserId
						userSID = $UserSID
						name = $UserName
						session = $Session
					}
					$UserSIDCustomObject = New-Object -TypeName PSObject -Prop $UserSIDCustomObjectProperties

					# Create new UserId hashtable entry
					$IdHashtable.Add($UserSId , $UserSIDCustomObject)
				}
			}

			#Write-Host "DEBUG: We should have UserID and UserSID`n$($IdHashtable | ConvertTo-Json)"
			
			# Cleanup variables
			$UserIdCustomObject = $null
			$UserSIDCustomObject = $null
			$UserId = $null
			$UserSID = $null
			
		}
		# endregion Find UserId and UserSID
#>

		# region EMS Agent Started/Stopped/received shutdown signal
		# Excluded -or ($LogEntryList[$i].Message -eq 'EMS Agent received shutdown signal')
		if(($LogEntryList[$i].Message -eq 'EMS Agent Started') -or ($LogEntryList[$i].Message -eq 'EMS Agent Stopped')) {

				$LogEntryList[$i].ProcessRunTime = "$($LogEntryList[$i].Message)"
				
				RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status 'Info' -detail $LogEntryList[$i].Message -logEntry "Line $($LogEntryList[$i].Line)"
				
				# Continue to next line in Do-while -loop
				$i++
				Continue
		}
		# endregion EMS Agent Started/Stopped/received shutdown signal
		
		# region Find ESP SkipESPStatusPage(s)
		if($skipDeviceStatusPage -eq 'NotFoundYet') {
			$Matches=$null
			if($LogEntryList[$i].Message -Match '^\[Win32App\] CheckDeviceOnlyFromFirstSyncReg set skipDeviceStatusPage: (.*)$') {
				$skipDeviceStatusPage = $Matches[1]
				$LogEntryList[$i].ProcessRunTime = "Skip Device ESP: $skipDeviceStatusPage"
				
				$skipDeviceStatusPageIndex=$i
				
				# Continue to next line in Do-while -loop
				$i++
				Continue
			}
		}

		# region Find ESP SkipESPStatusPage(s)
		if($skipUserStatusPage -eq 'NotFoundYet') {
			$Matches=$null
			if($LogEntryList[$i].Message -Match '^\[Win32App\] CheckDeviceOnlyFromFirstSyncReg set skipUserStatusPage: (.*)$') {
				$skipUserStatusPage = $Matches[1]
				$LogEntryList[$i].ProcessRunTime = "Skip User ESP  : $skipUserStatusPage"
				
				$skipUserStatusPageIndex=$i
				
				# Continue to next line in Do-while -loop
				$i++
				Continue
			}
		}
		# endregion Find ESP SkipESPStatusPage(s)
		
		
		# region Find ESP (Enrollment Status Page) phases
		if(-not $DetectedInEspPhase) {
			if($LogEntryList[$i].Message -eq '[Win32App] In EspPhase: DeviceSetup') {
				$DetectedInEspPhase = $True

				# This is here because these lines exists later after ESP
				# So only showing this information if we are in ESP
				if($skipDeviceStatusPage -eq 'True') {
					
					RecordStatusToTimeline -date $LogEntryList[$skipDeviceStatusPageIndex].DateTime -status "Info" -detail "Skip Device ESP: $skipDeviceStatusPage" -logEntry "Line $($LogEntryList[$skipDeviceStatusPageIndex].Line)" -color 'Yellow'
					
				} else {
					
					RecordStatusToTimeline -date $LogEntryList[$skipDeviceStatusPageIndex].DateTime -status "Info" -detail "Skip Device ESP: $skipDeviceStatusPage" -logEntry "Line $($LogEntryList[$skipDeviceStatusPageIndex].Line)"
				}
				
				# This is here because these lines exists later after ESP
				# So only showing this information if we are in ESP
				if($skipUserStatusPage -eq 'True') {
					
					RecordStatusToTimeline -date $LogEntryList[$skipUserStatusPageIndex].DateTime -status "Info" -detail "Skip User   ESP: $skipUserStatusPage" -logEntry "Line $($LogEntryList[$skipUserStatusPageIndex].Line)" -color 'Yellow'
				} else {
					
					RecordStatusToTimeline -date $LogEntryList[$skipUserStatusPageIndex].DateTime -status "Info" -detail "Skip User   ESP: $skipUserStatusPage" -logEntry "Line $($LogEntryList[$skipUserStatusPageIndex].Line)"
				}

				$LogEntryList[$i].ProcessRunTime = "[Win32App] In Device ESP phase detected"
				
				RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Info" -detail "######################## Device ESP phase ########################" -logEntry "Line $($LogEntryList[$i].Line)"
				
				# Continue to next line in Do-while -loop
				$i++
				Continue
			}
		}
		# endregion Find ESP (Enrollment Status Page) phases

		# region User ESP finished
		if(-not $UserESPFinished) {
			if($LogEntryList[$i].Message -like '`[Win32App`] In EspPhase: AccountSetup. All Win32App has been added. Set SidecarReady: True*') {
				$UserESPFinished = $True

				$LogEntryList[$i].ProcessRunTime = "User ESP finished successfully"
				
				RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Info" -detail "######################## User ESP finished   ########################" -logEntry "Line $($LogEntryList[$i].Line)"
				
				# Continue to next line in Do-while -loop
				$i++
				Continue
			}
		}
		# endregion User ESP finished


		# region NotInEsp in session
		if(-not $NotInEspInSessionFound) {
			if($LogEntryList[$i].Message -eq '[Win32App] The EspPhase: NotInEsp in session') {
				$NotInEspInSessionFound = $True

				$LogEntryList[$i].ProcessRunTime = "[Win32App] The EspPhase: NotInEsp in session"
				
				RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Info" -detail "######################## Not in ESP detected ########################" -logEntry "Line $($LogEntryList[$i].Line)"
				
				# Continue to next line in Do-while -loop
				$i++
				Continue
			}
		}
		# endregion NotInEsp in session

		# region NotInESP found but did not find User ESP Finished
		# Show warning on this case. Did User ESP fail ???
		if((-not $UserESPFinishedNotFoundWarningMessageShown) -and ($DetectedInEspPhase) -and ((-not $UserESPFinished) -and ($NotInEspInSessionFound))) {
			$UserESPFinishedNotFoundWarningMessageShown=$True
		
			RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Warning" -detail "Did not detect User ESP finished. This is usually falsepositive warning." -logEntry "Line $($LogEntryList[$i].Line)" -Color 'Yellow'
			
		}
		# endregion NotInESP found but did not find User ESP Finished
		
		# region Find Logged on user
		if($LogEntryList[$i].Message -Like 'After impersonation: *') {
			$Matches=$null
			if($LogEntryList[$i].Message -Match '^After impersonation: (.*)$') {
				$LoggedOnUser = $Matches[1]
				
				#Write-Host "DEBUG: `$FirstLoggedOnUser`=$FirstLoggedOnUser"
				# Skip if defaultuser0 found (we are in device ESP phase still)
				if(-not $LoggedOnUserDefaultUser0Shown) {
					if($LoggedOnUser -like '*\defaultuser0') {
						$LoggedOnUserDefaultUser0Shown = $True
						
						RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "User Logon" -detail "Logged on user        :  $LoggedOnUser" -logEntry "Line $($LogEntryList[$i].Line)" -Color 'Yellow'
			
						# Set LoggedOnUser information to ProcessRuntime property
						$LogEntryList[$i].ProcessRunTime = "Logged on user: $LoggedOnUser"

						$i++
						Continue
					}
				} elseif (($LoggedOnUser -notlike '*\defaultuser0') -and ($LoggedOnUser -ne $PreviousLoggedOnUser)) {
					$PreviousLoggedOnUser=$LoggedOnUser

					RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "User Logon" -detail "Logged on user        :  $LoggedOnUser" -logEntry "Line $($LogEntryList[$i].Line)" -Color 'Yellow'
					
					$LogEntryList[$i].ProcessRunTime = "Logged on user: $LoggedOnUser"

				}


				# Not sure if we want to run this everytime or not
				# Now running it every time we find this log entry
				if($True) {
					# Try to find user session from previous log entry
					if($LogEntryList[$i-1].Message -Match '^Starting impersonation, session id = (.*)$') {
						$ImpersonationSessionId = $Matches[1]

						#Write-Host "Found ImpersonationSessionId $ImpersonationSessionId"

						# Try to find UserId from either of following lines
						# [PowerShell] Processing user session 3, user id = cdc3177c-df70-45e3-a469-4ef1be2b8345
						# [PowerShell] Get 3 policies for user cdc3177c-df70-45e3-a469-4ef1be2b8345 in session 3
						$b = $i + 1
						While(($LogEntryList[$b].Message -notlike 'After impersonation: *') -and ($b -lt $LogEntryList.Count)) {
							#Write-Host "DEBUG: Processing line: $($LogEntryList[$b].Message)"
							
							if(($LogEntryList[$i].Thread -eq $LogEntryList[$b].Thread) -and (($LogEntryList[$b].Message -Match '^\[PowerShell\] Processing user session (.*), user id = (.*)$') -or ($LogEntryList[$b].Message -Match '^\[PowerShell\] Get (.*) policies for user (.*) in session (.*)$') -or ($LogEntryList[$b].Message -Match '^\[Win32App\] Got (.*) Win32App(s) for user (.*) in session (.*)$'))) {
								#Write-Host "DEBUG: Found matching line: $($LogEntryList[$b].Message)"
								
								# Save to temporary variables because next -Match will overwrite existing $Matches -variable
								$Matches1 = $Matches[1]
								$Matches2 = $Matches[2]
								
								# Either one is GUID so we need to test which $Matches variable is valid GUID and which is sessionId
								if($Matches1 -Match '[[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}') {
									$UserId = $Matches1
									$UserSession = $Matches2
								} elseif ($Matches2 -Match '[[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}') {
									$UserSession = $Matches1
									$UserId = $Matches2
								} else {
									# Something strange happened and we did not get GUID
									# We should not be here
									$UserSession = $null
									$UserId = $null
								}

								#Write-Host "`$UserId=$UserId"
								#Write-Host "`$UserSession=$UserSession"
								#Write-Host "`$ImpersonationSessionId=$ImpersonationSessionId"
								
								if($ImpersonationSessionId -eq $UserSession) {
									# We are quite sure we have right id for UserName
									
									# UserSID object does not exist in Hashtable
									# Create object which we will add to Hashtable
									if($IdHashtable.ContainsKey($UserId)) {
										$IdHashtable[$UserId].id = $UserId
										$IdHashtable[$UserId].userId = $UserId
										$IdHashtable[$UserId].name = $LoggedOnUser
										$IdHashtable[$UserId].displayName = $LoggedOnUser
										$IdHashtable[$UserId].session = $UserSession

									} else {
										$UserIdCustomObjectProperties = @{
											id = $UserId
											userId = $UserId
											userSID = $null
											name = $LoggedOnUser
											displayName = $LoggedOnUser
											session = $UserSession
										}
										$UserIdCustomObject = New-Object -TypeName PSObject -Prop $UserIdCustomObjectProperties

										# Create new UserId hashtable entry
										$IdHashtable.Add($UserId , $UserIdCustomObject)									
									}
									
									#Write-Host "DEBUG: We should have UserID and Username`n$($IdHashtable | ConvertTo-Json)"
								}
								
								# Break out from While-loop
								Break
							}
							$b++
						}
					}
				}

				
				# Continue to next line in Do-while -loop
				$i++
				Continue
			}
		}
		# endregion Find Logged on user
		
		
		# region Find Powershell Script Start log entry
		if($LogEntryList[$i].Message -like '`[PowerShell`] Processing policy with id = *') {
			if($LogEntryList[$i].Message -Match '^\[PowerShell\] Processing policy with id = (.*) for user (.*)$') {
				$PowershellScriptPolicyIdStart = $Matches[1]
				$PowershellScriptForUserIdStart = $Matches[2]
				
				$PowershellScriptStartLogEntryIndex=$i

				# This variables is probably not needed because
				# we can detect end and start log entries separately
				$PowershellScriptStartLogEntryFound=$True
				
				if($PowershellScriptForUserIdStart -eq '00000000-0000-0000-0000-000000000000') {
					$PowershellScriptForUserNameStart = 'System'
				} else {

					# Get real name if found from Hashtable
					if($IdHashtable.ContainsKey($PowershellScriptForUserIdStart)) {
						$Name = $IdHashtable[$PowershellScriptForUserIdStart].name
						if($Name) {
							$PowershellScriptForUserNameStart = $IdHashtable[$PowershellScriptForUserIdStart].name
						} else {
							$PowershellScriptForUserNameStart = $PowershellScriptForUserIdStart
						}
					} else {
						$PowershellScriptForUserNameStart = $PowershellScriptForUserIdStart
					}
				}
				
				# Get Powershell script name if parameter -Online was specified
				if($IdHashtable.ContainsKey($PowershellScriptPolicyIdStart)) {
					$PowershellScriptPolicyNameStart = $IdHashtable[$PowershellScriptPolicyIdStart].name
				} else {
					$PowershellScriptPolicyNameStart = $PowershellScriptPolicyIdStart
				}
				
				if($ShowAllTimelineEntries) {

					RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Process / Powershell" -detail "Powershell script     :  $PowershellScriptPolicyNameStart for user $PowershellScriptForUserNameStart" -logEntry "Line $($LogEntryList[$i].Line)"
				}
						
				# Set Powershell Script Start information to ProcessRuntime property
				$LogEntryList[$i].ProcessRunTime = "Processing Powershell script ($PowershellScriptPolicyNameStart) for user $PowershellScriptForUserNameStart"

				# Continue to next line in Do-while -loop
				$i++
				Continue
			}
		}
		# endregion Find Powershell Script Start log entry

		# region Find Powershell Script End log entry
		if($LogEntryList[$i].Message -like '`[PowerShell`] User Id = *') {
			if($LogEntryList[$i].Message -Match '^\[PowerShell\] User Id = (.*), Policy id = (.*), policy result = (.*)$') {
				$PowershellScriptForUserIdEndProcess = $Matches[1]
				$PowershellScriptPolicyIdEndProcess = $Matches[2]
				$PowershellScriptPolicyResultEndProcess = $Matches[3]
				
				# This variables is probably not needed because
				# we can detect end and start log entries separately
				$PowershellScriptEndLogEntryFound=$True
				
				if($PowershellScriptForUserIdEndProcess -eq '00000000-0000-0000-0000-000000000000') {
					$PowershellScriptForUserNameEndProcess = 'System'
				} else {
					$PowershellScriptForUserNameEndProcess = $PowershellScriptForUserIdEndProcess
					
					# Get real name if found from Hashtable
					if($IdHashtable.ContainsKey($PowershellScriptForUserIdEndProcess)) {
						$Name = $IdHashtable[$PowershellScriptForUserIdEndProcess].name
						if($Name) {
							$PowershellScriptForUserNameEndProcess = $IdHashtable[$PowershellScriptForUserIdEndProcess].name
						}
					}
				}
				
				# Get Powershell script name if parameter -Online was specified
				if($IdHashtable.ContainsKey($PowershellScriptPolicyIdEndProcess)) {
					$PowershellScriptPolicyNameEnd = $IdHashtable[$PowershellScriptPolicyIdEndProcess].name
				} else {
					$PowershellScriptPolicyNameEnd = $PowershellScriptPolicyIdEndProcess
				}
				
				# Get Process runtime
				if($PowershellScriptStartLogEntryIndex) {
					$PowershellScriptStartTime = $LogEntryList[$PowershellScriptStartLogEntryIndex].LogEntryDateTime
					$PowershellScriptEndTime = $LogEntryList[$i].LogEntryDateTime

					$PowershellScriptRunTime = New-Timespan $PowershellScriptStartTime $PowershellScriptEndTime
					$PowershellScriptRunTimeTotalSeconds = $PowershellScriptRunTime.TotalSeconds


				} else {
					$PowershellScriptRunTimeTotalSeconds = 'N/A'
				}

				if($PowershellScriptPolicyResultEndProcess -eq 'Success') {
					
					RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Success / Powershell" -detail "Powershell script     :  $PowershellScriptPolicyNameEnd for user $PowershellScriptForUserNameEndProcess" -Seconds $PowershellScriptRunTimeTotalSeconds -logEntry "Line $($LogEntryList[$i].Line)" -Color 'Green'

					# Check if Powershell script runtime is longer than configured (default 180 seconds)
					if($PowershellScriptRunTimeTotalSeconds -gt $LongRunningPowershellNotifyThreshold) {
						
						RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Warning" -detail "                            Long Powershell script runtime found" -Seconds $PowershellScriptRunTimeTotalSeconds -logEntry "Line $($LogEntryList[$i].Line)" -Color 'Yellow'
					}
					
					# Set Powershell Script End information to ProcessRuntime property
					$LogEntryList[$i].ProcessRunTime = "Succeeded Powershell script ($PowershellScriptPolicyNameEnd) for user $PowershellScriptForUserNameEndProcess (Runtime: $PowershellScriptRunTimeTotalSeconds seconds)"
					
				} elseif ($PowershellScriptPolicyResultEndProcess -eq 'Failed') {

					RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Failed  / Powershell" -detail "Powershell script     :  $PowershellScriptPolicyNameEnd for user $PowershellScriptForUserNameEndProcess" -Seconds $PowershellScriptRunTimeTotalSeconds -logEntry "Line $($LogEntryList[$i].Line)" -Color 'Red'

					# Set Powershell Script End information to ProcessRuntime property
					$LogEntryList[$i].ProcessRunTime = "Failed Powershell script ($PowershellScriptPolicyNameEnd) for user $PowershellScriptForUserNameEndProcess (Runtime: $PowershellScriptRunTimeTotalSeconds seconds)"

					# Check if there is error message for failed Powershell script
					
					# Check if next log entry exists
					if($i + 1 -le $LogEntryList.Count - 1) {
						if($LogEntryList[$i+1].Message -Match '^\[PowerShell\] Fail, the details are (.*)$') {
							$PowershellFailJsonString = $Matches[1] 

							# Try to convert JsonString
							Try {
								$PowershellFailFromJson=$null
								$PowershellFailFromJson = $PowershellFailJsonString | ConvertFrom-Json
								if($ExecutionMsg=$PowershellFailFromJson.ExecutionMsg) {
									if($ExecutionMsg) {
										# This makes time line harder to read so moved Error message after timeline

										$PowershellScriptErrorLogs += "$($LogEntryList[$i].DateTime) End   Powershell script ($PowershellScriptPolicyNameEnd) for user $PowershellScriptForUserNameEndProcess Success: $PowershellScriptPolicyResultEndProcess (Runtime: $PowershellScriptRunTimeTotalSeconds seconds) ($($LogEntryList[$i].FileName) line $($LogEntryList[$i].Line))"
										$PowershellScriptErrorLogs += "$($LogEntryList[$i+1].DateTime) Powershell script error message ($($LogEntryList[$i].FileName) line $($LogEntryList[$i+1].Line))"

										# Split Error Message by newline and enter error message line by line
										$ExecutionMsg.Split("`n") | Foreach-Object {
											$ErrorMessageLine = $_
											#$PowershellScriptErrorLogs += "`t`t`t   $ErrorMessageLine"
											$PowershellScriptErrorLogs += "$ErrorMessageLine"
										}

										if($ShowErrorsInTimeline) {
											$ExecutionMsg.Split("`n") | Foreach-Object {
												$ErrorMessageLine = $_

												RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "ErrorLog" -detail "$ErrorMessageLine" -logEntry "$($LogEntryList[$i].FileName) line $($LogEntryList[$i].Line)" -Color 'Red'
											}
										} else {
											
											RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Info" -detail "Errors found. Use Parameter -ShowErrorsInTimeline to show errors on Timeline" -logEntry "Line $($LogEntryList[$i].Line)" -Color 'Yellow'
											
										}
									}
								}
							} catch {
								# Did not find Powershell error message
							}
						}
					}

				} else {
					
					RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Script ended" -detail "Powershell script $PowershellScriptPolicyNameEnd for user $PowershellScriptForUserNameEndProcess ($PowershellScriptRunTimeTotalSeconds seconds)" -logEntry "Line $($LogEntryList[$i].Line)" -Color 'Yellow'
				}
					
				$PowershellScriptStartLogEntryIndex=$null

				# Continue to next line in Do-while -loop
				$i++
				Continue
			}
		}
		# endregion Find Powershell Script End log entry


<#
		# Remove me from Production
		# Left here for possible future purposes
		# This catches most Powershell script processes and we could for example check Process runtime
		# and alert for long running processes

		# region Find Powershell processess
		# Try to find Powershell process End log entry
		if($LogEntryList[$i].Message -like '*execution is done*') {
			$ProcessStartTime = $null
			$ProcessEndTime = $LogEntryList[$i].LogEntryDateTime
			$Thread = $LogEntryList[$i].Thread
			Write-Verbose "Found Powershell process   Ending message (Line $($LogEntryList[$i].Line) Thread $Thread): $($LogEntryList[$i].Message)"

			# Try to find Powershell process Start log entry (upwards)
			$b = $i - 1
			for($b = $i - 1; $b -ge 0; $b-- ) {
				if(($LogEntryList[$b].Message -like '*process id =*') -and ($LogEntryList[$b].Thread -eq $Thread)) {
					Write-Verbose "Found Powershell process Starting message (Line $($LogEntryList[$b].Line) Thread $($LogEntryList[$b].Thread)): $($LogEntryList[$b].Message)"
					$ProcessStartTime = $LogEntryList[$b].LogEntryDateTime
					$ProcessRunTime = New-Timespan $ProcessStartTime $ProcessEndTime
					$ProcessRunTimeTotalSeconds = $ProcessRunTime.TotalSeconds
					$ProcessRunTimeText = "{0,-16} {1,8:n1}" -f "End Powershell", $ProcessRunTimeTotalSeconds
					
					# Set RunTime to Powershell script Install Start log entry
					$LogEntryList[$b].ProcessRunTime = "{0,-16} {1,8:n1}" -f "Start Powershell", $ProcessRunTimeTotalSeconds
					
					# Set RunTime to Powershell script Install End log entry
					$LogEntryList[$i].ProcessRunTime = "{0,-16} {1,8:n1}" -f "End Powershell", $ProcessRunTimeTotalSeconds


					# Show warning for long running Powershell script
#					if($ProcessRunTimeTotalSeconds -ge $LongRunningPowershellNotifyThreshold) {
#						Write-Host "$($LogEntryList[$b].DateTime) Long running ($ProcessRunTimeTotalSeconds seconds) Powershell script ($LogFileName line $($LogEntryList[$b].Line)): $($LogEntryList[$b].Message)" -ForegroundColor Yellow
#					}


					# Break out from for-loop
					Break
				}
			} # For-loop end
			Write-Verbose ""
			
			# Cleanup variables
			$ProcessStartTime = $null
			$ProcessEndTime = $null
			$Thread = $null
			$ProcessRunTimeTotalSeconds = $null
			$ProcessRunTimeText = $null
		}
		# endregion Find Powershell processess
		# Remove me from Production
#>

		###############################################

		# region Find Proactive Remediation Detect script type - Proactive Remediation or custom Compliance
		if($LogEntryList[$i].Message -like '`[HS`] ProcessScript PolicyId:*') {
			# Proactive Remediation script = policyType 6
			# custom Compliance script = policyType 8

			#Write-Host "Proactive Remediation Detect script type found" -ForegroundColor Yellow
			$Matches=$null
			if($LogEntryList[$i].Message -Match '^\[HS\] ProcessScript PolicyId: (.*) PolicyType: (.*)$') {
				$PolicyId = $Matches[1]
				$PolicyType = $Matches[2]
				
				# Check if PolicyId is in Hashtable
				if($IdHashtable.ContainsKey($PolicyId)) {
					# This can exist if -Online parameter was used
					
					if($IdHashtable[$PolicyId] | Get-Member policyType){ 
						# policyType Property exists in Hashtable object
						$IdHashtable[$PolicyId].policyType = $PolicyType
					} else {
						# policyType Property does NOT exist in Hashtable object
						# Add new Property
						$IdHashtable[$PolicyId] | Add-Member -MemberType noteProperty -Name policyType -Value $PolicyType
					}

				} else {
					# Add new Hashtable id entry
					$PolicyObjectProperties = @{
						id = $PolicyId
						policyType = $PolicyType
					}
					$PolicyCustomObject = New-Object -TypeName PSObject -Prop $PolicyObjectProperties

					# Create new UserId hashtable entry
					$IdHashtable.Add($PolicyId , $PolicyCustomObject)
				}
			}
		}
		# endregion Find Proactive Remediation Detect script type - Proactive Remediation or custom Compliance

		# region Try to find Proactive Detection script run
		if($LogEntryList[$i].Message -like '`[HS`] the pre-rem*diation detection script compliance result for*') {
			#Write-Host "Proactive Remediation or custom Compliance Script Detect found" -ForegroundColor Yellow
			$Matches=$null
			if($LogEntryList[$i].Message -Match '^\[HS\] the pre-remdiation detection script compliance result for (.*) is (.*)$') {
				$PolicyIdEnd = $Matches[1]
				$PreRemediationDetectResult = $Matches[2]

				# Use this PolicyId on next phase -> Remediate
				# Because there we don't have 2 log entries with PolicyId
				$ProactiveRemediationPolicyId = $PolicyIdStart

				$ProcessStartTime = $null
				$DetectionLogEntryEndTime = $LogEntryList[$i].LogEntryDateTime
				$Thread = $LogEntryList[$i].Thread

				# Try to find displayName
				if($IdHashtable.ContainsKey($PolicyIdEnd)) {
					if($IdHashtable[$PolicyIdEnd].displayName) {
						$ProactiveRemediationScriptName = $IdHashtable[$PolicyIdEnd].displayName
					} else {
						$ProactiveRemediationScriptName = $PolicyIdEnd
					}
				}
				
				# Try to find policyType
				if($IdHashtable.ContainsKey($PolicyIdEnd)) {
					$ProactiveRemediationScriptPolicyType = $IdHashtable[$PolicyIdEnd].policyType
				}

				if($ProactiveRemediationScriptPolicyType -eq 6) {
					$ProactiveRemediationScriptPolicyTypeName = 'Proactive Remediation'
				} elseif ($ProactiveRemediationScriptPolicyType -eq 8) {
					$ProactiveRemediationScriptPolicyTypeName = 'Custom Compliance'
				} else {
					# We should never get here because we have seen polityType log entry before this log entry
					# Proactive Remediation script type is unknown
					$ProactiveRemediationScriptPolicyTypeName = 'Proactive Remediation script (unknown type)'
				}

				# Check if next log entry is either
				#	[HS] remediation is not optional, kick off remediation script
				#	[HS] no remediation script, skip remediation script
				if(($i+1 -lt $LogEntryList.Count) -and (($LogEntryList[$i+1].Message -like '`[HS`] remediation is not optional, kick off remediation script') -or ($LogEntryList[$i+1].Message -like '`[HS`] no remediation script, skip remediation script'))) {
					$RemediationDetectPostActionMessageIndex = $i + 1
				} else {
					$RemediationDetectPostActionMessageIndex = $null
				}

				Write-Verbose "Found $ProactiveRemediationScriptPolicyTypeName ($ProactiveRemediationScriptName) Detect Ending message ($($LogEntryList[$i].FileName) Line $($LogEntryList[$i].Line) Thread $Thread): $($LogEntryList[$i].Message)"

				$ProactiveRemediationDetectStdOutput = $null
				$ProactiveRemediationDetectStdOutIndex = $null
				$ProactiveRemediationDetectErrorOutput = $null
				$ProactiveRemediationDetectErrorOutputIndex = $null

				$PowershellScriptStartTime = $null
				$PowershellScriptEndTime = $null
				$PowershellScriptStartIndex = $null
				$PowershellScriptEndIndex = $null

				$ProactiveRemediationDetectProcessId = $null
				$ProcessRunTimeTotalSeconds = $null
				$ProactiveRemediationDetectRunAs = 'N/A'

				# Try to find Proactive Remediation Detect Start log entry (going upwards)
				$b = $i - 1
				for($b = $i - 1; $b -ge 0; $b-- ) {
					if(($LogEntryList[$b].Message -Match '^\[HS\] std output =(.*)$') -and ($LogEntryList[$b].Thread -eq $Thread)) {
						$ProactiveRemediationDetectStdOutput = ($Matches[1]).Trim()
						$ProactiveRemediationDetectStdOutIndex = $b
						# Continue to previous log entry
						Continue
					}

					if(($LogEntryList[$b].Message -Match '^\[HS\] err output =(.*)$') -and ($LogEntryList[$b].Thread -eq $Thread)) {
						$ProactiveRemediationDetectErrorOutput = ($Matches[1]).Trim()
						$ProactiveRemediationDetectErrorOutputIndex = $b
						# Continue to previous log entry
						Continue
					}

					if(($LogEntryList[$b].Message -Match '^Powershell execution is done, exitCode = (.*)$') -and ($LogEntryList[$b].Thread -eq $Thread)) {
						$ProactiveRemediationDetectExitCode = $Matches[1]
						
						# Set RunTime to Powershell script Install End log entry
						$LogEntryList[$b].ProcessRunTime = "End Proactive Detect Powershell (ExitCode=$ProactiveRemediationDetectExitCode)"

						$PowershellScriptEndTime = $LogEntryList[$b].LogEntryDateTime
						$PowershellScriptEndIndex = $b
						
						# Continue to previous log entry
						Continue
					}

					if(($LogEntryList[$b].Message -Match '^process id = (.*)$') -and ($LogEntryList[$b].Thread -eq $Thread)) {
						$ProactiveRemediationDetectProcessId = $Matches[1]
						
						# Set RunTime to Powershell script Install End log entry
						$LogEntryList[$b].ProcessRunTime = "Start Proactive Detect Powershell"

						$PowershellScriptStartTime = $LogEntryList[$b].LogEntryDateTime
						
						# Calculate Powershell script runtime
						if($PowershellScriptStartTime -and $PowershellScriptEndTime) {
							$PowershellScriptRuntime = New-Timespan $PowershellScriptStartTime $PowershellScriptEndTime
							$ProcessRunTimeTotalSeconds = $PowershellScriptRuntime.TotalSeconds
						} else {
							$ProcessRunTimeTotalSeconds = 'N/A'
						}
						
						$PowershellScriptStartIndex = $b

						# This if seems not to work
						# Left here if it would be fixed at some time
						if($ProactiveRemediationScriptPolicyType -eq 8) {
							# Custom Compliance Script
							
							# Set RunTime to Powershell script Start log entry
							$LogEntryList[$b].ProcessRunTime = "{0,-16} {1,8:n1}" -f "       Start Custom Compliance Powershell", $ProcessRunTimeTotalSeconds

							# Add RunTime to Powershell script End log entry
							$LogEntryList[$PowershellScriptEndIndex].ProcessRunTime = "{0,-16} {1,8:n1}" -f "       End Custom Compliance Powershell", $ProcessRunTimeTotalSeconds
						} else {
							# Proactive Remediation script
							
							# Set RunTime to Powershell script Start log entry
							$LogEntryList[$b].ProcessRunTime = "{0,-16} {1,8:n1}" -f "       Start Proactive Detect Powershell", $ProcessRunTimeTotalSeconds

							# Add RunTime to Powershell script End log entry
							$LogEntryList[$PowershellScriptEndIndex].ProcessRunTime = "{0,-16} {1,8:n1}" -f "       End Proactive Detect Powershell", $ProcessRunTimeTotalSeconds
						}

						# Continue to previous log entry
						Continue
					}

					if(($LogEntryList[$b].Message -Match '^\[HS\] The policy needs be run as (.*)$') -and ($LogEntryList[$b].Thread -eq $Thread)) {
						$ProactiveRemediationDetectRunAs = $Matches[1]
						
						# Continue to previous log entry
						Continue
					}

					# Custom Compliance policy name is shorter than Proactive Remediations
					# Add padding so we can align texts
					# And Yes, I should use -f formatting on these but maybe next time :)
					if($ProactiveRemediationScriptPolicyType -eq 8) {
						$Padding = '    '
					} else {
						$Padding = ''
					}
					
					if(($LogEntryList[$b].Message -Match '^\[HS\] Processing policy with id = (.*)$') -and ($LogEntryList[$b].Thread -eq $Thread)) {
						$PolicyIdStart = $Matches[1]
					
						# $b is startIndex
							
						if($PolicyIdStart -eq $PolicyIdEnd) {
							# We found matching starting log entry
							$DetectionLogEntryStartTime = $LogEntryList[$b].LogEntryDateTime
							if($DetectionLogEntryStartTime -and $DetectionLogEntryEndTime) {
								$ProactiveDetectionRuntime = New-Timespan $DetectionLogEntryStartTime $DetectionLogEntryEndTime
								$ProcessRunTimeTotalSeconds = $ProactiveDetectionRuntime.TotalSeconds
							} else {
								$ProcessRunTimeTotalSeconds = 'N/A'
							}



							# Start message to timeline
							if($ShowAllTimelineEntries) {
								RecordStatusToTimeline -date $LogEntryList[$b].DateTime -status "Process / Detect" -detail "$ProactiveRemediationScriptPolicyTypeName $($Padding):  $ProactiveRemediationScriptName" -Seconds $ProcessRunTimeTotalSeconds -logEntry "Line $($LogEntryList[$b].Line)"
							}
							
							# Set RunTime log entry
							$LogEntryList[$b].ProcessRunTime = "Processing $ProactiveRemediationScriptPolicyTypeName Detect script ($ProactiveRemediationScriptName) (run as $ProactiveRemediationDetectRunAs)"


							# End message to Timeline
							if($PreRemediationDetectResult -eq 'True') {
								
								RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Success / Detect" -detail "$ProactiveRemediationScriptPolicyTypeName $($Padding):  $ProactiveRemediationScriptName" -Seconds $ProcessRunTimeTotalSeconds -Color 'Green' -logEntry "Line $($LogEntryList[$i].Line)"
								
								# Set RunTime log entry
								$LogEntryList[$i].ProcessRunTime = "Succeeded $ProactiveRemediationScriptPolicyTypeName Detect script ($ProactiveRemediationScriptName) (Runtime: $ProcessRunTimeTotalSeconds sec)"
	
							} else {
								
								RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Failed  / Detect" -detail "$ProactiveRemediationScriptPolicyTypeName $($Padding):  $ProactiveRemediationScriptName" -Seconds $ProcessRunTimeTotalSeconds -Color 'Yellow' -logEntry "Line $($LogEntryList[$i].Line)"
								
								$LogEntryList[$i].ProcessRunTime = "Failed $ProactiveRemediationScriptPolicyTypeName Detect script ($ProactiveRemediationScriptName) (Runtime: $ProcessRunTimeTotalSeconds sec)"
							}


							# Script StdOut to Timeline if configured and if exists
							if($ShowStdOutInTimeline -and $ProactiveRemediationDetectStdOutIndex) {
																
								RecordStatusToTimeline -date $LogEntryList[$ProactiveRemediationDetectStdOutIndex].DateTime -status "StdOut" -detail "$ProactiveRemediationDetectStdOutput" -logEntry "Line $($LogEntryList[$ProactiveRemediationDetectStdOutIndex].Line)" -Color 'Yellow'
								
								$LogEntryList[$ProactiveRemediationDetectStdOutIndex].ProcessRunTime = "       Powershell Detect StdOut"
							}

							# Script Error to Timeline if configured and if exists
							if($ShowErrorsInTimeline -and $ProactiveRemediationDetectErrorOutputIndex) {
								
								RecordStatusToTimeline -date $LogEntryList[$ProactiveRemediationDetectErrorOutputIndex].DateTime -status "ErrorLog" -detail "$ProactiveRemediationDetectErrorOutput" -logEntry "Line $($LogEntryList[$ProactiveRemediationDetectStdOutIndex].Line)" -Color 'Yellow'
								
								$LogEntryList[$ProactiveRemediationDetectErrorOutputIndex].ProcessRunTime = "       Powershell Detect ErrorLog"
								
							}

							# Post-End message to Timeline if exists
							if($RemediationDetectPostActionMessageIndex) {
								
								$RemediationDetectPostActionMessage = $LogEntryList[$RemediationDetectPostActionMessageIndex].Message
								if($RemediationDetectPostActionMessage -eq '[HS] no remediation script, skip remediation script') {
									$RemediationDetectPostActionMessage = '                            remediation script does not exist'
								}

								if($RemediationDetectPostActionMessage -eq '[HS] remediation is not optional, kick off remediation script') {
									$RemediationDetectPostActionMessage = '                            kick off remediation script'
								}


								RecordStatusToTimeline -date $LogEntryList[$RemediationDetectPostActionMessageIndex].DateTime -status "Info" -detail "$RemediationDetectPostActionMessage" -logEntry "Line $($LogEntryList[$RemediationDetectPostActionMessageIndex].Line)" -Color 'Yellow'
							}

							
						} else {
							# We missed Proactive Remediation Detect start log entry
							# We should never get here
							
							if($PreRemediationDetectResult -eq 'True') {
								
								RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Success / Detect" -detail "$ProactiveRemediationScriptPolicyTypeName $($Padding):  $ProactiveRemediationScriptName" -Color 'Green' -logEntry "Line $($LogEntryList[$i].Line)"
								
								# Set RunTime log entry
								$LogEntryList[$i].ProcessRunTime = "Succeeded $ProactiveRemediationScriptPolicyTypeName Detect script ($ProactiveRemediationScriptName)"
	
							} else {
								RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Failed  / Detect" -detail "$ProactiveRemediationScriptPolicyTypeName $($Padding):  $ProactiveRemediationScriptName" -Color 'Yellow' -logEntry "Line $($LogEntryList[$i].Line)"
								
								$LogEntryList[$i].ProcessRunTime = "Failed/NotCompliant $ProactiveRemediationScriptPolicyTypeName Detect script ($ProactiveRemediationScriptName)"
							}
							
							# Script StdOut to Timeline if configured and if exists
							if($ShowStdOutInTimeline -and $ProactiveRemediationDetectStdOutIndex) {
																
								RecordStatusToTimeline -date $LogEntryList[$ProactiveRemediationDetectStdOutIndex].DateTime -status "StdOut" -detail "$ProactiveRemediationDetectStdOutput" -logEntry "Line $($LogEntryList[$ProactiveRemediationDetectStdOutIndex].Line)" -Color 'Yellow'
								
								$LogEntryList[$ProactiveRemediationDetectStdOutIndex].ProcessRunTime = "       Powershell Detect StdOut"
							}

							# Script Error to Timeline if configured and if exists
							if($ShowErrorsInTimeline -and $ProactiveRemediationDetectErrorOutputIndex) {
								
								RecordStatusToTimeline -date $LogEntryList[$ProactiveRemediationDetectErrorOutputIndex].DateTime -status "ErrorLog" -detail "$ProactiveRemediationDetectErrorOutput" -logEntry "Line $($LogEntryList[$ProactiveRemediationDetectStdOutIndex].Line)" -Color 'Yellow'
								
								$LogEntryList[$ProactiveRemediationDetectErrorOutputIndex].ProcessRunTime = "       Powershell Detect ErrorLog"
								
							}

							# Post-End message to Timeline if exists
							if($RemediationDetectPostActionMessageIndex) {
								
								RecordStatusToTimeline -date $LogEntryList[$RemediationDetectPostActionMessageIndex].DateTime -status "Info" -detail "                         $($LogEntryList[$RemediationDetectPostActionMessageIndex].Message)" -logEntry "Line $($LogEntryList[$RemediationDetectPostActionMessageIndex].Line)" -Color 'Yellow'
							}
							
							Break
						}
						
						# We are done
						
						# Break out from For-loop
						Break
					}

					
					if($LogEntryList[$b - 1].Message -like '`[HS`] the pre-remdiation detection script compliance result for*') {
						# We missed Proactive Remediation Detect start log entry
						
						#Write-Host "Missed Proactive Remedation Detect script start log entry" -ForegroundColor Red
						
						if($PreRemediationDetectResult -eq 'True') {
							
							RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Success / Detect" -detail "$ProactiveRemediationScriptPolicyTypeName $($Padding):  $ProactiveRemediationScriptName" -Color 'Green' -logEntry "Line $($LogEntryList[$i].Line)"
							
							# Set RunTime log entry
							$LogEntryList[$i].ProcessRunTime = "Succeeded $ProactiveRemediationScriptPolicyTypeName Detection script ($ProactiveRemediationScriptName)"

						} else {
							RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Failed  / Not Compliant" -detail "$ProactiveRemediationScriptPolicyTypeName $($Padding):  $ProactiveRemediationScriptName" -Color 'Yellow' -logEntry "Line $($LogEntryList[$i].Line)"
							
							$LogEntryList[$i].ProcessRunTime = "Failed/NotCompliant $ProactiveRemediationScriptPolicyTypeName Detection script ($ProactiveRemediationScriptName)"
						}
						
						Break
					}
					
				}
			} # For-loop end
			Write-Verbose ""

		}
		# endregion Try to find Proactive Detection script run


		# region Try to find Proactive Remediate script run
		if($LogEntryList[$i].Message -like '`[HS`] remediation script exit code is*') {
			#Write-Host "DEBUG: Proactive Remediation Remediate Script found" -ForegroundColor Yellow
			$Matches=$null
			if($LogEntryList[$i].Message -Match '\[HS\] remediation script exit code is (.*)$') {
				$RemediationScriptExitCode = $Matches[1]

				#Write-Host "DEBUG: $($LogEntryList[$i].Message)"

				$ProcessStartTime = $null
				$RemediationScriptLogEntryEndTime = $LogEntryList[$i].LogEntryDateTime
				$Thread = $LogEntryList[$i].Thread


				$RemediationScriptStdOutput = $null
				$RemediationScriptStdOutputIndex = $null
				$RemediationScriptErrorOutput = $null
				$RemediationScriptErrorOutputIndex = $null

				$PowershellScriptStartTime = $null
				$PowershellScriptEndTime = $null
				$PowershellScriptStartIndex = $null
				$PowershellScriptEndIndex = $null

				$RemediationPowershellScriptExitCode = $null

				$RemediationPowershellScriptProcessId = $null
				$ProcessRunTimeTotalSeconds = $null
				$RemediationScriptRunAs = 'N/A'

				# Try to find Proactive Remediation Remediate Start log entry (going upwards)
				$b = $i - 1
				for($b = $i - 1; $b -ge 0; $b-- ) {
					
					#Write-Host "DEBUG: $($LogEntryList[$b].Message)"
					
					if(($LogEntryList[$b].Message -Match '^\[HS\] std output =(.*)$') -and ($LogEntryList[$b].Thread -eq $Thread)) {
						$RemediationScriptStdOutput = ($Matches[1]).Trim()
						$RemediationScriptStdOutputIndex = $b
						# Continue to previous log entry
						Continue
					}

					if(($LogEntryList[$b].Message -Match '^\[HS\] err output =(.*)$') -and ($LogEntryList[$b].Thread -eq $Thread)) {
						$RemediationScriptErrorOutput = ($Matches[1]).Trim()
						$RemediationScriptErrorOutputIndex = $b
						# Continue to previous log entry
						Continue
					}

					if(($LogEntryList[$b].Message -Match '^Powershell execution is done, exitCode = (.*)$') -and ($LogEntryList[$b].Thread -eq $Thread)) {
						$RemediationPowershellScriptExitCode = $Matches[1]
						
						# Set RunTime to Powershell script Install End log entry
						$LogEntryList[$b].ProcessRunTime = "End Proactive Remediate Powershell (ExitCode=$RemediationPowershellScriptExitCode)"

						$PowershellScriptEndTime = $LogEntryList[$b].LogEntryDateTime
						$PowershellScriptEndIndex = $b
						
						# Continue to previous log entry
						Continue
					}

					if(($LogEntryList[$b].Message -Match '^process id = (.*)$') -and ($LogEntryList[$b].Thread -eq $Thread)) {
						$RemediationPowershellScriptProcessId = $Matches[1]

						$PowershellScriptStartTime = $LogEntryList[$b].LogEntryDateTime
						
						# Calculate Powershell script runtime
						if($PowershellScriptStartTime -and $PowershellScriptEndTime) {
							$PowershellScriptRuntime = New-Timespan $PowershellScriptStartTime $PowershellScriptEndTime
							$ProcessRunTimeTotalSeconds = $PowershellScriptRuntime.TotalSeconds
						} else {
							$ProcessRunTimeTotalSeconds = 'N/A'
						}
						
						$PowershellScriptStartIndex = $b

						# Set RunTime to Powershell script Start log entry
						$LogEntryList[$b].ProcessRunTime = "{0,-16} {1,8:n1}" -f "       Start Proactive Remediate Powershell", $ProcessRunTimeTotalSeconds

						# Add RunTime to Powershell script End log entry
						$LogEntryList[$PowershellScriptEndIndex].ProcessRunTime = "{0,-16} {1,8:n1}" -f "       End Proactive Remediate Powershell", $ProcessRunTimeTotalSeconds
						
						# Continue to previous log entry
						Continue
					}

					if(($LogEntryList[$b].Message -Match '^\[HS\] The policy needs be run as (.*)$') -and ($LogEntryList[$b].Thread -eq $Thread)) {
						$RemediationScriptRunAs = $Matches[1]
						
						# Continue to previous log entry
						Continue
					}
					
					if(($LogEntryList[$b].Message -Match '^\[HS\] Processing policy with id = (.*)$') -and ($LogEntryList[$b].Thread -eq $Thread)) {
						$PolicyIdStart = $Matches[1]
					
						# $b is startIndex
						
						# Try to find displayName
						if($IdHashtable.ContainsKey($PolicyIdStart)) {
							if($IdHashtable[$PolicyIdStart].displayName) {
								$ProactiveRemediationScriptName = $IdHashtable[$PolicyIdStart].displayName
							} else {
								$ProactiveRemediationScriptName = $PolicyIdStart
							}
						}


						$RemediationScriptStartTime = $LogEntryList[$b].LogEntryDateTime
						if($RemediationScriptStartTime -and $RemediationScriptLogEntryEndTime) {
							$ProactiveRemediateRuntime = New-Timespan $RemediationScriptStartTime $RemediationScriptLogEntryEndTime
							$ProcessRunTimeTotalSeconds = $ProactiveDetectionRuntime.TotalSeconds
						} else {
							$ProcessRunTimeTotalSeconds = 'N/A'
						}

						#Write-Host "DEBUG: `$ProcessRunTimeTotalSeconds=$ProcessRunTimeTotalSeconds"

						$Padding = ''

						# Start message to Timeline
						if($ShowAllTimelineEntries) {
							RecordStatusToTimeline -date $LogEntryList[$b].DateTime -status "Process / Remediate" -detail "$ProactiveRemediationScriptPolicyTypeName $($Padding):  $ProactiveRemediationScriptName (run as $RemediationScriptRunAs)" -logEntry "Line $($LogEntryList[$b].Line)"
						}
						
						# Set RunTime log entry
						$LogEntryList[$b].ProcessRunTime = "Processing $ProactiveRemediationScriptPolicyTypeName Remediate script ($ProactiveRemediationScriptName) (run as $RemediationScriptRunAs)"

						# End message to Timeline
						if($RemediationScriptExitCode -eq 0) {
							
							RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Success / Remediate" -detail "$ProactiveRemediationScriptPolicyTypeName $($Padding):  $ProactiveRemediationScriptName" -Seconds $ProcessRunTimeTotalSeconds -Color 'Green' -logEntry "Line $($LogEntryList[$i].Line)"
							
							# Set RunTime log entry
							$LogEntryList[$i].ProcessRunTime = "Succeeded $ProactiveRemediationScriptPolicyTypeName Detection script ($ProactiveRemediationScriptName) (Runtime: $ProcessRunTimeTotalSeconds sec)"

						} else {
															
							RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Failed  / Remediate" -detail "$ProactiveRemediationScriptPolicyTypeName $($Padding):  $ProactiveRemediationScriptName" -Seconds $ProcessRunTimeTotalSeconds -Color 'Red' -logEntry "Line $($LogEntryList[$i].Line)"
							
							$LogEntryList[$i].ProcessRunTime = "Failed $ProactiveRemediationScriptPolicyTypeName Remediate script ($ProactiveRemediationScriptName) (Runtime: $ProcessRunTimeTotalSeconds sec)"
						}

						# Script StdOut to Timeline if configured and if exists
						if($ShowStdOutInTimeline -and $RemediationScriptStdOutputIndex) {
															
							RecordStatusToTimeline -date $LogEntryList[$RemediationScriptStdOutputIndex].DateTime -status "StdOut" -detail "$RemediationScriptStdOutput" -logEntry "Line $($LogEntryList[$RemediationScriptStdOutputIndex].Line)" -Color 'Yellow'
							
							$LogEntryList[$RemediationScriptStdOutputIndex].ProcessRunTime = "       Powershell Remediate StdOut"
						}

						# Script Error to Timeline if configured and if exists
						if($ShowErrorsInTimeline -and $RemediationScriptErrorOutputIndex) {
															
							RecordStatusToTimeline -date $LogEntryList[$RemediationScriptErrorOutputIndex].DateTime -status "ErrorLog" -detail "$RemediationScriptErrorOutput" -logEntry "Line $($LogEntryList[$RemediationScriptErrorOutputIndex].Line)" -Color 'Yellow'
							
							$LogEntryList[$RemediationScriptErrorOutput].ProcessRunTime = "       Powershell Remediate ErrorLog"
							
						}
					
						# Break out from For-loop
						Break					
					}

				} # For-loop end

					
				if($LogEntryList[$b - 1].Message -like '`[HS`] the pre-remdiation detection script compliance result for*') {
					# We missed Proactive Remediation Detect start log entry
					
					#Write-Host "Missed Proactive Remedation Detect script start log entry" -ForegroundColor Red
					
					if($PreRemediationDetectResult -eq 'True') {
						
						RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Succeeded" -detail "$ProactiveRemediationScriptPolicyTypeName $($Padding):  $ProactiveRemediationScriptName" -Color 'Green' -logEntry "Line $($LogEntryList[$i].Line)"
						
						# Set RunTime log entry
						$LogEntryList[$i].ProcessRunTime = "Succeeded $ProactiveRemediationScriptPolicyTypeName Detection script ($ProactiveRemediationScriptName)"

					} else {
						RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Failed  / Not Compliant" -detail "$ProactiveRemediationScriptPolicyTypeName $($Padding):  $ProactiveRemediationScriptName" -Color 'Yellow' -logEntry "Line $($LogEntryList[$i].Line)"
						
						$LogEntryList[$i].ProcessRunTime = "Failed/NotCompliant $ProactiveRemediationScriptPolicyTypeName Detection script ($ProactiveRemediationScriptName)"
					}
					
					# Script StdOut to Timeline if configured and if exists
					if($ShowStdOutInTimeline -and $RemediationScriptStdOutputIndex) {
														
						RecordStatusToTimeline -date $LogEntryList[$RemediationScriptStdOutputIndex].DateTime -status "StdOut" -detail "$RemediationScriptStdOutput" -logEntry "Line $($LogEntryList[$RemediationScriptStdOutputIndex].Line)" -Color 'Yellow'
						
						$LogEntryList[$RemediationScriptStdOutputIndex].ProcessRunTime = "       Powershell Remediate StdOut"
					}

					# Script Error to Timeline if configured and if exists
					if($ShowErrorsInTimeline -and $RemediationScriptErrorOutputIndex) {
														
						RecordStatusToTimeline -date $LogEntryList[$RemediationScriptErrorOutputIndex].DateTime -status "ErrorLog" -detail "$RemediationScriptErrorOutput" -logEntry "Line $($LogEntryList[$RemediationScriptErrorOutputIndex].Line)" -Color 'Yellow'
						
						$LogEntryList[$RemediationScriptErrorOutput].ProcessRunTime = "       Powershell Remediate ErrorLog"
						
					}
					
					Break
				} 
			} 
		} 
		# endregion Try to find Proactive Remediate script run

		# region Try to find Proactive PostDetection script run
		if($LogEntryList[$i].Message -like '`[HS`] the post detection script result for *') {
			#Write-Host "Proactive Remediation PostDetect found" -ForegroundColor Yellow
			
			$Matches=$null
			if($LogEntryList[$i].Message -Match '^\[HS\] the post detection script result for (.*) is (.*)$') {
				$PolicyIdEnd = $Matches[1]
				$PostRemediationDetectResult = $Matches[2]

				$ProcessStartTime = $null
				$DetectionLogEntryEndTime = $LogEntryList[$i].LogEntryDateTime
				$Thread = $LogEntryList[$i].Thread

				# Try to find displayName
				if($IdHashtable.ContainsKey($PolicyIdEnd)) {
					if($IdHashtable[$PolicyIdEnd].displayName) {
						$ProactiveRemediationScriptName = $IdHashtable[$PolicyIdEnd].displayName
					} else {
						$ProactiveRemediationScriptName = $PolicyIdEnd
					}
				}
				
				# Try to find policyType
				if($IdHashtable.ContainsKey($PolicyIdEnd)) {
					$ProactiveRemediationScriptPolicyType = $IdHashtable[$PolicyIdEnd].policyType
				}

				if($ProactiveRemediationScriptPolicyType -eq 6) {
					$ProactiveRemediationScriptPolicyTypeName = 'Proactive Remediation'
				} elseif ($ProactiveRemediationScriptPolicyType -eq 8) {
					$ProactiveRemediationScriptPolicyTypeName = 'Custom Compliance'
				} else {
					# We should never get here because we have seen polityType log entry before this log entry
					# Proactive Remediation script type is unknown
					$ProactiveRemediationScriptPolicyTypeName = 'Proactive Remediation script (unknown type)'
				}

				Write-Verbose "Found $ProactiveRemediationScriptPolicyTypeName ($ProactiveRemediationScriptName) Detect Ending message ($($LogEntryList[$i].FileName) Line $($LogEntryList[$i].Line) Thread $Thread): $($LogEntryList[$i].Message)"

				$ProactiveRemediationPostDetectStdOutput = $null
				$ProactiveRemediationPostDetectStdOutIndex = $null
				$ProactiveRemediationPostDetectErrorOutput = $null
				$ProactiveRemediationPostDetectErrorOutputIndex = $null

				$PowershellScriptStartTime = $null
				$PowershellScriptEndTime = $null
				$PowershellScriptStartIndex = $null
				$PowershellScriptEndIndex = $null

				$ProactiveRemediationPostDetectProcessId = $null
				$ProcessRunTimeTotalSeconds = $null
				$ProactiveRemediationPostDetectRunAs = 'N/A'

				# Try to find Proactive Remediation Detect Start log entry (going upwards)
				$b = $i - 1
				for($b = $i - 1; $b -ge 0; $b-- ) {
					if(($LogEntryList[$b].Message -Match '^\[HS\] std output =(.*)$') -and ($LogEntryList[$b].Thread -eq $Thread)) {
						$ProactiveRemediationPostDetectStdOutput = ($Matches[1]).Trim()
						$ProactiveRemediationPostDetectStdOutIndex = $b
						# Continue to previous log entry
						Continue
					}

					if(($LogEntryList[$b].Message -Match '^\[HS\] err output =(.*)$') -and ($LogEntryList[$b].Thread -eq $Thread)) {
						$ProactiveRemediationPostDetectErrorOutput = ($Matches[1]).Trim()
						$ProactiveRemediationPostDetectErrorOutputIndex = $b
						# Continue to previous log entry
						Continue
					}

					if(($LogEntryList[$b].Message -Match '^Powershell execution is done, exitCode = (.*)$') -and ($LogEntryList[$b].Thread -eq $Thread)) {
						$ProactiveRemediationDetectExitCode = $Matches[1]
						
						# Set RunTime to Powershell script Install End log entry
						$LogEntryList[$b].ProcessRunTime = "End Proactive PostDetect Powershell (ExitCode=$ProactiveRemediationDetectExitCode)"

						$PowershellScriptEndTime = $LogEntryList[$b].LogEntryDateTime
						$PowershellScriptEndIndex = $b
						
						# Continue to previous log entry
						Continue
					}

					if(($LogEntryList[$b].Message -Match '^process id = (.*)$') -and ($LogEntryList[$b].Thread -eq $Thread)) {
						$ProactiveRemediationPostDetectProcessId = $Matches[1]
						
						# Set RunTime to Powershell script Install End log entry
						$LogEntryList[$b].ProcessRunTime = "Start Proactive PostDetect Powershell"

						$PowershellScriptStartTime = $LogEntryList[$b].LogEntryDateTime
						
						# Calculate Powershell script runtime
						if($PowershellScriptStartTime -and $PowershellScriptEndTime) {
							$PowershellScriptRuntime = New-Timespan $PowershellScriptStartTime $PowershellScriptEndTime
							$ProcessRunTimeTotalSeconds = $PowershellScriptRuntime.TotalSeconds
						} else {
							$ProcessRunTimeTotalSeconds = 'N/A'
						}
						
						$PowershellScriptStartIndex = $b

						# Set RunTime to Powershell script Start log entry
						$LogEntryList[$b].ProcessRunTime = "{0,-16} {1,8:n1}" -f "       Start Proactive PostDetect Powershell", $ProcessRunTimeTotalSeconds

						# Add RunTime to Powershell script End log entry
						$LogEntryList[$PowershellScriptEndIndex].ProcessRunTime = "{0,-16} {1,8:n1}" -f "       End Proactive PostDetect Powershell", $ProcessRunTimeTotalSeconds
						
						# Continue to previous log entry
						Continue
					}

					if(($LogEntryList[$b].Message -Match '^\[HS\] The policy needs be run as (.*)$') -and ($LogEntryList[$b].Thread -eq $Thread)) {
						$ProactiveRemediationPostDetectRunAs = $Matches[1]
						
						# Continue to previous log entry
						Continue
					}
					
					if(($LogEntryList[$b].Message -Match '^\[HS\] Processing policy with id = (.*)$') -and ($LogEntryList[$b].Thread -eq $Thread)) {
						$PolicyIdStart = $Matches[1]
					
						# $b is startIndex
						
						if($PolicyIdStart -eq $PolicyIdEnd) {
							# We found matching starting log entry
							$DetectionLogEntryStartTime = $LogEntryList[$b].LogEntryDateTime
							if($DetectionLogEntryStartTime -and $DetectionLogEntryEndTime) {
								$ProactiveDetectionRuntime = New-Timespan $DetectionLogEntryStartTime $DetectionLogEntryEndTime
								$ProcessRunTimeTotalSeconds = $ProactiveDetectionRuntime.TotalSeconds
							} else {
								$ProcessRunTimeTotalSeconds = 'N/A'
							}

							$Padding = ''

							# Start message to Timeline
							if($ShowAllTimelineEntries) {
								
								RecordStatusToTimeline -date $LogEntryList[$b].DateTime -status "Process / PostDetect" -detail "$ProactiveRemediationScriptPolicyTypeName $($Padding):  $ProactiveRemediationScriptName (run as $ProactiveRemediationPostDetectRunAs)" -logEntry "Line $($LogEntryList[$b].Line)"
							}
							
							# Set RunTime log entry
							$LogEntryList[$b].ProcessRunTime = "Processing $ProactiveRemediationScriptPolicyTypeName PostDetect script ($ProactiveRemediationScriptName) (run as $ProactiveRemediationPostDetectRunAs)"


							# End message to Timeline
							if($PostRemediationDetectResult -eq 'True') {
								
								RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Success / PostDetect" -detail "$ProactiveRemediationScriptPolicyTypeName $($Padding):  $ProactiveRemediationScriptName" -Seconds $ProcessRunTimeTotalSeconds -Color 'Green' -logEntry "Line $($LogEntryList[$i].Line)"
								
								# Set RunTime log entry
								$LogEntryList[$i].ProcessRunTime = "Succeeded $ProactiveRemediationScriptPolicyTypeName PostDetect script ($ProactiveRemediationScriptName) (Runtime: $ProcessRunTimeTotalSeconds sec)"
	
							} else {
								
								RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Failed  / PostDetect" -detail "$ProactiveRemediationScriptPolicyTypeName $($Padding):  $ProactiveRemediationScriptName" -Seconds $ProcessRunTimeTotalSeconds -Color 'Yellow' -logEntry "Line $($LogEntryList[$i].Line)"
								
								$LogEntryList[$i].ProcessRunTime = "Failed $ProactiveRemediationScriptPolicyTypeName PostDetect script ($ProactiveRemediationScriptName) (Runtime: $ProcessRunTimeTotalSeconds sec)"
							}

							# Script StdOut to Timeline if configured and if exists
							if($ShowStdOutInTimeline -and $ProactiveRemediationPostDetectStdOutIndex) {
																
								RecordStatusToTimeline -date $LogEntryList[$ProactiveRemediationPostDetectStdOutIndex].DateTime -status "StdOut" -detail "$ProactiveRemediationPostDetectStdOutput" -logEntry "Line $($LogEntryList[$ProactiveRemediationPostDetectStdOutIndex].Line)" -Color 'Yellow'
								
								$LogEntryList[$ProactiveRemediationPostDetectStdOutIndex].ProcessRunTime = "       Powershell PostDetect StdOut"
							}

							# Script Error to Timeline if configured and if exists
							if($ShowErrorsInTimeline -and $ProactiveRemediationPostDetectErrorOutputIndex) {
								
								RecordStatusToTimeline -date $LogEntryList[$ProactiveRemediationPostDetectErrorOutputIndex].DateTime -status "ErrorLog" -detail "$ProactiveRemediationPostDetectErrorOutput" -logEntry "Line $($LogEntryList[$ProactiveRemediationPostDetectStdOutIndex].Line)" -Color 'Yellow'
								
								$LogEntryList[$ProactiveRemediationPostDetectErrorOutputIndex].ProcessRunTime = "       Powershell PostDetect ErrorLog"
								
							}

					
						} else {
							# We missed Proactive Remediation Detect start log entry
							# We should never get here
							
							$Padding = ''
							if($PostRemediationDetectResult -eq 'True') {
								
								RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Success / PostDetect" -detail "$ProactiveRemediationScriptPolicyTypeName $($Padding):  $ProactiveRemediationScriptName" -Color 'Green' -logEntry "Line $($LogEntryList[$i].Line)"
								
								# Set RunTime log entry
								$LogEntryList[$i].ProcessRunTime = "Succeeded $ProactiveRemediationScriptPolicyTypeName PostDetect script ($ProactiveRemediationScriptName)"
	
							} else {
								RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Failed  / PostDetect" -detail "$ProactiveRemediationScriptPolicyTypeName $($Padding):  $ProactiveRemediationScriptName" -Color 'Yellow' -logEntry "Line $($LogEntryList[$i].Line)"
								
								$LogEntryList[$i].ProcessRunTime = "Failed/NotCompliant $ProactiveRemediationScriptPolicyTypeName PostDetect script ($ProactiveRemediationScriptName)"
							}
							
							
							# Script StdOut to Timeline if configured and if exists
							if($ShowStdOutInTimeline -and $ProactiveRemediationPostDetectStdOutIndex) {
																
								RecordStatusToTimeline -date $LogEntryList[$ProactiveRemediationPostDetectStdOutIndex].DateTime -status "StdOut" -detail "$ProactiveRemediationPostDetectStdOutput" -logEntry "Line $($LogEntryList[$ProactiveRemediationPostDetectStdOutIndex].Line)" -Color 'Yellow'
								
								$LogEntryList[$ProactiveRemediationPostDetectStdOutIndex].ProcessRunTime = "       Powershell PostDetect StdOut"
							}

							# Script Error to Timeline if configured and if exists
							if($ShowErrorsInTimeline -and $ProactiveRemediationPostDetectErrorOutputIndex) {
								
								RecordStatusToTimeline -date $LogEntryList[$ProactiveRemediationPostDetectErrorOutputIndex].DateTime -status "ErrorLog" -detail "$ProactiveRemediationPostDetectErrorOutput" -logEntry "Line $($LogEntryList[$ProactiveRemediationPostDetectStdOutIndex].Line)" -Color 'Yellow'
								
								$LogEntryList[$ProactiveRemediationPostDetectErrorOutputIndex].ProcessRunTime = "       Powershell PostDetect ErrorLog"
								
							}

							Break
						}
						
						# We are done
						
						# Break out from For-loop
						Break
					}

					$Padding = ''
					if($LogEntryList[$b - 1].Message -like '`[HS`] the pre-remdiation detection script compliance result for*') {
						# We missed Proactive Remediation Detect start log entry
						
						#Write-Host "Missed Proactive Remedation Detect script start log entry" -ForegroundColor Red
						
						if($PostRemediationDetectResult -eq 'True') {
							
							RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Success / PostDetect" -detail "$ProactiveRemediationScriptPolicyTypeName $($Padding):  $ProactiveRemediationScriptName" -Color 'Green' -logEntry "Line $($LogEntryList[$i].Line)"
							
							# Set RunTime log entry
							$LogEntryList[$i].ProcessRunTime = "Succeeded $ProactiveRemediationScriptPolicyTypeName PostDetect script ($ProactiveRemediationScriptName)"

						} else {
							RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Failed  / PostDetect" -detail "$ProactiveRemediationScriptPolicyTypeName $($Padding):  $ProactiveRemediationScriptName" -Color 'Yellow' -logEntry "Line $($LogEntryList[$i].Line)"
							
							$LogEntryList[$i].ProcessRunTime = "Failed/NotCompliant $ProactiveRemediationScriptPolicyTypeName PostDetect script ($ProactiveRemediationScriptName)"
						}
						
						# Script StdOut to Timeline if configured and if exists
						if($ShowStdOutInTimeline -and $ProactiveRemediationPostDetectStdOutIndex) {
															
							RecordStatusToTimeline -date $LogEntryList[$ProactiveRemediationPostDetectStdOutIndex].DateTime -status "StdOut" -detail "$ProactiveRemediationPostDetectStdOutput" -logEntry "Line $($LogEntryList[$ProactiveRemediationPostDetectStdOutIndex].Line)" -Color 'Yellow'
							
							$LogEntryList[$ProactiveRemediationPostDetectStdOutIndex].ProcessRunTime = "       Powershell PostDetect StdOut"
						}

						# Script Error to Timeline if configured and if exists
						if($ShowErrorsInTimeline -and $ProactiveRemediationPostDetectErrorOutputIndex) {
							
							RecordStatusToTimeline -date $LogEntryList[$ProactiveRemediationPostDetectErrorOutputIndex].DateTime -status "ErrorLog" -detail "$ProactiveRemediationPostDetectErrorOutput" -logEntry "Line $($LogEntryList[$ProactiveRemediationPostDetectStdOutIndex].Line)" -Color 'Yellow'
							
							$LogEntryList[$ProactiveRemediationPostDetectErrorOutputIndex].ProcessRunTime = "       Powershell PostDetect ErrorLog"
							
						}
					
						Break
					}
					
				}
			} # For-loop end
			
		}
		# endregion Try to find Proactive PostDetection script run
		


		###############################################
		# Applications
		# Win32App
		# WinGetApp


		# Check for application policies
		if($LogEntryList[$i].Message -like 'Get policies = *') {
			if($LogEntryList[$i].Message -Match '^Get policies = \[(.*)\]$') {
				$IntuneAppPolicies = "[$($Matches[1])]" | ConvertFrom-Json
				
				# Add policies to hash. Remove existing hash object if exist already
				Foreach($AppPolicy in $IntuneAppPolicies) {
					
					# Check if App exists in Hashtable
					if($IdHashtable.ContainsKey($AppPolicy.Id)) {
						
						# Remove existing object
						$IdHashtable.Remove($AppPolicy.Id)
					}
					
					# Add object to Hashtable
					$IdHashtable.Add($AppPolicy.Id , $AppPolicy)
				}
			}
		}

		# Get App Name from AppDownload log entry
		if($LogEntryList[$i].Message -Like '`[StatusService`] Downloading app (id = (.*), name (.*)\).*$') {
			if($LogEntryList[$i].Message -Match '^\[StatusService\] Downloading app \(id = (.*), name (.*)\).*') {
				
				# Message matches Win32App Name message
				$Win32AppId = $Matches[1]
				$Win32AppName = $Matches[2]
				
				Write-Verbose "Found Win32App Name $Win32AppName ($Win32AppId)"

				# Add AppId and name to Hashtable
				if($IdHashtable.ContainsKey($Win32AppId)) {
					#$IdHashtable[$Win32AppId].id = $Win32AppId
					
					if($IdHashtable[$Win32AppId].name -ne $Win32AppName) {
						Write-Host "Win32App name mismatch found. We should never get this warning" -ForegroundColor Yellow
						Write-Host "($($IdHashtable[$Win32AppId].name)) ($Win32AppName) ($Win32AppId)" -ForegroundColor Yellow
					}

					#$IdHashtable[$Win32AppId].displayName = $Win32AppName

				} else {
					$AppIdCustomObjectProperties = @{
						id = $Win32AppId
						name = $Win32AppName
						displayName = $Win32AppName
						intent = $null
					}
					$AppIdCustomObject = New-Object -TypeName PSObject -Prop $AppIdCustomObjectProperties

					# Create new UserId hashtable entry
					$IdHashtable.Add($Win32AppId, $AppIdCustomObject)
				}
			}
		}

		# Check for Applications to be installed in ESP phase (Device ESP and User ESP has it's own log entries)
		if($LogEntryList[$i].Message -like '`[Win32App`]`[ESPAppLockInProcessor`] Found * apps which need to be installed for current phase of ESP. AppIds: *') {
			if($LogEntryList[$i].Message -Match '^\[Win32App\]\[ESPAppLockInProcessor\] Found (.*) apps which need to be installed for current phase of ESP. AppIds: (.*)$') {
				$AppCountToInstallInESPPhase = $Matches[1]				
				$AppIdsToInstallInESPPhase = $Matches[2]

				if($AppIdsToInstallInESPPhase) {

					RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Info    / Win32App" -detail "Apps in ESP              $AppCountToInstallInESPPhase Apps to install at ESP phase" -logEntry "Line $($LogEntryList[$i].Line)"


					# Several apps are separated by ,
					$AppIdsToInstallInESPPhaseArray = $AppIdsToInstallInESPPhase.Split(',').Trim()

					Foreach($AppId in $AppIdsToInstallInESPPhaseArray) {
						# Get name for App
						if($IdHashtable.ContainsKey($AppId)) {
							$AppName = $IdHashtable[$AppId].name
							RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Info    / Win32App" -detail "Apps in ESP                 $AppName" -logEntry "Line $($LogEntryList[$i].Line)"
						}

					}

				}

				#$LogEntryList[$i].ProcessRunTime = "User ESP finished successfully"

				# Continue to next line in Do-while -loop
				$i++
				Continue

			}
		}

		# Case where different version of App of detected after Uninstall than expected
		# Usually in this case either Uninstall failed or there is wrong detection check in older app (eg. version -gt used)
		if($LogEntryList[$i].Message -like '`[Win32App`]`[ActionProcessor`] Encountered unexpected state for app with id:*') {
			# Set RunTime message
			$LogEntryList[$i].ProcessRunTime = "Warning: Encountered unexpected state for app with id"
		}

		
		# region Find Win32App installations
		# Check if this is ending message for Win32 application installation
		if(($LogEntryList[$i].Message -like '`[Win32App`] Installation is done, collecting result') -or ($LogEntryList[$i].Message -like '`[Win32App`] Failed to create installer process. Error code = *')) {
			$FailedToCreateInstallerProcess = $False


			# We will follow this Thread
			$Thread = $LogEntryList[$i].Thread

			if($LogEntryList[$i].Message -Match '^\[Win32App\] Failed to create installer process. Error code = (.*)$') {
				$ErrorMessage = $Matches[0]
				$ErrorCode = $Matches[1]
				$FailedToCreateInstallerProcess = $True

				$Win32AppInstallStartFailedIndex = $i

				# Set RunTime to Win32App Install Start Failed log entry
				$LogEntryList[$Win32AppInstallStartFailedIndex].ProcessRunTime = "$ErrorMessage"

				#Write-Host "DEBUG: Found Win32 App error: $ErrorMessage"
				Write-Verbose "Found Win32 App error: $ErrorMessage"

				# Set index so our search continues in next steps
				# Because we skip Loop below
				$b = $i - 1

			} else {

				$Win32AppInstallStartIndex = $null
				$Win32AppInstallEndIndex = $i
				
				$ProcessStartTime = $null
				$ProcessEndTime = $LogEntryList[$i].LogEntryDateTime

				Write-Verbose "Found Win32 App install end process message (Line $($LogEntryList[$i].Line) Thread $Thread): $($LogEntryList[$i].Message)"

				$Win32AppExitCode='N/A'
				$Win32AppInstallSuccess='N/A'

				# We don't know App Name or Intent yet
				$LogEntryList[$Win32AppInstallEndIndex].ProcessRunTime = "End Win32App"

				# Check if we can find exitCode and Success/failed information in next 20 lines
				for($b = $Win32AppInstallEndIndex + 1; $b -lt $Win32AppInstallEndIndex + 20; $b++ ) {
					$Matches=$null
					
					# Run this if first because it is much faster than run -Match for every line
					if($Win32AppExitCode -eq 'N/A') {
						if(($LogEntryList[$b].Message -Match '^\[Win32App\] lpExitCode (.*)$') -and ($LogEntryList[$b].Thread -eq $Thread)) {
							$Win32AppExitCode = $Matches[1]
							Write-Verbose "Found Win32App install exitCode: $Win32AppExitCode"

							# We don't know App Name or Intent yet
							$LogEntryList[$i].ProcessRunTime = "(ExitCode $Win32AppExitCode"

							# Continue For-loop to next log entry
							Continue
						}
					}
					
					# Run this if-clause first because it is much faster than run -Match for every line
					if($Win32AppInstallSuccess -eq 'N/A') {
						if(($LogEntryList[$b].Message -Match '^\[Win32App\] lpExitCode is defined as (.*)$') -and ($LogEntryList[$b].Thread -eq $Thread)) {
							$Win32AppInstallSuccess = $Matches[1]
							Write-Verbose "Found Win32App install success: $Win32AppInstallSuccess"

							# We don't know App Name or Intent yet
							$LogEntryList[$i].ProcessRunTime = "$($LogEntryList[$i].ProcessRunTime) $Win32AppInstallSuccess)"

							# Break out from For-loop
							Break
						}
					}
				} # For-loop end
				
				# Check if we can find previous LogEntry that matches same process thread start message
				$b = $Win32AppInstallEndIndex - 1
				While ($b -ge 0) {
					if(($LogEntryList[$b].Message -like '`[Win32App`] process id =*') -and ($LogEntryList[$b].Thread -eq $Thread)) {
						$Win32AppInstallStartIndex = $b

						# Message matches Win32App install start message
						$ProcessStartTime = $LogEntryList[$b].LogEntryDateTime
						$ProcessRunTime = New-Timespan $ProcessStartTime $ProcessEndTime
						$ProcessRunTimeTotalSeconds = $ProcessRunTime.TotalSeconds
						#$ProcessRunTime = "{0,-16} {1,8:n1}" -f "End Win32App", $ProcessRunTimeTotalSeconds
						
						Write-Verbose "Found Win32 App install Start process message (Line $($LogEntryList[$b].Line)): $($LogEntryList[$b].Message)"

						# FIXME
						# Set Win32App Name to Win32App Install Start entry
						#$LogEntryList[$Win32AppInstallStartIndex].ProcessRunTime = "$($LogEntryList[$Win32AppInstallStartIndex].ProcessRunTime)   $Win32AppName"
						
						# Set RunTime to Win32App Install Start log entry
						$LogEntryList[$Win32AppInstallStartIndex].ProcessRunTime = "{0,-16} {1,8:n1}" -f "Start Win32App", $ProcessRunTimeTotalSeconds

						# Set RunTime to Powershell script Install End log entry
						#$LogEntryList[$i].ProcessRunTime = "{0,-16} {1,8:n1}" -f "End Win32App", $ProcessRunTimeTotalSeconds
						$LogEntryList[$Win32AppInstallEndIndex].ProcessRunTime = "$($LogEntryList[$i].ProcessRunTime) $($ProcessRunTimeTotalSeconds)sec"
						
						Write-Verbose "Runtime: $ProcessRunTimeTotalSeconds"
						
						Break
					} else {
						# Message does not match Win32App install start message
					}							
					$b--
				}
			}

			# Try to find Win32App name if we have start and end time
			$Win32AppId = $null
			$Win32AppName = $null
			$Intent = 'Unknown Intent'
			$TempAppId = $null
			$BackupIntent = $null
			$intentDoubleCheck = $null
			$AppIntentNumberClosestToProcess = $null
			$AppSupersededUninstall = $null
			$AppDeliveryOptimization = $null

			if($ProcessStartTime -or $ProcessEndTime -or $FailedToCreateInstallerProcess) {
				# Remove this line
				#$Win32AppInstallStartIndex = $b
				
				# We are going upwards now
				While ($b -ge 0) {
					#Write-Host "$b $($LogEntryList[$b].Message)"

					# Find Win32App Id
					if(-not $Win32AppId) {
						if(($LogEntryList[$b].Message -Match '^\[Win32App\] SetCurrentDirectory: C:\\Windows\\IMECache\\(.*)_.*$') -and ($LogEntryList[$b].Thread -eq $Thread)) {
							$Win32AppId = $Matches[1]

							#Write-Host "DEBUG: Found Win32AppId: $Win32AppId"

							# Find name from Hashtable
							if($IdHashtable.ContainsKey($Win32AppId)) {
								if($IdHashtable[$Win32AppId].name) {
									$Win32AppName = $IdHashtable[$Win32AppId].name
									#Write-Host "DEBUG: Found existing Win32AppName: $Win32AppName`n"
								}
							}

							$b--
							Continue
						}
					}

					# Double check intent
					# If old application is Superceded then in policy intent is not targeted
					# But if old app version is detected then uninstall is triggered and intent
					# Need to be checked from this log entry
					if(-not $AppIntentNumberClosestToProcess) {
						if(($LogEntryList[$b].Message -Match '^\[Win32App\] ===Step=== InstallBehavior RegularWin32App, Intent (.*), UninstallCommandLine .*$') -and ($LogEntryList[$b].Thread -eq $Thread)) {
							$AppIntentNumberClosestToProcess = $Matches[1]

							#Write-Host "DEBUG: Found InstallBehavior: $AppIntentNumberClosestToProcess"
							
							$b--
							Continue
						}
					}

					
					# Get Delivery Optimization information
					if(-not $AppDeliveryOptimization -and $Win32AppId) {
						# ? is any character because escape `[ and `] didn't seem to work with doublequotes "   ???
						if(($LogEntryList[$b].Message -Like "?DO TEL? = {*$($Win32AppId)*}") -and ($LogEntryList[$b].Thread -eq $Thread)) {
							if(($LogEntryList[$b].Message -Match '^\[DO TEL\] = \{(.*)\}$') -and ($LogEntryList[$b].Thread -eq $Thread)) {
								$AppDeliveryOptimization = "{$($Matches[1])}" | ConvertFrom-Json
								
								# Disabled now because we don't have Win32App Processsing Start and End entries
								#$LogEntryList[$b].ProcessRunTime = "   Download and Delivery Optimization summary"

								#Write-Host "DEBUG: Found Delivery Optimization information:`n$($AppDeliveryOptimization | ConvertTo-Json)"
								
								$b--
								Continue
							}
						}
					}
					

<#					
					# Get intent: RequiredInstall, RequiredUninstall, AvailableInstall, AvailableUninstall?
					# We will catch AvailableInstall here which are apps installed from Intune Company Portal
					#
					# We will also catch NotTargeted which is superceded application
					# If we find NotTargeted then we should try to find actual app install after previous version uninstall
					if(($LogEntryList[$b].Message -Match '^\[Win32App\]\[ActionProcessor\] App with id: (.*), targeted intent: (.*),.*$') -and ($LogEntryList[$b].Thread -eq $Thread)) {

						$TempAppId = $Matches[1]
						if($TempAppId -eq $Win32AppId) {
							$BackupIntent = $Matches[2]
							#Write-Host "Found backupIntent: $Win32AppName $BackupIntent ($TempAppId) ($Win32AppId)"
						} else {
							#Write-Host "Found backupIntent for wrong id: $Win32AppName $BackupIntent ($TempAppId) ($Win32AppId)"
						}

						
						# End loop
						Break
					}
#>
					

					# Check if we reached start of this round
					if($LogEntryList[$b].Message -Like '^`[Win32App`]`[V3Processor`] Processing * subgraphs.$') {

						#Write-Host "We reached start Win32App processing phase"
						#Write-Verbose "We reached start Win32App processing phase"
						#Write-Host "$($LogEntryList[$b].Message)"
						
						# End loop
						Break
					}

					# Check if we reached start of this round
					# In some rare cases this does not catch Win32App installs in ESP phase
					# Look for ExecMaanger below
					if($LogEntryList[$b].Message -Match '^\[Win32App\]\[V3Processor\] Processing subgraph with app ids: (.*)$') {
						$TempAppId = $Matches[1]

						#Write-Host "We reached start Win32App processing phase: $TempAppId"
						#Write-Verbose "We reached start Win32App processing phase: $TempAppId"
						#Write-Host "$($LogEntryList[$b].Message)"
						
						# End loop
						Break
					}

					# This is start entry for Win32App processing
					# Only shown in ESP phase for rare cases or not shown at all?
					if($LogEntryList[$b].Message -Like '`[Win32App`] ExecManager: processing targeted app*') {

						#Write-Host "We reached start Win32App processing phase"
						#Write-Verbose "We reached start Win32App processing phase"
						#Write-Host "$($LogEntryList[$b].Message)"
						
						# End loop
						Break
					}
					
					# This is start entry for Win32App processing during ESP insome cases
					# We should never get here
					if($LogEntryList[$b].Message -Like '`[Win32App`]ExeManager: start processing app policies with count = *') {

						#Write-Host "We reached start Win32App processing phase"
						#Write-Verbose "We reached start Win32App processing phase"
						#Write-Host "$($LogEntryList[$b].Message)"
						
						# End loop
						Break
					}



					# This is one possible end of subgraph
					if($LogEntryList[$b].Message -Match '^\[Win32App\]\[V3Processor\] Done processing subgraph.$') {

						#Write-Host "We reached start Win32App processing phase: $TempAppId"
						#Write-Verbose "We reached start Win32App processing phase: $TempAppId"
						#Write-Host "$($LogEntryList[$b].Message)"
						
						# End loop
						Break
					}

					# This is one possible end of subgraph
					if($LogEntryList[$b].Message -Match '^\[Win32App\]\[V3Processor\] All apps in the subgraph are not applicable due to assignment filters. Processing will be skipped.$') {

						#Write-Host "We reached start Win32App processing phase: $TempAppId"
						#Write-Verbose "We reached start Win32App processing phase: $TempAppId"
						#Write-Host "$($LogEntryList[$b].Message)"
						
						# End loop
						Break
					}

					$b--
				}

<#
				# DEBUG Intents
				if($FailedToCreateInstallerProcess) {
					Write-Host "DEBUG: $Win32AppName (Line: $($LogEntryList[$Win32AppInstallStartFailedIndex].Line)) AppIntentNameClosestToProcess    : $(Get-AppIntentNameForNumber $AppIntentNumberClosestToProcess)"
					Write-Host "DEBUG: $Win32AppName (Line: $($LogEntryList[$Win32AppInstallStartFailedIndex].Line)) Last known Intent from App policy: $(Get-AppIntentNameForNumber $IdHashtable[$Win32AppId].Intent)`n"
				} else {
					Write-Host "DEBUG: $Win32AppName (Line: $($LogEntryList[$Win32AppInstallEndIndex].Line)) AppIntentNameClosestToProcess    : $(Get-AppIntentNameForNumber $AppIntentNumberClosestToProcess)"
					Write-Host "DEBUG: $Win32AppName (Line: $($LogEntryList[$Win32AppInstallEndIndex].Line)) Last known Intent from App policy: $(Get-AppIntentNameForNumber $IdHashtable[$Win32AppId].Intent)`n"
				}
#>

				# Set Intent placeholder value which is closest Intent found from process start (going upwards)
				if($AppIntentNumberClosestToProcess) {
					$intent = Get-AppIntentNameForNumber $AppIntentNumberClosestToProcess
				} else {
					# Set fallback Inten value from App Policy
					$intent = Get-AppIntent $Win32AppId
				}

				# Get App Intent Name from policy
				$AppIntentNameFromPolicy = Get-AppIntent $Win32AppId

				# Get App Intent Name from closest Intent found from process start (going upwards)
				if($AppIntentNumberClosestToProcess) {
					$AppIntentNameClosestToProcess = Get-AppIntentNameForNumber $AppIntentNumberClosestToProcess
				}

				# In case of Available Install we have different values
				# but we will use value from policy if action is RequiredInstall
				if(($AppIntentNameFromPolicy -eq 'Available Install') -and ($AppIntentNameClosestToProcess -eq 'Required Install')) {
					$intent = 'Available Install'
				}

				# In case of supercedence in policy we usually have NotTargeted but action is RequiredUninstall
				# but can have previous value of Available Install if we have installed old App version just previously (maybe not normal case)
				#if(($AppIntentNameFromPolicy -eq 'Not Targeted') -and ($AppIntentNameClosestToProcess -eq 'Required Uninstall')) {
				if(($AppIntentNameFromPolicy -ne 'Required Uninstall') -and ($AppIntentNameClosestToProcess -eq 'Required Uninstall')) {
					$intent = 'Required Uninstall'
					$AppSupersededUninstall = $True
				}


				if($AppDeliveryOptimization) {
					#$DownloadDuration = $AppDeliveryOptimization.DownloadDuration
					$DownloadDurationDateTimeObject = Get-Date $AppDeliveryOptimization.DownloadDuration
					
					# Full seconds
					$DownloadDuration = New-Timespan -Hours $DownloadDurationDateTimeObject.Hour -Minutes $DownloadDurationDateTimeObject.Minute -Seconds $DownloadDurationDateTimeObject.Second
					$DownloadDurationSeconds = $DownloadDuration.TotalSeconds
					
					$BytesFromPeers = $AppDeliveryOptimization.BytesFromPeers
					$BytesFromPeersMBRounded = [math]::Round($BytesFromPeers /1MB, 0)
					
					$TotalBytesDownloaded = $AppDeliveryOptimization.TotalBytesDownloaded
					$TotalBytesDownloadedMBExact = $TotalBytesDownloaded /1MB
					$TotalBytesDownloadedMBRounded = [math]::Round($TotalBytesDownloaded /1MB, 0)
					
					$DownloadMBps = $TotalBytesDownloadedMBExact / $DownloadDurationSeconds
					$DownloadMBps = [math]::Round($DownloadMBps, 1)
					if($DownloadMBps -eq 0) {
						$DownloadMBps = 'N/A'
					}


					if($BytesFromPeers -gt 0) {
						$BytesFromPeersPercentage = ($BytesFromPeers / $TotalBytesDownloaded).ToString('#%')
					} else {
						$BytesFromPeersPercentage = '0%'
					}

					# Add download statistics object to list
					$ApplicationDownloadStatistics.add([PSCustomObject]@{
						'AppType' = 'Win32App'
						'AppName' = $Win32AppName
						'DL Sec' = $DownloadDurationSeconds
						'Size (MB)' = $TotalBytesDownloadedMBRounded
						'MB/s' = $DownloadMBps
						'Delivery Optimization %' = $BytesFromPeersPercentage
						})
					$ApplicationDownloadStatistics.Add($DownloadStats)

					#$DownloadStats = "DL:$($DownloadDurationSeconds)s,$($TotalBytesDownloadedMBRounded)MB,$($DownloadMBps)MB/s,DO:$($BytesFromPeersPercentage)"
					#Write-Host "$($Win32AppName): $DownloadStats"
				}
				
				if($FailedToCreateInstallerProcess) {
					# Failed to create process

					if($AppSupersededUninstall) {
						$detail = "{0,$ColonIndentInTimeLine}{1}  {2}" -f "$Intent", ':', "$Win32AppName (Superseded) $(($LogEntryList[$Win32AppInstallStartFailedIndex].Message).Replace('[Win32App]','').Trim())"

					} else {
						$detail = "{0,$ColonIndentInTimeLine}{1}  {2}" -f "$Intent", ':', "$Win32AppName $(($LogEntryList[$Win32AppInstallStartFailedIndex].Message).Replace('[Win32App]','').Trim())"
					}

					# Set RunTime to Win32App Install Start Failed log entry
					$LogEntryList[$Win32AppInstallStartFailedIndex].ProcessRunTime = "Failed $Win32AppName $intent ($ErrorMessage)"

					RecordStatusToTimeline -date $LogEntryList[$Win32AppInstallStartFailedIndex].DateTime -status "Failed  / Win32App" -detail $detail -logEntry "Line $($LogEntryList[$Win32AppInstallStartFailedIndex].Line)" -Seconds '' -Color 'Red'

					# Set Info message about possible missing uninstall executable
					# First check that uninstall command is correct.
					#
					# Another case could be that there was earlier uninstall process started while application was running
					# Uninstaller then could not delete main executable because it was locked but it deleted uninstall.exe
					# And if Detection Check is for main executable then partially uninstalled App is detected for old version
					# So Intune always detects old App version and is trying to uninstall it before installing new App version but uninstaller is missing
					$detail = "{0,$ColonIndentInTimeLine}{1}  {2}" -f "", ':', "   Incorrect uninstall command line or uninstall .exe missing?"

					RecordStatusToTimeline -date $LogEntryList[$Win32AppInstallStartFailedIndex].DateTime -status "Info    / Win32App" -detail $detail -logEntry "Line $($LogEntryList[$Win32AppInstallStartFailedIndex].Line)" -Seconds '' -Color 'Yellow'

					if($IdHashtable.ContainsKey($Win32AppId)) {
						if($IdHashtable[$Win32AppId].UninstallCommandLine) {
							$UninstallCommandLine = $IdHashtable[$Win32AppId].UninstallCommandLine
							$detail = "{0,$ColonIndentInTimeLine}{1}  {2}" -f "", ':', "   Uninstall command: $UninstallCommandLine"

							RecordStatusToTimeline -date $LogEntryList[$Win32AppInstallStartFailedIndex].DateTime -status "Info    / Win32App" -detail $detail -logEntry "Line $($LogEntryList[$Win32AppInstallStartFailedIndex].Line)" -Seconds '' -Color 'Yellow'

						}
					}

					#Write-Host "DEBUG: Write Fail info to Timeline. `$detail=$detail"

				} else {
					# Set RunTime to Win32App Install Start log entry
					$LogEntryList[$Win32AppInstallStartIndex].ProcessRunTime = "Start Win32App $Win32AppName $Intent $($ProcessRunTimeTotalSeconds)sec"

					# Set RunTime to Powershell script Install End log entry
					#$LogEntryList[$i].ProcessRunTime = "{0,-16} {1,8:n1}" -f "End Win32App", $ProcessRunTimeTotalSeconds
					$LogEntryList[$Win32AppInstallEndIndex].ProcessRunTime = "End Win32App $Win32AppName $Intent $($LogEntryList[$Win32AppInstallEndIndex].ProcessRunTime)"


					if($AppSupersededUninstall) {
						$detail = "{0,$ColonIndentInTimeLine}{1}  {2}" -f "$Intent", ':', "$Win32AppName ($Win32AppExitCode $Win32AppInstallSuccess) (Superseded)"
					} else {
						$detail = "{0,$ColonIndentInTimeLine}{1}  {2}" -f "$Intent", ':', "$Win32AppName ($Win32AppExitCode $Win32AppInstallSuccess)"
					}
					
					# Change text color depending on Win32App install Success status
					if($Win32AppInstallSuccess -eq 'Success') {
						
						RecordStatusToTimeline -date $LogEntryList[$Win32AppInstallEndIndex].DateTime -status "Success / Win32App" -detail $detail -logEntry "Line $($LogEntryList[$Win32AppInstallEndIndex].Line)" -Seconds $ProcessRunTimeTotalSeconds -Color 'Green'

					} else {
						
						RecordStatusToTimeline -date $LogEntryList[$Win32AppInstallEndIndex].DateTime -status "Failed  / Win32App" -detail $detail -logEntry "Line $($LogEntryList[$Win32AppInstallEndIndex].Line)" -Seconds $ProcessRunTimeTotalSeconds -Color 'Red'
					}
				}
			}

			
			if(($b -eq 0) -and (-not $ProcessStartTime)) {
				# We did not find Win32App install start time
				
				Write-Host "Did NOT find matching Win32App install start time for Win32App install in line $LineNumber (Thread=$Thead)`n" -ForegroundColor Yellow
			}
			
			# Cleanup variables
			$ProcessStartTime = $null
			$ProcessEndTime = $null
			$Thread = $null
			$ProcessRunTimeTotalSeconds = $null
			$ProcessRunTimeText = $null
		}
		# endregion Find Win32App installations


		# region Find WinGet Application install
		if($LogEntryList[$i].Message -like '`[Win32App`]`[V3Processor`] Processing subgraph with app ids:*') {
			if($LogEntryList[$i].Message -Match '^\[Win32App\]\[V3Processor\] Processing subgraph with app ids: (.*)$') {
				$WinGetAppId = $Matches[1]
				$WinGetProcessingStartIndex = $i

				$Thread = $LogEntryList[$i].Thread
				$WinGetInstallationStartedDateTime = $LogEntryList[$i].DateTime

				$WinGetDownloadFirstEntryFound=$false
				$BackupIntent=$null

				$WinGetAppName = Get-AppName $WinGetAppId

				#Write-Host "DEBUG: Found WinGet installation start. `$i=$i"

				# Find next WinGet installation log entries
				# Now we go down (to end)
				for($b = $i + 1; $b -lt $LogEntryList.Count; $b++ ) {

					# Use this to debug WinGet log entries
					#Write-Host "DEBUG: $b - $($LogEntryList[$b].Message)"

					if(($LogEntryList[$b].Message -Match '^\[Win32App\]\[WinGetApp\]\[WinGetAppExecutionExecutor\] Starting execution of app with id: (.*) and context: (.*)\.$') -and ($LogEntryList[$b].Thread -eq $Thread)) {
						$TempAppId = $Matches[1]
						$WinGetInstallContext = $Matches[2]
						
						#Write-Host "DEBUG: Found WinGet id ($TempAppId) and context ($WinGetInstallContext)"
						
						if($TempAppId -ne $WinGetAppId) {
							# We got different AppIds
							# Something is big time wrong here
							
							#Write-Host "DEBUG: WinGet Application Id mismatch in log entries. Aborting WinGet AppId: $WinGetAppId"

							# Break For-loop
							Break
						}
						
						# Continue For-loop to next round
						Continue
					}


					if(-not $WinGetDownloadFirstEntryFound) {
						if(($LogEntryList[$b].Message -Match '^\[StatusService\] Downloading app .id = (.*), name (.*). via WinGet, bytes (.*)\/(.*) for user (.*)$')) {
							$WinGetDownloadFirstEntryFound=$True
							
							$WinGetDownloadStartDateTime = $LogEntryList[$b].DateTime
							
							$AppId = $Matches[1]
							$WinGetAppName = $Matches[2]
							$WinGetAppBytes = $Matches[4]
							$WinGetUserId = $Matches[5]

							# Add WinGetAppId and name to Hashtable
							if($IdHashtable.ContainsKey($AppId)) {
								$IdHashtable[$AppId].id = $AppId
								$IdHashtable[$AppId].name = $WinGetAppName
								#$IdHashtable[$AppId].displayName = $WinGetAppName
							} else {
								$AppIdCustomObjectProperties = @{
									id = $AppId
									name = $WinGetAppName
									displayName = $WinGetAppName
								}
								$AppIdCustomObject = New-Object -TypeName PSObject -Prop $AppIdCustomObjectProperties

								# Create new UserId hashtable entry
								$IdHashtable.Add($AppId , $AppIdCustomObject)
							}


							# Get real UserName if found from Hashtable
							if($IdHashtable.ContainsKey($WinGetUserId)) {
								$Name = $IdHashtable[$WinGetUserId].name
								if($Name) {
									$WinGetUserName = $IdHashtable[$WinGetUserId].name
								} else {
									$WinGetUserName = $WinGetUserId
								}
							}

							$LogEntryList[$b].ProcessRunTime = "       Start Download ($WinGetAppName) for user $WinGetUserName"

							#Write-Host "DEBUG: Found WinGet first download entry"
							
							# Continue For-loop to next round
							Continue
						}
					}

					$DownloadCompleteString1 = 'State transition processed - Previous State:Download In Progress received event: Download Finished which took state machine into Download Complete state.'
					$DownloadCompleteString2 = 'Processing state transition - Current State:Download Complete With Event: Continue Install.'
					if(($LogEntryList[$b].Message -eq $DownloadCompleteString1) -or ($LogEntryList[$b].Message -eq $DownloadCompleteString2)) {

						$WinGetDownloadEnddDateTime = $LogEntryList[$b].DateTime						

						$LogEntryList[$b].ProcessRunTime = "       Download complete ($WinGetAppName)"

						#Write-Host "DEBUG: Found WinGet download complete"
						
						# Continue For-loop to next round
						Continue
					}

					
					$WinGetInstallStartString = 'State transition processed - Previous State:Download Complete received event: Continue Install which took state machine into Install In Progress Download Complete state.'
					if(($LogEntryList[$b].Message -eq $WinGetInstallStartString)) {
						#  -and ($LogEntryList[$b].Thread -eq $Thread)
						$WinGetInstallStartdDateTime = $LogEntryList[$b].DateTime						
						$LogEntryList[$b].ProcessRunTime = "       Start Install process ($WinGetAppName)"
						#Write-Host "WinGet Found install start"
					
					}


					# Get intent: RequiredInstall, RequiredUninstall, AvailableInstall, AvailableUninstall?	
					if(($LogEntryList[$b].Message -Match '^\[Win32App\]\[ActionProcessor\] App with id: (.*), targeted intent: (.*),.*$') -and ($LogEntryList[$b].Thread -eq $Thread)) {

						$TempAppId = $Matches[1]
						$BackupIntent = $Matches[2]

					}


					if(($LogEntryList[$b].Message -Match '^\[WinGet\] Install Operation Result has arrived (.*)$')) {
						$WinGetInstallResult = $Matches[1]

						# No updates available
						if($WinGetInstallResult -eq 'NoApplicableUpgrade') {

							#$WinGetAppName = Get-AppName $WinGetAppId

							$LogEntryList[$i].ProcessRunTime = "Start WinGet processing App ($WinGetAppName)"
							$LogEntryList[$b].ProcessRunTime = "NoApplicableUpgrade for WinGetApp ($WinGetAppName)"

							Break
						}
			
						
			
						# Calculate install time
						$ProcessStartTime = $LogEntryList[$i].LogEntryDateTime
						$ProcessEndTime = $LogEntryList[$b].LogEntryDateTime
						$ProcessRunTime = New-Timespan $ProcessStartTime $ProcessEndTime
						$ProcessRunTimeTotalSeconds = $ProcessRunTime.TotalSeconds
						
						$WinGetInstallEndDateTime = $LogEntryList[$b].DateTime
						$LogEntryList[$i].ProcessRunTime = "Start WinGet processing App ($WinGetAppName)"

						$intent = Get-AppIntent $AppId
						if($intent -eq 'Unknown Intent') {
							if($BackupIntent) {
								
								# This case when App is installed from Intune Company Portal
								# Then that App is not shown in Get Policies -list
								# Intent in AvailableInstall then
								$intent = $BackupIntent
							}
						}

						$detail = "{0,$ColonIndentInTimeLine}{1}  {2}" -f "$Intent", ':', "$WinGetAppName for user $WinGetUserName"
											
						if($WinGetInstallResult -eq 'OK') {
							
							RecordStatusToTimeline -date $LogEntryList[$b].DateTime -status "Success / WinGet" -detail $detail -Seconds $ProcessRunTimeTotalSeconds -Color 'Green' -logEntry "Line $($LogEntryList[$b].Line)"

							$LogEntryList[$b].ProcessRunTime = "Stop WinGet Install Success ($WinGetAppName) for user $WinGetUserName ($ProcessRunTimeTotalSeconds)"
							
						} else {
							
							RecordStatusToTimeline -date $LogEntryList[$b].DateTime -status "Failed  / WinGet" -detail $detail -Seconds $ProcessRunTimeTotalSeconds -Color 'Red' -logEntry "Line $($LogEntryList[$b].Line)"
							
							$LogEntryList[$b].ProcessRunTime = "Stop WinGet Install Failed ($WinGetAppName) for user $WinGetUserName ($ProcessRunTimeTotalSeconds)"
						}
						
						#Write-Host "WinGet Found install OK"
						
						# Break For-loop
						Break
					
					}

					if(((($LogEntryList[$b].Message -eq '[Win32App][V3Processor] Done processing subgraph.') -or `
						 ($LogEntryList[$b].Message -like '`[Win32App`]`[V3Processor`] Done processing * subgraphs.') -or `
						 ($LogEntryList[$b].Message -eq '[Win32App][V3Processor] All apps in the subgraph are not applicable due to assignment filters. Processing will be skipped.')))`
						 -and ($LogEntryList[$b].Thread -eq $Thread)) {

						#Write-Host "DEBUG: Found WinGet Done Processing subgraph log entry"
						#Write-Verbose "Subgraph $WinGetAppId did not include WinGet Application install ($($LogEntryList[$i].FileName) Line $($LogEntryList[$i].Line))"
						
						# Break For-loop
						Break
					}
					
					
					# Check that we didn't go over end log of this installation process
					# This is just fail safe check. We should not get here
					if($LogEntryList[$b].Message -like '`[Win32App`]`[V3Processor`] Processing subgraph with app ids:*') {
						# New WinGet installation started.
						# We didn't catch all log entries to process WinGet installation
						Write-Verbose "Could not detect WinGet installation for WinGet AppId: $WinGetAppId (Line $($LogEntryList[$i].Line))"
						
						# Break For-loop
						Break
					}

					
				}
			}
			
		} # Region ends here
		# region Find WinGet Application install
		
		$i++
	}
	# endregion analyze log entries


	# Application download statistics
    Write-Host ""
    Write-Host "Application download statistics" -ForegroundColor Magenta
	$ApplicationDownloadStatistics | Format-Table


	# This is aligned with Michael Niehaus's Get-AutopilotDiagnostics script just in case
    Write-Host ""
    Write-Host "OBSERVED TIMELINE" -ForegroundColor Magenta
	$observedTimeline |
        Format-Table @{
			label= 'Date'
			Expression = { $_.Date.Split('.')[0] }
        },
        @{
            Label = "Status"
            Expression =
            {
                Switch ($_.Color)
                {
                    'Red'    { $color = "91"; break }
                    'Yellow' { $color = '93'; break }
                    'Green'  { $color = "92"; break }
                    default { $color = "0" }
                }
                $e = [char]27
                "$e[${color}m$($_.Status)$e[0m"
            }
        },
		@{
            Label = "Detail"
            Expression =
            {
                Switch ($_.Color)
                {
                    'Red'    { $color = "91"; break }
                    #'Yellow' { $color = '93'; break }
                    'Green'  { $color = "92"; break }
                    default { $color = "0" }
                }
                $e = [char]27
                "$e[${color}m$($_.Detail)$e[0m"
            }
        },
		@{
			label= 'Seconds'
			Expression = { "{0,6:N0}" -f $_.Seconds }
		},logEntry

    Write-Host ""


	# Print $IdHastable
	#Write-Host "DEBUG: `$IdHastable:`n$($IdHashtable | ConvertTo-Json)"


	# Command line Parameter -ShowErrorsSummary
	if($ShowErrorsSummary) {
		if($PowershellScriptErrorLogs) {
			Write-Host "`n`n`n`n"
			Write-Host "####################################"
			Write-Host "Powershell scripts error logs found:"
			Write-Host "$($PowershellScriptErrorLogs -Join "`n")"
			Write-Host "`n"
		}
	}


	############################################
	#region Export info to text file
	if($ExportTextFileName) {
		if(Test-Path $ExportTextFileName) {
			Write-Host "File $ExportTextFileName already exist. Give another filename to export info to." -ForegroundColor Yellow
			Write-Host "We will not overwrite existing files" -ForegroundColor Yellow
			$ExportTextFileName = $null
		}
	}

	$ExportDateTime = Get-Date -Format 'yyyyMMdd-HHmmss'

	if($ExportTextFileName) {
		Write-Host "Exporting information to text file: $ExportTextFileName"

		"Get-IntuneManagementExtensionDiagnostics.ps1 v1.1 report" | Out-File -FilePath $ExportTextFileName -Force
		"https://github.com/petripaavola/Get-IntuneManagementExtensionDiagnostics`n" | Out-File -FilePath $ExportTextFileName -Append
		"Export DateTime: $ExportDateTime`n`n" | Out-File -FilePath $ExportTextFileName -Append

		"Application download statistics`n" | Out-File -FilePath $ExportTextFileName -Append
		$ApplicationDownloadStatistics | Format-Table | Out-File -FilePath $ExportTextFileName -Append

		"`nOBSERVED TIMELINE`n" | Out-File -FilePath $ExportTextFileName -Append
		$observedTimeline | Select-Object -Property Date, Status, Detail, Seconds, LogEntry | Format-Table | Out-File -FilePath $ExportTextFileName -Append

		if($PowershellScriptErrorLogs) {		
			"`n`nPowershell scripts errors`n" | Out-File -FilePath $ExportTextFileName -Append
			"$($PowershellScriptErrorLogs -Join "`n")"  | Out-File -FilePath $ExportTextFileName -Append
		}
	}
	#endregion Export info to text file
	###################expor#########################



	# Command line Parameter -ConvertAllKnownGuidsToClearText
	if($ConvertAllKnownGuidsToClearText) {
		Foreach($HashtableKey in $IdHashtable.Keys) {
			#Write-Host "Processing Hashtable object $($IdHashtable[$HashtableKey].id)"
			
			Foreach($LogEntry in $LogEntryList) {
				$Logentry.Message = ($Logentry.Message).Replace($HashtableKey, "'      $($IdHashtable[$HashtableKey].name)      '")
			}
			
			Foreach($LogEntry in $LogEntryList | Where-Object ProcessRuntime -ne $null) {
				$Logentry.ProcessRuntime = ($Logentry.ProcessRuntime).Replace($HashtableKey, "'$($IdHashtable[$HashtableKey].name)'")
			}
		}
	}

<#  # This is work in progress so disabled at this time
	# Export data to json file
	if($ExportJson) {
		Write-Host "Export all processed data to json file: IntuneManagementExtensionDiagnostics_export.json"
		[array]$ExportJson = $LogEntryList | Select-Object -Property Index, FileName, Line, DateTime, Multiline, ProcessRunTime, Message, Component, Context, Type, Thread, File
		#$ExportJson = $LogEntryList | Select-Object -Property Index, FileName, Line, DateTime, Multiline, ProcessRunTime, Message, Component, Context, Type, Thread, File #| Sort-Object -Property Index
		$ExportJson | ConvertTo-Json | Out-File -Filepath ./IntuneManagementExtensionDiagnostics_export.json -Force
		Write-Host "Done"
	}
#>

	if($ShowLogViewerUI -or $LogViewerUI) {
		Write-Host "Show logs in Out-GridView"
		
		# Show log in Out-GridView
		if($SelectedLogFiles -is [Array]) {
			# Sort log entries and add index number
			# This is needed with multiple files so log entries
			# can be sorted based on this index column

			$SelectedLines = $LogEntryList | Select-Object -Property Index, FileName, Line, DateTime, Multiline, ProcessRunTime, Message, Component, Context, Type, Thread, File | Sort-Object -Property Index | Out-GridView -Title "Intune IME Log Viewer $($SelectedLogFiles.Name -join " ")" -OutputMode Multiple
		} else {
			
			$SelectedLines = $LogEntryList | Select-Object -Property Index, FileName, Line, DateTime, Multiline, ProcessRunTime, Message, Component, Context, Type, Thread, File | Out-GridView -Title "Intune IME Log Viewer $LogFilePath" -OutputMode Multiple
		}
	} else {

		Write-Host "`nTip: Use Parameter -ShowLogViewerUI to get LogViewerUI for graphical log viewing/debugging`n" -ForegroundColor Cyan
		
		# Not showing Out-GridView so we need to make sure Do-While loop will exit
		$SelectedLines = $null
		Break
	}
	
} While ($SelectedLines.Message -eq 'RELOAD LOG FILE(S).      Select this line and OK button from right botton corner')

#Write-Host "Script end"
