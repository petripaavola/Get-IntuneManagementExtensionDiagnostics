<#PSScriptInfo

.VERSION 2.3

.GUID f321e845-7139-41d7-b0dd-356ea87931e3

.AUTHOR Petri.Paavola@yodamiitti.fi

.COMPANYNAME Yodamiitti Oy

.COPYRIGHT Petri.Paavola@yodamiitti.fi

.TAGS Intune Windows Autopilot troubleshooting log analyzer

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
Version 2.0:  Huge new feature is to create html report
              Html report is primary reporting and console observed timeline is secondary
			  All future development will be done to html report
			  Console timeline will be available for example for OOBE troubleshooting scenarios
			  Added App detection events to timeline
			  Html report entries support HoverOn ToolTips which include more information
Version 2.3:  Updated script to use Microsoft.Graph.Authentication module to download data from Graph API
#>

<#
.Synopsis
   This script analyzes Microsoft Intune Management Extension (IME) log(s) and creates timeline report from found actions.

.DESCRIPTION
   This script analyzes Microsoft Intune Management Extension (IME) log(s) and creates timeline report from found log events.
   
   Report is saved to HTML file. Events are also shown in Powershell console window.
   
   Timeline report includes information about Intune Win32App, WinGetApp, Powershell scripts, Remedation scripts and custom Compliance Policy scripts events. Windows Autopilot ESP phases are also shown on timeline.
   
   Script also includes really capable Log Viewer UI if scripts is started with parameter -ShowLogViewerUI

   LogViewerUI (Out-GridView) looks a lot like cmtrace.exe tool but it is better because all found log actions are added to log for easier debugging.
   
   LogViewerUI has good search and filtering capabilities. Try to filter known log entries in Timeline: Add criteria -> ProcessRunTime -> is not empty.
   
   What really differentiates this LogViewer from other tools is it's capability to convert GUIDs to known names
   try parameter -ConvertAllKnownGuidsToClearText and you can see for example real application names instead of GUIDs on log events.
   
   Selecting last line (RELOAD) and OK will reload log file.
   
   Script can merge multiple log files so especially in LogViewerUI you can see Powershell command outputs from AgentExecutor.log
   
   Powershell command outputs and errors can be also shown in Timeline view with parameters -ShowStdOutInReport and -ShowErrorsInReport
   This shows instantly what is possible problem in Powershell scripts.


   Possible Microsoft 365 App and MSI Line-of-Business Apps (maybe change to Win32App ;) installations are not seen by this report because they are not installed with Intune Management Agent.


   Author:
   Petri.Paavola@yodamiitti.fi
   Senior Modern Management Principal
   Microsoft MVP - Windows and Devices
   
   2024-05-09

   https://github.com/petripaavola/Get-IntuneManagementExtensionDiagnostics

.PARAMETER Online
Download Powershell, Remediation and custom Compliance policy scripts to get displayName to Timeline report
Install Microsoft Graph module with command: Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser

.PARAMETER LogFile
Specify log file fullpath

.PARAMETER LogFilesFolder
Specify folder where to check log files. Will show UI where you can select what logs to process

.PARAMETER LogStartDateTime
Specify date and time to start log entries. For example -

.PARAMETER LogEndDateTime
Specify date and time to stop log entries

.PARAMETER ShowLogViewerUI
Shows graphical LogViewerUI where all log events are easily browsed, searched and filtered in graphical UI
This parameter will always show file selection UI and event selection UI.

.PARAMETER LogViewerUI
Shows graphical LogViewerUI where all log events are easily browsed, searched and filtered in graphical UI
This parameter will always show file selection UI and event selection UI.

.PARAMETER AllLogEvents
Process all found log events.
Selecting this parameter will disable UI which asks date/time/hour selection for logs (use for silent commands or scripts)
This is default option (aka silent and no selection UI shown)

.PARAMETER LogEventsSelectionUI
Selecting this parameter will enable UI which asks date/time/hour selection for logs

.PARAMETER AllLogFiles
Process all found supported log file(s) automatically. This includes *AgentExecutor*.log and *IntuneManagementExtension*.log
Selecting this parameter will disable UI which asks which log files to process (use for silent commands or scripts)
This is default option (aka silent and no selection UI shown)

.PARAMETER LogFilesSelectionUI
Selecting this parameter will enable UI which asks which log files to process

.PARAMETER Today
Show log entries from today (from midnight)

.PARAMETER ShowAllTimelineEvents
Shows more entries in report. This option will show starting messages for events which are not shown by default

.PARAMETER ShowStdOutInReport
Show script StdOut in events. This shows for example what Remediation script will return back to Intune

.PARAMETER ShowErrorsInReport
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

.PARAMETER FindAllLongRunningPowershellScripts
Poweruser option to try to find all long running Powershell scripts over threshold which default is 180 seconds
This option could find long running scripts which we don't even have event in our report.

.PARAMETER DoNotOpenReportAutomatically
Do not open html report file automatically in browser


.EXAMPLE
   .\Get-IntuneManagementExtensionDiagnostics.ps1
.EXAMPLE
   .\Get-IntuneManagementExtensionDiagnostics.ps1 -Online
.EXAMPLE
   .\Get-IntuneManagementExtensionDiagnostics.ps1 -Online -AllLogEvents -AllLogFiles
.EXAMPLE
   .\Get-IntuneManagementExtensionDiagnostics.ps1 -Online -AllLogEvents -ShowAllTimelineEvents
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
   .\Get-IntuneManagementExtensionDiagnostics.ps1 -Online -ShowStdOutInReport
.EXAMPLE
   .\Get-IntuneManagementExtensionDiagnostics.ps1 -Online -ShowErrorsInReport
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
    [Switch]$AllLogEntries=$true,
	[Parameter(Mandatory=$false)]
	[Switch]$LogEntriesSelectionUI,
	[Parameter(Mandatory=$false)]
    [Switch]$AllLogFiles=$true,
	[Parameter(Mandatory=$false)]
	[Switch]$LogFilesSelectionUI,
	[Parameter(Mandatory=$false)]
    [Switch]$Today,
	[Parameter(Mandatory=$false)]
    [Switch]$ShowAllTimelineEvents,
	[Parameter(Mandatory=$false)]
	[Switch]$ShowStdOutInReport,
	[Parameter(Mandatory=$false)]
    [Switch]$ShowErrorsInReport,
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
    [String]$ExportTextFileName=$null,
	[Parameter(Mandatory=$false)]
	[String]$ExportHTMLReportPath=$null,
	[Parameter(Mandatory=$false)]
	[Switch]$FindAllLongRunningPowershellScripts,
	[Parameter(Mandatory=$false)]
	[Switch]$DoNotOpenReportAutomatically
)


$ScriptVersion = "2.3"
$TimeOutBetweenGraphAPIRequests = 300


Write-Host "Get-IntuneManagementExtensionDiagnostics.ps1 $ScriptVersion" -ForegroundColor Cyan
Write-Host "Author: Petri.Paavola@yodamiitti.fi / Microsoft MVP - Windows and Devices"
Write-Host ""


$ExportHTML=$True

# Make script to not show start selection UIs by default
# This should be fixed in the code
# but this was quicker hack to make script silent by default

$AllLogEntries=$true
$AllLogFiles = $true

if($LogEntriesSelectionUI) {
	$AllLogEntries=$false
}
if($LogFilesSelectionUI) {
	$AllLogFiles = $false
}


# Show selection UIs with LogViewerUI because we would like to
# limit as less as events possible to save memory and speed up Out-GridView

# With LogViewerUI show Events selection UI always
# With LogViewerUI show file selection UI always
if($ShowLogViewerUI -or $LogViewerUI) {
	$LogEntriesSelectionUI = $true
	$LogFilesSelectionUI = $true

	$AllLogEntries=$false
	$AllLogFiles = $false
	
	Write-Host "Parameter -ShowLogViewerUI selected. Script will show log file and event selection UIs."
}



# Set variables automatically if we are in Windows Autopilot ESP (Enrollment Status Page)
# Idea is that user can just run the script without checking Parameters first
if($env:UserName -eq 'defaultUser0') {

	Write-Host "Detected running in Windows Autopilot Enrollment Status Page (ESP)" -ForegroundColor Yellow

	#$LOGFile='C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log'

	# Do not open HTML report to browser
	$DoNotOpenReportAutomatically = $true

	if((-not $LogFilesFolder) -or (-not $LOGFile)) {
		Write-Host "Configuring parameters automatically"
		Write-Host "Selected: All log files from default Intune IME logs folder"
		Write-Host "Selected: Do not open HTML report automatically"
		Write-Host

		$LogFilesFolder = 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs'

		if(-not (Test-Path $LogFilesFolder)) {
			Write-Host "Log folder does not exist yet: $LogFilesFolder"
			Write-Host "Try again in a moment..."  -ForegroundColor Yellow
			Write-Host ""
			Exit 0
		}
	}

	# Save Computer Name
	$ComputerNameForReport = $env:ComputerName

	# Process all found supported log files
	# This will not show file selection UI
	$AllLogFiles=$True

	# Process all log entries
	# This will not show time selection UI
	$AllLogEntries=$True

	# Show all entries in Timeline
	# This especially useful if some script or application hangs for a long time
	# so then you can see start entry for that script or application and you know what is current running Intune deployment
	$ShowAllTimelineEvents=$True
	
	if(-not (Test-Path 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log')) {
		Write-Host "Log file does not exist yet: C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log"  -ForegroundColor Yellow
		Write-Host "Try again in a moment..."  -ForegroundColor Yellow
		Write-Host ""
		Exit 0
	}
}


# Hashtable with ids and names
$IdHashtable = @{}

# Save timeline objects to this List
$observedTimeline = [System.Collections.Generic.List[PSObject]]@()

# Save application download statistics to this list
$ApplicationDownloadStatistics = [System.Collections.Generic.List[PSObject]]@()

# Save filtered oud applications to to this list
$ApplicationAssignmentFilterApplied = [System.Collections.Generic.List[PSObject]]@()

# TimeLine entry index
# This might be used in HTML table for sorting entries
$Script:observedTimeLineIndexToHTMLTable=0


################ Functions ################

# This is aligned with Michael Niehaus's Get-AutopilotDiagnostics script just in case
# region Functions
    Function RecordStatusToTimeline {
        param
        (
            [Parameter(Mandatory=$true)] [String] $date,
			[Parameter(Mandatory=$true)] [String] $status,
			[Parameter(Mandatory=$false)] [String] $type,
			[Parameter(Mandatory=$false)] [String] $intent,
			[Parameter(Mandatory=$false)] [String] $detail,
			[Parameter(Mandatory=$false)] $seconds,
            [Parameter(Mandatory=$false)] [String] $logEntry,
            [Parameter(Mandatory=$false)] [String] $color,
			[Parameter(Mandatory=$false)] [String] $DetailToolTip
        )

		$Script:observedTimeLineIndexToHTMLTable++

		# Round seconds to full seconds
		if($seconds) {
			$seconds = [math]::Round($seconds)
		}

		$observedTimeline.add([PSCustomObject]@{
			'Index' = $Script:observedTimeLineIndexToHTMLTable
			'Date' = $date
			'Status' = $status
			'Type' = $type
			'Intent' = $intent
			'Detail' = $detail
			'Seconds' = $seconds
			'LogEntry' = $logEntry
			'Color' = $color
			'DetailToolTip' = $DetailToolTip
			
			})
	}

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
			0	{ $intent = 'Not Targeted' }
			1	{ $intent = 'Available Install' }
			3	{ $intent = 'Required Install' }
			4	{ $intent = 'Required Uninstall' }
			default { $intent = 'Unknown Intent' }
		}
		
		return $intent
	}

	Function Get-AppDetectionResultForNumber {
		Param(
			$DetectionNumber
			)

		Switch ($DetectionNumber)
		{
			0	{ $DetectionState = 'Unknown' }
			1	{ $DetectionState = 'Detected' }
			2	{ $DetectionState = 'Not Detected' }
			3	{ $DetectionState = 'Unknown' }
			4	{ $DetectionState = 'Unknown' }
			5	{ $DetectionState = 'Unknown' }
			default { $DetectionState = 'Unknown' }
		}
		
		return $DetectionState
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


	Function Get-AppType {
		Param(
			$AppId
			)

		$AppType = 'App'

		if($AppId) {
			if($IdHashtable.ContainsKey($AppId)) {
				$AppPolicy=$IdHashtable[$AppId]

				if($AppPolicy.InstallerData) {
					# This should be New Store App
					$AppType = 'WinGetApp'
				} else {
					# This should be Win32App
					$AppType = 'Win32App'
				}
			}
		}
		
		return $AppType
	}

	Function Convert-AppDetectionValuesToHumanReadable {
		Param(
			$DetectionRulesObject
		)

		# Object has DetectionType and DetectionText objects
		# Object is array of objects
		<#		
		[
		  {
			"DetectionType":  2,
			"DetectionText":  {
								  "Path":  "C:\\Program Files (x86)\\Foo",
								  "FileOrFolderName":  "bar.exe",
								  "Check32BitOn64System":  true,
								  "DetectionType":  1,
								  "Operator":  0,
								  "DetectionValue":  null
							  }
		  }
		]
		#>

		foreach($DetectionRule in $DetectionRulesObject) {


			# Change DetectionText properties values to text
			
			# DetectionType: Registry
			if($DetectionRule.DetectionType -eq 0) {
			
				# Registry Detection Type values
				# https://learn.microsoft.com/en-us/graph/api/resources/intune-apps-win32lobappregistrydetectiontype?view=graph-rest-beta
				Switch ($DetectionRule.DetectionText.DetectionType)  {
					0 { $DetectionRule.DetectionText.DetectionType = 'Not configure' }
					1 { $DetectionRule.DetectionText.DetectionType = 'Value exists' }
					2 { $DetectionRule.DetectionText.DetectionType = 'Value does not exist' }
					3 { $DetectionRule.DetectionText.DetectionType = 'String comparison' }
					4 { $DetectionRule.DetectionText.DetectionType = 'Integer comparison' }
					5 { $DetectionRule.DetectionText.DetectionType = 'Version comparison' }
				}
				
				# Registry Detection Operation values
				# https://learn.microsoft.com/en-us/graph/api/resources/intune-apps-win32lobappruleoperator?view=graph-rest-beta
				Switch ($DetectionRule.DetectionText.Operator)  {
					0 { $DetectionRule.DetectionText.Operator = 'Not configured' }
					1 { $DetectionRule.DetectionText.Operator = 'Equals' }
					2 { $DetectionRule.DetectionText.Operator = 'Not equal to' }
					4 { $DetectionRule.DetectionText.Operator = 'Greater than' }
					5 { $DetectionRule.DetectionText.Operator = 'Greater than or equal to' }
					8 { $DetectionRule.DetectionText.Operator = 'Less than' }
					9 { $DetectionRule.DetectionText.Operator = 'Less than or equal to' }
				}
			}


			# DetectionType: File
			if($DetectionRule.DetectionType -eq 2) {

				# File Detection Type values
				# https://learn.microsoft.com/en-us/graph/api/resources/intune-apps-win32lobappfilesystemdetectiontype?view=graph-rest-beta
				Switch ($DetectionRule.DetectionText.DetectionType)  {
					0 { $DetectionRule.DetectionText.DetectionType = 'Not configure' }
					1 { $DetectionRule.DetectionText.DetectionType = 'File or folder exists' }
					2 { $DetectionRule.DetectionText.DetectionType = 'Date modified' }
					3 { $DetectionRule.DetectionText.DetectionType = 'Date created' }
					4 { $DetectionRule.DetectionText.DetectionType = 'String (version)' }
					5 { $DetectionRule.DetectionText.DetectionType = 'Size in MB' }
					6 { $DetectionRule.DetectionText.DetectionType = 'File or folder does not exist' }
				}
				
				# File Detection Operator values
				# https://learn.microsoft.com/en-us/graph/api/resources/intune-apps-win32lobappdetectionoperator?view=graph-rest-beta
				Switch ($DetectionRule.DetectionText.Operator)  {
					0 { $DetectionRule.DetectionText.Operator = 'Not configured' }
					1 { $DetectionRule.DetectionText.Operator = 'Equals' }
					2 { $DetectionRule.DetectionText.Operator = 'Not equal to' }
					4 { $DetectionRule.DetectionText.Operator = 'Greater than' }
					5 { $DetectionRule.DetectionText.Operator = 'Greater than or equal to' }
					8 { $DetectionRule.DetectionText.Operator = 'Less than' }
					9 { $DetectionRule.DetectionText.Operator = 'Less than or equal to' }
				}
			}
			

			# DetectionType: Custom script
			if($DetectionRule.DetectionType -eq 3) {

				# Convert base64 script to clear text
				#$DetectionRule.DetectionText.ScriptBody

				# Decode Base64 content
				$b = [System.Convert]::FromBase64String("$($DetectionRule.DetectionText.ScriptBody)")
				$DetectionRule.DetectionText.ScriptBody = [System.Text.Encoding]::UTF8.GetString($b)

			}
			
			
			<#
			# Change DetectionType value to text
			Switch ($DetectionRule.DetectionType) {
				0 { $DetectionRule.DetectionType = 'Registry' }
				1 { $DetectionRule.DetectionType = 'MSI' }
				2 { $DetectionRule.DetectionType = 'File' }
				3 { $DetectionRule.DetectionType = 'Custom script' }
				default { $DetectionRule.DetectionType = $DetectionRule.DetectionType }
			}
			#>
			
			# Add new property with DetectionType value as text
			Switch ($DetectionRule.DetectionType) {
				0 { $DetectionRule | Add-Member -MemberType noteProperty -Name DetectionTypeAsText -Value 'Registry' }
				1 { $DetectionRule | Add-Member -MemberType noteProperty -Name DetectionTypeAsText -Value 'MSI' }
				2 { $DetectionRule | Add-Member -MemberType noteProperty -Name DetectionTypeAsText -Value 'File' }
				3 { $DetectionRule | Add-Member -MemberType noteProperty -Name DetectionTypeAsText -Value 'Custom script' }
				default { $DetectionRule | Add-Member -MemberType noteProperty -Name DetectionTypeAsText -Value $DetectionRule.DetectionType }
			}			
			
		}			
		
		return $DetectionRulesObject
	}


	function Invoke-MgGraphRequestGetAllPages {
		param (
			[Parameter(Mandatory = $true)]
			[String]$uri
		)

		$MgGraphRequest = $null
		$AllMSGraphRequest = $null

		Start-Sleep -Milliseconds $TimeOutBetweenGraphAPIRequests

		try {

			# Save results to this variable
			$allGraphAPIData = @()

			do {

				$MgGraphRequest = $null
				$MgGraphRequest = Invoke-MgGraphRequest -Uri $uri -Method 'Get' -OutputType PSObject -ContentType "application/json"

				if($MgGraphRequest) {

					# Test if object has attribute named Value (whether value is null or not)
					#if((Get-Member -inputobject $MgGraphRequest -name 'Value' -Membertype Properties) -and (Get-Member -inputobject $MgGraphRequest -name '@odata.context' -Membertype Properties)) {
					if(Get-Member -inputobject $MgGraphRequest -name 'Value' -Membertype Properties) {
						# Value property exists
						$allGraphAPIData += $MgGraphRequest.Value

						# Check if we have value starting https:// in attribute @odate.nextLink
						# and check that $Top= parameter was NOT used. With $Top= parameter we can limit search results
						# but that almost always results .nextLink being present if there is more data than specified with top
						# If we specified $Top= ourselves then we don't want to fetch nextLink values
						#
						# So get GraphAllPages if there is valid nextlink and $Top= was NOT used in url originally
						if (($MgGraphRequest.'@odata.nextLink' -like 'https://*') -and (-not ($uri.Contains('$top=')))) {
							# Save nextLink url to variable and rerun do-loop
							$uri = $MgGraphRequest.'@odata.nextLink'
							Start-Sleep -Milliseconds $TimeOutBetweenGraphAPIRequests

							# Continue to next round in Do-loop
							Continue

						} else {
							# We dont have nextLink value OR
							# $top= exists so we return what we got from first round
							#return $allGraphAPIData
							$uri = $null
						}
						
					} else {
						# Sometimes we get results without Value-attribute (eg. getting user details)
						# We will return all we got as is
						# because there should not be nextLink page in this case ???
						return $MgGraphRequest
					}
				} else {
					# Invoke-MGGraphRequest failed so we return false
					return $null
				}
				
			} while ($uri) # Always run once and continue if there is nextLink value


			# We should not end here but just in case
			return $allGraphAPIData

		} catch {
			Write-Error "There was error with MGGraphRequest with url $url!"
			return $null
		}
	}


	### HTML Report helper functions ###
	
	function Fix-HTMLSyntax {
		Param(
			$html
		)

		$html = $html.Replace('&lt;', '<')
		$html = $html.Replace('&gt;', '>')
		$html = $html.Replace('&quot;', '"')

		return $html
	}

	function Fix-HTMLColumns {
		Param(
			$html
		)

		# Rename column headers
		$html = $html -replace '<th>@odata.type</th>','<th>App type</th>'
		$html = $html -replace '<th>displayname</th>','<th>App name</th>'
		$html = $html -replace '<th>assignmentIntent</th>','<th>Assignment Intent</th>'
		$html = $html -replace '<th>assignmentTargetGroupDisplayName</th>','<th>Target Group</th>'
		$html = $html -replace '<th>assignmentFilterDisplayName</th>','<th>Filter name</th>'
		$html = $html -replace '<th>FilterIncludeExclude</th>','<th>Filter Intent</th>'
		$html = $html -replace '<th>publisher</th>','<th>Publisher</th>'
		$html = $html -replace '<th>productVersion</th>','<th>Version</th>'
		$html = $html -replace '<th>filename</th>','<th>Filename</th>'
		$html = $html -replace '<th>createdDateTime</th>','<th>Created</th>'
		$html = $html -replace '<th>lastModifiedDateTime</th>','<th>Modified</th>'

		return $html
	}


	function return-ObjectPropertiesAsAlignedString {
		Param(
			$object
		)
		# Calculate the maximum length of property names for alignment
		$maxWidth = ($object.PSObject.Properties.Name | Measure-Object -Maximum -Property Length).Maximum

		# Build the formatted string
		$result = $object.PSObject.Properties | ForEach-Object {
			"{0,-$maxWidth} : {1}" -f $_.Name, $_.Value
		}

		# Convert array to single string
		$output = $result -join "`n"
		return $output
	}

# endregion Functions

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
		# Running report using local computer's Intune log files

		# Save Computer Name
		$ComputerNameForReport = $env:ComputerName
		
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
		'Log Start and End time' = 'Last 7 days';
		'LogStartDateTimeObject' = (Get-Date).AddDays(-7);
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

	$GraphAuthenticationModule = $null
	$MgContext = $null

	# Test if we are in Enrollment Status Page (ESP) phase
	# Detect defaultuser0 loggedon
	if($env:UserName -eq 'defaultUser0') {

		# Make sure we can connect
		$GraphAuthenticationModule = Import-Module Microsoft.Graph.Authentication -PassThru -ErrorAction Ignore
		if (-not $GraphAuthenticationModule) {

			Write-Host "Installing module Microsoft.Graph.Authentication"
			Install-Module Microsoft.Graph.Authentication -Force
			$Success = $?
			
			if($Success) {
				Write-Host "Success`n" -ForegroundColor Green

				Write-Host "Import module Microsoft.Graph.Authentication"
				$GraphAuthenticationModule = Import-Module Microsoft.Graph.Authentication -PassThru -ErrorAction Ignore

				if($GraphAuthenticationModule) {
					# Module imported successfully
					Write-Host "Success`n" -ForegroundColor Green					
				} else {
					# Failed to import module
					Write-Host "Failed to import module! Skip downloading script names...`n" -ForegroundColor Red
				}
			} else {
				Write-Host "Failed to install module! Skip downloading script names...`n" -ForegroundColor Red
			}
		}
		
		if($GraphAuthenticationModule) {
			
			Write-Host "Connect to Microsoft Graph API"
			
			$scopes = "DeviceManagementConfiguration.Read.All"
			$MgGraph = Connect-MgGraph -scopes $scopes
			$Success = $?

			if ($Success -and $MgGraph) {
				Write-Host "Success`n" -ForegroundColor Green

				# Get MgGraph session details
				$MgContext = Get-MgContext
				
				if($MgContext) {
				
					$TenantId = $MgContext.TenantId
					$AdminUserUPN = $MgContext.Account

					Write-Host "Connected to Intune tenant:`n$TenantId`n$AdminUserUPN`n"

				} else {
					Write-Host "Error getting MgContext information! Skip downloading script names..." -ForegroundColor Red
				}
				
			} else {
				Write-Host "Could not connect to Graph API! Skip downloading script names..." -ForegroundColor Red
			}
			
		} else {
			Write-Host "Could not connect to Graph API! Skip downloading script names..."  -ForegroundColor Red
		}
	
	} else {
		Write-Host "Connecting to Intune using module Microsoft.Graph.Authentication"

		Write-Host "Import module Microsoft.Graph.Authentication"
		Import-Module Microsoft.Graph.Authentication
		$Success = $?

		if($Success) {
			# Module imported successfully
			Write-Host "Success`n" -ForegroundColor Green
		} else {
			Write-Host "Failed"  -ForegroundColor Red
			Write-Host "Make sure you have installed module Microsoft.Graph.Authentication"
			Write-Host "You can install module without admin rights to your user account with command:`n`nInstall-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser" -ForegroundColor Yellow
			Write-Host "`nor you can install machine-wide module with with admin rights using command:`nInstall-Module -Name Microsoft.Graph.Authentication"
			Write-Host ""
			Exit 1
		}


		Write-Host "Connect to Microsoft Graph API"

		$scopes = "DeviceManagementConfiguration.Read.All"
		$MgGraph = Connect-MgGraph -scopes $scopes
		$Success = $?

		if ($Success -and $MgGraph) {
			Write-Host "Success`n" -ForegroundColor Green

			# Get MgGraph session details
			$MgContext = Get-MgContext
			
			if($MgContext) {
			
				$TenantId = $MgContext.TenantId
				$AdminUserUPN = $MgContext.Account

				Write-Host "Connected to Intune tenant:`n$TenantId`n$AdminUserUPN`n"

			} else {
				Write-Host "Error getting MgContext information!`nScript will exit!" -ForegroundColor Red
				Exit 1
			}
			
		} else {
			Write-Host "Could not connect to Graph API!" -ForegroundColor Red
			Exit 1
		}
	}

	# Download Intune scripts information if we have connection to Microsoft Graph API
	if($MgContext) {

		Write-Host "Download Intune Powershell scripts"
		# Get PowerShell Scripts
		$uri = 'https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts'
		$AllIntunePowershellScripts = Invoke-MgGraphRequestGetAllPages -Uri $uri

		if($AllIntunePowershellScripts) {
			Write-Host "Done" -ForegroundColor Green
			
			# Add Name Property to object
			$AllIntunePowershellScripts | Foreach-Object { $_ | Add-Member -MemberType noteProperty -Name name -Value $_.displayName }
			
			# Add all PowershellScripts to Hashtable
			$AllIntunePowershellScripts | Foreach-Object { $id = $_.id; $value=$_; $IdHashtable["$id"] = $value }
			
		} else {
			Write-Error "Did not find Intune Powershell scripts"
		}

		Start-Sleep -MilliSeconds 500
		
		Write-Host "Download Intune Remediations Scripts"
		# Get Proactive Remediations Scripts
		$uri = 'https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts'
		$AllIntuneProactiveRemediationsScripts = Invoke-MgGraphRequestGetAllPages -Uri $uri

		if($AllIntuneProactiveRemediationsScripts) {
			Write-Host "Done" -ForegroundColor Green
			
			# Add Name Property to object
			$AllIntuneProactiveRemediationsScripts | Foreach-Object { $_ | Add-Member -MemberType noteProperty -Name name -Value $_.displayName }
			
			# Add all PowershellScripts to Hashtable
			$AllIntuneProactiveRemediationsScripts | Foreach-Object { $id = $_.id; $value=$_; $IdHashtable["$id"] = $value }
				
		} else {
			Write-Error "Did not find Intune Remediation scripts"
		}

		Start-Sleep -MilliSeconds 500

		Write-Host "Download Intune Windows Device Compliance custom Scripts"
		# Get Windows Device Compliance custom Scripts
		$uri = 'https://graph.microsoft.com/beta/deviceManagement/deviceComplianceScripts'
		$AllIntuneCustomComplianceScripts = Invoke-MgGraphRequestGetAllPages -Uri $uri

		if($AllIntuneCustomComplianceScripts) {
			Write-Host "Done" -ForegroundColor Green
			
			# Add Name Property to object
			$AllIntuneCustomComplianceScripts | Foreach-Object { $_ | Add-Member -MemberType noteProperty -Name name -Value $_.displayName }
			
			# Add all PowershellScripts to Hashtable
			$AllIntuneCustomComplianceScripts | Foreach-Object { $id = $_.id; $value=$_; $IdHashtable["$id"] = $value }
			
		} else {
			Write-Error "Did not find Intune Windows custom Compliance scripts"
		}
		
		Start-Sleep -MilliSeconds 500
		
		Write-Host "Download Intune Filters"
		$uri = 'https://graph.microsoft.com/beta/deviceManagement/assignmentFilters?$select=*'
		$AllIntuneFilters = Invoke-MgGraphRequestGetAllPages -Uri $uri

		if($AllIntuneFilters) {
			Write-Host "Done" -ForegroundColor Green
			
			# Add all Filters to Hashtable
			$AllIntuneFilters | Foreach-Object { $id = $_.id; $value=$_; $IdHashtable["$id"] = $value }
		} else {
			Write-Error "Did not find Intune filters"
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

	# Info message about Windows 365 Apps and LOB MSI -apps not showing
	$Detail = 'Possible Microsoft 365 Apps and Intune LOB MSI Apps are not shown in this report'
	$DetailToolTip = "Microsoft 365 Apps and `"legacy`" Line-of-business apps (MSI) are installed by Windows MDM channel`nand not installed by Intune Management Extension agent.`n`nCurrent report version does not track those installation.`n`nFuture version might track mentioned app types in Windows Autopilot Enrollment Status Page (ESP) phase.`n`nThis could be a good idea so you would see all Windows Autopilot enrollment process app installations in one report."
	
	RecordStatusToTimeline -date "1900-01-01 23:59:59.0000000" -status 'Info' -detail $Detail -DetailToolTip $DetailToolTip -logEntry "0"

	
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
						
						RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Info" -type "User" -Intent "Logon" -detail "$LoggedOnUser" -logEntry "Line $($LogEntryList[$i].Line)" -Color 'Yellow'
			
						# Set LoggedOnUser information to ProcessRuntime property
						$LogEntryList[$i].ProcessRunTime = "Logged on user: $LoggedOnUser"

						# Extract computer name
						$ComputerNameForReport = $LoggedOnUser.Split('\')[0]

						$i++
						Continue
					}
				} elseif (($LoggedOnUser -notlike '*\defaultuser0') -and ($LoggedOnUser -ne $PreviousLoggedOnUser)) {
					$PreviousLoggedOnUser=$LoggedOnUser

					RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Info" -type "User" -Intent "Logon" -detail "$LoggedOnUser" -logEntry "Line $($LogEntryList[$i].Line)" -Color 'Yellow'
					
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
				
				if($ShowAllTimelineEvents) {

					RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Start" -Type "Powershell script" -Intent "Execute" -detail "`t$PowershellScriptPolicyNameStart for user $PowershellScriptForUserNameStart" -logEntry "Line $($LogEntryList[$i].Line)"
				}
						
				# Set Powershell Script Start information to ProcessRuntime property
				$LogEntryList[$i].ProcessRunTime = "Start Powershell script ($PowershellScriptPolicyNameStart) for user $PowershellScriptForUserNameStart"

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
					
					RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Success" -Type "Powershell script" -Intent "Execute" -detail "$PowershellScriptPolicyNameEnd for user $PowershellScriptForUserNameEndProcess" -Seconds $PowershellScriptRunTimeTotalSeconds -logEntry "Line $($LogEntryList[$i].Line)" -Color 'Green'

					# Check if Powershell script runtime is longer than configured (default 180 seconds)
					if($PowershellScriptRunTimeTotalSeconds -gt $LongRunningPowershellNotifyThreshold) {
						
						$PowershellScriptRunTimeTotalSecondsRounded = [math]::Round($PowershellScriptRunTimeTotalSeconds, 0)
						
						RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Warning" -Type "Powershell script" -Intent "Execute" -detail "   Long Powershell script runtime found ($PowershellScriptRunTimeTotalSecondsRounded seconds)" -detailToolTip "Script runtime was $PowershellScriptRunTimeTotalSecondsRounded seconds" -Seconds $PowershellScriptRunTimeTotalSeconds -logEntry "Line $($LogEntryList[$i].Line)" -Color 'Red'
					}
					
					# Set Powershell Script End information to ProcessRuntime property
					$LogEntryList[$i].ProcessRunTime = "Succeeded Powershell script ($PowershellScriptPolicyNameEnd) for user $PowershellScriptForUserNameEndProcess (Runtime: $PowershellScriptRunTimeTotalSeconds seconds)"
					
				} elseif ($PowershellScriptPolicyResultEndProcess -eq 'Failed') {

					RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Failed" -Type "Powershell script" -Intent "Execute" -detail "$PowershellScriptPolicyNameEnd for user $PowershellScriptForUserNameEndProcess" -Seconds $PowershellScriptRunTimeTotalSeconds -logEntry "Line $($LogEntryList[$i].Line)" -Color 'Red'

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

										

										if($ShowErrorsInReport) {
											
											# Show Powershell error message lines as new entries
											# This was not that good idea???
											<#
											
											# Split Error Message by newline and enter error message line by line
											$ExecutionMsg.Split("`n") | Foreach-Object {
												$ErrorMessageLine = $_
												#$PowershellScriptErrorLogs += "`t`t`t   $ErrorMessageLine"
												$PowershellScriptErrorLogs += "$ErrorMessageLine"
											}
											
											$ExecutionMsg.Split("`n") | Foreach-Object {
												$ErrorMessageLine = $_

												RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "ErrorLog" -Type "Powershell script" -Intent "Run" -detail "$ErrorMessageLine" -logEntry "$($LogEntryList[$i].FileName) line $($LogEntryList[$i].Line)" -Color 'Red'
											}
											#>
											
											# Show Powershell error message as one log entry (one string)
											$ExecutionMsgAsString = $ExecutionMsg #| Out-String
											
											RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "ErrorLog" -Type "Powershell script" -Intent "Execute" -detail "$ExecutionMsgAsString" -logEntry "$($LogEntryList[$i].FileName) line $($LogEntryList[$i].Line)" -Color 'Red'
										} else {
											
											RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Failed" -Type "Powershell script" -Intent "Execute" -detail "See Powershell error message. Hover on with mouse to see details." -detailToolTip $ExecutionMsg -logEntry "Line $($LogEntryList[$i].Line)" -Color 'Red'
											
										}
									}
								}
							} catch {
								# Did not find Powershell error message
							}
						}
					}

				} else {
					
					RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Script end" -Type "Powershell script" -Intent "Execute" -detail "Powershell script $PowershellScriptPolicyNameEnd for user $PowershellScriptForUserNameEndProcess ($PowershellScriptRunTimeTotalSeconds seconds)" -logEntry "Line $($LogEntryList[$i].Line)" -Color 'Yellow'
				}
					
				$PowershellScriptStartLogEntryIndex=$null

				# Continue to next line in Do-while -loop
				$i++
				Continue
			}
		}
		# endregion Find Powershell Script End log entry


		if($FindAllLongRunningPowershellScripts) {
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
						if($ProcessRunTimeTotalSeconds -ge $LongRunningPowershellNotifyThreshold) {
							Write-Host "$($LogEntryList[$b].DateTime) Long running ($ProcessRunTimeTotalSeconds seconds) Powershell script ($LogFileName line $($LogEntryList[$b].Line)): $($LogEntryList[$b].Message)" -ForegroundColor Yellow

							RecordStatusToTimeline -date $LogEntryList[$b].DateTime -status 'Warning' -Type "Powershell script" -Intent 'Execute' -detail "Long running ($ProcessRunTimeTotalSeconds seconds) Powershell script" -detailToolTip "" -Seconds $ProcessRunTimeTotalSeconds -Color 'Red' -logEntry "Line $($LogEntryList[$b].Line)"

						}


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
		}

		###############################################

		# region Find Remediation script type - Proactive Remediation or custom Compliance
		# 20230920 Update. Log entries had changed so this is new approach to find all remediations


		# Find Remediation Detect and Remediate tasks
		if($LogEntryList[$i].Message -like '"C:\Program Files (x86)\Microsoft Intune Management Extension\agentexecutor.exe"  -remediationScript  ""C:\Windows\IMECache\HealthScripts\*\*.ps1""*') {
			# Proactive Remediation script = policyType 6
			# custom Compliance script = policyType 8

			#Write-Host "Remediation Detect script found" -ForegroundColor Red
			$Matches=$null
			#if($LogEntryList[$i].Message -Match '^\"C:\\Program Files \(x86\)\\Microsoft Intune Management Extension\\agentexecutor.exe\"  -remediationScript  \"\"C:\\Windows\\IMECache\\HealthScripts\\(.*)_?\\detect.ps1\"\".*$') {
				
			if($LogEntryList[$i].Message -Match '^"C:\\Program Files \(x86\)\\Microsoft Intune Management Extension\\agentexecutor.exe"  -remediationScript  ""C:\\Windows\\IMECache\\HealthScripts\\(.*)_[0-9]+\\(.*)\.ps1"".*$') {
				$PolicyId = $Matches[1]
				
				# detect
				# remediate
				$RemediationAction = $Matches[2]

				if($RemediationAction -eq 'detect') {
					$intent = 'Detect'
				} elseif($RemediationAction -eq 'remediate') {
					$intent = 'Remediate'
				} else {
					$intent = 'unknown'	
				}
				

				$Thread = $LogEntryList[$i].Thread
				
				$RemediationScriptStartedDateTime = $LogEntryList[$i].DateTime
				$RemediationScriptEndDateTime = $null
				$RemediationDurationDateTimeObject = $null
				$ProcessRunTimeTotalSeconds = 'N/A'

				$RemediationExitCode = $null
				$RemediationScriptName = $null
				$RemediationScriptPowershellScript = $null

				# Try to find displayName
				if($IdHashtable.ContainsKey($PolicyId)) {
					if($IdHashtable[$PolicyId].displayName) {
						$RemediationScriptName = $IdHashtable[$PolicyId].displayName
					} else {
						$RemediationScriptName = $PolicyId
					}
				} else {
					$RemediationScriptName = $PolicyId
				}

				
				# Try to find policyType
				if($IdHashtable.ContainsKey($PolicyId)) {
					$RemediationScriptPolicyType = $IdHashtable[$PolicyId].policyType
				}

				if($RemediationScriptPolicyType -eq 6) {
					$RemediationScriptPolicyTypeName = 'Remediation'
				} elseif ($RemediationScriptPolicyType -eq 8) {
					$RemediationScriptPolicyTypeName = 'Custom Compliance'
				} else {
					# We should never get here because we have seen polityType log entry before this log entry
					# Proactive Remediation script type is unknown
					#$RemediationScriptPolicyTypeName = 'Remediation script (unknown type)'
					$RemediationScriptPolicyTypeName = 'Remediation'
				}


				$LogEntryList[$i].ProcessRunTime = "Start Remediation $RemediationScriptName $intent script"


				#Write-Host "Found Remediation or CompliancePolicy policyId $PolicyId" -ForegroundColor Red


				# Find stdOut, error and end message
				# Find next Remediations log entries

				# These are used to detect if some lines where found
				# If same is found again then something went wrong
				$PowershellCmdLineFound = $False
				
				$AgentExecutorPowershellExitCode = $null
				$AgentExecutorPowershellExitCodeFound = $False

				$AgentExecutorPowershellSuccessMessage = $null
				$AgentExecutorPowershellSuccessMessageFound = $False
				
				$scriptOutputStringArray = $null
				
				# Now we go down (to end)
				for($b = $i + 1; $b -lt $LogEntryList.Count; $b++ ) {

					# Use this to debug Remediations log entries
					#Write-Host "DEBUG: $b - $($LogEntryList[$b].Message)"


					# Try to find stdOut, errorOut and other information
					# before remediation process end log entry
					# This is somewhat tricky as we need to find information from AgentExecutor log
					# And in that log threads are always 1
					# and sometimes for example remediation script and Win32App detection script run can overlap
					# so we can NOT rely that information in AgentExecutor is always for our specific process
					# This is more like best guess and we provide additional information for example in ToolTips


					if(($LogEntryList[$b].Message -Like "cmd line for running powershell is -NoProfile -executionPolicy bypass -file  `"C:\Windows\IMECache\HealthScripts\$($PolicyId)*\*.ps1`"") -and ($LogEntryList[$b].Component -eq 'AgentExecutor')) {
						#Write-Host "Found Remediation powershell command AgentExecutor. Line: $($LogEntryList[$b].Line)" -ForegroundColor Red
						
						$PowershellCmdLineFound = $True
						
						# Continue to next For-loop round
						Continue
					}


					# We are looking for these next lines
					<#
						Powershell exit code is 1	AgentExecutor
						lenth of out=24	AgentExecutor
						lenth of error=2	AgentExecutor
					<----		error from script =	AgentExecutor
					---->			AgentExecutor
						Powershell script is failed to execute	or  Powershell script is successfully executed.
					<----		write output done. output = We need to remediate	AgentExecutor
					 --- 			AgentExecutor
					 --- 		, error = 	AgentExecutor
					---->			AgentExecutor
					#>



					if(($LogEntryList[$b].Message -Match '^Powershell exit code is (.*)$') -and ($LogEntryList[$b].Component -eq 'AgentExecutor') -and ($PowershellCmdLineFound)) {
						$AgentExecutorPowershellExitCode = $Matches[1]
						$AgentExecutorPowershellExitCodeFound = $True
						
						#Write-Host "Found Remediation powershell command AgentExecutor exit code: $($AgentExecutorPowershellExitCode). Line: $($LogEntryList[$b].Line)" -ForegroundColor Red

						# Continue to next For-loop round
						Continue
						
					}


					# Powershell script is failed to execute
					# Powershell script is successfully executed
					if(($LogEntryList[$b].Message -Match '^Powershell script is (.*)$') -and ($LogEntryList[$b].Component -eq 'AgentExecutor') -and ($AgentExecutorPowershellExitCodeFound)) {
						$AgentExecutorPowershellSuccessMessage = $Matches[1]
						$AgentExecutorPowershellSuccessMessageFound = $True
						
						#Write-Host "Found Remediation powershell command AgentExecutor success message: $($AgentExecutorPowershellSuccessMessage). Line: $($LogEntryList[$b].Line)" -ForegroundColor Red
						
						# Next multiline log entry should be possible output and error
						
						# "goto" next log line
						$b++
						
						$scriptOutputStringArray = @()
						if(($LogEntryList[$b].Message -Match '^write output done. (output =.*)$') -and ($LogEntryList[$b].Component -eq 'AgentExecutor') -and ($AgentExecutorPowershellSuccessMessageFound)) {
							$FirstOutputLine = $Matches[1]
							$scriptOutputStringArray += $FirstOutputLine
							
							$b++
							
							# Loop through next lines until we find multiline log entry ----> end character from Multiline property
							While(($LogEntryList[$b].Multiline -ne '---->') -and ($LogEntryList[$b].Component -eq 'AgentExecutor')) {
								$scriptOutputStringArray += $LogEntryList[$b].Message
								
								$b++
							}
						}
					}




					# Remediation script script is done
					# This is our end message we are looking for
					if(($LogEntryList[$b].Message -Match '^Powershell execution is done, exitCode = (.*)$') -and ($LogEntryList[$b].Thread -eq $Thread)) {
						$RemediationExitCode = $Matches[1]
						
						$RemediationScriptEndDateTime = $LogEntryList[$b].DateTime
						
						# Calculate runtime
						if($RemediationScriptStartedDateTime -and $RemediationScriptEndDateTime) {
							$RemediationDurationDateTimeObject = New-Timespan (Get-Date $RemediationScriptStartedDateTime) (Get-Date $RemediationScriptEndDateTime)
							$ProcessRunTimeTotalSeconds = $RemediationDurationDateTimeObject.TotalSeconds
							$ProcessRunTimeTotalSeconds = [math]::Round($ProcessRunTimeTotalSeconds,0)
						}
						#Write-Host "DEBUG: Found Remediation end with exitCode $RemediationExitCode"
						

						$detail = "$RemediationScriptName"
						
						
						if($scriptOutputStringArray) {
							$detailToolTip = $scriptOutputStringArray | Out-String
						} else {
							$detailToolTip = "Runtime information not available"
						}
						
						
						
						if($RemediationExitCode -eq 0) {
							# Remediation Detect Successful
							
							RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Success" -Type "$RemediationScriptPolicyTypeName" -Intent $intent -detail $detail -detailToolTip $detailToolTip -Seconds $ProcessRunTimeTotalSeconds -Color 'Green' -logEntry "Line $($LogEntryList[$i].Line)"
							
							$LogEntryList[$b].ProcessRunTime = "Remediation Detected with exitCode 0 $($RemediationDurationDateTimeObject.Seconds) seconds"
						} else {
							# Remediation Detect was NOT successful
						
							if($RemediationScriptPolicyTypeName -eq 'remediate') {
								$status = 'Failed'
								$color = 'Red'
							} else {
								# Thist should be Detection with Not Detected status
								$status = 'Not Detected'
								$color = 'Yellow'
							}
						
							RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status $status -Type "$RemediationScriptPolicyTypeName" -Intent $intent -detail $detail -detailToolTip $detailToolTip -Seconds $ProcessRunTimeTotalSeconds -Color $color -logEntry "Line $($LogEntryList[$i].Line)"
							
							$LogEntryList[$b].ProcessRunTime = "Remediation Not Detected with exitCode $RemediationExitCode $($RemediationDurationDateTimeObject.Seconds) seconds"
						}
						
						# Break for loop
						Break
						
					}


					# We didn't catch Remediation task end log entry previously
					# because this will catch next remediation script start
					#
					# We should never get here
					if($LogEntryList[$b].Message -like '"C:\Program Files (x86)\Microsoft Intune Management Extension\agentexecutor.exe"  -remediationScript  ""C:\Windows\IMECache\HealthScripts\*\*.ps1""*') {
					
						$status = 'Start'
						$detail = "$RemediationScriptName"
						$detailToolTip = $null
						
						# Print Remediation started but not ended
						RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status $status -Type "$RemediationScriptPolicyTypeName" -Intent $intent -detail $detail -detailToolTip $detailToolTip -Seconds 0 -Color 'Red' -logEntry "Line $($LogEntryList[$i].Line)"
						
						$status = 'Warning'
						$detail = "`tRemediation started but we did not detect process end"
						RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status $status -Type "$RemediationScriptPolicyTypeName" -Intent $intent -detail $detail -detailToolTip $detailToolTip -Seconds 0 -Color 'Red' -logEntry "Line $($LogEntryList[$i].Line)"
						
						# Break for loop
						Break
					}


				} # For-loop end

			}
			
			# Continue to next log entry in main While-loop
			$i++
			Continue
		}


<#
		# endregion Find Remediation script type - Proactive Remediation or custom Compliance

		###############################################

		# region Find Proactive Remediation Detect script type - Proactive Remediation or custom Compliance
		#
		# Note! These log entries have moved to HealthScripts.log so commented out for now
		#
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
					$ProactiveRemediationScriptPolicyTypeName = 'Remediation'
				} elseif ($ProactiveRemediationScriptPolicyType -eq 8) {
					$ProactiveRemediationScriptPolicyTypeName = 'Custom Compliance'
				} else {
					# We should never get here because we have seen polityType log entry before this log entry
					# Proactive Remediation script type is unknown
					$ProactiveRemediationScriptPolicyTypeName = 'Remediation script (unknown type)'
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
							if($ShowAllTimelineEvents) {
								RecordStatusToTimeline -date $LogEntryList[$b].DateTime -status "Start" -Type "$ProactiveRemediationScriptPolicyTypeName" -Intent "Detect" -detail "`t$ProactiveRemediationScriptName" -Seconds $ProcessRunTimeTotalSeconds -logEntry "Line $($LogEntryList[$b].Line)"
							}
							
							# Set RunTime log entry
							$LogEntryList[$b].ProcessRunTime = "Start $ProactiveRemediationScriptPolicyTypeName Detect script ($ProactiveRemediationScriptName) (run as $ProactiveRemediationDetectRunAs)"

							$detailToolTip = $null
							if($ProactiveRemediationDetectStdOutput) {
								$detailToolTip = "Remediation Detect stdOutput:`n$ProactiveRemediationDetectStdOutput"
							}
							
							if($ProactiveRemediationDetectErrorOutput) {
								$detailToolTip = "$detailToolTip`nRemediation Detect errorOutput:`n`n$ProactiveRemediationDetectErrorOutput"
							}


							# End message to Timeline
							if($PreRemediationDetectResult -eq 'True') {
								
								RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Compliant" -Type "$ProactiveRemediationScriptPolicyTypeName" -Intent "Detect" -detail "$ProactiveRemediationScriptName" -detailToolTip $detailToolTip -Seconds $ProcessRunTimeTotalSeconds -Color 'Green' -logEntry "Line $($LogEntryList[$i].Line)"
								
								# Set RunTime log entry
								$LogEntryList[$i].ProcessRunTime = "Compliant $ProactiveRemediationScriptPolicyTypeName Detect script ($ProactiveRemediationScriptName) (Runtime: $ProcessRunTimeTotalSeconds sec)"
	
							} else {
								
								RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Not Compliant" -Type "$ProactiveRemediationScriptPolicyTypeName" -Intent "Detect" -detail "$ProactiveRemediationScriptName" -detailToolTip $detailToolTip -Seconds $ProcessRunTimeTotalSeconds -Color 'Yellow' -logEntry "Line $($LogEntryList[$i].Line)"
								
								$LogEntryList[$i].ProcessRunTime = "Not Compliant $ProactiveRemediationScriptPolicyTypeName Detect script ($ProactiveRemediationScriptName) (Runtime: $ProcessRunTimeTotalSeconds sec)"
							}


							# Script StdOut to Timeline if configured and if exists
							if($ShowStdOutInReport -and $ProactiveRemediationDetectStdOutIndex) {
																
								RecordStatusToTimeline -date $LogEntryList[$ProactiveRemediationDetectStdOutIndex].DateTime -status "Info" -Type "$ProactiveRemediationScriptPolicyTypeName" -Intent "StdOut" -detail "$ProactiveRemediationDetectStdOutput" -logEntry "Line $($LogEntryList[$ProactiveRemediationDetectStdOutIndex].Line)" -Color 'Yellow'
								
								$LogEntryList[$ProactiveRemediationDetectStdOutIndex].ProcessRunTime = "       Powershell Detect StdOut"
							}

							# Script Error to Timeline if configured and if exists
							if($ShowErrorsInReport -and $ProactiveRemediationDetectErrorOutputIndex) {
								
								RecordStatusToTimeline -date $LogEntryList[$ProactiveRemediationDetectErrorOutputIndex].DateTime -status "Error" -Intent "ErrorLog" -Type "$ProactiveRemediationScriptPolicyTypeName" -detail "$ProactiveRemediationDetectErrorOutput" -logEntry "Line $($LogEntryList[$ProactiveRemediationDetectStdOutIndex].Line)" -Color 'Red'
								
								$LogEntryList[$ProactiveRemediationDetectErrorOutputIndex].ProcessRunTime = "       Powershell Detect ErrorLog"
								
							}

							# Post-End message to Timeline if exists
							if($RemediationDetectPostActionMessageIndex) {
								
								$RemediationDetectPostActionMessage = $LogEntryList[$RemediationDetectPostActionMessageIndex].Message
								if($RemediationDetectPostActionMessage -eq '[HS] no remediation script, skip remediation script') {
									$RemediationDetectPostActionMessage = "`tremediation script does not exist"
								}

								if($RemediationDetectPostActionMessage -eq '[HS] remediation is not optional, kick off remediation script') {
									$RemediationDetectPostActionMessage = "`tkick off remediation script"
								}


								RecordStatusToTimeline -date $LogEntryList[$RemediationDetectPostActionMessageIndex].DateTime -status "Info" -Type "$ProactiveRemediationScriptPolicyTypeName" -detail "$RemediationDetectPostActionMessage" -logEntry "Line $($LogEntryList[$RemediationDetectPostActionMessageIndex].Line)" -Color 'Yellow'
							}

							
						} else {
							# We missed Proactive Remediation Detect start log entry
							# We should never get here
							
							if($PreRemediationDetectResult -eq 'True') {
								
								RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Compliant" -Type "$ProactiveRemediationScriptPolicyTypeName" -Intent "Detect" -detail "$ProactiveRemediationScriptName" -Color 'Green' -logEntry "Line $($LogEntryList[$i].Line)"
								
								# Set RunTime log entry
								$LogEntryList[$i].ProcessRunTime = "Compliant $ProactiveRemediationScriptPolicyTypeName Detect script ($ProactiveRemediationScriptName)"
	
							} else {
								RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Not Compliant" -Type "$ProactiveRemediationScriptPolicyTypeName" -Intent "Detect" -detail "$ProactiveRemediationScriptName" -Color 'Yellow' -logEntry "Line $($LogEntryList[$i].Line)"
								
								$LogEntryList[$i].ProcessRunTime = "Not Compliant $ProactiveRemediationScriptPolicyTypeName Detect script ($ProactiveRemediationScriptName)"
							}
							
							# Script StdOut to Timeline if configured and if exists
							if($ShowStdOutInReport -and $ProactiveRemediationDetectStdOutIndex) {
																
								RecordStatusToTimeline -date $LogEntryList[$ProactiveRemediationDetectStdOutIndex].DateTime -status "Info" -Type "$ProactiveRemediationScriptPolicyTypeName" -Intent "StdOut" -detail "$ProactiveRemediationDetectStdOutput" -logEntry "Line $($LogEntryList[$ProactiveRemediationDetectStdOutIndex].Line)" -Color 'Yellow'
								
								$LogEntryList[$ProactiveRemediationDetectStdOutIndex].ProcessRunTime = "   Powershell Detect StdOut"
							}

							# Script Error to Timeline if configured and if exists
							if($ShowErrorsInReport -and $ProactiveRemediationDetectErrorOutputIndex) {
								
								RecordStatusToTimeline -date $LogEntryList[$ProactiveRemediationDetectErrorOutputIndex].DateTime -status "Error" -Type "$ProactiveRemediationScriptPolicyTypeName" -Intent "ErrorLog" -detail "$ProactiveRemediationDetectErrorOutput" -logEntry "Line $($LogEntryList[$ProactiveRemediationDetectStdOutIndex].Line)" -Color 'Red'
								
								$LogEntryList[$ProactiveRemediationDetectErrorOutputIndex].ProcessRunTime = "   Powershell Detect ErrorLog"
								
							}

							# Post-End message to Timeline if exists
							if($RemediationDetectPostActionMessageIndex) {
								
								RecordStatusToTimeline -date $LogEntryList[$RemediationDetectPostActionMessageIndex].DateTime -status "Info" -Type "$ProactiveRemediationScriptPolicyTypeName" -detail "   $($LogEntryList[$RemediationDetectPostActionMessageIndex].Message)" -logEntry "Line $($LogEntryList[$RemediationDetectPostActionMessageIndex].Line)" -Color 'Yellow'
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
							
							RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Compliant" -Type "$ProactiveRemediationScriptPolicyTypeName" -Intent "Detect" -detail "$ProactiveRemediationScriptName" -Color 'Green' -logEntry "Line $($LogEntryList[$i].Line)"
							
							# Set RunTime log entry
							$LogEntryList[$i].ProcessRunTime = "Compliant $ProactiveRemediationScriptPolicyTypeName Detection script ($ProactiveRemediationScriptName)"

						} else {
							RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Not Compliant" -Type "$ProactiveRemediationScriptPolicyTypeName" -detail "$ProactiveRemediationScriptName" -Color 'Yellow' -logEntry "Line $($LogEntryList[$i].Line)"
							
							$LogEntryList[$i].ProcessRunTime = "Not Compliant $ProactiveRemediationScriptPolicyTypeName Detection script ($ProactiveRemediationScriptName)"
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
						if($ShowAllTimelineEvents) {
							RecordStatusToTimeline -date $LogEntryList[$b].DateTime -status "Start" -Type "$ProactiveRemediationScriptPolicyTypeName" -Intent "Remediate" -detail "`t$ProactiveRemediationScriptName (run as $RemediationScriptRunAs)" -logEntry "Line $($LogEntryList[$b].Line)"
						}
						
						# Set RunTime log entry
						$LogEntryList[$b].ProcessRunTime = "Processing $ProactiveRemediationScriptPolicyTypeName Remediate script ($ProactiveRemediationScriptName) (run as $RemediationScriptRunAs)"


						$detailToolTip = $null
						if($RemediationScriptStdOutput) {
							$detailToolTip = "Remediation Script stdOutput:`n$RemediationScriptStdOutput"
						}
						
						if($RemediationScriptErrorOutput) {
							$detailToolTip = "$detailToolTip`nRemediation Script errorOutput:`n`n$RemediationScriptErrorOutput"
						}


						# End message to Timeline
						if($RemediationScriptExitCode -eq 0) {
							
							RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Success" -Type "$ProactiveRemediationScriptPolicyTypeName" -Intent "Remediate" -detail "$ProactiveRemediationScriptName" -detailToolTip $detailToolTip -Seconds $ProcessRunTimeTotalSeconds -Color 'Green' -logEntry "Line $($LogEntryList[$i].Line)"
							
							# Set RunTime log entry
							$LogEntryList[$i].ProcessRunTime = "Succeeded $ProactiveRemediationScriptPolicyTypeName Detection script ($ProactiveRemediationScriptName) (Runtime: $ProcessRunTimeTotalSeconds sec)"

						} else {
															
							RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Failed" -Type "$ProactiveRemediationScriptPolicyTypeName" -Intent "Remediate" -detail "$ProactiveRemediationScriptName" -detailToolTip $detailToolTip -Seconds $ProcessRunTimeTotalSeconds -Color 'Red' -logEntry "Line $($LogEntryList[$i].Line)"
							
							$LogEntryList[$i].ProcessRunTime = "Failed $ProactiveRemediationScriptPolicyTypeName Remediate script ($ProactiveRemediationScriptName) (Runtime: $ProcessRunTimeTotalSeconds sec)"
						}

						# Script StdOut to Timeline if configured and if exists
						if($ShowStdOutInReport -and $RemediationScriptStdOutputIndex) {
															
							RecordStatusToTimeline -date $LogEntryList[$RemediationScriptStdOutputIndex].DateTime -status "Info" -Type "$ProactiveRemediationScriptPolicyTypeName" -Intent "StdOut" -detail "$RemediationScriptStdOutput" -logEntry "Line $($LogEntryList[$RemediationScriptStdOutputIndex].Line)" -Color 'Yellow'
							
							$LogEntryList[$RemediationScriptStdOutputIndex].ProcessRunTime = "       Powershell Remediate StdOut"
						}

						# Script Error to Timeline if configured and if exists
						if($ShowErrorsInReport -and $RemediationScriptErrorOutputIndex) {
															
							RecordStatusToTimeline -date $LogEntryList[$RemediationScriptErrorOutputIndex].DateTime -status "Error" -Type "$ProactiveRemediationScriptPolicyTypeName" -Intent "ErrorLog" -detail "$RemediationScriptErrorOutput" -logEntry "Line $($LogEntryList[$RemediationScriptErrorOutputIndex].Line)" -Color 'Red'
							
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
						
						RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Compliant" -Type "$ProactiveRemediationScriptPolicyTypeName" -detail "$ProactiveRemediationScriptName" -Color 'Green' -logEntry "Line $($LogEntryList[$i].Line)"
						
						# Set RunTime log entry
						$LogEntryList[$i].ProcessRunTime = "Compliant $ProactiveRemediationScriptPolicyTypeName Detection script ($ProactiveRemediationScriptName)"

					} else {
						RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Not Compliant" -Type "$ProactiveRemediationScriptPolicyTypeName" -detail "$ProactiveRemediationScriptName" -Color 'Yellow' -logEntry "Line $($LogEntryList[$i].Line)"
						
						$LogEntryList[$i].ProcessRunTime = "Not Compliant $ProactiveRemediationScriptPolicyTypeName Detection script ($ProactiveRemediationScriptName)"
					}
					
					# Script StdOut to Timeline if configured and if exists
					if($ShowStdOutInReport -and $RemediationScriptStdOutputIndex) {
														
						RecordStatusToTimeline -date $LogEntryList[$RemediationScriptStdOutputIndex].DateTime -status "Info" -Intent "StdOut" -Type "$ProactiveRemediationScriptPolicyTypeName" -detail "$RemediationScriptStdOutput" -logEntry "Line $($LogEntryList[$RemediationScriptStdOutputIndex].Line)" -Color 'Yellow'
						
						$LogEntryList[$RemediationScriptStdOutputIndex].ProcessRunTime = "       Powershell Remediate StdOut"
					}

					# Script Error to Timeline if configured and if exists
					if($ShowErrorsInReport -and $RemediationScriptErrorOutputIndex) {
														
						RecordStatusToTimeline -date $LogEntryList[$RemediationScriptErrorOutputIndex].DateTime -status "Error" -Intent "ErrorLog" -Type "$ProactiveRemediationScriptPolicyTypeName" -detail "$RemediationScriptErrorOutput" -logEntry "Line $($LogEntryList[$RemediationScriptErrorOutputIndex].Line)" -Color 'Red'
						
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
					$ProactiveRemediationScriptPolicyTypeName = 'Remediation'
				} elseif ($ProactiveRemediationScriptPolicyType -eq 8) {
					$ProactiveRemediationScriptPolicyTypeName = 'Custom Compliance'
				} else {
					# We should never get here because we have seen polityType log entry before this log entry
					# Proactive Remediation script type is unknown
					$ProactiveRemediationScriptPolicyTypeName = 'Remediation script (unknown type)'
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
							if($ShowAllTimelineEvents) {
								
								RecordStatusToTimeline -date $LogEntryList[$b].DateTime -status "Start" -Type "$ProactiveRemediationScriptPolicyTypeName" -Intent "PostDetect" -detail "`t$ProactiveRemediationScriptName (run as $ProactiveRemediationPostDetectRunAs)" -logEntry "Line $($LogEntryList[$b].Line)"
							}
							
							# Set RunTime log entry
							$LogEntryList[$b].ProcessRunTime = "Start $ProactiveRemediationScriptPolicyTypeName PostDetect script ($ProactiveRemediationScriptName) (run as $ProactiveRemediationPostDetectRunAs)"

							$detailToolTip = $null
							if($ProactiveRemediationPostDetectStdOutput) {
								$detailToolTip = "Remediation PostDetect stdOutput:`n$ProactiveRemediationPostDetectStdOutput"
							}
							
							if($ProactiveRemediationDetectErrorOutput) {
								$detailToolTip = "$detailToolTip`nRemediation PostDetect errorOutput:`n`n$ProactiveRemediationPostDetectErrorOutput"
							}

							# End message to Timeline
							if($PostRemediationDetectResult -eq 'True') {
								
								RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Compliant" -Type "$ProactiveRemediationScriptPolicyTypeName" -Intent "PostDetect" -detail "$ProactiveRemediationScriptName" -detailToolTip $detailToolTip -Seconds $ProcessRunTimeTotalSeconds -Color 'Green' -logEntry "Line $($LogEntryList[$i].Line)"
								
								# Set RunTime log entry
								$LogEntryList[$i].ProcessRunTime = "Compliant $ProactiveRemediationScriptPolicyTypeName PostDetect script ($ProactiveRemediationScriptName) (Runtime: $ProcessRunTimeTotalSeconds sec)"
	
							} else {
								
								RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Not Compliant" -Type "$ProactiveRemediationScriptPolicyTypeName" -Intent "PostDetect" -detail "$ProactiveRemediationScriptName" -detailToolTip $detailToolTip -Seconds $ProcessRunTimeTotalSeconds -Color 'Red' -logEntry "Line $($LogEntryList[$i].Line)"
								
								$LogEntryList[$i].ProcessRunTime = "Not Compliant $ProactiveRemediationScriptPolicyTypeName PostDetect script ($ProactiveRemediationScriptName) (Runtime: $ProcessRunTimeTotalSeconds sec)"
							}

							# Script StdOut to Timeline if configured and if exists
							if($ShowStdOutInReport -and $ProactiveRemediationPostDetectStdOutIndex) {
																
								RecordStatusToTimeline -date $LogEntryList[$ProactiveRemediationPostDetectStdOutIndex].DateTime -status "Info" -Type "$ProactiveRemediationScriptPolicyTypeName" -Intent "StdOut" -detail "$ProactiveRemediationPostDetectStdOutput" -logEntry "Line $($LogEntryList[$ProactiveRemediationPostDetectStdOutIndex].Line)" -Color 'Yellow'
								
								$LogEntryList[$ProactiveRemediationPostDetectStdOutIndex].ProcessRunTime = "       Powershell PostDetect StdOut"
							}

							# Script Error to Timeline if configured and if exists
							if($ShowErrorsInReport -and $ProactiveRemediationPostDetectErrorOutputIndex) {
								
								RecordStatusToTimeline -date $LogEntryList[$ProactiveRemediationPostDetectErrorOutputIndex].DateTime -status "Error" -Type "$ProactiveRemediationScriptPolicyTypeName" -Intent "ErrorLog" -detail "$ProactiveRemediationPostDetectErrorOutput" -logEntry "Line $($LogEntryList[$ProactiveRemediationPostDetectStdOutIndex].Line)" -Color 'Red'
								
								$LogEntryList[$ProactiveRemediationPostDetectErrorOutputIndex].ProcessRunTime = "       Powershell PostDetect ErrorLog"
								
							}

					
						} else {
							# We missed Proactive Remediation Detect start log entry
							# We should never get here
							
							$Padding = ''
							if($PostRemediationDetectResult -eq 'True') {
								
								RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Compliant" -Type "$ProactiveRemediationScriptPolicyTypeName" -Intent "PostDetect" -detail "$ProactiveRemediationScriptName" -Color 'Green' -logEntry "Line $($LogEntryList[$i].Line)"
								
								# Set RunTime log entry
								$LogEntryList[$i].ProcessRunTime = "Compliant $ProactiveRemediationScriptPolicyTypeName PostDetect script ($ProactiveRemediationScriptName)"
	
							} else {
								RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Not Compliant" -Type "$ProactiveRemediationScriptPolicyTypeName" -Intent "PostDetect" -detail "$ProactiveRemediationScriptName" -Color 'Red' -logEntry "Line $($LogEntryList[$i].Line)"
								
								$LogEntryList[$i].ProcessRunTime = "Not Compliant $ProactiveRemediationScriptPolicyTypeName PostDetect script ($ProactiveRemediationScriptName)"
							}
							
							
							# Script StdOut to Timeline if configured and if exists
							if($ShowStdOutInReport -and $ProactiveRemediationPostDetectStdOutIndex) {
																
								RecordStatusToTimeline -date $LogEntryList[$ProactiveRemediationPostDetectStdOutIndex].DateTime -status "Info" -Intent "StdOut" -Type "$ProactiveRemediationScriptPolicyTypeName" -Intent "PostDetect" -detail "$ProactiveRemediationPostDetectStdOutput" -logEntry "Line $($LogEntryList[$ProactiveRemediationPostDetectStdOutIndex].Line)" -Color 'Yellow'
								
								$LogEntryList[$ProactiveRemediationPostDetectStdOutIndex].ProcessRunTime = "       Powershell PostDetect StdOut"
							}

							# Script Error to Timeline if configured and if exists
							if($ShowErrorsInReport -and $ProactiveRemediationPostDetectErrorOutputIndex) {
								
								RecordStatusToTimeline -date $LogEntryList[$ProactiveRemediationPostDetectErrorOutputIndex].DateTime -status "Error" -Type "$ProactiveRemediationScriptPolicyTypeName" -Intent "ErrorLog" -detail "$ProactiveRemediationPostDetectErrorOutput" -logEntry "Line $($LogEntryList[$ProactiveRemediationPostDetectStdOutIndex].Line)" -Color 'Red'
								
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
							
							RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Compliant" -Type "$ProactiveRemediationScriptPolicyTypeName" -Intent "PostDetect" -detail "$ProactiveRemediationScriptName" -Color 'Green' -logEntry "Line $($LogEntryList[$i].Line)"
							
							# Set RunTime log entry
							$LogEntryList[$i].ProcessRunTime = "Compliant $ProactiveRemediationScriptPolicyTypeName PostDetect script ($ProactiveRemediationScriptName)"

						} else {
							RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Not Compliant" -Type "$ProactiveRemediationScriptPolicyTypeName" -Intent "PostDetect" -detail "$ProactiveRemediationScriptName" -Color 'Red' -logEntry "Line $($LogEntryList[$i].Line)"
							
							$LogEntryList[$i].ProcessRunTime = "Not Compliant $ProactiveRemediationScriptPolicyTypeName PostDetect script ($ProactiveRemediationScriptName)"
						}
						
						# Script StdOut to Timeline if configured and if exists
						if($ShowStdOutInReport -and $ProactiveRemediationPostDetectStdOutIndex) {
															
							RecordStatusToTimeline -date $LogEntryList[$ProactiveRemediationPostDetectStdOutIndex].DateTime -status "Info" -Type "$ProactiveRemediationScriptPolicyTypeName" -Intent "StdOut" -detail "$ProactiveRemediationPostDetectStdOutput" -logEntry "Line $($LogEntryList[$ProactiveRemediationPostDetectStdOutIndex].Line)" -Color 'Yellow'
							
							$LogEntryList[$ProactiveRemediationPostDetectStdOutIndex].ProcessRunTime = "       Powershell PostDetect StdOut"
						}

						# Script Error to Timeline if configured and if exists
						if($ShowErrorsInReport -and $ProactiveRemediationPostDetectErrorOutputIndex) {
							
							RecordStatusToTimeline -date $LogEntryList[$ProactiveRemediationPostDetectErrorOutputIndex].DateTime -status "Error" -Type "$ProactiveRemediationScriptPolicyTypeName" -Intent "ErrorLog" -detail "$ProactiveRemediationPostDetectErrorOutput" -logEntry "Line $($LogEntryList[$ProactiveRemediationPostDetectStdOutIndex].Line)" -Color 'Red'
							
							$LogEntryList[$ProactiveRemediationPostDetectErrorOutputIndex].ProcessRunTime = "       Powershell PostDetect ErrorLog"
							
						}
					
						Break
					}
					
				}
			} # For-loop end
			
		}
		# endregion Try to find Proactive PostDetection script run
#>


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
					

					#####################################
					# Save App Assignments with filters applied
					# We will show this in own table in report
					#
					# "AppApplicabilityStateDueToAssginmentFilters":  null
					# no filters
					#
					# "AppApplicabilityStateDueToAssginmentFilters":  0
					# filters are applied but not preventing app installation
					#
					# "AppApplicabilityStateDueToAssginmentFilters":  1010
					# filters are applied and preventing app installation
					#
					#         "AppApplicabilityStateDueToAssginmentFilters":  1010,
					#		  "AssignmentFilterIds":  [
                    #                "149021e6-c4cf-4f1b-ba76-b98d792a4a4a"
                    #            ],
					
					if($AppPolicy.AssignmentFilterIds) {
						if(-not ($ApplicationAssignmentFilterApplied | Where-Object AppId -eq $AppPolicy.Id)) {
							# App was not in list before
							
							if($AppPolicy.InstallerData) {
								# This should be New Store App
								$AppType = 'WinGetApp'
							} else {
								# This should be Win32App
								$AppType = 'Win32App'
							}

							if($AppPolicy.AppApplicabilityStateDueToAssginmentFilters -eq 1010) {
								# This should be notApplicable
								$Applicable = 'Not Applicable'
							} else {
								# This should be Applicable
								$Applicable = 'Applicable'
							}
						
							# Get AssignmentFilterNames
							$AssignmentFilterNames = $null
							Foreach($FilterId in $AppPolicy.AssignmentFilterIds) {
								if($IdHashtable.ContainsKey($FilterId)) {
									$FilterName = $IdHashtable[$FilterId].displayName
								} else {
									$FilterName = "$FilterId"
								}
								if($AssignmentFilterNames) {
									# There was existing FilterName so we have multiple Filters applied
									$AssignmentFilterNames = "$AssignmentFilterNames, $FilterName"
								} else {
									# There was no previous FilterName
									$AssignmentFilterNames = "$FilterName"
								}
							}
							
							# Add Assignment Filter information to list
							$ApplicationAssignmentFilterApplied.add([PSCustomObject]@{
								'AppType' = $AppType
								'Applicable' = $Applicable
								'AppName' = $AppPolicy.Name
								'AssignmentFilterNames' = $AssignmentFilterNames
								'AssignmentFilterIds' = $AppPolicy.AssignmentFilterIds
								'AppPolicy' = $AppPolicy
								'AppId' = $AppPolicy.Id
								})
						}
					}
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

					RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Info" -Type "Apps" -Intent "Required in ESP" -detail "$AppCountToInstallInESPPhase Apps to install at ESP phase" -logEntry "Line $($LogEntryList[$i].Line)"


					# Several apps are separated by ,
					$AppIdsToInstallInESPPhaseArray = $AppIdsToInstallInESPPhase.Split(',').Trim()

					Foreach($AppId in $AppIdsToInstallInESPPhaseArray) {
						
						# Get name for App
						if($IdHashtable.ContainsKey($AppId)) {
							$AppName = $IdHashtable[$AppId].name
							$AppType = Get-AppType $AppId
							#$DetailToolTip = $IdHashtable[$AppId] | ConvertFrom-Json | ConvertTo-Json -Depth 4
							
							$DetailToolTip = $null
							#$DetailToolTip = $IdHashtable[$AppId] | Select-Object -Property * -ExcludeProperty AppApplicabilityStateDueToAssginmentFilters, ToastState, FlatDepedencies, MetadataVersion, RelationVersion, RebootEx, StartDeadlineEx, RemoveUserData, DOPriority, newFlatDependencies, FlatDependencies, ContentCacheDuration, ReevaluationInterval, SupportState, AssignmentFilterIdToEvalStateMap | Out-String
							
							$DetailToolTip = $IdHashtable[$AppId] | Select-Object -Property Name, InstallCommandLine, UninstallCommandLine, id | Format-List -Property * | Out-String
							
							# DEBUG App Data
							#$IdHashtable[$AppId] | ConvertTo-Json -Depth 5 | Set-Clipboard
							
							RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "Info" -Type "$AppType" -Intent "Required in ESP" -detail "`t$AppName" -DetailToolTip $DetailToolTip -logEntry "Line $($LogEntryList[$i].Line)"
							
							$DetailToolTip = $null
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

			$DetailToolTip = $null

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
								
								$DetailToolTip = $IdHashtable[$Win32AppId] | Select-Object -Property Name, InstallCommandLine, UninstallCommandLine, id | Format-List -Property * | Out-String
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

					#$ApplicationDownloadStatistics.Add($DownloadStats)
					#$DownloadStats = "DL:$($DownloadDurationSeconds)s,$($TotalBytesDownloadedMBRounded)MB,$($DownloadMBps)MB/s,DO:$($BytesFromPeersPercentage)"
					#Write-Host "$($Win32AppName): $DownloadStats"
				}
				
				if($FailedToCreateInstallerProcess) {
					# Failed to create process

					if($AppSupersededUninstall) {
						#$detail = "{0,$ColonIndentInTimeLine}{1}  {2}" -f "$Intent", ':', "$Win32AppName (Superseded) $(($LogEntryList[$Win32AppInstallStartFailedIndex].Message).Replace('[Win32App]','').Trim())"
						$detail = "$Win32AppName (Superseded) $(($LogEntryList[$Win32AppInstallStartFailedIndex].Message).Replace('[Win32App]','').Trim())"

					} else {
						#$detail = "{0,$ColonIndentInTimeLine}{1}  {2}" -f "$Intent", ':', "$Win32AppName $(($LogEntryList[$Win32AppInstallStartFailedIndex].Message).Replace('[Win32App]','').Trim())"
						$detail = "$Win32AppName $(($LogEntryList[$Win32AppInstallStartFailedIndex].Message).Replace('[Win32App]','').Trim())"
					}

					# Set RunTime to Win32App Install Start Failed log entry
					$LogEntryList[$Win32AppInstallStartFailedIndex].ProcessRunTime = "Failed $Win32AppName $intent ($ErrorMessage)"

					RecordStatusToTimeline -date $LogEntryList[$Win32AppInstallStartFailedIndex].DateTime -status "Failed" -Type "Win32App" -Intent $Intent -detail $detail -DetailToolTip $DetailToolTip -logEntry "Line $($LogEntryList[$Win32AppInstallStartFailedIndex].Line)" -Seconds '' -Color 'Red'

					# Set Info message about possible missing uninstall executable
					# First check that uninstall command is correct.
					#
					# Another case could be that there was earlier uninstall process started while application was running
					# Uninstaller then could not delete main executable because it was locked but it deleted uninstall.exe
					# And if Detection Check is for main executable then partially uninstalled App is detected for old version
					# So Intune always detects old App version and is trying to uninstall it before installing new App version but uninstaller is missing

					#$detail = "{0,$ColonIndentInTimeLine}{1}  {2}" -f "", ':', "   Incorrect uninstall command line or uninstall .exe missing?"
					$detail = "   Incorrect uninstall command line or uninstall .exe missing?"

					RecordStatusToTimeline -date $LogEntryList[$Win32AppInstallStartFailedIndex].DateTime -status "Info" -Intent $intent -Type "Win32App" -detail $detail -DetailToolTip $DetailToolTip -logEntry "Line $($LogEntryList[$Win32AppInstallStartFailedIndex].Line)" -Seconds '' -Color 'Yellow'

					if($IdHashtable.ContainsKey($Win32AppId)) {
						if($IdHashtable[$Win32AppId].UninstallCommandLine) {
							$UninstallCommandLine = $IdHashtable[$Win32AppId].UninstallCommandLine
							#$detail = "{0,$ColonIndentInTimeLine}{1}  {2}" -f "", ':', "   Uninstall command: $UninstallCommandLine"
							$detail = "   Uninstall command: $UninstallCommandLine"

							RecordStatusToTimeline -date $LogEntryList[$Win32AppInstallStartFailedIndex].DateTime -status "Info" -Type "Win32App" -Intent $intent -detail $detail -DetailToolTip $DetailToolTip -logEntry "Line $($LogEntryList[$Win32AppInstallStartFailedIndex].Line)" -Seconds '' -Color 'Yellow'

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
						#$detail = "{0,$ColonIndentInTimeLine}{1}  {2}" -f "$Intent", ':', "$Win32AppName ($Win32AppExitCode $Win32AppInstallSuccess) (Superseded)"
						$detail = "$Win32AppName ($Win32AppExitCode $Win32AppInstallSuccess) (Superseded)"
					} else {
						#$detail = "{0,$ColonIndentInTimeLine}{1}  {2}" -f "$Intent", ':', "$Win32AppName ($Win32AppExitCode $Win32AppInstallSuccess)"
						$detail = "$Win32AppName ($Win32AppExitCode $Win32AppInstallSuccess)"
					}
					
					# Change text color depending on Win32App install Success status
					if($Win32AppInstallSuccess -eq 'Success') {
						
						RecordStatusToTimeline -date $LogEntryList[$Win32AppInstallEndIndex].DateTime -status "Success" -Type "Win32App" -Intent $intent -detail $detail -DetailToolTip $DetailToolTip -logEntry "Line $($LogEntryList[$Win32AppInstallEndIndex].Line)" -Seconds $ProcessRunTimeTotalSeconds -Color 'Green'

					} else {
						
						RecordStatusToTimeline -date $LogEntryList[$Win32AppInstallEndIndex].DateTime -status "Failed" -Type "Win32App" -Intent $intent -detail $detail -DetailToolTip $DetailToolTip -logEntry "Line $($LogEntryList[$Win32AppInstallEndIndex].Line)" -Seconds $ProcessRunTimeTotalSeconds -Color 'Red'
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

<#
# Not Applicable checks
# For future TODO list

<![LOG[[Win32App][V3Processor] Processing subgraph with app ids: 9a7d745d-9cd2-4b05-adc2-7796f8c8cc87]LOG]!><time="08:45:59.2422749" date="9-13-2023" component="IntuneManagementExtension" context="" type="1" thread="5" file="">
<![LOG[[Win32App][V3Processor] All apps in the subgraph are not applicable due to assignment filters. Skipping processing.]LOG]!><time="08:45:59.2422749" date="9-13-2023" component="IntuneManagementExtension" context="" type="1" thread="5" file="">

#>



		# region Find WinGet Application install
		# Actually this start log entry applies to Win32Apps too
		if($LogEntryList[$i].Message -like '`[Win32App`]`[V3Processor`] Processing subgraph with app ids:*') {
			if($LogEntryList[$i].Message -Match '^\[Win32App\]\[V3Processor\] Processing subgraph with app ids: (.*)$') {
				$WinGetAppId = $Matches[1]
				$WinGetProcessingStartIndex = $i

				$Thread = $LogEntryList[$i].Thread
				$WinGetInstallationStartedDateTime = $LogEntryList[$i].DateTime

				$WinGetDownloadFirstEntryFound=$false
				$BackupIntent=$null
				
				$WinGetAppBytes = $null
				$WinGetDownloadStartDateTime = $null
				$WinGetDownloadEnddDateTime = $null
				$DownloadDurationDateTimeObject = $null
				$TotalBytesDownloadedMBRounded = $null
				$TotalBytesDownloadedMBExact = $null
				$AppSizeInMB = $null

				$WinGetAppName = Get-AppName $WinGetAppId

				$DetailToolTip = $null

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

							# Extract AppSize
							# WinGet StoreApps show only xx/100
							# WinGet Win32App show downloaded/fullsize
							if($WinGetAppBytes -eq '100') {
								# We are downloading Store App and we don't know app size :(
								$WinGetAppBytes = 'N/A'
								$TotalBytesDownloadedMBRounded = 'N/A'
							} else {
								# We are downloading WinGet Win32 app
								$TotalBytesDownloadedMBExact = $WinGetAppBytes /1MB
								$TotalBytesDownloadedMBRounded = [math]::Round($WinGetAppBytes /1MB, 0)
							}

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

							if($WinGetAppBytes) {
								$LogEntryList[$b].ProcessRunTime = "       Start Download WinGetApp ($WinGetAppName) ($($TotalBytesDownloadedMBRounded)MB) for user $WinGetUserName"
							} else {
								$LogEntryList[$b].ProcessRunTime = "       Start Download WinGetApp ($WinGetAppName) for user $WinGetUserName"
							}

							#Write-Host "DEBUG: Found WinGet first download entry"
							
							# Continue For-loop to next round
							Continue
						}
					}

					$DownloadCompleteString1 = 'State transition processed - Previous State:Download In Progress received event: Download Finished which took state machine into Download Complete state.'
					$DownloadCompleteString2 = 'Processing state transition - Current State:Download Complete With Event: Continue Install.'
					if(($LogEntryList[$b].Message -eq $DownloadCompleteString1) -or ($LogEntryList[$b].Message -eq $DownloadCompleteString2)) {

						$WinGetDownloadEnddDateTime = $LogEntryList[$b].DateTime						

						# Our if clause has or so in normal case we will find 2 log entries for download complete
						# We need to make sure we only process one of those log entries
						if(-not $DownloadDurationDateTimeObject) {
							if($WinGetDownloadStartDateTime -and $WinGetDownloadEnddDateTime) {
								$DownloadDurationDateTimeObject = New-Timespan (Get-Date $WinGetDownloadStartDateTime) (Get-Date $WinGetDownloadEnddDateTime)
							}

							$LogEntryList[$b].ProcessRunTime = "       Download complete ($WinGetAppName) in $($DownloadDurationDateTimeObject.Seconds) seconds"

						}

						#Write-Host "DEBUG: Found WinGet download complete"
						
						# Continue For-loop to next round
						Continue
					}

					
					$WinGetInstallStartString = 'State transition processed - Previous State:Download Complete received event: Continue Install which took state machine into Install In Progress Download Complete state.'
					if(($LogEntryList[$b].Message -eq $WinGetInstallStartString)) {
						#  -and ($LogEntryList[$b].Thread -eq $Thread)
						$WinGetInstallStartdDateTime = $LogEntryList[$b].DateTime						
						$LogEntryList[$b].ProcessRunTime = "       Start WinGetApp Install process ($WinGetAppName)"
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

							$LogEntryList[$i].ProcessRunTime = "Start WinGetApp processing ($WinGetAppName)"
							$LogEntryList[$b].ProcessRunTime = "NoApplicableUpgrade for WinGetApp ($WinGetAppName)"

							Break
						}
			
						
			
						# Calculate install time
						$ProcessStartTime = $LogEntryList[$i].LogEntryDateTime
						$ProcessEndTime = $LogEntryList[$b].LogEntryDateTime
						$ProcessRunTime = New-Timespan $ProcessStartTime $ProcessEndTime
						$ProcessRunTimeTotalSeconds = $ProcessRunTime.TotalSeconds
						
						$WinGetInstallEndDateTime = $LogEntryList[$b].DateTime
						$LogEntryList[$i].ProcessRunTime = "Start WinGetApp processing ($WinGetAppName)"

						$intent = Get-AppIntent $AppId
						if($intent -eq 'Unknown Intent') {
							if($BackupIntent) {
								
								# This case when App is installed from Intune Company Portal
								# Then that App is not shown in Get Policies -list
								# Intent in AvailableInstall then
								$intent = $BackupIntent
							}
						}

						#$detail = "{0,$ColonIndentInTimeLine}{1}  {2}" -f "$Intent", ':', "$WinGetAppName for user $WinGetUserName"
						$detail = "$WinGetAppName for user $WinGetUserName"
											
						if($WinGetInstallResult -eq 'OK') {
							
							RecordStatusToTimeline -date $LogEntryList[$b].DateTime -status "Success" -Type "WinGetApp" -Intent $intent -detail $detail -Seconds $ProcessRunTimeTotalSeconds -Color 'Green' -logEntry "Line $($LogEntryList[$b].Line)"

							$LogEntryList[$b].ProcessRunTime = "Stop WinGetApp Install Success ($WinGetAppName) for user $WinGetUserName ($ProcessRunTimeTotalSeconds)"
							
						} else {
							
							RecordStatusToTimeline -date $LogEntryList[$b].DateTime -status "Failed" -Type "WinGetApp" -Intent $intent -detail $detail -Seconds $ProcessRunTimeTotalSeconds -Color 'Red' -logEntry "Line $($LogEntryList[$b].Line)"
							
							$LogEntryList[$b].ProcessRunTime = "Stop WinGetApp Install Failed ($WinGetAppName) for user $WinGetUserName ($ProcessRunTimeTotalSeconds)"
						}
						
						# Add App Download information to list
						if($DownloadDurationDateTimeObject) {
							$BytesFromPeersMBRounded = 'N/A'
							
							$DownloadDurationSeconds = $DownloadDurationDateTimeObject.Seconds
							
							if($WinGetAppBytes -ne 'N/A') {
								$DownloadMBps = $TotalBytesDownloadedMBExact / $DownloadDurationSeconds
								$DownloadMBps = [math]::Round($DownloadMBps, 1)
								if($DownloadMBps -eq 0) {
									$DownloadMBps = 'N/A'
								}
							} else {
								$DownloadMBps = 'N/A'
							}
							
							$BytesFromPeersPercentage = 'N/A'
							
							# Add download statistics object to list
							$ApplicationDownloadStatistics.add([PSCustomObject]@{
								'AppType' = 'WinGetApp'
								'AppName' = $WinGetAppName
								'DL Sec' = $DownloadDurationSeconds
								'Size (MB)' = $TotalBytesDownloadedMBRounded
								'MB/s' = $DownloadMBps
								'Delivery Optimization %' = $BytesFromPeersPercentage
								})
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
					# Except we will get here if log start entry was for App witch is notApplicable
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

<#
		# region Application Detection check (Win32App and WinGetApp)

		#<![LOG[[Win32App][ActionProcessor] App with id: 334e8b62-281e-419a-b084-c7c806e4975f, targeted intent: RequiredInstall, and enforceability: Enforceable has projected enforcement classification: EnforcementPoint with desired state: Present. Current state is:
		if($LogEntryList[$i].Message -like '`[Win32App`]`[ActionProcessor`] App with id: *, targeted intent: *, and enforceability: * Current state is:') {
			if($LogEntryList[$i].Message -Match '^\[Win32App\]\[ActionProcessor\] App with id: (.*), targeted intent: (.*), and enforceability:.*classification: (.*) with desired state: (.*). Current state is:$') {
				$AppId = $Matches[1]

				# Example values
				# RequiredInstall
				# RequiredUninstall ?
				# NotTargeted (usually this application is superceeded)
				$AppIntent = $Matches[2]
				
				# Example values
				# EnforcementPoint
				# InstalledOrNotInstalled
				$AppEnforcementClassification = $Matches[3]
				
				# Example values
				# Present
				# NotPresent
				$AppDesiredState = $Matches[4]

				$AppDetectionStartIndex = $i

				$Thread = $LogEntryList[$i].Thread
				$DetectEndDateTime = $LogEntryList[$i].DateTime

				$DetectionAppName = Get-AppName $AppId
				
				#Write-Verbose "DEBUG Application Detection: $DetectEndDateTime $DetectionAppName - $AppIntent"

				$AppDetectionResult = $null
				$AppApplicabilityResult = $null
				$AppIntuneCompanyPortalReport = $null
				$AppIntuneCompanyPortalReportJSON = $null

				# Find DetectionCheck status information
				# Now we go down (to end)
				for($b = $i + 1; $b -lt $LogEntryList.Count; $b++ ) {

					#Detection = NotDetected
					#Applicability =  Applicable
					#Reboot = Clean
					#Local start time = 01/01/0001 0.00.00
					#Local deadline time = 01/01/0001 0.00.00

					# Use this to debug log entries
					#Write-Host "DEBUG: $b - $($LogEntryList[$b].Message)"

					if((-not $AppDetectionResult) -and (($LogEntryList[$b].Message -Match '^Detection = (.*)$') -and ($LogEntryList[$b].Thread -eq $Thread))) {
						$AppDetectionResult = $Matches[1]
						
						# Continue For-loop to next round
						Continue
					}
					
					if((-not $AppApplicabilityResult) -and (($LogEntryList[$b].Message -Match '^Applicability =  (.*)$') -and ($LogEntryList[$b].Thread -eq $Thread))) {
						$AppApplicabilityResult = $Matches[1]
						
						# Continue For-loop to next round
						Continue
					}

					# Try to find Intune Company Portal report aboud application detection
					# This may not work during Autopilot enrollment ESP phase
					# But it will work later
					
					# [Win32App][ReportingManager] Sending status to company portal based on report: {"ApplicationId":"1f3f5171-25a2-44da-b3e8-fba81184d1df","ResultantAppState":1,"ReportingImpact":{"DesiredState":3,"Classification":2,"ConflictReason":0,"ImpactingApps":[]},"WriteableToStorage":true,"CanGenerateComplianceState":true,"CanGenerateEnforcementState":true,"IsAppReportable":true,"IsAppAggregatable":true,"AvailableAppEnforcementFlag":0,"DesiredState":2,"DetectionState":1,"DetectionErrorOccurred":false,"DetectionErrorCode":null,"ApplicabilityState":0,"ApplicabilityErrorOccurred":false,"ApplicabilityErrorCode":null,"EnforcementState":1000,"EnforcementErrorCode":null,"TargetingMethod":0,"TargetingType":2,"InstallContext":1,"Intent":3,"InternalVersion":1,"DetectedIdentityVersion":"1.18.1462.0","RemovalReason":null}

					if(($LogEntryList[$b].Message -Match '^\[Win32App\]\[ReportingManager\] Sending status to company portal based on report: (.*)$') -and ($LogEntryList[$b].Thread -eq $Thread)) {
						$AppIntuneCompanyPortalReportJSON = $Matches[1]
						
						# Try to convert from report json
						$AppIntuneCompanyPortalReport = $AppIntuneCompanyPortalReportJSON | ConvertFrom-JSON -ErrorAction SilentlyContinue

						# Check we got IntuneCompanyPortal report for applicationId we are processing
						# If there are Superceded applications or required application
						# Then we do detection checks to other AppId's than the "main application"
						if($AppIntuneCompanyPortalReport.ApplicationId -eq $AppId) {
							# We found "main" application information
						} else {
							# We found some other applicationId's Company Portal report
							# Set variables to $null
							$AppIntuneCompanyPortalReportJSON = $null
							$AppIntuneCompanyPortalReport = $null
						}
						
						# Break For-loop
						Break
					}

					# Check if we reached end of subgraph processing
					# We should not end here
					if((($LogEntryList[$b].Message -eq '[Win32App][V3Processor] Done processing subgraph.') -or `
						 ($LogEntryList[$b].Message -like '`[Win32App`]`[V3Processor`] Done processing * subgraphs.')) -and ($LogEntryList[$b].Thread -eq $Thread)) {
						$AppApplicabilityResult = $Matches[1]
						
						#Write-Verbose "DEBUG: reached end of subgraph processing in App detection check. We should not get here."
						
						# Break For-loop
						Break
					}
					
					# Double check if we reached end of previous subgraph processing
					# We should not end here
					if($LogEntryList[$b].Message -like '`[Win32App`]`[V3Processor`] Processing subgraph with app ids:*') {
						# Previous application processsing ended and we are already processing new application
						# We should never get here
						
						# Break For-loop
						Break
					}
				} # For-loop ends
				
				
				if($AppIntuneCompanyPortalReport.DetectedIdentityVersion) {
						$DetectAppVersion = $AppIntuneCompanyPortalReport.DetectedIdentityVersion
					} else {
						$DetectAppVersion = $null
					}
					
					if($DetectAppVersion) {
						# This should be WinGetApp
						$detail = "{0,$ColonIndentInTimeLine}{1}  {2}" -f "$AppIntent", ':', "$($DetectionAppName) $($DetectAppVersion) - $($AppDetectionResult) - $($AppApplicabilityResult)"
						
					} else {
						# This should be Win32App
						$detail = "{0,$ColonIndentInTimeLine}{1}  {2}" -f "$AppIntent", ':', "$($DetectionAppName) - $($AppDetectionResult) - $($AppApplicabilityResult)"
					}

					RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "AppDetection" -detail $detail -Seconds $null -Color 'Yellow' -logEntry "Line $($LogEntryList[$i].Line)"

					$LogEntryList[$i].ProcessRunTime = "Detection check for app: $DetectionAppName"
					
			}
		}
		# region Application Detection check (Win32App and WinGetApp)
#>
		
		
		# region Application Detection check (Win32App and WinGetApp) nextGen


		if($LogEntryList[$i].Message -like '`[Win32App`]`[ReportingManager`] Sending status to company portal based on report: *') {
			if($LogEntryList[$i].Message -Match '^\[Win32App\]\[ReportingManager\] Sending status to company portal based on report: (.*)$') {
				$AppIntuneCompanyPortalReportJSON = $Matches[1]

				$AppIntuneCompanyPortalReport = $null
				
				$ResultantAppState = $null
				$DesiredState = $null
				$DetectionState = $null
				$InstallContext = $null
				$Intent = $null
				$InternalVersion = $null
				$DetectedIdentityVersion = $null
				$ImpactingApps = $null
				
				$DetectAppVersion = $null
				$AppIntent = $null
				$AppDetectionResult = $null
				$AppApplicabilityResult = $null
				
				$AppDetectionReportIndex = $i

				$Thread = $LogEntryList[$i].Thread
				$DetectEndDateTime = $LogEntryList[$i].DateTime

				Try {
					$AppIntuneCompanyPortalReport = $AppIntuneCompanyPortalReportJSON | ConvertFrom-JSON -ErrorAction SilentlyContinue
				} catch {
					Write-Verbose "Error converting Application Intune Company Portal detection report from JSON"
				}

				# Continue only if we have value in ResultantAppState property
				# This should get only 2 detections per application
				# one detection before app install and second detection after app install
				if($AppIntuneCompanyPortalReport.ResultantAppState) {
					$ApplicationId = $AppIntuneCompanyPortalReport.ApplicationId
					$ResultantAppState = $AppIntuneCompanyPortalReport.ResultantAppState
					$DesiredState = $AppIntuneCompanyPortalReport.DesiredState
					$DetectionState = $AppIntuneCompanyPortalReport.DetectionState
					$InstallContext = $AppIntuneCompanyPortalReport.InstallContext
					$Intent = $AppIntuneCompanyPortalReport.Intent
					$InternalVersion = $AppIntuneCompanyPortalReport.InternalVersion
					$DetectedIdentityVersion = $AppIntuneCompanyPortalReport.DetectedIdentityVersion
					$ImpactingApps = $AppIntuneCompanyPortalReport.ReportingImpact.ImpactingApps

					$DetectionAppName = Get-AppName $ApplicationId

					# Required Install
					# Required Uninstall
					# Available
					$AppIntent = Get-AppIntentNameForNumber $Intent

					# Detected
					# Not Detected
					$AppDetectionResult = Get-AppDetectionResultForNumber $DetectionState

					# Win32App
					# WinGetApp
					$DetectionAppType = Get-AppType $ApplicationId

					if($AppIntuneCompanyPortalReport.DetectedIdentityVersion) {
							$DetectAppVersion = $AppIntuneCompanyPortalReport.DetectedIdentityVersion
					} else {
						$DetectAppVersion = $null
					}
					
					if($DetectAppVersion) {
						# This should be WinGetApp
						#$detail = "{0,$ColonIndentInTimeLine}{1}  {2}" -f "$AppIntent", ':', "$($DetectionAppName) $($DetectAppVersion) - $($AppDetectionResult)"
						#$detail = "$($DetectionAppName) $($DetectAppVersion) - $($AppDetectionResult)"
						$detail = "`t$($DetectionAppName) $($DetectAppVersion)"
						
					} else {
						# This should be Win32App
						#$detail = "{0,$ColonIndentInTimeLine}{1}  {2}" -f "$AppIntent", ':', "$($DetectionAppName) - $($AppDetectionResult)"
						#$detail = "$($DetectionAppName) - $($AppDetectionResult)"
						$detail = "`t$($DetectionAppName)"
					}

					$Color = 'Yellow'

					# Check if there was error in detection check
					if($AppIntuneCompanyPortalReport.DetectionErrorOccurred) {
						# There was error in Detection Check
						$DetectionErrorCode = $AppIntuneCompanyPortalReport.DetectionErrorCode
						$detail = "$($detail). DetectionErrorOccurred: $DetectionErrorCode"
						$Color = 'Red'
					} else {
						# There was no error in Detection Check
					}

					# Check if there was enforcement error (app was not detected after successful install)
					if($AppIntuneCompanyPortalReport.EnforcementErrorCode) {
						# There was error in Detection Enforcement Check -> App was not Detected after successful install
						$EnforcementErrorCode = $AppIntuneCompanyPortalReport.EnforcementErrorCode
						
						if($AppDetectionResult -eq 'Not Detected') {
							$detail = "$($detail). Not Detected after App install ($($EnforcementErrorCode))"
						} else {
							$detail = "$($detail). EnforcementErrorCode: $EnforcementErrorCode"
						}
						
						$Color = 'Red'
					} else {
						# There was no error in Detection Enforcement Check -> App was detected after successfull install
					}

					# Create App Detection Details column ToolTip information
					$Win32AppInformationForToolTip = $null
					if($IdHashtable.ContainsKey($ApplicationId)) {
						$Win32AppInformationForToolTip = $IdHashtable[$ApplicationId]
						if($Win32AppInformationForToolTip) {
							
							# DEMO for conference session
							# DEBUG
							#$Win32AppInformationForToolTip | ConvertTo-Json -Depth 6 | Set-Clipboard
							
							# Not working, empty variable
							#Write-Host "DEBUG: $($Win32AppInformationForToolTip.DetectionRule.value)" -ForegroundColor Red
							
							# And we notice that DetectionRule itself is also Json
							# So we need to convert this from-json first
							#Write-Host "DEBUG: $($Win32AppInformationForToolTip.DetectionRule)" -ForegroundColor Red
							
							$detailToolTip = $null
							$DetectionRule = $null
							$DetectionType = $null
							$DetectionText = $null
							if($Win32AppInformationForToolTip.DetectionRule) {
								
								# DEBUG
								#$Win32AppInformationForToolTip.DetectionRule

								# Import JSON
								# Fix JSON double escapes
								$DetectionRulesObject = $Win32AppInformationForToolTip.DetectionRule -replace '\\\\', '\\' | ConvertFrom-json
								
								#DEBUG
								#$DetectionRulesObject
								#$DetectionRulesObject | ConvertTo-Json -Depth 5 | Set-Clipboard
								
								# Process DetectionText properties
								# First convert DetectionText JSON to Powershell objects
								# Add each DetectionText property to temporary custom Powershell object
								# and then replace DetectionText json with created custom object
								foreach($rule in $DetectionRulesObject) {
									
									$detectionType = $rule.DetectionType
									$DetectionText = $rule.DetectionText -replace '\\\\', '\\' | ConvertFrom-Json
									
									# Temp object to save each property
									$tempObject = New-Object PSCustomObject
									
									foreach($property in $DetectionText.PSObject.Properties) {
										$propertyName = $property.Name
										$propertyValue = $property.Value
										
										$tempObject | Add-Member -MemberType noteProperty -Name $propertyName -Value $propertyValue
									}
									
									$rule.DetectionText = $tempObject
								}								
								
								#DEBUG
								#$DetectionRulesObject
								#$DetectionRulesObject | ConvertTo-Json -Depth 5 | Set-Clipboard

								# Convert DetectionText to human readable values
								$DetectionRulesObject = Convert-AppDetectionValuesToHumanReadable $DetectionRulesObject

								
								# Show information as list
								$DetectionText = $DetectionRulesObject | Select-Object -Property DetectionTypeAsText -ExpandProperty DetectionText | Select-Object -Property DetectionTypeAsText,* -ErrorAction SilentlyContinue | Format-List -Property * | Out-String
								
								$detailToolTip = "Detection rules:`n$DetectionText" | Out-String
							} else {
								# There was no DetectionRule property
								$detailToolTip = $null
							}
						} else {
							# There was no Application information available
						}
					}
					
					RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status $AppDetectionResult -Type $DetectionAppType -Intent $AppIntent -detail $detail -detailToolTip $detailToolTip -Seconds $null -Color $Color -logEntry "Line $($LogEntryList[$i].Line)"	

					$LogEntryList[$i].ProcessRunTime = "$AppDetectionResult app: $DetectionAppName"
				}
			} # end log entry match



<#				
				#Write-Verbose "DEBUG Application Detection: $DetectEndDateTime $DetectionAppName - $AppIntent"

				$AppDetectionResult = $null
				$AppApplicabilityResult = $null
				$AppIntuneCompanyPortalReport = $null
				$AppIntuneCompanyPortalReportJSON = $null

				# Find DetectionCheck status information
				# Now we go down (to end)
				for($b = $i + 1; $b -lt $LogEntryList.Count; $b++ ) {

					#Detection = NotDetected
					#Applicability =  Applicable
					#Reboot = Clean
					#Local start time = 01/01/0001 0.00.00
					#Local deadline time = 01/01/0001 0.00.00

					# Use this to debug log entries
					#Write-Host "DEBUG: $b - $($LogEntryList[$b].Message)"

					if((-not $AppDetectionResult) -and (($LogEntryList[$b].Message -Match '^Detection = (.*)$') -and ($LogEntryList[$b].Thread -eq $Thread))) {
						$AppDetectionResult = $Matches[1]
						
						# Continue For-loop to next round
						Continue
					}
					
					if((-not $AppApplicabilityResult) -and (($LogEntryList[$b].Message -Match '^Applicability =  (.*)$') -and ($LogEntryList[$b].Thread -eq $Thread))) {
						$AppApplicabilityResult = $Matches[1]
						
						# Continue For-loop to next round
						Continue
					}

					# Try to find Intune Company Portal report aboud application detection
					# This may not work during Autopilot enrollment ESP phase
					# But it will work later
					
					# [Win32App][ReportingManager] Sending status to company portal based on report: {"ApplicationId":"1f3f5171-25a2-44da-b3e8-fba81184d1df","ResultantAppState":1,"ReportingImpact":{"DesiredState":3,"Classification":2,"ConflictReason":0,"ImpactingApps":[]},"WriteableToStorage":true,"CanGenerateComplianceState":true,"CanGenerateEnforcementState":true,"IsAppReportable":true,"IsAppAggregatable":true,"AvailableAppEnforcementFlag":0,"DesiredState":2,"DetectionState":1,"DetectionErrorOccurred":false,"DetectionErrorCode":null,"ApplicabilityState":0,"ApplicabilityErrorOccurred":false,"ApplicabilityErrorCode":null,"EnforcementState":1000,"EnforcementErrorCode":null,"TargetingMethod":0,"TargetingType":2,"InstallContext":1,"Intent":3,"InternalVersion":1,"DetectedIdentityVersion":"1.18.1462.0","RemovalReason":null}

					if(($LogEntryList[$b].Message -Match '^\[Win32App\]\[ReportingManager\] Sending status to company portal based on report: (.*)$') -and ($LogEntryList[$b].Thread -eq $Thread)) {
						$AppIntuneCompanyPortalReportJSON = $Matches[1]
						
						# Try to convert from report json
						$AppIntuneCompanyPortalReport = $AppIntuneCompanyPortalReportJSON | ConvertFrom-JSON -ErrorAction SilentlyContinue

						# Check we got IntuneCompanyPortal report for applicationId we are processing
						# If there are Superceded applications or required application
						# Then we do detection checks to other AppId's than the "main application"
						if($AppIntuneCompanyPortalReport.ApplicationId -eq $AppId) {
							# We found "main" application information
						} else {
							# We found some other applicationId's Company Portal report
							# Set variables to $null
							$AppIntuneCompanyPortalReportJSON = $null
							$AppIntuneCompanyPortalReport = $null
						}
						
						# Break For-loop
						Break
					}

					# Check if we reached end of subgraph processing
					# We should not end here
					if((($LogEntryList[$b].Message -eq '[Win32App][V3Processor] Done processing subgraph.') -or `
						 ($LogEntryList[$b].Message -like '`[Win32App`]`[V3Processor`] Done processing * subgraphs.')) -and ($LogEntryList[$b].Thread -eq $Thread)) {
						$AppApplicabilityResult = $Matches[1]
						
						#Write-Verbose "DEBUG: reached end of subgraph processing in App detection check. We should not get here."
						
						# Break For-loop
						Break
					}
					
					# Double check if we reached end of previous subgraph processing
					# We should not end here
					if($LogEntryList[$b].Message -like '`[Win32App`]`[V3Processor`] Processing subgraph with app ids:*') {
						# Previous application processsing ended and we are already processing new application
						# We should never get here
						
						# Break For-loop
						Break
					}
				} # For-loop ends
				
				
				if($AppIntuneCompanyPortalReport.DetectedIdentityVersion) {
						$DetectAppVersion = $AppIntuneCompanyPortalReport.DetectedIdentityVersion
					} else {
						$DetectAppVersion = $null
					}
					
					if($DetectAppVersion) {
						# This should be WinGetApp
						$detail = "{0,$ColonIndentInTimeLine}{1}  {2}" -f "$AppIntent", ':', "$($DetectionAppName) $($DetectAppVersion) - $($AppDetectionResult) - $($AppApplicabilityResult)"
						
					} else {
						# This should be Win32App
						$detail = "{0,$ColonIndentInTimeLine}{1}  {2}" -f "$AppIntent", ':', "$($DetectionAppName) - $($AppDetectionResult) - $($AppApplicabilityResult)"
					}

					RecordStatusToTimeline -date $LogEntryList[$i].DateTime -status "AppDetection" -detail $detail -Seconds $null -Color 'Yellow' -logEntry "Line $($LogEntryList[$i].Line)"

					$LogEntryList[$i].ProcessRunTime = "Detection check for app: $DetectionAppName"
					
			}
#>			
			
		}
		# region Application Detection check (Win32App and WinGetApp) nextGen
		
		
		
		
		# Region ends here
		
		
		$i++
	}
	# endregion analyze log entries


	# Application download statistics
    Write-Host ""
    Write-Host "Application download statistics" -ForegroundColor Magenta
	$ApplicationDownloadStatistics | Format-Table

	# Filtered Assignments
    Write-Host ""
    Write-Host "Application Assignment Filters applied" -ForegroundColor Magenta
	$ApplicationAssignmentFilterApplied | Select-Object -Property AppType, Applicable, AppName, AssignmentFilterNames | Format-Table

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
            Label = "Type"
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
                "$e[${color}m$($_.Type)$e[0m"
            }
        },
		@{
            Label = "Intent"
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
                "$e[${color}m$($_.Intent)$e[0m"
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

		"Get-IntuneManagementExtensionDiagnostics.ps1 $ScriptVersion report" | Out-File -FilePath $ExportTextFileName -Force
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

			if($IdHashtable[$HashtableKey].name) {			
				Foreach($LogEntry in $LogEntryList) {
						$Logentry.Message = ($Logentry.Message).Replace($HashtableKey, "'      $($IdHashtable[$HashtableKey].name)      '")
				}
				
				Foreach($LogEntry in $LogEntryList | Where-Object ProcessRuntime -ne $null) {
					$Logentry.ProcessRuntime = ($Logentry.ProcessRuntime).Replace($HashtableKey, "'$($IdHashtable[$HashtableKey].name)'")
				}
				
				# Continue to next HashtableKey
				# So we skip next if clause in case object has both name and displayName
				Continue
			}

			if($IdHashtable[$HashtableKey].displayName) {
				Foreach($LogEntry in $LogEntryList) {
						$Logentry.Message = ($Logentry.Message).Replace($HashtableKey, "'      $($IdHashtable[$HashtableKey].displayName)      '")
				}
				
				Foreach($LogEntry in $LogEntryList | Where-Object ProcessRuntime -ne $null) {
					$Logentry.ProcessRuntime = ($Logentry.ProcessRuntime).Replace($HashtableKey, "'$($IdHashtable[$HashtableKey].displayName)'")
				}
			}
		}
	}



	if($ExportHTML) {

		########################################################################################################################
		# Create HTML report

	$head = @'
	<style>
		body {
			background-color: #FFFFFF;
			font-family: Arial, sans-serif;
		}

		  header {
			background-color: #444;
			color: white;
			padding: 10px;
			display: flex;
			align-items: center;
		  }

		  header h1 {
			margin: 0;
			font-size: 24px;
			margin-right: 20px;
		  }

		  header .additional-info {
			display: flex;
			flex-direction: column;
			align-items: flex-start;
			justify-content: center;
		  }

		  header .additional-info p {
			margin: 0;
			line-height: 1.2;
		  }

		  header .author-info {
			display: flex;
			flex-direction: column;
			align-items: flex-end;
			justify-content: center;
			margin-left: auto;
		  }

		  header .author-info p {
			margin: 0;
			line-height: 1.2;
		  }
		  
		  header .author-info a {
			color: white;
			text-decoration: none;
		  }

		  header .author-info a:hover {
			text-decoration: underline;
		  }

		table {
			border-collapse: collapse;
			width: 100%;
			text-align: left;
		}

		table, table#TopTable {
			border: 2px solid #1C6EA4;
			/* background-color: #f7f7f4; */
			background-color: #012456; /* dark background */
			color: #f0f0f0;  /* light color for the text */
		}

		table td, table th {
			/* border: 2px solid #AAAAAA; */
			border: 1px solid #1a2f6a;
			padding: 2px;
		}

		table td {
			font-size: 15px;
			white-space: pre-wrap; /* or 'pre' if you don't want any wrapping */
		}

		table th {
			font-size: 18px;
			font-weight: bold;
			color: #FFFFFF;
			background: #1C6EA4;
			background: -moz-linear-gradient(top, #5592bb 0%, #327cad 66%, #1C6EA4 100%);
			background: -webkit-linear-gradient(top, #5592bb 0%, #327cad 66%, #1C6EA4 100%);
			background: linear-gradient(to bottom, #5592bb 0%, #327cad 66%, #1C6EA4 100%);
		}

		table#TopTable td, table#TopTable th {
			vertical-align: top;
			text-align: center;
		}

		table thead th:first-child {
			border-left: none;
		}

		table thead th span {
			font-size: 14px;
			margin-left: 4px;
			opacity: 0.7;
		}

		table tfoot {
			font-size: 16px;
			font-weight: bold;
			color: #FFFFFF;
			background: #D0E4F5;
			background: -moz-linear-gradient(top, #dcebf7 0%, #d4e6f6 66%, #D0E4F5 100%);
			background: -webkit-linear-gradient(top, #dcebf7 0%, #d4e6f6 66%, #D0E4F5 100%);
			background: linear-gradient(to bottom, #dcebf7 0%, #d4e6f6 66%, #D0E4F5 100%);
			border-top: 2px solid #444444;
		}

		table tfoot .links {
			text-align: right;
		}

		table tfoot .links a {
			display: inline-block;
			background: #1C6EA4;
			color: #FFFFFF;
			padding: 2px 8px;
			border-radius: 5px;
		}

		/* Every other table row to different background color
		table tbody tr:nth-child(even) {
		  background-color: #D0E4F5;
		}
		*/

		/* Set table row background color */
		table tbody tr {
		  /* background-color: #f7f7f4; */
		  background-color: #012456;
		}

		select {
		  font-family: "Courier New", monospace;
		}

	  /* Set DownloadTable settings */
	  #ApplicationDownloadStatistics {
		width: auto;
		max-width: 1000px;
		margin: 0 auto; /* Centers the table if it doesn't span the entire width of its container */
		
		/* Float -> align table to left
		   This setting needs to be cleared after table with html syntax
		   <div style="clear:both;"></div>
		   because otherwise page bottom element will float on right of table
		 */
		float: left;
	  }

		
	  footer {
		background-color: #444;
		color: white;
		padding: 10px;
		display: flex;
		align-items: center;
		justify-content: center;
	  }


	  footer .creator-info {
		display: flex;
		flex-direction: row;
		align-items: center;
		margin-right: 20px;
	  }

	  footer .creator-info p {
		line-height: 1.2;
		margin: 0;
	  }

	  footer .creator-info p.author-text {
		margin-right: 20px; /* Add margin-right rule here */
	  }

	  .profile-container {
		position: relative;
		width: 50px;
		height: 50px;
		border-radius: 50%;
		overflow: hidden;
		margin-right: 10px;
	  }

	  .profile-container img {
		width: 100%;
		height: 100%;
		object-fit: cover;
		transition: opacity 0.3s;
	  }

	  .profile-container img.black-profile {
		position: absolute;
		top: 0;
		left: 0;
		z-index: 1;
	  }

	  .profile-container:hover img.black-profile {
		opacity: 0;
	  }

	  footer .company-logo {
		width: 100px;
		height: auto;
		margin: 0 20px;
	  }

	  footer a {
		color: white;
		text-decoration: none;
	  }

	  footer a:hover {
		text-decoration: underline;
	  }

	  
		.filter-row {
		  display: flex;
		  align-items: center;
		}

		.control-group {
		  display: flex;
		  flex-direction: column;
		  align-items: flex-start;
		  margin-right: 16px;
		}

		.control-group label {
		  font-weight: bold;
		}

		/* Tooltip container */
		.tooltip {
		  position: relative;
		  display: inline-block;
		  cursor: pointer;
		  /* text-decoration: none; */ /* Remove underline from hyperlink */
		  color: inherit; /* Make the hyperlink have the same color as the text */
		}

		/* Tooltip text */
		.tooltip .tooltiptext {
		  visibility: hidden;
		  
		  /* make it up to the viewport width minus some padding 
		     bigger value is used so reaaaaaally wide tooltips would not overflow from left
			 which it still will do depending on window size
		  */
		  max-width: calc(100vw - 160px); 
		  
		  /* width: 120px; */
		  background-color: #555;
		  color: #fff;
		  text-align: left;
		  border-radius: 6px;
		  position: absolute;
		  z-index: 1;
		  bottom: 125%;

		  /* ensures the tooltip is centered with respect to the hovered element */
		  /* These are important settings! */
		  transform: translateX(-50%);
		  left: 50%;

		  /* margin-left: -60px; */
		  opacity: 0;
		  transition: opacity 1s;
		  /* white-space: pre; */
		  padding: 10px; /* Change this value to suit your needs */
		  
		  /* monospaced font so output is more readable */
		  font-family: 'Courier New', monospace;
		  /* word-wrap: break-word; */ /* break long words to fit within the tooltip */
		  
		  overflow-x: auto; /* This adds horizontal scrolling if content exceeds max-width */
		  
		  /* white-space: nowrap; */ /* This ensures content stays in a single line */
		  
		  /* preserve whitespaces */
		  white-space: pre;
		  
		  
		}

		/* Show tooltip text when hovering */
		.tooltip:hover .tooltiptext {
		  visibility: visible;
		  opacity: 1;
		}
	</style>
'@

		############################################################
		# Create HTML report

		Write-Host "Create HTML report information"

		######################
		#Write-Host "Create Observed	Timeline HTML fragment."

		try {

			$PreContent = @"
				<div class="filter-row">
					<div class="control-group">
					  <label><input type="checkbox" class="filterCheckbox" value="Win32App" onclick="toggleCheckboxes(this)"> Win32App</label>
					  <label><input type="checkbox" class="filterCheckbox" value="WinGetApp" onclick="toggleCheckboxes(this)"> WinGetApp</label>
					  <label><input type="checkbox" class="filterCheckbox" value="Powershell script" onclick="toggleCheckboxes(this)"> Powershell script</label>
					  <label><input type="checkbox" class="filterCheckbox" value="Remediation" onclick="toggleCheckboxes(this)"> Remediation</label>
					</div>
					<!-- Dropdown 1 -->
					<div class="control-group">
						<label for="dropdown1">Status</label>
						<select id="filterDropdown1" multiple>
						  <option value="all" selected>All</option>
						</select>
					</div>
					<!-- Dropdown 2 -->
					<div class="control-group">
						<label for="dropdown1">Type</label>
						<select id="filterDropdown2">
						  <option value="all">All</option>
						</select>
					</div>
					<!-- Dropdown 3 -->
					<div class="control-group">
					<label for="dropdown1">Intent</label>
					<select id="filterDropdown3">
						  <option value="all">All</option>
						</select>
					</div>
				</div>
				<br>
				<div>
					<input type="text" id="searchInput" placeholder="Search...">
					<button id="clearSearch" onclick="clearSearch()">X</button>
					<button id="resetFilter" onclick="resetFilters()">Reset filters</button>
				</div>
"@
			
			$observedTimelineHTML = $observedTimeline | ConvertTo-Html -As Table -Fragment -PreContent $PreContent

			# Fix &lt; &quot; etc...
			$observedTimelineHTML = Fix-HTMLSyntax $observedTimelineHTML

			# Fix column names
			#$AllAppsByDisplayNameHTML = Fix-HTMLColumns $AllAppsByDisplayNameHTML

			# Add TableId
			$TableId = 'ObservedTimeline'
			$observedTimelineHTML = $observedTimelineHTML.Replace('<table>',"<table id=`"$TableId`">")

			# Convert HTML Array to String which is requirement for HTTP PostContent
			$observedTimelineHTML = $observedTimelineHTML | Out-String


			# Debug- save $html1 to file
			#$observedTimelineHTML | Out-File "$PSScriptRoot\observedTimelineHTML.html"

		}
		catch {
			Write-Error "$($_.Exception.GetType().FullName)"
			Write-Error "$($_.Exception.Message)"
			Write-Error "Error creating HTML fragment information"
			Write-Host "Script will exit..."
			Pause
			Exit 1        
		}
		#############################
		# Create html
		Write-Host "Creating HTML report..."

		try {

			$ReportRunDateTime = (Get-Date).ToString("yyyyMMddHHmm")
			$ReportRunDateTimeHumanReadable = (Get-Date).ToString("yyyy-MM-dd HH:mm")
			$ReportRunDateFileName = (Get-Date).ToString("yyyyMMddHHmm")

			if($ExportHTMLReportPath) {
				if(Test-Path $ExportHTMLReportPath) {
					$ReportSavePath = $ExportHTMLReportPath
				} else {
					Write-Host "Warning: Parameter -ExportHTMLReportPath specified but destination directory does not exists ($ExportHTMLReportPath) !" -ForegroundColor Yellow
					Write-Host "Warning: Defaulting to running directory ($PSScriptRoot)!" -ForegroundColor Yellow
					$ReportSavePath = $PSScriptRoot
				}
			} else {
				$ReportSavePath = $PSScriptRoot
			}
			
			
			if($ComputerNameForReport -eq 'N/A') {
				$HTMLFileName = "$($ReportRunDateFileName)_Intune_Logs_Report.html"	
			} else {
				$HTMLFileName = "$($ReportRunDateFileName)_$($ComputerNameForReport)_Intune_Logs_Report.html"	
			}
			
			
			$PreContent = @"
		<header>
		  <h1>Get-IntuneManagementExtensionDiagnostics ver $ScriptVersion</h1>
		  <div class="additional-info">
			<p><strong>Report run:</strong> $ReportRunDateTimeHumanReadable</p>
			<p><strong>Computer Name:</strong> $ComputerNameForReport</p>
			<!-- <p><strong>Tenant name:</strong> $TenantDisplayName</p> -->
			<!-- <p><strong>Tenant id:</strong> $($ConnectMSGraph.TenantId)</p> -->
		  </div>
		  <div class="author-info">
			<p><a href="https://github.com/petripaavola/Get-IntuneManagementExtensionDiagnostics" target="_blank"><strong>Download Report tool from GitHub</strong></a><br>Author: Petri Paavola - Microsoft MVP</p>
		  </div>
		</header>
		<br>
"@

			$JavascriptPostContent = @'
		<div style="clear:both;"></div>
		<p><br></p>
		<footer>
			<div class="creator-info">
			<p class="author-text">Author:</p>
			  <div class="profile-container">
				<img src="data:image/png;base64,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" alt="Profile Picture Petri Paavola">
				<img class="black-profile" src="data:image/png;base64,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" alt="Black Profile Picture Petri Paavola">
			  </div>
			  <p><strong>Petri Paavola</strong><br>
				<a href="mailto:Petri.Paavola@yodamiitti.fi">Petri.Paavola@yodamiitti.fi</a><br>
				Senior Modern Management Principal
			  </p>
			</div>
			<img class="company-logo" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGQAAAAoCAYAAAAIeF9DAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAALiIAAC4iAari3ZIAAAAZdEVYdFNvZnR3YXJlAEFkb2JlIEltYWdlUmVhZHlxyWU8AAAS4klEQVRoQ+1bB3RVZbrdF0JIIRXSwITeQi8SupCAMogKjCDiQgZ9wltvRBxELAiCNGdQUFEERh31DYo61lEUEWQQmCC9SpUkEBJIJQlpQO7b+7/nxAQJ3JTnems99uLm3vOfc/7y7a/+5+DA1K1O3MD/GdSyvmsWJeT4col1cAOVQc0TcqkEtWo7EB5YFyi+bDXegLuoWUJIBoouo+D5GKTM7o7adWvfIKWSqDlCbDKW9oWnh6vbSy/0goeXx29PipMuU5/fChrrItdfA+usGUJEBidU9GpfeNUp3+XFRT3h6U1SSFaNIPcikFMMFFyyGgj9Vlsez0k4+TzW57cgRWNw3FtaB2L0zaEuUtRWxbGrn2VZZFykZXgwdlQEryfiUXSBQpIbqypIxu7Z3eDPPj4/kImpH5wwzfNHNsV93UJwMqMQA2duh/P9QabdMXETUI/K4Kh4XtVG/kWsHN8aD/UMM4eO8d8DHg7UCqqLEpPc8CPr0Rw8qay1rj2X6lmIyODnEi3jWmQIhX/uCe96dapnKczcOjf0RbP6XpjSP4KWwb7Y36Te4WhMAXRu5GsyvOPpBcgtO44EI81V5qffmre+L5bRZp2X4GzN1reuV7uuv7JNgtYx5yAyEjIL4Ri5Fi3bBML55gDc07k+QAVsw/muvL8V5gxv7CJD414DVSfEkOHEZZJR+zqs28hnsPf1qyYpRBbdUS1pXIAnQEWo7+OBnMJLuCQh8Xe3xfsR+dxOwJvWWFQCB+d3R5cGCAuoCw/Gt05R9YxAh7Sn0HQPhTqofTC6NvWj67PcHtuCqEC6LyrEy7VejhvCMe/o3AAecs089grzMXPK4O+g6GDc0tzfHIdqnb4eaMl7AxhHU9nvGJFkk1sBquaypElk2kkyqgK/p7chTwuvrPvKLjLat/FEDgZw4YOWH0QK+zk4vTM2Hj9vLCSILqPwvTjUpeAd929AD7qybVM6mNsTs4pw99tHsf1PHXA2txhhfp7weXKbURQbB1Ly0WF6PEYNvgkfUrNtOO5bj0dHN8eSu5pYLUD0X/bgvftacFwSfBXM+voUDqRegI9nbQxuFYAF65NxNDWfSlSxHVTeQiyzrioZQu6CGARIu6toKZtIiDCoVSDiWgYgk9qZyvhie810uopCzZPtNhmOKVvQ8YW98JUfJy4UlyB89g6sGNXMHEfw95xvT6F9hA9atwvGk7ENTbtj+Fr4U4FAaxEZR9MKGCc2mHO7p3ZElxnbze8dp/LMGKPeOWKOH/zgOOauO4VPt6fh0NkCPPp5Ao6euTYZQuUIMWTwHwN4dZE9vwcCA0lKYRVJ+TkHw9oG4VZmNx/vy6AWll9KidyOSCc+3Jth/HkOLayOxdq0LxJx9mg2YuS+iFRa2Md7M83vgS388cXBLPP74Ip+yCXZjSMZn4gP96RTDk58eyTbWKGJY8RFxYbMIqMMQloev6VwXrWxOykX2Wq/IgO9GtwnRGQQzqV9zHdNIGteDwQHs6KvJCneXNg/KGRpc1xLf2ZcWS7hlIHxw1bbZcUJwfoSLkuAdcq6TEep8tajK53z9hHM+iYJ0YwRUsCsfNccHVbGZhPL2wxMGGWbhxVP6UP4l791vTp2M866R4jI4GDOV2qODBsZc29GRCiDJoOyu6hLQj5j2it4U6hfHcwkIeUXXFuCoMYK93ZtYAJsWKg3EysXK0auXNPWhFxzHNk6AGMYwAVpfzR/z339EF75IcW05ajOIcYwoEvrB7YIQJ4swCZbYJ92EqWsz5AlS60Erk+IyOBizy/sgQmrj5vP61tSrZMujHvvmPGZc9aewk760tHvHsUb8Wetsy7Itz704QkkMbDet+oYJn10whzPWJOEM892R3MFRjdJCaU/P3U42zoiOGY4A7SfdgWIEN86rgKVgot9/aBpc77cB/FT2rNscFm6rEznJ6x21TJJM7vhidhG2JaYi33bz+Ff/9UOzs+H4JF+ETioQJxdjKmMA80beMH51gBzz+AVh9iRa0xlerKEDcfOm+NXRzbD8yOaVjpOXjvLsi1jSW+cpLY1Y9BSWiltcy7ubS7JuHARDSaxAKMQIuhCDj3RGUEPbERgM39kMU4Iu5Pz0HVaPGKYry++swn6MICiLgUiN+DFb6bPuSv6o/2ivUhMucA21yJ/BdYAUaxBcrjI7PPFCOfvS1TJdAorhPHIkwJJZg0SwfZadBHJrA3A4O3j74kBjAtrSaIWK+09Q40v4jlXPVKCuI71cY4V934RTcLrMCYpk8ug79/1E+OJ0lilvaE+6M54spZWVKL4Ucdh5lTAdDbtPDNH9ufFdDuWycYGxqVCpt3uuiuhYgtR8WORIWjhqjRDwrxLXYHwjRZAMupxUqpHAqUxFE4200cb645Sa+iXJ/QIcfXDfof2CIXz77GI6x5qzg1YdhAJz3RFQy643LZIWVDgSemFrgDJVDKVQTpdroQWnMb0NzmLBLA9he2GDPluziefMWoNY45iiarnk+cKXGRIToozHH893d7+pDzAn4Ln/C6SpHX7MrGLyYMhQz7Opw7HKcLXTCJK5KpoYRK25pSWQZnwPgVukbBmVzoKCytHhnB1QkQGfbNNRim4GFMNc+A10hpizU/ZCKZv9iNZJrMh+jJt1LX7zlDbie9ECBdwN7WwUIURr7Nj8Fv3NDfF0s7TFAaR/Gw3RIZXQIr6lwsQqfqtxcqKJVz1JwIEfdu/BQlKA+oeE2T5UR+6z7TxGlqpiC0FDxvT4pvwA15moNPql6SXEzT7jG0X5JqLwD57MvvzUp2l/iuBMrO2IDIoXOfiXlZDGVCo2roQRITwBTXrd20CkW0yJdckh7Pq1YKMZRDrLDdQn5ZUrIUTJugSqgeEYG9qoYWkWd0QJUFcSQqtwLm8P6KlFOqHLmPVg22wbGxLV6amxUtzbSHYx/yEUPPbmfs4HgU3eUBDLBxp+Xj2e3kZU3n91q0ii9fdxXWM0Fok6LJ92t/qW+C1b49t4Zqv7mU/zw+LYlzjmnRsf9xAeULooxsqaL14hWXY4ARiGjP4UrhfHnJlOXknczG8QzCKtFALOtaEVVEb0KQHadvABrUrgcH9q0NZiGO1rck+3Ne1OWcjkUG2B7WsVOssaNvk8YEs2qQ4/IztGoJsCUIC4rlGSqPl0nTM78j6rgdlozs1MAI2QqS1LGX29Ae5SwqvdpAnjqfRxZHgulRGL350j9ak2GD6kjzLCpeK1EiZlBlLdQfjhyxH13GMPLotwxcVxVeuTSgjo4rwKwtRH9dCMUm7vUN9JJKIExS0MJQFWoml+UJzxhP5bmUch+mv1elICcMGBbKbZA1bsAspvCaGqeScIVHWyWuArkYVsYpBLTSa3xrD5P5MEHKYST03JBIJ2gphZnSYKfXMwTdhBOPVgzGhmMikoo/ulUBpFf5KKCijB3j+lc0paM3K/9M/tMZPTEx60O0WUBlySJI2B2NbBQBMJJxLmPpzTXue7oL5t0fhq0c7mF3orswSV45rCedLVGYmPioUlRaPuyUCn01ojbR5NyNUT1G1OXkNlCeE/jGFQnY8ttVq+DU0yeHt6S+Jed+dNoL3UQppm6+FvrSSfAbWd5hCipFRZS2EmtKTi988twdOvdqX6ahre6MsGs/diR+ZyVxZ3Spd/YRBtRkF9BjdznzOQd5vaN9wLPnXGTz44l7sZ0IR1jIQrRnbJr52EJ9uSsHzG5LNtVvoYk08YT+r92SgHd3t/d1C8Oa2czhyMgdD/7KH1yWb3doiS8nO0wJtD6BMTGvuPD0ekz856ar0aUGZzDYnLt6H4X87jCdvjaScLiOPVvbiHU0weM5O3PPuMTwzqNGvLP5KlF+tILOjOTqm/ttqKI9CTtKYPvHujjREt6XGGZQnZITcFiX1zvY0gKbdgPGjFNSecPr0Pk39cJO05gpopzZJWZoytiugjbrn1p3GtFsaottNvohn3aB41ISu6cg5Wiz9tqwyOtwbTefvgvO/YxHVzM9YkdnHsrMJEvJX1kr3022F8Z7C1ALcSSvaNqc7GjJNLrr4y3pEuIxKf0yFzy6SX+uHkVxjptwl+05kZidXnkkXpm0ceTkZQwDd1crJ7TGxV6h5hlMu2bgKrn5WpJBhx59+sRQNoAlJU0I4sDouoakOi3ZZi0zfzrIEQwi1IeVsAYapurVgruC1dsV8JRpRm06rELsKGUI91i9pTBL0DEKJhS/nqiJwixXLlJJr7F2nLiCBblFr+GR8a+RTyYy3sDWUGVU809rpsQ1d+1ac+4uskWKoDHuZHapPeULJT9slnsrA2ObL8WNZ+b/MGPQOC8UWcs9cSxcVtnRpI5lJavNTXkPbK9rbmsiKfwytZ72SnLKZ3FVQMV0ihQPZpHgqVbQ1hegpt0UfqfhxNTQN5kQVzEjg6E6/uCubtDLclUI7rmfOVUyGYshe7ZjS2hZtPINlW1IpMIcp8vZSuNLeoyv7Gzd5nml0wgu9cIS+fuyqY/gnXdWfhzVGXBfORaRI7Sk0bUy+QXcl7dYuRCpjhIg9klZgdpFzGZxnrz2Fzx9oizf/oy2+PXIeG3ams4IPx6rpnfAB3Z6y0h9Yr2xjmdCVmdx3W1NpMcWGlF5L9yPp5d44uKgnIrSZaulDRbj+8xBNnrQ5X6r5fayyCH92B86yyq6wSrchFyHCtM3CBRstkX+XO1LlLI2RFqoGUAYkWAUc6OdN/+Y+CxS6uVeuTGtV+izl00f3CLpc1xl3w/7Vh/rWNYIehNljqUm7GUqhpdSajykJCM3bvqcCuPeAymgUp8Is5n8DIbO2Iz2D/v96ZAhaqGashUnw0vSKIKJEkg8Fcx3fXQql0yLHJsNdyOTNWG6s4Rpwb5bSKI7neGSL1VBzqD9TZDAguknGBGZTs+9qjH5KQyWAq/k+gWREsaa68FoftIlgQXiddNOAZPwnkwV/CdX2ze5Ac6BrWz+1o8uSqgE31YawzNwxueZICXrmR2RqX8wunK4HErL87uZYvTsdq1kb3HlziEsAshp9ZBFljmewBmm1YA8OMxMzlb3OKZkw2szfdkGp661zpQandrXJNeq3+pb1qM1ODOxj9U0Cm6oIrSBZcRfuEyKIFHmJyZuthqojcMaPyGbgc5sMgWNnUYiHD2ThthWHMJrZWwyzvHG9w/DU0CgEMn3d8FhHjNUrOfTXd7YLwpvaK6N7eziuEb5hEeepeEGf/hYLwNUT25rMaGyvMGx6vJN5XqKXEXIkZOKzP7bDolG8n0lDR9YbI1ivyAp66EUGxqNJtKbvp3VEJ+1ekDCzcVpNVI4QwZDigOPhqpOilxzOUxCVfsmBcEpbWZxNpBC/Y+E4rlsD3MbibuGXieYJZCwLwyeYyrYO88aXh7Iw4+sk3NahPlrQfQ1ZegA/UMizhzcxD6HGv38cQXRn834Xif4Ld+Mc0/in4hrSYkrw86zumPTRz7hA63lqZDO0C/fBJI4Zt2Sf6zk92/VsZ+C8Xa5juc8aQOUJEUQKtc7xx82laay7qPdUFd84Eegp9Hj10PJ+iPDzxNtfJcKPxzO/PoWmTDc/2ptutH/6l0m0nvrIppDOkHil3WFMaxeMaIL2JEAF7Ru0HG3BZzGzU4oeR+13ns03hS9pN9sqZ+nqZrPve5kqF9MV/e1HpsdUhhNKQDiXbLqz9x+h1SnBqAHrEKpGiKBMhFlObVqKqV7dgF65uUAtrBIZAofU/lA0a6NRyw/R1DyZDDmM29YU9DhX/tyL89IWTy3pDbMlvcSnV3Ce/mcSfB/bip8T8hi/trseLzD70utCC2+PwqC+ESwgqensy+xGq2P2UcCqXaHFPLdnf1JCT1rgst83w71zduC0XG9ls7IKwBGqAUNKLXiQFPOS2jWgV0kLtA9UVTIsBPvUMUWcnV6qclYBlsjCbGjbQHRnFf3WmOZ4aVOK2a6RBb3wfTL+Mb4VujIOaNdgIDW+JV2YdokjQrxxe/8Is2mpR8OyQAlXr+6MvS0Sax6KNntgwRzPfm7vz4xQROux7UBeE6ldX2qEtkmqi+q/2ysoA6FGFr3Sp/TN97KoOz0exQqU1SRDY8QxiK9nbDBuk+Mq2GrrPF/ZEIX0zK2RWLUzDSeTL6B9Uz8kMIvLo9vqxECsIL9sy1mzLfJQz1D8Nf4cktMLMXVQI/MS3cebUxHTPhjbEvNkFphCYf9EYr5lf1GR9UzQPsPrB7QNwsa9GbiV5Gt7RBuO24/loD/ntulotktRq4iaIUQQKXQn+u8IZd+Ar/N4PC5JWKpaawJKOcv0b8aVD5eLkd/SeR1LKKo91C53outkxbpXK9axNF7n9JCMX+acrtH9gtp1Xtfxp7lGxxpD/asPG7pG91aDDKHmCBEsUvTfEmQpHtP+jcs8rjEy/h+gZgkRSEotapr8caqykRtkVArVs6+rgZahNzL0RsgNMiqPmidEkJ91dzPvBsoA+B+htJLVXhyOiAAAAABJRU5ErkJggg==" alt="Microsoft MVP">
			<p style="margin: 0;">
			  <a href="https://github.com/petripaavola/Get-IntuneManagementExtensionDiagnostics" target="_blank"><strong>Download Get-IntuneManagementExtensionDiagnostics from GitHub</strong<</a>
			</p>
		</footer>
		<script>
		
			const HTML_TABLE_ID = "ObservedTimeline";
			const COLUMN_INDEX_FOR_DROPDOWN1 = 2;
			const COLUMN_INDEX_FOR_DROPDOWN2 = 3;
			const COLUMN_INDEX_FOR_DROPDOWN3 = 4;
			const COLUMN_INDEX_FOR_CHECKBOXES = 3;
			const SORT_BY_COLUMN_INDEX = 0;
			const SORT_AS_INTEGER_COLUMNS = [0, 6];
			const INTENT_COLUMN_INDEX = 4;

			// Specify which columns will be sorted as integers
			const integerColumns = [0, 6];

			// Add column names you want to hide
			const COLUMN_NAMES_TO_HIDE = ['Index', 'Color', 'DetailToolTip'];

			// Add column names which are set to bold text
			const COLUMN_NAMES_TO_BOLD = ['Detail'];

			// Change table row background color based on Color column value
			// Green to success, red to fail and Yellow for Info
			const COLOR_COLUMN_NAME = 'Color';
			const NAMED_COLUMNS_TO_COLOR = ['Status', 'Type', 'Intent', 'Detail'];

	
			// Constant object with predefined hex color values
			
			const FONT_COLOR_CONSTANTS = {
				'Green': '#00ff00',
				'Red': '#ff0000',
				'Yellow': '#ffffcc'
				// Add other colors as needed
			};

	
			function updateRowBackgroundOnColumnValueChange(tableId, columnIndex) {
			  let table = document.getElementById(tableId);
			  let rows = table.getElementsByTagName("tr");
			  let previousValue = null;
			  let currentColor = "rgba(208, 228, 245, 1)";

			  table.setAttribute("data-last-color", currentColor);

			  for (let i = 1; i < rows.length; i++) {
				let row = rows[i];

				// Skip hidden rows
				if (row.style.display === "none") {
				  continue;
				}

				let currentValue = row.getElementsByTagName("td")[columnIndex].textContent;

				if (previousValue !== null && currentValue !== previousValue) {
				  currentColor = table.getAttribute("data-last-color") === "rgba(242, 242, 242, 1)" ? "rgba(208, 228, 245, 1)" : "rgba(242, 242, 242, 1)";
				  table.setAttribute("data-last-color", currentColor);
				}

				row.style.backgroundColor = currentColor;
				previousValue = currentValue;
			  }
			}


			function setColumnBold(tableId, columnIndex) {
			  let table = document.getElementById(tableId);
			  let rows = table.getElementsByTagName("tr");

			  // Unbold previously selected column
			  if (table.hasAttribute("data-bold-column")) {
				let previousBoldColumn = parseInt(table.getAttribute("data-bold-column"));

				rows[0].getElementsByTagName("th")[previousBoldColumn].style.fontWeight = "normal";
				for (let i = 1; i < rows.length; i++) {
				  rows[i].getElementsByTagName("td")[previousBoldColumn].style.fontWeight = "normal";
				}
			  }

			  // Set header text to bold
			  rows[0].getElementsByTagName("th")[columnIndex].style.fontWeight = "bold";

			  // Set column values to bold
			  for (let i = 1; i < rows.length; i++) {
				rows[i].getElementsByTagName("td")[columnIndex].style.fontWeight = "bold";
			  }

			  // Save current bold column index
			  table.setAttribute("data-bold-column", columnIndex);
			}


			function mergeSort(arr, comparator) {
			  if (arr.length <= 1) {
				return arr;
			  }

			  const mid = Math.floor(arr.length / 2);
			  const left = mergeSort(arr.slice(0, mid), comparator);
			  const right = mergeSort(arr.slice(mid), comparator);

			  return merge(left, right, comparator);
			}

			function merge(left, right, comparator) {
			  let result = [];
			  let i = 0;
			  let j = 0;

			  while (i < left.length && j < right.length) {
				if (comparator(left[i], right[j]) <= 0) {
				  result.push(left[i]);
				  i++;
				} else {
				  result.push(right[j]);
				  j++;
				}
			  }

			  return result.concat(left.slice(i)).concat(right.slice(j));
			}

			// Declare a sortingDirections object outside the sortTable function to store the sorting directions for each column.
			const sortingDirections = {};


			function sortTable(n, tableId, dateColumns = []) {
			  let table, rows;
			  table = document.getElementById(tableId);

			  // Initialize the sorting direction for the column if it hasn't been set yet
			  if (!(n in sortingDirections)) {
				sortingDirections[n] = "asc";
			  }

			  // Remove existing arrow icons
			  let headerRow = table.getElementsByTagName("th");
			  for (let i = 0; i < headerRow.length; i++) {
				headerRow[i].innerHTML = headerRow[i].innerHTML.replace(
				  /<span>.*<\/span>/,
				  "<span>&#8597;</span>"
				);
			  }

			  const isDateColumn = dateColumns.includes(n);
			  rows = Array.from(table.rows).slice(1);

			  const comparator = (a, b) => {
				const x = a.cells[n].innerHTML.toLowerCase();
				const y = b.cells[n].innerHTML.toLowerCase();
				const isIntegerColumn = integerColumns.includes(n);

				if (isDateColumn) {
				  const xDate = getDateFromString(x);
				  const yDate = getDateFromString(y);

				  if (sortingDirections[n] === "asc") {
					return xDate - yDate;
				  } else {
					return yDate - xDate;
				  }
				} else if (isIntegerColumn) {
					const xInt = parseInt(x, 10) || 0;  // Use 0 if parsing fails
					const yInt = parseInt(y, 10) || 0;  // Use 0 if parsing fails
					if (sortingDirections[n] === "asc") {
					  return xInt - yInt;
					} else {
					  return yInt - xInt;
					}
				} else {
				  if (sortingDirections[n] === "asc") {
					return x.localeCompare(y);
				  } else {
					return y.localeCompare(x);
				  }
				}
			  };

			  const sortedRows = mergeSort(rows, comparator);

			  // Reinsert sorted rows into the table
			  for (let i = 0; i < sortedRows.length; i++) {
				table.tBodies[0].appendChild(sortedRows[i]);
			  }

			  // Update arrow icon for the last sorted column
			  if (sortingDirections[n] === "asc") {
				headerRow[n].innerHTML = headerRow[n].innerHTML.replace(
				  /<span>.*<\/span>/,
				  "<span>&#x25B2;</span>"
				);
			  } else {
				headerRow[n].innerHTML = headerRow[n].innerHTML.replace(
				  /<span>.*<\/span>/,
				  "<span>&#x25BC;</span>"
				);
			  }

			  // Create row coloring based on selected column
			  updateRowBackgroundOnColumnValueChange(HTML_TABLE_ID, n);

			  // Bold selected column header and text
			  setColumnBold(HTML_TABLE_ID, n);

			  // Toggle sorting direction for the next click
			  sortingDirections[n] = sortingDirections[n] === "asc" ? "desc" : "asc";
			}


			function getDateFromString(dateStr) {
				let [day, month, year, hours, minutes, seconds] = dateStr.split(/[. :]/);
				return new Date(year, month - 1, day, hours, minutes, seconds);
			}

			// Populates the given select element with unique values from the specified table and column
			function populateSelectWithUniqueColumnValues(tableId, column, selectId) {
			  let table = document.getElementById(tableId);
			  let rows = table.getElementsByTagName("tr");
			  let uniqueValues = {};

			  for (let i = 1; i < rows.length; i++) {
				let cellValue = rows[i].getElementsByTagName("td")[column].innerText;

				if (uniqueValues[cellValue]) {
				  uniqueValues[cellValue]++;
				} else {
				  uniqueValues[cellValue] = 1;
				}
			  }

			  let select = document.getElementById(selectId);

			  // Convert the uniqueValues object to an array of key-value pairs
			  let uniqueValuesArray = Object.entries(uniqueValues);

			  // Sort the array by the keys (unique column values)
			  uniqueValuesArray.sort((a, b) => a[0].localeCompare(b[0]));

			  // Find the longest text
			  let longestTextLength = Math.max(...uniqueValuesArray.map(([value, count]) => (value + " (" + count + ")").length));

			  // Loop through the sorted array to create the options with padded number values
			  for (let [value, count] of uniqueValuesArray) {
				let optionText = value + " (" + count + ")";
				let paddingLength = longestTextLength - optionText.length;
				let padding = "\u00A0".repeat(paddingLength);
				let option = document.createElement("option");
				option.value = value;
				option.text = value + padding + " (" + count + ")";
				select.add(option);
			  }
			}

			// This is used to extract AssignmentGroupDisplayName from complex a tag with span
			function getDirectChildTextNotWorking(parentNode) {
				let childNodes = parentNode.childNodes;
				let textContent = '';

				for(let i = 0; i < childNodes.length; i++) {
					if(childNodes[i].nodeType === Node.TEXT_NODE) {
						textContent += childNodes[i].nodeValue;
					}
				}

				return textContent.trim();  // Remove leading/trailing whitespaces
			}


			// Returns the textContent of the node if the node doesn't have any child nodes
			// In some Columns we have a href tag so we in those cases we need to get display value from <a> tag
			function getDirectChildText(node) {
			  if (!node || !node.hasChildNodes()) return node.textContent.trim();
			  return Array.from(node.childNodes)
				.filter(child => child.nodeType === Node.TEXT_NODE)
				.map(textNode => textNode.textContent)
				.join("");
			}


			// Filters the table based on the selected dropdown values, checkboxes, and search input
			// Should find AssignmentGroup and Filter displayNames for filtering from a href tag
			function combinedFilter(tableId, columnIndexForDropdown1, columnIndexForDropdown2, columnIndexForDropdown3, columnIndexForCheckboxes) {
			  let table = document.getElementById(tableId);
			  let rows = table.getElementsByTagName("tr");
			  let checkboxes = document.getElementsByClassName("filterCheckbox");
			  let dropdown1 = document.getElementById("filterDropdown1");
			  let dropdown2 = document.getElementById("filterDropdown2");
			  let dropdown3 = document.getElementById("filterDropdown3");
			  let searchInput = document.getElementById("searchInput");
			  let searchText = searchInput.value.toLowerCase();

			  let selectedDropdownValues1 = Array.from(dropdown1.selectedOptions).map(option => option.value);

			  for (let i = 1; i < rows.length; i++) {
				let row = rows[i];
				let cell1 = row.getElementsByTagName("td")[columnIndexForDropdown1];
				let cell2 = row.getElementsByTagName("td")[columnIndexForDropdown2];
				let cell3 = row.getElementsByTagName("td")[columnIndexForDropdown3];

				let cellValueDropdown1 = getDirectChildText(cell1.querySelector('a') || cell1);
				let cellValueDropdown2 = getDirectChildText(cell2.querySelector('a') || cell2);
				let cellValueDropdown3 = getDirectChildText(cell3.querySelector('a') || cell3);

				let showRowByDropdown1 = selectedDropdownValues1.includes("all") || selectedDropdownValues1.includes(cellValueDropdown1);
				let showRowByDropdown2 = dropdown2.value === "all" || cellValueDropdown2 === dropdown2.value;
				let showRowByDropdown3 = dropdown3.value === "all" || cellValueDropdown3 === dropdown3.value;

				let showRowByCheckboxes = true;
				for (let checkbox of checkboxes) {
				  if (checkbox.checked) {
					let cellValue = row.getElementsByTagName("td")[columnIndexForCheckboxes].textContent;
					let checkboxValues = checkbox.value.split(",");
					if (!checkboxValues.includes(cellValue)) {
					  showRowByCheckboxes = false;
					  break;
					}
				  }
				}

				let showRowBySearch = true;
				if (searchText) {
				  showRowBySearch = false;
				  let cells = row.getElementsByTagName("td");
				  for (let cell of cells) {
					if (getDirectChildText(cell.querySelector('a') || cell).toLowerCase().includes(searchText)) {
					  showRowBySearch = true;
					  break;
					}
				  }
				}

				row.style.display = (showRowByDropdown1 && showRowByDropdown2 && showRowByDropdown3 && showRowByCheckboxes && showRowBySearch) ? "" : "none";
			  }

			  let visibleRowCount = 0;
			  for (let i = 1; i < rows.length; i++) {
				if (rows[i].style.display !== 'none') {
				  visibleRowCount++;
				}
			  }

			  const noResultsMessage = document.getElementById('noResultsMessage');
			  if (visibleRowCount === 0) {
				noResultsMessage.style.display = 'block';
			  } else {
				noResultsMessage.style.display = 'none';
			  }
			}
			// function combinedFilter ends


			// Unchecks the other checkboxes in the group and updates the table filters
			function toggleCheckboxes(checkbox) {
			  let checkboxes = document.getElementsByClassName("filterCheckbox");
			  
			  for (let cb of checkboxes) {
				if (cb !== checkbox) {
				  cb.checked = false;
				}
			  }
			  
			  combinedFilter(HTML_TABLE_ID, COLUMN_INDEX_FOR_DROPDOWN1, COLUMN_INDEX_FOR_DROPDOWN2, COLUMN_INDEX_FOR_DROPDOWN3, COLUMN_INDEX_FOR_CHECKBOXES);
			}


			// Clears the search input and updates the table filters
			function clearSearch() {
			  let searchInput = document.getElementById("searchInput");
			  searchInput.value = "";
			  combinedFilter(HTML_TABLE_ID, COLUMN_INDEX_FOR_DROPDOWN1, COLUMN_INDEX_FOR_DROPDOWN2, COLUMN_INDEX_FOR_DROPDOWN3, COLUMN_INDEX_FOR_CHECKBOXES);
			}

			// Resets all filters and updates the table
			function resetFilters() {
			  let searchInput = document.getElementById("searchInput");
			  searchInput.value = "";
			  
			  let checkboxes = document.getElementsByClassName("filterCheckbox");
			  for (let checkbox of checkboxes) {
				checkbox.checked = false;
			  }
			  
			  let filterDropdown1 = document.getElementById("filterDropdown1");
			  filterDropdown1.value = "all";
			  
			  let filterDropdown2 = document.getElementById("filterDropdown2");
			  filterDropdown2.value = "all";
			  
			  let filterDropdown3 = document.getElementById("filterDropdown3");
			  filterDropdown3.value = "all";
			  
			  combinedFilter(HTML_TABLE_ID, COLUMN_INDEX_FOR_DROPDOWN1, COLUMN_INDEX_FOR_DROPDOWN2, COLUMN_INDEX_FOR_DROPDOWN3, COLUMN_INDEX_FOR_CHECKBOXES);
			}


			
			// Change Intent column cell background color based on value (Included or Excluded)
			function colorCells(tableId, columnIndex) {
			  let table = document.getElementById(tableId);
			  for (let i = 0; i < table.rows.length; i++) {
				let cell = table.rows[i].cells[columnIndex];
				if (cell) {
				  switch (cell.innerText.trim()) {
					case 'Included':
					  cell.style.backgroundColor = 'lightgreen';
					  break;
					case 'Excluded':
					  cell.style.backgroundColor = 'lightSalmon';
					  break;
					default:
					  break;
				  }
				}
			  }
			}


			// Event listeners for the dropdowns and checkboxes
			document.getElementById("filterDropdown1").addEventListener("change", function() {
			  combinedFilter(HTML_TABLE_ID, COLUMN_INDEX_FOR_DROPDOWN1, COLUMN_INDEX_FOR_DROPDOWN2, COLUMN_INDEX_FOR_DROPDOWN3, COLUMN_INDEX_FOR_CHECKBOXES);
			});

			document.getElementById("filterDropdown2").addEventListener("change", function() {
			  combinedFilter(HTML_TABLE_ID, COLUMN_INDEX_FOR_DROPDOWN1, COLUMN_INDEX_FOR_DROPDOWN2, COLUMN_INDEX_FOR_DROPDOWN3, COLUMN_INDEX_FOR_CHECKBOXES);
			});
			
			document.getElementById("filterDropdown3").addEventListener("change", function() {
			  combinedFilter(HTML_TABLE_ID, COLUMN_INDEX_FOR_DROPDOWN1, COLUMN_INDEX_FOR_DROPDOWN2, COLUMN_INDEX_FOR_DROPDOWN3, COLUMN_INDEX_FOR_CHECKBOXES);
			});

			let checkboxes = document.getElementsByClassName("filterCheckbox");
			for (let checkbox of checkboxes) {
			  checkbox.addEventListener("change", function() {
				combinedFilter(HTML_TABLE_ID, COLUMN_INDEX_FOR_DROPDOWN1, COLUMN_INDEX_FOR_DROPDOWN2, COLUMN_INDEX_FOR_DROPDOWN3, COLUMN_INDEX_FOR_CHECKBOXES);
			  });
			}

			// Add an event listener for the search input
			document.getElementById("searchInput").addEventListener("input", function() {
			  combinedFilter(HTML_TABLE_ID, COLUMN_INDEX_FOR_DROPDOWN1, COLUMN_INDEX_FOR_DROPDOWN2, COLUMN_INDEX_FOR_DROPDOWN3, COLUMN_INDEX_FOR_CHECKBOXES);
			});

		// Another approach to get function to run when loading page
		// This is not needed but left here on purpose just in case needed in the future
		//window.addEventListener('load', function() {
		//  sortTable(2, HTML_TABLE_ID, SORT_AS_INTEGER_COLUMNS);
		//});		



		// Hide named columns
		function hideColumnsByNames(tableId, columnNames) {
			let table = document.getElementById(tableId);
			let headerCells = table.getElementsByTagName('th');
			let columnIndices = [];

			// Find the indices of the columns with the given names
			for (let i = 0; i < headerCells.length; i++) {
				if (columnNames.includes(headerCells[i].innerText.trim())) {
					columnIndices.push(i);
				}
			}

			// If columns with the given names are found, hide them
			for (let index of columnIndices) {
				for (let row of table.rows) {
					row.cells[index].style.display = 'none';
				}
			}
		}

		// This static approach is used
		// if we don't use column sorting which changes bolding also to sorted column
		function boldColumnValues(tableId, columnNames) {
			let table = document.getElementById(tableId);
			let headerCells = table.getElementsByTagName('th');
			let columnIndices = [];

			// Find the indices of the columns with the given names
			for (let i = 0; i < headerCells.length; i++) {
				if (columnNames.includes(headerCells[i].innerText.trim())) {
					columnIndices.push(i);
				}
			}

			// If columns with the given names are found, set their text to bold
			for (let index of columnIndices) {
				for (let row of table.rows) {
					let cell = row.cells[index];
					if (cell) {
						cell.style.fontWeight = 'bold';
						// No need to change the background color, as it will be preserved.
					}
				}
			}
		}


		// Change html table row background colors based on Color column value
		function adjustColumnColorsByAnotherColumn(tableId, colorColumnName, targetColumnNames) {
			let table = document.getElementById(tableId);
			let headerCells = table.getElementsByTagName('th');
			let colorColumnIndex = null;
			let targetColumnIndices = [];

			// Find the index of the "Color" column (or the column specified in colorColumnName)
			for (let i = 0; i < headerCells.length; i++) {
				if (headerCells[i].innerText.trim() === colorColumnName) {
					colorColumnIndex = i;
				}
				if (targetColumnNames.includes(headerCells[i].innerText.trim())) {
					targetColumnIndices.push(i);
				}
			}

			if (colorColumnIndex === null) {
				console.error(`Couldn't find the '${colorColumnName}' column.`);
				return;
			}

			// Iterate over each row in the table
			for (let row of table.rows) {
				let colorValue = row.cells[colorColumnIndex].innerText.trim();
				if (colorValue) {
					for (let targetIndex of targetColumnIndices) {
						row.cells[targetIndex].style.backgroundColor = colorValue;
					}
				}
			}
		}


		// Change html table cell text colors based on Color column value
		function adjustColumnColorsByAnotherColumn(tableId, colorColumnName, targetColumnNames) {
			let table = document.getElementById(tableId);
			let headerCells = table.getElementsByTagName('th');
			let colorColumnIndex = null;
			let targetColumnIndices = [];

			// Find the index of the "Color" column (or the column specified in colorColumnName)
			for (let i = 0; i < headerCells.length; i++) {
				if (headerCells[i].innerText.trim() === colorColumnName) {
					colorColumnIndex = i;
				}
				if (targetColumnNames.includes(headerCells[i].innerText.trim())) {
					targetColumnIndices.push(i);
				}
			}

			if (colorColumnIndex === null) {
				console.error(`Couldn't find the '${colorColumnName}' column.`);
				return;
			}

			// Iterate over each row in the table
			for (let row of table.rows) {
				let colorValue = row.cells[colorColumnIndex].innerText.trim();
				if (colorValue) {
					let fontColor = FONT_COLOR_CONSTANTS[colorValue] || colorValue;
					for (let targetIndex of targetColumnIndices) {
						row.cells[targetIndex].style.color = fontColor;
					}
				}
			}
		}
		
		
		// Add HoverOn ToolTip from named ToolTip Column to named Target Column
		function addTooltipFromSourceColumnToTargetColumn(tableId, tooltipSourceColumnName, targetColumnNames) {
			let table = document.getElementById(tableId);
			let headerCells = table.getElementsByTagName('th');
			let sourceColumnIndex = null;
			let targetColumnIndices = [];

			// Find the index of the tooltip source column
			for (let i = 0; i < headerCells.length; i++) {
				if (headerCells[i].innerText.trim() === tooltipSourceColumnName) {
					sourceColumnIndex = i;
				}
				if (targetColumnNames.includes(headerCells[i].innerText.trim())) {
					targetColumnIndices.push(i);
				}
			}

			if (sourceColumnIndex === null) {
				console.error(`Couldn't find the '${tooltipSourceColumnName}' column.`);
				return;
			}

			// Iterate over each row in the table
			for (let row of table.rows) {
				if (row.cells[sourceColumnIndex] && row.cells[sourceColumnIndex].innerText.trim() !== "") {
					let tooltipText = row.cells[sourceColumnIndex].innerText.trim();
					
					for (let targetIndex of targetColumnIndices) {
						if (row.cells[targetIndex]) {
							// Create the tooltip span element
							let tooltipSpan = document.createElement('span');
							tooltipSpan.className = 'tooltiptext';
							tooltipSpan.innerText = tooltipText;

							// Append the tooltip to the target cell and add the tooltip class to the cell
							row.cells[targetIndex].classList.add('tooltip');
							row.cells[targetIndex].appendChild(tooltipSpan);
						}
					}
				}
			}
		}
		
		
		window.onload = function() {
			
			// Call this function to populate the first dropdown with unique values from the specified table and column
			populateSelectWithUniqueColumnValues(HTML_TABLE_ID, COLUMN_INDEX_FOR_DROPDOWN1, "filterDropdown1");

			// Call this function to populate the second dropdown with unique values from the specified table and column
			populateSelectWithUniqueColumnValues(HTML_TABLE_ID, COLUMN_INDEX_FOR_DROPDOWN2, "filterDropdown2");
			
			// Call this function to populate the third dropdown with unique values from the specified table and column
			populateSelectWithUniqueColumnValues(HTML_TABLE_ID, COLUMN_INDEX_FOR_DROPDOWN3, "filterDropdown3");

			// Not needed anymore because sorting will also do this automatically
			//updateRowBackgroundOnColumnValueChange(HTML_TABLE_ID, 2);

			// Sort table by name so user knowns which column was sorted
			//sortTable(SORT_BY_COLUMN_INDEX, HTML_TABLE_ID, SORT_AS_INTEGER_COLUMNS);
			
			// Change Intent column background color
			colorCells(HTML_TABLE_ID, INTENT_COLUMN_INDEX);
			
			// Hide columns
			hideColumnsByNames(HTML_TABLE_ID, COLUMN_NAMES_TO_HIDE);
			
			// Bold named column(s)
			boldColumnValues(HTML_TABLE_ID, COLUMN_NAMES_TO_BOLD);

			// Change row background color value based on Color column value
			//adjustColumnColorsByAnotherColumn(HTML_TABLE_ID, COLOR_COLUMN_NAME, NAMED_COLUMNS_TO_COLOR);
			
			// Change cell text color value based on Color column value
			adjustColumnColorsByAnotherColumn(HTML_TABLE_ID, COLOR_COLUMN_NAME, NAMED_COLUMNS_TO_COLOR);
			
			// Add ToolTips to Detail column
			addTooltipFromSourceColumnToTargetColumn(HTML_TABLE_ID, 'DetailToolTip', 'Detail')
			
		};
		</script>
'@

		$AppDownloadPreContent = @"
		<p id="noResultsMessage" style="display: none;">No results found.</p>
		<h2 id=`"AppDownloadStatistics`">App Download Statistics</h2>
"@

		# Application Download Summary
		$ApplicationDownloadStatisticsHTML = $ApplicationDownloadStatistics | ConvertTo-Html -Fragment -PreContent $AppDownloadPreContent | Out-String

		# Add TableId
		$TableId = 'ApplicationDownloadStatistics'
		$ApplicationDownloadStatisticsHTML = $ApplicationDownloadStatisticsHTML.Replace('<table>',"<table id=`"$TableId`">")


		$Title = "Get-IntuneManagementExtensionDiagnostics Observed Timeline Report"
		ConvertTo-HTML -head $head -PostContent $observedTimelineHTML, $ApplicationDownloadStatisticsHTML, $JavascriptPostContent -PreContent $PreContent -Title $Title | Out-File "$ReportSavePath\$HTMLFileName"
		$Success = $?

		if (-not ($Success)) {
			Write-Error "Error creating HTML file."
			Write-Host "Script will exit..."
			Pause
			Exit 1
		}
		else {
			Write-Host "Get-IntuneManagementExtensionDiagnostics report HTML file created:`n`t$ReportSavePath\$HTMLFileName`n" -ForegroundColor Green
		}
			
		############################################################
		# Open HTML file

		# Check file exists and is bigger than 0
		# File should exist already but years ago slow computer/disk caused some problems
		# so this is hopefully not needed workaround
		# Wait max. of 20 seconds

		if(-not $DoNotOpenReportAutomatically) {
			$i = 0
			$filesize = 0
			do {
				Write-Host "Double check HTML file creation is really done (round $i)"
				$filesize = 0
				Start-Sleep -Seconds 2
				try {
					$HTMLFile = Get-ChildItem "$ReportSavePath\$HTMLFileName"
					$filesize = $HTMLFile.Length
				}
				catch {
					# Something went wrong, waiting for next round.
					Write-Host "Trouble getting file size, waiting 2 seconds and trying again..."
				}
				if ($filesize -eq 0) { Write-Host "Filesize is 0kB so waiting for a while for file creation to finish" }

				$i += 1
			} while (($i -lt 10) -and ($filesize -eq 0))

			Write-Host "Opening created file:`n$ReportSavePath\$HTMLFileName`n"
			try {
				Invoke-Item "$ReportSavePath\$HTMLFileName"
			}
			catch {
				Write-Host "Error opening file automatically to browser. Open file manually:`n$ReportSavePath\$HTMLFileName`n" -ForegroundColor Red
			}
		} else {
			Write-Host "`nNote! Parameter -DoNotOpenReportAutomatically specified. Report was not opened automatically to web browser`n" -ForegroundColor Yellow
		}
		
		} # Try creation HTML report end
		  catch {
			Write-Error "$($_.Exception.GetType().FullName)"
			Write-Error "$($_.Exception.Message)"
			Write-Error "Error creating HTML report: $ReportSavePath\$HTMLFileName"
		}
			
			
	} # If $ExportHTML end


	if($ShowLogViewerUI -or $LogViewerUI) {
		Write-Host "Show logs in Out-GridView"
		
		# Show log in Out-GridView
		if($SelectedLogFiles -is [Array]) {
			# Sort log entries and add index number
			# This is needed with multiple files so log entries
			# can be sorted based on this index column

			#$SelectedLines = $LogEntryList | Select-Object -Property Index, FileName, Line, DateTime, Multiline, ProcessRunTime, Message, Component, Context, Type, Thread, File | Sort-Object -Property Index | Out-GridView -Title "Intune IME Log Viewer $($SelectedLogFiles.Name -join " ")" -OutputMode Multiple
			
			$SelectedLines = $LogEntryList | Select-Object -Property Index, DateTime, Multiline, ProcessRunTime, Message, Component, Context, Type, Thread, File, FileName, Line | Sort-Object -Property Index | Out-GridView -Title "Intune IME Log Viewer $($SelectedLogFiles.Name -join " ")" -OutputMode Multiple
		} else {
			
			#$SelectedLines = $LogEntryList | Select-Object -Property Index, FileName, Line, DateTime, Multiline, ProcessRunTime, Message, Component, Context, Type, Thread, File | Out-GridView -Title "Intune IME Log Viewer $LogFilePath" -OutputMode Multiple
			
			$SelectedLines = $LogEntryList | Select-Object -Property Index, DateTime, Multiline, ProcessRunTime, Message, Component, Context, Type, Thread, File, FileName, Line | Out-GridView -Title "Intune IME Log Viewer $LogFilePath" -OutputMode Multiple
		}
	} else {

		Write-Host "`nTip: Use Parameter -ShowLogViewerUI to get LogViewerUI for graphical log viewing/debugging`n" -ForegroundColor Cyan
		
		# Not showing Out-GridView so we need to make sure Do-While loop will exit
		$SelectedLines = $null
		Break
	}
	
} While ($SelectedLines.Message -eq 'RELOAD LOG FILE(S).      Select this line and OK button from right botton corner')

#Write-Host "Script end"
