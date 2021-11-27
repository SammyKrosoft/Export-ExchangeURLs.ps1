<#PSScriptInfo

.VERSION 3.3.6

.GUID 0a1b89dc-e2b3-4e34-b1ad-e86ca7f6833d

.AUTHOR Sam Drey

.COMPANYNAME Microsoft Canada

#>
 
<#

.SYNOPSIS
	This script dumps the auth settings of all vDirs of your Exchange servers in a CSV file
	
.DESCRIPTION
	This script dumps the auth settings of all vDirs of your Exchange servers in a CSV file by default,
	or dumps the info to the screen (you can then redirect to a file if you wish) if you
	use the -DoNotExport switch with the script.

.PARAMETER E2010
	When specified, will export E2010 only, or if specified with E2007 / E2013 / 2016, will export
	the specified versions.
	
	When none of E2007, E2013, E2016 are specified, all versions are scanned.

.PARAMETER E2013
	When specified, will export E2013 only, or if specified with E2007 / E2016, will export
	the specified versions.

	When none of E2007, E2013, E2016 are specified, all versions are scanned.

.PARAMETER E2016
	When specified, will export E2016 only, or if specified with E2007 / E2013, will export
	the specified versions.

	When none of E2007, E2013, E2016 are specified, all versions are scanned.

.PARAMETER DoNotExport
	This parameter tells the script not to export to a file. The results will just be
	dumped to the screen.

.PARAMETER CheckVersion
	This parameter dumps the current script version - the script stops processing after displaying the
	version if this parameter is specified, no matter what other parameter is also specified.

.INPUTS
    None.

.OUTPUTS
    Exports a CSV file, or the screen.

.EXAMPLE
.\Export-vDirAuthSettings.ps1
Will export all URLs for all Exchange servers' virtual directories into a file named ExchangeURLs_Day_Date_Time.csv

.EXAMPLE
.\Export-vDirAuthSettings.ps1 -DoNotExport
Will just print the results into the PowerShell console.

.NOTES
V3.2 -> Fixed mistake : Typo in OWAexternalURL, ECP and EWS, and OAB ... fixed
V3 -> adding switches to select which Exchange version to check (not finished yet)

.LINK
	https://github.com/SammyKrosoft

#>

[CmdletBinding(DefaultParameterSetName = "E2010")]
Param(
	[Parameter(Mandatory = $False, Position = 1, ParameterSetName = "Ex2010")]
	[Parameter(Mandatory = $False, Position = 1, ParameterSetName = "Ex2013")]
	[Parameter(Mandatory = $False, Position = 1, ParameterSetName = "Ex2016")]
	[switch]$DoNotExport,
	[Parameter(Mandatory = $False, Position = 2, ParameterSetName = "Ex2010")][switch]$E2010,
	[Parameter(Mandatory = $False, Position = 3, ParameterSetName = "Ex2013")][switch]$E2013,
	[Parameter(Mandatory = $False, Position = 4, ParameterSetName = "Ex2016")][switch]$E2016,
	[Parameter(Mandatory = $False, Position = 5, ParameterSetName = "checkversion")][Switch] $CheckVersion
)

<# ------- SCRIPT_HEADER (Only Get-Help comments and Param() above this point) ------- #>
#Initializing a $Stopwatch variable to use to measure script execution
$stopwatch = [system.diagnostics.stopwatch]::StartNew()
#Using Write-Debug and playing with $DebugPreference -> "Continue" will output whatever you put on Write-Debug "Your text/values"
# and "SilentlyContinue" will output nothing on Write-Debug "Your text/values"
$DebugPreference = "Continue"
# Set Error Action to your needs
$ErrorActionPreference = "SilentlyContinue"
#Script Version
$ScriptVersion = '3.3.6'
<# Version History
v3.3.6 -> fixed another mistake : used Get-ExchangeServer | which roles = "*Client*" - but in Exchange 2016/2019 there are no more Client Access roles ... server role is just "Mailbox"
v3.3.5 - output on executing's user Documents folder instead of script directory
v3.3.4 - added export server Site information as ServerSite
v3.3.3 - added -ADPropertiesOnly for Get-MAPIVirtualDirectory
v3.3.2 - removed "AutoDiscName" column as it's the same value as "ServerName"
v3.3.1 - changed author's name to Sam Drey and current company
v3.3 - added MAPI URLs export
v3.2 - See Notes
v3.1 : fixed Exchange version switches, Get-ClientAccessSErvice is for E2016 only, NOT E2013.
v1.0 -> v2
Added export of Outlook Anywhere with External Hostname (E2010, E2013, E2016) and Internal Hostname (not existing in E2010)
Fixed output file name
Added -DoNoExport switch, to not export to a file...
#> 
# Log or report file definition
# NOTE: use #PSScriptRoot in Powershell 3.0 and later or use $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition in Powershell 2.0
# $LogOrReportFile1 = "$PSScriptRoot\ReportOrLogFile_$(get-date -f yyyy-MM-dd-hh-mm-ss).csv"
# Other Option for Log or report file definition (use one of these)
# $LogOrReportFile2 = "$PSScriptRoot\PowerShellScriptExecuted-$(Get-Date -Format 'dd-MMMM-yyyy-hh-mm-ss-tt').txt"
If ($CheckVersion) {Write-Host "Script Version v$ScriptVersion";exit}
<# ---------------------------- /SCRIPT_HEADER ---------------------------- #>
<# -------------------------- DECLARATIONS -------------------------- #>
<# /DECLARATIONS #>
<# -------------------------- FUNCTIONS -------------------------- #>
<# /FUNCTIONS #>
<# -------------------------- EXECUTIONS -------------------------- #>
#Loading Exchange 2010 snapins enabling script to be executed on a basic Powershell session
#Note: you must have Exchange Admin tools installed on the machine where you run this.
Add-PSSnapin microsoft.exchange.management.powershell.admin -erroraction 'SilentlyContinue' | OUT-NULL
Add-PSsnapin Microsoft.Exchange.Management.PowerShell.E2010 -erroraction 'SilentlyContinue' | OUT-NULL
Add-PSsnapin Microsoft.Exchange.Management.PowerShell.Setup -erroraction 'SilentlyContinue' | OUT-NULL
Add-PSsnapin Microsoft.Exchange.Management.PowerShell.Support -erroraction 'SilentlyContinue' | OUT-NULL
#For Exchange 2007 and 2013, add the corresponding modules/snapins, or simply execute the script into an Exchange MAnagement Shell :-)

#Getting all Exchange servers in an array
#Note: you can target only one server, or get servers list from a file,
#just change the $Servers = @(Get-ClientAccessServer) line with $Servers = @(Get-content ServersList.txt) for example to get servers from a list...
$Servers = @()

#Sample Server Filtering:
# $Exchange2010CASServers = Get-ExchangeServer * | ? {$_.serverRole -match "Client" -and $_.AdminDisplayVersion -match '14.'} | Ft name, serverRole, AdminDisplayVersion

$ServerVersionFilter = $null

$ExchangeServers = Get-ExchangeServer

#Filtering out Exchange versions
# 14.x => Exchange 2010
# 15.0 => Exchange 2013
# 15.1 => Exchange 2016
if ($E2010) {
	Write-debug "E2010 switch on";
	$ExchangeServers = $ExchangeServers | ? {$_.ServerRole -match "Client" -and $_.AdminDisplayVersion -match '14\.'}
	If ($ExchangeServers -eq $null) {
		$msg = "No Exchange 2010 servers found - Try -E2013 or E2016 or no switch ... exiting.";
		Write-host $msg
		exit
	} Else {
		$msg = "Found $($ExchangeServers.count) Exchange 2010 servers."
		Write-host $msg
	}
}
if ($E2013) {
	Write-debug "E2013 switch on";$ExchangeServers = $ExchangeServers | ? {$_.AdminDisplayVersion -match '15\.0'}
	If ($ExchangeServers -eq $null) {
		$msg = "No Exchange 2013 servers found - Try -E2013 or E2016 or no switch ... exiting.";
		Write-host $msg
		exit
	}Else {
		$msg = "Found $($ExchangeServers.count) Exchange 2013 servers."
		Write-host $msg
	}
}
if ($E2016) {
	Write-debug "E2016 switch on";$ExchangeServers = $ExchangeServers | ? {$_.AdminDisplayVersion -match '15\.1'}
	If ($ExchangeServers -eq $null) {
		$msg = "No Exchange 2016 servers found - Try -E2013 or E2016 or no switch ... exiting.";
		Write-host $msg
		exit
	}Else {
		$msg = "Found $($ExchangeServers.count) Exchange 2016 servers."
		Write-host $msg
	}
}

If (!$E2010 -and !$E2013 -and !$E2016) {
	Write-debug "No switches - Discovering all servers`nNOTE: always run scripts from highest version of Exchange otherwise you won't see all servers";
	If ($ExchangeServers -eq $null) {
		$msg = "No Exchange Exchange 2010, 2013 or 2016 servers found ... exiting.";
		Write-host $msg
		exit
	}Else {
		$msg = "Found $($ExchangeServers.count) Exchange servers."
		Write-host $msg
	}
}

$Servers = $ExchangeServers
$Servers | Select Name, ServerRole, @{Label = "Exchange Version"; Expression = {$_.AdminDisplayVersion}} | ft -a

# DEBUG PURPOSE : MANUAL EXIT (BREAK) BELOW - COMMENT WHEN NOT DEBUGGING
# exit


#Initializing counters to setup a progress bar based on the number of servers browsed
# (more useful in an environment where you have dozen of servers - had 45 in mine)
	$Counter=0
    $Total=$Servers.count	
#Initializing the variable where I'll put all the results of my object browsing
    $report = @()
#For each server discovered in the "$Servers = Get-ClientAccessServer" line, 
# grab the Virtal Directories properties and store it in a custom Powershell object, 
# and then add this object in the $report array variable to eventually dump the whole result in a text (CSV) file.
foreach( $Server in $Servers)
{
    #$Computername=$Server.Name   <- not needed for now
	#This is to print the progress bar incrementing on each server (increment is later in the script $Counter++ it is...
    $Pct=($Counter/$Total)*100    
    Write-Progress -Activity "Processing Server $Server" -status "Server $Counter of $Total" -percentcomplete $pct
	#For the current server, get the main vDir settings (including AutodiscoverServiceInternalURI which is important to determine 
	#whether the Autodiscover service will be hit using the Load Balancer (recommended).
	$EAS = Get-ActiveSyncVirtualDirectory -Server $Server -ADPropertiesOnly | Select name,server,InternalAuthenticationMethods,ExternalAuthenticationMethods
	$OAB = Get-OabVirtualDirectory -Server $Server -ADPropertiesOnly | ? {$_.Name -like "*OAB*"} | Select name,server,InternalAuthenticationMethods,ExternalAuthenticationMethods
	$OWA = Get-OwaVirtualDirectory -Server $Server -ADPropertiesOnly | Select name,server,InternalAuthenticationMethods,ExternalAuthenticationMethods
	$ECP = Get-EcpVirtualDirectory -Server $Server -ADPropertiesOnly | Select name,server,InternalAuthenticationMethods,ExternalAuthenticationMethods
	#testing if there is an Exchange 2013/2016 in the $ExchangeServers collection - If TRUE then use Get-ClientAccessService, ELSE user Get-ClientAccessServer

	$EWS = Get-WebServicesVirtualDirectory -Server $Server -ADPropertiesOnly | Select name,server,InternalAuthenticationMethods,ExternalAuthenticationMethods
    $OA = Get-OutlookAnywhere -Server $Server -ADPropertiesOnly | Select name,server,InternalClientAuthenticationMethod,ExternalClientAuthenticationMethod
    
    #If you want to dump more things, use the below line as a sample:
	#$ServiceToDump = Get-Whatever -Server $Server | Select Property1, property2, ....   <- don't need the "Select property", you can omit this, it will just get all attributes...
	$MAPI = Get-MAPIVirtualDirectory -Server $Server -ADPropertiesOnly | Select name,server,InternalAuthenticationMethods,ExternalAuthenticationMethods # ....   <- don't need the "Select property", you can omit this, it will just get all attributes...

   	#Initializing a new Powershell object to store our discovered properties
    $Obj = New-Object PSObject

	#the below is a template if you need to dump more things into the final report
	#just replace the "ServiceToDump" string with the service you with to dump - don't forget to 
	#Get something above like the $Service = Get-whatever -Server
	#$Obj | Add-Member -MemberType NoteProperty -Name "ServiceToDump-vDirNAme" -Value $ServiceToDump.Name
	#$Obj | Add-Member -MemberType NoteProperty -Name "ServiceToDump-InternalURL" -Value $ServiceToDump.InternalURL
	#$Obj | Add-Member -MemberType NoteProperty -Name "ServiceToDump-ExternalURL" -Value $ServiceToDump.ExternalURL	

		
	$Obj | Add-Member -MemberType NoteProperty -Name "ServerName" -Value $Server.Name
	$Obj | Add-Member -MemberType NoteProperty -Name "ServerSite" -Value $Server.Site
	$Obj | Add-Member -MemberType NoteProperty -Name "ServerVersion" -Value $Server.AdminDisplayVersion
	#$Obj | Add-Member -MemberType NoteProperty -Name "EASName" -Value $EAS.Name
    $Obj | Add-Member -MemberType NoteProperty -Name "EAS_InternalAuthenticationMethods" -Value $($EAS.InternalAuthenticationMethods | out-string)
	$Obj | Add-Member -MemberType NoteProperty -Name "EAS_ExternalAuthenticationMethods" -Value $EAS.ExternalAuthenticationMethods.tostring()
	# $Obj | Add-Member -MemberType NoteProperty -Name "OABName" -Value $OAB.Name
	$Obj | Add-Member -MemberType NoteProperty -Name "OAB_InternalAuthenticationMethods" -Value $OAB.InternalAuthenticationMethods
	$Obj | Add-Member -MemberType NoteProperty -Name "OAB_ExternalAuthenticationMethods" -Value $OAB.ExternalAuthenticationMethods
	#$Obj | Add-Member -MemberType NoteProperty -Name "OWAName" -Value $OWA.Name
	$Obj | Add-Member -MemberType NoteProperty -Name "OWA_InternalAuthenticationMethods" -Value $OWA.InternalAuthenticationMethods
	$Obj | Add-Member -MemberType NoteProperty -Name "OWA_ExternalAuthenticationMethods" -Value $OWA.ExternalAuthenticationMethods
	# $Obj | Add-Member -MemberType NoteProperty -Name "ECPName" -Value $ECP.Name
	$Obj | Add-Member -MemberType NoteProperty -Name "ECP_InternalAuthenticationMethods" -Value $ECP.InternalAuthenticationMethods
	$Obj | Add-Member -MemberType NoteProperty -Name "ECP_ExternalAuthenticationMethods" -Value $ECP.ExternalAuthenticationMethods	
	# $Obj | Add-Member -MemberType NoteProperty -Name "EWSName" -Value $EWS.Name
	$Obj | Add-Member -MemberType NoteProperty -Name "EWS_InternalAuthenticationMethods" -Value $EWS.InternalAuthenticationMethods
    $Obj | Add-Member -MemberType NoteProperty -Name "EWS_ExternalAuthenticationMethods" -Value $EWS.ExternalAuthenticationMethods
    $Obj | Add-Member -MemberType NoteProperty -Name "OA_InternalAuthenticationMethod" -Value $OA.InternalClientAuthenticationMethod
    $Obj | Add-Member -MemberType NoteProperty -Name "OA_ExternalAuthenticationMethod" -Value $OA.ExternalClientAuthenticationMethod
	$Obj | Add-Member -MemberType NoteProperty -Name "MAPIHTTP_InternalAuthenticationMethods" -Value $MAPI.InternalAuthenticationMethods
	$Obj | Add-Member -MemberType NoteProperty -Name "MAPIHTTP_ExternalAuthenticationMethods" -Value $MAPI.ExternalAuthenticationMethods

		
		#Appending the current object into the $report variable (it's an array, remember)
        $report += $Obj
		
		#Incrementing the Counter for the progress bar
        $Counter++
    }
	
	
	If (!($DoNotExport)){
		#Building the file name string using date, time, seconds ...
		$DateAppend = Get-Date -Format "ddd-dd-MM-yyyy-\T\i\m\e-HH-mm-ss"
		$CSVFilename=($env:userprofile) + ("\Documents")+"\ExchangeURLs_"+$DateAppend+".csv"
		#Exporting the final result into the output file (see just above for the file string building...
		$report | Export-csv -notypeinformation -encoding Unicode $CSVFilename
		Notepad $CSVFilename
	} Else {
		Write-Host "Won't create a file because you specified the -DoNotExport parameter" -ForegroundColor Yellow -BackgroundColor Blue
		Write-Host "Just dumping to the screen this time ..." -ForegroundColor DarkBlue -BackgroundColor red
		$Report
	}

<# /EXECUTIONS #>
<# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
$stopwatch.Stop()
Write-Host "`n`nThe script took $($StopWatch.Elapsed.TotalSeconds) seconds to execute..."
<# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>
