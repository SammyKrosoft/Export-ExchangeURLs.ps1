<#
.SYNOPSIS
    Exports URLs of Exchange services (OWA, ECP, ActiveSync, OAB, Outlook Anywhere and AutoDiscover)
    on a CSV file, located on the same folder as the script.

.DESCRIPTION
    Same as Synopsis (for now)

.PARAMETER None
    No parameters for now

.INPUTS
    None. You cannot pipe objects to that script.

.OUTPUTS
    CSV file with the Exchange servers names and URLs

.EXAMPLE
.\Export-ExchangeURLs.ps1
This will launch the script and export the URLs as a CSV file

.NOTES
None

.LINK
    https://gallery.technet.microsoft.com/Powershell-script-to-ccde9d5f

#>
[CmdLetBinding(DefaultParameterSetName = "NormalRun")]
Param(
    [Parameter(Mandatory = $false, Position = 3, ParameterSetName = "CheckOnly")][switch]$CheckVersion
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
$ScriptVersion = "1"
<# Version changes
v0.1 : first script version
v0.1 -> v0.5 : 
#>
$ScriptName = $MyInvocation.MyCommand.Name
If ($CheckVersion) {Write-Host "SCRIPT NAME     : $ScriptName `nSCRIPT VERSION  : $ScriptVersion";exit}
# Log or report file definition
# NOTE: use $PSScriptRoot in Powershell 3.0 and later or use $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition in Powershell 2.0
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$OutputReport = "$ScriptPath\$($ScriptName)_$(get-date -f yyyy-MM-dd-hh-mm-ss).csv"
# Other Option for Log or report file definition (use one of these)
$ScriptLog = "$ScriptPath\$($ScriptName)-$(Get-Date -Format 'dd-MMMM-yyyy-hh-mm-ss-tt').txt"
<# ---------------------------- /SCRIPT_HEADER ---------------------------- #>
<# -------------------------- DECLARATIONS -------------------------- #>

<# /DECLARATIONS #>
<# -------------------------- FUNCTIONS -------------------------- #>

<# /FUNCTIONS #>
<# -------------------------- EXECUTIONS -------------------------- #>
$ErrorActionPreference='SilentlyContinue'

#Loading Exchange 2010 snapins enabling script to be executed on a basic Powershell session
#Note: you must have Exchange Admin tools installed on the machine where you run this.
Add-PSSnapin microsoft.exchange.management.powershell.admin -erroraction 'SilentlyContinue' | OUT-NULL
Add-PSsnapin Microsoft.Exchange.Management.PowerShell.E2010 -erroraction 'SilentlyContinue' | OUT-NULL
Add-PSsnapin Microsoft.Exchange.Management.PowerShell.Setup -erroraction 'SilentlyContinue'  | OUT-NULL
Add-PSsnapin Microsoft.Exchange.Management.PowerShell.Support -erroraction 'SilentlyContinue'  | OUT-NULL
#For Exchange 2007 and 2013, add the corresponding modules/snapins, or simply execute the script into an Exchange MAnagement Shell :-)

#Saving script path to use the same path to store the output file
$ScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition

#Getting all Exchange servers in an array
#Note: you can target only one server, or get servers list from a file,
#just change the $Servers = @(Get-ClientAccessServer) line with $Servers = @(Get-content ServersList.txt) for example to get servers from a list...
$Servers = @(Get-ClientAccessServer)

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
	$EAS = Get-ActiveSyncVirtualDirectory -Server $Server| Select Name, InternalURL,externalURL
	$OAB = Get-OabVirtualDirectory -Server $Server| Select Name,internalURL,externalURL
	$OWA = Get-OwaVirtualDirectory -Server $Server| Select Name,InternalURL,externalURL
	$ECP = Get-EcpVirtualDirectory -Server $Server| Select Name,InternalURL,externalURL
	$AutoDisc = get-ClientAccessServer $Server | Select name,identity,AutodiscoverServiceInternalUri
    $EWS = Get-WebServicesVirtualDirectory -Server $Server| Select NAme,identity,internalURL,externalURL
    $OA = Get-OutlookAnywhere -Server $Server | Select Name,ExternalHostname
	#If you want to dump more things, use the below line as a sample:
	#$ServiceToDump = Get-Whatever -Server $Server | Select Property1, property2, ....   <- don't need the "Select property", you can omit this, it will just get all attributes...

		#the below is a template if you need to dump more things into the final report
		#just replace the "ServiceToDump" string with the service you with to dump - don't forget to 
		#Get something above like the $Service = Get-whatever -Server
		#$Obj | Add-Member -MemberType NoteProperty -Name "ServiceToDump-vDirNAme" -Value $ServiceToDump.Name
		#$Obj | Add-Member -MemberType NoteProperty -Name "ServiceToDump-InternalURL" -Value $ServiceToDump.InternalURL
		#$Obj | Add-Member -MemberType NoteProperty -Name "ServiceToDump-ExernalURL" -Value $ServiceToDump.ExternalURL	
	   
	   	#Initializing a new Powershell object to store our discovered properties
        $Obj = New-Object PSObject
		
		#the below is a template if you need to dump more things into the final report
		#just replace the "ServiceToDump" string with the service you with to dump - don't forget to 
		#Get something above like the $Service = Get-whatever -Server
		#$Obj | Add-Member -MemberType NoteProperty -Name "ServiceToDump-vDirNAme" -Value $ServiceToDump.Name
		#$Obj | Add-Member -MemberType NoteProperty -Name "ServiceToDump-InternalURL" -Value $ServiceToDump.InternalURL
		#$Obj | Add-Member -MemberType NoteProperty -Name "ServiceToDump-ExernalURL" -Value $ServiceToDump.ExternalURL	

        
		$Obj | Add-Member -MemberType NoteProperty -Name "ServerName" -Value $Server.Name
		$Obj | Add-Member -MemberType NoteProperty -Name "AutoDisc-vDirNAme" -Value $AutoDisc.Name
		$Obj | Add-Member -MemberType NoteProperty -Name "AutoDisc-URI" -Value $AutoDisc.AutodiscoverServiceInternalURI
        $Obj | Add-Member -MemberType NoteProperty -Name "EAS-vDirNAme" -Value $EAS.Name
        $Obj | Add-Member -MemberType NoteProperty -Name "EAS-InternalURL" -Value $EAS.InternalURL
		$Obj | Add-Member -MemberType NoteProperty -Name "EAS-ExternalURL" -Value $EAS.ExternalURL
		$Obj | Add-Member -MemberType NoteProperty -Name "OAB-vDirNAme" -Value $OAB.Name
		$Obj | Add-Member -MemberType NoteProperty -Name "OAB-InternalURL" -Value $OAB.InternalURL
		$Obj | Add-Member -MemberType NoteProperty -Name "OAB-ExernalURL" -Value $OAB.ExternalURL
		$Obj | Add-Member -MemberType NoteProperty -Name "OWA-vDirNAme" -Value $OWA.Name
		$Obj | Add-Member -MemberType NoteProperty -Name "OWA-InternalURL" -Value $OWA.InternalURL
		$Obj | Add-Member -MemberType NoteProperty -Name "OWA-ExernalURL" -Value $OWA.ExternalURL
		$Obj | Add-Member -MemberType NoteProperty -Name "ECP-vDirNAme" -Value $ECP.Name
		$Obj | Add-Member -MemberType NoteProperty -Name "ECP-InternalURL" -Value $ECP.InternalURL
		$Obj | Add-Member -MemberType NoteProperty -Name "ECP-ExernalURL" -Value $ECP.ExternalURL	
		$Obj | Add-Member -MemberType NoteProperty -Name "EWS-vDirNAme" -Value $EWS.Name
		$Obj | Add-Member -MemberType NoteProperty -Name "EWS-InternalURL" -Value $EWS.InternalURL
        $Obj | Add-Member -MemberType NoteProperty -Name "EWS-ExernalURL" -Value $EWS.ExternalURL
        $Obj | Add-Member -MemberType NoteProperty -Name "OA Hostname" -Value $OA.Name
        $Obj | Add-Member -MemberType NoteProperty -Name "OA Hostname" -Value $OA.ExternalHostname
		
		
		#Appending the current object into the $report variable (it's an array, remember)
        $report += $Obj
		
		#Incrementing the Counter for the progress bar
        $Counter++
    }
	
	#Building the file name string using date, time, seconds ...
	$DateAppend = Get-Date -Format "ddd-dd-MM-yyyy-\T\i\m\e-HH-mm-ss"
    $CSVFilename=$ScriptPath+"\EASInformation"+$DateAppend+".csv"
	
	#Exporting the final result into the output file (see just above for the file string building...
    $report | Export-csv -notypeinformation -encoding Unicode $CSVFilename

<# /EXECUTIONS #>
<# -------------------------- CLEANUP VARIABLES -------------------------- #>
$report = $null
$DateAppend = $null
$CSVFilename = $null
$Obj = $null
$Counter = $null
<# /CLEANUP VARIABLES#>
<# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
$stopwatch.Stop()
$msg = "`n`nThe script took $([math]::round($($StopWatch.Elapsed.TotalSeconds),2)) seconds to execute..."
Write-Host $msg
$msg = $null
$StopWatch = $null
<# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>
