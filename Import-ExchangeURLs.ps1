<#
.SYNOPSIS
    This script imports a specially formatted CSV into server's configuration

.DESCRIPTION
    The CSV file to import must have the following headers:
    ServerName,"ServerVersion","EASInternalURL","EASExternalURL","OABInternalURL",
    "OABExernalURL","OWAInternalURL","OWAExernalURL","ECPInternalURL","ECPExernalURL",
    "AutoDiscName","AutoDiscURI","EWSInternalURL","EWSExernalURL","OutlookAnywhere-InternalHostName(NoneForE2010)",
    "OutlookAnywhere-ExternalHostNAme(E2010+)"
        NOTE: the ServerVersion is optional
        NOTE2: a blank value will set the corresponding attribute to $null

.PARAMETER InputVSB
    Specifies the CSV to input (will be validated in the script)

.PARAMETER CheckVersion
    This parameter will just dump the script current version.

.INPUTS
    CSV file

.OUTPUTS
    Set Exchange servers value - will warn if specified server in the CSV doesn't exist

.EXAMPLE
.\Do-Something.ps1
This will launch the script and do someting

.EXAMPLE
.\Do-Something.ps1 -CheckVersion
This will dump the script name and current version like :
SCRIPT NAME : Do-Something.ps1
VERSION : v1.0

.NOTES
None

.LINK
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comment_based_help?view=powershell-6

.LINK
    https://github.com/SammyKrosoft
#>
[CmdLetBinding(DefaultParameterSetName = "NormalRun")]
Param(
    [Parameter(Mandatory = $True, Position = 1, ParameterSetName = "NormalRun")][String]$InputCSV,
    [Parameter(Mandatory = $false, Position = 2, ParameterSetName = "CheckOnly")][switch]$CheckVersion
)

<# ------- SCRIPT_HEADER (Only Get-Help comments and Param() above this point) ------- #>
#Initializing a $Stopwatch variable to use to measure script execution
$stopwatch = [system.diagnostics.stopwatch]::StartNew()
#Using Write-Debug and playing with $DebugPreference -> "Continue" will output whatever you put on Write-Debug "Your text/values"
# and "SilentlyContinue" will output nothing on Write-Debug "Your text/values"
$DebugPreference = "Stop"
# Set Error Action to your needs
$ErrorActionPreference = "Stop"
#Script Version
$ScriptVersion = "0.1"
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
# $ErrorActionPreference = Continue
<# /DECLARATIONS #>
<# -------------------------- FUNCTIONS -------------------------- #>
function import-ValidCSV {
    <#
    .SYNOPSIS
        Imports a CSV file, validating that it has the columns specified in the -requiredColumns parameter.
        Throws a critical error if a required column is not found.

    .NOTES
        IMPORTANT INFORMATION:
        Please always mention the original authors of scripts or script extracts, it helps for traceability, and
        more importantly gives back to Caesar what belong to Caesar ;-)
        
            Author : Jason Coleman
            Role: Californian PowerShell and Virtualization Genius
            Link: https://virtuallyjason.blogspot.com/2016/08/import-validcsv-powershell-function.html

    .LINK
        https://virtuallyjason.blogspot.com/2016/08/import-validcsv-powershell-function.html
    #>

    param
        (
                [parameter(Mandatory=$true)]
                [ValidateScript({test-path $_ -type leaf})]
                [string]$inputFile,
                [string[]]$requiredColumns
        )
        $csvImport = import-csv $inputFile
        $inputTest = $csvImport | gm
        foreach ($requiredColumn in $requiredColumns)
        {
                if (!($inputTest | ? {$_.name -eq $requiredColumn}))
                {
                        Write-host "$inputFile is missing the $requiredColumn column" -BackgroundColor yellow -ForegroundColor red
                        exit 10
                }
        }
        $csvImport | ft
}
<# /FUNCTIONS #>
<# -------------------------- EXECUTIONS -------------------------- #>

$RequiredColumnsCollection = "ServerName","ServerVersion","EASInternalURL","EASExternalURL","OABInternalURL","OABExernalURL","OWAInternalURL","OWAExernalURL","ECPInternalURL","ECPExernalURL","AutoDiscURI","EWSInternalURL","EWSExernalURL","OutlookAnywhere-InternalHostName(NoneForE2010)","OutlookAnywhere-ExternalHostNAme(E2010+)"

# $ServerConfig = Import-Csv $InputCSV

import-ValidCSV -inputFile $InputCSV -requiredColumns $RequiredColumnsCollection

$ServersConfigs = import-csv $inputFile

Foreach ($server in $ServersConfigs) {
    Set-ActiveSyncVirtualDirectory -ActiveSyncServer E2016-01

    Write-Host "Getting Exchange server $($_.ServerName)"
    $CurrentServer = Get-ExchangeServer $_.ServerName

    Write-Host "Setting EAS InternalURL to $($_.EASInternalURL) and EAS ExternalURL to $_.EASExternalURL"
    $CurrentServer | Get-ActiveSyncVirtualDirectory -ADPropertiesOnly | Set-ActiveSyncVirtualDirectory -InternalURL $_.EASInternalURL -ExternalURL $_.EASExternalURL

    Write-Host "Setting OAB InternalURL to $($_.OABInternalURL) and OAB ExternalURL to $_.OABExternalURL"
    $CurrentServer | Get-OabVirtualDirectory -ADPropertiesOnly | Set-OabVirtualDirectory -InternalURL $_.OABInternalURL -ExternalUrl $_.OABExternalURL

    Write-Host "Setting EWS InternalURL to $($_.EWSInternalURL) and EWS ExternalURL to $_.EWSExternalURL"
    $CurrentServer | Get-EWSVirtualDirectory -ADPropertiesOnly | Set-EWSVirtualDirectory -InternalURL $_.EWSInternalURL -ExternalUrl $_.EWSExternalURL

    Write-Host "Setting ECP InternalURL to $($_.ECPInternalURL) and ECP ExternalURL to $_.ECPExternalURL"
    $CurrentServer | Get-ECPVirtualDirectory -ADPropertiesOnly | Set-ECPVirtualDirectory -InternalURL $_.ECPInternalURL -ExternalUrl $_.ECPExternalURL

    Write-Host "Setting EWS InternalURL to $($_.EWSInternalURL) and EWS ExternalURL to $_.EWSExternalURL"
    $CurrentServer | Get-WebServicesVirtualDirectory -ADPropertiesOnly | Set-WebServicesVirtualDirectory -InternalURL $_.EWSInternalURL -ExternalUrl $_.EWSExternalURL

    Write-Host "Setting OutlookAnywhere InternalURL to $($_.OutlookAnywhereInternalURL) and OutlookAnywhere ExternalURL to $_.OutlookAnywhereExternalURL"
    If ($CurrentServer.AdminDisplayVersion -match "15."){
        Write-Host "Server is E2013 or E2016, setting both OA Internal and External Host"
        $CurrentServer | Get-OutlookAnywhere -ADPropertiesOnly | Set-OutlookAnywhere -InternalHostName $_."OutlookAnywhere-InternalHostName(NoneForE2010)" -ExternalHostname $_."OutlookAnywhere-ExternalHostNAme(E2010+)"
    } Else {
        Write-Host "Server is E2010, setting only External Host"
        $CurrentServer | Get-OutlookAnywhere -ADPropertiesOnly | Set-OutlookAnywhere -ExternalHostname $_."OutlookAnywhere-ExternalHostNAme(E2010+)"
    }
    If ($CurrentServer.AdminDisplayVersion -match "15."){
        Set-ClientAccessService $CurrentServer -AutoDiscoverServiceInternalUri $_.AutodiscURI
    } Else {
        Set-ClientAccessServer $CurrentServer -AutoDiscoverServiceInternalUri $_.AutodiscURI
    }

}

<# /EXECUTIONS #>
<# -------------------------- CLEANUP VARIABLES -------------------------- #>

<# /CLEANUP VARIABLES#>
<# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...
$stopwatch.Stop()
$msg = "`n`nThe script took $([math]::round($($StopWatch.Elapsed.TotalSeconds),2)) seconds to execute..."
Write-Host $msg
$msg = $null
$StopWatch = $null
<# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>
