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
    [Parameter(Mandatory = $False, Position = 2, ParameterSetName = "NormalRun")][switch]$TestCSV,
    [Parameter(Mandatory = $false, Position = 3, ParameterSetName = "CheckOnly")][switch]$CheckVersion
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
$IsThereE2013orE2016 = $false
<# /DECLARATIONS #>
<# -------------------------- FUNCTIONS -------------------------- #>
Function Test-ExchTools(){
    <#
    .SYNOPSIS
    This small function will just check if you have Exchange tools installed or available on the
    current PowerShell session.
    
    .DESCRIPTION
    The presence of Exchange tools are checked by trying to execute "Get-ExBanner", one of the basic Exchange
    cmdlets that runs when the Exchange Management Shell is called.
    
    Just use Test-ExchTools in your script to make the script exit if not launched from an Exchange
    tools PowerShell session...
    
    .EXAMPLE
    Test-ExchTools
    => will exit the script/program si Exchange tools are not installed
    #>
        Try
        {
            #Get-command Get-ExBanner -ErrorAction Stop
            Get-command Get-Mailbox -ErrorAction Stop
            $ExchInstalledStatus = $true
            $Message = "Exchange tools are present !"
            Write-Host $Message -ForegroundColor Blue -BackgroundColor Red
        }
        Catch [System.SystemException]
        {
            $ExchInstalledStatus = $false
            $Message = "Exchange Tools are not present ! This script/tool need these. Exiting..."
            Write-Host $Message -ForegroundColor red -BackgroundColor Blue
            Exit
        }
        Return $ExchInstalledStatus
    }

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
        Return $csvImport
}
<# /FUNCTIONS #>
<# -------------------------- EXECUTIONS -------------------------- #>
If (!($TestCSV)){Test-ExchTools}

$RequiredColumnsCollection = "ServerName","ServerVersion","EASInternalURL","EASExternalURL","OABInternalURL","OABExernalURL","OWAInternalURL","OWAExernalURL","ECPInternalURL","ECPExernalURL","AutoDiscURI","EWSInternalURL","EWSExernalURL","OutlookAnywhere-InternalHostName(NoneForE2010)","OutlookAnywhere-ExternalHostNAme(E2010+)"

# $ServerConfig = Import-Csv $InputCSV

$ServersConfigs = import-ValidCSV -inputFile $InputCSV -requiredColumns $RequiredColumnsCollection
# $ServersConfigsDirectFromFile = Import-CSV $InputCSV

If($TestCSV){
    Write-Host "There are $($ServersConfigs.count) servers to parse on this CSV, here's the list:"
    $ServersConfigs | ft ServerName,@{Label = "Version"; Expression={"V." + $(($_.ServerVersion).Substring(8,4))}}
}

$test = ($ServersConfigs | % {$_.ServerVersion -match "15."}) -join ";"

If ($test -match "$true"){
    $IsThereE2013orE2016 = $True
} Else {
    $IsThereE2013orE2016 = $false
}

If ($IsThereE2013orE2016){
    Write-Host "There are some E2013/E2016 in the list... Make sure you run this tool from E2013/2016 EMS ! Using Get-ClientAccessServices instead of Get-ClientAccessServer" -BackgroundColor darkblue -fore red
} Else {
    Write-Host "No E2013/2016 in the list ... using Get-ClientAccessServer"
}

# $ServersConfigs = import-csv $inputFile

Foreach ($CurrentServer in $ServersConfigs) {

    Write-Host "Getting Exchange server $($CurrentServer.ServerName)"
    If (!$TestCSV){$CurrentServer = Get-ExchangeServer $CurrentServer.ServerName}

    Write-Host "Setting EAS InternalURL to $($CurrentServer.EASInternalURL) and EAS ExternalURL to $($CurrentServer.EASExternalURL)"
    If (!$TestCSV){$CurrentServer | Get-ActiveSyncVirtualDirectory -ADPropertiesOnly | Set-ActiveSyncVirtualDirectory -InternalURL $CurrentServer.EASInternalURL -ExternalURL $CurrentServer.EASExternalURL}

    Write-Host "Setting OAB InternalURL to $($CurrentServer.OABInternalURL) and OAB ExternalURL to $($CurrentServer.OABExternalURL)"
    If (!$TestCSV){$CurrentServer | Get-OabVirtualDirectory -ADPropertiesOnly | Set-OabVirtualDirectory -InternalURL $CurrentServer.OABInternalURL -ExternalUrl $CurrentServer.OABExternalURL}

    Write-Host "Setting EWS InternalURL to $($CurrentServer.EWSInternalURL) and EWS ExternalURL to $($CurrentServer.EWSExternalURL)"
    If (!$TestCSV){$CurrentServer | Get-EWSVirtualDirectory -ADPropertiesOnly | Set-EWSVirtualDirectory -InternalURL $CurrentServer.EWSInternalURL -ExternalUrl $CurrentServer.EWSExternalURL}

    Write-Host "Setting ECP InternalURL to $($CurrentServer.ECPInternalURL) and ECP ExternalURL to $($CurrentServer.ECPExternalURL)"
    If (!$TestCSV){$CurrentServer | Get-ECPVirtualDirectory -ADPropertiesOnly | Set-ECPVirtualDirectory -InternalURL $CurrentServer.ECPInternalURL -ExternalUrl $CurrentServer.ECPExternalURL}

    Write-Host "Setting EWS InternalURL to $($CurrentServer.EWSInternalURL) and EWS ExternalURL to $($CurrentServer.EWSExternalURL)"
    If (!$TestCSV){$CurrentServer | Get-WebServicesVirtualDirectory -ADPropertiesOnly | Set-WebServicesVirtualDirectory -InternalURL $CurrentServer.EWSInternalURL -ExternalUrl $CurrentServer.EWSExternalURL}

    Write-Host "Setting OutlookAnywhere InternalURL to $($CurrentServer.OutlookAnywhereInternalURL) and OutlookAnywhere ExternalURL to $($CurrentServer.OutlookAnywhereExternalURL)"
    If ($CurrentServer.AdminDisplayVersion -match "15."){
        Write-Host "Server is E2013 or E2016, setting both OA Internal and External Host"
        If (!$TestCSV){$CurrentServer | Get-OutlookAnywhere -ADPropertiesOnly | Set-OutlookAnywhere -InternalHostName $CurrentServer."OutlookAnywhere-InternalHostName(NoneForE2010)" -ExternalHostname $CurrentServer."OutlookAnywhere-ExternalHostNAme(E2010+)"}
    } Else {
        Write-Host "Server is E2010, setting only External Host"
        If (!$TestCSV){$CurrentServer | Get-OutlookAnywhere -ADPropertiesOnly | Set-OutlookAnywhere -ExternalHostname $CurrentServer."OutlookAnywhere-ExternalHostNAme(E2010+)"}
    }
    Write-Host "Setting Autodiscover URI (SCP) to $($CurrentServer.AutodiscURI)"
    If ($IsThereE2013orE2016){
        Write-Host "Using Get-ClientAccessService (assuming you run the script from an E2013/2016 EMS)" -ForegroundColor Yellow
        If (!$TestCSV){Set-ClientAccessService $CurrentServer -AutoDiscoverServiceInternalUri $CurrentServer.AutodiscURI}
    } Else {
        Write-Host "Using Get-ClientAccessServer (assuming you run the script from an 2010 EMS)" -ForegroundColor Yellow
        If (!$TestCSV){Set-ClientAccessServer $CurrentServer -AutoDiscoverServiceInternalUri $CurrentServer.AutodiscURI}
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
