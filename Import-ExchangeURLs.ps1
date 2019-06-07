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
    [Parameter(Mandatory = $False, Position = 3, ParameterSetName = "NormalRun")][switch]$DebugVerbose,
    [Parameter(Mandatory = $false, Position = 4, ParameterSetName = "CheckOnly")][switch]$CheckVersion
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
function IsNotEmpty($Param){
    If ($Param -ne "" -and $Param -ne $Null -and $Param -ne 0) {
        Return $True
    } Else {
        Return $False
    }
}
Function Title1 ($title, $TotalLength = 100, $Back = "Yellow", $Fore = "Black") {
    $TitleLength = $Title.Length
    [string]$StarsBeforeAndAfter = ""
    $RemainingLength = $TotalLength - $TitleLength
    If ($($RemainingLength % 2) -ne 0) {
        $Title = $Title + " "
    }
    $Counter = 0
    For ($i=1;$i -le $(($RemainingLength)/2);$i++) {
        $StarsBeforeAndAfter += "*"
        $counter++
    }
    
    $Title = $StarsBeforeAndAfter + $Title + $StarsBeforeAndAfter
    Write-host
    Write-Host $Title -BackgroundColor $Back -foregroundcolor $Fore
    Write-Host
    
}
Function LogMag ($Message){
    Write-Host $message -ForegroundColor Magenta
}

Function LogGreen ($message){
    Write-Host $message -ForegroundColor Green
}

Function LogYellow ($message){
    Write-Host $message -ForegroundColor Yellow
}

Function LogBlue ($message){
    Write-Host $message -ForegroundColor Blue
}

#Examples
# cls

# Title1 "Part 1 - Checking mailboxes"

# For ($i=0;$i -le 10;$i++){
#     LogGreen "Mailbox $i - ok"
# }

# Title1 "Part 2 - Checking databases"

# For ($i=0;$i -le 5;$i++){
#     LogGreen "Database $i - ok"
# }

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
#Trimming all values to ensure no leading or trailing spaces
$ServersConfigs | ForEach-Object {$_.PSObject.Properties | ForEach-Object {$_.Value = $_.Value.Trim()}}

If($TestCSV){
    Write-Host "There are $($ServersConfigs.count) servers to parse on this CSV, here's the list:"
    $ServersConfigs | ft ServerName,@{Label = "Version"; Expression={"V." + $(($_.ServerVersion).Substring(8,4))}}
}
#BOOKMARK
# exit

$test = ($ServersConfigs | % {$_.ServerVersion -match "15\."}) -join ";"

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

    Title1 "Getting Exchange server $($CurrentServer.ServerName)"
    # If we don't just test the script, we query the Server Object in another variable. It's only to test if the server is reachable
    # If server is not joigneable, we stop the script. If -TestCSV switch enabled, we don't test the server reachability, and we keep just 
    # the name string to build the command line...
    If (!$TestCSV){
        Try {
            Get-ExchangeServer $CurrentServer.ServerName -ErrorAction Stop
        }
        Catch{
            Write-Host "Server does not exist - please recheck your server names in your CSV, and remove the unexisting server names..."
            Exit
        }
    }
    
    If ($DebugVerbose){
        $CurrentServer | fl
    }

    # Exchange ActiveSync aka EAS
    $StatusMsg = "Setting EAS InternalURL to $($CurrentServer.EASInternalURL) and EAS ExternalURL to $($CurrentServer.EASExternalURL)"
    # Write-Host $StatusMsg -BackgroundColor Blue -ForegroundColor Red
    LogMag $StatusMsg
    $EAScmd = "$($CurrentServer.ServerName) | Get-ActiveSyncVirtualDirectory -ADPropertiesOnly | Set-ActiveSyncVirtualDirectory"
    #region #### VALUE TEST ROUTINE FOR DEBUG ######
    If ($DebugVerbose){
        LogGreen "Status of EAS InternalURL: "
        LogGreen "Value: $($CurrentServer.EASInternalURL)"
        LogGreen "Is it blank ?"
        LogGreen "$($CurrentServer.EASInternalURL -eq """")"
        LogGreen "Is it `$null ?"
        LogGreen "$($CurrentServer.EASInternalURL -eq $null)"
    }
    #endregion #### END OF TEST ROUTING FOR DEBUG ######
    If (IsNotEmpty $CurrentServer.EASInternalURL){
        If ($DebugVerbose){
            LogMag "EAS InternalURL is NOT null and is equal to $($CurrentServer.EASInternalURL)"
            LogMag "Length of characters : $($CurrentServer.EASInternalURL.Length)"
        }
        $EAScmd += " -InternalURL $($CurrentServer.EASInternalURL)"
    } Else {
        $EAScmd += " -InternalURL `$null"
    }
    #region #### VALUE TEST ROUTINE FOR DEBUG ######
    If ($DebugVerbose){
        If ($DebugVerbose){
            LogGreen "Status of EAS ExternalURL: "
            LogGreen "Value: $($CurrentServer.EASExternalURL)"
            LogGreen "Is it blank ?"
            LogGreen "$($CurrentServer.EASExternalURL -eq """")"
            LogGreen "Is it `$null ?"
            LogGreen "$($CurrentServer.EASExternalURL -eq $null)"
        }
    }
    #endregion #### END OF TEST ROUTING FOR DEBUG ######
    If (IsNotEmpty $CurrentServer.EASexternalURL){
        If ($DebugVerbose){
            LogMag "EAS ExternalURL is NOT null and is equal to $($CurrentServer.EASExternalURL)"
            LogMag "Length of characters : $($CurrentServer.EASExternalURL.Length)"
        }
        $EAScmd += " -ExternalURL $($CurrentServer.EASExternalURL)"
    } Else {
        If ($DebugVerbose){
            LogMag "EAS ExternalURL is null and is equal to $($CurrentServer.EASExternalURL)"
            LogMag "Length of characters : $($CurrentServer.EASExternalURL.Length)"
        }
        $EAScmd += " -ExternalURL `$null"
    }
    # If we have the -TestCSV switch enabled, we just print the generated command line. Otherwise, we run it using Invoke-Expression...
    If (!$TestCSV){
        Invoke-Expression $EAScmd
    } Else {
        Write-Host $EAScmd -BackgroundColor blue -ForegroundColor Yellow
    }

    # Exchange OfflineAddressBook 
    $StatusMsg = "Setting OAB InternalURL to $($CurrentServer.OABInternalURL) and OAB ExternalURL to $($CurrentServer.OABExternalURL)"
    # Write-Host $StatusMsg -BackgroundColor Blue -ForegroundColor Red
    LogMag $StatusMsg
    # $($CurrentServer.ServerName) | Get-OabVirtualDirectory -ADPropertiesOnly | Set-OabVirtualDirectory -InternalURL $CurrentServer.OABInternalURL -ExternalUrl $CurrentServer.OABExternalURL
    $OABCmd = "$($CurrentServer.ServerName) | Get-OabVirtualDirectory -ADPropertiesOnly | Set-OabVirtualDirectory"
    #region #### VALUE TEST ROUTINE FOR DEBUG ######
    If ($DebugVerbose){
        LogGreen "Status of OAB InternalURL: "
        LogGreen "Value: $($CurrentServer.OABInternalURL)"
        LogGreen "Is it blank ?"
        LogGreen "$($CurrentServer.OABInternalURL -eq """")"
        LogGreen "Is it `$null ?"
        LogGreen "$($CurrentServer.OABInternalURL -eq $null)"
    }
    #endregion #### END OF TEST ROUTING FOR DEBUG ######
    If ($CurrentServer.OABInternalURL -ne $null) {
        $OABcmd += " -InternalURL $($CurrentServer.OABInternalURL)"
    } Else {
        LogMag "OAB Internal URL is NULL and equal to $($CurrentServer.OABInternalURL)"
        $OABcmd += " -InternalURL `$null"
    }
    #region #### VALUE TEST ROUTINE FOR DEBUG ######
    If ($DebugVerbose){
        LogGreen "Status of OAB ExternalURL: "
        LogGreen "Value: $($CurrentServer.OABExternalURL)"
        LogGreen "Is it blank ?"
        LogGreen "$($CurrentServer.OABExternalURL -eq """")"
        LogGreen "Is it `$null ?"
        LogGreen "$($CurrentServer.OABExternalURL -eq $null)"
    }
    #endregion #### END OF TEST ROUTING FOR DEBUG ######
    If ($CurrentServer.OABExternalURL -ne $null) {
        $OABcmd += " -ExternalURL $($CurrentServer.OABExternalURL)"
    } Else {
        $OABcmd += " -ExternalURL `$null"
    }
    # If we have the -TestCSV switch enabled, we just print the generated command line. Otherwise, we run it using Invoke-Expression...
    If (!$TestCSV){
        Invoke-Expression $OABcmd
    } Else {
        Write-Host $OABcmd -BackgroundColor blue -ForegroundColor Yellow
    }

    # Outlook Web Access aka OWA
    $StatusMsg = "Setting OWA InternalURL to $($CurrentServer.OWAInternalURL) and OWA ExternalURL to $($CurrentServer.OWAExternalURL)"
    # Write-Host $StatusMsg -BackgroundColor Blue -ForegroundColor Red
    LogMag $StatusMsg
    $OWAcmd = "$($CurrentServer.ServerName) | Get-OWAVirtualDirectory -ADPropertiesOnly | Set-OWAVirtualDirectory"
    #region #### VALUE TEST ROUTINE FOR DEBUG ######
    If ($DebugVerbose){
        LogGreen "Status of OWA InternalURL: "
        LogGreen "Value: $($CurrentServer.OWAInternalURL)"
        LogGreen "Is it blank ?"
        LogGreen "$($CurrentServer.OWAInternalURL -eq """")"
        LogGreen "Is it `$null ?"
        LogGreen "$($CurrentServer.OWAInternalURL -eq $null)"
    }
    #endregion #### END OF TEST ROUTING FOR DEBUG ######
    If ($CurrentServer.OWAInternalURL -ne $null){
        $OWAcmd += " -InternalURL $($CurrentServer.OWAInternalURL)"
    } Else {
        $OWAcmd += " -InternalURL `$null"
    }
        #region #### VALUE TEST ROUTINE FOR DEBUG ######
        If ($DebugVerbose){
            LogGreen "Status of OWA InternalURL: "
            LogGreen "Value: $($CurrentServer.OWAInternalURL)"
            LogGreen "Is it blank ?"
            LogGreen "$($CurrentServer.OWAInternalURL -eq """")"
            LogGreen "Is it `$null ?"
            LogGreen "$($CurrentServer.OWAInternalURL -eq $null)"
        }
        #endregion #### END OF TEST ROUTING FOR DEBUG ######
    If ($CurrentServer.OWAExternalURL -ne $null){
        $OWAcmd += " -ExternalURL $($CurrentServer.OWAExternalURL)"
    } Else {
        $OWAcmd += " -ExternalURL `$null"
    }
    # If we have the -TestCSV switch enabled, we just print the generated command line. Otherwise, we run it using Invoke-Expression...
    If (!$TestCSV){
        Invoke-Expression $OWAcmd
    } Else {
        Write-Host $OWAcmd -BackgroundColor blue -ForegroundColor Yellow
    }

    # Outlook Web Access aka ECP
    $StatusMsg = "Setting ECP InternalURL to $($CurrentServer.ECPInternalURL) and ECP ExternalURL to $($CurrentServer.ECPExternalURL)"
    # Write-Host $StatusMsg -BackgroundColor Blue -ForegroundColor Red
    LogMag $StatusMsg
    $ECPcmd = "$($CurrentServer.ServerName) | Get-ECPVirtualDirectory -ADPropertiesOnly | Set-ECPVirtualDirectory"
    If ($CurrentServer.ECPInternalURL -ne $null){
        $ECPcmd += " -InternalURL $($CurrentServer.ECPInternalURL)"
    } Else {
        $ECPcmd += " -InternalURL `$null"
    }
    If ($CurrentServer.ECPExternalURL -ne $null){
        $ECPcmd += " -ExternalURL $($CurrentServer.ECPExternalURL)"
    } Else {
        $ECPcmd += " -ExternalURL `$null"
    }
    # If we have the -TestCSV switch enabled, we just print the generated command line. Otherwise, we run it using Invoke-Expression...
    If (!$TestCSV){
        Invoke-Expression $ECPcmd
    } Else {
        Write-Host $ECPcmd -BackgroundColor blue -ForegroundColor Yellow
    }

    # Exchange Exchange Web Services
    $StatusMsg = "Setting EWS InternalURL to $($CurrentServer.EWSInternalURL) and EWS ExternalURL to $($CurrentServer.EWSExternalURL)"
    # Write-Host $StatusMsg -BackgroundColor Blue -ForegroundColor Red
    LogMag $StatusMsg
    #$($CurrentServer.ServerName) | Get-WebServicesVirtualDirectory -ADPropertiesOnly | Set-WebServicesVirtualDirectory -InternalURL $CurrentServer.EWSInternalURL -ExternalUrl $CurrentServer.EWSExternalURL
    $EWSCmd = "$($CurrentServer.ServerName) | Get-WebServicesVirtualDirectory -ADPropertiesOnly | Set-WebServicesVirtualDirectory"
    If ($CurrentServer.EWSInternalURL -ne $null) {
        $EWScmd += " -InternalURL $($CurrentServer.EWSInternalURL)"
    } Else {
        $EWScmd += " -InternalURL `$null"
    }
    If ($CurrentServer.EWSExternalURL -ne $null) {
        $EWScmd += " -ExternalURL $($CurrentServer.EWSExternalURL)"
    } Else {
        $EWScmd += " -ExternalURL `$null"
    }
    # If we have the -TestCSV switch enabled, we just print the generated command line. Otherwise, we run it using Invoke-Expression...
    If (!$TestCSV){
        Invoke-Expression $EWScmd
    } Else {
        Write-Host $EWScmd -BackgroundColor blue -ForegroundColor Yellow
    }

    # Outlook Anywhere aka OA
    $StatusMsg = "Setting OutlookAnywhere InternalURL to $($CurrentServer."OutlookAnywhere-InternalHostName(NoneForE2010)") and OutlookAnywhere ExternalURL to $($CurrentServer."OutlookAnywhere-ExternalHostNAme(E2010+)")"
    # Write-Host $StatusMsg -BackgroundColor Blue -ForegroundColor Red
    LogMag $StatusMsg
    $OAcmd = "$($CurrentServer.ServerName) | Get-OutlookAnywhere -ADPropertiesOnly | Set-OutlookAnywhere"
    If ($CurrentServer.ServerVersion -match "15\."){
        If ($CurrentServer."OutlookAnywhere-InternalHostName(NoneForE2010)" -ne $null){
            $OAcmd += " -InternalHostName $($CurrentServer."OutlookAnywhere-InternalHostName(NoneForE2010)")"
        } Else {
            $OAcmd += " -InternalHostName `$null"
        }
    }
    If ($CurrentServer."OutlookAnywhere-ExternalHostNAme(E2010+)" -ne $null){
        $OAcmd += " -ExternalHostName $($CurrentServer."OutlookAnywhere-ExternalHostNAme(E2010+)")"
    } Else {
        $OAcmd += " -ExternalHostName `$null"
    }
    # If we have the -TestCSV switch enabled, we just print the generated command line. Otherwise, we run it using Invoke-Expression...
    If (!$TestCSV){
        Invoke-Expression $OAcmd
    } Else {
        Write-Host $OAcmd -BackgroundColor blue -ForegroundColor Yellow
    }

    # Autodiscover
    $StatusMsg = "Setting Autodiscover URI (SCP) to $($CurrentServer.AutodiscURI)"
    # Write-Host $StatusMsg -BackgroundColor Blue -ForegroundColor Red
    LogMag $StatusMsg
    If ($IsThereE2013orE2016){
        LogGreen "Using Get-ClientAccessService (assuming you run the script from an E2013/2016 EMS)"
        $SCPcmd = "Set-ClientAccessService $($CurrentServer.ServerName) -AutoDiscoverServiceInternalUri $($CurrentServer.AutodiscURI)"
    } Else {
        LogGreen "Using Get-ClientAccessServer (assuming you run the script from an 2010 EMS)" -ForegroundColor Yellow
        $SCPcmd = "Set-ClientAccessServer $($CurrentServer.ServerName) -AutoDiscoverServiceInternalUri $($CurrentServer.AutodiscURI)"
    }
    If (!$TestCSV){
        Invoke-Expression $SCPcmd
    } Else {
        Write-Host $SCPcmd -BackgroundColor blue -ForegroundColor Yellow
        
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
