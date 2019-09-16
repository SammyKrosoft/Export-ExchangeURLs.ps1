<#PSScriptInfo

.VERSION 3.3.2

.GUID 0a1b89dc-e2b3-4e34-b1ad-e86ca7f6833d

.AUTHOR Sam Drey

.COMPANYNAME Microsoft Canada

#> 

<#
.SYNOPSIS
    This script will help you set up your Exchange Virtual Directories, including AutoDiscoverfrom a CSV file
    it takes in Input. This CSV file can contain one or as many servers as you need to configure the 
    Virtual Directories for.

.DESCRIPTION
    This script takes in input a CSV file issued form an Export-ExchangeURLsv3.ps1 script or a
    CSV file containing the following headers:
    
    ServerName,ServerVersion,EASInternalURL,EASExternalURL,OABInternalURL,OABExternalURL,
    OWAInternalURL,OWAExternalURL,ECPInternalURL,ECPExternalURL,AutoDiscName,AutoDiscURI,
    EWSInternalURL,EWSExternalURL,OutlookAnywhere-InternalHostName(NoneForE2010),
    OutlookAnywhere-ExternalHostNAme(E2010+)
    
    IMPORTANT NOTE: a blank value will set the corresponding attribute to $null

    IMPORTANT ADVICE: Always use Export-ExchangeURLsv3.ps1 to export all your current URLs before any
    modifications => that way you'll just have to use that export with this script to set
    the Virtual Directories and/or Autodiscover value(s) back to what they were initially.

.PARAMETER InputCSV
    Specifies the CSV to input (will be validated in the script)

.PARAMETER GenerateCommandsOnly 
    This switch will just get all the values in the CSV file specified in the InputCSV property, and 
    print on screen all the actions that the script will perform without the -GenerateCommandsOnly switch.

.PARAMETER DebugVerbose
    This swich Will enable output of additional details regarding the attributes values and test whether
    it's set to $null or empty string ""

.PARAMETER CheckVersion
    This parameter will just dump the script current version.

.INPUTS
    CSV file containing the above mentionned headers.

.OUTPUTS
    Set Exchange servers value - the script will stop if specified server in the CSV doesn't exist

.EXAMPLE
.\Import-ExchangeURLs.ps1 -CheckVersion
This will dump the script name and current version like :
SCRIPT NAME : Import-ExchangeURLs.ps1
VERSION : v1.0

.EXAMPLE
.\Import-ExchangeURLs.ps1 -InputCSV .\ServersConfig.csv -GenerateCommandsOnly
This will launch the script and print only without executing the PowerShell command lines it will execute to update
the Exchange Virtual Directories according to the information provided in the ServersConfig.csv file
specified on the -InputCSV parameter.

.EXAMPLE
.\Import-ExchangeURLs.ps1 -InputCSV .\ServersConfig.csv -DebugVerbose -GenerateCommandsOnly
This will launch the script and print only without executing the PowerShell command lines it will execute to update
the Exchange Virtual Directories according to the information provided in the ServersConfig.csv file
specified on the -InputCSV parameter, as well as output additional details about he actions, comparisons, and some other stuff the script does...

.EXAMPLE
.\Import-ExchangeURLs.ps1 -InputCSV .\ServersConfig.csv
This will set all Exchange Virtual Directories (OAB, EWS, OWA, ECP, ...) including Autodiscover SCP 
with the URLs present in the CSV file. If a value is blank for a Virtual Directory property, the script
will set it to a $null aka blank value.

.EXAMPLE
.\Import-ExchangeURLs.ps1 -InputCSV .\ServersConfig.csv -DebugVerbose
This will launch the script, set the Virtual Directories values according to the specified CSV, and
output additional details about the actions, comparisons, and some other stuff the script does...

.NOTES
Again, it's strongly recommended to:
#1 - Export your current URLs and Autodiscover settings using the Export-ExchangeURLs.ps1 script to be able to easily roll back
if need be, and keep the original CSV export. Create a copy of that exported CSV that we will modify and use that modified copy
with the script to set the Exchange vDir properties.
#2 - Always run the Import-ExchangeURLs.ps1 script with the -GenerateCommandsOnly first, and review all the command lines that the script will
execute
#3 - the -DebugVerbose switch is not mandatory, it's mostly for debug purposes if the script crashes or doesn't behave as intended

.LINK
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comment_based_help?view=powershell-6

.LINK
    https://github.com/SammyKrosoft
#>
[CmdLetBinding(DefaultParameterSetName = "NormalRun")]
Param(
    [Parameter(Mandatory = $True, Position = 1, ParameterSetName = "NormalRun")][String]$InputCSV,
    [Parameter(Mandatory = $False, Position = 2, ParameterSetName = "NormalRun")][switch]$GenerateCommandsOnly,
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
$ScriptVersion = "3.3.2"
<# Version changes
v3.3.2 - jumped version to match and sync with Export-ExchangeURLs.ps1
v1.6.1 - changed author to Sam Drey and current company
v1.6.0 : added MAPI vdir update
v1.5 : Fixed typo in ExternalURL fields (was ExernalURL without the "t")
v1.4 : changed Get-ClientAccessService only for Exchange 2016 (NOT for Exchange 2013 - because E2013 still can have CAS Server separated, not E2016, hence change of cmdlet name)
v1.3: renamed -TestCSV switch to -GenerateCommandsOnly
v1.1 -> v1.2 : added # character on output text to enable users to use generated scripts when using -GenerateCommandsOnly instead of letting the script to set all URLs
v1.0-> v1.1 : fixed Set-OutlookAnywhere, added -Internal/ExternalClientsRequireSSL when server is Exchange 2013/2016, added usage of IsNotEmpty function
instead of comparison with $null as sometimes blank CSV cells is reported as empty string, sometimes as $null
v0.1 -> v1 : finalized version. To be fixed next : output text when setting a property to "$null". This is cosmetic minor change
that does not impact the script purpose and actions.
v0.1 : first script version
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
$IsThereE2016 = $false
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
    
    $Title = "# " + $StarsBeforeAndAfter + $Title + $StarsBeforeAndAfter
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
If (!($GenerateCommandsOnly)){Test-ExchTools}

$RequiredColumnsCollection = "ServerName","ServerVersion","EASInternalURL","EASExternalURL","OABInternalURL","OABExternalURL","OWAInternalURL","OWAExternalURL","ECPInternalURL","ECPExternalURL","AutoDiscURI","EWSInternalURL","EWSExternalURL","OutlookAnywhere-InternalHostName(NoneForE2010)","OutlookAnywhere-ExternalHostNAme(E2010+)"

# $ServerConfig = Import-Csv $InputCSV

$ServersConfigs = import-ValidCSV -inputFile $InputCSV -requiredColumns $RequiredColumnsCollection
#Trimming all values to ensure no leading or trailing spaces
$ServersConfigs | ForEach-Object {$_.PSObject.Properties | ForEach-Object {If(IsNotEmpty $_.Value){$_.Value = $_.Value.Trim()}}}

If($GenerateCommandsOnly){
    Write-Host "There are $($ServersConfigs.count) servers to parse on this CSV, here's the list:"
    $ServersConfigs | ft ServerName,@{Label = "Version"; Expression={"V." + $(($_.ServerVersion).Substring(8,4))}}
}
#BOOKMARK
# exit

$test = ($ServersConfigs | % {$_.ServerVersion -match "15\.1"}) -join ";"

If ($test -match "$true"){
    $IsThereE2016 = $True
} Else {
    $IsThereE2016 = $false
}

If ($IsThereE2016){
    Write-Host "There are some E2016 in the list... Make sure you run this tool from 2016 EMS ! Using Get-ClientAccessServices instead of Get-ClientAccessServer" -BackgroundColor darkblue -fore red
} Else {
    Write-Host "No 2016 in the list ... using Get-ClientAccessServer"
}

# $ServersConfigs = import-csv $inputFile

Foreach ($CurrentServer in $ServersConfigs) {

    Title1 "Getting Exchange server $($CurrentServer.ServerName)"
    # If we don't just test the script, we query the Server Object in another variable. It's only to test if the server is reachable
    # If server is not joigneable, we stop the script. If -GenerateCommandsOnly switch enabled, we don't test the server reachability, and we keep just 
    # the name string to build the command line...
    If (!$GenerateCommandsOnly){
        Try {
            Get-ExchangeServer $CurrentServer.ServerName -ErrorAction Stop | Select Name,domain,Site,ServerRole
        }
        Catch{
            Write-Host "Server does not exist - please recheck your server names in your CSV, and remove the unexisting server names..."
            Exit
        }
    }
    
    If ($DebugVerbose){
        $CurrentServer
    }

    # Exchange ActiveSync aka EAS
    $StatusMsg = "# Setting EAS InternalURL to $($CurrentServer.EASInternalURL) and EAS ExternalURL to $($CurrentServer.EASExternalURL)"
    # Write-Host $StatusMsg -BackgroundColor Blue -ForegroundColor Red
    LogMag $StatusMsg
    $EAScmd = "Get-ActiveSyncVirtualDirectory -Server $($CurrentServer.ServerName) -ADPropertiesOnly | Set-ActiveSyncVirtualDirectory"
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
    # If we have the -GenerateCommandsOnly switch enabled, we just print the generated command line. Otherwise, we run it using Invoke-Expression...
    If (!$GenerateCommandsOnly){
        Invoke-Expression $EAScmd
    } Else {
        Write-Host $EAScmd -BackgroundColor blue -ForegroundColor Yellow
    }

    # Exchange OfflineAddressBook 
    $StatusMsg = "# Setting OAB InternalURL to $($CurrentServer.OABInternalURL) and OAB ExternalURL to $($CurrentServer.OABExternalURL)"
    # Write-Host $StatusMsg -BackgroundColor Blue -ForegroundColor Red
    LogMag $StatusMsg
    # $($CurrentServer.ServerName) | Get-OabVirtualDirectory -ADPropertiesOnly | Set-OabVirtualDirectory -InternalURL $CurrentServer.OABInternalURL -ExternalUrl $CurrentServer.OABExternalURL
    $OABCmd = "Get-OabVirtualDirectory -Server $($CurrentServer.ServerName) -ADPropertiesOnly | Set-OabVirtualDirectory"
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
    If (IsNotEmpty $CurrentServer.OABInternalURL) {
        $OABcmd += " -InternalURL $($CurrentServer.OABInternalURL)"
    } Else {
        LogMag "OAB Internal URL is NULL and equal to $($CurrentServer.OABInternalURL)"
        $OABcmd += " -InternalURL `$null"
    }
    #region #### VALUE TEST ROUTINE FOR DEBUG ######
    If ($DebugVerbose){
        LogGreen "Status of OAB ExternalURL:"
        LogGreen "Value: $($CurrentServer.OABExternalURL)"
        LogGreen "Is it blank ?"
        LogGreen "$($CurrentServer.OABExternalURL -eq """")"
        LogGreen "Is it `$null ?"
        LogGreen "$($CurrentServer.OABExternalURL -eq $null)"
    }
    #endregion #### END OF TEST ROUTING FOR DEBUG ######
    If (IsNotEmpty $CurrentServer.OABExternalURL) {
        $OABcmd += " -ExternalURL $($CurrentServer.OABExternalURL)"
    } Else {
        $OABcmd += " -ExternalURL `$null"
    }
    # If we have the -GenerateCommandsOnly switch enabled, we just print the generated command line. Otherwise, we run it using Invoke-Expression...
    If (!$GenerateCommandsOnly){
        Invoke-Expression $OABcmd
    } Else {
        Write-Host $OABcmd -BackgroundColor blue -ForegroundColor Yellow
    }

    # Outlook Web Access aka OWA
    $StatusMsg = "# Setting OWA InternalURL to $($CurrentServer.OWAInternalURL) and OWA ExternalURL to $($CurrentServer.OWAExternalURL)"
    # Write-Host $StatusMsg -BackgroundColor Blue -ForegroundColor Red
    LogMag $StatusMsg
    $OWAcmd = "Get-OWAVirtualDirectory -Server $($CurrentServer.ServerName) -ADPropertiesOnly | Set-OWAVirtualDirectory"
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
    If (IsNotEmpty $CurrentServer.OWAInternalURL){
        $OWAcmd += " -InternalURL $($CurrentServer.OWAInternalURL)"
    } Else {
        $OWAcmd += " -InternalURL `$null"
    }
    #region #### VALUE TEST ROUTINE FOR DEBUG ######
    If ($DebugVerbose){
        LogGreen "Status of OWA ExternalURL: "
        LogGreen "Value: $($CurrentServer.OWAExternalURL)"
        LogGreen "Is it blank ?"
        LogGreen "$($CurrentServer.OWAExternalURL -eq """")"
        LogGreen "Is it `$null ?"
        LogGreen "$($CurrentServer.OWAExternalURL -eq $null)"
    }
    #endregion #### END OF TEST ROUTING FOR DEBUG ######
    If (IsNotEmpty $CurrentServer.OWAExternalURL){
        $OWAcmd += " -ExternalURL $($CurrentServer.OWAExternalURL)"
    } Else {
        $OWAcmd += " -ExternalURL `$null"
    }
    # If we have the -GenerateCommandsOnly switch enabled, we just print the generated command line. Otherwise, we run it using Invoke-Expression...
    If (!$GenerateCommandsOnly){
        Invoke-Expression $OWAcmd
    } Else {
        Write-Host $OWAcmd -BackgroundColor blue -ForegroundColor Yellow
    }

    # Outlook Web Access aka ECP
    $StatusMsg = "# Setting ECP InternalURL to $($CurrentServer.ECPInternalURL) and ECP ExternalURL to $($CurrentServer.ECPExternalURL)"
    # Write-Host $StatusMsg -BackgroundColor Blue -ForegroundColor Red
    LogMag $StatusMsg
    $ECPcmd = "Get-ECPVirtualDirectory -Server $($CurrentServer.ServerName) -ADPropertiesOnly | Set-ECPVirtualDirectory"
    If (IsNotEmpty $CurrentServer.ECPInternalURL){
        $ECPcmd += " -InternalURL $($CurrentServer.ECPInternalURL)"
    } Else {
        $ECPcmd += " -InternalURL `$null"
    }
    If (IsNotEmpty $CurrentServer.ECPExternalURL){
        $ECPcmd += " -ExternalURL $($CurrentServer.ECPExternalURL)"
    } Else {
        $ECPcmd += " -ExternalURL `$null"
    }
    # If we have the -GenerateCommandsOnly switch enabled, we just print the generated command line. Otherwise, we run it using Invoke-Expression...
    If (!$GenerateCommandsOnly){
        Invoke-Expression $ECPcmd
    } Else {
        Write-Host $ECPcmd -BackgroundColor blue -ForegroundColor Yellow
    }

    # Exchange Exchange Web Services
    $StatusMsg = "# Setting EWS InternalURL to $($CurrentServer.EWSInternalURL) and EWS ExternalURL to $($CurrentServer.EWSExternalURL)"
    # Write-Host $StatusMsg -BackgroundColor Blue -ForegroundColor Red
    LogMag $StatusMsg
    #$($CurrentServer.ServerName) | Get-WebServicesVirtualDirectory -ADPropertiesOnly | Set-WebServicesVirtualDirectory -InternalURL $CurrentServer.EWSInternalURL -ExternalUrl $CurrentServer.EWSExternalURL
    $EWSCmd = "Get-WebServicesVirtualDirectory -Server $($CurrentServer.ServerName) -ADPropertiesOnly | Set-WebServicesVirtualDirectory"
    If (IsNotEmpty $CurrentServer.EWSInternalURL) {
        $EWScmd += " -InternalURL $($CurrentServer.EWSInternalURL)"
    } Else {
        $EWScmd += " -InternalURL `$null"
    }
    If (IsNotEmpty $CurrentServer.EWSExternalURL) {
        $EWScmd += " -ExternalURL $($CurrentServer.EWSExternalURL)"
    } Else {
        $EWScmd += " -ExternalURL `$null"
    }
    # If we have the -GenerateCommandsOnly switch enabled, we just print the generated command line. Otherwise, we run it using Invoke-Expression...
    If (!$GenerateCommandsOnly){
        Invoke-Expression $EWScmd
    } Else {
        Write-Host $EWScmd -BackgroundColor blue -ForegroundColor Yellow
    }

    # Outlook Anywhere aka OA
    $StatusMsg = "# Setting OutlookAnywhere InternalURL to $($CurrentServer."OutlookAnywhere-InternalHostName(NoneForE2010)") and OutlookAnywhere ExternalURL to $($CurrentServer."OutlookAnywhere-ExternalHostNAme(E2010+)")"
    # Write-Host $StatusMsg -BackgroundColor Blue -ForegroundColor Red
    LogMag $StatusMsg
    $OAcmd = "Get-OutlookAnywhere -Server $($CurrentServer.ServerName) -ADPropertiesOnly | Set-OutlookAnywhere"
    # Exchange 2010 does NOT have any  InternalHohstNAme ... just setting InternalHostName if it's Exchange 2013/2016/2019
    If ($CurrentServer.ServerVersion -match "15\."){
        If (IsNotEmpty $CurrentServer."OutlookAnywhere-InternalHostName(NoneForE2010)"){
            $OAcmd += " -InternalHostName $($CurrentServer."OutlookAnywhere-InternalHostName(NoneForE2010)") -InternalClientsRequireSsl `$true"
        } Else {
            $OAcmd += " -InternalHostName `$null"
        }
    }
    If (IsNotEmpty $CurrentServer."OutlookAnywhere-ExternalHostNAme(E2010+)"){
        If ($CurrentServer.ServerVersion -match "15\."){
        $OAcmd += " -ExternalHostName $($CurrentServer."OutlookAnywhere-ExternalHostNAme(E2010+)") -ExternalClientsRequireSsl `$true -ExternalClientAuthenticationMethod Negotiate  -IISAuthenticationMethods @('Ntlm', 'Basic', 'Negotiate')"
        } Elseif ($CurrentServer.ServerVersion -match "14\."){
            # If Exchange 2010 server, then we set ExternalHostName, but without -ExternalClientsRequireSSL switch (that switch is for E2013/2016/2019 only)
            $OAcmd += " -ExternalHostName $($CurrentServer."OutlookAnywhere-ExternalHostNAme(E2010+)") -ExternalClientAuthenticationMethod NTLM"
        }
    } Else {
        $OAcmd += " -ExternalHostName `$null"
    }
    # If we have the -GenerateCommandsOnly switch enabled, we just print the generated command line. Otherwise, we run it using Invoke-Expression...
    If (!$GenerateCommandsOnly){
        Invoke-Expression $OAcmd
    } Else {
        Write-Host $OAcmd -BackgroundColor blue -ForegroundColor Yellow
    }

    # Autodiscover
    $StatusMsg = "# Setting Autodiscover URI (SCP) to $($CurrentServer.AutodiscURI)"
    # Write-Host $StatusMsg -BackgroundColor Blue -ForegroundColor Red
    LogMag $StatusMsg
    If (IsNotEmpty $CurrentServer.AutoDiscURI){
        If ($IsThereE2016){
            LogGreen "# Using Get-ClientAccessService (assuming you run the script from an 2016 EMS)"
            $SCPcmd = "Set-ClientAccessService $($CurrentServer.ServerName) -AutoDiscoverServiceInternalUri $($CurrentServer.AutodiscURI)"
        } Else {
            LogGreen "# Using Get-ClientAccessServer (assuming you run the script from an 2010 EMS)" -ForegroundColor Yellow
            $SCPcmd = "Set-ClientAccessServer $($CurrentServer.ServerName) -AutoDiscoverServiceInternalUri $($CurrentServer.AutodiscURI)"
        }
        If (!$GenerateCommandsOnly){
            Invoke-Expression $SCPcmd
        } Else {
            Write-Host $SCPcmd -BackgroundColor blue -ForegroundColor Yellow
            
        }
    } Else {
        LogGreen "# The CSV input had a blank value for the Autodiscover URI - we won't set it to `$null, instead we just don't touch it"
        LogMag "# Otherwise we can just replace that line of the script with `$SCPcmd = `"Set-ClientAccessService `$(`$CurrentServer.ServerName) -AutodiscoverServiceInternalURI `$(`$CurrentServer.AutodiscURI)`""
        LogMag "# and Invoke-Expression `$SCPcmd without the -TestCSV switch, or just dump `$SCPcmd with the -TestCSV switch"
    }

    # MAPI over HTTP
    If (IsNotEmpty $($CurrentServer.MAPIInternalURL)){
        If ($CurrentServer.ServerVersion -match "15\.") {
            $StatusMsg = "# Setting MAPI InternalURL to $($CurrentServer.MAPIInternalURL) and MAPI ExternalURL to $($CurrentServer.MAPIExternalURL)"
            # Write-Host $StatusMsg -BackgroundColor Blue -ForegroundColor Red
            LogMag $StatusMsg
            $MAPIcmd = "Get-MAPIVirtualDirectory -Server $($CurrentServer.ServerName) -ADPropertiesOnly | Set-MAPIVirtualDirectory"
            #region #### VALUE TEST ROUTINE FOR DEBUG ######
            If ($DebugVerbose){
                LogGreen "Status of MAPI InternalURL: "
                LogGreen "Value: $($CurrentServer.MAPIInternalURL)"
                LogGreen "Is it blank ?"
                LogGreen "$($CurrentServer.MAPIInternalURL -eq """")"
                LogGreen "Is it `$null ?"
                LogGreen "$($CurrentServer.MAPIInternalURL -eq $null)"
            }
            #endregion #### END OF TEST ROUTING FOR DEBUG ######
            If (IsNotEmpty $CurrentServer.MAPIInternalURL){
                $MAPIcmd += " -InternalURL $($CurrentServer.MAPIInternalURL)"
            } Else {
                $MAPIcmd += " -InternalURL `$null"
            }
            #region #### VALUE TEST ROUTINE FOR DEBUG ######
            If ($DebugVerbose){
                LogGreen "Status of MAPI ExternalURL: "
                LogGreen "Value: $($CurrentServer.MAPIExternalURL)"
                LogGreen "Is it blank ?"
                LogGreen "$($CurrentServer.MAPIExternalURL -eq """")"
                LogGreen "Is it `$null ?"
                LogGreen "$($CurrentServer.MAPIExternalURL -eq $null)"
            }
            #endregion #### END OF TEST ROUTING FOR DEBUG ######
            If (IsNotEmpty $CurrentServer.MAPIExternalURL){
                $MAPIcmd += " -ExternalURL $($CurrentServer.MAPIExternalURL)"
            } Else {
                $MAPIcmd += " -ExternalURL `$null"
            }
            # If we have the -GenerateCommandsOnly switch enabled, we just print the generated command line. Otherwise, we run it using Invoke-Expression...
            If (!$GenerateCommandsOnly){
                Invoke-Expression $MAPIcmd
            } Else {
                Write-Host $MAPIcmd -BackgroundColor blue -ForegroundColor Yellow
            }
        } Else {
            Write-Host "Not an Exchange 2013 or 2016 server - skipping MAPI over HTTP setup ;-)"
        }
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
