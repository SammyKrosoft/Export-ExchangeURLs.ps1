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
$DebugPreference = "Continue"
# Set Error Action to your needs
$ErrorActionPreference = "SilentlyContinue"
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
                        write-error "$inputFile is missing the $requiredColumn column"
                        exit 10
                }
        }
        $csvImport
}
<# /FUNCTIONS #>
<# -------------------------- EXECUTIONS -------------------------- #>


$ServerConfig = Import-Csv $InputCSV




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
