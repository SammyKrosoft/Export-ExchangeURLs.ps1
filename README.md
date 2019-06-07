# Export-ExchangeURLsv3.ps1
Export Exchange URLs to a CSV file, for backup purposes, or to modify the URLs / AutoDiscoverServiceInternalURI aka SCP aka Service Connection Point and import back the information on your Exchange servers using the below script

# Import-ExchangeURLs.ps1
Import Exchange URLs from a CSV issued by the above command and modified if need be to set the correct URLs / Autodiscover / SCP values on one or more of your Exchange Servers

# Get-Help
Use ```Get-Help .\Import-ExchangeURLs.ps1``` to get the script description and syntax
Use ```Get-Help .\Import-ExchangeURLs.ps1 -Examples``` to get the script examples sections
Use ```Get-Help .\Import-ExchangeURLs.ps1 -Examples``` to get the script's full help

# Help dump
I'm dumping temporarily the Import-ExchangeURLs.ps1 script help here for quick reference:
```
NAME
    .\Import-ExchangeURLs.ps1

SYNOPSIS
    This script will help you set up your Exchange Virtual Directories, including AutoDiscoverfrom a CSV file
    it takes in Input. This CSV file can contain one or as many servers as you need to configure the
    Virtual Directories for.


SYNTAX
    .\Import-ExchangeURLs.ps1 [-InputCSV] <String> [[-TestCSV]] [[-DebugVerbose]]
    [<CommonParameters>]

    .\Import-ExchangeURLs.ps1 [[-CheckVersion]] [<CommonParameters>]


DESCRIPTION
    This script takes in input a CSV file issued form an Export-ExchangeURLsv3.ps1 script or a
    CSV file containing the following headers:

    ServerName,ServerVersion,EASInternalURL,EASExternalURL,OABInternalURL,OABExernalURL,
    OWAInternalURL,OWAExernalURL,ECPInternalURL,ECPExernalURL,AutoDiscName,AutoDiscURI,
    EWSInternalURL,EWSExernalURL,OutlookAnywhere-InternalHostName(NoneForE2010),
    OutlookAnywhere-ExternalHostNAme(E2010+)

    IMPORTANT NOTE: a blank value will set the corresponding attribute to $null

    IMPORTANT ADVICE: Always use Export-ExchangeURLsv3.ps1 to export all your current URLs before any
    modifications => that way you'll just have to use that export with this script to set
    the Virtual Directories and/or Autodiscover value(s) to what they were initially.


PARAMETERS
    -InputCSV <String>
        Specifies the CSV to input (will be validated in the script)

        Required?                    true
        Position?                    2
        Default value
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -TestCSV [<SwitchParameter>]
        This switch will just get all the values in the CSV file specified in the InputCSV property, and
        print on screen all the actions that the script will perform without the -TestCSV switch.

        Required?                    false
        Position?                    3
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -DebugVerbose [<SwitchParameter>]
        This swich Will enable output of additional details regarding the attributes values and test whether
        it's set to $null or empty string ""

        Required?                    false
        Position?                    4
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters?  false

    -CheckVersion [<SwitchParameter>]
        This parameter will just dump the script current version.

        Required?                    false
        Position?                    5
        Default value                False
        Accept pipeline input?       false
        Accept wildcard characters?  false

    <CommonParameters>
        This cmdlet supports the common parameters: Verbose, Debug,
        ErrorAction, ErrorVariable, WarningAction, WarningVariable,
        OutBuffer, PipelineVariable, and OutVariable. For more information, see
        about_CommonParameters (https:/go.microsoft.com/fwlink/?LinkID=113216).

INPUTS
    CSV file containing the above mentionned headers.


OUTPUTS
    Set Exchange servers value - the script will stop if specified server in the CSV doesn't exist


NOTES


        Again, it's strongly recommended to:
        #1 - Export your current URLs and Autodiscover settings using the Export-ExchangeURLsv3.ps1 script to be able to easily roll back
        if need be, and keep the original CSV export. Create a copy of that exported CSV that we will modify and use that modified copy
        with the script to set the Exchange vDir properties.
        #2 - Always run the Import-ExchangeURLs.ps1 script with the -TestCSV first, and review all the command lines that the script will
        execute
        #3 - the -DebugVerbose switch is not mandatory, it's mostly for debug purposes if the script crashes or doesn't behave as intended

    -------------------------- EXAMPLE 1 --------------------------

    PS C:\>.\Import-ExchangeURLs.ps1 -CheckVersion

    This will dump the script name and current version like :
    SCRIPT NAME : Import-ExchangeURLs.ps1
    VERSION : v1.0




    -------------------------- EXAMPLE 2 --------------------------

    PS C:\>.\Import-ExchangeURLs.ps1 -InputCSV .\ServersConfig.csv -TestCSV

    This will launch the script and print only without executing the PowerShell command lines it will execute to update
    the Exchange Virtual Directories according to the information provided in the ServersConfig.csv file
    specified on the -InputCSV parameter.




    -------------------------- EXAMPLE 3 --------------------------

    PS C:\>.\Import-ExchangeURLs.ps1 -InputCSV .\ServersConfig.csv -DebugVerbose -TestCSV

    This will launch the script and print only without executing the PowerShell command lines it will execute to update
    the Exchange Virtual Directories according to the information provided in the ServersConfig.csv file
    specified on the -InputCSV parameter, as well as output additional details about he actions, comparisons, and some other stuff the script does...




    -------------------------- EXAMPLE 4 --------------------------

    PS C:\>.\Import-ExchangeURLs.ps1 -InputCSV .\ServersConfig.csv

    This will set all Exchange Virtual Directories (OAB, EWS, OWA, ECP, ...) including Autodiscover SCP
    with the URLs present in the CSV file. If a value is blank for a Virtual Directory property, the script
    will set it to a $null aka blank value.




    -------------------------- EXAMPLE 5 --------------------------

    PS C:\>.\Import-ExchangeURLs.ps1 -InputCSV .\ServersConfig.csv -DebugVerbose

    This will launch the script, set the Virtual Directories values according to the specified CSV, and
    output additional details about the actions, comparisons, and some other stuff the script does...





RELATED LINKS
    https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_comment_based_help?view=powershell-6
    https://github.com/SammyKrosoft
```


# Recent changes
Moved Exchange 2007 compatible Export script to the Archive section, since Exchange 2007 is not supported anymore, I won't maintain the Exchange 2007 version of the script.

