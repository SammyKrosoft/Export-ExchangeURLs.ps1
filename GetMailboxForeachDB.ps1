$OutputFileName = "MailboxExportsSimple_$(Get-Date -F MMddyyyy-HHmmss).csv"
$Databases = Get-MailboxDatabase

$Dbprogresscounter = 0
$MailboxesCollection = @()
$DatabasesCount = $Databases.Count
Foreach ($database in $Databases) {
    write-progress -Activity "Parsing databases" -Status "Now in database $($database.Name) ..." -PercentComplete $($Dbprogresscounter/$DatabasesCount*100)
    $Mailboxes = $null
    $Mailboxes = Get-Mailbox -ResultSize Unlimited -Database $Database -Filter {RecipientTypeDetails -ne "DiscoveryMailbox"}| Select Name,PrimarySMTPAddress,REcipientTypeDetails
    Write-Host "Found $($Mailboxes.count) mailboxes on database $($Database.name) ..." -ForegroundColor Green
    $MailboxesCollection +=$Mailboxes
    $Dbprogresscounter++
}

$MailboxesCollection | Export-CSV "$($env:USERPROFILE)\Documents\$OutputFileName" -NoTypeInformation -Encoding UTF8