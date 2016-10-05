<#

.NAME Exchange 2010 Top 20 Largest Mailboxes

.SYNOPSIS This power shell script gathers the top 20 mailbox sizes in an Exchange 2010 environment.

.DESCRIPTION The top portion of the script gathers the top 20 largest mailboxes in an exchange 2010 environment. The bottom portion formats the results in HTML and sends them out in an email.

>

[CmdletBinding()] param( [parameter(mandatory=$false)] [switch]$sendEmail )

.................................

Add the Microsoft Exchange PSSnapin

.................................

if (!(Get-PSSnapin | where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"})) { Write-Verbose "Loading the Exchange snapin" Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue . $env:ExchangeInstallPath\bin\RemoteExchange.ps1 Connect-ExchangeServer -auto -AllowClobber }

.................................

Initialize the variables

.................................

$mailboxDatabase = @(Get-MailboxDatabase)

$mailboxReport = @()

$now = Get-Date -Format F

.................................

SMTP Settings

.................................

$smtpsettings = @{ To = "dylan.macdermot@usafa.edu"#, "Liana.Jones@usafa.edu", "sabrina.vanderree@us.af.mil", "dylan.mac_dermot.2@us.af.mil", "steven.hunt@usafa.edu" From = "exchangeadministrators@usafa.edu" Subject = "Exchange Top 20 Report - $now" SmtpServer = "afaeduexch-02.usafa.ds.af.edu" }

.................................

Script

.................................

if(!$usafaMailboxes) { $usafaMailboxes = $null Write-Host -ForegroundColor Yellow "Gatherling Mailboxes..." $usafaMailboxes = Get-Mailbox -resultsize unlimited | where {$_.CustomAttribute1 -ne "Org"} | Get-MailboxStatistics | sort TotalItemSize -Descending | select -First 20 } else { Write-Host -ForegroundColor Green "Mailboxes already gathered." }

foreach($mailbox in $usafaMailboxes) { # Create custom object to get data from both get-mailbox and get-mailboxstatistics # List all the data to be gathered for each mailbox $objectHash = @{ "Display Name" = $null "Mailbox Size" = $null "Using Mailbox Database Default Limits" = $null "Prohibit Send Limit" = $null "Prohibit Send/Receive Limit" = $null "Item Count" = $null "Last Logon" = $null "Custom Attribute 1" = $null "Custom Attribute 3" = $null }

# Create custom object based on the has we created above. Each separate value will be populated below.
$mailboxObj = New-Object PSObject -Property $objectHash

$displayName = $mailbox.displayname
$mailboxObj | Add-Member NoteProperty -Name "Display Name" -Value $displayName -Force

$mailboxSize = $mailbox.totalitemsize
$mailboxObj | Add-Member NoteProperty -Name "Mailbox Size" -Value $mailboxSize -force

$usingMailboxDatabaseDefaults = Get-Mailbox $mailbox
$mailboxObj | Add-Member NoteProperty -Name "Using Mailbox Database Default Limits" -Value $usingMailboxDatabaseDefaults.UseDatabaseQuotaDefaults -Force

if($usingMailboxDatabaseDefaults.usedatabasequotadefaults)
{
    $prohibitSendLimit = $mailboxDatabase[0]
    $mailboxObj | Add-Member NoteProperty -Name "Prohibit Send Limit" -Value $prohibitSendLimit.ProhibitSendQuota -Force

    $prohibitSendReceiveLimit = $mailboxDatabase[0]
    $mailboxObj | Add-Member NoteProperty -Name "Prohibit Send/Receive Limit" -Value $prohibitSendReceiveLimit.ProhibitSendReceiveQuota -Force
}
else
{
    $prohibitSendLimit = Get-Mailbox $mailbox
    $mailboxObj | Add-Member NoteProperty -Name "Prohibit Send Limit" -Value $prohibitSendLimit.ProhibitSendQuota -Force

    $prohibitSendReceiveLimit = Get-Mailbox $mailbox
    $mailboxObj | Add-Member NoteProperty -Name "Prohibit Send/Receive Limit" -Value $prohibitSendReceiveLimit.ProhibitSendReceiveQuota -Force
}

$itemCount = $mailbox.ItemCount
$mailboxObj | Add-Member NoteProperty -Name "Item Count" -Value $itemCount -Force

$lastLogon = $mailbox.LastLogonTime
$mailboxObj | Add-Member NoteProperty -Name "Last Logon" -Value $lastLogon -Force

$customAttribute1 = Get-Mailbox $mailbox
$mailboxObj | Add-Member NoteProperty -Name "Custom Attribute 1" -Value $customAttribute1.CustomAttribute1 -Force

$customAttribute3 = Get-Mailbox $mailbox
$mailboxObj | Add-Member NoteProperty -Name "Custom Attribute 3" -Value $customAttribute3.CustomAttribute3 -Force

$mailboxReport += $mailboxObj
}
