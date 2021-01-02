
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -ShowProgress $true

$DownLoadDirectory = Get-ItemPropertyValue 'HKCU:\software\microsoft\windows\currentversion\explorer\shell folders\' -Name '{374DE290-123F-4565-9164-39C4925E467B}'
$TimeStamp = Get-Date -UFormat %Y-%m-%d_%H-%M

$MailboxFileName ="MailboxExport-"+$TimeStamp +".csv"
$MailboxFile = $DownLoadDirectory + "\" + $MailboxFileName

Get-Mailbox -ResultSize Unlimited | Select * |Export-CSV -NoTypeInformation -LiteralPath $MailboxFile

$MailUserFileName ="MailUserExport-"+$TimeStamp +".csv"
$MailUserFile = $DownLoadDirectory + "\" + $MailUserFileName

Get-MailUser -ResultSize Unlimited |select * | export-csv -NoTypeInformation -LiteralPath $MailUserFile

$UserFileName ="UserExport-"+$TimeStamp +".csv"
$UserFile = $DownLoadDirectory + "\" + $UserFileName

Get-User -ResultSize Unlimited |select * | export-csv -NoTypeInformation -LiteralPath $UserFile

$PermissionFileName ="PermissionExport-"+$TimeStamp +".csv"
$PermissionFile = $DownLoadDirectory + "\" + $PermissionFileName


Get-Mailbox | Get-MailboxPermission |select *| Export-Csv -NoTypeInformation -LiteralPath $PermissionFile

Disconnect-ExchangeOnline