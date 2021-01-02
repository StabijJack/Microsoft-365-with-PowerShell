# set export Directory
Add-Type -AssemblyName System.Windows.Forms
$fileBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{ 
    Description = 'Select a Export Directory'
    RootFolder = 'personal'
}
$null = $fileBrowser.ShowDialog()
$DownLoadDirectory = $fileBrowser.SelectedPath

Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline -ShowProgress $true

# $DownLoadDirectory = Get-ItemPropertyValue 'HKCU:\software\microsoft\windows\currentversion\explorer\shell folders\' -Name '{374DE290-123F-4565-9164-39C4925E467B}'


$MailboxFileName ="MailboxExport.csv"
$MailboxFile = $DownLoadDirectory + "\" + $MailboxFileName

Get-Mailbox -ResultSize Unlimited | Select-Object * |Export-CSV -NoTypeInformation -LiteralPath $MailboxFile

$MailUserFileName ="MailUserExport.csv"
$MailUserFile = $DownLoadDirectory + "\" + $MailUserFileName

Get-MailUser -ResultSize Unlimited |Select-Object * | export-csv -NoTypeInformation -LiteralPath $MailUserFile

$UserFileName ="UserExport.csv"
$UserFile = $DownLoadDirectory + "\" + $UserFileName

Get-User -ResultSize Unlimited |Select-Object * | export-csv -NoTypeInformation -LiteralPath $UserFile

$PermissionFileName ="PermissionExport.csv"
$PermissionFile = $DownLoadDirectory + "\" + $PermissionFileName


Get-Mailbox | Get-MailboxPermission |Select-Object *| Export-Csv -NoTypeInformation -LiteralPath $PermissionFile

Disconnect-ExchangeOnline