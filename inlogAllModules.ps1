$acctName="jack.schneider@presentnederland.nl"
$orgName="presentnederland"
Connect-ExchangeOnline -UserPrincipalName $acctName -ShowProgress $true
Connect-AzureAD -AccountId $acctName
Connect-MicrosoftTeams -AccountId $acctName

<#Connect-MsolService -AccountId $acctName
Connect-SPOService -Url https://presentnederland-admin.sharepoint.com -AccountId $acctName
#>