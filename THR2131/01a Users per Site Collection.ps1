$SiteCollections  = Get-SPOSite -Limit All
foreach ($site in $SiteCollections)
{
try {
    for ($i=0;;$i+=50) {
        $ExternalUsers += Get-SPOExternalUser -SiteUrl $site.Url -PageSize 50 -Position $i -ea Stop | Select DisplayName,EMail,AcceptedAs,WhenCreated,InvitedBy,@{Name = "Url" ; Expression = { $site.url } }
    }
}
catch {
}
}

$ExternalUsers | Export-Csv -Path "C:\PowerShell\ExternalUsersPerSC.csv" -NoTypeInformation
