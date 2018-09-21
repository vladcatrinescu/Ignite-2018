$SiteCollections  = Get-SPOSite -Limit All
foreach ($site in $SiteCollections)
{
    try 
    {
	$ExternalUserswithModern += Get-SPOUser -Limit All -Site $site.Url | Where {$_.LoginName -like "*urn:spo:guest*" -or $_.LoginName -like "*#ext#*"} | Select DisplayName,LoginName,@{Name = "Url" ; Expression = { $site.url }}
    }
    catch
    { 
    Write-Host "Could not get user from the following URL! Make sure you are a SC Admin of it! " $site.Url
    }
}

$ExternalUserswithModern | Export-Csv -Path "C:\PowerShell\ExternalUsersPerSCWithModern.csv" -NoTypeInformation
