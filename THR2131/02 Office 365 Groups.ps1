function GetAllOffice365Groups{
	$Groups = Get-UnifiedGroup
	foreach ($group in $Groups) {
		$LastSPSiteAccess = Search-UnifiedAuditLog -StartDate ((get-date).adddays(-90)) -EndDate (Get-Date) -ObjectId $group.SharePointSiteUrl | Sort-Object -Descending -Property CreationDate | Select-Object -First 1 -ExpandProperty CreationDate
		
		if ($LastSPSiteAccess -eq $null) { 
			$LastSPSiteAccess = "90 days +"
		}
		else
		{
			$LastSPSiteAccess  = $LastSPSiteAccess.tostring("MM-dd-yyyy")
		}
		
		$LastSPDocumentLib = Search-UnifiedAuditLog -StartDate ((get-date).adddays(-90)) -EndDate (Get-Date) -ObjectId $group.SharePointDocumentsUrl | Sort-Object -Descending -Property CreationDate | Select-Object -First 1 -ExpandProperty CreationDate 
		
		if ($LastSPDocumentLib -eq $null) {
			$LastSPDocumentLib = "90 days +"
		}
		else
		{
			$LastSPDocumentLib  = $LastSPDocumentLib.tostring("MM-dd-yyyy")
		}
		$Mailbox = Get-MailboxFolderStatistics -Identity $Group.Alias -IncludeOldestAndNewestITems -FolderScope Inbox
		if ((NEW-TIMESPAN -Start $Mailbox.NewestItemReceivedDate -End (Get-Date)).Days -gt 90)
		{
			$LastConversation = "90 days +"
		}
		else
		{
			$LastConversation = $Mailbox.NewestItemReceivedDate
		}
		$props = @{'Name'=$group.DisplayName;
		'Address'=$group.PrimarySmtpAddress;
		'Privacy Type'=$group.AccessType;
		'Members'=Get-UnifiedGroupLinks -Identity $group.Alias -LinkType Members | measure | % { $_.Count };
		'Owners'=Get-UnifiedGroupLinks -Identity $group.Alias -LinkType Owners | measure | % { $_.Count };
		'LastSPView'=$LastSPSiteAccess;
		'LastSPDocLib'=$LastSPDocumentLib;
		'LastInbox'=$LastConversation;
		'External Members'=$group.GroupExternalMemberCount
		}
		New-Object -TypeName PSObject -Property $props
	}
}


function GetExpiredOffice365Groups{
	$Groups = Get-AzureADMSDeletedGroup | Sort-Object -Property DeletedDateTime
	foreach ($group in $groups) {
		$props = @{'Name'=$group.DisplayName;
		'Address'=$group.Mail;
		'Deleted Date'=$group.DeletedDateTime;
		'DaysLeft'=(NEW-TIMESPAN -Start (Get-Date) -End (($group.DeletedDateTime).AddDays(30))).Days}
		New-Object -TypeName PSObject -Property $props
	}
}

$Date = Get-Date -Format d
$TenantName =  Get-AzureADTenantDetail | Select DisplayName

$params = @{'As'='Table';
'PreContent'='<h2>All Groups</h2>';
'EvenRowCssClass'='even';
'OddRowCssClass'='odd';
'TableCssClass'='grid';
'MakeTableDynamic'=$true;
'Properties'='Name','Address','Privacy Type', @{n='#Members';e={$_.Members};css={if ($_.Members -lt "1") { 'red' }}},@{n='#Owners';e={$_.Owners};css={if ($_.Owners -lt "2") { 'red' }}},@{n='Last SharePoint Home Visit';e={$_.LastSPView};css={if ($_.LastSPView -eq "90 days +") { 'red' }}},@{n='Last SharePoint Document Activity';e={$_.LastSPDocLib};css={if ($_.LastSPDocLib -eq "90 days +") { 'red' }}},@{n='Last Conversation';e={$_.LastInbox};css={if ($_.LastInbox -eq "90 days +") { 'red' }}} }

$AllGroupsInfo = GetAllOffice365Groups | ConvertTo-EnhancedHTMLFragment @params

$params = @{'As'='Table';
'PreContent'='<h2>Recently Deleted Groups</h2>';
'EvenRowCssClass'='even';
'OddRowCssClass'='odd';
'TableCssClass'='grid';
'Properties'='Name','Address','Deleted Date',@{n='Days left to restore';e={$_.DaysLeft};css={if ($_.DaysLeft -lt "18") { 'red' }}}}

$ExpiredGroupsInfo = GetExpiredOffice365Groups | ConvertTo-EnhancedHTMLFragment @params

ConvertTo-EnhancedHTML -HTMLFragments $ExpiredGroupsInfo, $AllGroupsInfo -CssUri 'D:\Dropbox\SharePoint Stuff\Session Submissions\2018-09-25 Ignite\THR2131\Demos\styles2.css' -PreContent "<H1>Office 365 Groups Report for $TenantName on $date</H1>" | Out-File "C:\PowerShell\Groups.html"

