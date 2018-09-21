function GetPictureInfo ($AllUsers) {
	foreach ($user in $AllUsers) {
        $Picture = Get-UserPhoto -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue
        if ($Picture -eq $null)
        {
        $props = @{'Username'=$user.UserPrincipalName;
		'DisplayName'=$user.DisplayName}
		New-Object -TypeName PSObject -Property $props
        }
        else
        {}}}

function GetManagerInfo ($AllUsers) {
	foreach ($user in $AllUsers) {
        $Manager = Get-AzureADUserManager -ObjectId $user.UserPrincipalName
        if ($Manager -eq $null)
        {
        $props = @{'Username'=$user.UserPrincipalName;
		'DisplayName'=$user.DisplayName}
		New-Object -TypeName PSObject -Property $props
        }
        else
        {}}}


function SharePointInfo ($AllUsers) {
	foreach ($user in $AllUsers) {
        $SPProfile  = Get-PnPUserProfileProperty -Account $user.UserPrincipalName -ErrorAction SilentlyContinue
        if ($SPProfile -ne $null)
        {
          if ($SPProfile.UserProfileProperties.'SPS-Birthday' -ne "")
            {
               $Birthday = ([datetime]$SPProfile.UserProfileProperties.'SPS-Birthday').ToString('dd/MM')
            }
          else
            {
                $Birthday = ""
            }
        $props = @{'Username'=$user.UserPrincipalName;
		'AboutMe'=$SPProfile.UserProfileProperties.AboutMe;
		'CellPhone'=$SPProfile.UserProfileProperties.CellPhone;
        'PastProjects'=$SPProfile.UserProfileProperties.'SPS-PastProjects';
        'Skills'= $SPProfile.UserProfileProperties.'SPS-Skills';
        'Birthday'= $Birthday
        }
		New-Object -TypeName PSObject -Property $props
        }
        else
        {}}}

$Date = Get-Date -Format d
$Users = Get-AzureADUser | Where {$_.UserType -eq 'Member' -and $_.AssignedLicenses -ne $null}


$params = @{'As'='Table';
'PreContent'='<h2>Users With No Picture</h2>';
'EvenRowCssClass'='even';
'OddRowCssClass'='odd';
'TableCssClass'='grid';
'Properties'='Username','DisplayName'}

$PictureInfo = GetPictureInfo $Users | ConvertTo-EnhancedHTMLFragment @params


$params = @{'As'='Table';
'PreContent'='<h2>Users With No Manager</h2>';
'EvenRowCssClass'='even';
'OddRowCssClass'='odd';
'TableCssClass'='grid';
'Properties'='Username','DisplayName'}

$ManagerInfo = GetManagerInfo $Users | ConvertTo-EnhancedHTMLFragment @params


$params = @{'As'='Table';
	'PreContent'='<h2> + All Users Highlight</h2>';
	'EvenRowCssClass'='even';
	'OddRowCssClass'='odd';
	'MakeTableDynamic'=$true;
	'MakeHiddenSection'=$true;
	'TableCssClass'='grid';
	'Properties'='UserPrincipalName',@{n='First Name';e={$_.GivenName};css={if ($_.GivenName -eq $null) { 'redCell' }}},@{n='Family Name';e={$_.Surname};css={if ($_.Surname -eq $null) { 'redCell' }}},@{n='Display Name';e={$_.DisplayName};css={if ($_.DisplayName -eq $null) { 'redCell' }}},@{n='Job Title';e={$_.JobTitle};css={if ($_.JobTitle -eq $null) { 'redCell' }}},@{n='Department';e={$_.Department};css={if ($_.Department -eq $null) { 'redCell' }}},@{n='Cell Phone';e={$_.Mobile};css={if ($_.Mobile -eq $null) { 'redCell' }}},@{n='City';e={$_.City};css={if ($_.City -eq $null) { 'redCell' }}},@{n='Office Phone';e={$_.TelephoneNumber};css={if ($_.TelephoneNumber -eq $null) { 'redCell' }}},@{n='Country';e={$_.Country};css={if ($_.Country -eq $null) { 'redCell' }}},@{n='State';e={$_.State};css={if ($_.State -eq $null) { 'redCell' }}}}

$AlluserInfo = Get-AzureADUser | Where {$_.UserType -eq 'Member' -and $_.AssignedLicenses -ne $null} | ConvertTo-EnhancedHTMLFragment @params -ErrorAction SilentlyContinue

$params = @{'As'='Table';
	'PreContent'='<h2> + SharePoint Profile Properties</h2>';
	'EvenRowCssClass'='even';
	'OddRowCssClass'='odd';
	'MakeTableDynamic'=$true;
	'MakeHiddenSection'=$true;
	'TableCssClass'='grid';
	'Properties'='Username',@{n='About Me';e={$_.AboutMe};css={if ($_.AboutMe -eq "") { 'redCell' }}},@{n='Mobile Phone';e={$_.CellPhone};css={if ($_.CellPhone -eq "") { 'redCell' }}},@{n='Past Projects';e={$_.PastProjects};css={if ($_.PastProjects -eq "") { 'redCell' }}},@{n='Skills';e={$_.Skills};css={if ($_.Skills -eq "") { 'redCell' }}},@{n='Birthday';e={$_.Birthday};css={if ($_.Birthday -eq "") { 'redCell' }}}}

$SPInfo = SharePointInfo $Users | ConvertTo-EnhancedHTMLFragment @params


ConvertTo-EnhancedHTML -HTMLFragments $PictureInfo, $ManagerInfo, $AlluserInfo, $SPInfo -CssUri 'D:\Dropbox\SharePoint Stuff\Session Submissions\2018-09-25 Ignite\THR2131\Demos\styles2.css' -PreContent "<H1>Profile Completeness Report on $date</H1>" | Out-File "C:\PowerShell\Profile.html"
