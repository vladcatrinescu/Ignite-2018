$cred = get-Credential
$cred.Password.MakeReadOnly()

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred -Authentication Basic -AllowRedirection
Import-PSSession $Session

Connect-SPOService -Url https://globomanticsorg-admin.sharepoint.com/ -Credential $cred
Connect-PnpOnline -Url https://globomanticsorg-admin.sharepoint.com/ -Credentials $cred

Import-Module AzureADPreview
Connect-AzureAD -Credential $cred
