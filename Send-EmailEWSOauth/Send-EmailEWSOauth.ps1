param (
    [string]$File,
    [string]$Recipient,
    [string]$MessageBodyFile,
    [int]$Count=10,
    [string]$Subject = "Test Emails, Number",
    [string]$Shared,
    [System.Management.Automation.PSCredential]$Credentials)

Import-Module -Name "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
<<<<<<< HEAD
#$AppId = '571ae6ec-cc83-4fc3-aaae-bd04558e79df'
$AppId = "1477e423-9714-4154-aac4-543f72e32d6a"
$AppSecret = 'H6I7Q~8QESf5AI51UnqrWZqPV0h9vHaCPRwUf'
=======
$AppId = '<app id here>'
$AppSecret = '<app secret here>'
>>>>>>> 446fe47a64909e6aa921377282cab86a72151148
if ($Credentials) {
    $username = $Credentials.UserName
    $Password = $Credentials.GetNetworkCredential().Password
} else {
    $Username = read-host -prompt "Enter email address credentials"
    $Password = read-host -prompt "Enter password" -AsSecureString    
    $Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($password))
}

$Scope = "https://outlook.office365.com/EWS.AccessAsUser.All"
# This value comes from App Registration / Endpoints / OAuth 2.0 token endpoint (v2)
<<<<<<< HEAD
$TokenEndpoint = "https://login.microsoftonline.com/4f50b5ad-6de4-412f-86f6-bcb64cb1b29a/oauth2/v2.0/token" 
=======
$TokenEndpoint = "https://login.microsoftonline.com/<endpoint>/oauth2/v2.0/token" 
>>>>>>> 446fe47a64909e6aa921377282cab86a72151148

# Create body
$Body = @{
    client_id = $AppId
	client_secret = $AppSecret
	scope = $Scope
	grant_type = 'password'
    username = $username
    password = $password
}
$PostThis = @{
    ContentType = 'application/x-www-form-urlencoded'
    Method = 'POST'
    Body = $Body
    Uri = $TokenEndpoint
}
# Request the access token
$accessToken = Invoke-RestMethod @PostThis

$exchService = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ExchangeService
$exchService.Credentials = New-Object -TypeName Microsoft.Exchange.WebServices.Data.OAuthCredentials($accessToken.access_token)
#if ($shared) { $exchService.AutodiscoverUrl($shared, {$true}) } else { $exchService.AutodiscoverUrl($Username, {$true}) }
$exchService.Url = [system.URI]"https://outlook.office365.com/EWS/Exchange.asmx"
if ($Shared) { $Username = $Shared }

$messageBodyPath = convert-path ($MessageBodyFile)
$messageBody = [System.IO.File]::ReadAllText($messageBodyPath)

if ($File) { 
    for ($i = 1; $i -le $count; $i++) { 
        $msg = New-Object -TypeName Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $exchService
        $msg.Subject = ($Subject + " " + $i)
        $msg.Body = $messageBody
        get-content $file | ForEach-Object {
            [void]$msg.ToRecipients.Add($_)
        }
        $msg.SendAndSaveCopy() 
        write-host ("{0} {1} ({2})" -f "Email sent",$msg.Subject,$i)
    }
}
if ($Recipient) {
    for ($i = 1; $i -le $Count; $i++) { 
        $msg = New-Object -TypeName Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $exchService
        $msg.Subject = $Subject
        $msg.Body = $messageBody
        [void]$msg.ToRecipients.Add($Recipient)
        $msg.SendAndSaveCopy() 
        write-host ("{0} {1} ({2})" -f "Email sent:",$msg.Subject,$i)
    }
}
