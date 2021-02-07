param (
    [string]$File,
    [string]$Recipient,
    [string]$MessageBodyFile,
    [int]$Count=10,
    [string]$Subject = "Test Emails, Number",
    [System.Management.Automation.PSCredential]$Credentials)

Function Create-JSON {
    param (
        [string]$Subject,
        [string]$MessageBody,
        [string]$Recipients)

$emailBody = 
@"
{
"message" : {
"subject": "$Subject",
"body" : {
"contentType": "HTML",
"content": "$MessageBody"
},
"toRecipients": [$Recipients]
}
}
"@

    return $emailBody
}

$AppId = '571ae6ec-cc83-4fc3-aaae-bd04558e79df'
$AppSecret = 'P.pTwpG5_i3n2vpqOoL7KrdT9J7jc.-.j4'
if ($Credentials) {
    $username = $Credentials.UserName
    $Password = $Credentials.GetNetworkCredential().Password
} else {
    $Username = read-host -prompt "Enter email address credentials"
    $Password = read-host -prompt "Enter password" -AsSecureString    
    $Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($password))
}
$Scope = "https://graph.microsoft.com/.default"
# This value comes from App Registration / Endpoints / OAuth 2.0 token endpoint (v2)
$TokenEndpoint = "https://login.microsoftonline.com/4f50b5ad-6de4-412f-86f6-bcb64cb1b29a/oauth2/v2.0/token" 

# Prep for access token request
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
# Header for the requests with the access token (token_type should be Bearer)
$Header = @{
    Authorization = "$($accessToken.token_type) $($accessToken.access_token)"
}

$uri = "https://graph.microsoft.com/v1.0/me/sendMail"

$messageBodyPath = convert-path ($MessageBodyFile)
$messageBody = [System.IO.File]::ReadAllText($messageBodyPath).Replace('"',"'")

if ($File) { 
    $addresses = Get-Content $File
    $Recipients = $addresses | ForEach-Object {'{"EmailAddress": {"Address": "'+$_+'"}},'}
    $Recipients = ([string]$Recipients).Substring(0, ([string]$Recipients).Length - 1) # chop off the last comma
}

if ($Recipient) {
    $Recipients = '{ "emailAddress": {"Address": "'+$Recipient+'"}}'
}

For ($i = 1; $i -le $Count; $i++) { 
    $SubjectNew = $Subject+" ("+$i+")"
    $msg = Create-JSON -Subject $SubjectNew -MessageBody $messageBody -Recipients $Recipients
    write-host ("{0} {1} ({2})" -f "Sending",$Subject,$i)
    Invoke-RestMethod -Headers $Header -Uri $uri -Body $msg -Method Post -ContentType 'application/json' | out-null
}
