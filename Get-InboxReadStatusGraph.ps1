# Delegated example, good for the /me stuff
$AppId = '87f0c70f-7d8a-4e54-9c17-c04a58123456'
$AppSecret = '0610dN_s.vazVeGDZYoRe=.0cd123456'
$Scope = "https://graph.microsoft.com/.default"
$Username = read-host -prompt "Enter upn" # "meganb@M365xXXXXXX.onmicrosoft.com" for example
$Password = read-host -prompt "Enter password" -AsSecureString
$password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($password))
# This value comes from App Registration / Endpoints / OAuth 2.0 token endpoint (v2)
$TokenEndpoint = "https://login.microsoftonline.com/0bec5193-f835-4d13-94c3-ba5ec9123456/oauth2/v2.0/token" 
$counter = 0; $read = 0; $unread = 0

# Create body
$Body = @{
    client_id = $AppId
	client_secret = $AppSecret
	scope = $Scope
	grant_type = 'password'
    username = $username
    password = $password
}
# Arrange the parameters for Invoke-RestMethod for cleaner code
$PostThis = @{
    ContentType = 'application/x-www-form-urlencoded'
    Method = 'POST'
    Body = $Body
    Uri = $TokenEndpoint
}
# Request the access token
$Request = Invoke-RestMethod @PostThis

# Header for the requests with the access token (token_type should be Bearer)
$Header = @{
    Authorization = "$($Request.token_type) $($Request.access_token)"
}

$Uri = "https://graph.microsoft.com/v1.0/me/messages"

# Make the request, process the results
$results = Invoke-RestMethod -Uri $Uri -Headers $Header -Method Get -ContentType "application/json"
do {
    for ($i = 0; $i -lt $results.value.count; $i++) {
        New-Object -TypeName PSObject -Property @{
            isRead = $results.value[$i].IsRead
            subject = $results[$i].subject
            DateTimeReceived = $results.value[$i].receivedDateTime
            Sender = $results.value[$i].from.emailaddress.address 
            MessageID = $results.value[$i].InterNetMessageID
        } | export-csv -Path .\GraphMessages.csv  -Append
        write-host $results.value[$i].receivedDateTime $results.value[$i].from.emailaddress.address $results.value[$i].Subject
        $counter++; if ($results.value[$i].IsRead) { $read++ } else { $unread++ } 
    }
    $results = Invoke-RestMethod  -uri $results.'@odata.nextLink' -Headers $Header -Method Get -ContentType "application/json"
} while($results.'@odata.nextLink')
Write-Host $read "read /" $unread "unread /" $counter "total"
# How to set one of the messages to Read; make sure to specify a message id
#$json = '{
#    "isRead" : "true"
#}'
#$setRead = Invoke-RestMethod -Headers $Header -Method Patch -ContentType "application/json" -Body $json -Uri 'https://graph.microsoft.com/v1.0/me/messages/AAMkADZjYmRmNjRjLTkwODUtNGVhOS1iMzliLTJiOWVhYTQ5YWI1YwBGAAAAAAAmdWNG_SMFTZvzE7aNLXxrBwCSvslsWfTeQZXAUtmtJOixAAAAAAEMAACSvslsWfTeQZXAUtmtJOixAAA12345678='