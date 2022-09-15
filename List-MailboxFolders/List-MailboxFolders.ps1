param (
    [Alias("Mailbox")][string]$Recipient,
    [System.Management.Automation.PSCredential]$Credentials)

Import-Module -Name "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
$AppId = 'appid'
$AppSecret = 'appsecret'
if ($Credentials) {
    $username = $Credentials.UserName
    $Password = $Credentials.GetNetworkCredential().Password
}
else {
    $Username = Read-Host -Prompt "Enter email address credentials"
    $Password = Read-Host -Prompt "Enter password" -AsSecureString    
    $Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($password))
}

$Scope = "https://outlook.office365.com/EWS.AccessAsUser.All"
# This value comes from App Registration / Endpoints / OAuth 2.0 token endpoint (v2)
$TokenEndpoint = "https://login.microsoftonline.com/yourEndpoint/oauth2/v2.0/token"

# Create body
$Body = @{
    client_id     = $AppId
    client_secret = $AppSecret
    scope         = $Scope
    grant_type    = 'password'
    username      = $username
    password      = $password
}
$PostThis = @{
    ContentType = 'application/x-www-form-urlencoded'
    Method      = 'POST'
    Body        = $Body
    Uri         = $TokenEndpoint
}
# Request the access token
$accessToken = Invoke-RestMethod @PostThis

$exchService = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ExchangeService
$exchService.Credentials = New-Object -TypeName Microsoft.Exchange.WebServices.Data.OAuthCredentials($accessToken.access_token)
$exchService.Url = [system.URI]"https://outlook.office365.com/EWS/Exchange.asmx"

$fvFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(10)  
$folderidcnt = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::'MsgFolderRoot', $Recipient)
# Ensure all folders in the search path are returned  
$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep;  
$psPropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
$PR_FOLDER_TYPE = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(13825, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer) 
$PR_MESSAGE_SIZE_EXTENDED = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3592, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Long)
$PR_Folder_Path = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
$psPropertySet.Add($PR_MESSAGE_SIZE_EXTENDED)
$psPropertySet.Add($PR_Folder_Path)
$fvFolderView.PropertySet = $psPropertySet 
# Exclude any Search Folders  
$sfSearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PR_FOLDER_TYPE, "1")  
$fiResult = $null

do {  
    $fiResult = $exchService.FindFolders($folderidcnt, $sfSearchFilter, $fvFolderView)  
    foreach ($ffFolder in $fiResult.Folders) { 
        $ParentFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchService, $ffFolder.ParentFolderId)	
        Write-Host $ffFolder.DisplayName "| Total Items:" $ffFolder.TotalCount "| Parent Folder:" $ParentFolder.DisplayName -ForegroundColor Green
    }        
    $fvFolderView.Offset += $fiResult.Folders.Count
} while ($fiResult.MoreAvailable -eq $true)