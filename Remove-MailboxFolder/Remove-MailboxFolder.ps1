param (
    [Alias("Mailbox")][string]$Recipient,
    [string]$Folder,
    [switch]$Delete,
    [System.Management.Automation.PSCredential]$Credentials)

Import-Module -Name "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
$AppId = 'yourappid'
$AppSecret = 'yoursecret'
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

# Connect to the service
$exchService = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ExchangeService
$exchService.Credentials = New-Object -TypeName Microsoft.Exchange.WebServices.Data.OAuthCredentials($accessToken.access_token)
$exchService.Url = [system.URI]"https://outlook.office365.com/EWS/Exchange.asmx"

# Find the  folder 
$fvFolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(10)  
$sfSearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $Folder)
$folderidcnt = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::'Root', $Recipient)
$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep    
$fvFolderView.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly,
    [Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName)
$fiResult = $null
do {  
    $fiResult = $exchService.FindFolders($folderidcnt, $sfSearchFilter, $fvFolderView)
    foreach ($ffFolder in $fiResult.Folders) {          
        $Entries = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchService, $ffFolder.ID)		
        if ($Entries.Displayname -eq $Folder) { 
            $sourceFolderID = $ffFolder.ID
            Write-Host "Folder" $Entries.Displayname "found" -ForegroundColor Green 
        }
    }   
    $fvFolderView.Offset += $fiResult.Folders.Count
} while ($fiResult.MoreAvailable -eq $true)  

If ($sourceFolderID) {
    $sourceBind = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchService, $sourceFolderID)

    # Get the parent folder displayname
    $ParentFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchService, $sourceBind.ParentFolderId)	
    Write-Host $sourceBind.DisplayName "| Total Items:" $sourceBind.TotalCount "| Parent:" $ParentFolder.DisplayName -ForegroundColor Green
    
    if ($Delete) {
        Try {
            $results = $sourceBind.Delete('MoveToDeletedItems')
            $cont = $true
            foreach ($r in $results) {
                if ($r.result -ne 'Success') {
                    Write-Host "Error:" $r.ErrorMessage -ForegroundColor red
                    $r | Format-List *
                    $cont = $false
                    break
                }
            }
            if ($cont) { Write-Host "Ok deleting" $sourceBind.DisplayName  -ForegroundColor Green }
        }
        catch {
            Write-Host "Error deleting" $sourceBind.DisplayName -ForegroundColor red
            Write-Host $Error[0] | Format-List *
        }
    }
}
else {
    Write-Host "Could not find folder" $Folder -ForegroundColor Yellow
}