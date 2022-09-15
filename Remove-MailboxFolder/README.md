# Delete a folder from an Exchnage Online mailbox using EWS and OAUTH
Although it's clear the long term direction is Microsoft Graph, there are still some things that can only be done with EWS. In this simple examle, we will delete a folder from the mailbox. Really the most useful items here are the setup steps required in Azure. These are quite similar to the steps in the Get-InboxReadStatysGraph but there are some small differences between registering a Graph app and an EWS.
## Setup
1.	Follow this article [Authenticate an EWS application by using OAuth](https://nam06.safelinks.protection.outlook.com/?url=https%3A%2F%2Fdocs.microsoft.com%2Fen-us%2Fexchange%2Fclient-developer%2Fexchange-web-services%2Fhow-to-authenticate-an-ews-application-by-using-oauth&data=05%7C01%7Cdarosen%40microsoft.com%7C70a7694c9e194d0da8b808da929126b6%7C72f988bf86f141af91ab2d7cd011db47%7C1%7C0%7C637983449369693703%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C3000%7C%7C%7C&sdata=OEnzqoITDFRxDl37p3voDjPTGqfZ5TK9JpRU%2BRoXbNo%3D&reserved=0 "Authenticate an EWS application by using OAuth | Microsoft Docs") to create the app registration. 
Add a client secret and copy it into line 9 of the scripts. 
Copy the Application (client) id into line 8. 
Click on Endpoint, copy the OAuth 2.0 token endpoint (v2) value into line 21.
2.  Do follow the steps in the [Configure for delegated authentication to edit the manifest](https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth#configure-for-delegated-authentication "Configure for delegated authentication to edit the manifest") section. 
![Image1](https://github.com/dlRo/mband/blob/3db9ff2d9bb38890dc995c7bddfbd5dbd7e2db83/Remove-MailboxFolder/Images/image1.png)
3.	You should then see this ready for the granting of admin consent in the API Permissions section:
![Image2](https://github.com/dlRo/mband/blob/3db9ff2d9bb38890dc995c7bddfbd5dbd7e2db83/Remove-MailboxFolder/Images/Image2.png)
4.	Admin consent for EWS.AccessAsUser.all granted:
![Image3](https://github.com/dlRo/mband/blob/3db9ff2d9bb38890dc995c7bddfbd5dbd7e2db83/Remove-MailboxFolder/Images/Image3.png)
5.	Open a PowerShell session and create a credentials variable for the Exchange admin id; in my demo tenant this was admin@m365x123456.onmicrosoft.com 
**$admin = get-credential -UserName admin@m365x123456.onmicrosoft.com -message "Exchange Admin"**
This step can be skipped but means you'll be prompted for id and password on every iteration. Not practical for running against a lot of mailboxes!
## Execute
•	There are two scripts here. The first is List-MailboxFolders.ps1 which lists all folders in the mailbox under MsgFolderRoot. Helpful for troubleshooting. Any mailbox the admin id has rights to can be specified:
**\List-MailboxFolders.ps1 -Mailbox "meganb@m365xq123456onmicrosoft.com" -Credentials $admin**
![Image4](https://github.com/dlRo/mband/blob/3db9ff2d9bb38890dc995c7bddfbd5dbd7e2db83/Remove-MailboxFolder/Images/Image4.png)
•	The Remove-MailboxFolder.ps1 script without the -Delete switch will find the folder if it exists and list the number of items; good for prep / testing:
**.\Remove-MailboxFolder.ps1 -Recipient "meganb@m365x123456.onmicrosoft.com" -Credentials $admin -Folder 'Project Falcon'**
![Image5](https://github.com/dlRo/mband/blob/3db9ff2d9bb38890dc995c7bddfbd5dbd7e2db83/Remove-MailboxFolder/Images/Image5.png)
•	When you're ready to pull the trigger, add the -Delete switch, works ok even if the folder is not empty:
**.\Remove-MailboxFolderEWS.ps1 -Recipient "meganb@m365x123456.onmicrosoft.com" -Credentials $admin -Folder 'Project Falcon' -Delete**
![Image6](https://github.com/dlRo/mband/blob/3db9ff2d9bb38890dc995c7bddfbd5dbd7e2db83/Remove-MailboxFolder/Images/Image6.png)
•	Even after deletion, you will still be able to find the folder but note the new parent folder:
![Image7](https://github.com/dlRo/mband/blob/3db9ff2d9bb38890dc995c7bddfbd5dbd7e2db83/Remove-MailboxFolder/Images/Image7.png)
•	Change line 78 from MoveToDeletedItems to HardDelete or SoftDelete as needed. Parameterize if needed.
[DeleteMode Enum (Microsoft.Exchange.WebServices.Data) | Microsoft Docs](https://docs.microsoft.com/en-us/dotnet/api/microsoft.exchange.webservices.data.deletemode?view=exchange-ews-api "DeleteMode Enum (Microsoft.Exchange.WebServices.Data)")