# Delete a folder from an Exchnage Online mailbox using EWS and OAUTH
Although it's clear the long term direction is Microsoft Graph, there are still some things that can only be done with EWS. In this simple examle, we will delete a folder from the mailbox. Really the most useful items here are the setup steps required in Azure. These are quite similar to the steps in the Get-InboxReadStatysGraph but there are some small differences between registering a Graph app and an EWS.

1.	Follow this article ![Authenticate an EWS application by using OAuth | Microsoft Docs] (https://nam06.safelinks.protection.outlook.com/?url=https%3A%2F%2Fdocs.microsoft.com%2Fen-us%2Fexchange%2Fclient-developer%2Fexchange-web-services%2Fhow-to-authenticate-an-ews-application-by-using-oauth&data=05%7C01%7Cdarosen%40microsoft.com%7C70a7694c9e194d0da8b808da929126b6%7C72f988bf86f141af91ab2d7cd011db47%7C1%7C0%7C637983449369693703%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C3000%7C%7C%7C&sdata=OEnzqoITDFRxDl37p3voDjPTGqfZ5TK9JpRU%2BRoXbNo%3D&reserved=0) to create the app registration. 
Add a client secret and copy it into line 9 of the scripts. 
Copy the Application (client) id into line 8. 
Click on Endpoint, copy the OAuth 2.0 token endpoint (v2) value into line 21.
2. Do follow the steps in the Configure for delegated authentication to edit the manifest. 
