# Access Exchange Online Mailbox Inbox via Microsoft Graph / REST API using PowerShell
There aren't many simple examples of accessing your Inbox via Microsoft Graph / REST API using PowerShell. Graph Explorer is a great tool for testing out queries but it doesn't give the whole end to end process especially how to acquire the tokens - hopefully this example will be easy for you to do some testing with and can leverage into more complex solutions. Also a good example of cycling through all the results when they come back paged, which again was not easy to find examples for.

First there is some setup required in your Azure tenant - you need to create an App Registration. Here are the steps:
1. Access your tenant at https://portal.azure.com. If a demo tenant log on with the admin@M365xXXXXXX.onmicrosoft.com id. Administrator role is typically required to create an app registration.
2. Navigate to App Registrations (just put that in the Search bar)
3. Create a New Registration / give it a name, select who can access (typically Single tenant); a Redirect URI is not required. Make note of the Application (Client) ID.
4. Click on Endpoints and make note of the OAuth 2.0 token endpoint (v2)
5. Select Certificates & secrets, click New Client Secret, give it a name and select an Expiration term, make note of the Value.
6. Select API permissions. Click Add a permission. For Read access to your mailbox, click Microsoft Graph, Delegated Permissions, Mail, Mail.Read. If you'll be making updates to messages, select Mail.ReadWrite. Click Add Permission.
7. Click Grant admin consent. Done!

Now update the sample code with the App ID, App (client) Secret, and Token Endpoint, and execute the code. Enter upn and password of the mailbox you want to look at.

The output to the screen is the Date/Time Received, From address, and Subject.
The results are saved to a csv file (GraphMessages.csv) also and include those values plus the Read/UnRead status and the messageID. Additional attributes can be added easily.

The Delegated permissions in this example allow for access to the mailbox of the user whose id and password you entered. There are lots of permissions to grant as seen in the API permissions section, the Delegated rights are best for the "/me" URIs you'd typically see in Graph Explorer examples, which is a great resource for finding more URIs to test.
Differences between Delegated and Application are here https://docs.microsoft.com/en-us/azure/active-directory/develop/developer-glossary#permissions.

Finally, there are some commented out lines of code at the very bottom for setting the Read / UnRead flag on a single message, just to have an example of a Patch method.
