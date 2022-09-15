# List The Folders in an Exchange Online mailbox using EWS and OAUTH
Although it's clear the long term direction is Microsoft Graph, there are still some things that can only be done with EWS. This isn't actually one of them, but while writing the Remove-MailboxFolder script, it became obvious that a script to list all folders, its parent folder, and number of items would be useful.

For the app registraiom steps, refer to the Remove-MailboxFolder README.