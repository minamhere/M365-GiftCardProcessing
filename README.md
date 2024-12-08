# M365-GiftCardProcessing

Create an Azure App Registration:

Go to the Azure Portal → Azure Active Directory → App registrations → New registration.
Ensure redirect URI is set to https://login.microsoftonline.com/common/oauth2/nativeclient for simplicity.
Use Public Client/Native as the type.

Assign Mail.Read delegated permission under API Permissions for Microsoft Graph.

Copy the Application (client) ID and Directory (tenant) ID.

Install the MSAL.PS module:

```powershell
Install-Module -Name MSAL.PS -Scope CurrentUser
```
