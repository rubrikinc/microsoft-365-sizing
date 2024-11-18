1. This script can be used to get the M365 sizing info using app-only access.
2. This does not require a user login.
3. Follow https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps to create an app registration to be used for this script.
4. The API permissions required are: 'Reports.Read.All', 'User.Read.All', and 'Group.Read.All' from Microsoft Graph and `Exchange.ManageAsApp` from Office 365 Exchange Online.
5. To run this script successfully, you will need:
	a. Tenant ID.
	b. Client ID (App ID) of the app created above.
	c. Client Secret created on the app above.
6. Refer ../README.md for the options that can be used on the script.
