## Requirements

* `PowerShell >= 5.1` for PowerShell Gallery.
* Microsoft user permissions to run this script: Global Reader and Reports Reader

* There are two ways users can authenticate to Exchange Online:
* 1. App access:
*     Follow https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps to create an app        registration to be used for this script.
*     The API permissions required are: 'Reports.Read.All', 'User.Read.All', and 'Group.Read.All' from Microsoft Graph and `Exchange.ManageAsApp` from Office 365 Exchange Online.
*     To run this script successfully, you will need:
	    a. Tenant ID.
	    b. Client ID (App ID) of the app created above.
	    c. Client Secret created on the app above.

* 2. User access:
*     Login through the admin user account when prompted on the browser.


## Installation

1. Download the [Get-RubrikM365SizingInfo.ps1](https://github.com/rubrikinc/microsoft-365-sizing/archive/refs/heads/main.zip) PowerShell script to your local machine
2. Install the `Microsoft.Graph.Reports` and `ExchangeOnlineManagement` modules from the PowerShell Gallery

```powershell
Install-Module Microsoft.Graph.Reports, Microsoft.Graph.Groups, ExchangeOnlineManagement
```

## Usage

1. Open a PowerShell terminal and navigate to the folder/directory where you previously downloaded the [Get-RubrikM365SizingInfo.ps1](https://github.com/rubrikinc/microsoft-365-sizing/blob/main/Get-RubrikM365SizingInfo.ps1) file.

2. Run the script.

```
./Get-RubrikM365SizingInfo.ps1
```

> NOTE - If you receive a PowerShell execution policy error message you can run the following command:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

3. Authenticate and acknowledge report access permissions in the browser window/tab that appears. This will occur twice during the script execution.

> Note: There is a known issue with the Microsoft authentication process that may result in an error message during the initial authentication process. If this occurs, re-run the script and the error will no longer show.

4. The script will run and the results will be written to a html file in the directory in which it was run.

```
.\RubrikMS365Sizing.html
```

## Options

If you want to run the script with app access, use the following:
```
./Get-RubrikM365SizingInfo.ps1 -UseAppAccess $true
```

If you want to run the script against a single AD Group, use the following:
```
./Get-RubrikM365SizingInfo.ps1 -ADGroup "RubrikEmployees"
```

The script will try to gather In Place Archive sizes for each mailbox. However, to do so, the script needs to query each mailbox user's information which can timeout for larger environments. If that's the case, you can skip gathering In Place Archives with the following:
```
./Get-RubrikM365SizingInfo.ps1 -SkipArchiveMailbox $true
```

The script will try to gather stats for the Recoverable Items folder. This can also take awhile and timeout in larger environments. You can skip this by using:
```
./Get-RubrikM365SizingInfo.ps1 -SkipRecoverableItems $true
```

The script will calculate annual growth rates for 10%, 20%, and 30% annual growth rates. You can change the 30% to a custom value such as 40% by using the following flag:
```
./Get-RubrikM365SizingInfo.ps1 -AnnualGrowth 40
```



## What information does the script access?

The majority of the information collected is directly from the Microsoft 365 [Usage reports](https://docs.microsoft.com/en-us/microsoft-365/admin/activity-reports/activity-reports?view=o365-worldwide) that are found in the admin center.



# Microsoft 365 Sizing PowerShell Script


```
 ./Get-RubrikM365SizingInfo.ps1
[INFO] Starting the Rubrik Microsoft 365 sizing script (v5.0).
[INFO] Connecting to the Microsoft Graph API using 'Reports.Read.All', 'User.Read.All', and 'Group.Read.All' permissions.
[INFO] Retrieving usage info for ...
 - Exchange
 - Usage report for Exchange output to: .\getMailboxUsageDetail.csv
[INFO] Retrieving usage info for ...
 - OneDrive
 - Usage report for OneDrive output to: .\getOneDriveUsageAccountDetail.csv
[INFO] Retrieving usage info for ...
 - SharePoint
 - Usage report for SharePoint output to: .\getSharePointSiteUsageDetail.csv
[INFO] Retrieving historical usage reports
[INFO] Current usage data and historical reports may differ pending deletions
[INFO] OneDrive usage:
  - Current usage (calculated with per-user stats): 1043.92 GB
  - Usage on 2024-03-05: 90.76 GB
  - Usage on 2023-09-08: 976.86 GB
  - Growth over 180 days: 67.06 GB
  - Growth annualized per year: 135.98 GB, 13%
[INFO] SharePoint usage:
  - Current usage (calculated with per-user stats): 30.06 GB
  - Usage on 2024-03-05: 30.06 GB
  - Usage on 2023-09-08: 21.27 GB
  - Growth over 180 days: 8.79 GB
  - Growth annualized per year: 17.82 GB, 59%
[INFO] Exchange usage:
  - Current usage (calculated with per-user stats): 0.81 GB
  - Usage on 2024-03-05: 0.32 GB
  - Usage on 2023-09-08: 1.57 GB
  - Growth over 180 days: -0.76 GB
  - Growth annualized per year: -1.54 GB, -190%
[NOTE] If the growth looks odd, try using a different period (parameter: -Period 7, 30, 90, 180) days
[INFO] Calculating the forecasted total storage need for Rubrik.
[INFO] Disconnecting from the Microsoft Graph API.
Now gathering In Place Archive usage
This may take awhile since stats need to be gathered per user
Progress will be written as they are gathered
[INFO] Switching to the Microsoft Exchange Online Module for more detailed reporting capabilities.
[INFO] Retrieving all Exchange Mailbox In-Place Archive sizing
[INFO] Found 4 mailboxes with In Place Archives
[0 / 4] Processing mailboxes ...
[INFO] Finished gathering stats on mailboxes with In Place Archive
[INFO] Total # of mailboxes with In Place Archive: 4
[INFO] Total size of mailboxes with In Place Archive: 0.01 GB
[INFO] Total # of items of mailboxes with In Place Archive: 74
[INFO] Disconnecting from the Microsoft Exchange Online Module

M365 Sizing information has been written to /home/Rubrik-M365-Sizing-2024-03-07.html

```


## Example Output

![image](https://user-images.githubusercontent.com/51362633/190453033-94379a84-8678-4592-9d9b-2b1dad96a521.png)




