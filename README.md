# Microsoft 365 Sizing PowerShell Script


```
PS ~/Development/microsoft-365-sizing> ./Get-RubrikM365SizingInfo.ps1
[INFO] Connecting to the Microsoft Graph API using 'Reports.Read.All' permissions.
[INFO] Retrieving Usage Details for Exchange.
[INFO] Retrieving Usage Details for Sharepoint.                                                                           
[INFO] Retrieving Usage Details for OneDrive.                                                                             
[INFO] Retrieving Storage Usage for Exchange.                                                                             
[INFO] Retrieving Storage Usage for Sharepoint.                                                                           
[INFO] Retrieving Storage Usage for OneDrive.                                                                             
[INFO] Retrieving the subscription License details.                                                                       
[INFO] Disconnecting from the Microsoft Graph API.                                                                        
[INFO] Connecting to the Microsoft Exchange Online Module.                                                                
[INFO] Retrieving all In-Place Archive Exchange Mailbox sizing information.
[INFO] Retrieving Exchange Mailbox Shared Mailbox information.                                                                                                                                                        
[INFO] Disconnecting from the Microsoft Exchange Online Module
[INFO] Calculating the forecasted total storage need for Rubrik.     

M365 Sizing information has been written to ~/Development/microsoft-365-sizing/Rubrik-MS365-Sizing.html   
```

## Requirements

* `PowerShell >= 5.1` for PowerShell Gallery.



## Installation

1. Download the [Get-RubrikM365SizingInfo.ps1](https://github.com/rubrikinc/microsoft-365-sizing/blob/main/Get-RubrikM365SizingInfo.ps1) file to your local machine
2. Install the `Microsoft.Graph.Reports` and `ExchangeOnlineManagement` modules from the PowerShell Gallery

```powershell
Install-Module Microsoft.Graph.Reports, ExchangeOnlineManagement
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
.\RubrikMS365Sizing.txt
```

## Example Output

![Rubrik Sizing HTML Output](https://user-images.githubusercontent.com/8610203/135927453-14334d14-886b-4000-b749-93934d341bd9.png)
