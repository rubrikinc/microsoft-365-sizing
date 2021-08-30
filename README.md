# Microsoft 365 Sizing PowerShell Script

## Requirements

* `PowerShell >= 5.1` for PowerShell Gallery.
* User with MS Graph API access.

## Installation

1. Download the [Get-RubrikM365SizingInfo.ps1](https://github.com/rubrikinc/microsoft-365-sizing/blob/main/Get-RubrikM365SizingInfo.ps1) file to your local machine
2. Install the `Microsoft.Graph.Authenication` and `Microsoft.Graph.Reports` PowerShell module from the PowerShell Gallery

```powershell
Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Reports
```

## Usage

1. Open a PowerShell terminal and navigate to the folder/directory where you previously downloaded the [Get-RubrikM365SizingInfo.ps1](https://github.com/rubrikinc/microsoft-365-sizing/blob/main/Get-RubrikM365SizingInfo.ps1) file.
2. Run the script.


> NOTE - If you receive a PowerShell execution policy error message you can run the following command:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

3. Authenticate and acknowledge report access permissions in the browser window/tab that appears.

> Note: There is a known issue with the Microsoft authentication process that may result in an error message during the initial authentication process. If this occurs, re-run the script and the error will no longer show.
5. The script will run and the results will be written to a text file in the directory in which it was run. .\RubrikMS365Sizing.txt

## Example Output

```
Exchange

Name                         Value
----                         -----
NumberOfUsers                296
TotalSizeGB                  1.26
SizePerUserGB                0
AverageGrowthPercentage      8
OneYearStorageForecastInGB   1.3608
ThreeYearStorageForecastInGB 1.5624

==========================================================================
OneDrive

Name                         Value
----                         -----
NumberOfUsers                308
TotalSizeGB                  3139.39
SizePerUserGB                10.19
AverageGrowthPercentage      912
OneYearStorageForecastInGB   31770.6268
ThreeYearStorageForecastInGB 89033.1004

==========================================================================
Sharepoint

Name                         Value
----                         -----
NumberOfSites                17
TotalSizeGB                  4.24
SizePerUserGB                0.25
AverageGrowthPercentage      15
OneYearStorageForecastInGB   4.876
ThreeYearStorageForecastInGB 6.148

==========================================================================
Licensing

Name                         Value
----                         -----
MICROSOFT 365 BUSINESS BASIC 296

==========================================================================
TotalRubrikStorageNeeded

Name          Value
----          -----
OneYearInGB   31776.8636
ThreeYearInGB 89040.8108

==========================================================================
```
