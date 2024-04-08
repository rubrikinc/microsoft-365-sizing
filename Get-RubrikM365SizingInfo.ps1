<#
.SYNOPSIS
    Get-RubrikM365SizingInfo.ps1 returns M365 usage information for a subscription.
.DESCRIPTION
    Get-RubrikM365SizingInfo.ps1 returns M365 usage information for a subscprtion.
    Data is gathered using the Microsoft Graph APIs and Exchange module.

    The usage data should be similar to the metrics in the Admin Center under each
    workload's "Usage Reports" data. There could be some discrepency between the
    "Total Users" shown in the Usage Charts and what the script gathers because the
    script uses the detailed user .CSV and summarizes that information.

    The M365 Usage Reports do not contain information on Exchange In Place Archives.
    By default, the script will try to gather this information by looping through
    each user that has an In Place Archive and gathering that info directly. Unfortunately,
    if there are a lot of users, this may time out the script. If that is the case,
    you can use the flag to skip gathering in place archive data and try to provide
    an estimate.

.EXAMPLE
    PS C:\> .\Get-RubrikM365SizingInfo.ps1
    Opens a browser window to authenticate to M365 Graph APIs and Microsoft Exchange
    Module to pull usage information.

    PS C:\> .\Get-RubrikM365SizingInfo.ps1 -AnnualGrowth 35
    Calculate sizing to include a 35% annual growth rate

    PS C:\> .\Get-RubrikM365SizingInfo.ps1 -SkipArchiveMailbox $true
    Skip gathering In Place Archive mailboxes.

    PS C:\> .\Get-RubrikM365SizingInfo.ps1 -ADGroup <ad_group_name>
    Gather user info for only the AD Group specified.
.NOTES
    Author:         Chris Lumnah
    Created Date:   6/17/2021
    Updated: 4/7/24
    By: Steven Tong
#>

[CmdletBinding()]
param (
    [Parameter()]
    [bool]$EnableDebug = $false,
    # Provide your estimated annual growth rate, eg '10' for 10%
    [Parameter(HelpMessage="Estimated annual growth rate, eg 10 for 10%")]
    [int]$AnnualGrowth,
    # Provide AD Group name if you only want to gather data for a specific AD Group
    [Parameter()]
    [String]$ADGroup,
    # If gathering AD Group, CSV file to output AD Group membership info
    [Parameter()]
    [String]$ADGroupCSVFilename = './adgrouplist.csv',
    # Whether or not to skip gathering archived mailboxes, which can timeout
    [Parameter()]
    [bool]$SkipArchiveMailbox = $false,
    # Number of days to get historical stats for: 7, 30, 90, 180
    [Parameter()]
    [Int]$Period = 180
)

$date = Get-Date
$dateString = $date.ToString("yyyy-MM-dd")

# Filename to export the html report to
$outFilename = "./Rubrik-M365-Sizing-$dateString.html"

# Folder to export CSVs to
$ExportFolder = '.'

$Version = "6.0"

$ProgressPreference = 'SilentlyContinue'

# Define the capacity metric conversions
$GB = 1000000000
$GiB = 1073741824
$TB = 1000000000000
$TiB = 1099511627776

# Set which capacity metric to use
$capacityMetric = $GB
$capacityDisplay = 'GB'

Write-Host "Starting the Rubrik Microsoft 365 sizing script ($Version)."

$ExchangeHTMLTitle = "User"

if ($AnnualGrowth -eq '') {
  $AnnualGrowth = 30
}

# Function to use Graph APIs to download report info
Function Get-MgReport {
  [CmdletBinding()]
  param (
    # MS Graph API report name
    [Parameter(Mandatory)]
    [String]$ReportName,
    # Report Period (Days)
    [Parameter()]
    [ValidateSet("7", "30", "90", "180")]
    [String]$Period
  )
  try {
    if ($reportName -eq 'getMailboxUsageDetail' -or $reportName -eq 'getMailboxUsageStorage') {
      $graphApiVersion = "Beta"
    } else {
      $graphApiVersion = "Beta"
    }
    if ($Period -ne '') {
      Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/$($graphApiVersion)/reports/$($ReportName)(period=`'D$($Period)`')" -OutputFilePath "$ExportFolder\$ReportName.csv"
    } else {
      Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/$($graphApiVersion)/reports/$($ReportName)" -OutputFilePath "$ExportFolder\$ReportName.csv"
    }
    return "$ExportFolder\$ReportName.csv"
  }
  catch {
    $errorMessage = $_.Exception | Out-String
    if ($errorMessage.Contains('Response status code does not indicate success: Forbidden (Forbidden)')) {
      Disconnect-MgGraph
      throw "The user account used for authentication must have permissions covered by Reports Reader admin role."
    }
    throw $_.Exception
  }
}

# Function to calculate annual growth based on historical report
# This may not be a good representation of actual growth in an environment
# if many changes are happening - for example, a migration event which would
# inflate steady state growth numbers.
function Measure-AverageGrowth {
  param (
    [Parameter(Mandatory)]
    [string]$ReportCSV,
    [Parameter(Mandatory)]
    [string]$ReportName
  )
  $UsageReport = Import-Csv -Path $ReportCSV | Sort-Object -Property "Report Date" -Descending
  $ReportDays = $UsageReport[0].'Report Period'
  $LatestUsageGB = [math]::Round($UsageReport[0].'Storage Used (Byte)' / 1GB, 2)
  $EarliestUsageGB = [math]::Round($UsageReport[-1].'Storage Used (Byte)' / 1GB, 2)
  $GrowthOverPeriod = [math]::Round($LatestUsageGB - $EarliestUsageGB, 2)
  $AvgGrowthPerDay = $GrowthOverPeriod / $ReportDays
  $GrowthPerYearGB = [math]::Round($AvgGrowthPerDay * 365, 2)
  $GrowthPerYearPct = [math]::Round($GrowthPerYearGB / $LatestUsageGB, 2)
  Write-Host "$ReportName historical storage usage:"
  Write-Host "  - Usage on $($UsageReport[0].'Report Date'): $LatestUsageGB GB"
  Write-Host "  - Usage on $($UsageReport[-1].'Report Date'): $EarliestUsageGB GB"
  Write-Host "  - Growth over $ReportDays days: $GrowthOverPeriod GB"
  Write-Host "  - Growth annualized per year: $GrowthPerYearGB GB, $($GrowthPerYearPct * 100)%"
  return $GrowthPerYearPct
}


# Function to solve for licenses required
# Query M365Licsolver Azure Function
function Solve-License {
  param (
    [Parameter(Mandatory)]
    [int]$userLicense,
    [Parameter(Mandatory)]
    [int]$storageGB
  )
  # If less than 76GB Average per user then query the azure function that calculates the best mix of subscription types. If more than 76 then Unlimited is the best option.
  if (($storageGB) / $userLicense -le 76) {
    # Query the M365Licsolver Azure Function
    $SolverQuery = '{"users":"' + $userLicense + '","data":"' + $storageGB + '"}'
    try {
      $APIReturn = ConvertFrom-JSON (Invoke-WebRequest 'https://m365licsolver-azure.azurewebsites.net:/api/httpexample' -ContentType "application/json" -Body $SolverQuery -UseBasicParsing -Method 'POST')
    }
    catch {
      $errorMessage = $_.Exception | Out-String
      if ($errorMessage.Contains('Response status code does not indicate success: 404')) {
        Write-Host "[Info] Unable to calculate license recommendations."
      }
    }
    $FiveGBUsers = $APIReturn.FiveGBSubscriptions * 10
    $TwentyGBUsers = $APIReturn.TwentyGBSubscriptions * 10
    $FiftyGBUsers = $APIReturn.FiftyGBSubscriptions * 10
    $UnlimitedGBUsers = 0
  } else {
    $FiveGBUsers = 0
    $TwentyGBUsers = 0
    $FiftyGBUsers = 0
    $UnlimitedGBPacks = [math]::ceiling($userLicense / 10)
    $UnlimitedGBUsers = $UnlimitedGBPacks * 10
  }
  $licenseRequired = [PSCustomObject] @{
    "FiveGBUsers" = $FiveGBUsers
    "TwentyGBUsers" = $TwentyGBUsers
    "FiftyGBUsers" = $FiftyGBUsers
    "UnlimitedGBUsers" = $UnlimitedGBUsers
  }
  return $licenseRequired
}


# Validate that Period (days) for historical reports is valid
# Must be: 7, 30, 90, or 180
$PeriodValues = @(7, 30, 90, 180)
if ($Period -in $PeriodValues) {
} else {
    throw "Error: Period (days) needs to be: 7, 30, 90, or 180"
}

# Validate the required 'Microsoft.Graph.Reports' is installed
# and provide a user friendly message when it's not.
if (Get-Module -ListAvailable -Name Microsoft.Graph.Reports) {
} else {
  throw "The 'Microsoft.Graph.Reports' module is required for this script. Run the follow command to install: Install-Module Microsoft.Graph.Reports"
}

# Validate the required 'ExchangeOnlineManagement' is installed and provide a user friendly message when it's not.
if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
} else {
  throw "The 'ExchangeOnlineManagement' module is required for this script. Run the follow command to install: Install-Module ExchangeOnlineManagement"
}

$AzureAdRequired = $PSBoundParameters.ContainsKey('ADGroup')

if ($AzureAdRequired) {
  # Validate the required 'Azure.Graph.Authentication' is installed
  # and provide a user friendly message when it's not.
  if (Get-Module -ListAvailable -Name Microsoft.Graph.Groups) {
  } else {
    throw "The 'Microsoft.Graph.Groups' module is required for filtering by a specific Azure AD Group. Run the follow command to install: Install-Module Microsoft.Graph.Groups"
  }
}

Write-Host "[INFO] Connecting to the Microsoft Graph API using 'Reports.Read.All', 'User.Read.All', and 'Group.Read.All' permissions."
try {
  Connect-MgGraph -Scopes "Reports.Read.All", "User.Read.All", "Group.Read.All"  | Out-Null
}
catch {
  $errorException = $_.Exception
  $errorMessage = $errorException.Message
  Write-Host "[ERROR] Unable to Connect to the Microsoft Graph PowerShell Module: $errorMessage"
}

if ($SkipArchiveMailbox -eq $false) {
  Write-Host "Connecting to the Microsoft Exchange Online Module to gather per-mailbox In Place Archive stats."
  try {
    Connect-ExchangeOnline -ShowBanner:$false
  } catch {
    $errorException = $_.Exception
    $errorMessage = $errorException.Message
    Write-Host "[ERROR] Unable to Connect to the Microsoft Exchange PowerShell Module: $errorMessage"
  }
}

# If AD Group is provided, get the AD Group membership info
if ($AzureAdRequired) {
  Write-Host "Looking up AD Group users in: $ADGroup" -foregroundcolor green
  $AzureAdGroupDetails = Get-MgGroup -Filter "DisplayName eq '$ADGroup'"
  if ($AzureAdGroupDetails.Count -eq 0) {
    throw "The Azure AD Group '$ADGroup' does not exist. Exiting script."
  }
  $AzureAdGroupMembersById = Get-MgGroupMember -GroupId $AzureAdGroupDetails.Id -All
  if ($EnableDebug) {
    Write-Host "[DEBUG] Azure AD Group Members Size: $($AzureAdGroupMembersById.Count)"
  }
  $AzureAdGroupMembersByUserPrincipalName = @()
  $AzureAdGroupMembersById | Foreach-Object {
    if ($_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.user") {
      $AzureAdGroupMembersByUserPrincipalName += $_.AdditionalProperties["userPrincipalName"]
    }
  }
  if ($AzureAdGroupMembersByUserPrincipalName.Count -eq 0) {
    throw "The Azure AD Group '$ADGroup' does not contain any User Principal Names."
  }
  Write-Host "# of Azure AD Group members found: $($AzureAdGroupMembersByUserPrincipalName.Count)" -foregroundcolor green
  Write-Host "AD Group user principal names exported to CSV: $ADGroupCSVFilename" -foregroundcolor green
  $AzureAdGroupMembersByUserPrincipalName | Out-File -Path $ADGroupCSVFilename
}

if ($EnableDebug) {
  try {
    $user = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/me"
    $permissions = Get-MgUserOauth2PermissionGrant -UserId $user.id
    Write-Host "[DEBUG] The authenticated user account has the following permissions:$($permissions.Scope)"
  }
  catch {
    $errorMessage = $_.Exception | Out-String
    throw $_.Exception
  }
}

Write-Host ""
Write-Host "The data gathered here is the same as in the M365 Admin Center" -foreground green
Write-Host "under Reports -> Usage -> <Workload> -> Usage Reports"
Write-Host "Microsoft metrics could differ a bit depending on what you are looking at"
Write-Host ""

### Exchange - Get reports for Exchange and process them
Write-Host "*** Retrieving usage info for: Exchange ***" -foregroundcolor green
Write-Host "Data will be gathered from the chart data and per-mailbox usage report" -foregroundcolor green

Write-Host "Getting chart data - this should match the chart for the Users drop-down and Total mailboxes" -foregroundcolor green
$ReportCSV = Get-MgReport -ReportName 'getMailboxUsageMailboxCounts' -Period $Period
Write-Host "Exchange chart CSV saved to: $ReportCSV" -foregroundcolor green
$ExchangeChartData = Import-Csv -Path $ReportCSV
Write-Host "Total # of User (non-shared mailboxes) - chart data: $($ExchangeChartData[0].total)" -foregroundcolor green
Write-Host ""

# Get Exchange per user usage report
Write-Host "Getting the per-mailbox usage report - this should match the details in the Admin Center reports" -foregroundcolor green
$ReportCSV = Get-MgReport -ReportName 'getMailboxUsageDetail' -Period $Period
Write-Host "Exchange per-mailbox usage CSV saved to: $ReportCSV" -foregroundcolor green

$ExchangeUsageReport = Import-Csv -Path $ReportCSV
$ExchangeTotalUsersCount = $ExchangeUsageReport.count
$ExchangeNonDeletedUsersCount = $($ExchangeUsageReport | Where-Object { $_.'Is Deleted' -eq 'FALSE' }).count
$ExchangeDeletedUsersCount = $($ExchangeUsageReport | Where-Object { $_.'Is Deleted' -eq 'TRUE' }).count

Write-Host "Total # of mailboxes - from usage report: $ExchangeTotalUsersCount"
Write-Host "Total # of deleted mailboxes - from usage report: $ExchangeDeletedUsersCount"
Write-Host "Total # of active (non-deleted) mailboxes - from usage report: $ExchangeNonDeletedUsersCount"
Write-Host "Now performing additional filtering..."
Write-Host ""

# List of all active (non-deleted) user mailboxes
$ExchangeUsageReportUsers = $ExchangeUsageReport | Where-Object { $_.'Is Deleted' -eq 'FALSE' -and
  $_.'Recipient Type' -ne 'Shared'}

# List of all active (non-deleted) shared mailboxes
$ExchangeUsageReportShared = $ExchangeUsageReport | Where-Object { $_.'Is Deleted' -eq 'FALSE' -and
  $_.'Recipient Type' -eq 'Shared'}

if ($AzureAdRequired) {
  Write-Host "Filtering user mailboxes by Azure AD Group: $ADGroup"
  $FilterByField = "User Principal Name"
  $ExchangeUsageReportUsers = $ExchangeUsageReportUsers | Where-Object { $_.$FilterByField -in $AzureAdGroupMembersByUserPrincipalName }
  # If we didn't get any usage for the Azure AD group users, it might be because the reports are masking User IDs
  if ($ExchangeUsageReportUsers.count -eq 0) {
    Write-Host "[ERROR] Did not match any Azure AD Group users to the usage reports" -foregroundcolor red
    Write-Host "[ERROR] Check the Azure AD group csv ($ADGroupCSVFilename) to see what users are part of the Azure AD Group" -foregroundcolor red
    Write-Host "[ERROR] Check the mailbox csv ($reportCSV) to see if 'User Principal Name' are being masked" -foregroundcolor red
    Write-Host "[ERROR] Un-mask by going to M365 Admin Center -> Settings -> Org Settings -> Services"  -foregroundcolor red
    Write-Host "[ERROR] Then click on Reports and clear: Display concealed user, group, and site names in all reports, and then select Save"  -foregroundcolor red
    Write-Host "[ERROR] See: https://learn.microsoft.com/en-us/microsoft-365/troubleshoot/miscellaneous/reports-show-anonymous-user-name" -foregroundcolor red
    throw "Error running script with Azure AD group option - could not find any matching users. Exiting script."
  }
  Write-Host "Matched $($ExchangeUsageReportUsers.count) users in the provided Azure AD Group" -foregroundcolor green
}

# Calculate metrics for user mailboxes
$userMailboxStorageSum = $ExchangeUsageReportUsers | Measure-Object -Property 'Storage Used (Byte)' -Sum
$userMailboxStorageSumDisplay = [math]::Round($userMailboxStorageSum.Sum / $capacityMetric, 2)
$userMailboxItems = $ExchangeUsageReportUsers | Measure-Object -Property 'Item Count' -Sum

# Calculate metrics for shared mailboxes
$sharedMailboxStorageSum = $ExchangeUsageReportShared | Measure-Object -Property 'Storage Used (Byte)' -Sum
$sharedMailboxStorageSumDisplay = [math]::Round($sharedMailboxStorageSum.Sum / $capacityMetric, 2)
$sharedMailboxItems = $ExchangeUsageReportShared | Measure-Object -Property 'Item Count' -Sum

Write-Host "Total # of active user mailboxes - from usage report: $($ExchangeUsageReportUsers.count)" -foregroundcolor green
Write-Host "Active users storage used (not including in-place archve): $userMailboxStorageSumDisplay $capacityDisplay" -foregroundcolor green
Write-Host "Active users item count (not including in-place archve): $($userMailboxItems.sum)" -foregroundcolor green
Write-Host "Total # of active shared mailboxes - from usage report: $($ExchangeUsageReportShared.count)" -foregroundcolor green
Write-Host "Shared mailboxes storage used: $sharedMailboxStorageSumDisplay $capacityDisplay" -foregroundcolor green
Write-Host "Shared mailboxes item count (not including in-place archve): $($sharedMailboxItems.sum)" -foregroundcolor green
Write-Host ""

Write-Host "Getting historical Exchange storage growth" -foregroundcolor green
$ReportCSV = Get-MgReport -ReportName 'getMailboxUsageStorage' -Period $Period
$CalculatedGrowth = Measure-AverageGrowth -ReportCSV $ReportCSV -ReportName 'Exchange'

$ExchangeDetails = [PSCustomObject] @{
  "Chart User Mailboxes" = $ExchangeChartData[0].total
  "Total Mailboxes" = $($ExchangeUsageReportUsers.count + $ExchangeUsageReportShared.count)
  "User Mailboxes" = $ExchangeUsageReportUsers.count
  "User Storage Used No Archive" = $userMailboxStorageSum.sum
  "User Items No Archive" = $userMailboxItems.sum
  "Shared Mailboxes" = $ExchangeUsageReportShared.count
  "Shared Storage Used No Archive" = $sharedMailboxStorageSum.sum
  "Shared Items No Archive" = $sharedMailboxItems.sum
  "Total Storage Used" = $($userMailboxStorageSum.sum + $sharedMailboxStorageSum.sum)
  "Total Items" = $($userMailboxItems.sum + $sharedMailboxItems.sum)
  "Calculated Growth %" = $CalculatedGrowth
}

### OneDrive - Get reports for OneDrive and process them
Write-Host "*** Retrieving usage info for: OneDrive ***" -foregroundcolor green
Write-Host "Data will be gathered from the chart data and per-account usage report" -foregroundcolor green

Write-Host "Getting chart data - this should match the chart for Total Accounts" -foregroundcolor green
$ReportCSV = Get-MgReport -ReportName 'getOneDriveUsageAccountCounts' -Period $Period
Write-Host "OneDrive account chart CSV saved to: $ReportCSV" -foregroundcolor green
$OneDriveChartAccount = Import-Csv -Path $ReportCSV
Write-Host "Total # of OneDrive accounts - chart data: $($OneDriveChartAccount[0].total)" -foregroundcolor green
Write-Host ""

Write-Host "Getting chart data - this should match the chart for Total Storage" -foregroundcolor green
$ReportCSV = Get-MgReport -ReportName 'getOneDriveUsageStorage' -Period $Period
Write-Host "OneDrive storage chart CSV saved to: $ReportCSV" -foregroundcolor green
$OneDriveChartStorage = Import-Csv -Path $ReportCSV
$OneDriveChartStorageUsed = [math]::round($OneDriveChartStorage[0].'Storage Used (Byte)' / $capacityMetric, 2)
Write-Host "Total # of OneDrive storage - chart data: $oneDriveChartStorageUsed $capacityDisplay" -foregroundcolor green
Write-Host ""

Write-Host "Getting chart data - this should match the chart for Total Files" -foregroundcolor green
$ReportCSV = Get-MgReport -ReportName 'getOneDriveUsageFileCounts' -Period $Period
Write-Host "OneDrive files chart CSV saved to: $ReportCSV" -foregroundcolor green
$OneDriveChartFiles = Import-Csv -Path $ReportCSV
Write-Host "Total # of OneDrive files - chart data: $($OneDriveChartFiles[0].total)" -foregroundcolor green
Write-Host ""


Write-Host "Getting the per-account usage report - this should match the details in the Admin Center reports" -foregroundcolor green
Write-Host "However, this count may not match the Total count in the chart" -foregroundcolor green
$ReportCSV = Get-MgReport -ReportName 'getOneDriveUsageAccountDetail' -Period $Period
Write-Host "OneDrive per-account usage CSV saved to: $ReportCSV" -foregroundcolor green

$OneDriveUsageReport = Import-Csv -Path $ReportCSV
$OneDriveTotalUsersCount = $OneDriveUsageReport.count
$OneDriveNonDeletedUsersCount = $($OneDriveUsageReport | Where-Object { $_.'Is Deleted' -eq 'FALSE' }).count
$OneDriveDeletedUsersCount = $($OneDriveUsageReport | Where-Object { $_.'Is Deleted' -eq 'TRUE' }).count

Write-Host "Total # of accounts - from usage report: $($OneDriveUsageReportAccounts.count)"
Write-Host "Total # of deleted accounts - from usage report: $OneDriveDeletedUsersCount"
Write-Host "Total # of active (non-deleted) accounts - from usage report: $OneDriveNonDeletedUsersCount"
Write-Host "Now performing additional filtering..."
Write-Host ""

$OneDriveUsageReportAccounts = $OneDriveUsageReport | Where-Object { $_.'Is Deleted' -eq 'FALSE' }

if ($AzureAdRequired) {
  Write-Host "Filtering user accounts by Azure AD Group: $ADGroup"
  $FilterByField = "Owner Principal Name"
  $OneDriveUsageReportAccounts = $OneDriveUsageReportAccounts | Where-Object { $_.$FilterByField -in $AzureAdGroupMembersByUserPrincipalName }
  # If we didn't get any usage for the Azure AD group users, it might be because the reports are masking User IDs
  if ($OneDriveUsageReportAccounts.count -eq 0) {
    Write-Host "[ERROR] Did not match any Azure AD Group users to the usage reports" -foregroundcolor red
    Write-Host "[ERROR] Check the Azure AD group csv ($ADGroupCSVFilename) to see what users are part of the Azure AD Group" -foregroundcolor red
    Write-Host "[ERROR] Check the OneDrive csv ($reportCSV) to see if 'Owner Principal Name' are being masked" -foregroundcolor red
    Write-Host "[ERROR] Un-mask by going to M365 Admin Center -> Settings -> Org Settings -> Services"  -foregroundcolor red
    Write-Host "[ERROR] Then click on Reports and clear: Display concealed user, group, and site names in all reports, and then select Save"  -foregroundcolor red
    Write-Host "[ERROR] See: https://learn.microsoft.com/en-us/microsoft-365/troubleshoot/miscellaneous/reports-show-anonymous-user-name" -foregroundcolor red
    throw "Error running script with Azure AD group option - could not find any matching users. Exiting script."
  }
  Write-Host "Matched $($OneDriveUsageReportAccounts.count) users in the provided Azure AD Group" -foregroundcolor green
}

# Calculate metrics for OneDrive accounts
$oneDriveStorageSum = $OneDriveUsageReportAccounts | Measure-Object -Property 'Storage Used (Byte)' -Sum
$oneDriveStorageSumDisplay = [math]::Round($oneDriveStorageSum.Sum / $capacityMetric, 2)
$oneDriveFiles = $OneDriveUsageReportAccounts | Measure-Object -Property 'File Count' -Sum

Write-Host "Total # of Accounts - from usage report: $($OneDriveUsageReportAccounts.count)" -foregroundcolor green
Write-Host "Accounts storage used: $oneDriveStorageSumDisplay $capacityDisplay" -foregroundcolor green
Write-Host "Accounts file count: $($oneDriveFiles.sum)" -foregroundcolor green
Write-Host ""

Write-Host "Getting historical OneDrive storage growth" -foregroundcolor green
$ReportCSV = Get-MgReport -ReportName 'getOneDriveUsageStorage' -Period $Period
$CalculatedGrowth = Measure-AverageGrowth -ReportCSV $ReportCSV -ReportName 'Exchange'

$OneDriveDetails = [PSCustomObject] @{
  "Chart Accounts" = $OneDriveChartAccount[0].total
  "Chart Storage Used" = $OneDriveChartStorage[0].'Storage Used (Byte)'
  "Chart Files" = $OneDriveChartFiles[0].total
  "Usage Accounts" = $OneDriveUsageReportAccounts.count
  "Usage Account Storage Used" = $oneDriveStorageSum.sum
  "Usage Account Files" = $oneDriveFiles.sum
  "Calculated Growth %" = $CalculatedGrowth
}

# If AD Group is specified, assume each user is licensed so use that count
if ($AzureAdRequired) {
  $OneDriveDetails | Add-Member -MemberType NoteProperty -Name 'Accounts' -Value $OneDriveDetails.'Usage Accounts'
  $OneDriveDetails | Add-Member -MemberType NoteProperty -Name 'Storage Used' -Value $OneDriveDetails.'Usage Account Storage Used'
  $OneDriveDetails | Add-Member -MemberType NoteProperty -Name 'Total Files' -Value $OneDriveDetails.'Usage Account Files'
} else {
  # Otherwise, use the counts from the Chart
  $OneDriveDetails | Add-Member -MemberType NoteProperty -Name 'Accounts' -Value $OneDriveDetails.'Chart Accounts'
  $OneDriveDetails | Add-Member -MemberType NoteProperty -Name 'Storage Used' -Value $OneDriveDetails.'Chart Storage Used'
  $OneDriveDetails | Add-Member -MemberType NoteProperty -Name 'Total Files' -Value $OneDriveDetails.'Chart Files'
}

### SharePoint - Get reports for SharePoint and process them
Write-Host "*** Retrieving usage info for: SharePoint ***" -foregroundcolor green
Write-Host "Data will be gathered site usage report" -foregroundcolor green

Write-Host "Getting site usage report" -foregroundcolor green
$ReportCSV = Get-MgReport -ReportName 'getSharePointSiteUsageDetail' -Period $Period
Write-Host "SharePoint site usage CSV saved to: $ReportCSV" -foregroundcolor green
$SharePointUsageReport = Import-Csv -Path $ReportCSV

# Calculate metrics for SharePoint Sites
$SharePointSitesCount = $SharePointUsageReport.count
$SharePointNonDeletedSitesCount = $($SharePointUsageReportSites | Where-Object { $_.'Is Deleted' -eq 'FALSE' }).count
$SharePointDeletedSitesCount = $($SharePointUsageReportSites | Where-Object { $_.'Is Deleted' -eq 'TRUE' }).count

Write-Host "Total # of sites - from usage report: $SharePointSitesCount"
Write-Host "Total # of deleted mailboxes - from usage report: $SharePointDeletedSitesCount"
Write-Host "Total # of active (non-deleted) mailboxes - from usage report: $SharePointNonDeletedSitesCount"
Write-Host "Now performing additional filtering..."
Write-Host ""

$SharePointUsageReportSites = $SharePointUsageReport | Where-Object { $_.'Is Deleted' -eq 'FALSE' }

# Calculate metrics for SharePoint sites
$sharePointStorageSum = $SharePointUsageReportSites | Measure-Object -Property 'Storage Used (Byte)' -Sum
$sharePointStorageSumDisplay = [math]::Round($sharePointStorageSum.Sum / $capacityMetric, 2)
$sharePointFiles = $SharePointUsageReportSites | Measure-Object -Property 'File Count' -Sum

Write-Host "Total # of SharePoint Sites - from usage report: $($SharePointUsageReportSites.count)" -foregroundcolor green
Write-Host "SharePoint Sites storage used: $sharePointStorageSumDisplay $capacityDisplay" -foregroundcolor green
Write-Host "SharePoint Sites file count: $($sharePointFiles.sum)" -foregroundcolor green
Write-Host ""

Write-Host "Getting historical OneDrive storage growth" -foregroundcolor green
$ReportCSV = Get-MgReport -ReportName 'getSharePointSiteUsageStorage' -Period $Period
$CalculatedGrowth = Measure-AverageGrowth -ReportCSV $ReportCSV -ReportName 'Exchange'

$SharePointDetails = [PSCustomObject] @{
  "Sites" = $SharePointUsageReportSites.count
  "Sites Storage Used" = $sharePointStorageSum.sum
  "Account Files" = $sharePointFiles.sum
  "Calculated Growth %" = $CalculatedGrowth
}

Write-Host "[INFO] Disconnecting from the Microsoft Graph API."
Disconnect-MgGraph

# The Microsoft Exchange Reports do not contain In-Place Archive sizing information.DESCRIPTION
# We need to connect to the Exchange Online module to get this information

if ($SkipArchiveMailbox -eq $true) {
  Write-Host "Skipping gathering In Place Archive usage" -foregroundcolor green
} else {
  Write-Host "Now gathering In Place Archive usage" -foregroundcolor green
  Write-Host "This may take awhile since stats need to be gathered per user" -foregroundcolor green
  Write-Host "Progress will be written as they are gathered" -foregroundcolor green
  Write-Host "If this keeps timing out, run script with -SkipArchiveMailbox $true option" -foregroundcolor green
  $ConnectionUserPrincipalName = $(Get-ConnectionInformation).UserPrincipalName
  # $ActionRequiredLogMessage = "[ACTION REQUIRED] In order to periodically refresh the connection to Microsoft, we need the User Principal Name used during the authentication process."
  # $ActionRequiredPromptMessage = "Enter the User Principal Name"
  $FirstInterval = 500
  $SkipInternval = $FirstInterval
  $ArchiveMailboxSizeGb = 0
  $LargeAmountofArchiveMailboxCount = 5000
  $FilterByField = 'User Principal Name'
  Write-Host "[INFO] Retrieving all Exchange Mailbox In-Place Archive sizing"
  # Get a list of all users with In Place Archive mailboxes in the tenant
  # $ArchiveMailboxes = Get-ExoMailbox -Archive -ResultSize Unlimited
  $ArchiveMailboxes = $ExchangeUsageReportUsers | Where-Object { $_.'Has Archive' -eq 'TRUE' }
  $ArchiveMailboxesCount = $ArchiveMailboxes.Count
  $ArchiveMailboxList = @()
  $CurrentMailboxNum = 0
  Write-Host "Found $ArchiveMailboxesCount mailboxes with In Place Archives" -foregroundcolor green
  do {
    if ( ($CurrentMailboxNum % 10) -eq 0 ) {
      Write-Host "[$CurrentMailboxNum / $ArchiveMailboxesCount] Processing mailboxes ..."
    }
    $CurrentUser = $ArchiveMailboxes[$CurrentMailboxNum].'User Principal Name'
    try {
      $ArchiveMailboxStats = Get-EXOMailboxStatistics -Archive -Identity $CurrentUser
      $MatchArchiveSize = $ArchiveMailboxStats.TotalItemSize -match '\(([^)]+) bytes\)'
      $ArchiveSize = [long]($Matches[1] -replace ',', '')
      $ArchiveStats = [PSCustomObject] @{
        "UserPrincipalName" = $CurrentUser
        "ArchiveSize" = $ArchiveSize
        "ArchiveItems" = $ArchiveMailboxStats.ItemCount
      }
      $ArchiveMailboxList += $ArchiveStats
    } catch {
      Write-Error "Error getting info for mailbox: $CurrentUser"
    }
    $CurrentMailboxNum += 1
  } while ($CurrentMailboxNum -lt $ArchiveMailboxesCount)
  $ArchiveMeasurementSize = $ArchiveMailboxList | Measure-Object -Property 'ArchiveSize' -Sum -Average
  $ArchiveMeasurementItems = $ArchiveMailboxList | Measure-Object -Property 'ArchiveItems' -Sum -Average
  $TotalArchiveSize = [math]::Round($($ArchiveMeasurementSize.Sum / $capacityMetric), 2)
  $TotalArchiveItems = $ArchiveMeasurementItems.Sum
  Write-Host "Finished gathering stats on mailboxes with In Place Archive" -foregroundcolor green
  Write-Host "Total # of mailboxes with In Place Archive: $ArchiveMailboxesCount" -foregroundcolor green
  Write-Host "Total size of mailboxes with In Place Archive: $TotalArchiveSize $capacityDisplay" -foregroundcolor green
  Write-Host "Total # of items of mailboxes with In Place Archive: $TotalArchiveItems" -foregroundcolor green
  Write-Host "Disconnecting from the Microsoft Exchange Online Module"
  Write-Host ""
  Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
}

if ($SkipArchiveMailbox -eq $false) {
  $ExchangeDetails | Add-Member -MemberType NoteProperty -Name 'Archive Mailboxes' -Value $ArchiveMailboxesCount
  $ExchangeDetails | Add-Member -MemberType NoteProperty -Name 'Archive Storage Used' -Value $ArchiveMeasurementSize.Sum
  $ExchangeDetails | Add-Member -MemberType NoteProperty -Name 'Archive Items' -Value $ArchiveMeasurementItems.Sum
  $ExchangeTotalStorage = $ExchangeDetails.'Total Storage Used' + $ArchiveMeasurementSize.Sum
  $ExchangeDetails.'Total Storage Used' = $ExchangeTotalStorage
  $ExchangeTotalItems = $ExchangeDetails.'Total Items' + $ArchiveMeasurementItems.Sum
  $ExchangeDetails.'Total Items' = $ExchangeTotalItems
} else {
  $ExchangeDetails | Add-Member -MemberType NoteProperty -Name 'Archive Mailboxes' -Value 'Skipped'
  $ExchangeDetails | Add-Member -MemberType NoteProperty -Name 'Archive Storage Used' -Value '-'
  $ExchangeDetails | Add-Member -MemberType NoteProperty -Name 'Archive Items' -Value '-'
}

Write-Host "Calculating # of license needed:"
Write-Host "Exchange user mailboxes: $($ExchangeDetails.'User Mailboxes')"
Write-Host "Exchange shared mailboxes: $($ExchangeDetails.'Shared Mailboxes')"
Write-Host "OneDrive chart accounts: $($OneDriveDetails.'Chart Accounts')"
$UserLicensesRequired = $($ExchangeDetails.'User Mailboxes')
if ($ExchangeDetails.'Shared Mailboxes' -gt $UserLicensesRequired) {
  $UserLicensesRequired = $ExchangeDetails.'Shared Mailboxes'
}
if ($OneDriveDetails.'Chart Accounts' -gt $UserLicensesRequired) {
  $UserLicensesRequired = $OneDriveDetails.'Chart Accounts'
}
Write-Host "# of licenses required: $UserLicensesRequired" -foregroundcolor green

$totalStorage = $ExchangeDetails.'Total Storage Used' + $OneDriveDetails.'Storage Used' +
  $SharePointDetails.'Sites Storage Used'
$totalItems = $ExchangeDetails.'Total Items' + $OneDriveDetails.'Total Files' +
  $SharePointDetails.'Account Files'

$totalStorageGB = [math]::round($totalStorage / 1GB, 2)

[int]$UserLicenseRequired10 = [int]$UserLicensesRequired * 1.1
[int]$UserLicenseRequired20 = [int]$UserLicensesRequired * 1.2
[int]$UserLicenseRequiredCustom = [int]$UserLicensesRequired * (1 + ($AnnualGrowth / 100))

$totalStorage10GB = [math]::round($totalStorage / 1GB * 1.1, 2)
$totalStorage20GB = [math]::round($totalStorage / 1GB * 1.2, 2)
$totalStorageCustomGB = [math]::round($totalStorage / 1GB * (1 + ($AnnualGrowth / 100)), 2)

$growth10 = Solve-License -userLicense $UserLicenseRequired10 -storageGB $totalStorage10GB
$growth20 = Solve-License -userLicense $UserLicenseRequired20 -storageGB $totalStorage20GB
$growthCustom = Solve-License -userLicense $UserLicenseRequiredCustom -storageGB $totalStorageCustomGB


#region HTML Code for Output
$HTML_CODE = @"
<!DOCTYPE html>

<html>
<!---->
<!---->
<link rel="stylesheet" href="https://www.w3schools.com/w3css/4/w3.css">

<head>
    <style>
        body {
            background-color: #f4f4f4
        }

        .card-container {
            display: flex;
            width: 100%;
            align-items: center;
            justify-content: center;
            padding-bottom: 20px;
        }

        .card-header {
            display: flex;
        }

        .card-header-logo {
            flex-grow: 1;
        }

        .rubrik-snowflake {
            padding-top: 15px;
        }

        .card-header-text {
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.9rem;
            line-height: 2.4rem;
        }

        .navigation-bar {
            display: flex;
            background-color: #060745;
            width: 100%;
            top: 0;
            left: 0;
            position: fixed;
            max-height: 82px;
        }

        .logo {
            padding-top: 20px;
            padding-left: 10px;
            display: block;
            max-width: 150px;

        }

        .nav-bar-text {
            padding-top: 12px;
            flex-grow: 1;
            display: flex;
            color: white;
            align-items: center;
            justify-content: center;
            font-size: 1.9rem;
            line-height: 2.4rem;
            margin-bottom: 20px;
            margin-top: 0;

        }

        .rubrik-logo path {
            fill: white;
        }

        .margin {
            padding-bottom: 130px;
        }

        .card {
            box-shadow: 0 4px 10px 0 rgb(0 0 0 / 20%), 0 4px 20px 0 rgb(0 0 0 / 19%);
            width: 98%;
            padding: 0.01em 16px;

        }

        .styled-table {
            margin: 25px 0;
            width: 100%;

        }

        .styled-table thead tr {
            text-align: left;
        }

        .styled-table th,
        .styled-table td {
            padding: 12px 15px;
        }
    </style>
</head>

<body>
    <div class="navigation-bar">
        <div class="logo">
            <svg class="rubrik-logo" width=auto height="82">
                <defs>
                    <style>
                        .cls-1 {
                            fill: #fff
                        }

                        .cls-1,
                        .cls-2 {
                            fill-rule: evenodd
                        }
                    </style>
                    <mask id="mask" x="13.3" y="0" width="12.35" height="12.27" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-2">
                                <path id="path-1" class="cls-1"
                                    d="M19.51.22a.83.83 0 0 0-.32.2l-5.34 5.32a.84.84 0 0 0 0 1.19l5.34 5.32a.84.84 0 0 0 1.19 0l5.33-5.32a.84.84 0 0 0 0-1.19L20.38.42a.83.83 0 0 0-.32-.2h-.55z">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-2-2" x="13.3" y="26.53" width="12.35" height="12.25" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-4">
                                <path id="path-3" class="cls-1"
                                    d="M19.19 27l-5.34 5.32a.83.83 0 0 0 0 1.18l5.34 5.33a.85.85 0 0 0 .25.17h.69a.85.85 0 0 0 .25-.17l5.33-5.33a.83.83 0 0 0 0-1.18L20.38 27a.82.82 0 0 0-.6-.25.81.81 0 0 0-.59.25">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-3" x="26.6" y="13.22" width="12.35" height="12.32" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-6">
                                <path id="path-5" class="cls-1"
                                    d="M32.49 13.69L27.15 19a.86.86 0 0 0 0 1.19l5.34 5.32a.84.84 0 0 0 1.19 0L39 20.2a.84.84 0 0 0 0-1.2l-5.33-5.32a.84.84 0 0 0-.59-.24.85.85 0 0 0-.6.24">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-4-2" x="9.63" y="33.2" width="3.17" height="4.57" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-8">
                                <path id="path-7" class="cls-1"
                                    d="M12.51 33.61L10.14 36a.59.59 0 0 0 .15 1l2 1a.52.52 0 0 0 .78-.52v-3.63c0-.28-.1-.43-.25-.43a.53.53 0 0 0-.35.19">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-5" x="26.15" y="33.2" width="3.17" height="4.57" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-10">
                                <path id="path-9" class="cls-1"
                                    d="M26.46 33.85v3.56a.52.52 0 0 0 .77.52l2.05-1a.59.59 0 0 0 .14-1l-2.37-2.36a.52.52 0 0 0-.34-.19c-.15 0-.25.15-.25.43">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-6-2" x="26.15" y="26.04" width="6.49" height="6.48" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-12">
                                <path id="path-11" class="cls-1"
                                    d="M27.3 26.27a.84.84 0 0 0-.84.83v4.8a.85.85 0 0 0 .84.84h4.81a.85.85 0 0 0 .89-.84v-4.8a.84.84 0 0 0-.84-.83H27.3z">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-7" x="33.32" y="9.56" width="4.58" height="3.17" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-14">
                                <path id="path-13" class="cls-1"
                                    d="M36.19 10l-2.38 2.37c-.32.32-.21.59.25.59h3.57a.53.53 0 0 0 .52-.78l-1-2a.62.62 0 0 0-.54-.36.65.65 0 0 0-.45.21">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-8-2" x="26.15" y="1" width="3.17" height="4.57" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-16">
                                <path id="path-15" class="cls-1"
                                    d="M26.46 1.8v3.56c0 .46.26.57.59.25l2.37-2.37a.59.59 0 0 0-.14-1l-2.05-1a.55.55 0 0 0-.23-.01.5.5 0 0 0-.5.57">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-9" x="1.05" y="9.56" width="4.58" height="3.17" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-18">
                                <path id="path-17" class="cls-1"
                                    d="M2.39 10.14l-1 2a.52.52 0 0 0 .52.78H5.5c.47 0 .58-.27.25-.59L3.38 10a.65.65 0 0 0-.46-.21.59.59 0 0 0-.53.36">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-10-2" x="6.31" y="6.25" width="6.49" height="6.48" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-20">
                                <path id="path-19" class="cls-1"
                                    d="M7.46 6.47a.85.85 0 0 0-.84.84v4.8a.85.85 0 0 0 .84.84h4.81a.85.85 0 0 0 .84-.84v-4.8a.85.85 0 0 0-.84-.84H7.46z">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-11" x="9.63" y="1" width="3.17" height="4.57" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-22">
                                <path id="path-21" class="cls-1"
                                    d="M12.33 1.29l-2 1a.59.59 0 0 0-.15 1l2.37 2.37c.33.32.6.21.6-.25V1.8a.51.51 0 0 0-.5-.57.74.74 0 0 0-.28.06">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-12-2" x="33.32" y="26.04" width="4.58" height="3.17" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-24">
                                <path id="path-23" class="cls-1"
                                    d="M34.06 26.27c-.46 0-.57.26-.25.59l2.38 2.37a.6.6 0 0 0 1-.15l1-2a.52.52 0 0 0-.52-.77h-3.61z">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-13" x="6.31" y="26.04" width="6.49" height="6.48" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-26">
                                <path id="path-25" class="cls-1"
                                    d="M7.46 26.27a.84.84 0 0 0-.84.83v4.8a.85.85 0 0 0 .84.84h4.81a.85.85 0 0 0 .84-.84v-4.8a.84.84 0 0 0-.84-.83H7.46z">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-14-2" x="1.05" y="26.04" width="4.58" height="3.17" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-28">
                                <path id="path-27" class="cls-1"
                                    d="M1.94 26.27a.52.52 0 0 0-.52.77l1 2a.59.59 0 0 0 1 .15l2.37-2.37c.33-.33.22-.59-.25-.59h-3.6z">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-15" x="26.15" y="6.25" width="6.49" height="6.48" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-30">
                                <path id="path-29" class="cls-1"
                                    d="M27.3 6.47a.85.85 0 0 0-.84.84v4.8a.85.85 0 0 0 .84.84h4.81a.85.85 0 0 0 .84-.84v-4.8a.85.85 0 0 0-.84-.84H27.3z">
                                </path>
                            </g>
                        </g>
                    </mask>
                    <mask id="mask-16-2" x="0" y="13.22" width="12.35" height="12.32" maskUnits="userSpaceOnUse">
                        <g transform="translate(-.31 -.22)">
                            <g id="mask-32">
                                <path id="path-31" class="cls-1"
                                    d="M5.89 13.69L.55 19a.84.84 0 0 0 0 1.19l5.34 5.32a.84.84 0 0 0 1.19 0l5.33-5.32a.84.84 0 0 0 0-1.19l-5.33-5.31a.85.85 0 0 0-.6-.24.84.84 0 0 0-.59.24">
                                </path>
                            </g>
                        </g>
                    </mask>
                </defs>
                <g id="Symbols">
                    <g class="svgName">
                        <path class="name r" id="Fill-57"
                            d="M58 12.6c-1.58 0-2.29.43-3.74 2.16V14c0-.91-.12-1-1-1h-.74c-.91 0-1 .12-1 1v14.28c0 .9.12 1 1 1h.74c.91 0 1-.12 1-1V20.7a8.24 8.24 0 0 1 .63-3.93A3.06 3.06 0 0 1 58 15.31a3.8 3.8 0 0 1 .8.22.42.42 0 0 0 .31 0 .54.54 0 0 0 .24-.21 4.5 4.5 0 0 0 .39-.67l.23-.45a2.24 2.24 0 0 0 .28-.67c0-.51-1-.94-2.24-.94"
                            transform="translate(-.31 -.22)"></path>
                        <path class="name u" id="Fill-59"
                            d="M66.09 22.5a6.61 6.61 0 0 0 .51 3.07 3.87 3.87 0 0 0 6.34 0 6.61 6.61 0 0 0 .51-3.07V14c0-.91.12-1 1-1h.75c.9 0 1 .12 1 1v8.8c0 2.39-.39 3.69-1.49 4.91a7.1 7.1 0 0 1-10 0c-1.1-1.22-1.49-2.52-1.49-4.91V14c0-.91.11-1 1-1h.75c.9 0 1 .12 1 1z"
                            transform="translate(-.31 -.22)"></path>
                        <path class="name b" id="Fill-61"
                            d="M83.42 21.13c0 3.61 2.24 6.09 5.47 6.09s5.35-2.6 5.35-6.17a5.54 5.54 0 0 0-5.39-5.86c-3.19 0-5.43 2.44-5.43 5.94zm.2-5.82a7 7 0 0 1 5.7-2.67c4.49 0 7.79 3.58 7.79 8.49s-3.34 8.64-7.87 8.64a6.89 6.89 0 0 1-5.62-2.71v1.22c0 .9-.12 1-1 1h-.74c-.91 0-1-.12-1-1V1.68c0-.9.12-1 1-1h.74c.91 0 1 .12 1 1z"
                            transform="translate(-.31 -.22)"></path>
                        <path class="name r" id="Fill-55"
                            d="M107.72 12.6c-1.57 0-2.28.43-3.74 2.16V14c0-.91-.12-1-1-1h-.75c-.9 0-1 .12-1 1v14.28c0 .9.12 1 1 1h.77c.9 0 1-.12 1-1V20.7a8.37 8.37 0 0 1 .63-3.93 3.07 3.07 0 0 1 3.1-1.46 3.8 3.8 0 0 1 .8.22.42.42 0 0 0 .31 0 .63.63 0 0 0 .25-.21 4.44 4.44 0 0 0 .38-.67l.24-.45a2.42 2.42 0 0 0 .27-.67c0-.51-1-.94-2.24-.94"
                            transform="translate(-.31 -.22)"></path>
                        <path class="name i" id="Fill-63"
                            d="M116.4 28.28c0 .9-.12 1-1 1h-.75c-.9 0-1-.12-1-1V14c0-.91.12-1 1-1h.75c.9 0 1 .12 1 1zm.6-21.45a2 2 0 1 1-2-2 2 2 0 0 1 2 2z"
                            transform="translate(-.31 -.22)"></path>
                        <path class="name k" id="Fill-65"
                            d="M129.84 13.47c.47-.48.47-.48 1.14-.48h1.22c.71 0 1 .2 1 .63 0 .16-.15.4-.47.71L127.08 20l7.13 8c.27.36.43.59.43.75 0 .39-.32.59-1 .59h-1.24c-.71 0-.71 0-1.14-.51l-6.14-6.92-.71.71v5.7c0 .9-.12 1-1 1h-.75c-.91 0-1-.12-1-1V1.68c0-.9.12-1 1-1h.75c.9 0 1 .12 1 1V19z"
                            transform="translate(-.31 -.22)"></path>
                    </g>
                    <g class="svgLogo">
                        <g mask="url(#mask)">
                            <path id="Fill-68" class="cls-2"
                                d="M19.51.22a.83.83 0 0 0-.32.2l-5.34 5.32a.84.84 0 0 0 0 1.19l5.34 5.32a.84.84 0 0 0 1.19 0l5.33-5.32a.84.84 0 0 0 0-1.19L20.38.42a.83.83 0 0 0-.32-.2h-.55z"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-2-2)">
                            <path id="Fill-71" class="cls-2"
                                d="M19.19 27l-5.34 5.32a.83.83 0 0 0 0 1.18l5.34 5.33a.85.85 0 0 0 .25.17h.69a.85.85 0 0 0 .25-.17l5.33-5.33a.83.83 0 0 0 0-1.18L20.38 27a.82.82 0 0 0-.6-.25.81.81 0 0 0-.59.25"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-3)">
                            <path id="Fill-74" class="cls-2"
                                d="M32.49 13.69L27.15 19a.86.86 0 0 0 0 1.19l5.34 5.32a.84.84 0 0 0 1.19 0L39 20.2a.84.84 0 0 0 0-1.2l-5.33-5.32a.84.84 0 0 0-.59-.24.85.85 0 0 0-.6.24"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-4-2)">
                            <path id="Fill-77" class="cls-2"
                                d="M12.51 33.61L10.14 36a.59.59 0 0 0 .15 1l2 1a.52.52 0 0 0 .78-.52v-3.63c0-.28-.1-.43-.25-.43a.53.53 0 0 0-.35.19"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-5)">
                            <path id="Fill-80" class="cls-2"
                                d="M26.46 33.85v3.56a.52.52 0 0 0 .77.52l2.05-1a.59.59 0 0 0 .14-1l-2.37-2.36a.52.52 0 0 0-.34-.19c-.15 0-.25.15-.25.43"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-6-2)">
                            <path id="Fill-83" class="cls-2"
                                d="M27.3 26.27a.84.84 0 0 0-.84.83v4.8a.85.85 0 0 0 .84.84h4.81a.85.85 0 0 0 .89-.84v-4.8a.84.84 0 0 0-.84-.83H27.3z"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-7)">
                            <path id="Fill-86" class="cls-2"
                                d="M36.19 10l-2.38 2.37c-.32.32-.21.59.25.59h3.57a.53.53 0 0 0 .52-.78l-1-2a.62.62 0 0 0-.54-.36.65.65 0 0 0-.45.21"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-8-2)">
                            <path id="Fill-89" class="cls-2"
                                d="M26.46 1.8v3.56c0 .46.26.57.59.25l2.37-2.37a.59.59 0 0 0-.14-1l-2.05-1a.55.55 0 0 0-.23-.01.5.5 0 0 0-.5.57"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-9)">
                            <path id="Fill-92" class="cls-2"
                                d="M2.39 10.14l-1 2a.52.52 0 0 0 .52.78H5.5c.47 0 .58-.27.25-.59L3.38 10a.65.65 0 0 0-.46-.21.59.59 0 0 0-.53.36"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-10-2)">
                            <path id="Fill-95" class="cls-2"
                                d="M7.46 6.47a.85.85 0 0 0-.84.84v4.8a.85.85 0 0 0 .84.84h4.81a.85.85 0 0 0 .84-.84v-4.8a.85.85 0 0 0-.84-.84H7.46z"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-11)">
                            <path id="Fill-98" class="cls-2"
                                d="M12.33 1.29l-2 1a.59.59 0 0 0-.15 1l2.37 2.37c.33.32.6.21.6-.25V1.8a.51.51 0 0 0-.5-.57.74.74 0 0 0-.28.06"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-12-2)">
                            <path id="Fill-101" class="cls-2"
                                d="M34.06 26.27c-.46 0-.57.26-.25.59l2.38 2.37a.6.6 0 0 0 1-.15l1-2a.52.52 0 0 0-.52-.77h-3.61z"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-13)">
                            <path id="Fill-104" class="cls-2"
                                d="M7.46 26.27a.84.84 0 0 0-.84.83v4.8a.85.85 0 0 0 .84.84h4.81a.85.85 0 0 0 .84-.84v-4.8a.84.84 0 0 0-.84-.83H7.46z"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-14-2)">
                            <path id="Fill-107" class="cls-2"
                                d="M1.94 26.27a.52.52 0 0 0-.52.77l1 2a.59.59 0 0 0 1 .15l2.37-2.37c.33-.33.22-.59-.25-.59h-3.6z"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-15)">
                            <path id="Fill-110" class="cls-2"
                                d="M27.3 6.47a.85.85 0 0 0-.84.84v4.8a.85.85 0 0 0 .84.84h4.81a.85.85 0 0 0 .84-.84v-4.8a.85.85 0 0 0-.84-.84H27.3z"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                        <g mask="url(#mask-16-2)">
                            <path id="Fill-113" class="cls-2"
                                d="M5.89 13.69L.55 19a.84.84 0 0 0 0 1.19l5.34 5.32a.84.84 0 0 0 1.19 0l5.33-5.32a.84.84 0 0 0 0-1.19l-5.33-5.31a.85.85 0 0 0-.6-.24.84.84 0 0 0-.59.24"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                    </g>
                    <path id="Fill-116" class="cls-2"
                        d="M134.82 13.78h.09c.1 0 .18 0 .18-.12s-.05-.12-.17-.12h-.1zm0 .45h-.18v-.8h.3a.46.46 0 0 1 .28.06.21.21 0 0 1 .08.17.21.21 0 0 1-.17.19c.08 0 .12.08.15.19a.51.51 0 0 0 .06.2h-.2a.47.47 0 0 1-.06-.2.15.15 0 0 0-.17-.12h-.09zm-.49-.42a.62.62 0 0 0 .62.64.64.64 0 0 0 0-1.27.62.62 0 0 0-.62.63zm1.43 0a.8.8 0 0 1-.81.81.81.81 0 0 1-.82-.81.8.8 0 0 1 .82-.79.79.79 0 0 1 .81.79z"
                        transform="translate(-.31 -.22)"></path>
                </g>
            </svg>
        </div>
        <div class="nav-bar-text">Microsoft 365 Sizing</div>
    </div>
    <div class="margin"></div>

    <!-- Exchange Mailbox -->
    <div class="card-container">
        <div class="card">
            <div class="card-header">
                <div>
                    <svg xmlns="http://www.w3.org/2000/svg" height="62" width="auto" viewBox="-8.24997 -12 71.49974 72">
                        <path fill="#28a8ea"
                            d="M51.5095 0h-12.207a3.4884 3.4884 0 00-2.4677 1.0225L8.0222 29.835a3.4884 3.4884 0 00-1.0224 2.4677v12.207A3.49 3.49 0 0010.49 48h12.207a3.4884 3.4884 0 002.4678-1.0225l28.813-28.8125a3.49 3.49 0 001.022-2.4677V3.4903A3.49 3.49 0 0051.5095 0z" />
                        <path fill="#0078d4"
                            d="M51.5098 48H39.3025a3.49 3.49 0 01-2.4678-1.0222l-5.835-5.835V30.24a6.24 6.24 0 016.24-6.24h10.903l5.8349 5.835a3.49 3.49 0 011.0222 2.4678V44.51a3.49 3.49 0 01-3.49 3.49z" />
                        <path fill="#50d9ff"
                            d="M10.4898 0H22.697a3.49 3.49 0 012.4678 1.0222l5.835 5.835V17.76a6.24 6.24 0 01-6.24 6.24H13.8569l-5.835-5.835a3.49 3.49 0 01-1.0221-2.4677V3.49a3.49 3.49 0 013.49-3.49z" />
                        <path opacity=".2"
                            d="M28.9998 12.33v26.34a1.7344 1.7344 0 01-.04.3998A2.3138 2.3138 0 0126.6697 41h-19.67V10h19.67a2.326 2.326 0 012.33 2.33z" />
                        <path opacity=".1"
                            d="M29.9998 12.33v24.34A3.3617 3.3617 0 0126.6697 40h-19.67V9h19.67a3.3418 3.3418 0 013.33 3.33z" />
                        <path opacity=".2"
                            d="M28.9998 12.33v24.34A2.326 2.326 0 0126.6697 39h-19.67V10h19.67a2.326 2.326 0 012.33 2.33z" />
                        <path opacity=".1"
                            d="M27.9998 12.33v24.34A2.326 2.326 0 0125.6697 39h-18.67V10h18.67a2.326 2.326 0 012.33 2.33z" />
                        <rect fill="#0078d4" rx="2.3333" height="28" width="28" y="10" />
                        <path fill="#fff"
                            d="M18.5851 18.8812H12.038v3.8286h6.1454v2.4537H12.038v3.9766h6.8961v2.4434H9.066v-15.167h9.5191z" />
                    </svg>
                </div>
                <div class="card-header-text">
                    Exchange Online $ADGroup
                </div>
            </div>

            <table class="styled-table">
                <thead>
                    <tr>
                        <th>Total Mailboxes</th>
                        <th>Total Size (GB)</th>
                        <th>Total # of Items</th>
                        <th>Per User Size (GB)</th>
                        <th># of User Mailboxes</th>
                        <th># of Shared Mailboxes</th>
                        <th># of User Mailboxes w/Archive</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>$($ExchangeDetails.'Total Mailboxes')</td>
                        <td>$([math]::round($ExchangeDetails.'Total Storage Used' / 1GB, 2))</td>
                        <td>$($ExchangeDetails.'Total Items')</td>
                        <td>$([math]::round($ExchangeDetails.'Total Storage Used' / 1GB / $ExchangeDetails.'Total Mailboxes', 2) )</td>
                        <td>$($ExchangeDetails.'User Mailboxes')</td>
                        <td>$($ExchangeDetails.'Shared Mailboxes')</td>
                        <td>$($ExchangeDetails.'Archive Mailboxes')</td>
                    </tr>
                </tbody>
            </table>
        </div>
    </div>

    <!-- OneDrive -->
    <div class="card-container">
        <div class="card">
            <div class="card-header">
                <div>
                    <svg xmlns="http://www.w3.org/2000/svg" height="62" width="auto"
                        viewBox="-154.5063 -164.9805 1339.0546 989.883">
                        <path
                            d="M622.292 445.338l212.613-203.327C790.741 69.804 615.338-33.996 443.13 10.168a321.9 321.9 0 00-188.92 134.837c3.29-.083 368.082 300.333 368.082 300.333z"
                            fill="#0364B8" />
                        <path
                            d="M392.776 183.283l-.01.035A256.233 256.233 0 00257.5 144.921c-1.104 0-2.189.07-3.29.083C112.063 146.765-1.74 263.424.02 405.567a257.389 257.389 0 0046.244 144.04l318.528-39.894 244.21-196.915z"
                            fill="#0078D4" />
                        <path
                            d="M834.905 242.012c-4.674-.312-9.37-.528-14.123-.528a208.464 208.464 0 00-82.93 17.117l-.006-.022-128.844 54.22 142.041 175.456 253.934 61.728c54.8-101.732 16.752-228.625-84.98-283.424a209.23 209.23 0 00-85.09-24.546z"
                            fill="#1490DF" />
                        <path
                            d="M46.264 549.607C94.36 618.757 173.27 659.967 257.5 659.922h563.281c76.946.022 147.691-42.202 184.195-109.937L609.001 312.798z"
                            fill="#28A8EA" />
                    </svg>
                </div>
                <div class="card-header-text">
                    OneDrive $ADGroup
                </div>
            </div>

            <table class="styled-table">
                <thead>
                    <tr>
                        <th>Total Accounts</th>
                        <th>Total Size (GB)</th>
                        <th>Total # of Files</th>
                        <th>Per Account Size (GB)</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>$($OneDriveDetails.'Accounts')</td>
                        <td>$([math]::round($OneDriveDetails.'Storage Used' / 1GB, 2))</td>
                        <td>$($OneDriveDetails.'Total Files')</td>
                        <td>$([math]::round($OneDriveDetails.'Storage Used' / 1GB / $OneDriveDetails.'Accounts', 2) )</td>
                    </tr>
                </tbody>
            </table>
        </div>
    </div>

    <!-- SharePoint -->
    <div class="card-container">
        <div class="card">
            <div class="card-header">
                <div>
                    <svg xmlns="http://www.w3.org/2000/svg" height="82" width="auto"
                        viewBox="-298.8501 -486.5 2590.0342 2919">
                        <circle r="556" cy="556" cx="1019.333" fill="#036C70" />
                        <circle r="509.667" cy="1065.667" cx="1482.667" fill="#1A9BA1" />
                        <circle r="393.833" cy="1552.167" cx="1088.833" fill="#37C6D0" />
                        <path
                            d="M1112 501.79v988.753c-.23 34.357-21.05 65.222-52.82 78.303a82.12 82.12 0 01-31.97 6.487H695.463c-.463-7.877-.463-15.29-.463-23.167-.154-7.734.155-15.47.927-23.167 8.48-148.106 99.721-278.782 235.837-337.77v-86.18c-302.932-48.005-509.592-332.495-461.587-635.427.333-2.098.677-4.195 1.034-6.289a391.8 391.8 0 019.73-46.333h546.27c46.753.178 84.611 38.036 84.789 84.79z"
                            opacity=".1" />
                        <path
                            d="M980.877 463.333H471.21c-51.486 302.386 151.908 589.256 454.293 640.742a555.466 555.466 0 0027.573 3.986c-143.633 68.11-248.3 261.552-257.196 420.938a193.737 193.737 0 00-.927 23.167c0 7.877 0 15.29.463 23.167a309.212 309.212 0 006.023 46.333h279.39c34.357-.23 65.222-21.05 78.303-52.82a82.098 82.098 0 006.487-31.97V548.123c-.176-46.736-38.006-84.586-84.742-84.79z"
                            opacity=".2" />
                        <path
                            d="M980.877 463.333H471.21c-51.475 302.414 151.95 589.297 454.364 640.773a556.017 556.017 0 0018.607 2.844c-139 73.021-239.543 266-248.254 422.05h284.95c46.681-.353 84.437-38.109 84.79-84.79V548.123c-.178-46.754-38.036-84.612-84.79-84.79z"
                            opacity=".2" />
                        <path
                            d="M934.543 463.333H471.21c-48.606 285.482 130.279 560.404 410.977 631.616A765.521 765.521 0 00695.927 1529h238.617c46.754-.178 84.612-38.036 84.79-84.79V548.123c-.026-46.817-37.973-84.764-84.791-84.79z"
                            opacity=".2" />
                        <linearGradient gradientTransform="matrix(1 0 0 -1 0 1948)" y2="398.972" x2="842.255"
                            y1="1551.028" x1="177.079" gradientUnits="userSpaceOnUse" id="a">
                            <stop offset="0" stop-color="#058f92" />
                            <stop offset=".5" stop-color="#038489" />
                            <stop offset="1" stop-color="#026d71" />
                        </linearGradient>
                        <path
                            d="M84.929 463.333h849.475c46.905 0 84.929 38.024 84.929 84.929v849.475c0 46.905-38.024 84.929-84.929 84.929H84.929c-46.905 0-84.929-38.024-84.929-84.929V548.262c0-46.905 38.024-84.929 84.929-84.929z"
                            fill="url(#a)" />
                        <path
                            d="M379.331 962.621a156.785 156.785 0 01-48.604-51.384 139.837 139.837 0 01-16.912-70.288 135.25 135.25 0 0131.46-91.045 185.847 185.847 0 0183.678-54.581 353.459 353.459 0 01114.304-17.699 435.148 435.148 0 01150.583 21.082v106.567a235.031 235.031 0 00-68.11-27.8 331.709 331.709 0 00-79.647-9.545 172.314 172.314 0 00-81.871 17.329 53.7 53.7 0 00-32.433 49.206 49.853 49.853 0 0013.9 34.843 124.638 124.638 0 0037.067 26.503c15.444 7.691 38.611 17.916 69.5 30.673a70.322 70.322 0 019.915 3.985 571.842 571.842 0 0187.663 43.229 156.935 156.935 0 0151.801 52.171 151.223 151.223 0 0118.533 78.767 146.506 146.506 0 01-29.468 94.798 164.803 164.803 0 01-78.767 53.005 357.22 357.22 0 01-112.312 16.309 594.113 594.113 0 01-101.933-8.34 349.057 349.057 0 01-82.612-24.279v-112.358a266.237 266.237 0 0083.4 39.847 326.268 326.268 0 0092.018 14.734 158.463 158.463 0 0083.4-17.699 55.971 55.971 0 0028.449-49.994 53.284 53.284 0 00-15.753-38.271 158.715 158.715 0 00-43.414-30.256c-18.533-9.267-45.824-21.483-81.871-36.65a465.328 465.328 0 01-81.964-42.859z"
                            fill="#FFF" />
                    </svg>
                </div>
                <div class="card-header-text">
                    SharePoint
                </div>
            </div>

            <table class="styled-table">
                <thead>
                    <tr>
                        <th>Total Sites</th>
                        <th>Total Size (GB)</th>
                        <th>Total # of Files</th>
                        <th>Per Site Size (GB)</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>$($SharePointDetails.'Sites')</td>
                        <td>$([math]::round($SharePointDetails.'Sites Storage Used' / 1GB, 2))</td>
                        <td>$($SharePointDetails.'Account Files')</td>
                        <td>$([math]::round($SharePointDetails.'Sites Storage Used' / 1GB / $SharePointDetails.'Sites', 2) )</td>
                    </tr>
                </tbody>
            </table>
        </div>
    </div>

    <!-- Total Data Needed -->
    <div class="card-container">
        <div class="card">
            <div class="card-header ">
                <div class="M365">
                <svg xmlns="http://www.w3.org/2000/svg" height="72" width="72" viewBox="-8 -35000 278050 403334" shape-rendering="geometricPrecision" text-rendering="geometricPrecision" image-rendering="optimizeQuality" fill-rule="evenodd" clip-rule="evenodd">
                <path fill="#ea3e23" d="M278050 305556l-29-16V28627L178807 0 448 66971l-448 87 22 200227 60865-23821V80555l117920-28193-17 239519L122 267285l178668 65976v73l99231-27462v-316z"/></svg>
                </div>
                <div class="card-header-text">
                    Discovery Summary
                </div>
                </div>

                <table class="styled-table">
                    <thead>
                        <tr>
                            <th>Required # of Licenses</th>
                            <th>Total Size (GB)</th>
                            <th>Total # of Items & Files</th>
                            <th>Per User/Account Size (GB)</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td>$UserLicensesRequired</td>
                            <td>$totalStorageGB</td>
                            <td>$totalItems</td>
                            <td>$([math]::round($totalStorageGB / $totalItems, 2))</td>
                        </tr>
                    </tbody>
                    <thead>
                        <tr>
                            <th>Year 1 Licenses @ 10% Growth</th>
                            <th>Year 1 Licenses @ 20% Growth</th>
                            <th>Year 1 Licenses @ $AnnualGrowth% Growth</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td>$UserLicenseRequired10</td>
                            <td>$UserLicenseRequired20</td>
                            <td>$UserLicenseRequiredCustom</td>
                        </tr>
                    </tbody>
                    <thead>
                        <tr>
                            <th>Year 1 @ 10% Growth (GB)</th>
                            <th>Year 1 @ 20% Growth (GB)</th>
                            <th>Year 1 @ $AnnualGrowth% Growth (GB)</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td>$totalStorage10GB</td>
                            <td>$totalStorage20GB</td>
                            <td>$totalStorageCustomGB</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>

    <!-- Subscription Output -->
    <div class="card-container">
        <div class="card">
            <div class="card-header ">
                <div class="rubrik-snowflake">
                    <svg xmlns="http://www.w3.org/2000/svg" height="52" width="auto" viewBox="0 0 50 38.77">
                        <defs>
                            <style>
                                .cls-1 {
                                    fill: #fff
                                }
                                .cls-1,
                                .cls-2 {
                                    fill-rule: evenodd
                                }
                            </style>
                            <mask id="mask" x="13.3" y="0" width="12.35" height="12.27" maskUnits="userSpaceOnUse">
                                <g transform="translate(-.31 -.22)">
                                    <g id="mask-2">
                                        <path id="path-1" class="cls-1"
                                            d="M19.51.22a.83.83 0 0 0-.32.2l-5.34 5.32a.84.84 0 0 0 0 1.19l5.34 5.32a.84.84 0 0 0 1.19 0l5.33-5.32a.84.84 0 0 0 0-1.19L20.38.42a.83.83 0 0 0-.32-.2h-.55z">
                                        </path>
                                    </g>
                                </g>
                            </mask>
                            <mask id="mask-2-2" x="13.3" y="26.53" width="12.35" height="12.25"
                                maskUnits="userSpaceOnUse">
                                <g transform="translate(-.31 -.22)">
                                    <g id="mask-4">
                                        <path id="path-3" class="cls-1"
                                            d="M19.19 27l-5.34 5.32a.83.83 0 0 0 0 1.18l5.34 5.33a.85.85 0 0 0 .25.17h.69a.85.85 0 0 0 .25-.17l5.33-5.33a.83.83 0 0 0 0-1.18L20.38 27a.82.82 0 0 0-.6-.25.81.81 0 0 0-.59.25">
                                        </path>
                                    </g>
                                </g>
                            </mask>
                            <mask id="mask-3" x="26.6" y="13.22" width="12.35" height="12.32"
                                maskUnits="userSpaceOnUse">
                                <g transform="translate(-.31 -.22)">
                                    <g id="mask-6">
                                        <path id="path-5" class="cls-1"
                                            d="M32.49 13.69L27.15 19a.86.86 0 0 0 0 1.19l5.34 5.32a.84.84 0 0 0 1.19 0L39 20.2a.84.84 0 0 0 0-1.2l-5.33-5.32a.84.84 0 0 0-.59-.24.85.85 0 0 0-.6.24">
                                        </path>
                                    </g>
                                </g>
                            </mask>
                            <mask id="mask-4-2" x="9.63" y="33.2" width="3.17" height="4.57" maskUnits="userSpaceOnUse">
                                <g transform="translate(-.31 -.22)">
                                    <g id="mask-8">
                                        <path id="path-7" class="cls-1"
                                            d="M12.51 33.61L10.14 36a.59.59 0 0 0 .15 1l2 1a.52.52 0 0 0 .78-.52v-3.63c0-.28-.1-.43-.25-.43a.53.53 0 0 0-.35.19">
                                        </path>
                                    </g>
                                </g>
                            </mask>
                            <mask id="mask-5" x="26.15" y="33.2" width="3.17" height="4.57" maskUnits="userSpaceOnUse">
                                <g transform="translate(-.31 -.22)">
                                    <g id="mask-10">
                                        <path id="path-9" class="cls-1"
                                            d="M26.46 33.85v3.56a.52.52 0 0 0 .77.52l2.05-1a.59.59 0 0 0 .14-1l-2.37-2.36a.52.52 0 0 0-.34-.19c-.15 0-.25.15-.25.43">
                                        </path>
                                    </g>
                                </g>
                            </mask>
                            <mask id="mask-6-2" x="26.15" y="26.04" width="6.49" height="6.48"
                                maskUnits="userSpaceOnUse">
                                <g transform="translate(-.31 -.22)">
                                    <g id="mask-12">
                                        <path id="path-11" class="cls-1"
                                            d="M27.3 26.27a.84.84 0 0 0-.84.83v4.8a.85.85 0 0 0 .84.84h4.81a.85.85 0 0 0 .89-.84v-4.8a.84.84 0 0 0-.84-.83H27.3z">
                                        </path>
                                    </g>
                                </g>
                            </mask>
                            <mask id="mask-7" x="33.32" y="9.56" width="4.58" height="3.17" maskUnits="userSpaceOnUse">
                                <g transform="translate(-.31 -.22)">
                                    <g id="mask-14">
                                        <path id="path-13" class="cls-1"
                                            d="M36.19 10l-2.38 2.37c-.32.32-.21.59.25.59h3.57a.53.53 0 0 0 .52-.78l-1-2a.62.62 0 0 0-.54-.36.65.65 0 0 0-.45.21">
                                        </path>
                                    </g>
                                </g>
                            </mask>
                            <mask id="mask-8-2" x="26.15" y="1" width="3.17" height="4.57" maskUnits="userSpaceOnUse">
                                <g transform="translate(-.31 -.22)">
                                    <g id="mask-16">
                                        <path id="path-15" class="cls-1"
                                            d="M26.46 1.8v3.56c0 .46.26.57.59.25l2.37-2.37a.59.59 0 0 0-.14-1l-2.05-1a.55.55 0 0 0-.23-.01.5.5 0 0 0-.5.57">
                                        </path>
                                    </g>
                                </g>
                            </mask>
                            <mask id="mask-9" x="1.05" y="9.56" width="4.58" height="3.17" maskUnits="userSpaceOnUse">
                                <g transform="translate(-.31 -.22)">
                                    <g id="mask-18">
                                        <path id="path-17" class="cls-1"
                                            d="M2.39 10.14l-1 2a.52.52 0 0 0 .52.78H5.5c.47 0 .58-.27.25-.59L3.38 10a.65.65 0 0 0-.46-.21.59.59 0 0 0-.53.36">
                                        </path>
                                    </g>
                                </g>
                            </mask>
                            <mask id="mask-10-2" x="6.31" y="6.25" width="6.49" height="6.48"
                                maskUnits="userSpaceOnUse">
                                <g transform="translate(-.31 -.22)">
                                    <g id="mask-20">
                                        <path id="path-19" class="cls-1"
                                            d="M7.46 6.47a.85.85 0 0 0-.84.84v4.8a.85.85 0 0 0 .84.84h4.81a.85.85 0 0 0 .84-.84v-4.8a.85.85 0 0 0-.84-.84H7.46z">
                                        </path>
                                    </g>
                                </g>
                            </mask>
                            <mask id="mask-11" x="9.63" y="1" width="3.17" height="4.57" maskUnits="userSpaceOnUse">
                                <g transform="translate(-.31 -.22)">
                                    <g id="mask-22">
                                        <path id="path-21" class="cls-1"
                                            d="M12.33 1.29l-2 1a.59.59 0 0 0-.15 1l2.37 2.37c.33.32.6.21.6-.25V1.8a.51.51 0 0 0-.5-.57.74.74 0 0 0-.28.06">
                                        </path>
                                    </g>
                                </g>
                            </mask>
                            <mask id="mask-12-2" x="33.32" y="26.04" width="4.58" height="3.17"
                                maskUnits="userSpaceOnUse">
                                <g transform="translate(-.31 -.22)">
                                    <g id="mask-24">
                                        <path id="path-23" class="cls-1"
                                            d="M34.06 26.27c-.46 0-.57.26-.25.59l2.38 2.37a.6.6 0 0 0 1-.15l1-2a.52.52 0 0 0-.52-.77h-3.61z">
                                        </path>
                                    </g>
                                </g>
                            </mask>
                            <mask id="mask-13" x="6.31" y="26.04" width="6.49" height="6.48" maskUnits="userSpaceOnUse">
                                <g transform="translate(-.31 -.22)">
                                    <g id="mask-26">
                                        <path id="path-25" class="cls-1"
                                            d="M7.46 26.27a.84.84 0 0 0-.84.83v4.8a.85.85 0 0 0 .84.84h4.81a.85.85 0 0 0 .84-.84v-4.8a.84.84 0 0 0-.84-.83H7.46z">
                                        </path>
                                    </g>
                                </g>
                            </mask>
                            <mask id="mask-14-2" x="1.05" y="26.04" width="4.58" height="3.17"
                                maskUnits="userSpaceOnUse">
                                <g transform="translate(-.31 -.22)">
                                    <g id="mask-28">
                                        <path id="path-27" class="cls-1"
                                            d="M1.94 26.27a.52.52 0 0 0-.52.77l1 2a.59.59 0 0 0 1 .15l2.37-2.37c.33-.33.22-.59-.25-.59h-3.6z">
                                        </path>
                                    </g>
                                </g>
                            </mask>
                            <mask id="mask-15" x="26.15" y="6.25" width="6.49" height="6.48" maskUnits="userSpaceOnUse">
                                <g transform="translate(-.31 -.22)">
                                    <g id="mask-30">
                                        <path id="path-29" class="cls-1"
                                            d="M27.3 6.47a.85.85 0 0 0-.84.84v4.8a.85.85 0 0 0 .84.84h4.81a.85.85 0 0 0 .84-.84v-4.8a.85.85 0 0 0-.84-.84H27.3z">
                                        </path>
                                    </g>
                                </g>
                            </mask>
                            <mask id="mask-16-2" x="0" y="13.22" width="12.35" height="12.32"
                                maskUnits="userSpaceOnUse">
                                <g transform="translate(-.31 -.22)">
                                    <g id="mask-32">
                                        <path id="path-31" class="cls-1"
                                            d="M5.89 13.69L.55 19a.84.84 0 0 0 0 1.19l5.34 5.32a.84.84 0 0 0 1.19 0l5.33-5.32a.84.84 0 0 0 0-1.19l-5.33-5.31a.85.85 0 0 0-.6-.24.84.84 0 0 0-.59.24">
                                        </path>
                                    </g>
                                </g>
                            </mask>
                        </defs>
                        <g id="Symbols">
                            <g class="svgName">
                                <path class="name r" id="Fill-57"
                                    d="M58 12.6c-1.58 0-2.29.43-3.74 2.16V14c0-.91-.12-1-1-1h-.74c-.91 0-1 .12-1 1v14.28c0 .9.12 1 1 1h.74c.91 0 1-.12 1-1V20.7a8.24 8.24 0 0 1 .63-3.93A3.06 3.06 0 0 1 58 15.31a3.8 3.8 0 0 1 .8.22.42.42 0 0 0 .31 0 .54.54 0 0 0 .24-.21 4.5 4.5 0 0 0 .39-.67l.23-.45a2.24 2.24 0 0 0 .28-.67c0-.51-1-.94-2.24-.94"
                                    transform="translate(-.31 -.22)"></path>
                                <path class="name u" id="Fill-59"
                                    d="M66.09 22.5a6.61 6.61 0 0 0 .51 3.07 3.87 3.87 0 0 0 6.34 0 6.61 6.61 0 0 0 .51-3.07V14c0-.91.12-1 1-1h.75c.9 0 1 .12 1 1v8.8c0 2.39-.39 3.69-1.49 4.91a7.1 7.1 0 0 1-10 0c-1.1-1.22-1.49-2.52-1.49-4.91V14c0-.91.11-1 1-1h.75c.9 0 1 .12 1 1z"
                                    transform="translate(-.31 -.22)"></path>
                                <path class="name b" id="Fill-61"
                                    d="M83.42 21.13c0 3.61 2.24 6.09 5.47 6.09s5.35-2.6 5.35-6.17a5.54 5.54 0 0 0-5.39-5.86c-3.19 0-5.43 2.44-5.43 5.94zm.2-5.82a7 7 0 0 1 5.7-2.67c4.49 0 7.79 3.58 7.79 8.49s-3.34 8.64-7.87 8.64a6.89 6.89 0 0 1-5.62-2.71v1.22c0 .9-.12 1-1 1h-.74c-.91 0-1-.12-1-1V1.68c0-.9.12-1 1-1h.74c.91 0 1 .12 1 1z"
                                    transform="translate(-.31 -.22)"></path>
                                <path class="name r" id="Fill-55"
                                    d="M107.72 12.6c-1.57 0-2.28.43-3.74 2.16V14c0-.91-.12-1-1-1h-.75c-.9 0-1 .12-1 1v14.28c0 .9.12 1 1 1h.77c.9 0 1-.12 1-1V20.7a8.37 8.37 0 0 1 .63-3.93 3.07 3.07 0 0 1 3.1-1.46 3.8 3.8 0 0 1 .8.22.42.42 0 0 0 .31 0 .63.63 0 0 0 .25-.21 4.44 4.44 0 0 0 .38-.67l.24-.45a2.42 2.42 0 0 0 .27-.67c0-.51-1-.94-2.24-.94"
                                    transform="translate(-.31 -.22)"></path>
                                <path class="name i" id="Fill-63"
                                    d="M116.4 28.28c0 .9-.12 1-1 1h-.75c-.9 0-1-.12-1-1V14c0-.91.12-1 1-1h.75c.9 0 1 .12 1 1zm.6-21.45a2 2 0 1 1-2-2 2 2 0 0 1 2 2z"
                                    transform="translate(-.31 -.22)"></path>
                                <path class="name k" id="Fill-65"
                                    d="M129.84 13.47c.47-.48.47-.48 1.14-.48h1.22c.71 0 1 .2 1 .63 0 .16-.15.4-.47.71L127.08 20l7.13 8c.27.36.43.59.43.75 0 .39-.32.59-1 .59h-1.24c-.71 0-.71 0-1.14-.51l-6.14-6.92-.71.71v5.7c0 .9-.12 1-1 1h-.75c-.91 0-1-.12-1-1V1.68c0-.9.12-1 1-1h.75c.9 0 1 .12 1 1V19z"
                                    transform="translate(-.31 -.22)"></path>
                            </g>
                            <g class="svgLogo">
                                <g mask="url(#mask)">
                                    <path id="Fill-68" class="cls-2"
                                        d="M19.51.22a.83.83 0 0 0-.32.2l-5.34 5.32a.84.84 0 0 0 0 1.19l5.34 5.32a.84.84 0 0 0 1.19 0l5.33-5.32a.84.84 0 0 0 0-1.19L20.38.42a.83.83 0 0 0-.32-.2h-.55z"
                                        transform="translate(-.31 -.22)"></path>
                                </g>
                                <g mask="url(#mask-2-2)">
                                    <path id="Fill-71" class="cls-2"
                                        d="M19.19 27l-5.34 5.32a.83.83 0 0 0 0 1.18l5.34 5.33a.85.85 0 0 0 .25.17h.69a.85.85 0 0 0 .25-.17l5.33-5.33a.83.83 0 0 0 0-1.18L20.38 27a.82.82 0 0 0-.6-.25.81.81 0 0 0-.59.25"
                                        transform="translate(-.31 -.22)"></path>
                                </g>
                                <g mask="url(#mask-3)">
                                    <path id="Fill-74" class="cls-2"
                                        d="M32.49 13.69L27.15 19a.86.86 0 0 0 0 1.19l5.34 5.32a.84.84 0 0 0 1.19 0L39 20.2a.84.84 0 0 0 0-1.2l-5.33-5.32a.84.84 0 0 0-.59-.24.85.85 0 0 0-.6.24"
                                        transform="translate(-.31 -.22)"></path>
                                </g>
                                <g mask="url(#mask-4-2)">
                                    <path id="Fill-77" class="cls-2"
                                        d="M12.51 33.61L10.14 36a.59.59 0 0 0 .15 1l2 1a.52.52 0 0 0 .78-.52v-3.63c0-.28-.1-.43-.25-.43a.53.53 0 0 0-.35.19"
                                        transform="translate(-.31 -.22)"></path>
                                </g>
                                <g mask="url(#mask-5)">
                                    <path id="Fill-80" class="cls-2"
                                        d="M26.46 33.85v3.56a.52.52 0 0 0 .77.52l2.05-1a.59.59 0 0 0 .14-1l-2.37-2.36a.52.52 0 0 0-.34-.19c-.15 0-.25.15-.25.43"
                                        transform="translate(-.31 -.22)"></path>
                                </g>
                                <g mask="url(#mask-6-2)">
                                    <path id="Fill-83" class="cls-2"
                                        d="M27.3 26.27a.84.84 0 0 0-.84.83v4.8a.85.85 0 0 0 .84.84h4.81a.85.85 0 0 0 .89-.84v-4.8a.84.84 0 0 0-.84-.83H27.3z"
                                        transform="translate(-.31 -.22)"></path>
                                </g>
                                <g mask="url(#mask-7)">
                                    <path id="Fill-86" class="cls-2"
                                        d="M36.19 10l-2.38 2.37c-.32.32-.21.59.25.59h3.57a.53.53 0 0 0 .52-.78l-1-2a.62.62 0 0 0-.54-.36.65.65 0 0 0-.45.21"
                                        transform="translate(-.31 -.22)"></path>
                                </g>
                                <g mask="url(#mask-8-2)">
                                    <path id="Fill-89" class="cls-2"
                                        d="M26.46 1.8v3.56c0 .46.26.57.59.25l2.37-2.37a.59.59 0 0 0-.14-1l-2.05-1a.55.55 0 0 0-.23-.01.5.5 0 0 0-.5.57"
                                        transform="translate(-.31 -.22)"></path>
                                </g>
                                <g mask="url(#mask-9)">
                                    <path id="Fill-92" class="cls-2"
                                        d="M2.39 10.14l-1 2a.52.52 0 0 0 .52.78H5.5c.47 0 .58-.27.25-.59L3.38 10a.65.65 0 0 0-.46-.21.59.59 0 0 0-.53.36"
                                        transform="translate(-.31 -.22)"></path>
                                </g>
                                <g mask="url(#mask-10-2)">
                                    <path id="Fill-95" class="cls-2"
                                        d="M7.46 6.47a.85.85 0 0 0-.84.84v4.8a.85.85 0 0 0 .84.84h4.81a.85.85 0 0 0 .84-.84v-4.8a.85.85 0 0 0-.84-.84H7.46z"
                                        transform="translate(-.31 -.22)"></path>
                                </g>
                                <g mask="url(#mask-11)">
                                    <path id="Fill-98" class="cls-2"
                                        d="M12.33 1.29l-2 1a.59.59 0 0 0-.15 1l2.37 2.37c.33.32.6.21.6-.25V1.8a.51.51 0 0 0-.5-.57.74.74 0 0 0-.28.06"
                                        transform="translate(-.31 -.22)"></path>
                                </g>
                                <g mask="url(#mask-12-2)">
                                    <path id="Fill-101" class="cls-2"
                                        d="M34.06 26.27c-.46 0-.57.26-.25.59l2.38 2.37a.6.6 0 0 0 1-.15l1-2a.52.52 0 0 0-.52-.77h-3.61z"
                                        transform="translate(-.31 -.22)"></path>
                                </g>
                                <g mask="url(#mask-13)">
                                    <path id="Fill-104" class="cls-2"
                                        d="M7.46 26.27a.84.84 0 0 0-.84.83v4.8a.85.85 0 0 0 .84.84h4.81a.85.85 0 0 0 .84-.84v-4.8a.84.84 0 0 0-.84-.83H7.46z"
                                        transform="translate(-.31 -.22)"></path>
                                </g>
                                <g mask="url(#mask-14-2)">
                                    <path id="Fill-107" class="cls-2"
                                        d="M1.94 26.27a.52.52 0 0 0-.52.77l1 2a.59.59 0 0 0 1 .15l2.37-2.37c.33-.33.22-.59-.25-.59h-3.6z"
                                        transform="translate(-.31 -.22)"></path>
                                </g>
                                <g mask="url(#mask-15)">
                                    <path id="Fill-110" class="cls-2"
                                        d="M27.3 6.47a.85.85 0 0 0-.84.84v4.8a.85.85 0 0 0 .84.84h4.81a.85.85 0 0 0 .84-.84v-4.8a.85.85 0 0 0-.84-.84H27.3z"
                                        transform="translate(-.31 -.22)"></path>
                                </g>
                                <g mask="url(#mask-16-2)">
                                    <path id="Fill-113" class="cls-2"
                                        d="M5.89 13.69L.55 19a.84.84 0 0 0 0 1.19l5.34 5.32a.84.84 0 0 0 1.19 0l5.33-5.32a.84.84 0 0 0 0-1.19l-5.33-5.31a.85.85 0 0 0-.6-.24.84.84 0 0 0-.59.24"
                                        transform="translate(-.31 -.22)"></path>
                                </g>
                            </g>
                            <path id="Fill-116" class="cls-2"
                                d="M134.82 13.78h.09c.1 0 .18 0 .18-.12s-.05-.12-.17-.12h-.1zm0 .45h-.18v-.8h.3a.46.46 0 0 1 .28.06.21.21 0 0 1 .08.17.21.21 0 0 1-.17.19c.08 0 .12.08.15.19a.51.51 0 0 0 .06.2h-.2a.47.47 0 0 1-.06-.2.15.15 0 0 0-.17-.12h-.09zm-.49-.42a.62.62 0 0 0 .62.64.64.64 0 0 0 0-1.27.62.62 0 0 0-.62.63zm1.43 0a.8.8 0 0 1-.81.81.81.81 0 0 1-.82-.81.8.8 0 0 1 .82-.79.79.79 0 0 1 .81.79z"
                                transform="translate(-.31 -.22)"></path>
                        </g>
                    </svg>
                </div>
                <div class="card-header-text">

                    License Option

                </div>
            </div>

            <table class="styled-table">
                <thead>
                    <tr>
                        <th>Year 1 @ 10% Growth</th>
                        <th>Starter (5GB)</th>
                        <th>Foundation (20GB)</th>
                        <th>Business (50GB)</th>
                        <th>Enterprise (Unlimited)</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>====></td>
                        <td>$($growth10.FiveGBUsers)</td>
                        <td>$($growth10.TwentyGBUsers)</td>
                        <td>$($growth10.FiftyGBUsers)</td>
                        <td>$($growth10.UnlimitedGBUsers)</td>
                    </tr>
                </tbody>
                <thead>
                    <tr>
                        <th>Year 1 @ 20% Growth</th>
                        <th>Starter (5GB)</th>
                        <th>Foundation (20GB)</th>
                        <th>Business (50GB)</th>
                        <th>Enterprise (Unlimited)</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>====></td>
                        <td>$($growth20.FiveGBUsers)</td>
                        <td>$($growth20.TwentyGBUsers)</td>
                        <td>$($growth20.FiftyGBUsers)</td>
                        <td>$($growth20.UnlimitedGBUsers)</td>
                    </tr>
                </tbody>
                <thead>
                    <tr>
                        <th>Year 1 @ $AnnualGrowth% Growth</th>
                        <th>Starter (5GB)</th>
                        <th>Foundation (20GB)</th>
                        <th>Business (50GB)</th>
                        <th>Enterprise (Unlimited)</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>====></td>
                        <td>$($growthCustom.FiveGBUsers)</td>
                        <td>$($growthCustom.TwentyGBUsers)</td>
                        <td>$($growthCustom.FiftyGBUsers)</td>
                        <td>$($growthCustom.UnlimitedGBUsers)</td>
                    </tr>
                </tbody>
            </table>
        </div>
    </div>
    <footer>
        <p style="color:#D3D3D3;text-align:right;padding-right: 10px;"<td>$CurrentDate $Version</td>
    </footer>
</body>
</html>
"@
#endregion

# Remove any previously created files
Remove-Item -Path $outFilename -ErrorAction SilentlyContinue
Write-Output $HTML_CODE | Format-Table -AutoSize | Out-File -FilePath $outFilename -Append

Write-Host "`n`nM365 Sizing information has been written to $((Get-ChildItem $outFilename).FullName)`n`n" -foregroundcolor green
