<#
.SYNOPSIS
    Get-RubrikM365SizingInfo.ps1 returns statistics on number of accounts, sites and how much storage they are using in a Micosoft 365 Tenant
.DESCRIPTION
    Get-RubrikM365SizingInfo.ps1 returns statistics on number of accounts, sites and how much storage they are using in a Micosoft 365 Tenant
    In this script, Rubrik uses Microsoft Graph APIs to return data from the customer's M365 Tenant. Data is collected via the Graph API
    and then downloaded to the customer's machine. The downloaded reports can be found in the customers $systemTempFolder folder. This data is left 
    behind and never sent to Rubrik or viewed by Rubrik. 

.EXAMPLE
    PS C:\> .\Get-RubrikM365SizingInfo.ps1
    Will connect to customer's M365 Tenant. A browser page will open up linking to the customer's M365 Tenant authorization page. The 
    customer will need to provide authorization. The script will gather data for 180 days. Once this is done output will be written to the current working directory as a file called 
    RubrikM365Sizing.txt
.INPUTS
    Inputs (if any)
.OUTPUTS
    Rubrik-M365-Sizing.html. 
.NOTES
    Author:         Chris Lumnah
    Created Date:   6/17/2021
#>

[CmdletBinding()]
param (
    [Parameter()]
    [bool]$EnableDebug = $false,
    # Parameter help description
    [Parameter()]
    [String]$AzureAdGroupName,
    $OutputObject
)

$Period = '180'
$Version = "v3.8"
Write-Output "[INFO] Starting the Rubrik Microsoft 365 sizing script ($Version)."

# Provide OS agnostic temp folder path for raw reports
$systemTempFolder = [System.IO.Path]::GetTempPath()
$ProgressPreference = 'SilentlyContinue'

function Get-MgReport {
    [CmdletBinding()]
    param (
        # MS Graph API report name
        [Parameter(Mandatory)]
        [String]$ReportName,

        # Report Period (Days)
        [Parameter(Mandatory)]
        [ValidateSet("7","30","90","180")]
        [String]$Period
    )
    
    process {
        try {
            Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/reports/$($ReportName)(period=`'D$($Period)`')" -OutputFilePath "$systemTempFolder\$ReportName.csv"

            "$systemTempFolder\$ReportName.csv"

        }
        catch {

            $errorMessage = $_.Exception | Out-String
            
            if($errorMessage.Contains('Response status code does not indicate success: Forbidden (Forbidden)')) {
                Disconnect-MgGraph
                throw "The user account used for authentication must have permissions covered by Reports Reader admin role."
            } 
            
            throw $_.Exception
        }
        
    }
}
function Measure-AverageGrowth {
    param (
        [Parameter(Mandatory)]
        [string]$ReportCSV, 
        [Parameter(Mandatory)]
        [string]$ReportName

    )
    if ($ReportName -eq 'getOneDriveUsageStorage'){
        $UsageReport = Import-Csv -Path $ReportCSV | Where-Object {$_.'Site Type' -eq 'OneDrive'} |Sort-Object -Property "Report Date"
    }else{
        $UsageReport = Import-Csv -Path $ReportCSV | Sort-Object -Property "Report Date"
    }
    
    $Record = 1
    $StorageUsage = @()
    foreach ($item in $UsageReport) {
        if ($Record -eq 1){
            $StorageUsed = $Item."Storage Used (Byte)"
        }else {
            $StorageUsage += (
                New-Object psobject -Property @{
                    Growth =  [math]::Round(((($Item.'Storage Used (Byte)' / $StorageUsed) -1) * 100),2)
                }
            )
            $StorageUsed = $Item."Storage Used (Byte)"
        }
        $Record = $Record + 1
    }
    
    $AverageGrowth = ($StorageUsage | Measure-Object -Property Growth -Average).Average
    # AverageGrowth is based on 180 days. This is not annual growth. To provide an annual growth we will take the value of AverageGrowth * 2 and then round up to the nearest whole percentage. While this is not exact, it should be close enough for our purposes.
    $AverageGrowth = [math]::Ceiling(($AverageGrowth * 2)) 
    return $AverageGrowth
}
function ProcessUsageReport {
    param (
        [Parameter(Mandatory)]
        [string]$ReportCSV, 
        [Parameter(Mandatory)]
        [string]$ReportName,
        [Parameter(Mandatory)]
        [string]$Section
    )


    $ReportDetail = Import-Csv -Path $ReportCSV | Where-Object {$_.'Is Deleted' -eq 'FALSE'}
    if (($AzureAdRequired) -and ($Section -ne "SharePoint")) {
        # The OneDrive and Exchange Usage reports have different column names that need to be accounted for.
        if ($Section -eq "OneDrive") {
            $FilterByField = "Owner Principal Name"
        } else {
            $FilterByField = "User Principal Name"
        }
        
        $SummarizedData = $ReportDetail | Where-Object {$_.$FilterByField -in $AzureAdGroupMembersByUserPrincipalName} | Measure-Object -Property 'Storage Used (Byte)' -Sum -Average


    } else {

        $SummarizedData = $ReportDetail | Measure-Object -Property 'Storage Used (Byte)' -Sum -Average

    }
    switch ($Section) {
        'SharePoint' { $M365Sizing.$($Section).NumberOfSites = $SummarizedData.Count }
        Default {$M365Sizing.$($Section).NumberOfUsers = $SummarizedData.Count}
    }
    $M365Sizing.$($Section).TotalSizeGB = [math]::Round(($SummarizedData.Sum / 1GB), 2, [MidPointRounding]::AwayFromZero)
    $M365Sizing.$($Section).SizePerUserGB = [math]::Round((($SummarizedData.Average) / 1GB), 2)
} 

if ([string]::IsNullOrEmpty($AzureAdGroupName)) {
    $AzureAdRequired = $false
} else {
    $AzureAdRequired = $true
}


# Validate the required 'Microsoft.Graph.Reports' is installed
# and provide a user friendly message when it's not.
if (Get-Module -ListAvailable -Name Microsoft.Graph.Reports)
{
    
}
else
{
    throw "The 'Microsoft.Graph.Reports' is required for this script. Run the follow command to install: Install-Module Microsoft.Graph.Reports"
}

# Validate the required 'ExchangeOnlineManagement' is installed
# and provide a user friendly message when it's not.
if (Get-Module -ListAvailable -Name ExchangeOnlineManagement)
{
    
}
else
{
    throw "The 'ExchangeOnlineManagement' is required for this script. Run the follow command to install: Install-Module ExchangeOnlineManagement"
}

if ($AzureAdRequired) {
    # Validate the required 'Azure.Graph.Authentication' is installed
    # and provide a user friendly message when it's not.
    if (Get-Module -ListAvailable -Name Microsoft.Graph.Groups)
    {
        
    }
    else
    {
        throw "The 'Microsoft.Graph.Groups' is required for filtering by a specific Azure AD Group. Run the follow command to install: Install-Module Microsoft.Graph.Groups"
    }
}


Write-Output "[INFO] Connecting to the Microsoft Graph API using 'Reports.Read.All', 'User.Read.All', and 'Group.Read.All' (if filtering results by Azure AD Group) permissions."
try {
    Connect-MgGraph -Scopes "Reports.Read.All","User.Read.All","Group.Read.All"  | Out-Null
}
catch {
    $errorException = $_.Exception
    $errorMessage = $errorException.Message
    Write-Output "[ERROR] Unable to Connect to the Microsoft Graph PowerShell Module: $errorMessage"
}

Write-Output "[INFO] Looking up all users in the provided Azure AD Group."
if ($AzureAdRequired) {
    $AzureAdGroupDetails = Get-MgGroup -Filter "DisplayName eq '$AzureAdGroupName'"
    
    if ($AzureAdGroupDetails.Count -eq 0) {
        throw "The Azure AD Group '$AzureAdGroupName' does not exist."
    }

    $AzureAdGroupMembersById = Get-MgGroupMember -GroupId $AzureAdGroupDetails.Id -All

    if ($EnableDebug) {
        Write-Output "[DEBUG] Azure AD Group Members Size: $($AzureAdGroupMembersById.Count)"
    }

    $AzureAdGroupMembersByUserPrincipalName = @()
    $AzureAdGroupMembersById | Foreach-Object  {
            if ($_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.user"){
             $AzureAdGroupMembersByUserPrincipalName += $_.AdditionalProperties["userPrincipalName"]
         }
     }

    
     Write-Output "[INFO] Discovered $($AzureAdGroupMembersByUserPrincipalName.Count) users in the provided Azure AD Group."
}



if ($EnableDebug) {
    try {
        $user = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/me"
        $permissions = Get-MgUserOauth2PermissionGrant -UserId $user.id
        
        Write-Output "[DEBUG] The authenticated user account has the following permissions:$($permissions.Scope)"
        
    }
    catch {
        $errorMessage = $_.Exception | Out-String
                
                
        throw $_.Exception
    }
}

$M365Sizing = [ordered]@{
    Exchange = [ordered]@{
        NumberOfUsers = 0
        TotalSizeGB   = 0
        SizePerUserGB = 0
        AverageGrowthPercentage = 0
        OneYearStorageForecastInGB = 0
        ThreeYearStorageForecastInGB = 0
    }
    OneDrive = [ordered]@{
        NumberOfUsers = 0
        TotalSizeGB   = 0
        SizePerUserGB = 0
        AverageGrowthPercentage = 0
        OneYearStorageForecastInGB = 0
        ThreeYearStorageForecastInGB = 0
    }
    SharePoint = [ordered]@{
        NumberOfSites = 0
        TotalSizeGB   = 0
        SizePerUserGB = 0
        AverageGrowthPercentage = 0
        OneYearStorageForecastInGB = 0
        ThreeYearStorageForecastInGB = 0
    }
    Licensing = [ordered]@{
        # Commented out for now, but we can get the number of licensed users if required (Not just activated).
        # Exchange         = 0
        # OneDrive         = 0
        # SharePoint       = 0
        # Teams            = 0
    }
    TotalDataToProtect = [ordered]@{
        OneYearInGB = 0
        ThreeYearInGB   = 0
    }

    # Teams = @{
    #     NumberOfUsers = 0
    #     TotalSizeGB = 0
    #     SizePerUserGB = 0
    #     AverageGrowthPercentage = 0
    # }
}


#region Usage Detail Reports
# Run Usage Detail Reports for different sections to get counts, total size of each section and average size. 
# We will only capture data that [Is Deleted] is equal to false. If [Is Deleted] is equal to True then that account has been deleted 
# from the customers M365 Tenant. It should not be counted in the sizing reports as We will not backup those objects. 
$UsageDetailReports = @{}
$UsageDetailReports.Add('Exchange', 'getMailboxUsageDetail')
$UsageDetailReports.Add('OneDrive', 'getOneDriveUsageAccountDetail')
$UsageDetailReports.Add('SharePoint', 'getSharePointSiteUsageDetail')

Write-Output "[INFO] Retrieving the Total Storage Consumed for ..."
foreach($Section in $UsageDetailReports.Keys){
    Write-Output " - $Section"
    $ReportCSV = Get-MgReport -ReportName $UsageDetailReports[$Section] -Period $Period
    ProcessUsageReport -ReportCSV $ReportCSV -ReportName $UsageDetailReports[$Section] -Section $Section
    Remove-Item -Path $ReportCSV
}

#endregion


#region Storage Usage Reports
# Run Storage Usage Reports for each section get get a trend of storage used for the period provided. We will get the growth percentage
# for each day and then average them all across the period provided. This way we can take into account the growth or the reduction 
# of storage used across the entire period. 
$StorageUsageReports = @{}
$StorageUsageReports.Add('Exchange', 'getMailboxUsageStorage')
$StorageUsageReports.Add('OneDrive', 'getOneDriveUsageStorage')
$StorageUsageReports.Add('SharePoint', 'getSharePointSiteUsageStorage')
Write-Output "[INFO] Retrieving the Average Storage Growth Forecast for ..."



foreach($Section in $StorageUsageReports.Keys){
    Write-Output " - $Section"
    $ReportCSV = Get-MgReport -ReportName $StorageUsageReports[$Section] -Period $Period
    $AverageGrowth = Measure-AverageGrowth -ReportCSV $ReportCSV -ReportName $StorageUsageReports[$Section]
    $M365Sizing.$($Section).AverageGrowthPercentage = [math]::Round($AverageGrowth,2)
    Remove-Item -Path $ReportCSV
}


#endregion



#region License usage
# Write-Output "[INFO] Retrieving the subscription License details."
# $licenseReportPath = Get-MgReport -ReportName getOffice365ActiveUserDetail -Period 180
# $licenseReport = Import-Csv -Path $licenseReportPath | Where-Object 'is deleted' -eq 'FALSE'


# # Clean up temp CSV
# Remove-Item -Path $licenseReportPath

# $licensesToIgnore = "POWER APPS PER USER PLAN","DYNAMICS 365 REMOTE ASSIST","POWER AUTOMATE PER USER PLAN","BUSINESS APPS (FREE)","MICROSOFT BUSINESS CENTER","DYNAMICS 365 GUIDES","POWERAPPS PER APP BASELINE","MICROSOFT MYANALYTICS","MICROSOFT 365 PHONE SYSTEM","POWER BI PRO","AZURE ACTIVE DIRECTORY PREMIUM","MICROSOFT INTUNE","DYNAMICS 365 TEAM MEMBERS","SECURITY E3","ENTERPRISE MOBILITY","MICROSOFT WORKPLACE ANALYTICS","MICROSOFT POWER AUTOMATE FREE","MICROSOFT TEAMS EXPLORATORY","MICROSOFT STREAM TRIAL", "VISIO PLAN 2","MICROSOFT POWER APPS PLAN 2 TRIAL","DYNAMICS 365 CUSTOMER ENGAGEMENT PLAN","DYNAMICS 365 BUSINESS CENTRAL ESSENTIAL","PROJECT PLAN","DYNAMICS 365 BUSINESS CENTRAL FOR IWS","PROJECT ONLINE ESSENTIALS","MICROSOFT TEAMS TRIAL","POWERAPPS AND LOGIC FLOWS","DYNAMICS 365 CUSTOMER VOICE TRIAL","MICROSOFT DEFENDER FOR ENDPOINT","DYNAMICS 365 SALES PREMIUM VIRAL TRIAL","DYNAMICS 365 P1 TRIAL FOR INFORMATION WORKERS","POWER BI (FREE)","APP CONNECT", "AZURE ACTIVE DIRECTORY PREMIUM P1","DYNAMICS 365 UNIFIED OPERATIONS PLAN","MICROSOFT DYNAMICS AX7  USER TRIAL","MICROSOFT DYNAMICS AX7 USER TRIAL","MICROSOFT POWER APPS PLAN 2 (QUALIFIED OFFER)","POWER APPS PER USER PLAN - GLOBAL","POWERAPPS PER APP BASELINE ACCESS","RIGHTS MANAGEMENT ADHOC","VISIO PLAN 1",""

# $assignedProducts = $licenseReport | ForEach-Object {$_.'Assigned Products'.Split('+')} | Group-Object | Select-Object Name,Count

# $assignedProducts | ForEach-Object {if ($_.name -NotIn $licensesToIgnore) {$M365Sizing.Licensing.Add($_.name, $_.count)}}

Write-Output "[INFO] Disconnecting from the Microsoft Graph API."
Disconnect-MgGraph

# The Microsoft Exchange Reports do not contain In-Place Archive sizing information so we also need to connect to the Exchange Online module to
# get this information
Write-Output "[INFO] Switching to the Microsoft Exchange Online Module for more detailed reporting capabilities."
Connect-ExchangeOnline -ShowBanner:$false
$ManualUserPrincipalName = $null 
$ActionRequiredLogMessage = "[ACTION REQUIRED] In order to periodically refresh the connection to Microsoft, we need the User Principal Name used during the authentication process."
$ActionRequiredPromptMessage = "Enter the User Principal Name"
Write-Output "[INFO] Retrieving all Exchange Mailbox In-Place Archive sizing."


$FirstInterval = 500
$SkipInternval = $FirstInterval
$ArchiveMailboxSizeGb = 0
$LargeAmountofArchiveMailboxCount = 5000
try {
    
    $ArchiveMailboxes = Get-ExoMailbox -Archive -ResultSize Unlimited
    $ArchiveMailboxesCount = @($ArchiveMailboxes).Count

    $ArchiveMailboxesFolders = @()
    # Process the first N number of Archive Mailboxes. Where N = $FirstInterval
    $ArchiveMailboxesFirstInverval = $ArchiveMailboxes | Select-Object -First $FirstInterval
    if ($ArchiveMailboxesCount -le $LargeAmountofArchiveMailboxCount) {
        $ArchiveMailboxesFolders += $ArchiveMailboxesFirstInverval| Get-EXOMailboxFolderStatistics -Archive -Folderscope "Archive" | Select-Object name,FolderAndSubfolderSize

    } else {
        Write-Output "[INFO] Detected a large number of Archive Mailboxes. Implementing additional logic to account for Microsoft API performance limits. This may take some time."
        Write-Output ""
        Write-Output $ActionRequiredLogMessage
        Write-Output ""
        $ManualUserPrincipalName = Read-Host -Prompt $ActionRequiredPromptMessage
        $ArchiveMailboxesFolders +=  Start-RobustCloudCommand -UserPrincipalName $ManualUserPrincipalName -IdentifyingProperty "DisplayName" -recipients $ArchiveMailboxesFirstInverval -logfile "$systemTempFolder\archiveMailbox.log" -ScriptBlock {Get-EXOMailboxFolderStatistics -Identity $input.UserPrincipalName -Archive -Folderscope "Archive" | Select-Object name,FolderAndSubfolderSize }     
        Write-Output ""
       
    }
    
    # Process any remaining Archive Mailboxes at the pre-defined $FirstInterval
    if ($ArchiveMailboxesCount -ge $FirstInterval){

        while($ArchiveMailboxesCount -ge 0)
        {   
            $ArchiveMailboxesCount = $ArchiveMailboxesCount - $FirstInterval
            $ArchiveMailboxesSecondaryInverval = $ArchiveMailboxes | Select-Object -Skip $SkipInternval -First $FirstInterval 

            if ($ArchiveMailboxesCount -le $LargeAmountofArchiveMailboxCount) {
                $ArchiveMailboxesFolders += $ArchiveMailboxesSecondaryInverval | Get-EXOMailboxFolderStatistics -Archive -Folderscope "Archive" | Select-Object name,FolderAndSubfolderSize
            } else {
                $ArchiveMailboxesFolders += Start-RobustCloudCommand -UserPrincipalName $ManualUserPrincipalName -IdentifyingProperty "DisplayName" -recipients $ArchiveMailboxesSecondaryInverval -logfile "$systemTempFolder\archiveMailbox.log" -ScriptBlock { Get-EXOMailboxFolderStatistics -Identity $input.UserPrincipalName -Archive -Folderscope "Archive" | Select-Object name,FolderAndSubfolderSize }     
                Write-Output ""
            }
            $SkipInternval = $SkipInternval + $FirstInterval
        }

    }
    # Remove the Start-RobustCloudCommand log file if it exists
    Remove-Item -Path "$systemTempFolder\archiveMailbox.log" -ErrorAction SilentlyContinue
    
    foreach($Folder in $ArchiveMailboxesFolders){
        $FolderSize = $Folder.FolderAndSubfolderSize.ToString().split("(") | Select-Object -Index 1 
        $FolderSizeBytes = $FolderSize.split("bytes") | Select-Object -Index 0
        
        $FolderSizeInGb = [math]::Round(([int64]$FolderSizeBytes / 1GB), 3, [MidPointRounding]::AwayFromZero)

        $ArchiveMailboxSizeGb += $FolderSizeInGb
    }
}
catch {
    $errorException = $_.Exception
    $errorMessage = $errorException.Message
    Write-Output "[ERROR] Unable to retrieve In-Place Archive sizing. $errorMessage "
}

Write-Output "[INFO] Retrieving Exchange Mailbox Shared Mailbox sizing."
# Reset First and Skip interval values
$FirstInterval = 500
$SkipInternval = $FirstInterval
$SharedMailboxesSizeGb = 0
$LargeAmountofSharedMailboxCount = 5000

try {
    # Process the first N number of Shared Mailboxes. Where N = $FirstInterval
    $SharedMailboxes = Get-ExoMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited
    $SharedMailboxesCount = @($SharedMailboxes).Count
    
    
    $SharedMailboxeFirstInterval = $SharedMailboxes | Select-Object -First $FirstInterval
    if ($SharedMailboxesCount -le $LargeAmountofSharedMailboxCount) {
        $SharedMailboxesSize += $SharedMailboxeFirstInterval | Get-ExoMailboxStatistics | Select-Object TotalItemSize     

    } else {
        Write-Output "[INFO] Detected a large number of Shared Mailboxes. Implementing additional logic to account for Microsoft API performance limits. This may take some time."
        Write-Output ""
        if ($ManualUserPrincipalName -eq $null) {
            Write-Output ""
            Write-Output $ActionRequiredLogMessage
            Write-Output ""
            $ManualUserPrincipalName = Read-Host -Prompt $ActionRequiredPromptMessage
        }
       
        $SharedMailboxesSize +=  Start-RobustCloudCommand -UserPrincipalName $ManualUserPrincipalName -IdentifyingProperty "DisplayName" -recipients $SharedMailboxeFirstInterval -logfile "$systemTempFolder\sharedMailbox.log" -ScriptBlock {Get-ExoMailboxStatistics -Identity $input.UserPrincipalName | Select-Object TotalItemSize}
        Write-Output ""
    }

    # Process any remaining Shared Mailboxes at the pre-defined $FirstInterval
    if ($SharedMailboxesCount -ge $FirstInterval){


        while($SharedMailboxesCount -ge 0)
        {   
            $SharedMailboxesCount = $SharedMailboxesCount - $FirstInterval
            $SharedMailboxesSecondaryInterval = $SharedMailboxes | Select-Object -Skip $SkipInternval -First $FirstInterval
            if ($SharedMailboxesCount -le $LargeAmountofSharedMailboxCount) {
                $SharedMailboxesSize += $SharedMailboxesSecondaryInterval | Get-ExoMailboxStatistics| Select-Object TotalItemSize
        
            } else {
                $SharedMailboxesSize +=  Start-RobustCloudCommand -UserPrincipalName $ManualUserPrincipalName -IdentifyingProperty "DisplayName" -recipients $SharedMailboxesSecondaryInterval -logfile "$systemTempFolder\sharedMailbox.log" -ScriptBlock {Get-ExoMailboxStatistics -Identity $input.UserPrincipalName| Select-Object TotalItemSize}
                Write-Output ""
            }
            
            $SkipInternval = $SkipInternval + $FirstInterval
        }

    }

    # Remove the Start-RobustCloudCommand log file if it exists
    Remove-Item -Path "$systemTempFolder\sharedMailbox.log" -ErrorAction SilentlyContinue

    foreach($Folder in $SharedMailboxesSize){
        $FolderSize = $Folder.TotalItemSize.Value.ToString().split("(") | Select-Object -Index 1
        $FolderSizeBytes = $FolderSize.split("bytes") | Select-Object -Index 0
        
        $FolderSizeInGb = [math]::Round(([int64]$FolderSizeBytes / 1GB), 3, [MidPointRounding]::AwayFromZero)

        $SharedMailboxesSizeGb += $FolderSizeInGb
    }

}
catch {
    $errorException = $_.Exception
    $errorMessage = $errorException.Message
    Write-Output "[ERROR] Unable to retrieve Shared Mailbox sizing. $errorMessage"
}

$M365Sizing.Exchange.TotalSizeGB += $ArchiveMailboxSizeGb
$M365Sizing.Exchange.TotalSizeGB += $SharedMailboxesSizeGb

Write-Output "[INFO] Disconnecting from the Microsoft Exchange Online Module"
Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue

Write-Output "[INFO] Calculating the forecasted total storage need for Rubrik."
foreach($Section in $M365Sizing | Select-Object -ExpandProperty Keys){

    if ( $Section -NotIn @("Licensing", "TotalDataToProtect") )
    {
        $M365Sizing.$($Section).OneYearStorageForecastInGB = $M365Sizing.$($Section).TotalSizeGB * (1.0 + (($M365Sizing.$($Section).AverageGrowthPercentage / 100) * 1))
        $M365Sizing.$($Section).ThreeYearStorageForecastInGB = $M365Sizing.$($Section).TotalSizeGB * (1.0 + (($M365Sizing.$($Section).AverageGrowthPercentage / 100) * 3))
    
        $M365Sizing.TotalDataToProtect.OneYearInGB = $M365Sizing.TotalDataToProtect.OneYearInGB + $M365Sizing.$($Section).OneYearStorageForecastInGB
        $M365Sizing.TotalDataToProtect.ThreeYearInGB = $M365Sizing.TotalDataToProtect.ThreeYearInGB + $M365Sizing.$($Section).ThreeYearStorageForecastInGB
    }

}

# Calculate the total number of licenses required
if ($SharedMailboxesCount -gt $M365Sizing.Exchange.NumberOfUsers){
    Write-Output "[INFO] Detected more Shared Mailboxes than User Mailboxes. Automatically updating license count requirements."
    $M365Sizing.Exchange.NumberOfUsers = $SharedMailboxesCount
} 

if ($M365Sizing.Exchange.NumberOfUsers -gt $M365Sizing.OneDrive.NumberOfUsers){
    $UserLicensesRequired = $M365Sizing.Exchange.NumberOfUsers
} else {
    $UserLicensesRequired = $M365Sizing.OneDrive.NumberOfUsers
}

$Calculate_Users_Required=[math]::ceiling($UserLicensesRequired)
$Calculate_Storage_Required=[math]::ceiling($($M365Sizing[4].OneYearInGB))

# Query M365Licsolver Azure Function
# If less than 76GB Average per user then query the azure function that calculates the best mix of subscription types. If more than 76 then Unlimited is the best option.
if (($Calculate_Storage_Required)/$Calculate_Users_Required -le 76) {

    # Query the M365Licsolver Azure Function
    $SolverQuery = '{"users":"' + $Calculate_Users_Required + '","data":"' + $Calculate_Storage_Required + '"}'
    try {
        $APIReturn = ConvertFrom-JSON (Invoke-WebRequest 'https://m365licsolver-azure.azurewebsites.net:/api/httpexample' -ContentType "application/json" -Body $SolverQuery -Method 'POST')
    }
    catch {
        $errorMessage = $_.Exception | Out-String
        if($errorMessage.Contains('Response status code does not indicate success: 404')) {
            Write-Output "[Info] Unable to calculate license recommendations."
        } 
    }
    $FiveGBPacks=$APIReturn.FiveGBSubscriptions
    $TwentyGBPacks=$APIReturn.TwentyGBSubscriptions
    $FiftyGBPacks=$APIReturn.FiftyGBSubscriptions
    $UnlimitedGBPacks=0
    $UnlimitedGBUsers=0
    $FiveGBUsers=$FiveGBPacks*10
    $TwentyGBUsers=$TwentyGBPacks*10
    $FiftyGBUsers=$FiftyGBPacks*10
    $TotalAmountUsers=$FiveGBUsers + $TwentyGBUsers + $FiftyGBUsers
    $TotalAmountStorage=($FiveGBUsers*5) + ($TwentyGBUsers*20) + ($FiftyGBUsers*50)
} else {
    $FiveGBPacks=0
    $TwentyGBPacks=0
    $FiftyGBPacks=0
    $FiveGBUsers=0
    $TwentyGBUsers=0
    $FiftyGBUsers=0
    $UnlimitedGBPacks=$Calculate_Users_Required=[math]::ceiling($UserLicensesRequired/10)
    $UnlimitedGBUsers=$UnlimitedGBPacks*10
    $TotalAmountUsers=$UnlimitedGBUsers
    $TotalAmountStorage="Unlimited"
}

#region HTML Code for Output
$HTML_CODE=@"                            
<!DOCTYPE html>
<html>
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
                    Exchange Online
                </div>
            </div>

            <table class="styled-table">
                <thead>
                    <tr>
                        <th>Number of Users</th>
                        <th>Total Size</th>
                        <th>Per User Size</th>
                        <th>Average Growth Forecast (Yearly)</th>
                        <th>One Year Storage Forecast</th>
                        <th>Three Year Storage Forecast</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>$($M365Sizing[0].NumberOfUsers)</td>
                        <td>$($M365Sizing[0].TotalSizeGB) GB</td>
                        <td>$($M365Sizing[0].SizePerUserGB) GB</td>
                        <td>$($M365Sizing[0].AverageGrowthPercentage)%</td>
                        <td>$($M365Sizing[0].OneYearStorageForecastInGB) GB</td>
                        <td>$($M365Sizing[0].ThreeYearStorageForecastInGB) GB</td>

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
                    OneDrive
                </div>
            </div>

            <table class="styled-table">
                <thead>
                    <tr>
                        <th>Number of Users</th>
                        <th>Total Size</th>
                        <th>Per User Size</th>
                        <th>Average Growth Forecast (Yearly)</th>
                        <th>One Year Storage Forecast</th>
                        <th>Three Year Storage Forecast</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>$($M365Sizing[1].NumberOfUsers)</td>
                        <td>$($M365Sizing[1].TotalSizeGB) GB</td>
                        <td>$($M365Sizing[1].SizePerUserGB) GB</td>
                        <td>$($M365Sizing[1].AverageGrowthPercentage)%</td>
                        <td>$($M365Sizing[1].OneYearStorageForecastInGB) GB</td>
                        <td>$($M365Sizing[1].ThreeYearStorageForecastInGB) GB</td>

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
                        <th>Number of Sites</th>
                        <th>Total Size</th>
                        <th>Per Site Size</th>
                        <th>Average Growth Forecast (Yearly)</th>
                        <th>One Year Storage Forecast</th>
                        <th>Three Year Storage Forecast</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>$($M365Sizing[2].NumberOfSites)</td>
                        <td>$($M365Sizing[2].TotalSizeGB) GB</td>
                        <td>$($M365Sizing[2].SizePerUserGB) GB</td>
                        <td>$($M365Sizing[2].AverageGrowthPercentage)%</td>
                        <td>$($M365Sizing[2].OneYearStorageForecastInGB) GB</td>
                        <td>$($M365Sizing[2].ThreeYearStorageForecastInGB) GB</td>
                    </tr>

                    
                </tbody>
            </table>
        </div>
    </div>

    <!-- Licensing -->
    <!-- <div class="card-container">
        <div class="card">
            <h1>Licensing</h1>
            <table class="styled-table">
                <thead>
                    <tr>
                        <th>Number of Users</th>
                        <th>Total Size (GB)</th>
                        <th>Per User Size (GB)</th>
                        <th>Average Growth Forecast (Yearly)</th>
                        <th>One Year Storage Forecast (GB)</th>
                        <th>Three Year Storage Forecast (GB)</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>296</td>
                        <td>1.26</td>
                        <td>0</td>
                        <td>8</td>
                        <td>1.3608</td>
                        <td>1.5624</td>

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
                

                <svg xmlns="http://www.w3.org/2000/svg" height="62" width="70" viewBox="0 0 278050 333334" shape-rendering="geometricPrecision" text-rendering="geometricPrecision" image-rendering="optimizeQuality" fill-rule="evenodd" clip-rule="evenodd">
                <path fill="#ea3e23" d="M278050 305556l-29-16V28627L178807 0 448 66971l-448 87 22 200227 60865-23821V80555l117920-28193-17 239519L122 267285l178668 65976v73l99231-27462v-316z"/></svg>

                

                </div>
                <div class="card-header-text">
                    Discovery Summary
                </div>
                </div>

                <table class="styled-table">
                    <thead>
                        <tr>
                            <th>Required Number of Licenses</th>
                            <th>One Year Storage Forecast</th>
                            <th>Three Year Storage Forecast</th>
                            
    
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td>$UserLicensesRequired</td>
                            <td>$($M365Sizing[4].OneYearInGB) GB</td>
                            <td>$($M365Sizing[4].ThreeYearInGB) GB</td>
                     
    
    
                        </tr>
    
                        
                    </tbody>
                </table>
            </div>
        </div>




    <!-- Licensing -->
    <!-- <div class="card-container">
        <div class="card">
            <h1>Licensing</h1>
            <table class="styled-table">
                <thead>
                    <tr>
                        <th>Number of Users</th>
                        <th>Total Size (GB)</th>
                        <th>Per User Size (GB)</th>
                        <th>Average Growth Forecast (Yearly)</th>
                        <th>One Year Storage Forecast (GB)</th>
                        <th>Three Year Storage Forecast (GB)</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>296</td>
                        <td>1.26</td>
                        <td>0</td>
                        <td>8</td>
                        <td>1.3608</td>
                        <td>1.5624</td>

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

                    License Recommendation 

       
                </div>
            </div>

            <table class="styled-table">
                <thead>
                    <tr>
                        <th>Starter (5GB)</th>
                        <th>Foundation (20GB)</th>
                        <th>Business (50GB)</th>
                        <th>Enterprise (Unlimited)</th>
                     
                        

                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>$FiveGBUsers</td>
                        <td>$TwentyGBUsers</td>
                        <td>$FiftyGBUsers</td>
                        <td>$UnlimitedGBUsers </td>
                     

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

#region Start-RobustCloudCommand
# Source: https://github.com/Canthv0/RobustCloudCommand
# MIT License

# Copyright (c) 2019 Canthv0

# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:

# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.

# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.
Function Start-RobustCloudCommand {

    <#

.SYNOPSIS
Generic wrapper script that tries to ensure that a script block successfully finishes execution in O365 against a large object count.

Works well with intense operations that may cause throttling

.DESCRIPTION
Wrapper script that tries to ensure that a script block successfully finishes execution in O365 against a large object count.

It accomplishs this by doing the following:
* Monitors the health of the Remote powershell session and restarts it as needed.
* Restarts the session every X number seconds to ensure a valid connection.
* Attempts to work past session related errors and will skip objects that it can't process.
* Attempts to calculate throttle exhaustion and sleep a sufficient time to allow throttle recovery

.PARAMETER ActiveThrottle
Calculated value based on your tenants powershell recharge rate.
You tenant recharge rate can be calculated using a Micro Delay Warning message.

Look for the following line in your Micro Delay Warning Message
Balance: -1608289/2160000/-3000000

The middle value is the recharge rate.
Divide this value by the number of milliseconds in an hour (3600000)
And subtract the result from 1 to get your AutomaticThrottle value

1 - (2160000 / 3600000) = 0.4

Default Value is .25

.PARAMETER IdentifyingProperty
What property of the objects we are processing that will be used to identify them in the log file and host
If the value is not set by the user the script will attempt to determine if one of the following properties is present
"DisplayName","Name","Identity","PrimarySMTPAddress","Alias","GUID"

If the value is not set and we are not able to match a well known property the script will generate an error and terminate.

.PARAMETER LogFile
Location and file name for the log file.

.PARAMETER ManualThrottle
Manual delay of X number of milliseconds to sleep between each cmdlets call.
Should only be used if the AutomaticThrottle isn't working to introduce sufficent delay to prevent Micro Delays

.PARAMETER NonInteractive
Suppresses output to the screen.  All output will still be in the log file.

.PARAMETER Recipients
Array of objects to operate on. This can be mailboxes or any other set of objects.
Input must be an array!
Anything comming in from the array can be accessed in the script block using $input.property

.PARAMETER ResetSeconds
How many seconds to run the script block before we rebuild the session with O365.

.PARAMETER ScriptBlock
The script that you want to robustly execute against the array of objects.  The Recipient objects will be provided to the cmdlets in the script block
and can be accessed with $input as if you were pipelining the object.

.PARAMETER UserPrincipalName
UPN of the user that will be connecting to Exchange online.  Required so that sessions can automatically be set up using cached tokens.

.LINK
https://github.com/Canthv0/RobustCloudCommand

.OUTPUTS
Creates the log file specified in -logfile.  Logfile contains a record of all actions taken by the script.

.EXAMPLE
invoke-command -scriptblock {Get-mailbox -resultsize unlimited | select-object -property Displayname,PrimarySMTPAddress,Identity} -session (get-pssession) | export-csv c:\temp\mbx.csv

$mbx = import-csv c:\temp\mbx.csv

$cred = get-Credential

.\Start-RobustCloudCommand.ps1 -UserPrincipalName admin@contoso.com -recipients $mbx -logfile C:\temp\out.log -ScriptBlock {Set-Clutter -identity $input.PrimarySMTPAddress.tostring() -enable:$false}

Gets all mailboxes from the service returning only Displayname,Identity, and PrimarySMTPAddress.  Exports the results to a CSV
Imports the CSV into a variable
Gets your O365 Credential
Executes the script setting clutter to off using Legacy Credentials

.EXAMPLE
invoke-command -scriptblock {Get-mailbox -resultsize unlimited | select-object -property Displayname,PrimarySMTPAddress,Identity} -session (get-pssession) | export-csv c:\temp\recipients.csv

$recipients = import-csv c:\temp\recipients.csv

Start-RobustCloudCommand -UserPrincipalName admin@contoso.com -recipients $recipients -logfile C:\temp\out.log -ScriptBlock {Get-MobileDeviceStatistics -mailbox $input.PrimarySMTPAddress.tostring() | Select-Object -Property @{Name = "PrimarySMTPAddress";Expression={$input.PrimarySMTPAddress.tostring()}},DeviceType,LastSuccessSync,FirstSyncTime | Export-Csv c:\temp\stats.csv -Append }

Gets All Recipients and exports them to a CSV (for restart ability)
Imports the CSV into a variable
Executes the script to gather EAS Device statistics and output them to a csv file using ADAL with support for MFA


#>

    Param(
        [Parameter(Mandatory = $true)]
        [string]$LogFile,
        [Parameter(Mandatory = $true)]
        $Recipients,
        [Parameter(Mandatory = $true)]
		[ScriptBlock]$ScriptBlock,
		[Parameter(Mandatory = $true)]
		[String]$UserPrincipalName,
        [int]$ManualThrottle = 0,
        [double]$ActiveThrottle = .25,
        [int]$ResetSeconds = 870,
        [string]$IdentifyingProperty,
        [Switch]$NonInteractive
    )

    # Turns on strict mode https://technet.microsoft.com/library/03373bbe-2236-42c3-bf17-301632e0c428(v=wps.630).aspx
    Set-StrictMode -Version 2
    $InformationPreference = "Continue"
    $Global:ErrorActionPreference = "Stop"
    Write-Log ("Error Action Preference: " + $Global:ErrorActionPreference)
    Write-Log ("Information Preference: " + $InformationPreference)

    # Log the script block for debugging purposes
    Write-log $ScriptBlock

    # Setup our first session to O365
    $ErrorCount = 0
    New-CleanO365Session

    # Get when we started the script for estimating time to completion
    $ScriptStartTime = Get-Date
    [int]$ObjectsProcessed = 0
    [int]$ObjectCount = $Recipients.count

    # If we don't have an identifying property then try to find one
    if ([string]::IsNullOrEmpty($IdentifyingProperty)) {
        # Call our function for finding an identifying property and pass in the first recipient object
        $IdentifyingProperty = Get-ObjectIdentificationProperty -object $Recipients[0]
    }

    # Go thru each recipient object and execute the script block
    foreach ($object in $Recipients) {

        # Set our initial while statement values
        $TryCommand = $true
		$errorcount = 0
		$Global:Error.clear()

        # Try the command 3 times and exit out if we can't get it to work
        # Record the error and restart the session each time it errors out
        while ($TryCommand) {
            Write-log ("Running scriptblock for " + ($object.$IdentifyingProperty).tostring())

            # Test our connection and rebuild if needed
            Test-O365Session

            # Invoke the script block
            try {
                Invoke-Command -InputObject $object -ScriptBlock $ScriptBlock -ErrorAction Stop

                # Since we didn't get an error don't run again
                $TryCommand = $false

                # Increment the object processed count / Estimate time to completion
                $ObjectsProcessed = Get-EstimatedTimeToCompletion -ProcessedCount $ObjectsProcessed -TotalObjects $ObjectCount -StartTime $ScriptStartTime
            }
            catch {

                # Handle if we keep failing on the object
                if ($errorcount -ge 3) {
                    Write-Log ("[ERROR] - Object `"" + ($object.$IdentifyingProperty).tostring() + "`" has failed three times!")
                    Write-Log ("[ERROR] - Skipping Object")

                    # Increment the object processed count / Estimate time to completion
                    $ObjectsProcessed = Get-EstimatedTimeToCompletion -ProcessedCount $ObjectsProcessed -StartTime $ScriptStartTime

                    # Set trycommand to false so we abort the while loop
                    $TryCommand = $false
                }
                # Otherwise try the command again
                else {
                    if ($null -eq $Global:Error){
                        Write-Log "Global Error Null"
                        Write-Log ("Local Error: " + $Error)
                    }
                    else {
                        Write-Log $Global:Error
                    }

					Write-Log ("Rebuilding session and trying again")
					$ErrorCount++
                    # Create a new session in case the error was due to a session issue
                    New-CleanO365Session
                }
            }
        }
    }

    Write-Log "Script Complete Destroying PS Sessions"
    # Destroy any outstanding PS Session
    Get-PSSession | Remove-PSSession -Confirm:$false

    $Global:ErrorActionPreference = "Continue"
    Write-Log ("Error Action Preference: " + $Global:ErrorActionPreference)


}

# Writes output to a log file with a time date stamp
Function Write-Log {
    Param ([string]$string)

    # Get the current date
    [string]$date = Get-Date -Format G

    # Write everything to our log file
    ( "[" + $date + "] - " + $string) | Out-File -FilePath $LogFile -Append

    # If NonInteractive true then suppress host output
    if (!($NonInteractive)) {
        Write-Information ( "[" + $date + "] - " + $string)
    }
}

# Sleeps X seconds and displays a progress bar
Function Start-SleepWithProgress {
    Param([int]$sleeptime)

    # Loop Number of seconds you want to sleep
    For ($i = 0; $i -le $sleeptime; $i++) {
        $timeleft = ($sleeptime - $i);

        # Progress bar showing progress of the sleep
        Write-Progress -Activity "Sleeping" -CurrentOperation "$Timeleft More Seconds" -PercentComplete (($i / $sleeptime) * 100) -Status " "

        # Sleep 1 second
        start-sleep 1
    }

    Write-Progress -Completed -Activity "Sleeping" -Status " "
}

# Setup a new O365 Powershell Session
Function New-CleanO365Session {

    # Destroy any outstanding PS Session
    Write-Log "Removing all PS Sessions"
    Get-PSSession | Remove-PSSession -Confirm:$false

    # Force Garbage collection just to try and keep things more agressively cleaned up due to some issue with large memory footprints
    [System.GC]::Collect()

    # Sleep 15s to allow the sessions to tear down fully
    Write-Log ("Sleeping 15 seconds for Session Tear Down")
    Start-SleepWithProgress -SleepTime 15

    # Clear out all errors
    $Error.Clear()

    # Create the session
	Write-Log "Connecting to Exchange Online"
	Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowBanner:$false

    # Check for an error while creating the session
    if ($Error.Count -gt 0) {

        Write-Log "[ERROR] - Error while setting up session"
        Write-log $Error

        # Increment our error count so we abort after so many attempts to set up the session
        $ErrorCount++

        # if we have failed to setup the session > 3 times then we need to abort because we are in a failure state
        if ($ErrorCount -gt 3) {

            Write-log "[ERROR] - Failed to setup session after multiple tries"
            Write-log "[ERROR] - Aborting Script"
            exit

        }

        # If we are not aborting then sleep 60s in the hope that the issue is transient
        Write-Log "Sleeping 60s so that issue can potentially be resolved"
        Start-SleepWithProgress -sleeptime 60

        # Attempt to set up the sesion again
        New-CleanO365Session
    }

    # If the session setup worked then we need to set $errorcount to 0
    else {
        $ErrorCount = 0
    }

    # Set the Start time for the current session
    Set-Variable -Scope script -Name SessionStartTime -Value (Get-Date)
}

# Verifies that the connection is healthy
# Goes ahead and resets it every $ResetSeconds number of seconds either way
Function Test-O365Session {

    # Get the time that we are working on this object to use later in testing
    $ObjectTime = Get-Date

    # Reset and regather our session information
    $SessionInfo = $null
    $SessionInfo = Get-PSSession

    # Make sure we found a session
    if ($null -eq $SessionInfo) {
        Write-Log "[ERROR] - No Session Found"
        Write-log "Recreating Session"
        New-CleanO365Session
    }
    # Make sure it is in an opened state if not log and recreate
    elseif ($SessionInfo.State -ne "Opened") {
        Write-Log "[ERROR] - Session not in Open State"
        Write-log ($SessionInfo | Format-List | Out-String )
        Write-log "Recreating Session"
        New-CleanO365Session
    }
    # If we have looped thru objects for an amount of time gt our reset seconds then tear the session down and recreate it
    elseif (($ObjectTime - $SessionStartTime).totalseconds -gt $ResetSeconds) {
        Write-Log ("Session Has been active for greater than " + $ResetSeconds + " seconds" )
        Write-Log "Rebuilding Connection"

        # Estimate the throttle delay needed since the last session rebuild
        # Amount of time the session was allowed to run * our activethrottle value
        # Divide by 2 to account for network time, script delays, and a fudge factor
        # Subtract 15s from the results for the amount of time that we spend setting up the session anyway
        [int]$DelayinSeconds = ((($ResetSeconds * $ActiveThrottle) / 2) - 15)

        # If the delay is >15s then sleep that amount for throttle to recover
        if ($DelayinSeconds -gt 0) {

            Write-Log ("Sleeping " + $DelayinSeconds + " addtional seconds to allow throttle recovery")
            Start-SleepWithProgress -SleepTime $DelayinSeconds
        }
        # If the delay is <15s then the sleep already built into New-CleanO365Session should take care of it
        else {
            Write-Log ("Active Delay calculated to be " + ($DelayinSeconds + 15) + " seconds no addtional delay needed")
        }

        # new O365 session and reset our object processed count
        New-CleanO365Session
    }
    else {
        # If session is active and it hasn't been open too long then do nothing and keep going
    }

    # If we have a manual throttle value then sleep for that many milliseconds
    if ($ManualThrottle -gt 0) {
        Write-log ("Sleeping " + $ManualThrottle + " milliseconds")
        Start-SleepWithProgress -Milliseconds $ManualThrottle
    }
}

# If the $identifyingProperty has not been set then we attempt to locate a value for tracking modified objects
Function Get-ObjectIdentificationProperty {
    Param($object)

    Write-Log "Trying to identify a property for displaying per object progress"

    # Common properties to check
    [array]$PropertiesToCheck = "DisplayName", "Name", "Identity", "PrimarySMTPAddress", "Alias", "GUID"

    # Set our counter to 0
    $i = 0
    [string]$PropertiesString = $null
    [bool]$Found = $false

    # While we haven't found an ID property continue checking
    while ($found -eq $false) {

        # If we have gone thru the list then we need to throw an error because we don't have Identity information
        # Set the string to bogus just to ensure we will exit the while loop
        if ($i -gt ($PropertiesToCheck.length - 1)) {
            Write-Log "[ERROR] - Unable to find a common identity parameter in the input object"

            # Create an error message that has all of the valid property names that we are looking for
            ForEach ($value in $PropertiesToCheck) { [string]$PropertiesString = $PropertiesString + "`"" + $value + "`", " }
            $PropertiesString = $PropertiesString.TrimEnd(", ")
            [string]$errorstring = "Objects does not contain a common identity parameter " + $PropertiesString + " please use -IdentifyingProperty to set the identity value"

            # Throw error
            Write-Error -Message $errorstring -ErrorAction Stop
        }

        # Get the property we are testing out of our array
        [string]$Property = $PropertiesToCheck[$i]

        # Check the properties of the object to see if we have one that matches a well known name
        # If we have found one set the value to that property
        if ($null -ne $object.$Property) {
            Write-log ("Found " + $Property + " to use for displaying per object progress")
            $found = $true
            Return $Property
        }

        # Increment our position counter
        $i++

    }
}

# Gather and print out information about how fast the script is running
Function Get-EstimatedTimeToCompletion {
    param([int]$ProcessedCount,[int]$TotalObjects, [datetime]$StartTime)

    # Increment our count of how many objects we have processed
    $ProcessedCount++

    # Every 100 we need to estimate our completion time and write that out
    if (($ProcessedCount % 100) -eq 0) {

        # Get the current date
        $CurrentDate = Get-Date

        # Average time per object in seconds
        $AveragePerObject = (((($CurrentDate) - $StartTime).totalseconds) / $ProcessedCount)

        # Write out session stats and estimated time to completion
        Write-Log ("[STATS] - Total Number of Objects:     " + $TotalObjects)
        Write-Log ("[STATS] - Number of Objects processed: " + $ProcessedCount)
        Write-Log ("[STATS] - Average seconds per object:  " + $AveragePerObject)
        Write-Log ("[STATS] - Estimated completion time:   " + $CurrentDate.addseconds((($TotalObjects - $ProcessedCount) * $AveragePerObject)))
    }

    # Return number of objects processed so that the variable in incremented
    return $ProcessedCount
}

#endregion
# Remove any previously created files
Remove-Item -Path .\Rubrik-M365-Sizing.html -ErrorAction SilentlyContinue
Write-Output $HTML_CODE |Format-Table -AutoSize | Out-File -FilePath .\Rubrik-M365-Sizing.html -Append


 
Write-Output "`n`nM365 Sizing information has been written to $((Get-ChildItem Rubrik-M365-Sizing.html).FullName)`n`n"
if ($OutputObject) {
    return $M365Sizing
}
