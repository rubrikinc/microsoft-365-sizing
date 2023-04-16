[CmdletBinding()]
param (
    [Parameter()]
    [Array]$AzureAdGroups
    
)

Connect-MgGraph -Scopes "Reports.Read.All","User.Read.All","Group.Read.All"  | Out-Null
Write-Output "Listing all users for the Azure Ad Group:"

foreach ($AzureAdGroupName in $AzureAdGroups ){
   
        # Write-Output " -`$AzureAdGroupName`"
        Write-Output "   -$AzureAdGroupName"
    
        $AzureAdGroupDetails = Get-MgGroup -Filter "DisplayName eq '$AzureAdGroupName'"
    
        $AzureAdGroupMembersByUserPrincipalName = @()
    
    
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
    
    }


$AzureAdGroupMembersByUserPrincipalName| Export-Csv -Path ./Users.csv 

$AzureAdGroupMembersByUserPrincipalName | ConvertTo-Csv -NoTypeInformation



Write-Output $AzureAdGroupMembersByUserPrincipalName

