$Users = 2727
$GBSize = 23500

$Calculate_Users_Required=[math]::ceiling($Users) 
$Calculate_Storage_Required=[math]::ceiling($GBSize)

$SolverQuery = '{"users":"' + $Calculate_Users_Required + '","data":"' + $Calculate_Storage_Required + '"}'
$APIReturn = ConvertFrom-JSON (Invoke-WebRequest 'https://m365licsolver-azure.azurewebsites.net:/api/httpexample' -ContentType "application/json" -Body $SolverQuery -Method 'POST')

Write-Host "5GB:  "($APIReturn.FiveGBSubscriptions * 10)
Write-Host "20GB: "($APIReturn.TwentyGBSubscriptions * 10)
Write-Host "50GB: "($APIReturn.FiftyGBSubscriptions * 10)