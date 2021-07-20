Write-Output "Start loading user profile task"

[string]$AppPools = Get-VstsInput -Name AppPools -Require

Write-Output "Load profiles for AppPools: $AppPools"

Import-Module WebAdministration; 
              
if ($AppPools -eq $null){
    $AppPools='';
}

$AppPoolsArray = $AppPools.Split(",")
if ($AppPoolsArray.Count -gt 0){
    $AppPoolsArray | ForEach-Object {
        $appPool=$_.Trim()
        Set-ItemProperty "IIS:\AppPools\$appPool" -Name "processModel.loadUserProfile" -Value "True"
    }
}

Write-Output "Loading user profile task finished"