Write-Output "Start installing util task"

[string]$WebSiteName = Get-VstsInput -Name WebSiteName -Require
[string]$FilePath = Get-VstsInput -Name FilePath -Require

Write-Output "Start unlocking $SectionNames on IIS"

$dotnetPath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319"
Push-Location $dotnetPath

.\installutil.exe /instanceName="$WebSiteName" "$PhysicalPath\$FilePath"

Pop-Location

Write-Output "Unlocking sections on IIS task finished"
              