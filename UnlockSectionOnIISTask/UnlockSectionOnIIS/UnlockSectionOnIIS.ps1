Write-Output "Start unlocking sections on IIS task"

[string]$SectionNames = Get-VstsInput -Name SectionNames -Require
[string]$SiteName = Get-VstsInput -Name SiteName -Require

Write-Output "Start unlocking $SectionNames on IIS"

Import-Module WebAdministration; 
              
if ($SectionNames -eq $null){
    $SectionNames='';
}

$SectionNamesArray = $SectionNames.Split(",")
if ($SectionNamesArray.Count -gt 0){
    $SectionNamesArray | ForEach-Object {
        $sectionName=$_.Trim()
        Set-WebConfiguration $sectionName -Location IIS:\sites\$SiteName -metadata overrideMode -value Allow
        Write-Output "  Removed lock from $sectionName PSPath IIS:\sites\$SiteName"
    }
}

Write-Output "Unlocking sections on IIS task finished"
              