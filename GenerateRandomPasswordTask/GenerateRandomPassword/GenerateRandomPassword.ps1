Write-Output "Start password genereting"

[int]$MinimumPasswordLength = Get-VstsInput -Name MinimumPasswordLength -Require
[int]$MaximumPasswordLength = Get-VstsInput -Name MaximumPasswordLength -Require
[int]$NumberOfNonAlphanumericCharacters = Get-VstsInput -Name NumberOfNonAlphanumericCharacters -Require
[string]$ResultName = Get-VstsInput -Name ResultName -Require

Add-Type -AssemblyName 'System.Web'
$length = Get-Random -Minimum $MinimumPasswordLength -Maximum $MaximumPasswordLength
$password = [System.Web.Security.Membership]::GeneratePassword($length,$NumberOfNonAlphanumericCharacters)

Write-Output "##vso[task.setvariable variable=$ResultName]$password"    




    