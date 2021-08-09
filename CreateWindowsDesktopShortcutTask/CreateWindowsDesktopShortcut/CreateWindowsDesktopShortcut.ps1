Write-Output "Starting: Create desktop shortcut"

[String] $SourceFileLocation = Get-VstsInput -Name SourceFileLocation -Require
[String] $WorkingDirectory = Get-VstsInput -Name WorkingDirectory
[String] $ShortcutName = Get-VstsInput -Name ShortcutName -Require
[String] $IconLocation = Get-VstsInput -Name IconLocation -Default "%SystemRoot%\system32\SHELL32.dll, 46"
[String] $Description = Get-VstsInput -Name Description
[Boolean] $RunAsAdministrator = Get-VstsInput -Name RunAsAdministrator -Default $false



if ([String]::IsNullOrWhiteSpace($WorkingDirectory)) {
    $WorkingDirectory = Split-Path -Path $SourceFileLocation
}

$WScriptShell = New-Object -ComObject WScript.Shell
$PublicDesktop = $env:PUBLIC + "\Desktop"
$ShortcutPath = "$PublicDesktop\$ShortcutName.lnk"

Write-Output "Shortcut path: $ShortcutPath"

if(Test-Path $ShortcutPath) {
    Remove-Item $ShortcutPath -Force
    Write-Output "Removed $ShortcutPath"
}

$Shortcut = $WScriptShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $SourceFileLocation
$Shortcut.WorkingDirectory = $WorkingDirectory
$Shortcut.IconLocation = $IconLocation
$Shortcut.Description = $Description
$Shortcut.Save()

Write-Output "Shortcut created"

if ($RunAsAdministrator) {
  $bytes = [System.IO.File]::ReadAllBytes($ShortcutPath)
  $bytes[0x15] = $bytes[0x15] -bor
  [System.IO.File]::WriteAllBytes($ShortcutPath, $bytes)
}

Write-Output "Finishing: Create desktop shortcut"