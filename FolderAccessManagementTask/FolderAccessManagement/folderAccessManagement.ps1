Write-Output "Start adding access rights"

[String]$Username = Get-VstsInput -Name Username -Require
[String]$FolderName = Get-VstsInput -Name FolderName -Require
[String]$AccessRights = Get-VstsInput -Name AccessRights -Require

Write-Output "Adding user $Username access rights $AccessRights on folder $FolderName"

if ($AccessRights -eq $null){
    $AccessRights='';
}

$AccessRightsArray = $AccessRights.Split(",")
if ($AccessRightsArray.Count -gt 0){
    $Acl = Get-Acl $FolderName

    $AccessRightsArray | ForEach-Object {
        $accessRight=$_.Trim()
        $Ar = New-Object  system.security.accesscontrol.filesystemaccessrule($Username,$accessRight,"ContainerInherit,ObjectInherit","None","Allow")
        $Acl.SetAccessRule($Ar)
        Set-Acl $FolderName $Acl
    }

    remove-variable accessRight
}

remove-variable AccessRightsArray
Write-Output "Added user $Username access rights $AccessRights on folder $FolderName"