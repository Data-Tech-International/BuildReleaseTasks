Write-Output "Start user creating"

[String]$ComputerName = Get-VstsInput -Name ComputerName -Require
[String]$Username = Get-VstsInput -Name Username -Require
[String]$Password = Get-VstsInput -Name Password -Require
[String]$LocalGroups = Get-VstsInput -Name LocalGroups

Write-Output "Adding local user $Username on $ComputerName and to local groups $LocalGroups"

$localUser = Get-WmiObject Win32_UserAccount -Filter "Domain='$ComputerName' and Name='$Username'"
if(!$localUser)
{
    # Create new local user 
    $ComputerSvc = [ADSI]"WinNT://$ComputerName,Computer"
    $LocalUserSvc = $ComputerSvc.Create("User", $Username)
    try {
        $LocalUserSvc.SetPassword($Password)
        $LocalUserSvc.SetInfo()
        $LocalUserSvc.Description = "Automatically created by pipeline task"
        $LocalUserSvc.SetInfo()
        $LocalUserSvc.UserFlags = 65536 # ADS_UF_DONT_EXPIRE_PASSWD
        $LocalUserSvc.SetInfo()
        Write-Output "  Local user $Username is created"
    } catch {
        $ErrorMessage = $_.Exception.Message -replace [System.Environment]::NewLine
        Write-Output "  Create new Local user $Username failed: $ErrorMessage"
    }

    remove-variable LocalUserSvc,ComputerSvc
}
else
{
    try {
        $LocalUserSvc = [adsi]"WinNT://$ComputerName/$Username,user"
        $LocalUserSvc.UserFlags = 65536 # ADS_UF_DONT_EXPIRE_PASSWD
        $LocalUserSvc.SetInfo()
        $LocalUserSvc.SetPassword($Password)
        $LocalUserSvc.SetInfo()
        Write-Output "  User $Username already exists, password updated"
    } catch {
        $ErrorMessage = $_.Exception.Message -replace [System.Environment]::NewLine
        Write-Output "  Update Local user $localUser password failed: $ErrorMessage"
    }

    remove-variable LocalUserSvc
}
#add user to IIS_IUSERS group
if ($LocalGroups -eq $null){
    $LocalGroups='';
}

$LocalGroupsArray = $LocalGroups.Split(",")
if ($LocalGroupsArray.Count -gt 0){
    try {
        $LocalUserSvc = [adsi]"WinNT://$ComputerName/$Username,user"
        $LocalGroupsArray | ForEach-Object {
            $groupName=$_.Trim()
            Write-Output "Adding $Username to group $groupName"
            $Group = [ADSI]"WinNT://$ComputerName/$groupName,Group"

            if ($Group -ne $null){
                $membersObj = @($Group.psbase.Invoke("Members")) 
                $members = ($membersObj | foreach {$_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)})

                If ($members -contains $Username) {
                    Write-Output "$Username exists in the group $groupName"
                } Else {
                    $Group.Add($LocalUserSvc.Path)
                }
            } Else {
                Write-Output "Group $groupName doesn't exists on Computer: $ComputerName"
            }
            
            remove-variable Group
        }
    }
    catch  {
        $ErrorMessage = $_.Exception.Message -replace [System.Environment]::NewLine
        Write-Output "Adding user $Username to local groups $LocalGroups failed: $ErrorMessage"
    }   
}

remove-variable LocalUserSvc
    