Write-Output "Start event log creating"

[String]$SourceName = Get-VstsInput -Name SourceName -Require
[String]$LogName = Get-VstsInput -Name LogName -Require
[Int]$MaximumLogSize = Get-VstsInput -Name MaximumLogSize -Require
[String]$AutoBackupLogFilesName = Get-VstsInput -Name AutoBackupLogFilesName

if($SourceName -ne ""){
    if((Get-EventLog -List).Log -notcontains $LogName)
    {
        New-Eventlog -LogName $LogName -Source $SourceName -ErrorAction SilentlyContinue
        Write-Output "  Eventlog $LogName created and Source $SourceName registered"
    }
    else
    {
        New-Eventlog -LogName $LogName -Source $SourceName -ErrorAction SilentlyContinue
        Write-Output "  Source $SourceName registerd for Eventlog $LogName created"
    }
    if((Get-EventLog -List | Where-Object {$_.log -eq $LogName}).MaximumKilobytes -ne $MaximumLogSize*1024)
        {Limit-EventLog -LogName $LogName -MaximumSize "$($MaximumLogSize)MB";
        Write-Output "  Eventlog $LogName MaximumSize set to $MaximumLogSize MB"}
    $regkey = "HKLM:\SYSTEM\CurrentControlSet\services\eventlog\$LogName"
    if((Get-ItemProperty $regkey).PSObject.Properties.Name -contains $AutoBackupLogFilesName)
        {
            if((Get-ItemProperty $regkey).$AutoBackupLogFilesName -ne 1)
            {
                Set-ItemProperty $regkey -Name AutoBackupLogFilesName -value 1;
                Write-Output "  Registry value $regkey\$AutoBackupLogFilesName updated to value 1"
            }
        }
    else
        {New-ItemProperty $regkey -Name $AutoBackupLogFilesName -value 1;
            Write-Output "  Registry value $regkey\$AutoBackupLogFilesName Created adn set to value 1"}
    remove-variable regkey
}
    