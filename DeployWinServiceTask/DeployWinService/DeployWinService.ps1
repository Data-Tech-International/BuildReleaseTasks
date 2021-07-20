Write-Output "Start deploying windows service task"

[string]$ServiceName = Get-VstsInput -Name ServiceName -Require
[string]$DisplayName = Get-VstsInput -Name DisplayName -Require
[string]$SourcePath = Get-VstsInput -Name SourcePath -Require
[string]$RootPath = Get-VstsInput -Name RootPath -Require
[string]$ExcludeList = Get-VstsInput -Name ExcludeList -Require
[string]$Executable = Get-VstsInput -Name Executable -Require
[string]$StartType = Get-VstsInput -Name StartType -Require

Write-Output "Start deploying $ServiceName windows service"

try
{
#1. Stop process
    Write-Output "Start remote Stop-Process"
    Get-Process $ServiceName -ErrorAction SilentlyContinue | ForEach-object { $_ | Stop-Process -Force}
    Write-Output "Finished remote Stop-Process"
#2. Stop and uninstall service
    $svc=Get-Service $ServiceName -ErrorAction SilentlyContinue
    if ($svc)
    {
        if ($svc.Status -eq 'Running') 
        {
            Write-Host "Stopping windows service $($ServiceName)"
            Stop-Service $ServiceName
            Write-Host "Windows Service $($ServiceName) is stopped"
        }
            
        $CmdLine = '{0}\Microsoft.NET\Framework64\v4.0.30319\installutil.exe /u /instanceName="{1}" "{2}"' -f $env:SystemRoot,$ServiceName,$Executable
        Invoke-Expression $CmdLine	
        Write-Host "Windows Service $($ServiceName) is uninstalled."
    }

#3. Clean destination path    
    Write-Output "Destination path: $($RootPath)"
    if (Test-Path $RootPath) 
    {
        Write-Output "Destination path exists"
        if (($ExcludeList -ne $null) -and ($ExcludeList -ne "")){
            Remove-Item -recurse "$RootPath\*" -exclude $ExcludeList
        }
        Write-Output "Finished deleting"
    }
    else
    {
        Write-Output "Destination path does not exist"
        mkdir $RootPath
        Write-Output "Destination path: $($RootPath) created"
    }
#4. Copy content
	Write-Output "Start coping"
    Copy-Item "$SourcePath\*" -Destination $RootPath -Verbose -Recurse -Force
    Write-Output " Finished coping"
#5. Install service
    $CmdLine = $env:SystemRoot + '\Microsoft.NET\Framework64\v4.0.30319\installutil.exe /instanceName="{0}" /displayName="{1}" /startType="{3}" "{2}"' -f $ServiceName,$DisplayName,$executable,$StartType
    Write-Output "Invoking: $CmdLine"
    Invoke-Expression $CmdLine	
    Write-Output "Invoked: $CmdLine"
#6. Start service
    $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if($svc -eq $null)
    {
        Write-Error "Service $($ServiceName) not found!"
    }
    else
    {
        Write-Output "Found service $($svc.Name)"
        $svc | Format-List

        if ($svc.Status -ne 'Running' -and $svc.StartType -eq 'Automatic')
        {
            $svc.Start();
            Start-Sleep -Seconds 5

            $svc.WaitForStatus('Running','00:02:00')
                    
            if($svc.Status -ne 'Running')
            {
                Write-Error "`nService remained in state $($svc.Status)"
            }

            Write-Output "Started service $svc.Name"
        }
    }
}
catch
{
    Write-Error $_
    $LASTEXITCODE = 1
    exit $LASTEXITCODE
}
Write-Output "Deploying windows service finished"
              