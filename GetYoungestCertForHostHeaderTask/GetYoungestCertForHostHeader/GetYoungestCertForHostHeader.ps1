Write-Output "Start searching for the youngest cert"

[string]$HostHeader = Get-VstsInput -Name HostHeader -Require
[string]$FriendlyName = Get-VstsInput -Name FriendlyName -Require
[string]$Resultname = Get-VstsInput -Name ResultName

$youngestCert = $null
$certAgeInDays = 0
$now=[System.DateTime]::Now
$certs = Get-ChildItem -Path Cert:\LocalMachine\My

foreach($singleCert in $certs)
{
    $allExtensions =  $singleCert.Extensions
    foreach($singleExtension in $allExtensions)
    {
        if($singleExtension.Oid.FriendlyName -eq $FriendlyName)
        {
            if($singleExtension.Format(1).Contains("DNS Name=$HostHeader") -or $singleCert.Subject.Contains("CN=$HostHeader"))
            {
                If($($now-$singleCert.NotBefore).TotalDays -lt $certAgeInDays -or $certAgeInDays -eq 0)
                {
                    $youngestCert=$singleCert
                    $certAgeInDays = $($trenutak-$singleCert.NotBefore).TotalDays
                }

            }
        }
    }
}
$youngestCertThumbprint=$youngestCert.Thumbprint
Write-Host "##vso[task.setvariable variable=$ResultName]$youngestCertThumbprint"

    