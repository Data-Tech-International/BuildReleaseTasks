Write-Output "Start searching for the youngest cert"

[string]$HostHeader = Get-VstsInput -Name HostHeader -Require
[string]$FriendlyName = Get-VstsInput -Name FriendlyName -Require
[string]$Resultname = Get-VstsInput -Name ResultName

$youngestCert = $null
$certAge = [System.DateTime]::MinValue
$dateTimeMinValue=[System.DateTime]::MinValue
$certs = Get-ChildItem -Path Cert:\LocalMachine\My
$youngestCertThumbprint = $null

if ($HostHeader -ne $null -and $HostHeader -ne ""){
    $hostHeader=$HostHeader.Trim()

    foreach($singleCert in $certs)
    {
        $allExtensions =  $singleCert.Extensions
        foreach($singleExtension in $allExtensions)
        {
            if($singleExtension.Oid.FriendlyName -eq $FriendlyName)
            {
                if($singleExtension.Format(1).Contains("DNS Name=$hostHeader") -or $singleCert.Subject.Contains("CN=$hostHeader"))
                {
                    If($certAge -lt $singleCert.NotBefore -or $certAge -eq $dateTimeMinValue)
                    {
                        $youngestCert=$singleCert
                        $certAge = $singleCert.NotBefore
                    }
                }
            }
        }
    }
    Write-Output "HostHeader: $hostHeader"
    $youngestCertThumbprint=$youngestCert.Thumbprint
    Write-Output "Thumbprint: $youngestCertThumbprint"
}

Write-Host "##vso[task.setvariable variable=$ResultName]$youngestCertThumbprint"

    