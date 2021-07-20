Write-Output "Start preparing bindings for site"

[string]$Bindings = Get-VstsInput -Name Bindings -Require
[string]$Resultname = Get-VstsInput -Name ResultName


$bindingsHashtable = $Bindings | ConvertFrom-Json
$result=@()
$certs = Get-ChildItem -Path Cert:\LocalMachine\My
foreach($bindingHashtable in $bindingsHashtable)
{
    If ($bindingHashtable.protocol -eq 'https'){
        $hostHeader=$bindingHashtable.hostHeader.Trim()
        $youngestCert = $null
        $certAge = [System.DateTime]::MinValue
        $dateTimeMinValue=[System.DateTime]::MinValue

        foreach($singleCert in $certs)
        {
            $allExtensions =  $singleCert.Extensions
            foreach($singleExtension in $allExtensions)
            {
                if($singleExtension.Oid.FriendlyName -eq "Subject Alternative Name")
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

        $result+=@{ipAddress="*";protocol="https";hostname="$hostHeader"; port="$($bindingHashtable.port)"; sslThumbprint="$($youngestCert.Thumbprint)"; sniFlag=$True;}

        Write-Output "HostHeader: $hostHeader"
        Write-Output "Thumbprint: $($youngestCert.Thumbprint)"
    } else {
            $result+=@{ipAddress="*";protocol="http";hostname="$hostHeader"; port="$($bindingHashtable.port)"; sslThumbprint=""; sniFlag=$False;}
    }
}

$resultJson = $result | ConvertTo-Json -Compress

Write-Host "##vso[task.setvariable variable=$ResultName]$resultJson"    