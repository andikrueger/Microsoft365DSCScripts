# Visit https://microsoft365dsc.com for more information

[CmdletBinding()]
param (
    $BasePath = 'C:\TenantConfig\',
    $ApplicationId = '',
    $CertificateThumbprint = 'C:\TenantConfig\Microsoft365DSC.pfx',
    $TenantId = ''
)

$date = Get-Date -Format 'yyyyMMdd-HHmm'

$directories = @(
    (Join-Path $BasePath 'Scripts'),
    (Join-Path $BasePath 'Exports'),
    (Join-Path $BasePath 'Reports')
)

foreach ($directory in $directories)
{
    if (-not (Test-Path $directory))
    {
        New-Item -Path $directory -ItemType Directory -Force | Out-Null
    }
}

$targetPath = "$($BasePath)\Exports\$($date)\"
$reportPath = "$($BasePath)\Reports\$($date)\"

$authenticationParameters = @{
    ApplicationId         = $ApplicationId
    TenantId              = $TenantId
    CertificateThumbprint = $CertificateThumbprint
}

$exportConfiguration = @{
    AAD                = @(
        'AADTenantDetails'
    )
    Exchange           = @(
        'EXOAcceptedDomain'
    )
    Intune             = @(
        'IntuneDeviceCategory'
    )
    Office365          = @(
        'O365AdminAuditLogConfig'
    )
    PowerPlatform      = @(
        'PPPowerAppsEnvironment'
    )
    SecurityCompliance = @(
        'SCSensitivityLabel'
    )
    SharePoint         = @(
        'ODSettings',
        'SPOTenantCdnEnabled'
    )
    Teams              = @(
        'TeamsAppPermissionPolicy'
    )
}

$lastExportTargetFolder = Get-ChildItem -Path "$($BasePath)\Exports" | Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1

foreach ($component in $exportConfiguration.Keys)
{
    try
    {
        $components = $exportConfiguration[$component]
        Export-M365DSCConfiguration @authenticationParameters -Path "$($targetPath)$($component)" -Components $Components | Out-Null

        $lastExportTargetFolderPath = "$($BasePath)\Exports\$($lastExportTargetFolder[0].Name)\"

        $lastExportFilePath = "$($lastExportTargetFolderPath)$($component)\M365TenantConfig.ps1"

        if (-not (Test-Path -Path $reportPath) )
        {
            New-Item -Type Directory -Path $reportPath | Out-Null
        }

        New-M365DSCDeltaReport -Source $lastExportFilePath -Destination "$($targetPath)$($component)\M365TenantConfig.ps1" -OutputPath "$($reportPath)$($component).html" -Type Html -DriftOnly $true

        if ((Get-ChildItem -Path $reportPath).Count -eq 0)
        {
            Remove-Item -Path $reportPath -Confirm:$false -Force | Out-Null
        }
    }
    catch
    {
        <#Do this if a terminating exception happens#>
    }
}


