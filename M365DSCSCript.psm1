<#
# .SYNOPSIS
#   Installs the Microsoft365DSC PowerShell module.
# .DESCRIPTION
#   Installs the Microsoft365DSC PowerShell module.
# .EXAMPLE
#   Install-Microsoft365DSC
#>

function Install-Microsoft365DSC
{

    # Test if the script is running as administrator
    # Test if the script is running in PowerShell 5
    if ($PSVersionTable.MajorVersion -ne 5 -or `
            -not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator') )
    {
        Write-Error 'This script requires PowerShell 5 and needs to be run as administrator'
        return
    }

    # Install the required modules
    Install-Module -Name Microsoft365DSC
    Update-M365DSCDependencies
    Uninstall-M365DSCOutdatedDependencies

}

<#
# .SYNOPSIS
#   Creates a new Azure AD application with full read permissions for Microsoft365DSC.
# .DESCRIPTION
#   Creates a new Azure AD application with full read permissions for Microsoft365DSC.
# .EXAMPLE
#   New-Microsoft365DSCApplicationWithFullReadPermission -ApplicationName "Microsoft365DSC" -CertificatePath "C:\Microsoft365DSC\Microsoft365DSC.pfx"
#>
function New-Microsoft365DSCApplicationWithFullReadPermission
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [System.String]
        $ApplicationName,

        [Parameter(Mandatory = $true)]
        [System.String]
        $CertificatePath,

        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]
        $Credential
    )

    $allResources = Get-M365DSCAllResources

    $allPermissions = Get-M365DSCCompiledPermissionList -ResourceNameList $allResources -PermissionType Application -AccessType Read

    Update-M365DSCAzureAdApplication -ApplicationName $ApplicationName `
        -Permissions $allPermissions `
        -Type Certificate `
        -CreateSelfSignedCertificate `
        -AdminConsent `
        -CertificatePath $CertificatePath `
        -MonthsValid 12 `
        -Credential $Credential

    Write-Host 'Application created with full read permissions'
    Write-Host "Application Name: $ApplicationName"
    Write-Host "Certificate Path: $CertificatePath"
    Write-Host 'Please export the certificate with the private key and store it in a computer certificate store.'

    Write-Host 'Please assign global reader role to the application.'
    
    Write-Host 'Please follow these instructions to configure the applications permissions: https://microsoft365dsc.com/user-guide/get-started/authentication-and-permissions/'
}

<#
# .SYNOPSIS
#   Creates the directory structure for the Microsoft365DSC script.
# .DESCRIPTION
#   Creates the directory structure for the Microsoft365DSC script.
# .PARAMETER BasePath
#   The base path where the directory structure should be created.
# .EXAMPLE
#   Confirm-Microsoft365DSCDirectoryStructure -BasePath "C:\Microsoft365DSC"
#>
function Confirm-Microsoft365DSCDirectoryStructure
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $BasePath
    )

    $directories = @(
        (Join-Path $BasePath 'Scripts'),
        (Join-Path $BasePath 'Exports'),
        (Join-Path $BasePath 'Reports')
    )

    foreach ($directory in $directories)
    {
        if (-not (Test-Path $directory))
        {
            New-Item -Path $directory -ItemType Directory -Force
        }
    }
}

Export-ModuleMember -Function Install-Microsoft365DSC, New-Microsoft365DSCApplicationWithFullReadPermission, Confirm-Microsoft365DSCDirectoryStructure
