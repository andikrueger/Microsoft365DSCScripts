# Set the various parameters for the script

$basepath = 'C:\Microsoft365DSC'

$ApplicationName = 'DemoM365DSC'
$CertificatePath = "$($basepath)\DemoM365DSC.cer"

# Import the helper module to setup the environment
Import-Module .\M365DSCSCript.psm1

# Install the module
Install-Microsoft365DSC

# Create the directory structure
Confirm-Microsoft365DSCDirectoryStructure -BasePath $basepath

# Create the application with full read permissions
$cred = Get-Credential
New-Microsoft365DSCApplicationWithFullReadPermission -ApplicationName $ApplicationName -CertificatePath $CertificatePath -Credential $cred


