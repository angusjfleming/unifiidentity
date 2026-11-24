
$ErrorActionPreference = 'Stop'
$toolsDir   = "$(Split-Path -parent $MyInvocation.MyCommand.Definition)"
$url        = 'https://fw-update.ui.com/api/v2/firmware-latest/data?filter=eq~~product~~uid-identity-standard-desktop-app-msi&filter=eq~~platform~~windows&filter=eq~~channel~~release'

$packageArgs = @{
  packageName   = $env:ChocolateyPackageName
  unzipLocation = $toolsDir
  fileType      = 'MSI'
  url           = $url

  softwareName  = 'unifiidentity*'

  checksum      = '718C71B7CDF1EC5271202BFE474D52B172AC93B02D66548FBD609CAC0FB8D424'
  checksumType  = 'sha256'

  silentArgs    = "/qn /norestart /l*v `"$($env:TEMP)\$($packageName).$($env:chocolateyPackageVersion).MsiInstall.log`""
  validExitCodes= @(0, 3010, 1641)
}

Install-ChocolateyPackage @packageArgs


