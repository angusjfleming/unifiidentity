
<#
.SYNOPSIS
  Download MSI(s), compute checksum(s), update Chocolatey script checksums (replace-only),
  and update unifiidentity.nuspec version from MSI ProductVersion.

.DESCRIPTION
  - Updates ONLY `checksum` and `checksum64` in tools\chocolateyinstall.ps1.
  - Does NOT modify checksumType fields.
  - PRESERVES formatting (indent, spaces, original quotes, trailing comments).
  - DOES NOT append new lines if the keys are missing (replace-only).
  - Extracts MSI ProductVersion (prefers x64 MSI) and updates ..\unifiidentity.nuspec <version>
    using namespace-aware XPath, with regex fallback.

.PARAMETER ScriptPath
  Path to the Chocolatey install script (e.g., tools\chocolateyinstall.ps1).

.PARAMETER Url
  Download URL for x86 MSI (optional if Url64 is provided).

.PARAMETER Url64
  Download URL for x64 MSI (optional if Url is provided).

.PARAMETER Algorithm
  Hash algorithm for checksums: sha256 (default), sha1, sha384, sha512, md5.

.PARAMETER DownloadDir
  Optional directory for downloaded files; defaults to "<repo-root>\downloads".
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$ScriptPath,

  [Parameter(Mandatory = $false)]
  [string]$Url,

  [Parameter(Mandatory = $false)]
  [string]$Url64,

  [Parameter(Mandatory = $false)]
  [ValidateSet('sha1','sha256','sha384','sha512','md5')]
  [string]$Algorithm = 'sha256',

  [Parameter(Mandatory = $false)]
  [string]$DownloadDir
)

$ErrorActionPreference = 'Stop'

# --- Resolve paths ---
$scriptFullPath = Resolve-Path -Path $ScriptPath

# If this script is .\build\Update-ChocoChecksums.ps1, repo root is the parent of $PSScriptRoot
$repoRoot   = Split-Path -Parent $PSScriptRoot
$NuspecPath = Join-Path $repoRoot 'unifiidentity.nuspec'

if ([string]::IsNullOrWhiteSpace($DownloadDir)) {
  $DownloadDir = Join-Path $repoRoot 'downloads'
}
if (-not (Test-Path -Path $DownloadDir)) {
  New-Item -ItemType Directory -Path $DownloadDir | Out-Null
}

# --- Algorithm for Get-FileHash ---
$algoUpper = switch ($Algorithm.ToLowerInvariant()) {
  'sha1'   { 'SHA1' }
  'sha256' { 'SHA256' }
  'sha384' { 'SHA384' }
  'sha512' { 'SHA512' }
  'md5'    { 'MD5' }
  default  { throw "Unsupported algorithm: $Algorithm" }
}

function New-DownloadFileName {
  param([Parameter(Mandatory = $true)][string]$Label)
  $ts = Get-Date -Format 'yyyyMMddHHmmssfff'
  Join-Path $DownloadDir ("$Label-$ts.msi")
}

function Invoke-DownloadWithRetry {
  param(
    [Parameter(Mandatory = $true)][string]$SourceUrl,
    [Parameter(Mandatory = $true)][string]$DestinationPath,
    [int]$MaxAttempts = 5,
    [int]$DelaySeconds = 2
  )
  for ($i = 1; $i -le $MaxAttempts; $i++) {
    try {
      Write-Host "Downloading ($i/$MaxAttempts): $SourceUrl"
      Invoke-WebRequest -Uri $SourceUrl -OutFile $DestinationPath -UseBasicParsing
      if (Test-Path $DestinationPath -PathType Leaf) { return $DestinationPath }
      throw "File not found after download: $DestinationPath"
    } catch {
      if ($i -eq $MaxAttempts) { throw }
      Start-Sleep -Seconds $DelaySeconds
    }
  }
}

function Get-FileHashHex {
  param(
    [Parameter(Mandatory = $true)][string]$Path,
    [Parameter(Mandatory = $true)][string]$Algorithm
  )
  (Get-FileHash -Path $Path -Algorithm $Algorithm).Hash.ToUpperInvariant()
}

# Read MSI ProductVersion via COM
function Get-MsiVersion {
  param([Parameter(Mandatory = $true)][string]$MsiPath)
  if (-not (Test-Path $MsiPath)) { throw "MSI not found: $MsiPath" }

  $installer = $null; $db = $null; $view = $null; $record = $null
  try {
    $installer = New-Object -ComObject WindowsInstaller.Installer
    $db = $installer.OpenDatabase($MsiPath, 0)
    $view = $db.OpenView("SELECT `Value` FROM `Property` WHERE `Property`='ProductVersion'")
    $view.Execute()
    $record = $view.Fetch()
    if ($record) { return $record.StringData(1).Trim() }
    throw "ProductVersion not found in MSI Property table."
  } finally {
    foreach ($obj in @($record, $view, $db, $installer)) {
      if ($obj -ne $null) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) }
    }
  }
}

# --- Download & compute hashes / preferred MSI for version ---
$checksum      = $null
$checksum64    = $null
$msiForVersion = $null

if ($Url) {
  $dst = New-DownloadFileName -Label 'x86'
  $dl  = Invoke-DownloadWithRetry -SourceUrl $Url -DestinationPath $dst
  $checksum = Get-FileHashHex -Path $dl -Algorithm $algoUpper
  if (-not $Url64) { $msiForVersion = $dl }
}
if ($Url64) {
  $dst64 = New-DownloadFileName -Label 'x64'
  $dl64  = Invoke-DownloadWithRetry -SourceUrl $Url64 -DestinationPath $dst64
  $checksum64    = Get-FileHashHex -Path $dl64 -Algorithm $algoUpper
  $msiForVersion = $dl64
}

# --- Read the Chocolatey script ---
$content = Get-Content -Path $scriptFullPath -Raw

# Use here-strings for robust regex patterns (no escaping headaches)
$patternChecksum = @'
(?mi)^(?<indent>\s*)\$?checksum(?<sp>\s*)=(?<sp2>\s*)(?<quote>["'])(?<val>[^"']*)(?<quote2>["'])(?<trail>\s*#.*)?$
'@

$patternChecksum64 = @'
(?mi)^(?<indent>\s*)\$?checksum64(?<sp>\s*)=(?<sp2>\s*)(?<quote>["'])(?<val>[^"']*)(?<quote2>["'])(?<trail>\s*#.*)?$
'@

function Replace-LineValue {
  param(
    [Parameter(Mandatory = $true)][string]$InContent,
    [Parameter(Mandatory = $true)][string]$Pattern,
    [Parameter(Mandatory = $true)][string]$Name,
    [Parameter(Mandatory = $true)][string]$NewValue
  )

  $opts = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
          [System.Text.RegularExpressions.RegexOptions]::Multiline

  if ([System.Text.RegularExpressions.Regex]::IsMatch($InContent, $Pattern, $opts)) {
    $evaluator = [System.Text.RegularExpressions.MatchEvaluator]{
      param([System.Text.RegularExpressions.Match]$m)
      $q1 = $m.Groups['quote'].Value
      $q2 = $m.Groups['quote2'].Value
      if ([string]::IsNullOrEmpty($q1)) { $q1 = "'" }
      if ([string]::IsNullOrEmpty($q2)) { $q2 = "'" }
      ($m.Groups['indent'].Value + $Name +
       $m.Groups['sp'].Value     + '=' +
       $m.Groups['sp2'].Value    + $q1 + $NewValue + $q2 +
       $m.Groups['trail'].Value)
    }
    return [System.Text.RegularExpressions.Regex]::Replace($InContent, $Pattern, $evaluator, $opts)
  }

  # Replace-only: do NOT append if not found
  Write-Host "Line '$Name' not foundâ€”no changes made to $Name."
  return $InContent
}  # closing brace present

# Replace-only updates
if ($null -ne $checksum)   { $content = Replace-LineValue -InContent $content -Pattern $patternChecksum   -Name 'checksum'   -NewValue $checksum }
if ($null -ne $checksum64) { $content = Replace-LineValue -InContent $content -Pattern $patternChecksum64 -Name 'checksum64' -NewValue $checksum64 }

# Write back Chocolatey script
Set-Content -Path $scriptFullPath -Value $content -Encoding UTF8

# --- Update nuspec <version> from MSI ProductVersion ---
if ($msiForVersion -and (Test-Path $NuspecPath)) {
  $msiVersion = Get-MsiVersion -MsiPath $msiForVersion

  # Namespace-aware XML update (falls back to regex if needed)
  try {
    [xml]$doc = Get-Content -Path $NuspecPath -Raw
    $nsUri = $doc.DocumentElement.NamespaceURI
    $nsMgr = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    if ([string]::IsNullOrWhiteSpace($nsUri)) {
      # Default nuspec schema if DocumentElement has no namespace
      $nsMgr.AddNamespace('n', 'http://schemas.microsoft.com/packaging/2015/06/nuspec.xsd')
    } else {
      $nsMgr.AddNamespace('n', $nsUri)
    }
    $verNode = $doc.SelectSingleNode('/n:package/n:metadata/n:version', $nsMgr)
    if ($verNode -ne $null) {
      $verNode.InnerText = $msiVersion
      $doc.Save($NuspecPath)
    } else {
      throw "version node not found via XML namespace"
    }
  } catch {
    # Regex fallback: update first <version> under <metadata>
    $txt = Get-Content -Path $NuspecPath -Raw

    # Hardened pattern (attributes tolerated; singleline; ignorecase)
    $patternVersion = '(?is)(<metadata\b[^>]*>.*?<version\b[^>]*>)([^<]+)(</version>)'

    if ([System.Text.RegularExpressions.Regex]::IsMatch(
          $txt, $patternVersion,
          [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
          [System.Text.RegularExpressions.RegexOptions]::Singleline)) {

      $eval = [System.Text.RegularExpressions.MatchEvaluator]{
        param([System.Text.RegularExpressions.Match]$m)
        $m.Groups[1].Value + $msiVersion + $m.Groups[3].Value
      }

      $txt = [System.Text.RegularExpressions.Regex]::Replace(
               $txt, $patternVersion, $eval,
               [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor
               [System.Text.RegularExpressions.RegexOptions]::Singleline)

      Set-Content -Path $NuspecPath -Value $txt -Encoding UTF8
    } else {
      # No angle brackets to avoid parser confusion
      Write-Warning "Could not locate version element in nuspec; no update applied."
    }
  }
} else {
  Write-Host "Skipping nuspec update (no MSI or nuspec missing)."
}

