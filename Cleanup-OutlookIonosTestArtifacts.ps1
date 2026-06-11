<#
.SYNOPSIS
  Removes known Outlook/IONOS test artifacts without touching production profiles by default.

.DESCRIPTION
  Defaults to dry-run. Use -Execute to apply. On multi-profile machines, prefer -RemoveAdTestProfilesOnly.

.NOTES
  Version: 0.0.17
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
  [string]$Email,

  [string]$KeepProfile = '',

  [switch]$RemoveAdTestProfilesOnly,

  [switch]$RemoveRootAutodiscoverXml,

  [switch]$RemoveTmpFiles,

  [switch]$RemoveRedirectServers,

  [switch]$StopOutlook,

  [switch]$Execute
)

$ErrorActionPreference = 'Stop'

function Invoke-Step { param([string]$Message, [scriptblock]$Action) if ($Execute) { Write-Host "DO:    $Message"; & $Action } else { Write-Host "WOULD: $Message" } }
function Get-DomainFromEmail { param([string]$Address) return ($Address -split '@', 2)[1].ToLowerInvariant() }

if ($StopOutlook) { Get-Process OUTLOOK -ErrorAction SilentlyContinue | Stop-Process -Force }

$profileRoots = @(
  'HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles',
  'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles'
)

foreach ($root in $profileRoots) {
  if (-not (Test-Path $root)) { continue }
  Get-ChildItem $root | Where-Object {
    if ($RemoveAdTestProfilesOnly) { $_.PSChildName -like 'ADTest*' }
    else { $_.PSChildName -like 'ADTest*' -or $_.PSChildName -like 'IONOS-Exchange-AutoXML*' -or $_.PSChildName -like 'IONOS-Exchange-CredSeed*' -or $_.PSChildName -like 'IONOS-Exchange-Basic*' }
  } | Where-Object { -not $KeepProfile -or $_.PSChildName -ne $KeepProfile } | ForEach-Object {
    $name = $_.PSChildName
    Invoke-Step "Remove test Outlook profile $name from $root" { Remove-Item $_.PSPath -Recurse -Force }
  }
}

$outlookRoot = Join-Path $env:LOCALAPPDATA 'Microsoft\Outlook'
if ($RemoveRootAutodiscoverXml) {
  Get-ChildItem $outlookRoot -File -Filter '*-Autodiscover.xml' -ErrorAction SilentlyContinue | ForEach-Object {
    $file = $_.FullName
    Invoke-Step "Remove root Autodiscover cache file $file" { Remove-Item $file -Force }
  }
}
if ($RemoveTmpFiles) {
  Get-ChildItem $outlookRoot -File -Filter '*.tmp' -ErrorAction SilentlyContinue | ForEach-Object {
    $file = $_.FullName
    Invoke-Step "Remove temporary Outlook file $file" { Remove-Item $file -Force }
  }
}

if ($RemoveRedirectServers) {
  foreach ($key in @('HKCU:\Software\Microsoft\Office\16.0\Outlook\AutoDiscover\RedirectServers','HKCU:\Software\Policies\Microsoft\Office\16.0\Outlook\AutoDiscover\RedirectServers')) {
    if (Test-Path $key) {
      foreach ($name in @('autodiscover.1and1.info','exchange2019.ionos.eu')) {
        Invoke-Step "Remove $key::$name" { Remove-ItemProperty -Path $key -Name $name -Force -ErrorAction SilentlyContinue }
      }
    }
  }
}

if (-not $Execute) { Write-Host "Dry-run only. Add -Execute to apply changes." }
