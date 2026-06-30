<#
.SYNOPSIS
  Restores Outlook/IONOS-related local state from a backup created by Backup-OutlookIonosState.

.NOTES
  Version: 0.0.22
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$BackupDirectory,

  [switch]$RestoreRegistry,

  [switch]$RestoreFiles,

  [switch]$RestoreOst,

  [switch]$StopOutlook,

  [switch]$Execute
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $BackupDirectory)) { throw "Backup directory not found: $BackupDirectory" }
if ($StopOutlook) { Get-Process OUTLOOK -ErrorAction SilentlyContinue | Stop-Process -Force }

function Invoke-Step {
  param([string]$Message, [scriptblock]$Action)
  if ($Execute) { Write-Host "DO:    $Message"; & $Action } else { Write-Host "WOULD: $Message" }
}

$regDir = Join-Path $BackupDirectory 'registry'
$fileDir = Join-Path $BackupDirectory 'files'
$outlookRoot = Join-Path $env:LOCALAPPDATA 'Microsoft\Outlook'
$autodiscoverDir = Join-Path $outlookRoot 'Autodiscover'

if ($RestoreRegistry) {
  Get-ChildItem $regDir -File -Filter '*.reg' -ErrorAction SilentlyContinue | ForEach-Object {
    $file = $_.FullName
    Invoke-Step "Import registry file $file" { reg import $file | Out-Null }
  }
}

if ($RestoreFiles) {
  $backupAutodiscover = Join-Path $fileDir 'Autodiscover'
  Invoke-Step "Replace $autodiscoverDir from backup" {
    Remove-Item $autodiscoverDir -Recurse -Force -ErrorAction SilentlyContinue
    if (Test-Path $backupAutodiscover) { Copy-Item $backupAutodiscover $autodiscoverDir -Recurse -Force }
  }

  $backupRootXml = Join-Path $fileDir 'OutlookRootAutodiscoverXml'
  Invoke-Step "Restore root *-Autodiscover.xml files" {
    Get-ChildItem $outlookRoot -File -Filter '*-Autodiscover.xml' -ErrorAction SilentlyContinue | Remove-Item -Force
    if (Test-Path $backupRootXml) { Get-ChildItem $backupRootXml -File | Copy-Item -Destination $outlookRoot -Force }
  }

  $backup16 = Join-Path $fileDir '16'
  $target16 = Join-Path $outlookRoot '16'
  Invoke-Step "Replace Outlook 16 cache directory from backup" {
    Remove-Item $target16 -Recurse -Force -ErrorAction SilentlyContinue
    if (Test-Path $backup16) { Copy-Item $backup16 $target16 -Recurse -Force }
  }
}

if ($RestoreOst) {
  $backupOst = Join-Path $fileDir 'OST'
  Invoke-Step "Restore OST files from backup" {
    if (Test-Path $backupOst) { Get-ChildItem $backupOst -File -Filter '*.ost' | Copy-Item -Destination $outlookRoot -Force }
  }
}

if (-not $Execute) { Write-Host "Dry-run only. Add -Execute to apply changes." }
