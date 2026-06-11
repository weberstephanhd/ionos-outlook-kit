<#
.SYNOPSIS
  Backs up the Outlook/IONOS-related local state before changes.

.DESCRIPTION
  Exports relevant Outlook registry keys and copies managed Autodiscover files. OST files are skipped by
  default because they can be large and are Exchange cache files. Use -OstHandling Move or Copy deliberately.

.NOTES
  Version: 0.0.17
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
  [string]$Email,

  [string]$BackupRoot = (Join-Path $env:USERPROFILE 'Desktop\outlook-ionos-backups'),

  [string]$TargetDirectory = (Join-Path $env:LOCALAPPDATA 'Microsoft\Outlook\Autodiscover'),

  [ValidateSet('Skip','Copy','Move')]
  [string]$OstHandling = 'Skip',

  [switch]$StopOutlook
)

$ErrorActionPreference = 'Stop'
$Version = '0.0.17'

function Get-DomainFromEmail { param([string]$Address) return ($Address -split '@', 2)[1].ToLowerInvariant() }
function ConvertTo-SafeFilePrefix { param([string]$Domain) return ($Domain.ToLowerInvariant() -replace '[^a-z0-9.-]', '-') }
function Export-KeyIfExists {
  param([string]$RegPath, [string]$OutputFile)
  $null = reg query $RegPath 2>$null
  if ($LASTEXITCODE -eq 0) { reg export $RegPath $OutputFile /y | Out-Null }
}

if ($StopOutlook) { Get-Process OUTLOOK -ErrorAction SilentlyContinue | Stop-Process -Force }

$domain = Get-DomainFromEmail -Address $Email
$prefix = ConvertTo-SafeFilePrefix -Domain $domain
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backupDir = Join-Path $BackupRoot ("$prefix-$stamp")
$regDir = Join-Path $backupDir 'registry'
$fileDir = Join-Path $backupDir 'files'
New-Item -ItemType Directory -Path $regDir,$fileDir -Force | Out-Null

Export-KeyIfExists 'HKCU\Software\Microsoft\Office\16.0\Outlook\AutoDiscover' (Join-Path $regDir 'hkcu-office-autodiscover.reg')
Export-KeyIfExists 'HKCU\Software\Policies\Microsoft\Office\16.0\Outlook\AutoDiscover' (Join-Path $regDir 'hkcu-policies-autodiscover.reg')
Export-KeyIfExists 'HKCU\Software\Microsoft\Office\16.0\Outlook\Profiles' (Join-Path $regDir 'hkcu-office-profiles.reg')
Export-KeyIfExists 'HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles' (Join-Path $regDir 'hkcu-messaging-profiles.reg')

if (Test-Path $TargetDirectory) {
  Copy-Item $TargetDirectory (Join-Path $fileDir 'Autodiscover') -Recurse -Force
}

$outlookRoot = Join-Path $env:LOCALAPPDATA 'Microsoft\Outlook'
$rootXmlDir = Join-Path $fileDir 'OutlookRootAutodiscoverXml'
New-Item -ItemType Directory -Path $rootXmlDir -Force | Out-Null
Get-ChildItem $outlookRoot -File -Filter '*-Autodiscover.xml' -ErrorAction SilentlyContinue |
  Copy-Item -Destination $rootXmlDir -Force

$sixteenDir = Join-Path $outlookRoot '16'
if (Test-Path $sixteenDir) {
  Copy-Item $sixteenDir (Join-Path $fileDir '16') -Recurse -Force
}

$ostDir = Join-Path $fileDir 'OST'
New-Item -ItemType Directory -Path $ostDir -Force | Out-Null
$ostFiles = @(Get-ChildItem $outlookRoot -File -Filter '*.ost' -ErrorAction SilentlyContinue)
foreach ($ost in $ostFiles) {
  if ($OstHandling -eq 'Copy') { Copy-Item $ost.FullName $ostDir -Force }
  elseif ($OstHandling -eq 'Move') { Move-Item $ost.FullName $ostDir -Force }
}

$metadata = [ordered]@{
  Version = $Version
  CreatedAt = (Get-Date).ToString('o')
  Email = $Email
  Domain = $domain
  BackupDirectory = $backupDir
  OstHandling = $OstHandling
}
$metadata | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $backupDir 'backup-metadata.json') -Encoding UTF8
Write-Host "Backup written: $backupDir"
