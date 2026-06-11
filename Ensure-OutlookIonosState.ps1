<#
.SYNOPSIS
  Re-applies the desired local Outlook Autodiscover registry state from one or more kit manifests.

.DESCRIPTION
  This script performs a local repair only. It does not fetch Autodiscover, does not ask for passwords,
  does not change Outlook profiles, and does not touch credentials. It checks that the managed XML files
  exist and re-applies the registry values described by the desired-state manifest.

  Version 0.0.17 uses reg.exe for registry writes and verifies the values after writing. This avoids
  ambiguous behavior where PowerShell registry-provider writes appeared to be planned but were not visible
  to subsequent checks on some systems.

.NOTES
  Version: 0.0.17
#>

[CmdletBinding(DefaultParameterSetName = 'ByEmail')]
param(
  [Parameter(ParameterSetName = 'ByEmail', Mandatory = $true)]
  [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
  [string]$Email,

  [Parameter(ParameterSetName = 'ByManifest', Mandatory = $true)]
  [string]$StateFile,

  [Parameter(ParameterSetName = 'All', Mandatory = $true)]
  [switch]$All,

  [string]$TargetDirectory = (Join-Path $env:LOCALAPPDATA 'Microsoft\Outlook\Autodiscover'),

  [switch]$StopOutlook,

  [switch]$Execute
)

$ErrorActionPreference = 'Stop'
$Version = '0.0.17'

function Get-DomainFromEmail {
  param([string]$Address)
  return ($Address -split '@', 2)[1].ToLowerInvariant()
}

function ConvertTo-SafeFilePrefix {
  param([string]$Domain)
  return ($Domain.ToLowerInvariant() -replace '[^a-z0-9.-]', '-')
}

function Convert-ToRegExePath {
  param([string]$Path)
  if ($Path -like 'HKCU:\*') {
    return ('HKCU\' + $Path.Substring(6))
  }
  if ($Path -like 'HKLM:\*') {
    return ('HKLM\' + $Path.Substring(6))
  }
  return ($Path -replace '^Registry::HKEY_CURRENT_USER', 'HKCU' -replace '^Registry::HKEY_LOCAL_MACHINE', 'HKLM')
}

function Invoke-RegExe {
  param([string[]]$Arguments)
  $output = & reg.exe @Arguments 2>&1
  if ($LASTEXITCODE -ne 0) {
    throw "reg.exe $($Arguments -join ' ') failed with exit code $LASTEXITCODE. Output: $output"
  }
}

function Write-Plan {
  param([string]$Message)
  if ($Execute) {
    Write-Host "DO:    $Message"
  }
  else {
    Write-Host "WOULD: $Message"
  }
}

function Get-RegistryValue {
  param(
    [string]$Path,
    [string]$Name
  )
  try {
    return (Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop).$Name
  }
  catch {
    return $null
  }
}

function Set-DWordValueIfNeeded {
  param(
    [string]$Path,
    [string]$Name,
    [int]$Value
  )

  $current = Get-RegistryValue -Path $Path -Name $Name
  if ($current -eq $Value) {
    Write-Host "OK:    $Path::$Name = $Value"
    return
  }

  Write-Plan "$Path::$Name = $Value (current: $current)"
  if ($Execute) {
    $regPath = Convert-ToRegExePath -Path $Path
    Invoke-RegExe -Arguments @('add', $regPath, '/v', $Name, '/t', 'REG_DWORD', '/d', ([string]$Value), '/f')

    $after = Get-RegistryValue -Path $Path -Name $Name
    if ($after -ne $Value) {
      throw "Post-check failed for $Path::$Name. Expected $Value, actual '$after'."
    }
    Write-Host "OK:    Verified $Path::$Name = $after"
  }
}

function Set-StringValueIfNeeded {
  param(
    [string]$Path,
    [string]$Name,
    [string]$Value
  )

  $current = Get-RegistryValue -Path $Path -Name $Name
  if ($current -eq $Value) {
    Write-Host "OK:    $Path::$Name = $Value"
    return
  }

  Write-Plan "$Path::$Name = $Value (current: $current)"
  if ($Execute) {
    $regPath = Convert-ToRegExePath -Path $Path
    Invoke-RegExe -Arguments @('add', $regPath, '/v', $Name, '/t', 'REG_SZ', '/d', $Value, '/f')

    $after = Get-RegistryValue -Path $Path -Name $Name
    if ($after -ne $Value) {
      throw "Post-check failed for $Path::$Name. Expected '$Value', actual '$after'."
    }
    Write-Host "OK:    Verified $Path::$Name = $after"
  }
}

function Remove-LastKnownGoodUrlIfPresent {
  param([string]$Path)

  $current = Get-RegistryValue -Path $Path -Name LastKnownGoodUrl
  if ($null -eq $current) {
    Write-Host "OK:    $Path::LastKnownGoodUrl is absent"
    return
  }

  Write-Plan "Remove $Path::LastKnownGoodUrl"
  if ($Execute) {
    $regPath = Convert-ToRegExePath -Path $Path
    & reg.exe delete $regPath /v LastKnownGoodUrl /f 2>$null | Out-Null

    $after = Get-RegistryValue -Path $Path -Name LastKnownGoodUrl
    if ($null -ne $after) {
      throw "Post-check failed for $Path::LastKnownGoodUrl. Value is still present."
    }
    Write-Host "OK:    Verified $Path::LastKnownGoodUrl is absent"
  }
}

function Read-StateFile {
  param([string]$Path)
  if (-not (Test-Path $Path)) {
    throw "State file not found: $Path"
  }
  return Get-Content $Path -Raw | ConvertFrom-Json
}

function New-InferredState {
  param(
    [string]$Address,
    [string]$Directory
  )

  $domain = Get-DomainFromEmail -Address $Address
  $prefix = ConvertTo-SafeFilePrefix -Domain $domain
  $fullXml = Join-Path $Directory ("$prefix-autodiscover-full.xml")
  $mapiXml = Join-Path $Directory ("$prefix-autodiscover-mapi-first.xml")
  $manifest = Join-Path $Directory ("$prefix-ionos-outlook-kit-state.json")

  return [pscustomobject]@{
    Version = $Version
    CreatedAt = (Get-Date).ToString('o')
    Email = $Address
    Domain = $domain
    ExchangeHost = 'exchange2019.ionos.eu'
    FullXmlPath = $fullXml
    MapiFirstXmlPath = $mapiXml
    ManifestFile = $manifest
    RegistryPaths = @(
      'HKCU:\Software\Microsoft\Office\16.0\Outlook\AutoDiscover',
      'HKCU:\Software\Policies\Microsoft\Office\16.0\Outlook\AutoDiscover'
    )
    ExpectedDwordValues = [pscustomobject]@{
      # IONOS-documented Microsoft 365 Autodiscover mitigation.
      ExcludeExplicitO365Endpoint = 1
      ExcludeHttpsRootDomain = 1

      # Additional values used by this kit for local XML pinning and stable repair.
      ExcludeHttpsAutoDiscoverDomain = 1
      ExcludeLastKnownGoodUrl = 1
      PreferLocalXML = 1
    }
    ExpectedStringValues = [pscustomobject]@{
      $domain = $mapiXml
    }
    ManagedFiles = @(
      $fullXml,
      $mapiXml,
      $manifest
    )
    Notes = @(
      'Inferred state created by Ensure-OutlookIonosState because no manifest existed yet.',
      'It intentionally does not store credentials.',
      'ExcludeExplicitO365Endpoint and ExcludeHttpsRootDomain reflect the IONOS-documented Microsoft 365 Autodiscover mitigation for hosted Exchange mailboxes.'
    )
  }
}

function Write-StateFileIfMissing {
  param(
    [object]$State,
    [string]$Path
  )

  if ((-not (Test-Path $Path)) -and $Execute) {
    $State | ConvertTo-Json -Depth 8 | Set-Content $Path -Encoding UTF8
    Write-Host "DO:    Wrote inferred desired-state manifest: $Path"
  }
}

function Invoke-EnsureState {
  param([object]$State)

  Write-Host "`n=== Ensuring $($State.Email) / $($State.Domain) ==="

  foreach ($file in @($State.FullXmlPath, $State.MapiFirstXmlPath)) {
    if (Test-Path $file) {
      Write-Host "OK:    File exists: $file"
    }
    else {
      throw "Managed file is missing: $file"
    }
  }

  foreach ($path in $State.RegistryPaths) {
    foreach ($prop in $State.ExpectedDwordValues.PSObject.Properties) {
      Set-DWordValueIfNeeded -Path $path -Name $prop.Name -Value ([int]$prop.Value)
    }
    foreach ($prop in $State.ExpectedStringValues.PSObject.Properties) {
      Set-StringValueIfNeeded -Path $path -Name $prop.Name -Value ([string]$prop.Value)
    }
    Remove-LastKnownGoodUrlIfPresent -Path $path
  }
}

if ($StopOutlook) {
  Get-Process OUTLOOK -ErrorAction SilentlyContinue | Stop-Process -Force
}

$stateFiles = @()
if ($PSCmdlet.ParameterSetName -eq 'ByEmail') {
  $domain = Get-DomainFromEmail -Address $Email
  $prefix = ConvertTo-SafeFilePrefix -Domain $domain
  $stateFiles += Join-Path $TargetDirectory ("$prefix-ionos-outlook-kit-state.json")
}
elseif ($PSCmdlet.ParameterSetName -eq 'ByManifest') {
  $stateFiles += $StateFile
}
else {
  $stateFiles = @(Get-ChildItem $TargetDirectory -File -Filter '*-ionos-outlook-kit-state.json' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName)
  if ($stateFiles.Count -eq 0) {
    throw "No state files found in $TargetDirectory. Run Update-IonosExchangeAutodiscover with -SetRegistryFromExistingXml first."
  }
}

foreach ($file in $stateFiles) {
  if (Test-Path $file) {
    $state = Read-StateFile -Path $file
    Invoke-EnsureState -State $state
  }
  elseif ($PSCmdlet.ParameterSetName -eq 'ByEmail') {
    Write-Host "WARN:  State file not found, inferring desired state from existing local XML paths: $file"
    $state = New-InferredState -Address $Email -Directory $TargetDirectory
    Invoke-EnsureState -State $state
    Write-StateFileIfMissing -State $state -Path $file
  }
  else {
    throw "State file not found: $file"
  }
}

if (-not $Execute) {
  Write-Host "`nDry-run only. Add -Execute to apply changes."
}
else {
  Write-Host "`nDone. Desired registry state has been verified after writing."
}
