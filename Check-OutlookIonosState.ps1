<#
.SYNOPSIS
  Checks whether the local Outlook/IONOS desired state is still present.

.NOTES
  Version: 0.0.15
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
  [string]$Email,

  [string]$ProfileName = '',

  [string]$ExchangeHost = 'exchange2019.ionos.eu',

  [string]$TargetDirectory = (Join-Path $env:LOCALAPPDATA 'Microsoft\Outlook\Autodiscover'),

  [switch]$OnlineCheck,

  [switch]$SslNoRevoke,

  [switch]$Detailed
)

$ErrorActionPreference = 'Continue'

function Add-Result {
  param(
    [string]$Status,
    [string]$Area,
    [string]$Check,
    [string]$Detail
  )
  [pscustomobject]@{
    Status = $Status
    Area = $Area
    Check = $Check
    Detail = $Detail
  }
}

function Get-DomainFromEmail {
  param([string]$Address)
  return ($Address -split '@', 2)[1].ToLowerInvariant()
}

function ConvertTo-SafeFilePrefix {
  param([string]$Domain)
  return ($Domain.ToLowerInvariant() -replace '[^a-z0-9.-]', '-')
}

function Test-DwordValue {
  param(
    [string]$Path,
    [string]$Name,
    [int]$Expected
  )
  try {
    $actual = (Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop).$Name
    if ($actual -eq $Expected) {
      Add-Result 'OK' 'Registry' "$Path::$Name" "Value is $actual"
    }
    else {
      Add-Result 'FAIL' 'Registry' "$Path::$Name" "Expected $Expected, actual $actual"
    }
  }
  catch {
    Add-Result 'FAIL' 'Registry' "$Path::$Name" 'Missing'
  }
}

function Test-StringValue {
  param(
    [string]$Path,
    [string]$Name,
    [string]$Expected
  )
  try {
    $actual = (Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop).$Name
    if ($actual -eq $Expected) {
      Add-Result 'OK' 'Registry' "$Path::$Name" "Expected path is set"
    }
    else {
      Add-Result 'FAIL' 'Registry' "$Path::$Name" "Expected '$Expected', actual '$actual'"
    }
  }
  catch {
    Add-Result 'FAIL' 'Registry' "$Path::$Name" 'Missing'
  }
}

$results = @()
$domain = Get-DomainFromEmail -Address $Email
$prefix = ConvertTo-SafeFilePrefix -Domain $domain
$stateFile = Join-Path $TargetDirectory ("$prefix-ionos-outlook-kit-state.json")
$fullXml = Join-Path $TargetDirectory ("$prefix-autodiscover-full.xml")
$mapiXml = Join-Path $TargetDirectory ("$prefix-autodiscover-mapi-first.xml")

$results += Add-Result 'INFO' 'Input' 'Email' $Email
$results += Add-Result 'INFO' 'Input' 'Domain' $domain
$results += Add-Result 'INFO' 'Input' 'ExchangeHost' $ExchangeHost

$state = $null
if (Test-Path $stateFile) {
  $results += Add-Result 'OK' 'Manifest' 'State file' $stateFile
  try {
    $state = Get-Content $stateFile -Raw | ConvertFrom-Json
    if ($state.MapiFirstXmlPath) { $mapiXml = $state.MapiFirstXmlPath }
    if ($state.FullXmlPath) { $fullXml = $state.FullXmlPath }
    if ($state.ExchangeHost) { $ExchangeHost = $state.ExchangeHost }
  }
  catch {
    $results += Add-Result 'FAIL' 'Manifest' 'State file parse' $_.Exception.Message
  }
}
else {
  $results += Add-Result 'WARN' 'Manifest' 'State file' "Missing: $stateFile"
}

if (Test-Path $fullXml) { $results += Add-Result 'OK' 'Files' 'Full XML exists' $fullXml } else { $results += Add-Result 'FAIL' 'Files' 'Full XML exists' "Missing: $fullXml" }
if (Test-Path $mapiXml) { $results += Add-Result 'OK' 'Files' 'MAPI/HTTP-first XML exists' $mapiXml } else { $results += Add-Result 'FAIL' 'Files' 'MAPI/HTTP-first XML exists' "Missing: $mapiXml" }

if (Test-Path $mapiXml) {
  [xml]$doc = Get-Content $mapiXml -Raw
  $nsUri = 'http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a'
  $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
  $ns.AddNamespace('out', $nsUri)
  $firstType = $doc.SelectSingleNode('//out:Account/out:Protocol[1]/out:Type', $ns)
  if ($firstType -and $firstType.InnerText -eq 'mapiHttp') { $results += Add-Result 'OK' 'XML' 'First protocol' 'mapiHttp' } else { $results += Add-Result 'FAIL' 'XML' 'First protocol' "Expected mapiHttp, actual '$($firstType.InnerText)'" }
  foreach ($pattern in @('outlook.office365.com','outlook.live.com','login.microsoftonline.com')) {
    $content = Get-Content $mapiXml -Raw
    if ($content -match [regex]::Escape($pattern)) { $results += Add-Result 'FAIL' 'XML' "No $pattern" 'Unexpected Microsoft online host found' } else { $results += Add-Result 'OK' 'XML' "No $pattern" 'Not found' }
  }
  foreach ($pattern in @('/mapi/emsmdb/','/mapi/nspi/','<Type>EXPR</Type>','<Type>EXCH</Type>')) {
    $content = Get-Content $mapiXml -Raw
    if ($content -match [regex]::Escape($pattern)) { $results += Add-Result 'OK' 'XML' $pattern 'Present' } else { $results += Add-Result 'FAIL' 'XML' $pattern 'Missing' }
  }
}

$expectedMapiXml = if ($state -and $state.MapiFirstXmlPath) { $state.MapiFirstXmlPath } else { $mapiXml }
$regPaths = @(
  'HKCU:\Software\Microsoft\Office\16.0\Outlook\AutoDiscover',
  'HKCU:\Software\Policies\Microsoft\Office\16.0\Outlook\AutoDiscover'
)
foreach ($path in $regPaths) {
  if (Test-Path $path) { $results += Add-Result 'OK' 'Registry' $path 'Key exists' } else { $results += Add-Result 'FAIL' 'Registry' $path 'Key missing' }
  foreach ($name in @('ExcludeExplicitO365Endpoint','ExcludeHttpsRootDomain','ExcludeHttpsAutoDiscoverDomain','ExcludeLastKnownGoodUrl','PreferLocalXML')) {
    $results += Test-DwordValue -Path $path -Name $name -Expected 1
  }
  $results += Test-StringValue -Path $path -Name $domain -Expected $expectedMapiXml
  try {
    $null = Get-ItemProperty -Path $path -Name LastKnownGoodUrl -ErrorAction Stop
    $results += Add-Result 'FAIL' 'Registry' "$path::LastKnownGoodUrl" 'Value should be absent'
  }
  catch {
    $results += Add-Result 'OK' 'Registry' "$path::LastKnownGoodUrl" 'Value absent'
  }
}

$profileRoot = 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles'
if ($ProfileName) {
  $profilePath = Join-Path $profileRoot $ProfileName
  if (Test-Path $profilePath) { $results += Add-Result 'OK' 'Profiles' "Profile $ProfileName" 'Present' } else { $results += Add-Result 'FAIL' 'Profiles' "Profile $ProfileName" 'Missing' }
}
$adTests = @(Get-ChildItem $profileRoot -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -like 'ADTest*' })
if ($adTests.Count -eq 0) { $results += Add-Result 'OK' 'Profiles' 'ADTest profiles' 'None found' } else { $results += Add-Result 'WARN' 'Profiles' 'ADTest profiles' (($adTests | Select-Object -ExpandProperty PSChildName) -join ', ') }
$profiles = @(Get-ChildItem $profileRoot -ErrorAction SilentlyContinue | Select-Object -ExpandProperty PSChildName)
$results += Add-Result 'INFO' 'Profiles' 'Existing Outlook profiles' ($profiles -join ', ')

$cred = cmdkey /list | Select-String -Pattern 'exchange2019|engelmann|MicrosoftOffice16|autodiscover|ionos' -SimpleMatch
if ($cred) { $results += Add-Result 'INFO' 'Credentials' 'Credential Manager hints' (($cred | ForEach-Object { $_.Line.Trim() }) -join '; ') } else { $results += Add-Result 'INFO' 'Credentials' 'Credential Manager hints' 'No matching cmdkey entries found' }

$outlookRoot = Join-Path $env:LOCALAPPDATA 'Microsoft\Outlook'
$rootAutodiscover = @(Get-ChildItem $outlookRoot -File -Filter '*-Autodiscover.xml' -ErrorAction SilentlyContinue)
if ($rootAutodiscover.Count -eq 0) { $results += Add-Result 'OK' 'Files' 'Root Autodiscover cache XML files' 'None found' } else { $results += Add-Result 'WARN' 'Files' 'Root Autodiscover cache XML files' (($rootAutodiscover | Select-Object -ExpandProperty Name) -join ', ') }
$ostFiles = @(Get-ChildItem $outlookRoot -File -Filter '*.ost' -ErrorAction SilentlyContinue)
$results += Add-Result 'INFO' 'Files' 'OST files' (($ostFiles | ForEach-Object { "$($_.Name) ($([math]::Round($_.Length / 1MB, 1)) MB)" }) -join ', ')

if ($OnlineCheck) {
  $mapiUrl = "https://$ExchangeHost/mapi/emsmdb/"
  $rpcUrl = "https://$ExchangeHost/rpc/rpcproxy.dll"
  foreach ($item in @(@('MAPI endpoint',$mapiUrl), @('RPC/HTTP endpoint',$rpcUrl))) {
    $name = $item[0]
    $url = $item[1]
    $curlArgs = @('-sS', '-I', '--http1.1')
    if ($SslNoRevoke) { $curlArgs += '--ssl-no-revoke' }
    $curlArgs += $url
    $output = & curl.exe @curlArgs 2>&1
    $statusLine = ($output | Select-String -Pattern '^HTTP/' | Select-Object -First 1).Line
    if ($statusLine -match '401') {
      $results += Add-Result 'OK' 'Network' $name $statusLine
    }
    elseif ($statusLine) {
      $level = if ($name -like 'RPC*') { 'INFO' } else { 'WARN' }
      $results += Add-Result $level 'Network' $name $statusLine
    }
    else {
      $results += Add-Result 'WARN' 'Network' $name ($output -join ' ')
    }
  }
}

$order = @{ 'FAIL' = 0; 'WARN' = 1; 'OK' = 2; 'INFO' = 3 }
$sorted = $results | Sort-Object @{ Expression = { $order[$_.Status] } }, Area, Check
if ($Detailed) {
  $sorted | Format-List
}
else {
  $sorted | Format-Table -AutoSize -Wrap
}

$failCount = @($results | Where-Object Status -eq 'FAIL').Count
$warnCount = @($results | Where-Object Status -eq 'WARN').Count
Write-Host "`nSummary: $failCount FAIL, $warnCount WARN, $($results.Count) total checks"
exit $(if ($failCount -gt 0) { 2 } elseif ($warnCount -gt 0) { 1 } else { 0 })
