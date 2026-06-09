<#
.SYNOPSIS
  Creates a local Outlook Autodiscover XML for IONOS Hosted Exchange 2019 and optionally registers it.

.DESCRIPTION
  The script fetches the live IONOS Autodiscover response for the mailbox, writes the full response,
  creates a MAPI/HTTP-first local Autodiscover XML, and can set the Outlook Autodiscover registry keys.
  It also writes a desired-state manifest used by Check/Ensure/Restore helper scripts.

.NOTES
  Version: 0.0.15
  The script does not store the mailbox password. The password is only used for the live Autodiscover request.
  Use -SslNoRevoke only as a temporary diagnostic workaround for Windows Schannel revocation failures.
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [ValidatePattern('^[^@\s]+@[^@\s]+\.[^@\s]+$')]
  [string]$Email,

  [string]$ExchangeHost = 'exchange2019.ionos.eu',

  [string]$AutodiscoverUrl = '',

  [string]$TargetDirectory = (Join-Path $env:LOCALAPPDATA 'Microsoft\Outlook\Autodiscover'),

  [switch]$SetRegistry,

  [switch]$SetRegistryFromExistingXml,

  [switch]$RegisterRedirectServers,

  [switch]$BackupBeforeChanges,

  [switch]$StopOutlook,

  [switch]$SslNoRevoke
)

$ErrorActionPreference = 'Stop'
$Version = '0.0.15'

function Write-Info {
  param([string]$Message)
  Write-Host $Message
}

function Get-DomainFromEmail {
  param([string]$Address)
  return ($Address -split '@', 2)[1].ToLowerInvariant()
}

function ConvertTo-SafeFilePrefix {
  param([string]$Domain)
  return ($Domain.ToLowerInvariant() -replace '[^a-z0-9.-]', '-')
}

function Stop-OutlookIfRequested {
  if ($StopOutlook) {
    Get-Process OUTLOOK -ErrorAction SilentlyContinue | Stop-Process -Force
  }
}

function Export-KeyIfExists {
  param(
    [string]$RegPath,
    [string]$OutputFile
  )

  $null = reg query $RegPath 2>$null
  if ($LASTEXITCODE -eq 0) {
    reg export $RegPath $OutputFile /y | Out-Null
  }
}

function Invoke-StateBackup {
  param(
    [string]$Domain,
    [string]$TargetDirectory
  )

  $backupRoot = Join-Path $env:USERPROFILE 'Desktop\outlook-ionos-backups'
  $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
  $backupDir = Join-Path $backupRoot ("$Domain-$stamp")
  New-Item -ItemType Directory -Path $backupDir -Force | Out-Null

  $regDir = Join-Path $backupDir 'registry'
  $fileDir = Join-Path $backupDir 'files'
  New-Item -ItemType Directory -Path $regDir,$fileDir -Force | Out-Null

  Export-KeyIfExists 'HKCU\Software\Microsoft\Office\16.0\Outlook\AutoDiscover' (Join-Path $regDir 'hkcu-office-autodiscover.reg')
  Export-KeyIfExists 'HKCU\Software\Policies\Microsoft\Office\16.0\Outlook\AutoDiscover' (Join-Path $regDir 'hkcu-policies-autodiscover.reg')
  Export-KeyIfExists 'HKCU\Software\Microsoft\Office\16.0\Outlook\Profiles' (Join-Path $regDir 'hkcu-office-profiles.reg')
  Export-KeyIfExists 'HKCU\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles' (Join-Path $regDir 'hkcu-messaging-profiles.reg')

  if (Test-Path $TargetDirectory) {
    $targetBackup = Join-Path $fileDir 'Autodiscover'
    Copy-Item $TargetDirectory $targetBackup -Recurse -Force
  }

  $outlookRoot = Join-Path $env:LOCALAPPDATA 'Microsoft\Outlook'
  $rootXmlDir = Join-Path $fileDir 'OutlookRootAutodiscoverXml'
  New-Item -ItemType Directory -Path $rootXmlDir -Force | Out-Null
  Get-ChildItem $outlookRoot -File -Filter '*-Autodiscover.xml' -ErrorAction SilentlyContinue |
    Copy-Item -Destination $rootXmlDir -Force

  $metadata = [ordered]@{
    Version = $Version
    CreatedAt = (Get-Date).ToString('o')
    Email = $Email
    Domain = $Domain
    BackupDirectory = $backupDir
    Note = 'Created by Update-IonosExchangeAutodiscover before changes.'
  }
  $metadata | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $backupDir 'backup-metadata.json') -Encoding UTF8
  Write-Info "Backup written: $backupDir"
}

function New-AutodiscoverRequestXml {
  param([string]$Address)

  return '<?xml version="1.0" encoding="utf-8"?>' +
    '<Autodiscover xmlns="http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006">' +
    '<Request>' +
    "<EMailAddress>$Address</EMailAddress>" +
    '<AcceptableResponseSchema>http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a</AcceptableResponseSchema>' +
    '</Request>' +
    '</Autodiscover>'
}

function Invoke-CurlAutodiscover {
  param(
    [string]$Uri,
    [string]$Address,
    [string]$OutputFile,
    [string]$HeaderFile,
    [switch]$NoRevoke
  )

  $sec = Read-Host "IONOS Exchange password for $Address" -AsSecureString
  $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
  $reqFile = Join-Path $env:TEMP ("ionos-autodiscover-request-$PID.xml")

  try {
    $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
    $pair = "$Address`:$plain"
    $b64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($pair))

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($reqFile, (New-AutodiscoverRequestXml -Address $Address), $utf8NoBom)

    $curlArgs = @(
      '-sS',
      '-D', $HeaderFile,
      '-o', $OutputFile,
      '--http1.1',
      '-X', 'POST',
      '-H', 'Content-Type: text/xml; charset=utf-8',
      '-H', 'User-Agent: Microsoft Office/16.0',
      '-H', "Authorization: Basic $b64",
      '--data-binary', "@$reqFile"
    )

    if ($NoRevoke) {
      Write-Warning 'Using curl.exe --ssl-no-revoke. This bypasses certificate revocation checking for this request and should only be used for diagnostics or temporary recovery.'
      $curlArgs += '--ssl-no-revoke'
    }

    $curlArgs += $Uri

    & curl.exe @curlArgs

    if ($LASTEXITCODE -ne 0) {
      throw "curl.exe failed with exit code $LASTEXITCODE."
    }
  }
  finally {
    if ($ptr) {
      [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
    }
    Remove-Variable plain,b64,pair,sec -ErrorAction SilentlyContinue
    Remove-Item $reqFile -Force -ErrorAction SilentlyContinue
  }
}

function Assert-AutodiscoverResponse {
  param(
    [string]$HeaderFile,
    [string]$XmlFile,
    [string]$HostName
  )

  $headers = Get-Content $HeaderFile -Raw
  $response = Get-Content $XmlFile -Raw

  if ($headers -notmatch 'HTTP/1\.1 200') {
    throw "Autodiscover request did not return HTTP 200. Headers:`n$headers"
  }
  if ($response -notmatch '<Action>settings</Action>') {
    throw 'Autodiscover response does not contain <Action>settings</Action>.'
  }
  if ($response -notmatch '<Type>EXPR</Type>') {
    throw 'Autodiscover response does not contain EXPR protocol settings.'
  }
  if ($response -notmatch [regex]::Escape($HostName)) {
    throw "Autodiscover response does not reference $HostName."
  }
}

function Add-OrUpdate-ChildText {
  param(
    [xml]$Document,
    [System.Xml.XmlElement]$Parent,
    [string]$Name,
    [string]$NamespaceUri,
    [string]$Value,
    [System.Xml.XmlNode]$AfterNode = $null
  )

  $manager = New-Object System.Xml.XmlNamespaceManager($Document.NameTable)
  $manager.AddNamespace('out', $NamespaceUri)
  $node = $Parent.SelectSingleNode("out:$Name", $manager)
  if (-not $node) {
    $node = $Document.CreateElement($Name, $NamespaceUri)
    if ($AfterNode) {
      [void]$Parent.InsertAfter($node, $AfterNode)
    }
    else {
      [void]$Parent.AppendChild($node)
    }
  }
  $node.InnerText = $Value
  return $node
}

function New-MapiFirstXml {
  param(
    [string]$SourceXml,
    [string]$TargetXml,
    [string]$Address,
    [string]$HostName
  )

  [xml]$doc = Get-Content $SourceXml -Raw
  $nsUri = 'http://schemas.microsoft.com/exchange/autodiscover/outlook/responseschema/2006a'
  $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
  $ns.AddNamespace('out', $nsUri)

  $account = $doc.SelectSingleNode('//out:Account', $ns)
  if (-not $account) {
    throw 'Account node not found in Autodiscover XML.'
  }

  # Remove mapiHttp blocks if the input already contains one.
  $existingMapi = @($account.SelectNodes("out:Protocol[out:Type='mapiHttp']", $ns))
  foreach ($node in $existingMapi) {
    [void]$account.RemoveChild($node)
  }

  $exprNodes = @($account.SelectNodes("out:Protocol[out:Type='EXPR']", $ns))
  $exchNodes = @($account.SelectNodes("out:Protocol[out:Type='EXCH']", $ns))
  $otherProtocolNodes = @($account.SelectNodes('out:Protocol', $ns)) | Where-Object {
    $typeNode = $_.SelectSingleNode('out:Type', $ns)
    $typeValue = if ($typeNode) { $typeNode.InnerText } else { '' }
    $typeValue -ne 'EXPR' -and $typeValue -ne 'EXCH' -and $typeValue -ne 'mapiHttp'
  }

  $mailboxId = $null
  if ($exchNodes.Count -gt 0) {
    $serverNode = $exchNodes[0].SelectSingleNode('out:Server', $ns)
    if ($serverNode -and $serverNode.InnerText) {
      $mailboxId = $serverNode.InnerText
    }
  }
  if (-not $mailboxId) {
    $mailboxId = $Address
  }
  $encodedMailboxId = [uri]::EscapeDataString($mailboxId)

  foreach ($protocol in @($account.SelectNodes('out:Protocol', $ns))) {
    [void]$account.RemoveChild($protocol)
  }

  $mapi = $doc.CreateElement('Protocol', $nsUri)
  $type = $doc.CreateElement('Type', $nsUri)
  $type.InnerText = 'mapiHttp'
  [void]$mapi.AppendChild($type)
  $login = $doc.CreateElement('LoginName', $nsUri)
  $login.InnerText = $Address
  [void]$mapi.AppendChild($login)

  $mailStore = $doc.CreateElement('MailStore', $nsUri)
  foreach ($name in @('ExternalUrl', 'InternalUrl')) {
    $node = $doc.CreateElement($name, $nsUri)
    $node.InnerText = "https://$HostName/mapi/emsmdb/?MailboxId=$encodedMailboxId"
    [void]$mailStore.AppendChild($node)
  }
  [void]$mapi.AppendChild($mailStore)

  $addressBook = $doc.CreateElement('AddressBook', $nsUri)
  foreach ($name in @('ExternalUrl', 'InternalUrl')) {
    $node = $doc.CreateElement($name, $nsUri)
    $node.InnerText = "https://$HostName/mapi/nspi/?MailboxId=$encodedMailboxId"
    [void]$addressBook.AppendChild($node)
  }
  [void]$mapi.AppendChild($addressBook)

  [void]$account.AppendChild($mapi)

  foreach ($node in $exprNodes) {
    $typeNode = $node.SelectSingleNode('out:Type', $ns)
    [void](Add-OrUpdate-ChildText -Document $doc -Parent $node -Name 'LoginName' -NamespaceUri $nsUri -Value $Address -AfterNode $typeNode)
    [void]$account.AppendChild($node)
  }
  foreach ($node in $exchNodes) {
    $typeNode = $node.SelectSingleNode('out:Type', $ns)
    [void](Add-OrUpdate-ChildText -Document $doc -Parent $node -Name 'LoginName' -NamespaceUri $nsUri -Value $Address -AfterNode $typeNode)
    [void]$account.AppendChild($node)
  }
  foreach ($node in $otherProtocolNodes) {
    [void]$account.AppendChild($node)
  }

  $settings = New-Object System.Xml.XmlWriterSettings
  $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
  $settings.Indent = $true
  $writer = [System.Xml.XmlWriter]::Create($TargetXml, $settings)
  try {
    $doc.Save($writer)
  }
  finally {
    $writer.Close()
  }
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

function Set-DWordValue {
  param(
    [string]$Path,
    [string]$Name,
    [int]$Value
  )
  $regPath = Convert-ToRegExePath -Path $Path
  Invoke-RegExe -Arguments @('add', $regPath, '/v', $Name, '/t', 'REG_DWORD', '/d', ([string]$Value), '/f')
}

function Set-StringValue {
  param(
    [string]$Path,
    [string]$Name,
    [string]$Value
  )
  $regPath = Convert-ToRegExePath -Path $Path
  Invoke-RegExe -Arguments @('add', $regPath, '/v', $Name, '/t', 'REG_SZ', '/d', $Value, '/f')
}

function Set-OutlookAutodiscoverRegistry {
  param(
    [string]$Domain,
    [string]$MapiFirstXml,
    [string]$ExchangeHost
  )

  $keys = @(
    'HKCU:\Software\Microsoft\Office\16.0\Outlook\AutoDiscover',
    'HKCU:\Software\Policies\Microsoft\Office\16.0\Outlook\AutoDiscover'
  )

  foreach ($key in $keys) {
    Set-DWordValue -Path $key -Name 'ExcludeExplicitO365Endpoint' -Value 1
    Set-DWordValue -Path $key -Name 'ExcludeHttpsRootDomain' -Value 1
    Set-DWordValue -Path $key -Name 'ExcludeHttpsAutoDiscoverDomain' -Value 1
    Set-DWordValue -Path $key -Name 'ExcludeLastKnownGoodUrl' -Value 1
    Set-DWordValue -Path $key -Name 'PreferLocalXML' -Value 1
    Set-StringValue -Path $key -Name $Domain -Value $MapiFirstXml
    & reg.exe delete (Convert-ToRegExePath -Path $key) /v LastKnownGoodUrl /f 2>$null | Out-Null

    if ($RegisterRedirectServers) {
      $redirectKey = Join-Path $key 'RedirectServers'
      New-Item -Path $redirectKey -Force | Out-Null
      Set-StringValue -Path $redirectKey -Name 'autodiscover.1and1.info' -Value ''
      Set-StringValue -Path $redirectKey -Name $ExchangeHost -Value ''
    }
  }
}

function Write-DesiredStateManifest {
  param(
    [string]$Domain,
    [string]$Address,
    [string]$HostName,
    [string]$FullXml,
    [string]$MapiFirstXml,
    [string]$ManifestFile
  )

  $registryPaths = @(
    'HKCU:\Software\Microsoft\Office\16.0\Outlook\AutoDiscover',
    'HKCU:\Software\Policies\Microsoft\Office\16.0\Outlook\AutoDiscover'
  )

  $state = [ordered]@{
    Version = $Version
    CreatedAt = (Get-Date).ToString('o')
    Email = $Address
    Domain = $Domain
    ExchangeHost = $HostName
    FullXmlPath = $FullXml
    MapiFirstXmlPath = $MapiFirstXml
    RegistryPaths = $registryPaths
    ExpectedDwordValues = [ordered]@{
      ExcludeExplicitO365Endpoint = 1
      ExcludeHttpsRootDomain = 1
      ExcludeHttpsAutoDiscoverDomain = 1
      ExcludeLastKnownGoodUrl = 1
      PreferLocalXML = 1
    }
    ExpectedStringValues = [ordered]@{
      $Domain = $MapiFirstXml
    }
    ManagedFiles = @(
      $FullXml,
      $MapiFirstXml,
      $ManifestFile
    )
    Notes = @(
      'This manifest describes the desired local Outlook Autodiscover state created by the kit.',
      'It intentionally does not store credentials.',
      'Use Ensure-OutlookIonosState to re-apply registry values without fetching Autodiscover again.'
    )
  }

  $state | ConvertTo-Json -Depth 8 | Set-Content $ManifestFile -Encoding UTF8
}

Stop-OutlookIfRequested

$domain = Get-DomainFromEmail -Address $Email
$prefix = ConvertTo-SafeFilePrefix -Domain $domain
if (-not $AutodiscoverUrl) {
  $AutodiscoverUrl = "https://$ExchangeHost/Autodiscover/autodiscover.xml"
}

New-Item -ItemType Directory -Path $TargetDirectory -Force | Out-Null

if ($BackupBeforeChanges) {
  Invoke-StateBackup -Domain $domain -TargetDirectory $TargetDirectory
}

$fullXml = Join-Path $TargetDirectory ("$prefix-autodiscover-full.xml")
$mapiFirstXml = Join-Path $TargetDirectory ("$prefix-autodiscover-mapi-first.xml")
$manifest = Join-Path $TargetDirectory ("$prefix-ionos-outlook-kit-state.json")
$tmpFull = Join-Path $env:TEMP ("ionos-autodiscover-full-$PID.xml")
$tmpHeaders = Join-Path $env:TEMP ("ionos-autodiscover-headers-$PID.txt")

if ($SetRegistryFromExistingXml) {
  if (-not (Test-Path $mapiFirstXml)) {
    throw "Existing MAPI/HTTP-first XML not found: $mapiFirstXml. Run the script once without -SetRegistryFromExistingXml after fixing HTTPS/certificate access."
  }

  if (-not (Test-Path $fullXml)) {
    Write-Warning "Full Autodiscover XML not found: $fullXml. The registry can still be restored from the existing MAPI/HTTP-first XML, but the state manifest will reference a missing full XML file."
  }

  Set-OutlookAutodiscoverRegistry -Domain $domain -MapiFirstXml $mapiFirstXml -ExchangeHost $ExchangeHost
  Write-Info "Registry updated from existing local XML for domain: $domain"

  Write-DesiredStateManifest -Domain $domain -Address $Email -HostName $ExchangeHost -FullXml $fullXml -MapiFirstXml $mapiFirstXml -ManifestFile $manifest
  Write-Info "Desired-state manifest written: $manifest"
  Write-Info 'Done.'
  return
}

try {
  Invoke-CurlAutodiscover -Uri $AutodiscoverUrl -Address $Email -OutputFile $tmpFull -HeaderFile $tmpHeaders -NoRevoke:$SslNoRevoke
  Assert-AutodiscoverResponse -HeaderFile $tmpHeaders -XmlFile $tmpFull -HostName $ExchangeHost

  Copy-Item $tmpFull $fullXml -Force
  Write-Info "Full Autodiscover XML written: $fullXml"

  New-MapiFirstXml -SourceXml $fullXml -TargetXml $mapiFirstXml -Address $Email -HostName $ExchangeHost
  Write-Info "MAPI/HTTP-first Autodiscover XML written: $mapiFirstXml"

  if ($SetRegistry) {
    Set-OutlookAutodiscoverRegistry -Domain $domain -MapiFirstXml $mapiFirstXml -ExchangeHost $ExchangeHost
    Write-Info "Registry updated for domain: $domain"
  }

  Write-DesiredStateManifest -Domain $domain -Address $Email -HostName $ExchangeHost -FullXml $fullXml -MapiFirstXml $mapiFirstXml -ManifestFile $manifest
  Write-Info "Desired-state manifest written: $manifest"
  Write-Info 'Done.'
}
finally {
  Remove-Item $tmpFull,$tmpHeaders -Force -ErrorAction SilentlyContinue
}
