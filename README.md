# IONOS Outlook Kit

Version: 0.0.17

The release ZIP is versioned. The repository files and the files inside the ZIP use stable, unversioned script names; the script version is recorded inside each file.

This kit helps classic Outlook for Microsoft 365 connect to IONOS Hosted Microsoft Exchange 2019 mailboxes when Outlook Autodiscover is unreliable, falls back to Microsoft 365 / Outlook.com endpoints, or selects an unusable RPC/HTTP transport path.

The kit creates and registers a local MAPI/HTTP-first Autodiscover XML for the mailbox domain and stores a desired-state manifest. This makes the working local configuration checkable and quickly repairable if Outlook, Microsoft 365, Windows, policy refresh, or security software later removes required registry values.

The scripts do not store the mailbox password.

## Table of contents

- [Problem solved](#problem-solved)
- [How the kit works](#how-the-kit-works)
- [Safety model](#safety-model)
- [Repository layout](#repository-layout)
- [Installation from ZIP](#installation-from-zip)
- [First-time setup](#first-time-setup)
- [Fast repair after Outlook, Microsoft, or Windows changes](#fast-repair-after-outlook-microsoft-or-windows-changes)
- [Check the local desired state](#check-the-local-desired-state)
- [Live refresh from IONOS](#live-refresh-from-ionos)
- [Create or repair the Outlook profile](#create-or-repair-the-outlook-profile)
- [Disposable Outlook Autodiscover test profiles](#disposable-outlook-autodiscover-test-profiles)
- [Targeted cleanup](#targeted-cleanup)
- [Rollback](#rollback)
- [Managed files and registry values](#managed-files-and-registry-values)
- [IONOS-documented Microsoft 365 Autodiscover mitigation](#ionos-documented-microsoft-365-autodiscover-mitigation)
- [Troubleshooting](#troubleshooting)
- [Limitations](#limitations)
- [References and background](#references-and-background)

## Problem solved

Classic Outlook can be difficult to connect to an IONOS Hosted Exchange mailbox when the Windows user or the mailbox domain is also related to Microsoft 365. Typical symptoms include:

- Outlook tries Microsoft 365, Outlook.com, `outlook.office365.com`, `outlook.live.com`, or `login.microsoftonline.com` instead of the IONOS Hosted Exchange endpoint.
- Outlook Autodiscover reaches IONOS but the resulting profile still does not open the mailbox.
- Outlook receives usable IONOS Autodiscover data but chooses an unusable RPC/HTTP path.
- `https://exchange2019.ionos.eu/rpc/rpcproxy.dll` returns `503 Service Unavailable` while `https://exchange2019.ionos.eu/mapi/emsmdb/` returns `401 Unauthorized`.
- Outlook works after manual registry tuning, but disconnects again after an Outlook, Microsoft 365, Windows, policy, or security-product change.
- Security software interferes with Exchange HTTPS traffic and Windows Schannel certificate revocation checks.

The goal is not to replace Outlook account setup. The goal is to make the known-good IONOS Exchange Autodiscover path explicit, local, verifiable, and repairable.

## How the kit works

The working approach is:

1. Fetch the live IONOS Autodiscover response for the mailbox.
2. Store a local full Autodiscover XML response.
3. Create a local MAPI/HTTP-first Autodiscover XML.
4. Register that local XML for the mailbox domain through Outlook Autodiscover registry values.
5. Store a desired-state manifest for later checks and fast repair.
6. Provide an `Ensure` script that can re-apply the desired registry state without network access and without a mailbox password.

The productive local XML is the `*-autodiscover-mapi-first.xml` file. It puts a `mapiHttp` block before the IONOS `EXPR` and `EXCH` protocol blocks and points Outlook to:

```text
https://exchange2019.ionos.eu/mapi/emsmdb/
https://exchange2019.ionos.eu/mapi/nspi/
```

This is useful when MAPI/HTTP is reachable but the older RPC/HTTP path is unavailable or returns `503 Service Unavailable`.

## Safety model

Use this kit conservatively on machines that may contain multiple productive Outlook profiles.

Important rules:

- Do not broadly delete Outlook profiles on multi-profile machines.
- Back up local Outlook/Autodiscover state before first-time changes.
- Treat `ADTest*` profiles as disposable test profiles.
- Do not delete or move productive `.ost` files unless explicitly intended.
- Use `Ensure-OutlookIonosState` for fast repair if local XML files are still present.
- Use live update only if the local XML has to be refreshed from IONOS.
- Use `-SslNoRevoke` only as a diagnostic or temporary recovery option, not as the normal operating mode.

OST handling defaults to `Skip`, because OST files are usually large Exchange cache files and can be rebuilt by Outlook.

## Repository layout

```text
ionos-outlook-kit/
  README.md
  Backup-OutlookIonosState.ps1
  Restore-OutlookIonosState.ps1
  Update-IonosExchangeAutodiscover.ps1
  Ensure-OutlookIonosState.ps1
  Check-OutlookIonosState.ps1
  Cleanup-OutlookIonosTestArtifacts.ps1
```

## Installation from ZIP

Assuming the ZIP is in the user's Downloads folder:

```powershell
$version = "0.0.17"
$zip = "$env:USERPROFILE\Downloads\ionos-outlook-kit-$version.zip"
$extractDir = "$env:TEMP\ionos-outlook-kit-$version"
$targetDir = "$env:LOCALAPPDATA\Microsoft\Outlook\Autodiscover"

Remove-Item $extractDir -Recurse -Force -ErrorAction SilentlyContinue
Expand-Archive -Path $zip -DestinationPath $extractDir -Force
New-Item -ItemType Directory -Path $targetDir -Force | Out-Null

Copy-Item "$extractDir\ionos-outlook-kit\*.ps1" $targetDir -Force
```

After installation, all scripts are available under:

```text
%LOCALAPPDATA%\Microsoft\Outlook\Autodiscover
```

## First-time setup

Use this flow when setting up a mailbox for the first time on a machine or when the local XML files do not yet exist.

### 1. Back up current state

```powershell
powershell.exe -ExecutionPolicy Bypass `
  -File "$env:LOCALAPPDATA\Microsoft\Outlook\Autodiscover\Backup-OutlookIonosState.ps1" `
  -Email "user@example.tld" `
  -StopOutlook
```

If a large OST must deliberately be moved out of the active Outlook directory:

```powershell
powershell.exe -ExecutionPolicy Bypass `
  -File "$env:LOCALAPPDATA\Microsoft\Outlook\Autodiscover\Backup-OutlookIonosState.ps1" `
  -Email "user@example.tld" `
  -StopOutlook `
  -OstHandling Move
```

Use `Move` only intentionally. It removes the OST from the active Outlook directory.

### 2. Create local IONOS Autodiscover XML and register it

```powershell
powershell.exe -ExecutionPolicy Bypass `
  -File "$env:LOCALAPPDATA\Microsoft\Outlook\Autodiscover\Update-IonosExchangeAutodiscover.ps1" `
  -Email "user@example.tld" `
  -SetRegistry `
  -BackupBeforeChanges `
  -StopOutlook
```

Default Exchange host:

```text
exchange2019.ionos.eu
```

Override it only if needed:

```powershell
-ExchangeHost "exchange2019.ionos.eu"
```

The update script creates:

```text
%LOCALAPPDATA%\Microsoft\Outlook\Autodiscover\<domain>-autodiscover-full.xml
%LOCALAPPDATA%\Microsoft\Outlook\Autodiscover\<domain>-autodiscover-mapi-first.xml
%LOCALAPPDATA%\Microsoft\Outlook\Autodiscover\<domain>-ionos-outlook-kit-state.json
```

### 3. Check the result

```powershell
powershell.exe -ExecutionPolicy Bypass `
  -File "$env:LOCALAPPDATA\Microsoft\Outlook\Autodiscover\Check-OutlookIonosState.ps1" `
  -Email "user@example.tld" `
  -ProfileName "IONOS-Exchange-MAPI" `
  -Detailed
```

A healthy local state should end with:

```text
Summary: 0 FAIL, 0 WARN
```

## Fast repair after Outlook, Microsoft, or Windows changes

Use this when Outlook was already working and later disconnects, especially after updates or policy/security-product changes.

First check the local state:

```powershell
powershell.exe -ExecutionPolicy Bypass `
  -File "$env:LOCALAPPDATA\Microsoft\Outlook\Autodiscover\Check-OutlookIonosState.ps1" `
  -Email "user@example.tld" `
  -ProfileName "IONOS-Exchange-MAPI" `
  -Detailed
```

If XML files and the profile are present but registry values are missing, re-apply the desired state without network access and without entering the mailbox password.

Dry-run:

```powershell
powershell.exe -ExecutionPolicy Bypass `
  -File "$env:LOCALAPPDATA\Microsoft\Outlook\Autodiscover\Ensure-OutlookIonosState.ps1" `
  -Email "user@example.tld"
```

Apply and verify:

```powershell
powershell.exe -ExecutionPolicy Bypass `
  -File "$env:LOCALAPPDATA\Microsoft\Outlook\Autodiscover\Ensure-OutlookIonosState.ps1" `
  -Email "user@example.tld" `
  -StopOutlook `
  -Execute
```

Apply all local IONOS desired-state manifests on the machine:

```powershell
powershell.exe -ExecutionPolicy Bypass `
  -File "$env:LOCALAPPDATA\Microsoft\Outlook\Autodiscover\Ensure-OutlookIonosState.ps1" `
  -All `
  -StopOutlook `
  -Execute
```

`Ensure-OutlookIonosState` writes registry values with `reg.exe` and verifies them after writing. If the desired-state manifest is missing but the local XML files are still present, it can infer the desired state from the mailbox domain and the standard local file names. With `-Execute`, it also writes the missing manifest.

This is the fastest repair path when a known-good local XML is still available and only the registry switches were lost.

## Check the local desired state

Basic check:

```powershell
powershell.exe -ExecutionPolicy Bypass `
  -File "$env:LOCALAPPDATA\Microsoft\Outlook\Autodiscover\Check-OutlookIonosState.ps1" `
  -Email "user@example.tld" `
  -ProfileName "IONOS-Exchange-MAPI" `
  -Detailed
```

Optional online endpoint check:

```powershell
powershell.exe -ExecutionPolicy Bypass `
  -File "$env:LOCALAPPDATA\Microsoft\Outlook\Autodiscover\Check-OutlookIonosState.ps1" `
  -Email "user@example.tld" `
  -ProfileName "IONOS-Exchange-MAPI" `
  -OnlineCheck `
  -Detailed
```

A successful MAPI endpoint check normally returns `HTTP/1.1 401 Unauthorized`. This is good: it means the Exchange endpoint is reachable and requires authentication.

For diagnostics only, the check script supports `-SslNoRevoke` for online endpoint checks:

```powershell
powershell.exe -ExecutionPolicy Bypass `
  -File "$env:LOCALAPPDATA\Microsoft\Outlook\Autodiscover\Check-OutlookIonosState.ps1" `
  -Email "user@example.tld" `
  -ProfileName "IONOS-Exchange-MAPI" `
  -OnlineCheck `
  -SslNoRevoke `
  -Detailed
```

A healthy Windows/Kaspersky configuration should not need `-SslNoRevoke`.

## Live refresh from IONOS

Use live refresh only when the local XML files are missing or need to be refreshed from IONOS.

```powershell
powershell.exe -ExecutionPolicy Bypass `
  -File "$env:LOCALAPPDATA\Microsoft\Outlook\Autodiscover\Update-IonosExchangeAutodiscover.ps1" `
  -Email "user@example.tld" `
  -SetRegistry `
  -BackupBeforeChanges `
  -StopOutlook
```

If `curl.exe` fails with `CRYPT_E_NO_REVOCATION_CHECK`, repair from existing local XML first if possible, then investigate TLS/certificate revocation scanning separately.

For live update diagnostics or temporary recovery only, the update script supports `-SslNoRevoke`. This passes `--ssl-no-revoke` to `curl.exe` for the Autodiscover download.

```powershell
powershell.exe -ExecutionPolicy Bypass `
  -File "$env:LOCALAPPDATA\Microsoft\Outlook\Autodiscover\Update-IonosExchangeAutodiscover.ps1" `
  -Email "user@example.tld" `
  -SetRegistry `
  -StopOutlook `
  -SslNoRevoke
```

Do not use `-SslNoRevoke` as the normal operating mode. Fix the Windows/Kaspersky certificate revocation path instead.

If the local XML already exists and only registry values must be restored, use this instead of live refresh:

```powershell
powershell.exe -ExecutionPolicy Bypass `
  -File "$env:LOCALAPPDATA\Microsoft\Outlook\Autodiscover\Update-IonosExchangeAutodiscover.ps1" `
  -Email "user@example.tld" `
  -SetRegistryFromExistingXml `
  -StopOutlook
```

## Create or repair the Outlook profile

After the local XML and registry values are in place, create or repair only the specific IONOS Outlook profile. Do not delete unrelated production profiles.

Suggested profile name:

```text
IONOS-Exchange-MAPI
```

In classic Outlook account setup:

1. Enter the IONOS mailbox address, for example `user@example.tld`.
2. Select manual setup.
3. Choose `Exchange`.
4. Allow IONOS Autodiscover redirects when the host is `exchange2019.ionos.eu`.
5. Use the IONOS Exchange mailbox password when Outlook asks for credentials.

If Outlook displays a redirect prompt such as:

```text
https://exchange2019.ionos.eu/autodiscover/autodiscover.xml
```

this is expected for IONOS Hosted Exchange. It can be allowed if the host is the expected IONOS Exchange host.

## Disposable Outlook Autodiscover test profiles

Disposable test profiles should use names like:

```text
ADTestMapiHttp
ADTestMapiHttp2
ADTestMapiHttp3
ADTestNtlmFresh
```

Use a new name for repeated tests because Outlook refuses to create a `PIM` profile if the name already exists.

Example:

```powershell
& "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE" /PIM "ADTestMapiHttp2"
```

Then use Ctrl + right-click on the small Outlook tray icon and select `Test E-mail AutoConfiguration`.

Only enable:

```text
[x] Use Autodiscover
[ ] Use Guessmart
[ ] Secure Guessmart Authentication
```

Expected result for the productive local XML:

```text
Protocol: Exchange-MAPI-HTTP
... exchange2019.ionos.eu/mapi/emsmdb/ ...
... exchange2019.ionos.eu/mapi/nspi/ ...
```

## Targeted cleanup

Prefer targeted cleanup over broad profile deletion.

Dry-run for disposable `ADTest*` profiles only:

```powershell
powershell.exe -ExecutionPolicy Bypass `
  -File "$env:LOCALAPPDATA\Microsoft\Outlook\Autodiscover\Cleanup-OutlookIonosTestArtifacts.ps1" `
  -Email "user@example.tld" `
  -RemoveAdTestProfilesOnly
```

Apply:

```powershell
powershell.exe -ExecutionPolicy Bypass `
  -File "$env:LOCALAPPDATA\Microsoft\Outlook\Autodiscover\Cleanup-OutlookIonosTestArtifacts.ps1" `
  -Email "user@example.tld" `
  -RemoveAdTestProfilesOnly `
  -StopOutlook `
  -Execute
```

On multi-profile systems, prefer this mode. Do not remove unrelated Outlook profiles.

## Rollback

Rollback can restore registry and managed files from a backup created before the update.

Dry-run:

```powershell
powershell.exe -ExecutionPolicy Bypass `
  -File "$env:LOCALAPPDATA\Microsoft\Outlook\Autodiscover\Restore-OutlookIonosState.ps1" `
  -BackupDirectory "$env:USERPROFILE\Desktop\outlook-ionos-backups\<backup-folder>" `
  -RestoreRegistry `
  -RestoreFiles `
  -StopOutlook
```

Apply:

```powershell
powershell.exe -ExecutionPolicy Bypass `
  -File "$env:LOCALAPPDATA\Microsoft\Outlook\Autodiscover\Restore-OutlookIonosState.ps1" `
  -BackupDirectory "$env:USERPROFILE\Desktop\outlook-ionos-backups\<backup-folder>" `
  -RestoreRegistry `
  -RestoreFiles `
  -StopOutlook `
  -Execute
```

Rollback removes managed Autodiscover files created after the backup if they were not present in the saved state.

## Managed files and registry values

### Managed files

The productive files for a domain are:

```text
%LOCALAPPDATA%\Microsoft\Outlook\Autodiscover\<domain>-autodiscover-full.xml
%LOCALAPPDATA%\Microsoft\Outlook\Autodiscover\<domain>-autodiscover-mapi-first.xml
%LOCALAPPDATA%\Microsoft\Outlook\Autodiscover\<domain>-ionos-outlook-kit-state.json
```

The desired-state manifest is intentionally kept. It allows fast detection and repair if registry values disappear after updates or policy refresh.

### Managed registry paths

The scripts manage these values under both paths:

```text
HKCU\Software\Microsoft\Office\16.0\Outlook\AutoDiscover
HKCU\Software\Policies\Microsoft\Office\16.0\Outlook\AutoDiscover
```

Expected DWORD values:

```text
ExcludeExplicitO365Endpoint        = 1
ExcludeHttpsRootDomain             = 1
ExcludeHttpsAutoDiscoverDomain     = 1
ExcludeLastKnownGoodUrl            = 1
PreferLocalXML                     = 1
```

`ExcludeExplicitO365Endpoint` and `ExcludeHttpsRootDomain` are the IONOS-documented Microsoft 365 Autodiscover mitigation values. The remaining DWORD values are used by this kit to pin Outlook to the local MAPI/HTTP-first XML and to avoid stale Autodiscover cache behavior.

Expected string value:

```text
<domain> = %LOCALAPPDATA%\Microsoft\Outlook\Autodiscover\<domain>-autodiscover-mapi-first.xml
```

The scripts also remove `LastKnownGoodUrl` from the managed Autodiscover keys if it exists.

Optional trusted redirect entries can exist under:

```text
HKCU\Software\Microsoft\Office\16.0\Outlook\AutoDiscover\RedirectServers
HKCU\Software\Policies\Microsoft\Office\16.0\Outlook\AutoDiscover\RedirectServers
```

For IONOS Hosted Exchange, the relevant redirect host is usually:

```text
exchange2019.ionos.eu
```

## IONOS-documented Microsoft 365 Autodiscover mitigation

IONOS documents a Microsoft 365 Autodiscover conflict for hosted Exchange mailboxes. In that scenario Outlook may try to sign in to the domain registered in Microsoft 365 instead of using the hosted Exchange mailbox endpoint.

The IONOS-documented registry values are:

    ExcludeExplicitO365Endpoint = 1
    ExcludeHttpsRootDomain      = 1

This kit treats those two values as the vendor-documented baseline and always writes and checks them when registry state is applied. The kit also sets additional values that are needed for the local MAPI/HTTP-first XML strategy:

    ExcludeHttpsAutoDiscoverDomain = 1
    ExcludeLastKnownGoodUrl        = 1
    PreferLocalXML                 = 1
    <mailbox-domain>               = <path-to-mapi-first-xml>

The additional values are not a replacement for the IONOS-documented mitigation. They make the known-good local Autodiscover XML explicit, stable, and quickly repairable if Outlook, Microsoft 365, Windows, policy refresh, or security software later removes the registry values.

## Troubleshooting

### Quick decision guide

Use this rough decision tree:

```text
Check script reports missing registry values, but XML/profile are OK
  -> Run Ensure-OutlookIonosState ... -Execute

Check script reports XML missing or stale
  -> Run live Update-IonosExchangeAutodiscover ... -SetRegistry

curl to /mapi/emsmdb/ returns HTTP/1.1 401 Unauthorized
  -> Exchange MAPI/HTTP endpoint is reachable

curl fails with CRYPT_E_NO_REVOCATION_CHECK
  -> Investigate Windows Schannel / Kaspersky / encrypted traffic scanning

/rpc/rpcproxy.dll returns 503 but /mapi/emsmdb/ returns 401
  -> Prefer the local MAPI/HTTP-first XML path

Outlook is connected and says all folders are up to date
  -> Do not change the profile
```

### Kaspersky / encrypted traffic scanning

A confirmed working configuration with Kaspersky Plus was:

```text
Program: C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE
Enabled: Do not block interaction with the AMSI Protection component
Enabled: Do not scan encrypted traffic
```

In the German Kaspersky Plus UI these were shown as:

```text
Interaktion mit der AMSI-Schutzkomponente nicht blockieren
Verschlüsselten Datenverkehr nicht untersuchen
```

The second setting is critical for `CRYPT_E_NO_REVOCATION_CHECK` during Exchange HTTPS access. Without it, `curl.exe` and Outlook may fail before the HTTP request is sent.

After applying the Outlook exception, this command should normally return `HTTP/1.1 401 Unauthorized`:

```powershell
curl.exe -v -I https://exchange2019.ionos.eu/mapi/emsmdb/
```

Other trusted-application settings that may be useful during troubleshooting include:

```text
Do not monitor application activity
Do not monitor child application activity
```

For the PowerShell scripts, an AMSI/script scanning exception for the local kit directory may be useful:

```text
%LOCALAPPDATA%\Microsoft\Outlook\Autodiscover
```

Only apply such exclusions when the scripts are trusted and stored in a controlled local directory.

### CRYPT_E_NO_REVOCATION_CHECK

Diagnostic pattern:

```powershell
curl.exe -v -I https://exchange2019.ionos.eu/mapi/emsmdb/
curl.exe -v -I --ssl-no-revoke https://exchange2019.ionos.eu/mapi/emsmdb/
```

Interpretation:

```text
without --ssl-no-revoke -> CRYPT_E_NO_REVOCATION_CHECK
with    --ssl-no-revoke -> HTTP/1.1 401 Unauthorized
```

This means the Exchange endpoint is reachable, but Windows Schannel cannot complete certificate revocation checking. Fix security software, encrypted traffic inspection, proxy, or Windows certificate revocation access instead of using `--ssl-no-revoke` permanently.

### Exchange endpoint checks

Autodiscover endpoint:

```powershell
curl.exe -v -I https://exchange2019.ionos.eu/Autodiscover/autodiscover.xml
```

MAPI/HTTP endpoint:

```powershell
curl.exe -v -I https://exchange2019.ionos.eu/mapi/emsmdb/
```

RPC/HTTP endpoint:

```powershell
curl.exe -v -I https://exchange2019.ionos.eu/rpc/rpcproxy.dll
```

Expected useful results:

```text
/Autodiscover/autodiscover.xml -> 401 Unauthorized
/mapi/emsmdb/                  -> 401 Unauthorized
/rpc/rpcproxy.dll              -> may return 503 Service Unavailable
```

`401 Unauthorized` is useful in these unauthenticated tests. It means the service is reachable and asks for authentication.

### Outlook status checks

If Outlook is disconnected but the kit check reports `0 FAIL, 0 WARN`:

1. Verify that `Send/Receive > Work Offline` is not enabled.
2. Check the Outlook connection status from the tray icon.
3. Re-test `curl.exe -v -I https://exchange2019.ionos.eu/mapi/emsmdb/`.
4. If the curl test fails with `CRYPT_E_NO_REVOCATION_CHECK`, fix Kaspersky/TLS inspection first.
5. If curl returns `401 Unauthorized` and Outlook still stays disconnected, inspect credentials, profile state, and Outlook connection diagnostics.

### Public folder redirect prompt

Outlook may ask whether it is allowed to configure settings for a website such as:

```text
https://exchange2019.ionos.eu/autodiscover/autodiscover.xml
```

A prompt mentioning an address like `PublicFolderMBX...@exchange2019.ionos.eu` is plausible for IONOS Exchange public folder/system mailbox discovery. It can be allowed if the host is the expected IONOS Exchange host.

## Limitations

The kit manages only the local files and registry settings it explicitly backs up or creates. It does not manage:

- IONOS server-side outages or configuration changes.
- Microsoft 365 tenant state.
- Windows account broker state.
- Saved mailbox passwords in Windows Credential Manager.
- Third-party security product policy refresh.
- Manual Outlook changes made outside this kit after the backup.
- Security implications of disabling encrypted traffic scanning for Outlook in a third-party security product.

## References and background

- IONOS: [Set up Microsoft Exchange in classic Outlook for Microsoft 365](https://www.ionos.de/hilfe/e-mail/akkordeons-zu-microsoftr-exchange/microsoftr-exchange-einrichten/microsoft-exchanger-im-klassischen-outlook-microsoft-365-einrichten/). This documents the normal IONOS setup flow for Microsoft Exchange in classic Outlook for Microsoft 365.
- IONOS: [Disable Microsoft 365 Autodiscover in Outlook](https://www.ionos.de/hilfe/e-mail/akkordeons-zu-microsoftr-exchange/microsoft-exchange-einrichten/autodiscover-fuer-microsoft-365-in-outlook-deaktivieren/). This documents the hosted-Exchange/Microsoft-365 Autodiscover conflict and includes `ExcludeExplicitO365Endpoint=1` and `ExcludeHttpsRootDomain=1`.
- IONOS: [Autodiscover explained](https://www.ionos.de/hilfe/domains/glossar-domain-fachbegriffe-verstaendlich-erklaert/autodiscover/). This describes the IONOS Autodiscover DNS mechanism.
- IONOS: [Set up Microsoft Exchange in Outlook 2024](https://www.ionos.de/hilfe/e-mail/akkordeons-zu-microsoftr-exchange/microsoftr-exchange-einrichten/microsoft-exchanger-in-outlook-2024-einrichten/). This includes the user-facing Autodiscover authorization prompt.
- Microsoft: [Control Outlook AutoDiscover by using Group Policy](https://learn.microsoft.com/en-us/microsoft-365-apps/outlook/profiles-and-accounts/control-autodiscover-via-group-policy). This documents Outlook Autodiscover behavior and configurable Autodiscover methods.
- Microsoft: [Suppress AutoDiscover redirect warnings in Outlook](https://learn.microsoft.com/en-us/troubleshoot/outlook/connectivity/suppress-autodiscover-redirect-warning). This documents the `RedirectServers` registry key and empty `REG_SZ` entries for trusted redirect targets.
- Microsoft: [Outlook cannot connect or web services do not work after migration to Microsoft 365](https://learn.microsoft.com/en-us/troubleshoot/outlook/profiles-and-accounts/cannot-connect-web-service-not-working-migrated-to-office-365). This documents `ExcludeLastKnownGoodUrl=1` under both user and policy Autodiscover registry paths.
- Kaspersky: [Threats and Exclusions](https://support.kaspersky.com/kaspersky-for-windows/21.8/201385). This documents trusted application exclusions.
- Kaspersky: [How to change encrypted connections settings](https://support.kaspersky.com/kaspersky-for-windows/21.16/157530). This documents encrypted connection scanning behavior and trusted addresses.
- Kaspersky: [Exclude scripts from AMSI scanning](https://support.kaspersky.com/kaspersky-for-windows/21.17/186114). This documents AMSI script scanning exclusions.
