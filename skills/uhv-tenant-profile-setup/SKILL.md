# Skill: UHV Tenant Profile Setup

## Purpose

Create a local, reusable auth/profile setup so deploy/update commands are fast and consistent.

## Local Profile File

1. Copy template:

```powershell
Copy-Item .\scripts\examples\UHV.LocalProfile.example.ps1 .\ignore\UHV.LocalProfile.ps1
```

2. Edit local file with tenant values:

```powershell
$env:UHV_CLIENT_ID = "<entra-app-client-id-guid>"
$env:UHV_TENANT = "<tenant>.onmicrosoft.com"

$Global:UhvTenantName = "<tenant>"
$Global:UhvAppCatalogUrl = "https://<tenant>.sharepoint.com/sites/appcatalog"
$Global:UhvTenantAdminUrl = "https://<tenant>-admin.sharepoint.com"
$Global:UhvSiteUrl = "https://<tenant>.sharepoint.com/sites/TestUHV2"
```

3. Load profile each session:

```powershell
. .\ignore\UHV.LocalProfile.ps1
```

## Connectivity Test

```powershell
Connect-PnPOnline -Url $Global:UhvSiteUrl -DeviceLogin -ClientId $env:UHV_CLIENT_ID -Tenant $env:UHV_TENANT -PersistLogin
Get-PnPWeb | Select-Object Title, Url
```

## Fast Deploy Using Profile

```powershell
.\scripts\Deploy-UHV-All.ps1 `
  -ClientId $env:UHV_CLIENT_ID `
  -Tenant $env:UHV_TENANT `
  -AppCatalogUrl $Global:UhvAppCatalogUrl `
  -TenantAdminUrl $Global:UhvTenantAdminUrl `
  -DeviceLogin
```

## Rules

- Keep profile files inside `ignore/` (non-committed).
- Do not hardcode secrets in committed scripts/docs.
- Prefer `-PersistLogin` for smoother repeated admin operations.

