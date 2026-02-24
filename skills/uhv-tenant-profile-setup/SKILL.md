# Skill: UHV Tenant Profile Setup

## Purpose

Create a local, reusable auth/profile setup so deploy/update commands are fast and consistent.

## Local Profile File (Required)

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

## Recommended Additions (Team Defaults)

```powershell
$Global:UhvTargetSites = @(
  "https://<tenant>.sharepoint.com/sites/TestUHV1",
  "https://<tenant>.sharepoint.com/sites/TestUHV2",
  "https://<tenant>.sharepoint.com"
)
```

## Connectivity Test (Expected: site title + URL)

```powershell
Connect-PnPOnline -Url $Global:UhvSiteUrl -DeviceLogin -ClientId $env:UHV_CLIENT_ID -Tenant $env:UHV_TENANT -PersistLogin
Get-PnPWeb | Select-Object Title, Url
```

## Fast Deploy Using Profile (Tenant Catalog)

```powershell
.\scripts\Deploy-UHV-All.ps1 `
  -ClientId $env:UHV_CLIENT_ID `
  -Tenant $env:UHV_TENANT `
  -AppCatalogUrl $Global:UhvAppCatalogUrl `
  -TenantAdminUrl $Global:UhvTenantAdminUrl `
  -DeviceLogin
```

## Site Catalog Variant

```powershell
.\scripts\Deploy-UHV-Wrapper.ps1 `
  -AppCatalogUrl $Global:UhvSiteUrl `
  -Scope Site `
  -ClientId $env:UHV_CLIENT_ID `
  -Tenant $env:UHV_TENANT `
  -DeviceLogin
```

## Decision Tree

1. Deploying for multiple sites in tenant?
  - Yes -> use `Deploy-UHV-All.ps1` with tenant app catalog.
  - No -> consider site app catalog path.
2. Device login prompts every run?
  - Ensure `-PersistLogin` is used and token cache is not being cleared.
3. Script says `ClientId is required` or `Tenant is required`?
  - Confirm profile loaded in current session (`. .\ignore\UHV.LocalProfile.ps1`).
4. `Connect-SPOService` fails but PnP works?
  - Use PnP/CSOM fallback for site-level operations as documented in `README.md`.

## Expected Outputs

- Profile load:
  - no errors, env vars present
- PnP connect:
  - warning about token cache may appear
  - `Get-PnPWeb` returns site title/url
- Deploy:
  - `Deployment completed.`
  - site statuses: `UpdatedOrCurrent` or `Installed`

## Quick Self-Check

```powershell
"UHV_CLIENT_ID=$env:UHV_CLIENT_ID"
"UHV_TENANT=$env:UHV_TENANT"
"AppCatalog=$Global:UhvAppCatalogUrl"
"TenantAdmin=$Global:UhvTenantAdminUrl"
"SiteUrl=$Global:UhvSiteUrl"
```

## Rules

- Keep profile files inside `ignore/` (non-committed).
- Do not hardcode secrets in committed scripts/docs.
- Prefer `-PersistLogin` for smoother repeated admin operations.
- Keep one canonical profile file per machine to avoid drift.
