# Skill: UHV Release Deploy

## Purpose

Ship a fresh UHV package and update target SharePoint sites with a repeatable flow.

## Inputs

- Tenant app/client values:
  - `ClientId`
  - `Tenant` (for example: `contoso.onmicrosoft.com`)
  - `AppCatalogUrl`
  - `TenantAdminUrl`
- Target sites list (for example: `TestUHV1`, `TestUHV2`, root site)

## Standard Flow

1. Build/package:

```powershell
.\scripts\Build-UHV.ps1 -QuietNpm
```

2. Deploy and update in one pass:

```powershell
.\scripts\Deploy-UHV-All.ps1 `
  -ClientId "<client-guid>" `
  -Tenant "<tenant>.onmicrosoft.com" `
  -AppCatalogUrl "https://<tenant>.sharepoint.com/sites/appcatalog" `
  -TenantAdminUrl "https://<tenant>-admin.sharepoint.com" `
  -SiteUrls @(
    "https://<tenant>.sharepoint.com/sites/TestUHV1",
    "https://<tenant>.sharepoint.com/sites/TestUHV2",
    "https://<tenant>.sharepoint.com"
  ) `
  -DeviceLogin
```

3. Verify installed version:

```powershell
$clientId = "<client-guid>"
$tenant = "<tenant>.onmicrosoft.com"
$sites = @(
  "https://<tenant>.sharepoint.com/sites/TestUHV1",
  "https://<tenant>.sharepoint.com/sites/TestUHV2",
  "https://<tenant>.sharepoint.com"
)

foreach ($site in $sites) {
  Connect-PnPOnline -Url $site -DeviceLogin -ClientId $clientId -Tenant $tenant -PersistLogin | Out-Null
  Get-PnPApp -Scope Tenant |
    Where-Object { $_.Title -like "*Universal HTML Viewer*" } |
    Select-Object Title, InstalledVersion, Deployed
}
```

## If Visual Changes Do Not Appear

1. Bump versions:
  - `spfx/UniversalHtmlViewer/config/package-solution.json` -> `solution.version`
  - `spfx/UniversalHtmlViewer/src/webparts/universalHtmlViewer/UniversalHtmlViewerWebPart.manifest.json` -> `version`
2. Rebuild and redeploy.
3. Hard refresh page (`Ctrl+F5`) and reopen web part picker.

## Known Pitfalls

- `npm ci` lock mismatch:
  - Ensure lockfile was generated with the same local toolchain used by `Build-UHV.ps1`.
- SPFx ship parse errors:
  - Keep webpack override aligned with SPFx build expectations (`spfx/UniversalHtmlViewer/package.json`).
- Sharing links in `/:u:/r/...` format:
  - Use canonical file URL for `SharePointFileContent` mode.

