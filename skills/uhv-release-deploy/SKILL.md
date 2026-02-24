# Skill: UHV Release Deploy

## Purpose

Ship a fresh UHV package and update target SharePoint sites with a repeatable, low-risk flow.

## Inputs

- `ClientId`
- `Tenant` (for example: `contoso.onmicrosoft.com`)
- `AppCatalogUrl`
- `TenantAdminUrl`
- `SiteUrls` (for example: `TestUHV1`, `TestUHV2`, root site)

## Recommended Path (Tenant App Catalog)

```powershell
.\scripts\Build-UHV.ps1 -QuietNpm

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

## Exact Evotec Example

```powershell
.\scripts\Build-UHV.ps1 -QuietNpm

.\scripts\Deploy-UHV-All.ps1 `
  -ClientId "f34fe56f-d9e7-4e0e-bd04-d62a0cdb2c1c" `
  -Tenant "evotecpoland.onmicrosoft.com" `
  -AppCatalogUrl "https://evotecpoland.sharepoint.com/sites/appcatalog" `
  -TenantAdminUrl "https://evotecpoland-admin.sharepoint.com" `
  -SiteUrls @(
    "https://evotecpoland.sharepoint.com/sites/TestUHV1",
    "https://evotecpoland.sharepoint.com/sites/TestUHV2",
    "https://evotecpoland.sharepoint.com"
  ) `
  -DeviceLogin
```

## Site App Catalog Variant (Single Site)

Use this only when tenant app catalog is not the chosen path:

```powershell
$siteUrl = "https://<tenant>.sharepoint.com/sites/TestUHV2"

.\scripts\Deploy-UHV-Wrapper.ps1 `
  -AppCatalogUrl $siteUrl `
  -Scope Site `
  -ClientId "<client-guid>" `
  -Tenant "<tenant>.onmicrosoft.com" `
  -DeviceLogin `
  -SkipBuild

.\scripts\Update-UHVSiteApp.ps1 `
  -SiteUrls @($siteUrl) `
  -InstallIfMissing `
  -ClientId "<client-guid>" `
  -Tenant "<tenant>.onmicrosoft.com" `
  -DeviceLogin
```

## Decision Tree

1. Need fresh code in package?
  - Yes -> run `Build-UHV.ps1`.
  - No -> run deploy with `-SkipBuild`.
2. Tenant-wide distribution needed?
  - Yes -> use `Deploy-UHV-All.ps1` or `Deploy-UHV-Wrapper.ps1 -Scope Tenant`.
  - No -> use `Deploy-UHV-Wrapper.ps1 -Scope Site` on target site catalog.
3. App present on site but version old?
  - Run `Update-UHVSiteApp.ps1 -InstallIfMissing`.
4. Visuals unchanged after successful deploy?
  - Bump versions in both manifest and solution, rebuild, redeploy, hard refresh.

## Expected Outputs

- Build success:
  - `Package created: ...\sharepoint\solution\universal-html-viewer.sppkg`
- Deploy success:
  - `Deployment completed.`
- Site update success:
  - `Status = UpdatedOrCurrent` or `Status = Installed`
- Verification success:
  - `InstalledVersion = <expected>`
  - `Deployed = True`

## Verification Block

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
  [PSCustomObject]@{
    SiteUrl = $site
    App = (Get-PnPApp -Scope Tenant | Where-Object { $_.Title -like "*Universal HTML Viewer*" } | Select-Object -First 1).Title
    InstalledVersion = (Get-PnPApp -Scope Tenant | Where-Object { $_.Title -like "*Universal HTML Viewer*" } | Select-Object -First 1).InstalledVersion
  }
}
```

## Post-Deploy Validation Checklist

- Web part picker shows expected icon/branding.
- `SharePointFileContent` mode renders inline (no forced download).
- Deep links update `?uhvPage=` and back/forward works.
- Initial load does not fight host scroll.
- `Open in new tab` opens content endpoint as expected.

## Rollback Mini-Playbook

```powershell
.\scripts\Rollback-UHV.ps1 `
  -AppCatalogUrl "https://<tenant>.sharepoint.com/sites/appcatalog" `
  -PreviousPackagePath "<path-to-older-sppkg>" `
  -ClientId "<client-guid>" `
  -Tenant "<tenant>.onmicrosoft.com" `
  -DeviceLogin
```

Then run:

```powershell
.\scripts\Update-UHVSiteApp.ps1 `
  -SiteUrls @(
    "https://<tenant>.sharepoint.com/sites/TestUHV1",
    "https://<tenant>.sharepoint.com/sites/TestUHV2",
    "https://<tenant>.sharepoint.com"
  ) `
  -InstallIfMissing `
  -ClientId "<client-guid>" `
  -Tenant "<tenant>.onmicrosoft.com" `
  -DeviceLogin
```

## Known Pitfalls

- `npm ci` lock mismatch:
  - Regenerate lockfile using the same runtime used by `Build-UHV.ps1`.
- SPFx ship parse errors:
  - Keep webpack override aligned with SPFx expectations in `spfx/UniversalHtmlViewer/package.json`.
- Stale visuals after deploy:
  - Bump both `solution.version` and webpart `manifest version`.
- Share links in `/:u:/r/...` format:
  - Use canonical file URL for `SharePointFileContent`.
