# UniversalHtmlViewer - SharePoint Online Deployment Guide

This guide shows how to build, publish, install, update, and rollback UniversalHtmlViewer (UHV) in SharePoint Online using the repository scripts.

📦 Workflow Status

[![SPFx Tests](https://github.com/EvotecIT/UltimateHtmlViewer/actions/workflows/spfx-tests.yml/badge.svg)](https://github.com/EvotecIT/UltimateHtmlViewer/actions/workflows/spfx-tests.yml)
[![Release SPPKG](https://github.com/EvotecIT/UltimateHtmlViewer/actions/workflows/release-sppkg.yml/badge.svg)](https://github.com/EvotecIT/UltimateHtmlViewer/actions/workflows/release-sppkg.yml)

## Documentation map

- Product overview and capabilities: `README.md`
- Deployment and operations (this guide): `docs/Deploy-SharePointOnline.md`
- Reusable day-to-day runbook: `docs/Operations-Runbook.md`

## What this guide covers

- One-time authentication setup for PnP.PowerShell 3.x.
- App catalog discovery/registration.
- Build and deploy flows (site scope and tenant app catalog).
- Site onboarding and page provisioning scripts.
- Update and rollback operations.
- Troubleshooting for real-world SharePoint behavior.

UHV positioning:

- UHV is an SPFx app delivering a reusable web part.
- The web part can be added to any modern page and mixed with other web parts.
- Page names like `Dashboard.aspx` in examples are conventions, not requirements.

## Prerequisites

- PowerShell 7+ recommended.
- Permission to publish apps to your target app catalog.
- PnP.PowerShell installed:

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

Notes:

- `Build-UHV.ps1` can bootstrap a local compatible Node runtime in `.tools/` when global Node is unsupported.
- Supported Node for this project is `>=22.14.0 <23.0.0`.
- SharePoint Online does not require cryptographic signing of `.sppkg`; trust is controlled via App Catalog governance.

## Quick start variables

```powershell
$clientId = "<client-guid>"
$tenant = "<tenant>.onmicrosoft.com"      # or tenant GUID
$appCatalogUrl = "https://<tenant>.sharepoint.com/sites/appcatalog"
$tenantAdminUrl = "https://<tenant>-admin.sharepoint.com"
```

Auth reuse defaults:

- Scripted connections now default to persisted login (`-PersistLogin $true`) to reduce repeat device/interactive prompts between runs.
- Use `-PersistLogin:$false` for a non-persistent session.
- Use `-ForceAuthentication` to force a fresh login prompt.
- If `-ClientId` / `-Tenant` are omitted, scripts can read `UHV_CLIENT_ID` / `UHV_TENANT` environment variables.
- For local-only values, use `ignore/UHV.LocalProfile.ps1` (template: `scripts/examples/UHV.LocalProfile.example.ps1`).

## 1) One-time auth setup (ClientId for PnP.PowerShell)

PnP.PowerShell 3.x requires a registered Entra ID app (`ClientId`) for interactive/device login.

```powershell
Register-PnPEntraIDAppForInteractiveLogin `
  -ApplicationName "UniversalHtmlViewer Deploy" `
  -Tenant "<tenant>.onmicrosoft.com" `
  -DeviceLogin `
  -SharePointDelegatePermissions "AllSites.FullControl"
```

Keep the returned app id (`ClientId`) and tenant identifier.

Permission guidance:

- Simplest working delegated permission for this deployment flow is `AllSites.FullControl`.
- The signed-in user still needs proper SharePoint/App Catalog role permissions.

## 2) Find or configure app catalog

Check tenant app catalog URL:

```powershell
Connect-PnPOnline -Url "https://<tenant>-admin.sharepoint.com" -DeviceLogin -ClientId $clientId -Tenant $tenant
Get-PnPTenantAppCatalogUrl
```

If blank, create/register one:

```powershell
Get-PnPTimeZoneId -Match "Warsaw"
Register-PnPAppCatalogSite -Url "https://<tenant>.sharepoint.com/sites/appcatalog" -Owner "<admin@tenant>" -TimeZoneId <id>
```

If site already exists but is not linked:

```powershell
Set-PnPTenantAppCatalogUrl -Url "https://<tenant>.sharepoint.com/sites/appcatalog"
```

## 3) Build package

```powershell
.\scripts\Build-UHV.ps1
```

Useful build variants:

```powershell
.\scripts\Build-UHV.ps1 -SkipInstall
.\scripts\Build-UHV.ps1 -QuietNpm
```

Expected package:

```text
spfx/UniversalHtmlViewer/sharepoint/solution/universal-html-viewer.sppkg
```

## 4) Deploy package

### Option A: Site app catalog scope

```powershell
.\scripts\Deploy-UHV.ps1 `
  -AppCatalogUrl "<site-app-catalog-url>" `
  -Scope Site `
  -DeviceLogin `
  -ClientId $clientId `
  -Tenant $tenant
```

### Option B: Tenant app catalog publish

```powershell
.\scripts\Deploy-UHV.ps1 `
  -AppCatalogUrl $appCatalogUrl `
  -Scope Tenant `
  -DeviceLogin `
  -ClientId $clientId `
  -Tenant $tenant `
  -TenantAdminUrl $tenantAdminUrl
```

Important package behavior:

- `spfx/UniversalHtmlViewer/config/package-solution.json` currently has `"skipFeatureDeployment": false`.
- This means true tenant-wide skip-feature rollout is not supported by current package metadata.
- Publishing to tenant app catalog still works, but site-level app installation is typically required.

### One-command wrapper

Build + deploy:

```powershell
.\scripts\Deploy-UHV-Wrapper.ps1 -AppCatalogUrl $appCatalogUrl -Scope Tenant -DeviceLogin -ClientId $clientId -Tenant $tenant -TenantAdminUrl $tenantAdminUrl
```

Deploy only (skip build):

```powershell
.\scripts\Deploy-UHV-Wrapper.ps1 -AppCatalogUrl $appCatalogUrl -Scope Tenant -DeviceLogin -ClientId $clientId -Tenant $tenant -TenantAdminUrl $tenantAdminUrl -SkipBuild
```

Build only (no SharePoint login):

```powershell
.\scripts\Deploy-UHV-Wrapper.ps1 -AppCatalogUrl "https://example.invalid" -NoDeploy
```

## 5) Install app on target site

When not using true tenant-wide skip-feature deployment, install app on each target site:

1. Open `Site contents`.
2. Select `Add an app` and install `Universal HTML Viewer` from organization apps.
3. Refresh page editor and add `Universal HTML Viewer` web part.

## 6) Recommended UHV config for SharePoint-hosted HTML bundles

For HTML report/app bundles in `SiteAssets` or another SharePoint library:

- `Configuration preset`: `SharePointLibraryRelaxed`
- `Content delivery mode`: `SharePoint file API (inline iframe)`
- `HTML source mode`: `Full URL`
- `Full URL`: your `.../SiteAssets/Index.html` (or any entry HTML)
- `Height mode`: `Auto`
- `Fit content to width`: `On`

Why this works:

- Avoids direct `.html` iframe download/header behavior.
- Preserves linked-file navigation for same-tenant report bundles.
- Supports nested iframe hydration in wrapper pages.

## 7) Site onboarding and page provisioning

### Full onboarding in one command

```powershell
.\scripts\Setup-UHVSite.ps1 `
  -SiteUrl "https://<tenant>.sharepoint.com/sites/Reports" `
  -SiteRelativeDashboardPath "SiteAssets/Index.html" `
  -PageName "Reports" `
  -PageTitle "Reports" `
  -ConfigurationPreset "SharePointLibraryRelaxed" `
  -ContentDeliveryMode "SharePointFileContent" `
  -ClientId $clientId `
  -Tenant $tenant `
  -DeviceLogin
```

Note: parameter name `-SiteRelativeDashboardPath` is kept for backward compatibility; it can point to any HTML entry file.

### Add/configure page directly

```powershell
.\scripts\Add-UHVPage.ps1 `
  -SiteUrl "https://<tenant>.sharepoint.com/sites/Reports" `
  -PageName "Operations" `
  -PageTitle "Operations" `
  -PageLayoutType "Article" `
  -FullUrl "https://<tenant>.sharepoint.com/sites/Reports/SiteAssets/Index.html" `
  -ConfigurationPreset "SharePointLibraryFullPage" `
  -ContentDeliveryMode "SharePointFileContent" `
  -Publish `
  -ClientId $clientId `
  -Tenant $tenant `
  -DeviceLogin
```

Useful switches:

- `-ForceOverwrite`
- `-EnsureSitePagesForceCheckout`
- `-SkipAddWebPart`
- `-SkipConfigureWebPartProperties`
- `-SetAsHomePage`

## 8) Update existing sites after publishing new package

You do not reupload per site; you update site app instances:

```powershell
.\scripts\Update-UHVSiteApp.ps1 `
  -SiteUrls @(
    "https://<tenant>.sharepoint.com/sites/Reports",
    "https://<tenant>.sharepoint.com/sites/Operations"
  ) `
  -InstallIfMissing `
  -ClientId $clientId `
  -Tenant $tenant `
  -DeviceLogin
```

## 9) Rollback

```powershell
.\scripts\Rollback-UHV.ps1 `
  -AppCatalogUrl $appCatalogUrl `
  -RollbackSppkgPath "C:\Releases\universal-html-viewer-1.0.11.sppkg" `
  -Scope Tenant `
  -SiteUrls @(
    "https://<tenant>.sharepoint.com/sites/Reports",
    "https://<tenant>.sharepoint.com/sites/Operations"
  ) `
  -AppCatalogScope Tenant `
  -InstallIfMissing `
  -ClientId $clientId `
  -Tenant $tenant `
  -DeviceLogin `
  -TenantAdminUrl $tenantAdminUrl
```

Use `-SkipSiteUpdate` if you only need to republish the old package.

## Troubleshooting

- `Connect-PnPOnline` interactive/device prompts fail:
  ensure valid `ClientId` app registration and allowed tenant policy.
- `Add-PnPApp ... webFullUrl null`:
  tenant app catalog URL is not configured.
- `Package does not have SkipFeatureDeployment set to true`:
  current package metadata blocks true tenant-wide skip-feature deployment.
- Direct file URL downloads or iframe times out:
  switch to `SharePoint file API (inline iframe)` mode.
- Linked pages or wrapper iframes fail:
  use `SharePointLibraryRelaxed`/`SharePointLibraryFullPage`, keep links relative and same-tenant.
- `SavePageCoAuth 400` while editing pages:
  usually SharePoint page authoring issue, not UHV. Temporary workaround:

```powershell
Connect-PnPOnline -Url "https://<tenant>.sharepoint.com/sites/<site>" -DeviceLogin -ClientId $clientId -Tenant $tenant
Set-PnPList -Identity "Site Pages" -ForceCheckout:$true
```

- `Can't edit this page` + `com.fluidframework.leaf.string` on script-created page:
  recreate page with latest `Add-UHVPage.ps1 -ForceOverwrite`.

## Security and governance note

UHV is an SPFx app-catalog-governed deployment model.  
Enterprise control is achieved through:

- App Catalog approval/governance,
- admin/site-owner role boundaries,
- tenant policies and SharePoint permissions.
