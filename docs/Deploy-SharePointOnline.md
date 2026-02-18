# Deploy UniversalHtmlViewer To SharePoint Online

This repo contains a SharePoint Framework (SPFx) package (`.sppkg`) and PowerShell scripts to build and deploy it.

## Prerequisites

- Permission to upload/publish apps to the target App Catalog (site app catalog or tenant app catalog).
- PowerShell 7+ recommended.
- PnP.PowerShell:

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```

Notes:
- You do not need a global Node.js installation if you use `scripts/Build-UHV.ps1`. It will use a local, known-good Node runtime under `.tools/`.
- If you do have Node installed globally, SPFx in this repo expects Node 22.14+ (Node 24+ is not supported by this project).
- PnP.PowerShell 3.x requires an Entra ID App Registration `ClientId` for `Connect-PnPOnline -Interactive` and `-DeviceLogin`.
- For repeat packaging runs, use `.\scripts\Build-UHV.ps1 -SkipInstall` to avoid reinstalling npm packages each time.
- To reduce npm warning noise during install, use `.\scripts\Build-UHV.ps1 -QuietNpm`.
- SharePoint Online does not require cryptographic signing of `.sppkg`; trust is managed through App Catalog governance and permissions.

## Evotec Quick Start (Your Test Tenant)

- Tenant domain: `evotecpoland.sharepoint.com`
- Tenant admin URL: `https://evotecpoland-admin.sharepoint.com`
- Candidate tenant app catalog URL: `https://evotecpoland.sharepoint.com/sites/appcatalog` (verify via SharePoint Admin Center or `Get-PnPTenantAppCatalogUrl`)
- Entra app registration (created): `UniversalHtmlViewer Deploy`
- ClientId: `f34fe56f-d9e7-4e0e-bd04-d62a0cdb2c1c`

Suggested session variables:

```powershell
$clientId = "f34fe56f-d9e7-4e0e-bd04-d62a0cdb2c1c"
$tenant = "evotecpoland.onmicrosoft.com"   # if your initial domain differs, replace this value
$appCatalogUrl = "https://evotecpoland.sharepoint.com/sites/appcatalog"
$tenantAdminUrl = "https://evotecpoland-admin.sharepoint.com"
```

## Step 0 (one-time): Create A ClientId For PnP Authentication

This creates an Entra ID app registration suitable for PnP interactive/device login and prints its `ClientId` (Application Id).

Pick your tenant identifier:
- Preferred: your initial domain like `yourtenant.onmicrosoft.com`
- Or: the tenant GUID

Example (Device Login is the most reliable):

```powershell
Register-PnPEntraIDAppForInteractiveLogin `
  -ApplicationName "UniversalHtmlViewer Deploy" `
  -Tenant "evotecpoland.onmicrosoft.com" `
  -DeviceLogin `
  -SharePointDelegatePermissions "AllSites.FullControl"
```

App UniversalHtmlViewer Deploy with id f34fe56f-d9e7-4e0e-bd04-d62a0cdb2c1c created.

Keep the returned `ClientId` and `Tenant` values; you will pass them to the deploy script.

### What permissions do we need?

For this repo's deployment flow (`Add-PnPApp` + `Publish-PnPApp` to an App Catalog), the simplest working permission is:
- SharePoint Delegated: `AllSites.FullControl`

This is a broad delegated permission (the signed-in user still needs the appropriate SharePoint/App Catalog roles). If your tenant has stricter policies, ask an M365/Entra admin to create the app registration and grant only the permissions your organization allows.

### Does Register-PnPEntraIDAppForInteractiveLogin require Connect-PnPOnline first?

No. `Register-PnPEntraIDAppForInteractiveLogin` performs its own sign-in flow (browser/device code) to create the app registration.
It does require that your account is allowed to create app registrations in Entra ID (many tenants restrict this to admins).

## Step 1: Determine Your App Catalog URL

`scripts/Deploy-UHV.ps1` needs the URL of the site that hosts the App Catalog.

Common examples (verify in your tenant):
- Tenant app catalog site: `https://<tenant>.sharepoint.com/sites/appcatalog`
- Site app catalog: the URL of the site collection where the "Site Collection App Catalog" feature is enabled

For your tenant, the URL will typically start with:
- `https://evotecpoland.sharepoint.com/`

How to verify you have the right URL:
- Open the URL in a browser.
- You should see an App Catalog site with an app library (often "Apps for SharePoint").

Optional PowerShell check (requires SharePoint admin access):

```powershell
Connect-PnPOnline -Url "https://evotecpoland-admin.sharepoint.com" -DeviceLogin -ClientId "<client-guid>" -Tenant "<yourtenant>.onmicrosoft.com"
Get-PnPTenantAppCatalogUrl
```

If `Get-PnPTenantAppCatalogUrl` returns blank, no tenant app catalog is configured yet.
Create one (admin permissions required):

```powershell
# Find timezone id, e.g. for Warsaw
Get-PnPTimeZoneId -Match "Warsaw"

# Create and register tenant app catalog
Register-PnPAppCatalogSite -Url "https://evotecpoland.sharepoint.com/sites/appcatalog" -Owner "<admin@tenant>" -TimeZoneId <id>
```

If the site already exists and just needs registration:

```powershell
Set-PnPTenantAppCatalogUrl -Url "https://evotecpoland.sharepoint.com/sites/appcatalog"
```

## Step 2: Build The SPFx Package (.sppkg)

From the repo root:

```powershell
.\scripts\Build-UHV.ps1
```

Optional fast/quiet variants:

```powershell
.\scripts\Build-UHV.ps1 -SkipInstall
.\scripts\Build-UHV.ps1 -QuietNpm
```

Expected output package:

```text
spfx/UniversalHtmlViewer/sharepoint/solution/universal-html-viewer.sppkg
```

## Step 3: Deploy

### Option A: Site-scoped (recommended for isolated testing)

Uploads and publishes to a site app catalog.

```powershell
.\scripts\Deploy-UHV.ps1 -AppCatalogUrl "<site-app-catalog-site-url>" -Scope Site -DeviceLogin -ClientId "<client-guid>" -Tenant "<tenant>.onmicrosoft.com"
```

After publishing, go to the target site and add/install the app as needed so the web part appears in the web part picker.

### Option B: Tenant app catalog (tenant-wide optional)

Uploads and publishes to the tenant app catalog.

Tenant-wide (Skip Feature Deployment):

```powershell
.\scripts\Deploy-UHV.ps1 -AppCatalogUrl "<tenant-app-catalog-site-url>" -Scope Tenant -TenantWide -DeviceLogin -ClientId "<client-guid>" -Tenant "<tenant>.onmicrosoft.com"
```

Example (common pattern, verify the actual app catalog site in your tenant first):

```powershell
.\scripts\Deploy-UHV.ps1 -AppCatalogUrl "https://evotecpoland.sharepoint.com/sites/appcatalog" -Scope Tenant -TenantWide -DeviceLogin -ClientId "<client-guid>" -Tenant "<yourtenant>.onmicrosoft.com"
```

If you do not use `-TenantWide`, the solution is still published to the tenant app catalog, but typically must be installed per-site before the web part is available there.

Important:
- For `-Scope Tenant`, the script connects to the tenant admin URL (`https://<tenant>-admin.sharepoint.com`) to avoid the PnP device-login context-switch limitation.
- For `-Scope Site`, it connects directly to `-AppCatalogUrl`.
- The deploy script uses `-Force` for app catalog operations, so you should not get repeated no-script confirmation prompts.

Tenant-wide note for this project:
- Current file `spfx/UniversalHtmlViewer/config/package-solution.json` has `"skipFeatureDeployment": false`.
- That means true tenant-wide rollout (`-SkipFeatureDeployment`) is not allowed by the package.
- If you pass `-TenantWide`, the script now falls back to normal publish and prints a warning.
- Result: solution is published to tenant app catalog, but each site may still need app installation.

Evotec command (skip rebuild):

```powershell
.\scripts\Deploy-UHV-Wrapper.ps1 -AppCatalogUrl $appCatalogUrl -Scope Tenant -TenantWide -DeviceLogin -ClientId $clientId -Tenant $tenant -TenantAdminUrl $tenantAdminUrl -SkipBuild
```

Evotec command (recommended with current package, no tenant-wide skip-feature deploy):

```powershell
.\scripts\Deploy-UHV-Wrapper.ps1 -AppCatalogUrl $appCatalogUrl -Scope Tenant -DeviceLogin -ClientId $clientId -Tenant $tenant -TenantAdminUrl $tenantAdminUrl -SkipBuild
```

## One-command Wrapper (deploy-uhv)

If you prefer a single command that builds and deploys:

```powershell
.\scripts\Deploy-UHV-Wrapper.ps1 -AppCatalogUrl "<app-catalog-site-url>" -DeviceLogin -ClientId "<client-guid>" -Tenant "<tenant>.onmicrosoft.com"
```

Tenant-wide:

```powershell
.\scripts\Deploy-UHV-Wrapper.ps1 -AppCatalogUrl "<tenant-app-catalog-site-url>" -TenantWide -DeviceLogin -ClientId "<client-guid>" -Tenant "<tenant>.onmicrosoft.com"
```

Build only (no SharePoint login):

```powershell
.\scripts\Deploy-UHV-Wrapper.ps1 -AppCatalogUrl "https://example.invalid" -NoDeploy
```

Windows shortcut command (same wrapper):

```powershell
.\scripts\deploy-uhv.cmd -AppCatalogUrl "<app-catalog-site-url>"
```

## One-command Site Onboarding

Install/update UHV on a site and create a configured dashboard page in one run:

```powershell
.\scripts\Setup-UHVSite.ps1 `
  -SiteUrl "https://contoso.sharepoint.com/sites/Reports" `
  -SiteRelativeDashboardPath "SiteAssets/Index.html" `
  -PageName "Dashboard" `
  -PageTitle "Dashboard" `
  -ConfigurationPreset "SharePointLibraryRelaxed" `
  -ContentDeliveryMode "SharePointFileContent" `
  -ClientId "<client-guid>" `
  -Tenant "<tenant>.onmicrosoft.com" `
  -DeviceLogin
```

Install/update app only (no page creation):

```powershell
.\scripts\Setup-UHVSite.ps1 `
  -SiteUrl "https://contoso.sharepoint.com/sites/Reports" `
  -InstallOnly `
  -ClientId "<client-guid>" `
  -Tenant "<tenant>.onmicrosoft.com" `
  -DeviceLogin
```

## Step 3.6: Create A UHV Page Directly From PowerShell (Optional)

If SharePoint page editing is unstable in the browser, create and configure the page via script:

```powershell
.\scripts\Add-UHVPage.ps1 `
  -SiteUrl "https://evotecpoland.sharepoint.com/sites/TestUHV1" `
  -PageName "Dashboard" `
  -PageTitle "Dashboard" `
  -PageLayoutType "Article" `
  -FullUrl "https://evotecpoland.sharepoint.com/sites/TestUHV1/SiteAssets/Index.html" `
  -ConfigurationPreset "SharePointLibraryFullPage" `
  -ContentDeliveryMode "SharePointFileContent" `
  -Publish `
  -ClientId $clientId `
  -Tenant $tenant `
  -DeviceLogin
```

Useful switches:
- `-SetAsHomePage` sets the created page as the site home page.
- `-ForceOverwrite` deletes an existing page with the same name and recreates it.
- `-EnsureSitePagesForceCheckout` applies the Site Pages workaround (`ForceCheckout = true`) before creating the page.
- `-SkipEnsureUhvAppOnSite` skips automatic app install/update check before adding the web part.
- `-PageLayoutType` defaults to `Article` (recommended for stability); supports `SingleWebPartAppPage` and `Home`.
- `-SkipAddWebPart` creates the page only (useful for isolating SharePoint page-authoring issues from web part issues).
- `-SkipConfigureWebPartProperties` adds UHV but leaves default properties untouched (useful when isolating page schema/write issues).

## Step 3.7: Upgrade Existing Sites After Publishing A New Package

After you upload/publish a new UHV package in App Catalog, you do not need to reupload per site.
You may still need to update the installed app instance on each site:

```powershell
.\scripts\Update-UHVSiteApp.ps1 `
  -SiteUrls @(
    "https://evotecpoland.sharepoint.com/sites/TestUHV1",
    "https://evotecpoland.sharepoint.com/sites/TestUHV2"
  ) `
  -InstallIfMissing `
  -ClientId $clientId `
  -Tenant $tenant `
  -DeviceLogin
```

This script:
- Finds UHV in app catalog (tenant first, then site catalog fallback).
- Installs it if missing on each site (`-InstallIfMissing`).
- Runs update so each site picks up the latest published package.

## Step 3.5: Install On The Target Site (Required When Not Tenant-Wide)

After publish to tenant app catalog, open the target site and ensure the app is installed there:

1. Go to `Site contents` on the target site.
2. If you do not see `universal-html-viewer-client-side-solution`, choose `Add an app` and install it from `From Your Organization`.
3. Refresh the page editor and add the `Universal HTML Viewer` web part.

## Step 4: Configure UHV For SharePoint-Hosted HTML

For enterprise tenants where direct `.html` iframe loading is blocked or downloaded:

1. Upload your HTML file to SharePoint (for example `Site Assets/Index.html`).
2. In the Universal HTML Viewer web part settings:
   - Set `Content delivery mode` to `SharePoint file API (inline iframe)`.
   - Set source mode/URL to the target HTML file.
   - For better fit on report bundles: set `Height mode` to `Auto` and enable `Fit content to width (inline mode)`.
   - Optional for edge-to-edge layout: set `Configuration preset` to `SharePoint library (full page)`.
   - Keep URL same-tenant (or site-relative).

This mode fetches file content via SharePoint REST and renders it inside the iframe using `srcdoc`, avoiding direct file-response header issues.
It also supports wrapper dashboards that contain inner iframes pointing to same-tenant `.html/.htm/.aspx` files (common for TheDashboard-style outputs).
Nested iframe targets that change at runtime (`iframe.src = ...`) are also auto-resolved in this mode.

For report bundles with many linked files:
- Upload the whole folder tree (not just `index.html`) to the same SharePoint library/folder.
- Keep links relative (for example `./details/report1.html`, `../index.html`).
- Keep links in the same tenant and same allowed path prefix.
- Use `.html`/`.htm`/`.aspx` links for in-place navigation in this mode.
- If a linked `.aspx` page must be server-rendered by SharePoint (not loaded as file text), use `Direct URL in iframe` mode for that scenario.

## Troubleshooting

- "PnP.PowerShell module not found": run the `Install-Module` command above.
- "Access denied" or app catalog errors: the account used in the interactive login must have App Catalog permissions.
- Package missing: run `.\scripts\Build-UHV.ps1` first.
- `Add-PnPApp : Value cannot be null. (Parameter 'webFullUrl')`: tenant app catalog is not configured. Run `Get-PnPTenantAppCatalogUrl`; if blank, register/set the tenant app catalog URL as shown above.
- `Publish-PnPApp ... Package does not have SkipFeatureDeployment set to true`: this package cannot be tenant-wide deployed in its current configuration (`skipFeatureDeployment=false`). Publish without tenant-wide and install per site, or change package config and rebuild.
- "Do we need scripting?": App Catalog operations may temporarily require toggling no-script on the catalog site. This does not mean you need to enable custom scripting across normal SharePoint sites.
- Direct file URL opens as download / iframe times out: use UHV `Content delivery mode = SharePoint file API (inline iframe)`.
- Linked HTML pages do not open inside the web part: use `Content delivery mode = SharePoint file API (inline iframe)`, keep links relative to the same library/folder tree, and prefer preset `SharePointLibraryRelaxed` (strict sandbox can prevent inline link interception).
- Wrapper pages with inner iframes do not load: use `SharePoint file API (inline iframe)` plus preset `SharePointLibraryRelaxed` or `SharePointLibraryFullPage` (strict sandbox blocks same-origin API access for nested iframe hydration).
- Browser console shows `about:srcdoc` inline-script CSP warnings: in many tenants this is currently report-only telemetry. If enforcement is enabled, move inline scripts to external `.js` files and reference them with `<script src="...">`.
- "Custom script is being retired / resets every 24h": UHV does not require enabling custom script on dashboard sites when using `Content delivery mode = SharePoint file API (inline iframe)`. Keep `DenyAddAndCustomizePages` enabled on business sites; only App Catalog maintenance may temporarily toggle no-script on the catalog site.
- `SavePageCoAuth` returns `400 Bad Request` and page edit exits after a few seconds: this is a SharePoint authoring/coauthoring issue (not UHV). Temporary workaround per affected site:

```powershell
Connect-PnPOnline -Url "https://<tenant>.sharepoint.com/sites/<site>" -DeviceLogin -ClientId "<client-guid>" -Tenant "<tenant>.onmicrosoft.com"
Set-PnPList -Identity "Site Pages" -ForceCheckout:$true
```

If multiple test sites are affected, apply the same setting per site collection.
- `Can't edit this page` with message about `com.fluidframework.leaf.string` on a script-created page: recreate that page with the latest `scripts/Add-UHVPage.ps1` and `-ForceOverwrite` (script uses safer `Set-PnPPageWebPart -PropertiesJson` flow).

## Rollback Procedure

Use `scripts/Rollback-UHV.ps1` to publish an older known-good package and then update target site app instances:

```powershell
.\scripts\Rollback-UHV.ps1 `
  -AppCatalogUrl "https://contoso.sharepoint.com/sites/appcatalog" `
  -RollbackSppkgPath "C:\Releases\universal-html-viewer-1.0.11.sppkg" `
  -Scope Tenant `
  -SiteUrls @(
    "https://contoso.sharepoint.com/sites/Reports",
    "https://contoso.sharepoint.com/sites/Operations"
  ) `
  -AppCatalogScope Tenant `
  -InstallIfMissing `
  -ClientId "<client-guid>" `
  -Tenant "<tenant>.onmicrosoft.com" `
  -DeviceLogin `
  -TenantAdminUrl "https://contoso-admin.sharepoint.com"
```

If you only need to republish the older package without site updates, add `-SkipSiteUpdate`.
- If the same error persists: run once with `-SkipAddWebPart` and confirm the page can be edited; then rerun without `-SkipAddWebPart` to isolate whether corruption occurs in page creation or web part insertion.
- If page creation works but adding/configuring UHV fails: rerun with `-SkipConfigureWebPartProperties`; if that page is editable, the issue is property-write payload and you can configure UHV in UI afterward.
