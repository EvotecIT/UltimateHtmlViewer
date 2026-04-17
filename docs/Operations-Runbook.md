# UHV Operations Runbook (Reusable)

This runbook is designed for reuse across tenants/sites without hardcoded organization values.

## 1) Optional local profile (recommended)

Create your local, non-committed profile once:

```powershell
Copy-Item .\scripts\examples\UHV.LocalProfile.example.ps1 .\ignore\UHV.LocalProfile.ps1
notepad .\ignore\UHV.LocalProfile.ps1
```

Load profile in your session:

```powershell
. .\ignore\UHV.LocalProfile.ps1
```

After loading, scripts can use:

- `$env:UHV_CLIENT_ID`
- `$env:UHV_TENANT`

## 2) Build package

```powershell
.\scripts\Build-UHV.ps1 -SkipInstall -QuietNpm
```

## 3) Deploy `.sppkg` to app catalog (tenant or site scope)

Tenant scope:

```powershell
.\scripts\Deploy-UHV-Wrapper.ps1 `
  -AppCatalogUrl $UhvAppCatalogUrl `
  -Scope Tenant `
  -DeviceLogin `
  -TenantAdminUrl $UhvTenantAdminUrl
```

Site scope:

```powershell
.\scripts\Deploy-UHV-Wrapper.ps1 `
  -AppCatalogUrl "https://<tenant>.sharepoint.com/sites/<site>" `
  -Scope Site `
  -DeviceLogin
```

Note: `-ClientId` and `-Tenant` remain supported explicitly. If omitted, scripts read from `UHV_CLIENT_ID` / `UHV_TENANT`.

## 4) Install/update app on one or more sites

```powershell
.\scripts\Update-UHVSiteApp.ps1 `
  -SiteUrls @(
    "https://<tenant>.sharepoint.com/sites/SiteA",
    "https://<tenant>.sharepoint.com/sites/SiteB"
  ) `
  -InstallIfMissing `
  -DeviceLogin
```

## 5) One-command page provisioning (site URL first)

This is the quickest reusable path for most teams:

```powershell
.\scripts\Setup-UHVSite.ps1 `
  -SiteUrl "https://<tenant>.sharepoint.com/sites/Reports" `
  -SiteRelativeDashboardPath "SiteAssets/Index.html" `
  -PageName "Reports" `
  -PageTitle "Reports" `
  -ConfigurationPreset "SharePointLibraryRelaxed" `
  -ContentDeliveryMode "SharePointFileContent" `
  -DeviceLogin
```

Note: `-SiteRelativeDashboardPath` is a backward-compatible parameter name and can point to any HTML entry file.

## 6) Validate app install/version

```powershell
Connect-PnPOnline -Url "https://<tenant>.sharepoint.com/sites/<site>" -DeviceLogin -ClientId $env:UHV_CLIENT_ID -Tenant $env:UHV_TENANT
Get-PnPApp -Scope Site | Where-Object { $_.Title -like "*Universal HTML Viewer*" } | Select-Object Title, InstalledVersion, Deployed
```

## 7) Recommended web part runtime settings

- Content delivery mode: `SharePoint file API (inline iframe)`
- HTML source mode: `Full URL` (or chosen preset path mode)
- Sandbox preset: `Relaxed` (unless stricter governance is required)
- Use same-tenant URLs for linked report pages

## 8) One host page for many files

If you have dozens of report files under one SharePoint library/folder tree, do not create one UHV page per file.
Provision one host page, point it at a default entry file, and allow query-driven deep links inside the approved path boundary.

Example:

```powershell
.\scripts\Setup-UHVSite.ps1 `
  -SiteUrl "https://<tenant>.sharepoint.com/sites/Reports" `
  -PageName "Reports" `
  -PageTitle "Reports" `
  -HtmlSourceMode "BasePathAndRelativePath" `
  -BasePath "/sites/Reports/Shared Documents/Reports/" `
  -RelativePath "Global/LegacyUsers.aspx" `
  -AllowedPathPrefixes "/sites/Reports/Shared Documents/Reports/" `
  -AllowedFileExtensions ".html,.htm,.aspx" `
  -AllowQueryStringPageOverride `
  -ConfigurationPreset "SharePointLibraryRelaxed" `
  -ContentDeliveryMode "SharePointFileContent" `
  -DeviceLogin
```

Then share links like:

```text
https://<tenant>.sharepoint.com/sites/Reports/SitePages/Reports.aspx?uhvPage=%2Fsites%2FReports%2FShared%20Documents%2FReports%2FGlobal%2FLegacyUsers.aspx
```

Permission notes:

- UHV does not bypass SharePoint permissions.
- Users still need access to both the host page and the target file.
- In mixed-permission libraries, prefer one host page per permission boundary or top-level secured folder, not per individual file.

## 9) Rollback

```powershell
.\scripts\Rollback-UHV.ps1 `
  -RollbackSppkgPath ".\release\universal-html-viewer.previous.sppkg" `
  -AppCatalogUrl "https://<tenant>.sharepoint.com/sites/appcatalog" `
  -SiteUrls @("https://<tenant>.sharepoint.com/sites/Reports") `
  -Scope Tenant `
  -DeviceLogin `
  -TenantAdminUrl "https://<tenant>-admin.sharepoint.com"
```
