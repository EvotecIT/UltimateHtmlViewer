# Contributor Instructions (Current)

This repository already contains the SPFx solution under:

- `spfx/UniversalHtmlViewer`

Use this file as a quick contributor map. For operations and deployment, use the docs listed below.

## Primary Docs

- Deployment guide: `docs/Deploy-SharePointOnline.md`
- Reusable operations runbook: `docs/Operations-Runbook.md`
- Product overview/configuration: `README.md`

## Common Flows

### Build package

```powershell
.\scripts\Build-UHV.ps1
```

### Deploy package to app catalog

```powershell
.\scripts\Deploy-UHV-Wrapper.ps1 -AppCatalogUrl "https://<tenant>.sharepoint.com/sites/appcatalog" -Scope Tenant -DeviceLogin
```

### One-command site setup

```powershell
.\scripts\Setup-UHVSite.ps1 -SiteUrl "https://<tenant>.sharepoint.com/sites/Reports" -DeviceLogin
```

## Local-Only Files

Use `ignore/` for non-committed local files (tenant values, private notes, temporary snippets).

- template: `scripts/examples/UHV.LocalProfile.example.ps1`
- local copy: `ignore/UHV.LocalProfile.ps1`

Scripts can read auth from env vars if parameters are omitted:

- `UHV_CLIENT_ID`
- `UHV_TENANT`

## Notes

- Avoid hardcoding tenant-specific URLs or identities in shared docs/scripts.
- Keep shared scripts parameterized and reusable across organizations.
