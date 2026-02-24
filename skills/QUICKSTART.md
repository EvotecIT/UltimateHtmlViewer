# UHV Skills Quickstart

Use this when you need to move fast and avoid re-learning the repo.

## 60-Second Flow

1. Load tenant profile
  - Open: `skills/uhv-tenant-profile-setup/SKILL.md`
  - Goal: confirm auth/env is ready (`UHV_CLIENT_ID`, `UHV_TENANT`)
2. Deploy/update
  - Open: `skills/uhv-release-deploy/SKILL.md`
  - Goal: build, deploy, update sites, verify version
3. Debug only if needed
  - Open: `skills/uhv-sharepoint-debug/SKILL.md`
  - Goal: resolve download/deep-link/scroll issues with decision tree

## Common Command Sequence

```powershell
. .\ignore\UHV.LocalProfile.ps1
.\scripts\Build-UHV.ps1 -QuietNpm
.\scripts\Deploy-UHV-All.ps1 `
  -ClientId $env:UHV_CLIENT_ID `
  -Tenant $env:UHV_TENANT `
  -AppCatalogUrl $Global:UhvAppCatalogUrl `
  -TenantAdminUrl $Global:UhvTenantAdminUrl `
  -DeviceLogin
```

## Done Criteria

- Package build succeeds.
- Deploy reports `Deployment completed.`.
- Site updates show `UpdatedOrCurrent` or `Installed`.
- Installed version is correct on target sites.

