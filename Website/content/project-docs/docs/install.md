---
title: "Install UltimateHtmlViewer"
description: "Build and deploy the UniversalHtmlViewer SPFx package."
layout: docs
---

Build the package from the repository:

```powershell
.\scripts\Build-UHV.ps1
```

Deploy it to a tenant or site app catalog using placeholder tenant values:

```powershell
.\scripts\Deploy-UHV-Wrapper.ps1 `
    -AppCatalogUrl 'https://<tenant>.sharepoint.com/sites/appcatalog' `
    -Scope Tenant `
    -DeviceLogin `
    -ClientId '<client-guid>' `
    -Tenant '<tenant>.onmicrosoft.com' `
    -TenantAdminUrl 'https://<tenant>-admin.sharepoint.com'
```

The generated `.sppkg` is an SPFx package and must be approved through the normal SharePoint app catalog process.

