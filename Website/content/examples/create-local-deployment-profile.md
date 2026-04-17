---
title: "Create a local deployment profile"
description: "Create a local-only profile for tenant deployment values."
layout: docs
---

This pattern is useful when you deploy repeatedly but do not want tenant values committed to source control.

It is adapted from `scripts/examples/UHV.LocalProfile.example.ps1`.

## Example

```powershell
Copy-Item .\scripts\examples\UHV.LocalProfile.example.ps1 .\ignore\UHV.LocalProfile.ps1
. .\ignore\UHV.LocalProfile.ps1

$env:UHV_CLIENT_ID = '<entra-app-client-id-guid>'
$env:UHV_TENANT = '<tenant>.onmicrosoft.com'

$Global:UhvAppCatalogUrl = 'https://<tenant>.sharepoint.com/sites/appcatalog'
$Global:UhvTenantAdminUrl = 'https://<tenant>-admin.sharepoint.com'
$Global:UhvSiteUrl = 'https://<tenant>.sharepoint.com/sites/Reports'
```

## What this demonstrates

- keeping tenant values in `ignore/`
- using placeholders in public documentation
- preparing values used by the deployment scripts

## Source

- [UHV.LocalProfile.example.ps1](https://github.com/EvotecIT/UltimateHtmlViewer/blob/master/scripts/examples/UHV.LocalProfile.example.ps1)

