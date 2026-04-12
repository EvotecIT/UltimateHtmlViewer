---
title: "Upload demo HTML to SiteAssets"
description: "Upload curated demo HTML files for a SharePoint-hosted UHV test."
layout: docs
---

This pattern is useful when you want a safe demo file before pointing UHV at real report output.

It is adapted from `samples/siteassets/README.md`.

## Example

```powershell
Import-Module PnP.PowerShell

Connect-PnPOnline `
    -Url 'https://<tenant>.sharepoint.com/sites/<site>' `
    -DeviceLogin `
    -ClientId $env:UHV_CLIENT_ID `
    -Tenant $env:UHV_TENANT

Add-PnPFile -Path '.\samples\siteassets\UHV-Screenshot-Demo.html' -Folder 'SiteAssets' -Overwrite
Add-PnPFile -Path '.\samples\siteassets\UHV-Feature-Showcase.html' -Folder 'SiteAssets' -Overwrite
```

Set UHV `Full URL` to:

```text
https://<tenant>.sharepoint.com/sites/<site>/SiteAssets/UHV-Screenshot-Demo.html
```

## What this demonstrates

- testing with curated demo HTML
- keeping tenant and site values as placeholders
- using the same sample files shown in the repository docs

## Source

- [samples/siteassets/README.md](https://github.com/EvotecIT/UltimateHtmlViewer/blob/master/samples/siteassets/README.md)

