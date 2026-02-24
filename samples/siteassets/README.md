# SiteAssets Demo Files

This folder contains ready-to-upload HTML files for polished UHV screenshots and demos.

## Included

- `UHV-Screenshot-Demo.html` - self-contained branded demo dashboard for screenshots.
- `UHV-Feature-Showcase.html` - visual "UHV capabilities" page for product-focused demos.

## Upload to SharePoint SiteAssets

```powershell
Connect-PnPOnline -Url "https://<tenant>.sharepoint.com/sites/<site>" -DeviceLogin -ClientId $env:UHV_CLIENT_ID -Tenant $env:UHV_TENANT
Add-PnPFile -Path ".\samples\siteassets\UHV-Screenshot-Demo.html" -Folder "SiteAssets" -Overwrite
Add-PnPFile -Path ".\samples\siteassets\UHV-Feature-Showcase.html" -Folder "SiteAssets" -Overwrite
```

## Use in UHV

Set UHV `Full URL` to:

```text
https://<tenant>.sharepoint.com/sites/<site>/SiteAssets/UHV-Screenshot-Demo.html
```

Or use:

```text
https://<tenant>.sharepoint.com/sites/<site>/SiteAssets/UHV-Feature-Showcase.html
```
