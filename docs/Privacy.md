# Universal HTML Viewer Privacy Statement

Universal HTML Viewer (UHV) is a client-side SharePoint Framework web part. It runs inside the Microsoft 365 tenant where it is installed and uses the current user's SharePoint permissions to render HTML content selected by a site owner.

## Data UHV Processes

UHV may process the following data inside the user's browser and SharePoint page context:

- Web part configuration, such as the selected HTML file URL, source mode, cache mode, display settings, and security options.
- HTML files and supporting assets that a site owner configures UHV to render from SharePoint or from explicitly allowed URLs.
- Optional tenant configuration loaded from a SharePoint URL configured by the tenant or site owner.
- Browser URL state used for report navigation, such as `uhvPage` query-string values or same-page hash fragments.

UHV does not include an Evotec Services-hosted service and does not send tenant content, configuration, telemetry, or usage analytics to Evotec Services sp. z o.o.

## Storage

UHV stores its web part configuration in SharePoint page/web part properties, using the normal Microsoft 365 storage model. Runtime HTML caching is in-memory in the browser and is used only to reduce repeat SharePoint reads during the current page session.

## Network Requests

By default, UHV reads configured HTML content through SharePoint APIs in the same tenant context. Site owners can enable more permissive modes, such as external HTTPS URLs or external script inlining, only through UHV security settings and allowed-host configuration.

Any third-party HTML, scripts, images, styles, or embedded content rendered through UHV are controlled by the tenant/site owner who configured that content. Those third-party resources may have their own privacy behavior.

## Administrator Responsibilities

Tenant and site administrators are responsible for:

- Choosing trusted HTML sources.
- Configuring allowed hosts, path prefixes, and file extensions.
- Reviewing any third-party resources included by the hosted HTML content.
- Applying Microsoft 365, SharePoint, and organizational data-governance policies.

## Contact

For product questions or issue reports, use the public GitHub repository:

https://github.com/EvotecIT/UltimateHtmlViewer/issues

Last updated: 2026-06-29
