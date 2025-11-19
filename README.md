# UniversalHtmlViewer SPFx Web Part

The **UniversalHtmlViewer** web part is a SharePoint Framework (SPFx) client-side web part that renders a single `<iframe>` pointing to a static HTML page stored in SharePoint Online. It is designed to be generic and reusable across a tenant for hosting dashboards, reports, and other generated HTML pages.

## HTML source modes

The web part supports three ways to define the iframe source URL, controlled by the **HTML source mode** property.

### 1. FullUrl

- Property: `fullUrl`
- The iframe source is used exactly as provided (after validation).
- Example: `https://contoso.sharepoint.com/sites/Reports/Dashboards/system1/index.html`

### 2. BasePathAndRelativePath

- Properties:
  - `basePath` – site-relative base path, e.g. `/sites/Reports/Dashboards/`
  - `relativePath` – relative path appended to the base, e.g. `system1/index.html`
- The final URL is calculated as:

```text
finalUrl = basePath + relativePath
```

### 3. BasePathAndDashboardId

- Properties:
  - `basePath` – site-relative base path, e.g. `/sites/Reports/Dashboards/`
  - `dashboardId` – fallback dashboard ID
  - `defaultFileName` – file name within the dashboard folder (defaults to `index.html` when empty)
  - `queryStringParamName` – name of the query string parameter used to resolve the dashboard ID (defaults to `dashboard` when empty)
- Resolution order for the effective dashboard ID:
  1. Value from the current page URL query string parameter `queryStringParamName`, if present.
  2. The `dashboardId` property as a fallback.
- The final URL is calculated as:

```text
finalUrl = basePath + dashboardId + '/' + defaultFileName
```

If no dashboard ID can be determined, the web part renders a friendly message instead of an iframe.

## Height modes

The web part supports two height modes for the iframe:

### Fixed

- Property: `heightMode = "Fixed"`
- Property: `fixedHeightPx` (default: `800`)
- The iframe height is set to the specified number of pixels, e.g. `height: 800px;`.

### Viewport

- Property: `heightMode = "Viewport"`
- The iframe height is set to fill the viewport: `height: 100vh;`.

## Property pane configuration

The web part exposes the following properties in the property pane:

- Group: **Source settings**
  - `htmlSourceMode` (dropdown) – selects one of `FullUrl`, `BasePathAndRelativePath`, or `BasePathAndDashboardId`.
  - `fullUrl` – used only when `htmlSourceMode = FullUrl`.
  - `basePath` – used when `htmlSourceMode` is not `FullUrl`.
  - `relativePath` – used when `htmlSourceMode = BasePathAndRelativePath`.
  - `dashboardId`, `defaultFileName`, `queryStringParamName` – used when `htmlSourceMode = BasePathAndDashboardId`.
- Group: **Layout**
  - `heightMode` (dropdown) – `Fixed` or `Viewport`.
  - `fixedHeightPx` – numeric value used when `heightMode = Fixed`.

## URL safety and validation

Before rendering the iframe, the web part:

- Rejects URLs using the `javascript:` scheme or any scheme other than `http`, `https`, or a site-relative path starting with `/`.
- When the URL is absolute (`http` or `https`), it ensures the host matches the current SharePoint page host (same tenant).
- If validation fails, the web part renders a clear error message:

> UniversalHtmlViewer: The target URL is invalid or not allowed.

If no URL can be computed, a friendly message is shown:

> UniversalHtmlViewer: No URL configured. Please update the web part settings.

## Styling

The web part uses minimal styling to keep the iframe clean:

- The container takes the full available width.
- The iframe:
  - Has no border.
  - Uses `width: 100%`.
  - Uses either a fixed pixel height or `100vh` depending on the configured height mode.

## Limitations

- Only same-tenant URLs are allowed when using absolute URLs.
- `javascript:` URLs and non-HTTP(S) schemes (e.g. `ftp:`) are blocked.
- Paths must either be absolute URLs on the same tenant or site-relative paths starting with `/`.

## Running tests

Unit tests for URL computation and validation logic are located in:

- `src/webparts/universalHtmlViewer/__tests__/UrlHelper.test.ts`

Jest configuration is defined in `jest.config.js`. To run the tests, ensure Jest and `ts-jest` are installed, then run:

```bash
npm test
```

