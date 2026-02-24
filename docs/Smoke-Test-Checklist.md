# UniversalHtmlViewer Smoke Test Checklist

Use this checklist after deploying the latest package to a test site.

## Build Artifact

- [ ] Confirm package exists: `spfx/UniversalHtmlViewer/sharepoint/solution/universal-html-viewer.sppkg`
- [ ] Confirm package version/date is the expected one for this run

## Test Setup

- [ ] Site page contains the UniversalHtmlViewer web part
- [ ] Web part is configured to:
  - `ContentDeliveryMode = SharePointFileContent`
  - `ConfigurationPreset = SharePointLibraryRelaxed` (or stricter as required)
  - source points to the root entry HTML file

## Navigation and Deep Links

- [ ] Open host page with no `uhvPage` query parameter
  - Expected: default entry page loads
- [ ] Navigate to 2-3 subpages via report menu links
  - Expected: page URL updates with `uhvPage=...`
- [ ] Copy current URL and open it in new browser tab
  - Expected: same subpage opens directly
- [ ] Use browser Back/Forward
  - Expected: iframe content follows history entries correctly

## Invalid Deep Link Handling

- [ ] Edit URL manually with invalid/forbidden `uhvPage`
  - Expected: clear error appears
  - Expected: `Reset to default` action is shown
- [ ] Click `Reset to default`
  - Expected: returns to default entry page and removes deep-link override

## Access and Request Access Flow

- [ ] Open a subpage where current user has access
  - Expected: content loads inline
- [ ] Open a subpage where current user does not have access (`401/403`)
  - Expected: access-denied style message appears
  - Expected: `Open file in SharePoint / Request access` action is shown
- [ ] Click `Open file in SharePoint / Request access`
  - Expected: SharePoint opens target file page
  - Expected: native `Request access` flow is available

## Download Regression Check

- [ ] Navigate between subpages multiple times
  - Expected: no unexpected file download prompts

## Content Selector (if enabled)

- [ ] Change content using selector in header
  - Expected: switches content without direct-file download behavior
  - Expected: URL deep-link updates and remains shareable

## Security Boundary Check

- [ ] Test relative links that stay inside allowed path prefix
  - Expected: they work
- [ ] Test links outside configured allowed path prefix
  - Expected: blocked (no unsafe bypass)

## Diagnostics

- [ ] Enable diagnostics temporarily if issue appears
  - Expected: status/error info is visible in diagnostics payload
- [ ] Disable diagnostics after testing

## Pass / Fail Summary

- [ ] PASS: deep-linking works
- [ ] PASS: back/forward works
- [ ] PASS: request-access route works
- [ ] PASS: no auto-download regression
- [ ] PASS: prefix security rules enforced

Notes:

- Build/package validated locally. Tenant validation is required for authentication/permissions behavior.
