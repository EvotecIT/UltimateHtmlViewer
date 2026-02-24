# Skill: UHV SharePoint Debug

## Purpose

Diagnose and fix SharePoint-hosted UHV runtime issues quickly and consistently.

## Symptom -> First Checks

1. HTML triggers file download
  - Use `contentDeliveryMode = SharePointFileContent`.
  - Use canonical file URL, not `/:u:/r/...` share link.

2. Deep links do not persist or back/forward fails
  - Verify `?uhvPage=...` appears in host page URL.
  - Confirm links are supported extensions (`.html`, `.htm`, `.aspx` by default).

3. Page jumps/scroll fights during initial load
  - Reproduce with `?uhvTraceScroll=1`.
  - Validate host-scroll stabilization and nested iframe hydration behavior.

4. "Open in new tab" downloads instead of opens
  - Confirm source URL points to a renderable page endpoint.
  - Prefer canonical site URL rather than temporary sharing URL.

## Repro Capture Checklist

1. Full host URL used.
2. UHV property-pane values:
  - source mode
  - content delivery mode
  - security mode
  - sandbox preset
3. Browser console filtered by `uhv`.
4. Screenshot of behavior and console.

## Quick Runtime Verification

Use these console probes:

```js
window.__uhvScrollTrace?.slice(-20)
window.scrollY
document.getElementById('uhv-scroll-top-marker')?.getBoundingClientRect()
```

## Common Fixes

1. Replace sharing URL with canonical path:
  - From: `https://<tenant>.sharepoint.com/:u:/r/sites/...?...`
  - To: `https://<tenant>.sharepoint.com/sites/<site>/Shared%20Documents/<file>.html`

2. Keep host page and target HTML in same site collection.

3. If UI still shows stale behavior:
  - Bump solution/webpart versions and redeploy.

4. For minimal UI during demos:
  - Use published page view.
  - Optionally append `?env=Embedded`.
  - Disable comments and site social bar where needed.

