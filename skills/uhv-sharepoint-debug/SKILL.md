# Skill: UHV SharePoint Debug

## Purpose

Diagnose and fix SharePoint-hosted UHV runtime issues quickly and consistently.

## Symptom Decision Tree

1. HTML downloads instead of rendering
  - Confirm `Content delivery mode = SharePointFileContent`.
  - Confirm URL is canonical path, not `/:u:/r/...`.
  - Confirm URL points to `.html/.htm/.aspx` file in allowed scope.
2. Deep links break or back/forward does nothing
  - Confirm host page URL gets `?uhvPage=...`.
  - Confirm target links are relative or same-site absolute.
  - Confirm extension allowlist includes target extension.
3. Page starts mid-scroll or jumps on load
  - Reproduce with `?uhvTraceScroll=1`.
  - Check for repeated `restore-top` and `scroll-lock` events.
  - Verify nested iframe hydration timing on heavy pages.
4. "Open in new tab" opens download flow
  - Use canonical site URL.
  - Test with `DirectUrl` mode if raw page behavior is desired.
5. Works in one site but not another
  - Compare security mode (`StrictTenant`, `Allowlist`, `AnyHttps`).
  - Compare permissions to page and target file library.

## Repro Capture Checklist

1. Full host URL used.
2. UHV property-pane values:
  - source mode
  - content delivery mode
  - security mode
  - sandbox preset
3. Browser console filtered by `uhv`.
4. Screenshot of behavior and console.
5. Whether issue occurs first load only or every navigation.

## Quick Runtime Verification

Use these console probes:

```js
window.__uhvScrollTrace?.slice(-20)
window.scrollY
document.getElementById('uhv-scroll-top-marker')?.getBoundingClientRect()
location.href
```

## Expected Trace Patterns

- Healthy initial deep-link flow:
  - `deep-link-evaluation`
  - `scroll-lock-start`
  - multiple `restore-top` while hydrating
  - `auto-release-stable`
  - `scroll-lock-released`
- Problem pattern:
  - no `scroll-lock-released` for long period
  - repeated `restore-top` while user cannot interact
  - host URL missing/losing `uhvPage`

## Known False Alarms

- Browser console errors such as `browser.pipe.aria.microsoft.com ... ERR_ADDRESS_INVALID` are often telemetry noise and not UHV runtime failures.
- iframe sandbox warning from SharePoint shell can appear even when UHV behavior is correct.
- `invalid contentSourceFilter config` may appear from host scripts and is not always tied to UHV rendering.

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

5. If only heavy reports misbehave:
  - Start with the lean showcase page to validate host baseline.
  - Then retest heavy report with trace enabled.

## Fast Escalation Package

When opening an issue/PR, include:

- Host URL + exact `uhvPage` value.
- Screenshot/GIF of behavior.
- Console output filtered by `uhv`.
- Current installed UHV version from site.
- Property-pane configuration export or screenshots.
