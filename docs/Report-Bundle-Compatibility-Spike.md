# Report Bundle Compatibility Spike

Status: spike, **2026-09-25**. Branch: `spike/report-bundles`. No runtime code changed.

This spike checks whether UHV can host the multi-file report bundles planned for
TestimoX (reporting platform plan, sections 6.2 "Bundle" and 6.4 "SharePoint
through UltimateHtmlViewer"), and lists the UHV changes proposed for plan phase 3.

No SharePoint tenant was used. Evidence comes from UHV's own transform functions
under Jest/jsdom and from reading the code. Everything that depends on SharePoint
Online response headers, SharePoint's own page CSP or real browser behavior is
marked **unverified** and listed in the manual checklist at the end.

## What was added

- `samples/report-bundle/`: a minimal bundle.
  - `index.html` loads `assets/app.css`, `assets/app.js` and `data/probes/0.js`
    with plain relative tags. It links to `page2.html#section`, and `app.js`
    creates a second `page2.html#section` link at runtime.
  - `page2.html` loads `data/probes/0.js` at runtime through a script element
    injected by `HfxData.load(...)`. It also links back to `index.html#summary`.
  - `assets/app.js` is a stand-in for the bundle runtime: `HfxData.register(id, chunk, payload)`,
    `HfxData.load(...)`, a runtime-created link and a listener for the
    proposed `uhv-host-theme` message.
  - `data/probes/0.js` is a sidecar that calls `HfxData.register("probes", 0, {...})`.
  - `manifest.json` lists pages, sections, assets and dataset chunks.
- `spfx/UniversalHtmlViewer/src/webparts/universalHtmlViewer/__tests__/ReportBundleCompatibility.test.ts`:
  21 Jest tests. They pass the sample HTML through `prepareInlineHtmlForSrcDoc`,
  `prepareInlineHtmlForBlobUrl`, `appendAdditionalCspHostSources`,
  `wireInlineAnchorRuntimeRewrite`, `resolveInlineDeepLinkTarget` and
  `inlineAllowedExternalScripts`. The inputs mirror what the web part passes:
  - source/base URL: `https://contoso.sharepoint.com/sites/Reports/Shared%20Documents/TestimoX/index.html`
  - host page: `https://contoso.sharepoint.com/sites/Reports/SitePages/Reports.aspx`
  - `allowQueryStringPageOverride` enabled, deep-link param `uhvPage`

### Test results

The suite ran on Node 22.14.0 with `npm test -- --runInBand`:

- Before the spike: 48 suites and 329 tests passed.
- After the spike: 49 suites and 350 tests passed, including the 21 new tests.
- `npm run lint` is clean, and `npm run build` (Heft) passes.

## Findings

### Works as-is

1. **The UHV CSP allows the bundle's own assets and sidecars (srcdoc mode).**
   `getDefaultSrcDocContentSecurityPolicy` lists the page origin and the base origin
   in the directives the bundle needs:

   - `script-src`: `assets/*.js`, static `data/*.js` and runtime-injected `data/*.js`
   - `style-src`: `assets/*.css`
   - `connect-src`: fetching `manifest.json`
   - `img-src`, `font-src`, `frame-src`

   The tests resolve every relative `script[src]` and `link[href]` in the sample
   against the base, then check each URL against the injected directives.
   Test: *injects a CSP whose script-src and style-src allow ...*.
2. **The strict inline CSP also works.** With `enforceStrictInlineCsp`,
   `'unsafe-inline'` and `'unsafe-eval'` are removed, but external scripts from
   the tenant origin still load. The sample bundle has no inline report script,
   so it runs under the strict policy.
3. **`<base href>` resolves relative paths to the library folder.** UHV injects
   `<base href="<source file URL>">` with the query and fragment removed. Inside a
   frame with that base, `document.baseURI` is the library file. The runtime
   sidecar URL (`new URL('data/probes/0.js', document.baseURI)`) resolves to
   `.../Shared%20Documents/TestimoX/data/probes/0.js`, and an injected sidecar
   element's `src` points there. The payload registers correctly.
   Tests: *resolves runtime sidecar URLs against the injected base*, *injects a runtime sidecar script ...*.
4. **Bundle links become host deep links and keep the fragment.**
   - `page2.html#section` becomes
     `Reports.aspx?uhvPage=%2Fsites%2FReports%2FShared%2520Documents%2FTestimoX%2Fpage2.html%23section`.
   - `data-uhv-inline-href` keeps the absolute target with `#section`.
   - `resolveInlineDeepLinkTarget` on that host URL returns `.../page2.html#section`.
   - The page2 back-link `index.html#summary` works the same way.
5. **Links created at runtime are rewritten too.** `wireInlineAnchorRuntimeRewrite`
   (MutationObserver) catches the link that `app.js` adds after load. The inline
   navigation bridge also intercepts clicks on it. The MutationObserver rewrite
   needs same-origin frame access, so it does not run under the Strict preset
   (gap 4).
6. **Blob mode resolves relative URLs.** `prepareInlineHtmlForBlobUrl` injects the
   same `<base href>`, navigation bridge and link rewrites as srcdoc mode. A
   `blob:` document uses `<base href>` for relative URL resolution, so `assets/`
   and `data/` resolve to the library folder.
7. **The `inlineExternalScripts` fallback covers static tags.**
   `ExternalScriptInliningHelper` always allows same-host scripts. It fetches
   `assets/app.js` and `data/probes/0.js` with `credentials: 'same-origin'` and
   inlines them.
8. **The report browser already shows freshness.** It lists `.html` files with
   `TimeLastModified`, so a bundle's `index.html` can be opened today.

### Gaps and blockers

1. **The fragment is not applied after a cross-page deep link.** The fragment
   survives in `uhvPage`, but it is dropped before the page renders:
   - The content comes from the REST `$value` endpoint. The server-relative path
     loses the hash through `new URL(...).pathname` for absolute URLs, or through
     `stripQueryAndHashFromPath` for server-relative input.
   - The iframe document is `about:srcdoc` or a `blob:` URL without the fragment,
     so the browser never scrolls to `#section` by itself.
   - `HostPageHashNavigationHelper` only reacts to the **host** page's
     `window.location.hash`, not to the target URL's fragment.

   What happens next depends on how the page was reached:
   - **In-frame navigation** (clicking the link): `resetIframeScrollPosition`
     skips its scroll-to-top when the target has a hash (`hasUrlHash`). But
     nothing scrolls to `#section`, so the view stays wherever the new document
     starts, normally the top.
   - **First load of a shared `?uhvPage=...%23section` link**:
     `shouldApplyInitialDeepLinkScrollLock` turns on the initial deep-link scroll
     lock. While the lock is active it calls `forceHostScrollTop()` and
     `resetInlineIframeScrollPositionForDeepLink()`. That in turn calls
     `resetIframeScrollPosition(iframe)` **without** a target URL, so the frame is
     actively forced to the top.

   I found no code that applies `#section` to the newly loaded document. Expect
   `page2.html#section` to open at the top of page2. **Unverified in a real
   browser.** See proposed change P1.
2. **The UHV CSP is scoped to the origin, not the library path.**
   - `script-src` allows any script on `https://<tenant>.sharepoint.com`, not only
     the bundle folder.
   - `appendAdditionalCspHostSources` reduces URLs to their origin, so an admin
     cannot scope the policy to a path either.
   - This is not a blocker, but the policy is looser than a bundle needs.
     See P2.
3. **The host page's CSP also applies (unverified).** Under the HTML spec,
   `about:srcdoc` and `blob:` documents inherit the embedding page's policy
   container. Any SharePoint Online page CSP therefore applies on top of UHV's
   `<meta>` CSP, and UHV cannot relax it. Blob mode has **only** the inherited
   policy, because UHV does not inject a CSP there, and its CSP host options are
   not applied (test *has no UHV CSP in blob mode ...*). The README's note that
   "SharePoint CSP blocks CDN script tags" does not prove inheritance: UHV's own
   origin-only `<meta>` CSP blocks CDN scripts too. Open questions for a tenant:
   - Does the SharePoint policy allow scripts from the tenant's own document
     library paths?
   - Is it nonce-based? UHV stamps the page nonce only on inline scripts without
     `src` (`applyPageScriptNonceToInlineScripts`). Under a nonce plus
     `'strict-dynamic'` policy, the parser-inserted `<script src="assets/app.js">`
     would be blocked. Sidecars injected by an already trusted script would load.
     Scripts inlined by `inlineExternalScripts` have no `src` left, so they do get
     the page nonce. That makes the fallback a workaround for static tags.
4. **The Strict sandbox preset breaks `fetch`.** Without `allow-same-origin`, the
   report has an opaque origin.
   - Classic `<script src>` loads still work, because they are no-CORS.
   - `fetch('manifest.json')` becomes a cross-origin CORS request. SharePoint does
     not send `Access-Control-Allow-Origin`, so the request fails. Even with CORS
     headers, the default `credentials: 'same-origin'` would send no cookies, and
     the request would come back 401/403. That is one more reason for sidecars to
     be scripts.
   - Whether SharePoint auth cookies go with no-CORS subresource requests from an
     opaque-origin frame depends on their `SameSite` attributes. **Unverified.**
   - UHV's host-side helpers read `iframe.contentDocument`. Under Strict that
     access is not available. This affects `wireInlineAnchorRuntimeRewrite`,
     nested frame hydration, `navigateToSamePageHash` and
     `scrollHostPageToIframeHashTarget`.
   - Only the in-frame navigation bridge (a capture-phase click listener that
     posts to the host) keeps working. The jsdom test for runtime link rewriting
     cannot show this, because jsdom always allows frame access.
5. **The `inlineExternalScripts` fallback has limits.**
   - It only sees static `<script src>` tags. Sidecars injected at runtime are
     never inlined (test *cannot inline sidecars that the runtime injects after load*).
   - Where no host page nonce exists, inlined scripts get no nonce, so the
     fallback does not combine with `enforceStrictInlineCsp` (the test asserts the
     missing nonce). Inside SharePoint the page nonce is stamped on them (see gap 3).
   - Big bundles would inline megabytes of JavaScript into `srcdoc` on every page
     view.
6. **Some bundle markup changes UHV behavior.** These are authoring rules for
   TestimoX:
   - A `<meta http-equiv="Content-Security-Policy">` in the bundle turns off UHV's
     CSP. `'self'` in that policy means an opaque origin under the Strict preset.
     **Do not ship a CSP meta tag.**
   - A bundle `<base>` stops UHV's base injection. A relative base like `./`
     resolves against the parent page's base, which is `/sites/Reports/SitePages/`
     for srcdoc. That breaks every asset. **Do not ship a `<base>` tag.**
   - Both checks are raw-text regexes over the whole HTML
     (`/<base\s+/i` and `http-equiv=...content-security-policy`). A literal
     `<base ` or CSP `http-equiv` string inside inline JavaScript or a template
     also turns off UHV's injection. **Keep those strings out of the bundle HTML.**
   - Links with `target` other than `_self`, or modifier clicks, keep native
     behavior (`shouldKeepNativeAnchorBehavior`). **In-bundle links must not set `target`.**
   - Only `.html`/`.htm`/`.aspx` links, or the configured extensions, become host
     navigation. Links to `.json`/`.csv` stay native downloads.
   - `history.pushState` with a hash URL is swallowed by the compatibility shim when
     the browser rejects it. The shell should not depend on `location.hash`
     routing. It should use element `id`s for sections, which
     `SamePageHashNavigationHelper` scrolls to.
7. **SharePoint serving `.js`, `.css` and `.json` from a library (unverified).**
   - SharePoint Online forces `.html` downloads (`Content-Disposition: attachment`,
     `X-Download-Options: noopen`). That is why UHV loads HTML through REST.
   - For subresources, `Content-Disposition` and `X-Download-Options` do not
     matter. The `Content-Type` together with `X-Content-Type-Options: nosniff`
     does: a `.js` served as `application/octet-stream` with `nosniff` is refused
     as a script.
   - JavaScript in Site Assets is commonly referenced, but this spike could not
     confirm the headers for `Shared Documents` on a modern site with custom
     script disabled. The same applies to Blob mode.
8. **Assets are cached across runs.** Plan 6.2 shares `assets/` across report runs
   in the same folder. After an atomic folder swap, the browser can still use a
   cached `assets/app.js`. TestimoX should content-hash asset file names or add a
   version query. UHV's in-memory cache (default 15 s) covers the prepared HTML.
   With `inlineExternalScripts` on, it also covers the inlined script text, with
   the same TTL. Browser caching of `assets/` depends on SharePoint's
   `Cache-Control`/`ETag` headers.

## Proposed UHV changes (plan phase 3)

### P1 — Fragment-aware navigation

- When the resolved target (initial `uhvPage` or inline navigation) has a
  fragment, apply it after load. Call `navigateToSamePageHash(iframeDocument, hash)`
  and `scrollHostPageToIframeHashTarget(...)`, with the same retry delays that
  `InlineNavigationHelper` already uses for host hashes. Keep the existing "skip
  scroll-to-top when hashed" rule.
- Exempt hashed targets from the initial deep-link scroll lock, or release the
  lock once the fragment is applied. Otherwise the lock's
  `resetInlineIframeScrollPositionForDeepLink()` resets the frame to the top again.
  Host scroll-to-top can stay.
- Send the fragment to the report as `{ type: 'uhv-inline-fragment', hash }`
  after `uhv-inline-ready`. This is **required** for the Strict preset, where the
  host cannot reach `iframe.contentDocument`. It also lets a report that renders
  sections late (after sidecar data arrives) scroll itself. Alternatively, the
  navigation bridge can apply a fragment that is embedded in its configuration.
- Tests: srcdoc and blob targets with `#section`, element present at load and
  element rendered late, initial load under the scroll lock, and the Strict
  preset (postMessage path only).

### P2 — Sidecar and CSP support for bundles

- New option `inlineCspBundleScope` (default off). When it is on, emit
  path-scoped sources for the bundle folder
  (`https://tenant/sites/Reports/Shared%20Documents/TestimoX/`) in `script-src`,
  `style-src`, `img-src`, `font-src` and `connect-src`, instead of the bare origin.
  Keep origin-only as the default for compatibility. Note that CSP ignores paths
  after redirects, so this narrows the policy but is not a hard boundary.
- Stamp the host page nonce on same-origin `<script src>` elements inside the
  allowed path prefixes, not only on inline scripts. Parser-inserted bundle
  scripts then pass a nonce plus `'strict-dynamic'` policy.
- Accept `additionalScriptSrcHosts` in blob mode only as documentation. Blob mode
  cannot add a policy that is looser than the inherited one. Say this explicitly
  in the property pane help.
- Diagnostics: count `securitypolicyviolation` events in the frame and show them
  in the existing diagnostics panel. Report the first blocked URL and directive.
  A tenant's CSP problems then become visible without DevTools.

### P3 — Theme handshake

- Host to report, after `uhv-inline-ready` and again on SharePoint theme change:
  `postMessage({ type: 'uhv-host-theme', version: 1, isDark, tokens: { accent, surface, text, ... } }, '*')`.
  Build the tokens from `this.context` theme data (`ThemeProvider`/`IReadonlyTheme`
  `semanticColors` and `palette`) and send only a fixed list of hex colours.
  `'*'` is needed because srcdoc and blob frames under the Strict preset have an
  opaque origin. The payload holds no secrets.
- Report side, already prototyped in `samples/report-bundle/assets/app.js`:
  - accept only messages where `event.source === window.parent`
  - accept only matching `type` and `version`
  - check each token against `/^#(?:[0-9a-f]{3,4}|[0-9a-f]{6}|[0-9a-f]{8})$/i`
  - map tokens to report CSS variables and set `data-report-theme`
  - clear tokens that a later message omits

  Tests cover a valid message, invalid and omitted tokens, and messages with the
  wrong source or version.
- Optionally, the report replies with `{ type: 'uhv-report-capabilities', theme: true }`.
  UHV then knows that the report adopted the theme and can skip any host-side
  fallback styling.

### P4 — Manifest-aware report picker and freshness

- In `SharePointReportBrowser` mode, a folder that contains `manifest.json`
  becomes one report entry. The entry shows:
  - title, `generatedAt` and entry page from the manifest
  - the page list as sub-navigation
  - the `assets/` and `data/` subfolders, hidden
- Read the manifest through the same `GetFileByServerRelativePath(...)/$value` REST
  path, with security trimming. Validate a small schema (`schema`, `entry`,
  `pages[].path` must stay relative and inside the folder) and cap its size.
- Freshness badge: "Last updated" from `generatedAt`, falling back to
  `TimeLastModified`. Optionally warn when `generatedAt` is older than a
  configured age.
- Auto-refresh (`FileLastModified` cache buster) can watch `manifest.json`
  instead of the HTML page, so one change signal covers the whole bundle.

### P5 — Documentation

- Add a "Report bundles" section to the README with the authoring rules from
  gap 6 and the recommended preset:
  - `SharePointFileContent`
  - `allowQueryStringPageOverride` on
  - `allowedPathPrefixes` set to the report root
- Point this spike's checklist to `docs/Smoke-Test-Checklist.md`.

## Manual verification checklist (real tenant)

Upload `samples/report-bundle/` to
`/sites/Reports/Shared Documents/TestimoX/` and keep its folder structure. Add a
UHV web part in `SharePointFileContent` mode with Single page URL `.../TestimoX/index.html`,
`allowQueryStringPageOverride` on and `allowedPathPrefixes`
`/sites/Reports/Shared Documents/TestimoX/`. Repeat the checks for each variation
below.

Variations:

- `SharePointFileContent` and `SharePointFileBlobUrl`
- sandbox presets `None`, `Relaxed` and `Strict`
- `enforceStrictInlineCsp` off and on
- Edge/Chrome and Firefox; Safari if available

Checks:

1. Direct GET of `.../assets/app.js`, `.../data/probes/0.js`, `.../assets/app.css`
   and `.../manifest.json` in DevTools Network. Record these headers:
   - `Content-Type` and `X-Content-Type-Options`
   - `Content-Disposition` and `X-Download-Options`
   - `Cache-Control` and `ETag`
   - any `Content-Security-Policy` or `-Report-Only`
2. Record the host page (`Reports.aspx`) response `Content-Security-Policy` and
   `Content-Security-Policy-Report-Only` headers. This is the policy that the
   srcdoc and blob frames inherit.
3. Index page: the summary shows "3 probes, 1 need attention.". There are no CSP
   violations in the console, and `data/probes/0.js` loaded with status 200.
4. Page 2 through a link: the probe table fills from the runtime-injected sidecar.
   The injected `<script data-hfx-sidecar>` resolves to the library folder.
5. Click "Probe details" (static link) and "Probe details (created at runtime)".
   Each opens page2 inside UHV. The host URL changes to `?uhvPage=...page2.html%23section`.
   Record whether the view scrolls to `#section` (gap 1).
6. Paste the deep link into a new tab. Page2 opens inside UHV. Record the scroll
   position again.
7. Back and Forward move between index and page2 inside the host.
8. Strict preset: check which requests carry auth cookies and whether sidecars
   return 200 or 401/403.
9. With the CSP diagnostics from P2 in place, or DevTools
   `securitypolicyviolation` logging: record every blocked URL and directive per
   variation.
10. Theme (after P3): switch the site theme or use a dark section. The report
    accent and background follow it.
11. Freshness (after P4): replace the bundle through the staged swap. The report
    browser shows the new `generatedAt`, and no stale `assets/app.js` is served
    (check the DevTools cache column).
