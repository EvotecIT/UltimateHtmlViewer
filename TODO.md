# UniversalHtmlViewer Program TODO

This file tracks the end‑to‑end plan and current status for making **UniversalHtmlViewer** the safe, flexible hosting surface for **any HTML dashboard** in SharePoint Online, plus the downstream work in **TheDashboard** and **SharePointEssentials**. It uses checkbox status so you can see progress at a glance.

## Project map (paths)
- UniversalHtmlViewer (SPFx web part + utilities)
  - `/mnt/c/Support/GitHub/UniversalHtmlViewer`
  - SPFx solution: `/mnt/c/Support/GitHub/UniversalHtmlViewer/spfx/UniversalHtmlViewer`
  - Duplicate utils (to be consolidated): `/mnt/c/Support/GitHub/UniversalHtmlViewer/src`
- TheDashboard (PowerShell module that generates static dashboards)
  - `/mnt/c/Support/GitHub/TheDashboard`
- SharePointEssentials (PowerShell sync/publish utilities)
  - `/mnt/c/Support/GitHub/SharePointEssentials`
- HtmlForgeX (C# HTML builder, PSWriteHTML equivalent)
  - `/mnt/c/Support/GitHub/HtmlForgeX`

---

## Current baseline (known state)
- [x] Initial code review of UniversalHtmlViewer and TheDashboard.
- [ ] Decide security defaults (strict tenant vs allowlist vs any https).
- [ ] Decide hosting target (SharePoint document library recommended).
- [ ] Decide whether to allow external CDN assets by default.

---

## Workstream 1 — UniversalHtmlViewer (primary focus)
**Goal:** Be the single, secure way to host any HTML/CSS/JS dashboard in SharePoint Online, with strong defaults and configurable controls.

### 1.1 Cleanup and modernization
- [x] Make `/spfx/UniversalHtmlViewer` the single source of truth (remove or archive duplicate `/src`).
- [x] Upgrade SPFx to current supported version (Node 22 compatible) and update dependencies.
- [x] Ensure Jest/unit tests still pass after uplift.

### 1.2 Security + URL policy (safe-by-default)
- [x] Block protocol‑relative URLs (`//...`) and unsafe schemes (`javascript:`, `data:`, `vbscript:`).
- [x] Add security profile option:
  - [x] `StrictTenant` (default)
  - [x] `Allowlist` (tenant + allowed hosts)
  - [x] `AnyHttps` (explicit opt‑in)
- [x] Add `allowedHosts` property (comma list) for CDN/FrontDoor/extra domains.
- [x] Add optional `allowedPathPrefixes` (e.g., `/sites/Reports/Dashboards/`).

### 1.3 Iframe controls (configurable)
- [x] Add `sandbox` options (off by default for compatibility; configurable).
- [x] Add `allow` attribute options (fullscreen, clipboard, etc.).
- [x] Add `referrerPolicy`, `loading`, `title/aria-label` options.
- [x] Optional auto‑refresh (`refreshIntervalMinutes`).

### 1.4 Cache‑busting
- [x] `cacheBusterMode: None | Timestamp | FileLastModified`.
- [x] Implement `FileLastModified` via SPO REST for server‑relative paths.

### 1.5 Tests + docs
- [x] Add tests for URL validation and allowlist behavior.
- [x] Update README with new properties and security guidance.

---

## Workstream 2 — TheDashboard (output + SharePoint profile)
**Goal:** Produce SharePoint‑safe output, allow “host anything,” and reduce regeneration overhead. Two output modes to support both legacy and future approaches.

### 2.1 Output correctness (must‑fix)
- [ ] Use `UrlName` (not filesystem folder leaf) when building HREFs.
- [ ] URL‑encode/sanitize file names and links for SharePoint compatibility.
- [ ] Apply `UrlPath` consistently to nav links, calendar links, and iframe sources.
- [ ] Fix dynamic scope in `New-HTMLReportPage` (use parameters instead of `$CurrentReport`).

### 2.2 SharePoint Hosting Profile
- [ ] Add `-HostingProfile SharePointOnline` (or switch) to enable:
  - [ ] `.html` output only (no .aspx)
  - [ ] SharePoint filename sanitization
  - [ ] Content fixes (strip `<%` patterns / CSP meta) where needed
  - [ ] Force `-Online` (CDN usage) where appropriate

### 2.3 Output strategy options (both supported)
- [ ] **WrapperPages** (current): keep per‑report wrapper HTML with menu + iframe.
- [ ] **ManifestShell** (new): generate a single shell page + `dashboard.json` manifest.
  - [ ] Shell loads menu + report list from JSON.
  - [ ] Deep‑link via query string or hash (`?report=...`).
  - [ ] Massive reduction in file count, faster generation and upload.

### 2.4 Dependencies / PSWriteHTML implications
- [ ] Evaluate PSWriteHTML changes needed for ManifestShell and/or link generation.
- [ ] If required, define minimal changes or HtmlForgeX alternative.

---

## Workstream 3 — SharePointEssentials (publish/sync)
**Goal:** Make deployment repeatable and safe with incremental sync and logging.

- [ ] Add `Publish-TheDashboardSharePoint` wrapper:
  - [ ] Calls TheDashboard generation
  - [ ] Syncs output using SharePointEssentials
  - [ ] Hash/mtime based uploads
  - [ ] Exclusions (`.json`, `.xml`, temp, logs)
  - [ ] Good logs and `-WhatIf` support

---

## Workstream 4 — HtmlForgeX / C# acceleration (optional)
**Goal:** Evaluate if C# is materially faster or more stable for large output (5k+ files).

- [ ] POC: Render a small dashboard in HtmlForgeX and compare speed/size.
- [ ] Decide on full port vs hybrid (PS for orchestration, C# for rendering).
- [ ] If adopted, define compatibility layer with PSWriteHTML output.

---

## Decision log (fill as we go)
- [ ] Default security profile: ___
- [ ] Allow external CDNs by default: ___
- [ ] Primary hosting location (library/site): ___
- [ ] Output strategy default (WrapperPages vs ManifestShell): ___

---

## Status notes
- UniversalHtmlViewer is the current priority.
- TheDashboard work may require PSWriteHTML or HtmlForgeX changes for ManifestShell mode.
- SharePoint .aspx hosting is ending in March 2026 → HTML hosting via UniversalHtmlViewer becomes the standard.
