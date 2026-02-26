# Security Notes

This project uses GitHub Dependabot and regular dependency maintenance, but some alerts are currently constrained by the SPFx toolchain dependency graph.

## Reporting Security Issues

Please report security issues privately through GitHub Security Advisories for this repository.

## Current Open Dependabot Alerts

Status captured on **2026-02-26** from:

- `repos/EvotecIT/UltimateHtmlViewer/dependabot/alerts?state=open`

| Alert | Package | Severity | Scope | Current status |
| --- | --- | --- | --- | --- |
| GHSA-2g4f-4pwh-qvx6 | `ajv` | medium | runtime | Open. Indirect through SPFx dependency graph; monitor SPFx-compatible upgrade path. |
| GHSA-7fh5-64p2-3v2j | `postcss` | medium | development | Open. Transitive tooling dependency; SPFx upgrade path required. |
| GHSA-p8p7-x288-28g6 | `request` | medium | development | Open. Legacy transitive dependency in tooling chain. |

Recently closed (after lockfile/override remediation in PR #25 and PR #27):

- `GHSA-8fgc-7cc6-rx7x` (`webpack`)
- `GHSA-38r7-794h-5758` (`webpack`)
- `GHSA-vghf-hv5q-vc2g` (`validator`)
- `GHSA-9965-vmph-33xx` (`validator`)
- `GHSA-grv7-fg5c-xmjg` (`braces`)
- `GHSA-72xf-g2v4-qvf3` (`tough-cookie`)

## Security Ownership Matrix (As Of 2026-02-26)

| Workstream | Advisory focus | Exposure class | Current mitigation | Owner | Next checkpoint | Target date | Status |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Lockfile and override hygiene | `webpack`, `validator`, `braces`, `tough-cookie` (recently closed advisories) | Build-time only | Overrides pinned to `webpack@5.105.3`, `braces@3.0.3`, `tough-cookie@4.1.4`; lockfile refreshed in PR #25 and PR #27 | Repo maintainers | Verify no alert regressions after each dependency refresh | 2026-03-05 | Monitoring |
| SPFx toolchain transitive chain | `postcss`, `request` | Build/dev chain | Keep direct dependencies patched where safe; avoid forced SPFx-major drift and reject invalid dependency-constraint overrides | Repo maintainers | Execute controlled uplift spike runbook | 2026-03-12 | In progress |
| Runtime package safety | `ajv` (indirect path) | Runtime policy + shared libs | Continue runtime boundary checks and strict URL policy defaults in UHV | Repo maintainers | Re-validate after each SPFx lockfile refresh | 2026-03-12 | Monitoring |

## Controlled SPFx Uplift Spike

A controlled spike plan is now tracked in:

- `docs/SPFx-Security-Uplift-Spike.md`

Spike objective:

- Reduce unresolved transitive tooling findings without destabilizing production SPFx builds.

Latest local spike snapshot (2026-02-26, iteration 2):

- `npm audit` moved from `73 total / 12 high / 61 moderate` to `71 total / 0 high / 71 moderate`
- `braces` and `tough-cookie` closed safely via constrained lockfile overrides
- follow-up attempts for `ajv` and `postcss` were rejected because they introduced invalid dependency constraints (`npm ls` `ELSPROBLEMS`)
- current live Dependabot open set: `ajv`, `postcss`, `request` (all medium)

## What We Do Today

- Patch direct and safely-overridable transitive dependencies where possible.
- Keep SPFx-compatible versions pinned to avoid broken ship builds.
- Re-evaluate unresolved alerts on each SPFx upgrade and lockfile refresh.

## Revalidation Checklist

```powershell
cd spfx/UniversalHtmlViewer
npm audit --omit=dev
npm audit
gh api -X GET "repos/EvotecIT/UltimateHtmlViewer/dependabot/alerts?state=open&per_page=100"
```
