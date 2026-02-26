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
| GHSA-8fgc-7cc6-rx7x | `webpack` | low | development | Open. Tooling-chain dependency; constrained by SPFx webpack compatibility. |
| GHSA-38r7-794h-5758 | `webpack` | low | development | Open. Tooling-chain dependency; constrained by SPFx webpack compatibility. |
| GHSA-vghf-hv5q-vc2g | `validator` | high | development | Open. Transitive tooling dependency; SPFx upgrade path required. |
| GHSA-9965-vmph-33xx | `validator` | medium | development | Open. Transitive tooling dependency; SPFx upgrade path required. |
| GHSA-grv7-fg5c-xmjg | `braces` | high | development | Open. Transitive tooling dependency; SPFx upgrade path required. |
| GHSA-7fh5-64p2-3v2j | `postcss` | medium | development | Open. Transitive tooling dependency; SPFx upgrade path required. |
| GHSA-72xf-g2v4-qvf3 | `tough-cookie` | medium | development | Open. Transitive tooling dependency; SPFx upgrade path required. |
| GHSA-p8p7-x288-28g6 | `request` | medium | development | Open. Legacy transitive dependency in tooling chain. |

## Security Ownership Matrix (As Of 2026-02-26)

| Workstream | Advisory focus | Exposure class | Current mitigation | Owner | Next checkpoint | Target date | Status |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Lockfile and override hygiene | `webpack` (`GHSA-8fgc-7cc6-rx7x`, `GHSA-38r7-794h-5758`) | Build-time only | Override pinned to `webpack@5.105.3`; lockfile refreshed in PR #25 | Repo maintainers | Re-run `npm audit` and Dependabot triage | 2026-03-05 | In progress |
| SPFx toolchain transitive chain | `braces`, `validator`, `postcss`, `tough-cookie`, `request` | Build/dev chain | Keep direct dependencies patched where safe; avoid forced SPFx-major drift | Repo maintainers | Execute controlled uplift spike runbook | 2026-03-12 | In progress |
| Runtime package safety | `ajv` (indirect path) | Runtime policy + shared libs | Continue runtime boundary checks and strict URL policy defaults in UHV | Repo maintainers | Re-validate after each SPFx lockfile refresh | 2026-03-12 | Monitoring |

## Controlled SPFx Uplift Spike

A controlled spike plan is now tracked in:

- `docs/SPFx-Security-Uplift-Spike.md`

Spike objective:

- Reduce remaining high-severity transitive tooling findings without destabilizing production SPFx builds.

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
