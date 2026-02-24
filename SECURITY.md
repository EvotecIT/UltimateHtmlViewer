# Security Notes

This project uses GitHub Dependabot and regular dependency maintenance, but some alerts are currently constrained by the SPFx toolchain dependency graph.

## Reporting Security Issues

Please report security issues privately through GitHub Security Advisories for this repository.

## Current Open Dependabot Alerts

Status captured on **2026-02-24** from:

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

