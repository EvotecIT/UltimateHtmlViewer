# Security Notes

This project uses GitHub Dependabot and regular dependency maintenance. Some alerts are currently constrained by SPFx toolchain compatibility and release-build requirements.

## Reporting Security Issues

Please report security issues privately through GitHub Security Advisories for this repository.

## Current Open Dependabot Alerts

Status captured on **2026-02-27** from:

- `repos/EvotecIT/UltimateHtmlViewer/dependabot/alerts?state=open`

| Alert | Package | Severity | Scope | Current status |
| --- | --- | --- | --- | --- |
| GHSA-vghf-hv5q-vc2g | `validator` | high | development | Open. Transitive via SPFx tooling chain. |
| GHSA-9965-vmph-33xx | `validator` | medium | development | Open. Same transitive path as above. |
| GHSA-4vvj-4cpr-p986 | `webpack` | medium | development | Open. Transitive in SPFx tooling. |
| GHSA-7fh5-64p2-3v2j | `postcss` | medium | development | Open. Transitive in legacy SPFx Sass chain. |
| GHSA-8fgc-7cc6-rx7x | `webpack` | low | development | Open. Transitive in SPFx tooling. |
| GHSA-38r7-794h-5758 | `webpack` | low | development | Open. Transitive in SPFx tooling. |

Recently closed:

- `GHSA-p8p7-x288-28g6` (`request`) via alias to `@cypress/request@3.0.0` in PR #65.

## Important Context (2026-02-27)

- A temporary override to pin `webpack` to `5.105.3` previously reduced webpack/validator alerts.
- During release prep for `v1.0.31.0`, that override caused `bundle:ship` failures (Terser parse error in SPFx ship build path).
- To restore a stable release pipeline, the explicit webpack override was removed in commit `990d718`.
- Result: release packaging and deployment are healthy again, but webpack/validator transitive alerts re-opened.

## Security Ownership Matrix (As Of 2026-02-27)

| Workstream | Advisory focus | Exposure class | Current mitigation | Owner | Next checkpoint | Target date | Status |
| --- | --- | --- | --- | --- | --- | --- | --- |
| Release-safe dependency posture | `webpack`, `validator` | Build-time only | Keep SPFx-native dependency graph to preserve `bundle:ship` / `package-solution:ship`; no incompatible override forcing | Repo maintainers | Validate against next SPFx-compatible uplift path | 2026-03-12 | In progress |
| SPFx Sass transitive chain | `postcss` | Build/dev chain | Monitor SPFx package updates; avoid breaking override combinations | Repo maintainers | Re-test with uplift spike branch | 2026-03-12 | In progress |
| Runtime boundary controls | URL/security policy surface | Runtime | Keep strict URL validation and secure defaults in UHV runtime | Repo maintainers | Revalidate on each release train | 2026-03-12 | Monitoring |

## Controlled SPFx Uplift Spike

A controlled spike plan is tracked in:

- `docs/SPFx-Security-Uplift-Spike.md`

Objective:

- Reduce transitive tooling findings without destabilizing production SPFx builds and release packaging.

## What We Do Today

- Patch direct and safely-overridable transitive dependencies where possible.
- Prefer release pipeline reliability over unsafe/incompatible override forcing.
- Re-evaluate unresolved alerts on each SPFx upgrade and lockfile refresh.

## Revalidation Checklist

```powershell
cd spfx/UniversalHtmlViewer
npx -y -p node@22.14.0 -p npm@10.9.2 -c "npm ci"
npx -y -p node@22.14.0 -p npm@10.9.2 -c "npm run bundle:ship"
npx -y -p node@22.14.0 -p npm@10.9.2 -c "npm run package-solution:ship"
gh api -X GET "repos/EvotecIT/UltimateHtmlViewer/dependabot/alerts?state=open&per_page=100"
```
