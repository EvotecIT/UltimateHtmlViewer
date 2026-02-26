# SPFx Security Uplift Spike

This spike formalizes a safe path to reduce remaining transitive dependency risk in the SPFx build chain.

Status: Started on **2026-02-26**

## Goal

- Reduce high-severity tooling-chain findings while preserving successful SPFx bundle and package flows.

## Scope

- `spfx/UniversalHtmlViewer/package.json`
- `spfx/UniversalHtmlViewer/package-lock.json`
- SPFx build/test/lint validation and release packaging smoke checks

## Out Of Scope

- Runtime feature changes in the UHV web part
- Forced framework downgrades/upgrades that break current SPFx runtime target

## Controlled Branch Plan

1. Create spike branch from `master`:
   - `spike/spfx-toolchain-uplift-2026-02`
2. Apply only constrained dependency changes per commit:
   - safe overrides first
   - lockfile refresh second
   - wider SPFx/tooling version trial last
3. Keep each experiment reversible:
   - one concern per commit
   - record audit delta after each step

## Execution Checklist

```powershell
cd spfx/UniversalHtmlViewer

# Baseline snapshot
npm audit --json > ../../ignore/audit-baseline.json

# After each change
npm install --package-lock-only
npm audit --json > ../../ignore/audit-after-change.json
npm run lint
npm test -- --runInBand
npm run bundle
```

## Acceptance Gates

- `npm run lint` passes
- `npm test -- --runInBand` passes
- `npm run bundle` passes
- `npm audit` total or high-severity count is reduced or unchanged (never worse)

## Exit Criteria

- A documented PR either:
  - merges safe dependency-risk reductions, or
  - records blocked upgrades and exact blockers with package names and versions

## Target Timeline

- Initial spike branch created: **2026-02-26**
- First spike findings review: **2026-03-05**
- Decision checkpoint (merge, split, or defer): **2026-03-12**

## Iteration Log

### Iteration 1 (2026-02-26)

- Branch: `spike/spfx-toolchain-uplift-2026-02`
- Change:
  - added lockfile overrides for:
    - `braces` -> `3.0.3`
    - `tough-cookie` -> `4.1.4`
- Audit delta (`spfx/UniversalHtmlViewer`):
  - before: `total 73`, `high 12`, `moderate 61`
  - after: `total 71`, `high 0`, `moderate 71`
- Validation:
  - `npm run lint` passed
  - `npm test -- --runInBand` passed
  - `npm run bundle` blocked in this environment due Node `v24.13.0` (SPFx requires `<23`)
  - follow-up validation passed with Node `v22` wrapper:
    - `npx -y node@22 ./node_modules/gulp/bin/gulp.js bundle`
    - `npx -y node@22 ./node_modules/gulp/bin/gulp.js package-solution`
- Notes:
  - This iteration intentionally avoided SPFx major-version changes.
  - Dependabot alert state should be re-checked after merge and GitHub re-scan.

### Iteration 2 (2026-02-26)

- Branch: `spike/spfx-security-uplift-iteration-2b`
- Goal:
  - attempt safe closure of remaining medium alerts (`ajv`, `postcss`, `request`) without SPFx-major upgrades
- Experiments:
  - nested overrides targeting `ajv` under `@microsoft/tsdoc-config` and `@rushstack/node-core-library`
  - direct `postcss` override uplift to `8.4.31`
  - constrained `@rushstack/eslint-config` alignment and lockfile-only refresh checks
- Result:
  - some combinations reduced local `npm audit` totals
  - all meaningful reductions required an invalid dependency graph (`npm ls` reported `ELSPROBLEMS` / invalid constraints)
  - therefore all iteration-2 dependency edits were reverted
- Current validated baseline after reverts:
  - `npm audit`: `total 71`, `high 0`, `moderate 71`
  - Dependabot open alerts: `GHSA-2g4f-4pwh-qvx6` (`ajv`), `GHSA-7fh5-64p2-3v2j` (`postcss`), `GHSA-p8p7-x288-28g6` (`request`)
- Decision:
  - do not force incompatible overrides
  - carry unresolved items as monitored until a compatible SPFx toolchain path is available
