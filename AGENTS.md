# UltimateHtmlViewer

SharePoint Framework (SPFx) web part, UniversalHtmlViewer (UHV), that hosts
static HTML reports and apps inside modern SharePoint pages with inline
rendering, deep links (`?uhvPage=...`), history handling and URL policy controls.

## Layout

- `spfx/UniversalHtmlViewer/`: the SPFx solution (Heft build, Jest tests).
  Web part code and its `__tests__/` are in
  `src/webparts/universalHtmlViewer/`.
- `scripts/`: PowerShell build, package, deploy, rollback and site setup
  (`Build-UHV.ps1`, `Deploy-UHV-Wrapper.ps1`, `Setup-UHVSite.ps1`, ...).
  `scripts/examples/` holds templates.
- `docs/`: deployment, operations, release checklist, security and marketplace notes.
- `samples/siteassets/`: sample HTML content. `Website/`: project page content.
- `skills/`: operator playbooks (see below).
- `ignore/`: gitignored local-only files such as tenant profiles.

## Build and test

Use Node 22.14.0 (`.nvmrc`, `.node-version`; `package.json` requires
`>=22.14.0 <23.0.0`). The npm `pre*` hooks fail on other versions. If the
system Node is newer, wrap commands:
`npx -y -p node@22.14.0 -c "<command>"`.

```sh
cd spfx/UniversalHtmlViewer
npm ci
npm run lint
npm test
npm run bundle
npm run bundle:ship && npm run package-solution:ship   # release .sppkg
```

```powershell
./scripts/Build-UHV.ps1        # bootstrap Node if needed, build and package
```

- Release-only: pushing a `v<version>` tag runs `release-sppkg.yml`; the tag
  must equal `solution.version` in `config/package-solution.json`. Follow
  `docs/Release-Checklist.md`.
- Deploy scripts (`Deploy-UHV*.ps1`, `Update-UHVSiteApp.ps1`, `Rollback-UHV.ps1`)
  change a real SharePoint tenant. Run them only when asked.

## Generated files

`lib/`, `dist/`, `temp/`, `release/` and `sharepoint/solution/*.sppkg` are build
output and are gitignored. Never commit them.

## Conventions

- Versions live in three places: `solution.version` in
  `config/package-solution.json` (the `.sppkg` and release tag), `version` in
  `UniversalHtmlViewerWebPart.manifest.json`, and the SPFx package versions in
  `package.json`.
- Keep web part logic in small single-purpose helpers next to the web part
  (`*Helper.ts`) with Jest tests.

## Gotchas

- Keep tenant URLs, site names and credentials out of the repo. Use
  `ignore/UHV.LocalProfile.ps1`, copied from
  `scripts/examples/UHV.LocalProfile.example.ps1`.
- Rendering, deep-link and scroll behaviour can only be fully verified in
  SharePoint Online; `docs/Smoke-Test-Checklist.md` lists the manual checks.

## Repository skills

Playbooks in `skills/` (plain Markdown, not auto-loaded by either agent):

- `skills/QUICKSTART.md`: short operator flow across profile, deploy and debug.
- `skills/uhv-release-deploy/`: build, deploy to the app catalog, update sites, verify the version.
- `skills/uhv-sharepoint-debug/`: triage downloads, deep-link failures, scroll jumps and iframe quirks.
- `skills/uhv-tenant-profile-setup/`: create and validate the local tenant profile.
