# UHV Release Checklist

Use this checklist before publishing a new release tag.

## 1) Preflight

- [ ] Confirm working tree is clean.
- [ ] Confirm target version in `spfx/UniversalHtmlViewer/src/webparts/universalHtmlViewer/UniversalHtmlViewerWebPart.manifest.json`.
- [ ] Confirm Node version matches project/tooling requirements (`>=22.14.0 <23.0.0` for local SPFx build scripts).

## 2) Validate locally

```powershell
cd spfx/UniversalHtmlViewer
npm ci
npm run lint
npm test -- --runInBand
```

If your local Node version is out of range, use `.nvmrc`/`.node-version` and rerun.

## 3) Tag and push

```powershell
git tag v<version>
git push origin v<version>
```

Tag push triggers `.github/workflows/release-sppkg.yml`.

## 4) Verify workflow output

- [ ] Workflow completed successfully.
- [ ] Versioned package exists in artifacts:
  - `release/universal-html-viewer-<version>.sppkg`
- [ ] `SHA256SUMS.txt` artifact present.

## 5) Verify GitHub Release

- [ ] Release entry exists under GitHub Releases for tag `v<version>`.
- [ ] Attached assets include:
  - `universal-html-viewer-<version>.sppkg`
  - `SHA256SUMS.txt`

## 6) Post-release sanity checks

- [ ] Downloaded `sppkg` filename and manifest version match.
- [ ] Deployment docs still reflect current process (`docs/Deploy-SharePointOnline.md`).
- [ ] Communicate release notes to operators/site admins.
