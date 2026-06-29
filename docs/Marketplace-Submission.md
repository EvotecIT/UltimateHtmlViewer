# Microsoft Marketplace Submission Notes

Use this file as the source checklist for preparing Universal HTML Viewer for Microsoft Marketplace / AppSource / SharePoint Store submission through Partner Center.

## Offer Shape

- Offer type: SharePoint Framework solution.
- Product name: Universal HTML Viewer.
- Publisher: Evotec Services sp. z o.o.
- Author: Przemysław Kłys.
- Pricing: Free listing. The app does not require a license gate, entitlement service, paid checkout flow, or Evotec-hosted backend for this submission.
- Package: `spfx/UniversalHtmlViewer/sharepoint/solution/universal-html-viewer.sppkg`.
- App/package version prepared for submission: `1.0.32.24`.
- Web part manifest version: `1.0.34`.
- Tenant-wide deployment: not enabled in the package (`skipFeatureDeployment=false`). Tenant or site admins should add the app to target sites after app catalog installation.

Choose the offer type carefully in Partner Center. Microsoft does not allow changing the offer type after the offer is created.

## Package Metadata

The Partner Center listing should match the SPFx package metadata where Microsoft validates consistency.

| Field | Value |
| --- | --- |
| Name | Universal HTML Viewer |
| Publisher | Evotec Services sp. z o.o. |
| Author | Przemysław Kłys |
| Short description | Render trusted HTML reports and dashboards inside SharePoint pages. |
| Long description | Universal HTML Viewer lets site owners embed and navigate trusted HTML reports and supporting assets in SharePoint pages while keeping navigation inside the host experience. |
| Website URL | `https://github.com/EvotecIT/UltimateHtmlViewer` |
| Privacy URL | `https://github.com/EvotecIT/UltimateHtmlViewer/blob/master/docs/Privacy.md` |
| Terms URL | `https://github.com/EvotecIT/UltimateHtmlViewer/blob/master/docs/Terms-of-Use.md` |
| Support URL | `https://github.com/EvotecIT/UltimateHtmlViewer/issues` |
| Categories | Productivity; Content management; Data + analytics |

## Suggested Marketplace Description

Universal HTML Viewer is a SharePoint Framework web part for rendering trusted static HTML reports, dashboards, and small HTML applications inside modern SharePoint pages.

Use it when teams publish generated HTML reports or documentation bundles to SharePoint and need a stable host page, predictable iframe behavior, SharePoint-file rendering, report navigation, and security controls around allowed paths, hosts, and file extensions.

Key capabilities:

- Render HTML from SharePoint files through the SharePoint file API.
- Keep report navigation inside the SharePoint host page.
- Share links to report states with `uhvPage` URL state or same-page hash fragments.
- Configure strict tenant, allowlist, or expert HTTPS security modes.
- Browse report folders when using the SharePoint report browser source mode.
- Use scripted build, deployment, update, rollback, and site setup flows.

UHV does not provide an external hosted service. It runs in the customer's Microsoft 365 tenant and uses the current user's SharePoint permissions.

## Screenshot Set

Recommended screenshots from `assets/`:

| File | Suggested caption |
| --- | --- |
| `assets/uhv-site-contents-app-tile.png` | Confirm UHV is installed in SharePoint site contents. |
| `assets/uhv-showcase-runtime-page.png` | Render an HTML report inside a modern SharePoint page. |
| `assets/uhv-showcase-editor-quick-setup.png` | Configure the SharePoint HTML file source and delivery mode. |
| `assets/uhv-source-mode-selector.png` | Choose between single page, report browser, and URL-builder source modes. |
| `assets/uhv-source-mode-report-browser.png` | Let users browse permitted report files from a SharePoint folder. |
| `assets/uhv-showcase-deeplink-security-ops.png` | Show deep links, security controls, and operational guidance. |

Asset dimensions checked during preparation:

| Asset family | Dimensions |
| --- | --- |
| `UHV-Small.png`, `UHV-Smaller.png` | 960 x 640 |
| `UHV-Big.png` | 1536 x 1024 |
| Runtime/source screenshots | about 1905 x 1900 |
| Site contents screenshot | 1905 x 933 |

## Microsoft Test Instructions

Provide Partner Center testers with a tenant or test site where UHV is already installed, or provide install instructions and test credentials in Partner Center only. Do not commit credentials to this repository.

Suggested tester flow:

1. Open the provided SharePoint test site.
2. Upload `samples/siteassets/UHV-Feature-Showcase.html` to a document library or Site Assets.
3. Create or edit a modern SharePoint page.
4. Add the Universal HTML Viewer web part.
5. Select `SharePoint file API (inline iframe)` as the content delivery mode.
6. Set the initial/default HTML page to the uploaded showcase HTML file.
7. Publish the page.
8. Confirm the HTML renders inline, navigation stays in the UHV host page, and a direct page hash such as `#security` scrolls to the embedded report section.

## Pre-Submission Checklist

- [ ] Partner Center publisher profile is complete and verified.
- [ ] Offer type is SharePoint Framework solution.
- [ ] Package solution name, descriptions, and URLs match Partner Center listing text.
- [ ] Privacy and terms URLs are public and reviewed.
- [ ] Support URL is public and monitored.
- [ ] Screenshots are uploaded with clear captions.
- [ ] Test credentials and test site instructions are entered in Partner Center only.
- [ ] `.sppkg` was built with `gulp bundle --ship` and `gulp package-solution --ship`.
- [ ] UHV was smoke-tested on a SharePoint test site.
- [ ] Legal/compliance owner approved privacy and terms text before submission.

## Validation Commands

Run from the repository root:

```powershell
npx -y -p node@22.14.0 -c "cd spfx/UniversalHtmlViewer && npm run lint"
npx -y -p node@22.14.0 -c "cd spfx/UniversalHtmlViewer && npm test -- --runInBand"
npx -y -p node@22.14.0 -c "cd spfx/UniversalHtmlViewer && npm run build"
.\scripts\Build-UHV.ps1 -SkipInstall -QuietNpm
```

## Known Submission Risks

- Privacy and terms text should be reviewed before using it as the final public legal/compliance language.
- If Partner Center requires company-hosted legal URLs instead of GitHub-hosted Markdown pages, publish `docs/Privacy.md` and `docs/Terms-of-Use.md` to an Evotec-owned website and update `package-solution.json`.
- Because `skipFeatureDeployment=false`, tenant-wide deployment should not be presented as the default install model.
- Switching to a paid listing later should be treated as a separate product decision because it would require license/entitlement behavior, billing terms, and no-license UX that UHV does not currently implement.
