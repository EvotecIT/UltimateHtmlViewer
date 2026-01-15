NOTE: The SPFx solution already exists under `spfx/UniversalHtmlViewer`
This file is legacy scaffolding guidance; keep for reference only.

Step 1 – Make sure Node is OK

  1. In PowerShell:

     node -v
  2. If the version starts with v22. you’re good for the latest SPFx (1.21).
  3. If it’s not v22.x:
      - Easiest “best practice” path: install Node 22 LTS from the Node.js site, then reopen PowerShell and re-run node -v until it shows v22.x.
      - If you really can’t change Node, tell me what node -v prints and I’ll adjust the instructions to an older SPFx version.

  I’ll assume from here you can get Node 22 LTS in place (recommended for SPO + SPFx 1.21).

  ———

  Step 2 – Install Yeoman + SPFx generator (so yo works)

  In PowerShell:

  npm install -g yo gulp @microsoft/generator-sharepoint

  After that:

  yo --version

  If you see a version number, yo is now available.

  ———

  Step 3 – Protect the code we already wrote

  Right now this folder already contains the web part code (src/webparts/universalHtmlViewer/...) plus tests and README.

  Before we let the SPFx generator run (it will create its own src etc.), move this existing src out of the way:

  cd C:\Support\GitHub\UniversalHtmlViewer

  Rename-Item -Path .\src -NewName src-pre-spfx

  (Optional) also keep our README separate so the SPFx README can be generated:

  Rename-Item -Path .\README.md -NewName UniversalHtmlViewer.README.md

  You should now have at least:

  - src-pre-spfx\... (our current code)
  - package.json, jest.config.js, etc.

  ———

  Step 4 – Scaffold the SPFx solution in this folder

  Still in the same folder:

  yo @microsoft/sharepoint

  When prompted, answer like this:

  - What is your solution name?
    UniversalHtmlViewer (or just press Enter if that’s the default)
  - Which baseline packages... ?
    Select: SharePoint Online only (latest)
  - Where do you want to place the files?
    Use the current folder
  - If it warns that the folder is not empty, choose Yes to continue.
  - Do you want to allow the solution to be deployed...
    Up to you; typical is Yes.
  - Which type of client-side component?
    WebPart
  - What is your Web part name?
    UniversalHtmlViewer
  - What is your Web part description?
    Something like Universal HTML Viewer.
  - Which framework would you like to use?
    No JavaScript framework

  Let Yeoman finish (it will run npm install inside this folder; that may take a bit).

  When it’s done you should now have a full SPFx solution structure: config, gulpfile.js, sharepoint, a new src folder, etc.

  ———

  Step 5 – Copy our web part implementation into the SPFx project

  Now we merge our existing code back into the scaffolded project.

  You can do this with Explorer/VS Code or via PowerShell; I’ll describe the structure and the files you should overwrite/copy.

  1. Locate the generator’s web part folder:
      - src\webparts\universalHtmlViewer\  (from the scaffold)
  2. Copy/overwrite the following files from the backup (src-pre-spfx) to the new src:
      - From src-pre-spfx\webparts\universalHtmlViewer\UniversalHtmlViewerWebPart.ts
        → To src\webparts\universalHtmlViewer\UniversalHtmlViewerWebPart.ts
      - From src-pre-spfx\webparts\universalHtmlViewer\UniversalHtmlViewerWebPart.module.scss
        → To src\webparts\universalHtmlViewer\UniversalHtmlViewerWebPart.module.scss
      - From src-pre-spfx\webparts\universalHtmlViewer\UniversalHtmlViewerWebPart.manifest.json
        → To src\webparts\universalHtmlViewer\UniversalHtmlViewerWebPart.manifest.json
        (It’s OK to overwrite the generated manifest; the id can be the one from our file.)
      - Copy helper files (new adds, no overwrite):
          - src-pre-spfx\webparts\universalHtmlViewer\UrlHelper.ts
            → src\webparts\universalHtmlViewer\UrlHelper.ts
          - src-pre-spfx\webparts\universalHtmlViewer\QueryStringHelper.ts
            → src\webparts\universalHtmlViewer\QueryStringHelper.ts
      - Copy tests:
          - The whole folder
            src-pre-spfx\webparts\universalHtmlViewer\__tests__
            → src\webparts\universalHtmlViewer\__tests__
  3. Keep jest.config.js in the project root (we already created it earlier). If it was overwritten or removed, you can restore the file content I gave you before; but likely it’s still there.

  After this, your active code lives under src\webparts\universalHtmlViewer inside the SPFx project.

  You can now delete the backup folder if you like:

  Remove-Item -Recurse -Force .\src-pre-spfx

  (Or keep it until everything works.)

  ———

  Step 6 – Make sure Jest dev dependencies and test script exist

  The scaffolded SPFx solution comes with its own package.json. We want to (re)add the Jest bits.

  In PowerShell:

  npm install --save-dev jest ts-jest @types/jest jest-environment-jsdom

  Then open package.json and make sure you have a test script, e.g.:

  "scripts": {
    "test": "jest",
    ...
  }

  (Keep the other SPFx scripts like gulp bundle, gulp serve etc. that the generator added.)

  jest.config.js should already be present with the content I gave you. If not, let me know and I’ll resend it.

  Now you can run:

  npm test

  and the UrlHelper tests should pass again, now inside the real SPFx project.

  ———

  Step 7 – Run the web part locally

  To run the SPFx workbench and see the Universal HTML Viewer web part:

  gulp serve

  (or npx gulp serve if you don’t have gulp globally)

  Then open (SharePoint Online workbench):

  - https://<your-tenant>.sharepoint.com/sites/<your-site>/_layouts/15/workbench.aspx

  Add the “Universal HTML Viewer” web part to the page, open the property pane, and try the different modes:

  - FullUrl
  - BasePathAndRelativePath
  - BasePathAndDashboardId
  - HeightMode (Fixed vs Viewport)

  ———

  Step 8 – Package and deploy (when you’re ready)

  Once you’re happy and want to deploy to your tenant:

  1. Build and package:

     gulp clean
     gulp bundle --ship
     gulp package-solution --ship
  2. Upload the generated .sppkg from:
      - sharepoint\solution\*.sppkg

     to your tenant app catalog (SharePoint admin center).
  3. Approve and deploy the app.

  After that, you can add the “Universal HTML Viewer” web part to real pages.

  ———

  If you’re OK with this plan, the next concrete thing I’d like you to do is:

  1. Run node -v and tell me the result.
  2. Then run npm install -g yo gulp @microsoft/generator-sharepoint and let me know if yo --version works.

  Once that’s done, I’ll walk you through the yo @microsoft/sharepoint prompts live using exactly what you see on screen.