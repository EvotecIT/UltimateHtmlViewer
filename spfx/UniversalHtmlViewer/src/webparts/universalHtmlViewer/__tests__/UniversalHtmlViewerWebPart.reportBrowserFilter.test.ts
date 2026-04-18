/* eslint-disable @typescript-eslint/no-explicit-any, @typescript-eslint/no-var-requires */
jest.mock('@microsoft/sp-core-library', () => ({
  Version: {
    parse: () => ({}),
  },
}));
jest.mock('@microsoft/sp-http', () => ({
  SPHttpClient: {
    configurations: {
      v1: {},
    },
  },
}));
jest.mock('@microsoft/sp-lodash-subset', () => ({
  escape: (value: string): string => value,
}));
jest.mock('@microsoft/sp-property-pane', () => ({
  PropertyPaneDropdown: jest.fn(),
  PropertyPaneSlider: jest.fn(),
  PropertyPaneTextField: jest.fn(),
  PropertyPaneToggle: jest.fn(),
}));
jest.mock('@microsoft/sp-webpart-base', () => ({
  BaseClientSideWebPart: class {},
}));

const {
  default: UniversalHtmlViewerWebPart,
}: {
  default: any;
} = require('../UniversalHtmlViewerWebPart');

describe('UniversalHtmlViewerWebPart report browser filter', () => {
  function createWebPartHarness(): any {
    const webPart = Object.create(UniversalHtmlViewerWebPart.prototype) as any;
    webPart.domElement = document.createElement('div');
    webPart.context = {
      pageContext: {
        web: {
          absoluteUrl: 'https://contoso.sharepoint.com/sites/TestSite1',
        },
      },
      spHttpClient: {},
    };
    return webPart;
  }

  it('reapplies the active filter after rendering a new report list', () => {
    const webPart = createWebPartHarness();
    webPart.domElement.innerHTML = `
      <input data-uhv-report-filter value="match" />
      <div data-uhv-report-status></div>
      <div data-uhv-report-list></div>`;

    webPart.renderReportBrowserItems(
      [
        {
          kind: 'File',
          name: 'Match.html',
          relativePath: 'Reports/Match.html',
          serverRelativeUrl: '/sites/TestSite1/SiteAssets/Reports/Match.html',
          timeLastModified: '2026-04-17T10:00:00Z',
        },
        {
          kind: 'File',
          name: 'Other.html',
          relativePath: 'Reports/Other.html',
          serverRelativeUrl: '/sites/TestSite1/SiteAssets/Reports/Other.html',
          timeLastModified: '2026-04-17T10:00:00Z',
        },
      ],
      '/sites/TestSite1/SiteAssets/Reports',
      '/sites/TestSite1/SiteAssets/Reports',
      'Folders',
    );

    const rows = webPart.domElement.querySelectorAll('[data-uhv-report-row]');
    expect(rows).toHaveLength(2);
    expect(rows[0].style.display).toBe('');
    expect(rows[1].style.display).toBe('none');
  });

  it('does not render the browser for FullUrl even when the legacy flag is enabled', () => {
    const webPart = createWebPartHarness();

    const html = webPart.buildReportBrowserHtml({
      htmlSourceMode: 'FullUrl',
      contentDeliveryMode: 'SharePointFileContent',
      showChrome: true,
      showReportBrowser: true,
    });

    expect(html).toBe('');
  });

  it('uses the configured browser root in SharePointReportBrowser mode', () => {
    const webPart = createWebPartHarness();

    const rootPath = webPart.getEffectiveReportBrowserRootPath(
      {
        htmlSourceMode: 'SharePointReportBrowser',
        reportBrowserRootPath: '/sites/TestSite1/SiteAssets/Reports',
        basePath: '/sites/TestSite1/SiteAssets/Stale',
      },
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/Index.html',
    );

    expect(rootPath).toBe('/sites/TestSite1/SiteAssets/Reports');
  });

  it('resets the picker view when the configured default view changes', () => {
    const webPart = createWebPartHarness();
    webPart.domElement.innerHTML = `
      <div data-uhv-report-browser>
        <button data-uhv-report-view="Folders"></button>
        <button data-uhv-report-view="Files"></button>
        <input data-uhv-report-filter />
        <div data-uhv-report-status></div>
        <div data-uhv-report-list></div>
      </div>`;
    webPart.reportBrowserRootPath = '/sites/TestSite1/SiteAssets/Reports';
    webPart.reportBrowserFolderPath = '/sites/TestSite1/SiteAssets/Reports/Nested';
    webPart.reportBrowserView = 'Folders';
    webPart.reportBrowserDefaultView = 'Folders';
    webPart.loadAndRenderReportBrowser = jest.fn(() => Promise.resolve());

    webPart.attachReportBrowserHandlers(
      '/sites/TestSite1/SiteAssets/Reports/Index.html',
      'None',
      'v',
      'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx',
      {
        htmlSourceMode: 'SharePointReportBrowser',
        contentDeliveryMode: 'SharePointFileContent',
        basePath: '/sites/TestSite1/SiteAssets/Reports',
        reportBrowserDefaultView: 'Files',
      },
    );

    expect(webPart.reportBrowserView).toBe('Files');
    expect(webPart.reportBrowserFolderPath).toBe('/sites/TestSite1/SiteAssets/Reports');
  });
});

export {};
