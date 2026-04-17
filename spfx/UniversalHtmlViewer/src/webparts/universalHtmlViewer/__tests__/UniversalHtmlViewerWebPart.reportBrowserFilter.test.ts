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
  it('reapplies the active filter after rendering a new report list', () => {
    const webPart = Object.create(UniversalHtmlViewerWebPart.prototype) as any;
    webPart.domElement = document.createElement('div');
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
});

export {};
