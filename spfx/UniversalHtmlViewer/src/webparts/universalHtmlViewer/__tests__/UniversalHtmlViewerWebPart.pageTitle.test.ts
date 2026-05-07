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

describe('UniversalHtmlViewerWebPart page title sync', () => {
  let originalTitle: string;

  beforeEach(() => {
    originalTitle = document.title;
    document.title = 'TheDashboardPage';
  });

  afterEach(() => {
    document.title = originalTitle;
  });

  it('syncs the browser tab title from loaded report html when enabled', () => {
    const webPart = Object.create(UniversalHtmlViewerWebPart.prototype) as any;

    webPart.syncPageTitleFromHtml(
      '<html><head><title>Active Directory Overall - Computers</title></head></html>',
      {
        syncPageTitle: true,
      },
    );

    expect(document.title).toBe('Active Directory Overall - Computers');
    expect(webPart.originalDocumentTitle).toBe('TheDashboardPage');
  });

  it('leaves the browser tab title alone when page title sync is disabled', () => {
    const webPart = Object.create(UniversalHtmlViewerWebPart.prototype) as any;

    webPart.syncPageTitleFromHtml(
      '<html><head><title>Active Directory Overall - Computers</title></head></html>',
      {
        syncPageTitle: false,
      },
    );

    expect(document.title).toBe('TheDashboardPage');
    expect(webPart.originalDocumentTitle).toBeUndefined();
  });

  it('restores the original browser tab title during cleanup', () => {
    const webPart = Object.create(UniversalHtmlViewerWebPart.prototype) as any;

    webPart.syncPageTitleFromHtml('<title>Users</title>', {
      syncPageTitle: true,
    });
    webPart.restoreOriginalDocumentTitle();

    expect(document.title).toBe('TheDashboardPage');
    expect(webPart.originalDocumentTitle).toBeUndefined();
  });
});

export {};
