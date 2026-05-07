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

  const createWebPart = (ownerId: string = 'test-owner'): any => {
    const webPart = Object.create(UniversalHtmlViewerWebPart.prototype) as any;
    webPart.pageTitleSyncOwnerId = ownerId;
    return webPart;
  };

  beforeEach(() => {
    originalTitle = document.title;
    document.title = 'TheDashboardPage';
    delete (window as any).__uhvPageTitleSync;
  });

  afterEach(() => {
    document.title = originalTitle;
    delete (window as any).__uhvPageTitleSync;
  });

  it('syncs the browser tab title from loaded report html when enabled', () => {
    const webPart = createWebPart();

    webPart.syncPageTitleFromHtml(
      '<html><head><title>Active Directory Overall - Computers</title></head></html>',
      {
        syncPageTitle: true,
      },
    );

    expect(document.title).toBe('Active Directory Overall - Computers');
    expect((window as any).__uhvPageTitleSync).toMatchObject({
      ownerId: 'test-owner',
      originalTitle: 'TheDashboardPage',
      syncedTitle: 'Active Directory Overall - Computers',
    });
  });

  it('leaves the browser tab title alone when page title sync is disabled', () => {
    const webPart = createWebPart();

    webPart.syncPageTitleFromHtml(
      '<html><head><title>Active Directory Overall - Computers</title></head></html>',
      {
        syncPageTitle: false,
      },
    );

    expect(document.title).toBe('TheDashboardPage');
    expect((window as any).__uhvPageTitleSync).toBeUndefined();
  });

  it('restores the original browser tab title during cleanup', () => {
    const webPart = createWebPart();

    webPart.syncPageTitleFromHtml('<title>Users</title>', {
      syncPageTitle: true,
    });
    webPart.restoreOriginalDocumentTitle();

    expect(document.title).toBe('TheDashboardPage');
    expect((window as any).__uhvPageTitleSync).toBeUndefined();
  });

  it('restores the original browser tab title when a synced report has no title', () => {
    const webPart = createWebPart();

    webPart.syncPageTitleFromHtml('<title>Users</title>', {
      syncPageTitle: true,
    });
    webPart.syncPageTitleFromHtml('<html><body>No title</body></html>', {
      syncPageTitle: true,
    });

    expect(document.title).toBe('TheDashboardPage');
    expect((window as any).__uhvPageTitleSync).toBeUndefined();
  });

  it('does not restore a title owned by another active viewer instance', () => {
    const firstWebPart = createWebPart('first');
    const secondWebPart = createWebPart('second');

    firstWebPart.syncPageTitleFromHtml('<title>Users</title>', {
      syncPageTitle: true,
    });
    secondWebPart.syncPageTitleFromHtml('<title>Computers</title>', {
      syncPageTitle: true,
    });
    firstWebPart.restoreOriginalDocumentTitle();

    expect(document.title).toBe('Computers');
    expect((window as any).__uhvPageTitleSync).toMatchObject({
      ownerId: 'second',
      originalTitle: 'TheDashboardPage',
      syncedTitle: 'Computers',
    });
  });

  it('treats title sync as active only for inline content modes', () => {
    const webPart = createWebPart();

    expect(webPart.shouldSyncPageTitle({ syncPageTitle: true }, 'SharePointFileContent')).toBe(
      true,
    );
    expect(webPart.shouldSyncPageTitle({ syncPageTitle: true }, 'SharePointFileBlobUrl')).toBe(
      true,
    );
    expect(webPart.shouldSyncPageTitle({ syncPageTitle: true }, 'DirectUrl')).toBe(false);
    expect(webPart.shouldSyncPageTitle({ syncPageTitle: false }, 'SharePointFileContent')).toBe(
      false,
    );
  });

  it('enables host deep-link anchor rewrites only when query overrides are active', () => {
    const webPart = createWebPart();

    expect(webPart.shouldRewriteInlineAnchorHrefs({})).toBe(false);
    expect(webPart.shouldRewriteInlineAnchorHrefs({ allowQueryStringPageOverride: true })).toBe(
      true,
    );
    expect(
      webPart.shouldRewriteInlineAnchorHrefs({
        allowQueryStringPageOverride: true,
        enableExpertSecurityModes: true,
        securityMode: 'AnyHttps',
      }),
    ).toBe(false);
  });
});

export {};
