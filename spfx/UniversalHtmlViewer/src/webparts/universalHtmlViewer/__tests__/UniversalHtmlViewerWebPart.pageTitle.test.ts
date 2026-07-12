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
      originalTitle: 'TheDashboardPage',
      entries: [
        {
          ownerId: 'test-owner',
          syncedTitle: 'Active Directory Overall - Computers',
        },
      ],
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

  it('does not mutate browser history when inline deep links are disabled', () => {
    const webPart = createWebPart();
    webPart.lastEffectiveProps = {
      contentDeliveryMode: 'SharePointFileContent',
      allowQueryStringPageOverride: false,
    };
    const pushStateSpy = jest.spyOn(window.history, 'pushState');

    webPart.onNavigatedToUrl(
      `${window.location.origin}/SiteAssets/Reports/Next.html`,
      `${window.location.origin}/SitePages/Dashboard.aspx`,
    );

    expect(pushStateSpy).not.toHaveBeenCalled();
    pushStateSpy.mockRestore();
  });

  it('uses the configured deep-link parameter when history integration is enabled', () => {
    const webPart = createWebPart();
    webPart.lastEffectiveProps = {
      contentDeliveryMode: 'SharePointFileContent',
      allowQueryStringPageOverride: true,
      inlineDeepLinkParamName: 'viewerTwoPage',
    };
    const pushStateSpy = jest
      .spyOn(window.history, 'pushState')
      .mockImplementation(() => undefined);

    webPart.onNavigatedToUrl(
      `${window.location.origin}/SiteAssets/Reports/Next.html`,
      `${window.location.origin}/SitePages/Dashboard.aspx`,
    );

    expect(pushStateSpy).toHaveBeenCalledWith(
      window.history.state,
      '',
      expect.stringContaining('viewerTwoPage='),
    );
    pushStateSpy.mockRestore();
  });

  it('replaces an active blob frame location without adding iframe history', () => {
    const webPart = createWebPart();
    const replace = jest.fn();
    const iframe = {
      contentWindow: {
        location: {
          replace,
        },
      },
      src: 'blob:https://contoso.sharepoint.com/old',
    } as unknown as HTMLIFrameElement;

    webPart.replaceInlineBlobFrameLocation(
      iframe,
      'blob:https://contoso.sharepoint.com/new',
    );

    expect(replace).toHaveBeenCalledWith('blob:https://contoso.sharepoint.com/new');
    expect(iframe.src).toBe('blob:https://contoso.sharepoint.com/old');
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
      originalTitle: 'TheDashboardPage',
      entries: [
        {
          ownerId: 'second',
          syncedTitle: 'Computers',
        },
      ],
    });
  });

  it('restores the previous active synced title when the latest viewer is cleaned up', () => {
    const firstWebPart = createWebPart('first');
    const secondWebPart = createWebPart('second');

    firstWebPart.syncPageTitleFromHtml('<title>Users</title>', {
      syncPageTitle: true,
    });
    secondWebPart.syncPageTitleFromHtml('<title>Computers</title>', {
      syncPageTitle: true,
    });
    secondWebPart.restoreOriginalDocumentTitle();

    expect(document.title).toBe('Users');
    expect((window as any).__uhvPageTitleSync).toMatchObject({
      originalTitle: 'TheDashboardPage',
      entries: [
        {
          ownerId: 'first',
          syncedTitle: 'Users',
        },
      ],
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

  it('enables validated host deep links for query overrides or the new-tab action', () => {
    const webPart = createWebPart();

    expect(webPart.shouldRewriteInlineAnchorHrefs({})).toBe(false);
    expect(webPart.shouldRewriteInlineAnchorHrefs({ allowQueryStringPageOverride: true })).toBe(
      true,
    );
    expect(
      webPart.shouldRewriteInlineAnchorHrefs({
        allowQueryStringPageOverride: false,
        showOpenInNewTab: true,
      }),
    ).toBe(false);
    expect(
      webPart.shouldRewriteInlineAnchorHrefs({
        allowQueryStringPageOverride: true,
        showOpenInNewTab: true,
        enableExpertSecurityModes: true,
        securityMode: 'AnyHttps',
      }),
    ).toBe(false);
  });

  it('keeps inbound query overrides behind their explicit toggle', () => {
    const webPart = createWebPart();

    expect(
      webPart.shouldAllowInlineDeepLinkOverride({
        allowQueryStringPageOverride: false,
        showOpenInNewTab: true,
      }),
    ).toBe(false);
    expect(
      webPart.shouldAllowInlineDeepLinkOverride({
        allowQueryStringPageOverride: true,
        showOpenInNewTab: false,
      }),
    ).toBe(true);
    expect(
      webPart.shouldAllowInlineDeepLinkOverride({
        allowQueryStringPageOverride: true,
        enableExpertSecurityModes: true,
        securityMode: 'AnyHttps',
      }),
    ).toBe(false);
    expect(
      webPart.shouldAllowInlineDeepLinkOverride({
        contentDeliveryMode: 'DirectUrl',
        allowQueryStringPageOverride: true,
        securityMode: 'Allowlist',
      }),
    ).toBe(false);
  });

  it('disables inline-only query overrides for direct URL presets', () => {
    const webPart = createWebPart();
    const allowlistProps = {
      allowQueryStringPageOverride: true,
    };
    const anyHttpsProps = {
      allowQueryStringPageOverride: true,
    };

    webPart.applyPreset('AllowlistCDN', allowlistProps);
    webPart.applyPreset('AnyHttps', anyHttpsProps);

    expect(allowlistProps).toMatchObject({
      contentDeliveryMode: 'DirectUrl',
      allowQueryStringPageOverride: false,
    });
    expect(anyHttpsProps).toMatchObject({
      contentDeliveryMode: 'DirectUrl',
      allowQueryStringPageOverride: false,
    });
  });

  it('keeps inline new-tab navigation and validated inbound deep links consistent', () => {
    const webPart = createWebPart();
    const props = {
      contentDeliveryMode: 'SharePointFileContent',
      showOpenInNewTab: true,
      allowQueryStringPageOverride: false,
    };

    webPart.normalizeInlineDeepLinkConfiguration(props);

    expect(props.allowQueryStringPageOverride).toBe(true);
    expect(webPart.shouldRewriteInlineAnchorHrefs(props)).toBe(true);
    expect(webPart.shouldAllowInlineDeepLinkOverride(props)).toBe(true);
  });

  it('explains when inline open-in-new-tab cannot honor a disabled query override', () => {
    const webPart = createWebPart();
    webPart.getCurrentPageUrl = jest
      .fn()
      .mockReturnValue('https://contoso.sharepoint.com/sites/Test/SitePages/Dashboard.aspx');

    const chromeHtml = webPart.buildChromeHtml(
      'https://contoso.sharepoint.com/sites/Test/SiteAssets/Reports/Current.html',
      'https://contoso.sharepoint.com/sites/Test/SiteAssets/Reports/Current.html',
      'https://contoso.sharepoint.com/sites/Test/SitePages/Dashboard.aspx',
      { securityMode: 'StrictTenant' },
      'None',
      {
        contentDeliveryMode: 'SharePointFileContent',
        showChrome: true,
        showOpenInNewTab: true,
        allowQueryStringPageOverride: false,
        showRefreshButton: false,
        showStatus: false,
      },
    );

    expect(chromeHtml).toContain('data-uhv-action="open-in-new-tab-disabled"');
    expect(chromeHtml).toContain('aria-disabled="true"');
    expect(chromeHtml).toContain('Enable page query override');
    expect(chromeHtml).not.toContain('data-uhv-action="open-in-new-tab"');
    expect(chromeHtml).not.toContain('uhvPage=');
  });

  it('renders a working inline open-in-new-tab deep link when query overrides are enabled', () => {
    const webPart = createWebPart();
    webPart.getCurrentPageUrl = jest
      .fn()
      .mockReturnValue('https://contoso.sharepoint.com/sites/Test/SitePages/Dashboard.aspx');

    const chromeHtml = webPart.buildChromeHtml(
      'https://contoso.sharepoint.com/sites/Test/SiteAssets/Reports/Current.html',
      'https://contoso.sharepoint.com/sites/Test/SiteAssets/Reports/Current.html',
      'https://contoso.sharepoint.com/sites/Test/SitePages/Dashboard.aspx',
      { securityMode: 'StrictTenant' },
      'None',
      {
        contentDeliveryMode: 'SharePointFileContent',
        showChrome: true,
        showOpenInNewTab: true,
        allowQueryStringPageOverride: true,
        showRefreshButton: false,
        showStatus: false,
      },
    );

    expect(chromeHtml).toContain('data-uhv-action="open-in-new-tab"');
    expect(chromeHtml).toContain('uhvPage=');
    expect(chromeHtml).not.toContain('open-in-new-tab-disabled');
  });
});

export {};
