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

describe('UniversalHtmlViewerWebPart deep-link scroll lock decision', () => {
  it('enables initial scroll lock only when a deep link is actually applied in inline mode', () => {
    const webPart = Object.create(UniversalHtmlViewerWebPart.prototype) as any;

    const shouldLock = webPart.shouldApplyInitialDeepLinkScrollLock(
      'SharePointFileContent',
      {
        deepLinkedUrl:
          'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/Ops.html',
      },
    );
    const shouldNotLockWithoutDeepLink = webPart.shouldApplyInitialDeepLinkScrollLock(
      'SharePointFileContent',
      {
        deepLinkedUrl: undefined,
      },
    );
    const shouldNotLockForDirectMode = webPart.shouldApplyInitialDeepLinkScrollLock(
      'DirectUrl',
      {
        deepLinkedUrl:
          'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/Ops.html',
      },
    );

    expect(shouldLock).toBe(true);
    expect(shouldNotLockWithoutDeepLink).toBe(false);
    expect(shouldNotLockForDirectMode).toBe(false);
  });
});

export {};
