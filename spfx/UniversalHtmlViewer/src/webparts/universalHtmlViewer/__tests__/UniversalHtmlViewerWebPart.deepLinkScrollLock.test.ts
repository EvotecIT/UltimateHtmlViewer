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

const createDeepLinkScrollLockHarness = (): any => {
  const webPart = Object.create(UniversalHtmlViewerWebPart.prototype) as any;

  webPart.domElement = {
    querySelector: jest.fn().mockReturnValue(undefined),
  };
  webPart.deepLinkScrollLockDiagnostics = {
    starts: 0,
    releases: 0,
    releasedByAutoStable: 0,
    releasedByUserInteraction: 0,
    releasedByTimeout: 0,
    releasedByManual: 0,
    releasedByReplace: 0,
    releasedByDispose: 0,
    active: false,
    lastReleaseReason: '',
    lastLockDurationMs: 0,
  };
  webPart.getPotentialHostScrollContainers = jest.fn().mockReturnValue([]);
  webPart.getInlineDeepLinkFrameMetrics = jest.fn().mockReturnValue(undefined);
  webPart.getDeepLinkScrollOffsets = jest.fn().mockReturnValue({
    windowTop: 0,
    hostMaxTop: 0,
    iframeTop: 0,
    maxOffset: 0,
  });
  webPart.forceHostScrollTop = jest.fn();
  webPart.restoreHostScrollPosition = jest.fn();
  webPart.resetInlineIframeScrollPositionForDeepLink = jest.fn();
  webPart.isScrollTraceEnabled = jest.fn().mockReturnValue(false);
  webPart.describeScrollElement = jest.fn().mockReturnValue('host');
  webPart.emitScrollTrace = jest.fn();

  return webPart;
};

describe('UniversalHtmlViewerWebPart deep-link scroll lock diagnostics', () => {
  beforeEach(() => {
    jest.useFakeTimers();
  });

  afterEach(() => {
    jest.runOnlyPendingTimers();
    jest.useRealTimers();
  });

  it('records dispose reason when cleanup is triggered during disposal flow', () => {
    const webPart = createDeepLinkScrollLockHarness();

    webPart.applyInitialDeepLinkScrollLock();
    expect(webPart.deepLinkScrollLockDiagnostics.starts).toBe(1);
    expect(webPart.deepLinkScrollLockDiagnostics.active).toBe(true);

    webPart.clearInitialDeepLinkScrollLock('dispose');
    expect(webPart.deepLinkScrollLockDiagnostics.releases).toBe(1);
    expect(webPart.deepLinkScrollLockDiagnostics.releasedByDispose).toBe(1);
    expect(webPart.deepLinkScrollLockDiagnostics.lastReleaseReason).toBe('dispose');
    expect(webPart.deepLinkScrollLockDiagnostics.active).toBe(false);
    expect(webPart.initialDeepLinkScrollLockCleanup).toBeUndefined();
  });

  it('records replace reason when a new lock replaces an active lock', () => {
    const webPart = createDeepLinkScrollLockHarness();

    webPart.applyInitialDeepLinkScrollLock();
    webPart.applyInitialDeepLinkScrollLock();

    expect(webPart.deepLinkScrollLockDiagnostics.starts).toBe(2);
    expect(webPart.deepLinkScrollLockDiagnostics.releases).toBe(1);
    expect(webPart.deepLinkScrollLockDiagnostics.releasedByReplace).toBe(1);
    expect(webPart.deepLinkScrollLockDiagnostics.lastReleaseReason).toBe('replace');
    expect(webPart.deepLinkScrollLockDiagnostics.active).toBe(true);
    expect(webPart.initialDeepLinkScrollLockCleanup).toBeDefined();

    webPart.clearInitialDeepLinkScrollLock();
  });

  it('records timeout reason when lock does not reach stable scroll state', () => {
    const webPart = createDeepLinkScrollLockHarness();
    webPart.getDeepLinkScrollOffsets = jest.fn().mockReturnValue({
      windowTop: 3,
      hostMaxTop: 3,
      iframeTop: 3,
      maxOffset: 3,
    });

    webPart.applyInitialDeepLinkScrollLock();
    jest.advanceTimersByTime(12050);

    expect(webPart.deepLinkScrollLockDiagnostics.starts).toBe(1);
    expect(webPart.deepLinkScrollLockDiagnostics.releases).toBe(1);
    expect(webPart.deepLinkScrollLockDiagnostics.releasedByTimeout).toBe(1);
    expect(webPart.deepLinkScrollLockDiagnostics.lastReleaseReason).toBe('timeout');
    expect(webPart.deepLinkScrollLockDiagnostics.active).toBe(false);
    expect(webPart.initialDeepLinkScrollLockCleanup).toBeUndefined();
  });

  it('records auto-stable reason when scroll remains stable after minimum lock duration', () => {
    const webPart = createDeepLinkScrollLockHarness();

    webPart.applyInitialDeepLinkScrollLock();
    jest.advanceTimersByTime(1500);

    expect(webPart.deepLinkScrollLockDiagnostics.starts).toBe(1);
    expect(webPart.deepLinkScrollLockDiagnostics.releases).toBe(1);
    expect(webPart.deepLinkScrollLockDiagnostics.releasedByAutoStable).toBe(1);
    expect(webPart.deepLinkScrollLockDiagnostics.lastReleaseReason).toBe('auto-stable');
    expect(webPart.deepLinkScrollLockDiagnostics.active).toBe(false);
    expect(webPart.initialDeepLinkScrollLockCleanup).toBeUndefined();
  });

  it('records manual reason when cleanup runs without an explicit reason', () => {
    const webPart = createDeepLinkScrollLockHarness();

    webPart.applyInitialDeepLinkScrollLock();
    webPart.clearInitialDeepLinkScrollLock();

    expect(webPart.deepLinkScrollLockDiagnostics.starts).toBe(1);
    expect(webPart.deepLinkScrollLockDiagnostics.releases).toBe(1);
    expect(webPart.deepLinkScrollLockDiagnostics.releasedByManual).toBe(1);
    expect(webPart.deepLinkScrollLockDiagnostics.lastReleaseReason).toBe('manual');
    expect(webPart.deepLinkScrollLockDiagnostics.active).toBe(false);
    expect(webPart.initialDeepLinkScrollLockCleanup).toBeUndefined();
  });
});

export {};
