/* eslint-disable @typescript-eslint/no-explicit-any, @typescript-eslint/no-var-requires */
jest.mock('@microsoft/sp-http', () => ({
  SPHttpClient: {
    configurations: {
      v1: {},
    },
  },
  SPHttpClientResponse: class {},
}));
jest.mock('@microsoft/sp-webpart-base', () => ({
  BaseClientSideWebPart: class {},
}));

const {
  UniversalHtmlViewerWebPartRuntimeBase,
}: {
  UniversalHtmlViewerWebPartRuntimeBase: any;
} = require('../UniversalHtmlViewerWebPartRuntimeBase');

function createRuntimeHarness(domElement: HTMLElement): any {
  const runtime = Object.create(
    UniversalHtmlViewerWebPartRuntimeBase.prototype,
  ) as any;

  runtime.domElement = domElement;
  runtime.refreshInProgress = false;
  runtime.iframeLoadFallbackState = {};
  runtime.hostScrollRestoreState = {};
  runtime.deferredScrollTimeoutState = {};
  runtime.properties = {
    contentDeliveryMode: 'DirectUrl',
  };
  runtime.lastEffectiveProps = runtime.properties;
  runtime.setLoadingVisible = jest.fn();
  runtime.resolveUrlWithCacheBuster = jest
    .fn()
    .mockResolvedValue('https://contoso.sharepoint.com/sites/TestSite1/SitePages/next.html');
  runtime.captureHostScrollPosition = jest.fn().mockReturnValue({ x: 12, y: 34 });
  runtime.restoreHostScrollPosition = jest.fn();
  runtime.updateStatusBadge = jest.fn();
  runtime.trySetIframeSrcDocFromSource = jest.fn().mockResolvedValue(false);

  return runtime;
}

describe('UniversalHtmlViewerWebPartRuntimeBase refreshIframe lifecycle', () => {
  it('clears previous host-scroll restore load listener before wiring a new one', async () => {
    const iframe = document.createElement('iframe');
    const addEventListenerSpy = jest.spyOn(iframe, 'addEventListener');
    const removeEventListenerSpy = jest.spyOn(iframe, 'removeEventListener');
    const container = document.createElement('div');
    container.appendChild(iframe);
    const runtime = createRuntimeHarness(container);

    await (runtime as any).refreshIframe(
      'https://contoso.sharepoint.com/sites/TestSite1/SitePages',
      'None',
      'v',
      'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx',
      false,
      true,
      false,
    );
    const firstLoadHandler = addEventListenerSpy.mock.calls[0]?.[1];

    await (runtime as any).refreshIframe(
      'https://contoso.sharepoint.com/sites/TestSite1/SitePages',
      'None',
      'v',
      'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx',
      false,
      true,
      false,
    );

    expect(firstLoadHandler).toEqual(expect.any(Function));
    expect(removeEventListenerSpy).toHaveBeenCalledWith('load', firstLoadHandler);
    expect(addEventListenerSpy).toHaveBeenCalledTimes(2);
  });

  it('clears pending host-scroll restore load listener during dispose', async () => {
    const iframe = document.createElement('iframe');
    const addEventListenerSpy = jest.spyOn(iframe, 'addEventListener');
    const removeEventListenerSpy = jest.spyOn(iframe, 'removeEventListener');
    const container = document.createElement('div');
    container.appendChild(iframe);
    const runtime = createRuntimeHarness(container);

    await (runtime as any).refreshIframe(
      'https://contoso.sharepoint.com/sites/TestSite1/SitePages',
      'None',
      'v',
      'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx',
      false,
      true,
      false,
    );
    const loadHandler = addEventListenerSpy.mock.calls[0]?.[1];
    (runtime as any).onDispose();

    expect(loadHandler).toEqual(expect.any(Function));
    expect(removeEventListenerSpy).toHaveBeenCalledWith('load', loadHandler);
    expect((runtime as any).hostScrollRestoreState.iframe).toBeUndefined();
    expect((runtime as any).hostScrollRestoreState.loadHandler).toBeUndefined();
  });

  it('clears stale host-scroll restore listener before returning when iframe is missing', async () => {
    const removeEventListenerSpy = jest.fn();
    const staleHandler = jest.fn();
    const staleIframe = {
      removeEventListener: removeEventListenerSpy,
    } as unknown as HTMLIFrameElement;

    const container = document.createElement('div');
    const runtime = createRuntimeHarness(container);
    (runtime as any).hostScrollRestoreState = {
      iframe: staleIframe,
      loadHandler: staleHandler,
    };

    await (runtime as any).refreshIframe(
      'https://contoso.sharepoint.com/sites/TestSite1/SitePages',
      'None',
      'v',
      'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx',
      false,
      true,
      false,
    );

    expect(removeEventListenerSpy).toHaveBeenCalledWith('load', staleHandler);
    expect((runtime as any).hostScrollRestoreState.iframe).toBeUndefined();
    expect((runtime as any).hostScrollRestoreState.loadHandler).toBeUndefined();
  });

  it('clears deferred iframe scroll reset timers during dispose', () => {
    jest.useFakeTimers();
    const clearTimeoutSpy = jest.spyOn(window, 'clearTimeout');
    try {
      const iframeDocument = document.implementation.createHTMLDocument('iframe');
      const iframeWindowScrollToSpy = jest.fn();
      const iframe = {
        contentWindow: {
          scrollTo: iframeWindowScrollToSpy,
        },
        contentDocument: iframeDocument,
      } as unknown as HTMLIFrameElement;
      const container = document.createElement('div');
      const runtime = createRuntimeHarness(container);
      runtime.resetIframeDeepScrollPosition = jest.fn();

      (runtime as any).resetIframeScrollPosition(
        iframe,
        'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Report.html',
      );

      expect(iframeWindowScrollToSpy).toHaveBeenCalledTimes(1);
      expect((runtime as any).deferredScrollTimeoutState.timeoutIds).toHaveLength(4);

      (runtime as any).onDispose();

      expect(clearTimeoutSpy).toHaveBeenCalled();
      expect((runtime as any).deferredScrollTimeoutState.timeoutIds).toEqual([]);

      jest.runOnlyPendingTimers();
      expect(iframeWindowScrollToSpy).toHaveBeenCalledTimes(1);
    } finally {
      clearTimeoutSpy.mockRestore();
      jest.useRealTimers();
    }
  });

  it('clears deferred host scroll restore timers during dispose', () => {
    jest.useFakeTimers();
    const clearTimeoutSpy = jest.spyOn(window, 'clearTimeout');
    const originalScrollTo = window.scrollTo;
    const scrollToSpy = jest.fn();
    Object.defineProperty(window, 'scrollTo', {
      configurable: true,
      value: scrollToSpy,
      writable: true,
    });

    try {
      const container = document.createElement('div');
      const runtime = createRuntimeHarness(container);
      runtime.getPotentialHostScrollContainers = jest.fn().mockReturnValue([]);
      runtime.restoreHostScrollPosition = (
        UniversalHtmlViewerWebPartRuntimeBase.prototype as any
      ).restoreHostScrollPosition.bind(runtime);

      (runtime as any).restoreHostScrollPosition({ x: 12, y: 34 });

      expect(scrollToSpy).toHaveBeenCalledTimes(1);
      expect((runtime as any).deferredScrollTimeoutState.timeoutIds).toHaveLength(4);

      (runtime as any).onDispose();

      expect(clearTimeoutSpy).toHaveBeenCalled();
      expect((runtime as any).deferredScrollTimeoutState.timeoutIds).toEqual([]);

      jest.runOnlyPendingTimers();
      expect(scrollToSpy).toHaveBeenCalledTimes(1);
    } finally {
      Object.defineProperty(window, 'scrollTo', {
        configurable: true,
        value: originalScrollTo,
        writable: true,
      });
      clearTimeoutSpy.mockRestore();
      jest.useRealTimers();
    }
  });
});
