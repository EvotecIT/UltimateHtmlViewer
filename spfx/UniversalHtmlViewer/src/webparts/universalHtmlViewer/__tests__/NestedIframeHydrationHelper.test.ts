import { wireNestedIframeHydration } from '../NestedIframeHydrationHelper';
import { UrlValidationOptions } from '../UrlHelper';

function createIframeStubWithDocument(
  iframeDocument: Document,
): HTMLIFrameElement {
  const listeners = new Map<string, Array<() => void>>();

  const iframeStub = {
    contentDocument: iframeDocument,
    addEventListener: (eventName: string, handler: () => void): void => {
      const existing = listeners.get(eventName) || [];
      existing.push(handler);
      listeners.set(eventName, existing);
    },
    removeEventListener: (eventName: string, handler: () => void): void => {
      const existing = listeners.get(eventName) || [];
      listeners.set(
        eventName,
        existing.filter((entry) => entry !== handler),
      );
    },
  };

  return iframeStub as unknown as HTMLIFrameElement;
}

describe('NestedIframeHydrationHelper', () => {
  let consoleErrorSpy: jest.SpyInstance;
  const validationOptions: UrlValidationOptions = {
    securityMode: 'StrictTenant',
    currentPageUrl:
      'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx',
    allowHttp: false,
    allowedPathPrefixes: ['/sites/testsite1/sitepages/'],
    allowedFileExtensions: ['.html', '.htm', '.aspx'],
  };

  beforeAll(() => {
    consoleErrorSpy = jest.spyOn(console, 'error').mockImplementation(() => undefined);
    Object.defineProperty(window, 'scrollTo', {
      value: jest.fn(),
      writable: true,
      configurable: true,
    });
  });

  afterAll(() => {
    consoleErrorSpy.mockRestore();
  });

  afterEach(() => {
    jest.clearAllMocks();
  });

  it('hydrates nested iframe URLs using current page URL fallback when base is not absolute', async () => {
    document.body.innerHTML = '<iframe src="reports/nested.html"></iframe>';
    Object.defineProperty(document, 'baseURI', {
      value: '/sites/TestSite1/SiteAssets/Reports/',
      configurable: true,
    });
    const parentIframe = createIframeStubWithDocument(document);
    const loadInlineHtml = jest
      .fn()
      .mockResolvedValue('<html><body>Nested content</body></html>');

    const cleanup = wireNestedIframeHydration({
      iframe: parentIframe,
      currentPageUrl: validationOptions.currentPageUrl,
      validationOptions,
      cacheBusterParamName: 'v',
      loadInlineHtml,
    });

    await Promise.resolve();
    await Promise.resolve();

    const nestedFrame = document.querySelector('iframe') as HTMLIFrameElement;
    expect(loadInlineHtml).toHaveBeenCalled();
    expect(loadInlineHtml.mock.calls[0][0]).toBe(
      'https://contoso.sharepoint.com/sites/TestSite1/SitePages/reports/nested.html',
    );
    expect(loadInlineHtml.mock.calls[0][1]).toBe(
      'https://contoso.sharepoint.com/sites/TestSite1/SitePages/reports/nested.html',
    );
    expect(nestedFrame.getAttribute('data-uhv-nested-state')).toBe('done');
    expect(nestedFrame.srcdoc).toContain('Nested content');

    cleanup();
  });

  it('does not hydrate nested iframes that resolve outside current tenant host', async () => {
    document.body.innerHTML = '<iframe src="https://example.org/report.html"></iframe>';
    const parentIframe = createIframeStubWithDocument(document);
    const loadInlineHtml = jest.fn().mockResolvedValue('<html></html>');

    const cleanup = wireNestedIframeHydration({
      iframe: parentIframe,
      currentPageUrl: validationOptions.currentPageUrl,
      validationOptions,
      cacheBusterParamName: 'v',
      loadInlineHtml,
    });

    await Promise.resolve();

    const nestedFrame = document.querySelector('iframe') as HTMLIFrameElement;
    expect(loadInlineHtml).not.toHaveBeenCalled();
    expect(nestedFrame.getAttribute('data-uhv-nested-state')).toBeNull();

    cleanup();
  });

  it('ignores stale in-flight hydration result after nested iframe source changes', async () => {
    jest.useFakeTimers();
    try {
      document.body.innerHTML = '<iframe src="/sites/TestSite1/SitePages/first.html"></iframe>';
      Object.defineProperty(document, 'baseURI', {
        value: 'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx',
        configurable: true,
      });

      const nestedFrame = document.querySelector('iframe') as HTMLIFrameElement;
      const pendingLoads: Array<{
        sourceUrl: string;
        resolve: (value: string | PromiseLike<string | undefined> | undefined) => void;
      }> = [];
      const onDiagnosticsEvent = jest.fn();
      const loadInlineHtml = jest.fn().mockImplementation((sourceUrl: string) => {
        return new Promise<string | undefined>((resolve) => {
          pendingLoads.push({
            sourceUrl,
            resolve,
          });
        });
      });

      const parentIframe = createIframeStubWithDocument(document);
      const cleanup = wireNestedIframeHydration({
        iframe: parentIframe,
        currentPageUrl: validationOptions.currentPageUrl,
        validationOptions,
        cacheBusterParamName: 'v',
        loadInlineHtml,
        onDiagnosticsEvent,
      });

      await Promise.resolve();
      await Promise.resolve();
      expect(loadInlineHtml).toHaveBeenCalledTimes(1);

      nestedFrame.setAttribute('src', '/sites/TestSite1/SitePages/second.html');
      await Promise.resolve();
      jest.advanceTimersByTime(40);
      await Promise.resolve();
      await Promise.resolve();

      expect(loadInlineHtml).toHaveBeenCalledTimes(2);
      const firstLoad = pendingLoads[0];
      const secondLoad = pendingLoads[1];
      expect(firstLoad.sourceUrl).toContain('/first.html');
      expect(secondLoad.sourceUrl).toContain('/second.html');

      firstLoad.resolve('<html><body>First</body></html>');
      await Promise.resolve();
      await Promise.resolve();
      expect(nestedFrame.getAttribute('data-uhv-nested-state')).toBe('processing');
      expect(nestedFrame.srcdoc || '').not.toContain('First');
      expect(onDiagnosticsEvent).toHaveBeenCalledWith('nestedHydrationStaleResultIgnored');

      secondLoad.resolve('<html><body>Second</body></html>');
      await Promise.resolve();
      await Promise.resolve();
      expect(nestedFrame.getAttribute('data-uhv-nested-state')).toBe('done');
      expect(nestedFrame.srcdoc).toContain('Second');
      expect(nestedFrame.srcdoc).not.toContain('First');
      expect(onDiagnosticsEvent).toHaveBeenCalledWith('nestedHydrationSucceeded');

      cleanup();
    } finally {
      jest.runOnlyPendingTimers();
      jest.useRealTimers();
    }
  });

  it('ignores stale in-flight nested click navigation result when a newer click wins', async () => {
    document.body.innerHTML =
      '<iframe src="https://contoso.sharepoint.com/sites/TestSite1/SitePages/seed.html"></iframe>';
    const nestedFrame = document.querySelector('iframe') as HTMLIFrameElement;
    const nestedDocument = document.implementation.createHTMLDocument('nested');
    nestedDocument.body.innerHTML = [
      '<a id="first-link" href="https://contoso.sharepoint.com/sites/TestSite1/SitePages/first-click.html">First</a>',
      '<a id="second-link" href="https://contoso.sharepoint.com/sites/TestSite1/SitePages/second-click.html">Second</a>',
    ].join('');
    Object.defineProperty(nestedDocument, 'baseURI', {
      value: 'https://contoso.sharepoint.com/sites/TestSite1/SitePages/seed.html',
      configurable: true,
    });
    Object.defineProperty(nestedFrame, 'contentDocument', {
      value: nestedDocument,
      configurable: true,
    });

    const pendingLoads = new Map<
      string,
      (value: string | PromiseLike<string | undefined> | undefined) => void
    >();
    const onDiagnosticsEvent = jest.fn();
    const loadInlineHtml = jest.fn().mockImplementation((sourceUrl: string) => {
      if (sourceUrl.includes('/seed.html')) {
        return Promise.resolve('<html><body>Seed</body></html>');
      }

      return new Promise<string | undefined>((resolve) => {
        pendingLoads.set(sourceUrl, resolve);
      });
    });
    const parentIframe = createIframeStubWithDocument(document);
    const cleanup = wireNestedIframeHydration({
      iframe: parentIframe,
      currentPageUrl: validationOptions.currentPageUrl,
      validationOptions,
      cacheBusterParamName: 'v',
      loadInlineHtml,
      onDiagnosticsEvent,
    });

    await Promise.resolve();
    await Promise.resolve();

    const firstLink = nestedDocument.getElementById('first-link') as HTMLAnchorElement;
    const secondLink = nestedDocument.getElementById('second-link') as HTMLAnchorElement;
    firstLink.dispatchEvent(
      new MouseEvent('click', {
        bubbles: true,
        cancelable: true,
        button: 0,
      }),
    );
    secondLink.dispatchEvent(
      new MouseEvent('click', {
        bubbles: true,
        cancelable: true,
        button: 0,
      }),
    );

    expect(loadInlineHtml).toHaveBeenCalledTimes(3);
    const firstClickTarget = loadInlineHtml.mock.calls[1][0] as string;
    const secondClickTarget = loadInlineHtml.mock.calls[2][0] as string;
    expect(firstClickTarget).toContain('/first-click.html');
    expect(secondClickTarget).toContain('/second-click.html');

    const resolveFirst = pendingLoads.get(firstClickTarget);
    const resolveSecond = pendingLoads.get(secondClickTarget);
    if (!resolveFirst || !resolveSecond) {
      throw new Error('Expected pending nested click loads to be registered.');
    }

    resolveFirst('<html><body>First click content</body></html>');
    await Promise.resolve();
    await Promise.resolve();

    expect(nestedFrame.getAttribute('data-uhv-nested-state')).toBe('processing');
    expect(nestedFrame.srcdoc || '').not.toContain('First click content');
    expect(onDiagnosticsEvent).toHaveBeenCalledWith('nestedNavigationStaleResultIgnored');

    resolveSecond('<html><body>Second click content</body></html>');
    await Promise.resolve();
    await Promise.resolve();

    expect(nestedFrame.getAttribute('data-uhv-nested-state')).toBe('done');
    expect(nestedFrame.srcdoc).toContain('Second click content');
    expect(nestedFrame.srcdoc).not.toContain('First click content');
    expect(onDiagnosticsEvent).toHaveBeenCalledWith('nestedNavigationStarted');
    expect(onDiagnosticsEvent).toHaveBeenCalledWith('nestedNavigationSucceeded');

    cleanup();
  });

  it('wires nested iframe window click listener and intercepts window-captured navigation clicks', async () => {
    document.body.innerHTML =
      '<iframe src="https://contoso.sharepoint.com/sites/TestSite1/SitePages/seed.html"></iframe>';
    const nestedFrame = document.querySelector('iframe') as HTMLIFrameElement;
    const nestedDocument = document.implementation.createHTMLDocument('nested-window');
    nestedDocument.body.innerHTML =
      '<a id="window-link" href="https://contoso.sharepoint.com/sites/TestSite1/SitePages/window-click.html">Window click</a>';
    Object.defineProperty(nestedDocument, 'baseURI', {
      value: 'https://contoso.sharepoint.com/sites/TestSite1/SitePages/seed.html',
      configurable: true,
    });
    Object.defineProperty(nestedFrame, 'contentDocument', {
      value: nestedDocument,
      configurable: true,
    });

    const nestedWindow = nestedFrame.contentWindow as Window;
    let windowClickHandler: ((event: Event) => void) | undefined;
    const originalAddWindowListener = nestedWindow.addEventListener.bind(nestedWindow);
    const addWindowListener = jest
      .spyOn(nestedWindow, 'addEventListener')
      .mockImplementation((eventName: string, handler: EventListenerOrEventListenerObject, options?: boolean | AddEventListenerOptions) => {
        if (eventName === 'click' && typeof handler === 'function') {
          windowClickHandler = handler as (event: Event) => void;
        }
        originalAddWindowListener(eventName, handler, options);
      });
    const removeWindowListener = jest.spyOn(nestedWindow, 'removeEventListener');

    const loadInlineHtml = jest.fn().mockImplementation((sourceUrl: string) => {
      if (sourceUrl.includes('/seed.html')) {
        return Promise.resolve('<html><body>Seed</body></html>');
      }
      return Promise.resolve('<html><body>Window click content</body></html>');
    });
    const parentIframe = createIframeStubWithDocument(document);
    const cleanup = wireNestedIframeHydration({
      iframe: parentIframe,
      currentPageUrl: validationOptions.currentPageUrl,
      validationOptions,
      cacheBusterParamName: 'v',
      loadInlineHtml,
    });

    await Promise.resolve();
    await Promise.resolve();

    expect(addWindowListener).toHaveBeenCalledWith('click', expect.any(Function), true);
    if (!windowClickHandler) {
      throw new Error('Expected nested iframe window click handler to be registered.');
    }

    const anchor = nestedDocument.getElementById('window-link') as HTMLAnchorElement;
    const clickEvent = {
      button: 0,
      bubbles: true,
      cancelable: true,
      metaKey: false,
      ctrlKey: false,
      shiftKey: false,
      altKey: false,
      target: anchor,
      preventDefault: jest.fn(),
      stopPropagation: jest.fn(),
      stopImmediatePropagation: jest.fn(),
      cancelBubble: false,
      returnValue: true,
    } as unknown as Event;
    windowClickHandler(clickEvent);

    await Promise.resolve();
    await Promise.resolve();

    expect(loadInlineHtml).toHaveBeenCalledTimes(2);
    expect(loadInlineHtml.mock.calls[1][0]).toContain('/window-click.html');
    expect(nestedFrame.srcdoc).toContain('Window click content');
    expect(clickEvent.preventDefault).toHaveBeenCalledTimes(1);
    expect(clickEvent.stopPropagation).toHaveBeenCalledTimes(1);
    expect(clickEvent.stopImmediatePropagation).toHaveBeenCalledTimes(1);
    expect((clickEvent as Event & { cancelBubble?: boolean }).cancelBubble).toBe(true);
    expect((clickEvent as Event & { returnValue?: boolean }).returnValue).toBe(false);

    cleanup();

    expect(removeWindowListener).toHaveBeenCalledWith('click', expect.any(Function), true);

    addWindowListener.mockRestore();
    removeWindowListener.mockRestore();
  });
});
