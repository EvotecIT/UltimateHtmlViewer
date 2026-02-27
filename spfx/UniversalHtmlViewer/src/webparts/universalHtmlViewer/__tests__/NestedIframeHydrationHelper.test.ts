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

      secondLoad.resolve('<html><body>Second</body></html>');
      await Promise.resolve();
      await Promise.resolve();
      expect(nestedFrame.getAttribute('data-uhv-nested-state')).toBe('done');
      expect(nestedFrame.srcdoc).toContain('Second');
      expect(nestedFrame.srcdoc).not.toContain('First');

      cleanup();
    } finally {
      jest.runOnlyPendingTimers();
      jest.useRealTimers();
    }
  });
});
