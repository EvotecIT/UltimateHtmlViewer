import { wireNestedIframeHydration } from '../NestedIframeHydrationHelper';
import { UrlValidationOptions } from '../UrlHelper';

function createIframeWithListeners(documentRef: Document): {
  iframe: HTMLIFrameElement;
  addEventListenerSpy: jest.Mock;
  removeEventListenerSpy: jest.Mock;
} {
  const addEventListenerSpy = jest.fn();
  const removeEventListenerSpy = jest.fn();
  const iframe = {
    contentDocument: documentRef,
    addEventListener: addEventListenerSpy,
    removeEventListener: removeEventListenerSpy,
  } as unknown as HTMLIFrameElement;

  return {
    iframe,
    addEventListenerSpy,
    removeEventListenerSpy,
  };
}

describe('NestedIframeHydrationHelper lifecycle', () => {
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
  });

  afterAll(() => {
    consoleErrorSpy.mockRestore();
  });

  afterEach(() => {
    jest.clearAllMocks();
  });

  it('cleanup removes nested frame click handlers and marker attributes', () => {
    document.body.innerHTML = '<iframe src="nested.html"></iframe>';
    const nestedFrame = document.querySelector('iframe') as HTMLIFrameElement;
    const nestedDocument = document.implementation.createHTMLDocument('nested');
    Object.defineProperty(nestedFrame, 'contentDocument', {
      value: nestedDocument,
      configurable: true,
    });

    const removeNestedDocumentListenerSpy = jest.spyOn(nestedDocument, 'removeEventListener');
    const {
      iframe: parentIframe,
      addEventListenerSpy,
      removeEventListenerSpy,
    } = createIframeWithListeners(document);

    const cleanup = wireNestedIframeHydration({
      iframe: parentIframe,
      currentPageUrl: validationOptions.currentPageUrl,
      validationOptions,
      cacheBusterParamName: 'v',
      loadInlineHtml: jest.fn().mockResolvedValue('<html><body>ok</body></html>'),
    });

    expect(addEventListenerSpy).toHaveBeenCalledWith('load', expect.any(Function));
    expect(nestedDocument.documentElement?.getAttribute('data-uhv-inline-nav')).toBe('1');

    cleanup();

    expect(removeEventListenerSpy).toHaveBeenCalledWith('load', expect.any(Function));
    expect(removeNestedDocumentListenerSpy).toHaveBeenCalledWith(
      'click',
      expect.any(Function),
      true,
    );
    expect(nestedDocument.documentElement?.getAttribute('data-uhv-inline-nav')).toBeNull();
  });

  it('removes handlers for nested frames deleted from DOM before final cleanup', async () => {
    document.body.innerHTML = '<iframe src="nested.html"></iframe>';
    const nestedFrame = document.querySelector('iframe') as HTMLIFrameElement;
    const nestedDocument = document.implementation.createHTMLDocument('nested');
    Object.defineProperty(nestedFrame, 'contentDocument', {
      value: nestedDocument,
      configurable: true,
    });

    const removeNestedDocumentListenerSpy = jest.spyOn(nestedDocument, 'removeEventListener');
    const { iframe: parentIframe } = createIframeWithListeners(document);
    const cleanup = wireNestedIframeHydration({
      iframe: parentIframe,
      currentPageUrl: validationOptions.currentPageUrl,
      validationOptions,
      cacheBusterParamName: 'v',
      loadInlineHtml: jest.fn().mockResolvedValue('<html><body>ok</body></html>'),
    });

    expect(nestedDocument.documentElement?.getAttribute('data-uhv-inline-nav')).toBe('1');

    nestedFrame.remove();
    await new Promise<void>((resolve) => {
      setTimeout(resolve, 60);
    });

    expect(removeNestedDocumentListenerSpy).toHaveBeenCalledWith(
      'click',
      expect.any(Function),
      true,
    );
    expect(nestedDocument.documentElement?.getAttribute('data-uhv-inline-nav')).toBeNull();

    cleanup();
  });

  it('debounces mutation-triggered nested scans under bursty DOM updates', async () => {
    jest.useFakeTimers();
    try {
      document.body.innerHTML = '';
      Object.defineProperty(document, 'baseURI', {
        value: 'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx',
        configurable: true,
      });
      const { iframe: parentIframe } = createIframeWithListeners(document);
      const loadInlineHtml = jest.fn().mockResolvedValue('<html><body>ok</body></html>');
      const cleanup = wireNestedIframeHydration({
        iframe: parentIframe,
        currentPageUrl: validationOptions.currentPageUrl,
        validationOptions,
        cacheBusterParamName: 'v',
        loadInlineHtml,
      });

      const firstFrame = document.createElement('iframe');
      firstFrame.setAttribute('src', '/sites/TestSite1/SitePages/frame-one.html');
      const secondFrame = document.createElement('iframe');
      secondFrame.setAttribute('src', '/sites/TestSite1/SitePages/frame-two.html');
      document.body.appendChild(firstFrame);
      document.body.appendChild(secondFrame);

      await Promise.resolve();
      expect(loadInlineHtml).toHaveBeenCalledTimes(0);

      jest.advanceTimersByTime(39);
      await Promise.resolve();
      expect(loadInlineHtml).toHaveBeenCalledTimes(0);

      jest.advanceTimersByTime(1);
      await Promise.resolve();
      await Promise.resolve();
      expect(loadInlineHtml).toHaveBeenCalledTimes(2);

      cleanup();
    } finally {
      jest.runOnlyPendingTimers();
      jest.useRealTimers();
    }
  });

  it('does not apply async nested hydration completion to detached frames', async () => {
    document.body.innerHTML = '<iframe src="/sites/TestSite1/SitePages/nested.html"></iframe>';
    Object.defineProperty(document, 'baseURI', {
      value: 'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx',
      configurable: true,
    });
    const nestedFrame = document.querySelector('iframe') as HTMLIFrameElement;
    const nestedDocument = document.implementation.createHTMLDocument('nested');
    Object.defineProperty(nestedFrame, 'contentDocument', {
      value: nestedDocument,
      configurable: true,
    });

    let resolveInlineHtml: ((value: string) => void) | undefined;
    const loadInlineHtml = jest.fn().mockImplementation(
      () => new Promise<string>((resolve) => {
        resolveInlineHtml = resolve;
      }),
    );
    const { iframe: parentIframe } = createIframeWithListeners(document);
    const cleanup = wireNestedIframeHydration({
      iframe: parentIframe,
      currentPageUrl: validationOptions.currentPageUrl,
      validationOptions,
      cacheBusterParamName: 'v',
      loadInlineHtml,
    });

    expect(loadInlineHtml).toHaveBeenCalled();
    nestedFrame.remove();

    if (!resolveInlineHtml) {
      throw new Error('Expected inline hydration resolver to be initialized.');
    }
    resolveInlineHtml('<html><body>Detached content</body></html>');
    await Promise.resolve();
    await Promise.resolve();

    expect(nestedFrame.getAttribute('data-uhv-nested-state')).not.toBe('done');
    expect((nestedFrame as unknown as { srcdoc?: string }).srcdoc || '').not.toContain(
      'Detached content',
    );

    cleanup();
  });
});
