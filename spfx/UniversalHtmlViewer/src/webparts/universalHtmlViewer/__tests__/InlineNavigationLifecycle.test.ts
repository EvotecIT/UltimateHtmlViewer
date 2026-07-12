import {
  wireInlineIframeNavigation,
} from '../InlineNavigationHelper';
import { INLINE_HOST_PAGE_URL_CHANGED_EVENT } from '../InlineHostPageUrlSyncHelper';
import { UrlValidationOptions } from '../UrlHelper';

describe('wireInlineIframeNavigation lifecycle', () => {
  const validationOptions: UrlValidationOptions = {
    securityMode: 'StrictTenant',
    currentPageUrl:
      'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx',
    allowHttp: false,
    allowedPathPrefixes: ['/sites/testsite1/siteassets/'],
    allowedFileExtensions: ['.html', '.htm', '.aspx'],
  };

  afterEach(() => {
    if (window.location.hash) {
      window.history.replaceState(null, document.title, `${window.location.pathname}${window.location.search}`);
    }
  });

  it('returns cleanup that unwires iframe/document listeners and marker attribute', () => {
    const iframeDocument = document.implementation.createHTMLDocument('iframe');
    const addLoadListener = jest.fn();
    const removeLoadListener = jest.fn();
    const removeDocumentListenerSpy = jest.spyOn(iframeDocument, 'removeEventListener');

    const iframeStub = {
      contentDocument: iframeDocument,
      addEventListener: addLoadListener,
      removeEventListener: removeLoadListener,
    } as unknown as HTMLIFrameElement;

    const cleanup = wireInlineIframeNavigation({
      iframe: iframeStub,
      currentPageUrl: validationOptions.currentPageUrl,
      validationOptions,
      cacheBusterParamName: 'v',
      onNavigate: () => {
        return;
      },
    });

    expect(addLoadListener).toHaveBeenCalledWith('load', expect.any(Function));
    expect(iframeDocument.documentElement?.getAttribute('data-uhv-inline-nav')).toBe('1');

    cleanup();

    expect(removeLoadListener).toHaveBeenCalledWith('load', expect.any(Function));
    expect(removeDocumentListenerSpy).toHaveBeenCalledWith(
      'click',
      expect.any(Function),
      true,
    );
    expect(iframeDocument.documentElement?.getAttribute('data-uhv-inline-nav')).toBeNull();
  });

  it('replaces iframe-window click listeners across document reloads and cleans up the last handler', () => {
    const initialDocument = document.implementation.createHTMLDocument('iframe-initial');
    const reloadedDocument = document.implementation.createHTMLDocument('iframe-reloaded');
    let activeDocument: Document = initialDocument;
    const addLoadListener = jest.fn();
    const removeLoadListener = jest.fn();
    const addWindowListener = jest.fn();
    const removeWindowListener = jest.fn();
    const iframeWindowStub = {
      addEventListener: addWindowListener,
      removeEventListener: removeWindowListener,
    } as unknown as Window;

    const iframeStub = {
      get contentDocument(): Document {
        return activeDocument;
      },
      contentWindow: iframeWindowStub,
      addEventListener: addLoadListener,
      removeEventListener: removeLoadListener,
    } as unknown as HTMLIFrameElement;

    const cleanup = wireInlineIframeNavigation({
      iframe: iframeStub,
      currentPageUrl: validationOptions.currentPageUrl,
      validationOptions,
      cacheBusterParamName: 'v',
      onNavigate: () => {
        return;
      },
    });

    expect(addWindowListener).toHaveBeenCalledTimes(3);
    expect(addWindowListener).toHaveBeenCalledWith(
      'click',
      expect.any(Function),
      true,
    );
    expect(addWindowListener).toHaveBeenCalledWith(
      'mousedown',
      expect.any(Function),
      true,
    );
    expect(addWindowListener).toHaveBeenCalledWith(
      'pointerdown',
      expect.any(Function),
      true,
    );

    const loadHandler = addLoadListener.mock.calls[1][1] as () => void;
    activeDocument = reloadedDocument;
    loadHandler();

    expect(addWindowListener).toHaveBeenCalledTimes(6);
    expect(removeWindowListener).toHaveBeenCalledTimes(3);

    cleanup();

    expect(removeWindowListener).toHaveBeenCalledTimes(6);
    expect(removeLoadListener).toHaveBeenCalledWith('load', loadHandler);
  });

  it('suppresses intercepted click propagation to avoid fallback browser navigation', () => {
    const iframeDocument = document.implementation.createHTMLDocument('iframe-click');
    iframeDocument.body.innerHTML =
      '<a id="report-link" href="https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/next.html">Next</a>';
    const addLoadListener = jest.fn();
    const removeLoadListener = jest.fn();
    const addDocumentListenerSpy = jest.spyOn(iframeDocument, 'addEventListener');
    const onNavigate = jest.fn();
    const iframeStub = {
      contentDocument: iframeDocument,
      addEventListener: addLoadListener,
      removeEventListener: removeLoadListener,
    } as unknown as HTMLIFrameElement;

    const cleanup = wireInlineIframeNavigation({
      iframe: iframeStub,
      currentPageUrl: validationOptions.currentPageUrl,
      validationOptions,
      cacheBusterParamName: 'v',
      onNavigate,
    });

    const clickRegistration = addDocumentListenerSpy.mock.calls.find(
      (call) => call[0] === 'click',
    );
    expect(clickRegistration).toBeDefined();
    const clickHandler = clickRegistration?.[1] as (event: Event) => void;
    const anchor = iframeDocument.getElementById('report-link') as HTMLAnchorElement;
    const syntheticEvent = {
      button: 0,
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

    clickHandler(syntheticEvent);

    expect(onNavigate).toHaveBeenCalledWith(
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/next.html',
    );
    expect(syntheticEvent.preventDefault).toHaveBeenCalledTimes(1);
    expect(syntheticEvent.stopPropagation).toHaveBeenCalledTimes(1);
    expect(syntheticEvent.stopImmediatePropagation).toHaveBeenCalledTimes(1);
    expect((syntheticEvent as Event & { cancelBubble?: boolean }).cancelBubble).toBe(true);
    expect((syntheticEvent as Event & { returnValue?: boolean }).returnValue).toBe(false);

    cleanup();
    addDocumentListenerSpy.mockRestore();
  });

  it('keeps fragment-only anchors inside the iframe document', () => {
    const iframeDocument = document.implementation.createHTMLDocument('iframe-fragment');
    iframeDocument.body.innerHTML =
      '<a id="section-link" href="#what">What it does</a><section id="what">Details</section>';
    const addLoadListener = jest.fn();
    const removeLoadListener = jest.fn();
    const addDocumentListenerSpy = jest.spyOn(iframeDocument, 'addEventListener');
    const onNavigate = jest.fn();
    const iframeStub = {
      contentDocument: iframeDocument,
      ownerDocument: document,
      getBoundingClientRect: () =>
        ({
          top: 240,
          left: 0,
          bottom: 1040,
          right: 800,
          width: 800,
          height: 800,
        }) as DOMRect,
      addEventListener: addLoadListener,
      removeEventListener: removeLoadListener,
    } as unknown as HTMLIFrameElement;

    const targetSection = iframeDocument.getElementById('what') as HTMLElement;
    targetSection.scrollIntoView = jest.fn();

    const cleanup = wireInlineIframeNavigation({
      iframe: iframeStub,
      currentPageUrl: validationOptions.currentPageUrl,
      validationOptions,
      cacheBusterParamName: 'v',
      onNavigate,
    });

    const clickRegistration = addDocumentListenerSpy.mock.calls.find(
      (call) => call[0] === 'click',
    );
    const clickHandler = clickRegistration?.[1] as (event: Event) => void;
    const anchor = iframeDocument.getElementById('section-link') as HTMLAnchorElement;
    const syntheticEvent = {
      button: 0,
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

    clickHandler(syntheticEvent);

    expect(onNavigate).not.toHaveBeenCalled();
    expect(syntheticEvent.preventDefault).toHaveBeenCalledTimes(1);
    expect(syntheticEvent.stopPropagation).toHaveBeenCalledTimes(1);
    expect(syntheticEvent.stopImmediatePropagation).toHaveBeenCalledTimes(1);
    expect(targetSection.scrollIntoView).toHaveBeenCalledTimes(1);
    expect((syntheticEvent as Event & { cancelBubble?: boolean }).cancelBubble).toBe(true);
    expect((syntheticEvent as Event & { returnValue?: boolean }).returnValue).toBe(false);

    cleanup();
    addDocumentListenerSpy.mockRestore();
  });

  it('leaves scripted placeholder hash anchors to the iframe page', () => {
    const iframeDocument = document.implementation.createHTMLDocument('iframe-script-control');
    iframeDocument.body.innerHTML = '<a id="toggle" href="#">Toggle</a>';
    const addLoadListener = jest.fn();
    const removeLoadListener = jest.fn();
    const addDocumentListenerSpy = jest.spyOn(iframeDocument, 'addEventListener');
    const onNavigate = jest.fn();
    const iframeStub = {
      contentDocument: iframeDocument,
      ownerDocument: document,
      addEventListener: addLoadListener,
      removeEventListener: removeLoadListener,
    } as unknown as HTMLIFrameElement;

    const cleanup = wireInlineIframeNavigation({
      iframe: iframeStub,
      currentPageUrl: validationOptions.currentPageUrl,
      validationOptions,
      cacheBusterParamName: 'v',
      onNavigate,
    });

    const clickRegistration = addDocumentListenerSpy.mock.calls.find(
      (call) => call[0] === 'click',
    );
    const clickHandler = clickRegistration?.[1] as (event: Event) => void;
    const anchor = iframeDocument.getElementById('toggle') as HTMLAnchorElement;
    const syntheticEvent = {
      button: 0,
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

    clickHandler(syntheticEvent);

    expect(onNavigate).not.toHaveBeenCalled();
    expect(syntheticEvent.preventDefault).not.toHaveBeenCalled();
    expect(syntheticEvent.stopPropagation).not.toHaveBeenCalled();
    expect(syntheticEvent.stopImmediatePropagation).not.toHaveBeenCalled();
    expect((syntheticEvent as Event & { cancelBubble?: boolean }).cancelBubble).toBe(false);
    expect((syntheticEvent as Event & { returnValue?: boolean }).returnValue).toBe(true);

    cleanup();
    addDocumentListenerSpy.mockRestore();
  });

  it('leaves missing fragment hash anchors to the iframe page', () => {
    const iframeDocument = document.implementation.createHTMLDocument('iframe-missing-fragment');
    iframeDocument.body.innerHTML = '<a id="missing-link" href="#missing">Missing</a>';
    const addLoadListener = jest.fn();
    const removeLoadListener = jest.fn();
    const addDocumentListenerSpy = jest.spyOn(iframeDocument, 'addEventListener');
    const onNavigate = jest.fn();
    const iframeStub = {
      contentDocument: iframeDocument,
      ownerDocument: document,
      addEventListener: addLoadListener,
      removeEventListener: removeLoadListener,
    } as unknown as HTMLIFrameElement;

    const cleanup = wireInlineIframeNavigation({
      iframe: iframeStub,
      currentPageUrl: validationOptions.currentPageUrl,
      validationOptions,
      cacheBusterParamName: 'v',
      onNavigate,
    });

    const clickRegistration = addDocumentListenerSpy.mock.calls.find(
      (call) => call[0] === 'click',
    );
    const clickHandler = clickRegistration?.[1] as (event: Event) => void;
    const anchor = iframeDocument.getElementById('missing-link') as HTMLAnchorElement;
    const syntheticEvent = {
      button: 0,
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

    clickHandler(syntheticEvent);

    expect(onNavigate).not.toHaveBeenCalled();
    expect(syntheticEvent.preventDefault).not.toHaveBeenCalled();
    expect(syntheticEvent.stopPropagation).not.toHaveBeenCalled();
    expect(syntheticEvent.stopImmediatePropagation).not.toHaveBeenCalled();

    cleanup();
    addDocumentListenerSpy.mockRestore();
  });

  it('navigates from inline bridge postMessage events emitted inside iframe srcdoc', () => {
    const iframeDocument = document.implementation.createHTMLDocument('iframe-message');
    const addLoadListener = jest.fn();
    const removeLoadListener = jest.fn();
    const addWindowListener = jest.fn();
    const removeWindowListener = jest.fn();
    const iframePostMessage = jest.fn();
    const iframeWindowStub = {
      addEventListener: addWindowListener,
      removeEventListener: removeWindowListener,
      postMessage: iframePostMessage,
    } as unknown as Window;
    const onNavigate = jest.fn();
    const iframeStub = {
      contentDocument: iframeDocument,
      contentWindow: iframeWindowStub,
      addEventListener: addLoadListener,
      removeEventListener: removeLoadListener,
    } as unknown as HTMLIFrameElement;

    const cleanup = wireInlineIframeNavigation({
      iframe: iframeStub,
      currentPageUrl: validationOptions.currentPageUrl,
      validationOptions,
      cacheBusterParamName: 'v',
      onNavigate,
    });

    const messageEvent = new MessageEvent('message', {
      data: {
        type: 'uhv-inline-nav',
        targetUrl: 'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/next-message.html?v=1',
      },
    });
    Object.defineProperty(messageEvent, 'source', {
      value: iframeWindowStub,
      configurable: true,
    });

    window.dispatchEvent(messageEvent);

    expect(onNavigate).toHaveBeenCalledWith(
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/next-message.html',
    );

    expect(iframePostMessage).not.toHaveBeenCalled();
    const readyEvent = new MessageEvent('message', {
      data: {
        type: 'uhv-inline-ready',
      },
    });
    Object.defineProperty(readyEvent, 'source', {
      value: iframeWindowStub,
      configurable: true,
    });
    window.dispatchEvent(readyEvent);
    expect(iframePostMessage).toHaveBeenCalledWith(
      {
        type: 'uhv-inline-host-page-url',
        hostPageUrl: window.location.href,
      },
      '*',
    );
    iframePostMessage.mockClear();
    window.dispatchEvent(new Event(INLINE_HOST_PAGE_URL_CHANGED_EVENT));
    expect(iframePostMessage).toHaveBeenCalledWith(
      {
        type: 'uhv-inline-host-page-url',
        hostPageUrl: window.location.href,
      },
      '*',
    );

    iframePostMessage.mockClear();
    const hostSyncLoadHandler = addLoadListener.mock.calls[0][1] as () => void;
    hostSyncLoadHandler();
    window.dispatchEvent(new Event(INLINE_HOST_PAGE_URL_CHANGED_EVENT));
    expect(iframePostMessage).not.toHaveBeenCalled();

    cleanup();
    iframePostMessage.mockClear();
    window.dispatchEvent(new Event(INLINE_HOST_PAGE_URL_CHANGED_EVENT));
    expect(iframePostMessage).not.toHaveBeenCalled();
  });
});
