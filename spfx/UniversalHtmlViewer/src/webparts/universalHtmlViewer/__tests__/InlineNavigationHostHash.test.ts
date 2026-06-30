import { wireInlineIframeNavigation } from '../InlineNavigationHelper';
import { scrollHostPageToIframeHashTarget } from '../HostPageHashNavigationHelper';
import { UrlValidationOptions } from '../UrlHelper';

describe('wireInlineIframeNavigation host page hash', () => {
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
    jest.useRealTimers();
    jest.restoreAllMocks();
  });

  it('applies a host page hash to the matching iframe target on load', () => {
    window.history.replaceState(null, document.title, '#what');

    const iframeDocument = document.implementation.createHTMLDocument('iframe-host-hash');
    iframeDocument.body.innerHTML = '<section id="what">Details</section>';
    const addLoadListener = jest.fn();
    const removeLoadListener = jest.fn();
    const onNavigate = jest.fn();
    const iframeStub = createIframeStub(iframeDocument, 240);

    iframeStub.addEventListener = addLoadListener;
    iframeStub.removeEventListener = removeLoadListener;

    const targetSection = iframeDocument.getElementById('what') as HTMLElement;
    targetSection.scrollIntoView = jest.fn();
    targetSection.getBoundingClientRect = () => createRect(900, 100);
    const scrollToSpy = jest.spyOn(window, 'scrollTo').mockImplementation(() => {
      return;
    });

    const cleanup = wireInlineIframeNavigation({
      iframe: iframeStub,
      currentPageUrl: validationOptions.currentPageUrl,
      validationOptions,
      cacheBusterParamName: 'v',
      onNavigate,
    });

    expect(onNavigate).not.toHaveBeenCalled();
    expect(targetSection.scrollIntoView).toHaveBeenCalledTimes(1);
    expect(scrollToSpy).toHaveBeenCalledWith({
      top: 1044,
      left: 0,
      behavior: 'auto',
    });
    expect(iframeDocument.documentElement?.getAttribute('data-uhv-host-hash')).toBe('#what');

    cleanup();
  });

  it('ignores a host page hash that does not exist inside the iframe document', () => {
    window.history.replaceState(null, document.title, '#missing');

    const iframeDocument = document.implementation.createHTMLDocument('iframe-missing-hash');
    iframeDocument.body.innerHTML = '<section id="what">Details</section>';
    const onNavigate = jest.fn();
    const iframeStub = createIframeStub(iframeDocument, 240);

    const targetSection = iframeDocument.getElementById('what') as HTMLElement;
    targetSection.scrollIntoView = jest.fn();
    const scrollToSpy = jest.spyOn(window, 'scrollTo').mockImplementation(() => {
      return;
    });

    const cleanup = wireInlineIframeNavigation({
      iframe: iframeStub,
      currentPageUrl: validationOptions.currentPageUrl,
      validationOptions,
      cacheBusterParamName: 'v',
      onNavigate,
    });

    expect(onNavigate).not.toHaveBeenCalled();
    expect(targetSection.scrollIntoView).not.toHaveBeenCalled();
    expect(scrollToSpy).not.toHaveBeenCalled();
    expect(iframeDocument.documentElement?.getAttribute('data-uhv-host-hash')).toBeNull();

    cleanup();
  });

  it('resets handled hash marker when a later host hash is not found', () => {
    window.history.replaceState(null, document.title, '#security');

    const iframeDocument = document.implementation.createHTMLDocument('iframe-marker-reset');
    iframeDocument.body.innerHTML = '<section id="security">Security</section>';
    const onNavigate = jest.fn();
    const iframeStub = createIframeStub(iframeDocument, 180);

    const targetSection = iframeDocument.getElementById('security') as HTMLElement;
    targetSection.scrollIntoView = jest.fn();
    targetSection.getBoundingClientRect = () => createRect(1200, 100);
    jest.spyOn(window, 'scrollTo').mockImplementation(() => {
      return;
    });

    const cleanup = wireInlineIframeNavigation({
      iframe: iframeStub,
      currentPageUrl: validationOptions.currentPageUrl,
      validationOptions,
      cacheBusterParamName: 'v',
      onNavigate,
    });

    expect(iframeDocument.documentElement?.getAttribute('data-uhv-host-hash')).toBe('#security');

    window.history.replaceState(null, document.title, '#missing');
    window.dispatchEvent(new Event('hashchange'));

    expect(iframeDocument.documentElement?.getAttribute('data-uhv-host-hash')).toBeNull();

    window.history.replaceState(null, document.title, '#security');
    window.dispatchEvent(new Event('hashchange'));

    expect(targetSection.scrollIntoView).toHaveBeenCalledTimes(2);
    expect(iframeDocument.documentElement?.getAttribute('data-uhv-host-hash')).toBe('#security');

    cleanup();
  });

  it('applies host page hash changes to matching iframe targets', () => {
    const iframeDocument = document.implementation.createHTMLDocument('iframe-hash-change');
    iframeDocument.body.innerHTML = '<section id="security">Security</section>';
    const onNavigate = jest.fn();
    const iframeStub = createIframeStub(iframeDocument, 180);

    const targetSection = iframeDocument.getElementById('security') as HTMLElement;
    targetSection.scrollIntoView = jest.fn();
    targetSection.getBoundingClientRect = () => createRect(1200, 100);
    const scrollToSpy = jest.spyOn(window, 'scrollTo').mockImplementation(() => {
      return;
    });

    const cleanup = wireInlineIframeNavigation({
      iframe: iframeStub,
      currentPageUrl: validationOptions.currentPageUrl,
      validationOptions,
      cacheBusterParamName: 'v',
      onNavigate,
    });

    window.history.replaceState(null, document.title, '#security');
    window.dispatchEvent(new Event('hashchange'));

    expect(onNavigate).not.toHaveBeenCalled();
    expect(targetSection.scrollIntoView).toHaveBeenCalledTimes(1);
    expect(scrollToSpy).toHaveBeenCalledWith({
      top: 1284,
      left: 0,
      behavior: 'auto',
    });
    expect(iframeDocument.documentElement?.getAttribute('data-uhv-host-hash')).toBe('#security');

    cleanup();
  });

  it('cancels stale host hash retry scrolls when the hash changes again', () => {
    jest.useFakeTimers();

    const iframeDocument = document.implementation.createHTMLDocument('iframe-stale-retry');
    iframeDocument.body.innerHTML =
      '<section id="security">Security</section><section id="overview">Overview</section>';
    const onNavigate = jest.fn();
    const iframeStub = createIframeStub(iframeDocument, 180);

    const securitySection = iframeDocument.getElementById('security') as HTMLElement;
    const overviewSection = iframeDocument.getElementById('overview') as HTMLElement;
    securitySection.scrollIntoView = jest.fn();
    overviewSection.scrollIntoView = jest.fn();
    securitySection.getBoundingClientRect = () => createRect(1200, 100);
    overviewSection.getBoundingClientRect = () => createRect(1600, 100);
    jest.spyOn(window, 'scrollTo').mockImplementation(() => {
      return;
    });

    const cleanup = wireInlineIframeNavigation({
      iframe: iframeStub,
      currentPageUrl: validationOptions.currentPageUrl,
      validationOptions,
      cacheBusterParamName: 'v',
      onNavigate,
    });

    window.history.replaceState(null, document.title, '#security');
    window.dispatchEvent(new Event('hashchange'));
    window.history.replaceState(null, document.title, '#overview');
    window.dispatchEvent(new Event('hashchange'));
    jest.runOnlyPendingTimers();

    expect(securitySection.scrollIntoView).toHaveBeenCalledTimes(1);
    expect(overviewSection.scrollIntoView).toHaveBeenCalled();

    cleanup();
  });

  it('scrolls the nearest host scroll container when SharePoint does not scroll window', () => {
    const iframeDocument = document.implementation.createHTMLDocument('iframe-scroll-container');
    iframeDocument.body.innerHTML = '<section id="security">Security</section>';
    const targetSection = iframeDocument.getElementById('security') as HTMLElement;
    targetSection.getBoundingClientRect = () => createRect(900, 100);

    const scrollContainer = document.createElement('div');
    scrollContainer.style.overflowY = 'scroll';
    Object.defineProperty(scrollContainer, 'clientHeight', { value: 500 });
    Object.defineProperty(scrollContainer, 'scrollHeight', { value: 1500 });
    scrollContainer.getBoundingClientRect = () => createRect(80, 500);
    const containerScrollTo = jest.fn(
      (optionsOrX?: ScrollToOptions | number, y?: number) => {
        scrollContainer.scrollTop =
          typeof optionsOrX === 'number' ? y || 0 : optionsOrX?.top || 0;
      },
    );
    scrollContainer.scrollTo = containerScrollTo as typeof scrollContainer.scrollTo;

    const iframe = document.createElement('iframe') as HTMLIFrameElement;
    Object.defineProperty(iframe, 'contentDocument', { value: iframeDocument });
    iframe.getBoundingClientRect = () => createRect(260, 800);
    scrollContainer.appendChild(iframe);
    document.body.appendChild(scrollContainer);

    const handled = scrollHostPageToIframeHashTarget(iframe, iframeDocument, '#security');

    expect(handled).toBe(true);
    expect(containerScrollTo).toHaveBeenCalledWith({
      top: 984,
      left: 0,
      behavior: 'auto',
    });
    expect(scrollContainer.scrollTop).toBe(984);

    scrollContainer.remove();
  });
});

function createIframeStub(iframeDocument: Document, iframeTop: number): HTMLIFrameElement {
  return {
    contentDocument: iframeDocument,
    ownerDocument: document,
    addEventListener: jest.fn(),
    removeEventListener: jest.fn(),
    getBoundingClientRect: () => createRect(iframeTop, 800),
  } as unknown as HTMLIFrameElement;
}

function createRect(top: number, height: number): DOMRect {
  return {
    top,
    left: 0,
    bottom: top + height,
    right: 800,
    width: 800,
    height,
  } as DOMRect;
}
