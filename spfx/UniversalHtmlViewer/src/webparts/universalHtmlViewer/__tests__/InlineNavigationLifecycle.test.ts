import { wireInlineIframeNavigation } from '../InlineNavigationHelper';
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
});
