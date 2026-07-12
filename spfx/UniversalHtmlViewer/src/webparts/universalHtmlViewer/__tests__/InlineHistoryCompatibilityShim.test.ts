import { prepareInlineHtmlForSrcDoc } from '../InlineHtmlTransformHelper';

describe('inline history compatibility shim', () => {
  it('swallows unsupported srcdoc hash history updates without navigating to the SharePoint base file', () => {
    const result = prepareInlineHtmlForSrcDoc(
      '<html><head></head><body></body></html>',
      'https://contoso.sharepoint.com/sites/TestSite2/SiteAssets/GPOzaurr/GPOBroken.html',
      'https://contoso.sharepoint.com/sites/TestSite2/SitePages/Dashboard.aspx',
    );
    const parsed = new DOMParser().parseFromString(result, 'text/html');
    const shim = parsed.querySelector('script[data-uhv-history-compat="1"]')?.textContent || '';
    const replace = jest.fn();
    const historyObject = {
      pushState: jest.fn((_state?: unknown, _title?: string, _url?: string) => {
        const error = new Error('srcdoc state URL is not permitted');
        error.name = 'SecurityError';
        throw error;
      }),
      replaceState: jest.fn(
        (_state?: unknown, _title?: string, _url?: string) => undefined,
      ),
    };
    const frameWindow = {
      history: historyObject,
      location: {
        href: 'about:srcdoc',
        replace,
      },
    };

    // eslint-disable-next-line no-new-func
    new Function('window', shim)(frameWindow);
    expect(() => historyObject.pushState(null, '', '#WizardStep-s9dl1j34')).not.toThrow();
    expect(replace).not.toHaveBeenCalled();
  });
});
