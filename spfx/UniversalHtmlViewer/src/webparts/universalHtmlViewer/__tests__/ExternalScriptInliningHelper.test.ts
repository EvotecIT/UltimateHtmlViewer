import {
  clearExternalScriptInliningCacheForTests,
  getInlineExternalScriptAllowedHosts,
  inlineAllowedExternalScripts,
} from '../ExternalScriptInliningHelper';

describe('ExternalScriptInliningHelper', () => {
  const pageUrl = 'https://contoso.sharepoint.com/sites/TestSite2/SitePages/Dashboard.aspx';
  const baseUrlForRelativeLinks = '/sites/TestSite2/SiteAssets/Reports/';
  const originalFetch = globalThis.fetch;

  beforeEach(() => {
    clearExternalScriptInliningCacheForTests();
  });

  afterEach(() => {
    Object.defineProperty(globalThis, 'fetch', {
      configurable: true,
      value: originalFetch,
    });
  });

  it('inlines allowed external report scripts', async () => {
    const mockFetch = jest.fn().mockResolvedValue({
      ok: true,
      text: jest.fn().mockResolvedValue(
        'window.jQuery = function(){}; window.$ = window.jQuery;',
      ),
    });
    Object.defineProperty(globalThis, 'fetch', {
      configurable: true,
      value: mockFetch,
    });

    const result = await inlineAllowedExternalScripts(
      '<html><head><script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script></head><body></body></html>',
      baseUrlForRelativeLinks,
      pageUrl,
      {
        enabled: true,
      },
    );

    expect(mockFetch).toHaveBeenCalledWith(
      'https://cdnjs.cloudflare.com/ajax/libs/jquery/3.5.1/jquery.min.js',
      { credentials: 'omit', mode: 'cors' },
    );
    expect(result).toContain('data-uhv-inlined-external-script');
    expect(result).toContain('window.$ = window.jQuery;');
    expect(result).not.toContain('src="https://cdnjs.cloudflare.com/ajax/libs/jquery');
  });

  it('includes credentials for same-origin SharePoint script assets', async () => {
    const mockFetch = jest.fn().mockResolvedValue({
      ok: true,
      text: jest.fn().mockResolvedValue('window.localReportScriptLoaded = true;'),
    });
    Object.defineProperty(globalThis, 'fetch', {
      configurable: true,
      value: mockFetch,
    });

    await inlineAllowedExternalScripts(
      '<html><head><script src="/sites/TestSite2/SiteAssets/report-support.js"></script></head><body></body></html>',
      baseUrlForRelativeLinks,
      pageUrl,
      {
        enabled: true,
      },
    );

    expect(mockFetch).toHaveBeenCalledWith(
      'https://contoso.sharepoint.com/sites/TestSite2/SiteAssets/report-support.js',
      { credentials: 'same-origin', mode: 'cors' },
    );
  });

  it('does not inline external scripts from non-allowlisted hosts', async () => {
    const mockFetch = jest.fn();
    Object.defineProperty(globalThis, 'fetch', {
      configurable: true,
      value: mockFetch,
    });

    const result = await inlineAllowedExternalScripts(
      '<html><head><script src="https://evil.example.com/report.js"></script></head><body></body></html>',
      baseUrlForRelativeLinks,
      pageUrl,
      {
        enabled: true,
        allowedHosts: ['cdn.jsdelivr.net'],
      },
    );

    expect(mockFetch).not.toHaveBeenCalled();
    expect(result).toContain('src="https://evil.example.com/report.js"');
  });

  it('leaves report HTML unchanged when the option is disabled', async () => {
    const html =
      '<html><head><script src="https://cdnjs.cloudflare.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script></head><body></body></html>';
    const result = await inlineAllowedExternalScripts(
      html,
      baseUrlForRelativeLinks,
      pageUrl,
      {
        enabled: false,
      },
    );

    expect(result).toBe(html);
  });

  it('uses the stable DataTables CDN before the old nightly URL', async () => {
    const mockFetch = jest.fn().mockResolvedValueOnce({
      ok: true,
      text: jest.fn().mockResolvedValue('jQuery.fn.dataTable = {};'),
    });
    Object.defineProperty(globalThis, 'fetch', {
      configurable: true,
      value: mockFetch,
    });

    const result = await inlineAllowedExternalScripts(
      '<html><head><script src="https://nightly.datatables.net/js/jquery.dataTables.min.js"></script></head><body></body></html>',
      baseUrlForRelativeLinks,
      pageUrl,
      {
        enabled: true,
      },
    );

    expect(mockFetch).toHaveBeenNthCalledWith(
      1,
      'https://cdn.datatables.net/1.10.24/js/jquery.dataTables.min.js',
      { credentials: 'omit', mode: 'cors' },
    );
    expect(mockFetch).toHaveBeenCalledTimes(1);
    expect(result).toContain(
      'data-uhv-inlined-external-script="https://cdn.datatables.net/1.10.24/js/jquery.dataTables.min.js"',
    );
    expect(result).toContain('jQuery.fn.dataTable = {};');
    expect(result).not.toContain('src="https://nightly.datatables.net');
  });

  it('defaults to common PSWriteHTML CDN hosts', () => {
    expect(getInlineExternalScriptAllowedHosts()).toEqual([
      'code.jquery.com',
      'cdnjs.cloudflare.com',
      'cdn.datatables.net',
      'cdn.jsdelivr.net',
      'nightly.datatables.net',
      'unpkg.com',
    ]);
  });
});
