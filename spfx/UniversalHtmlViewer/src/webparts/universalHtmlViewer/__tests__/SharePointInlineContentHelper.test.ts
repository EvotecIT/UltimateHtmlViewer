import { prepareInlineHtmlForSrcDoc } from '../InlineHtmlTransformHelper';
import {
  clearInlineHtmlCacheForTests,
  loadSharePointFileContentForInline,
} from '../SharePointInlineContentHelper';

describe('prepareInlineHtmlForSrcDoc', () => {
  it('injects base href and history compatibility shim when head exists', () => {
    const inputHtml = '<html><head><title>Report</title></head><body><h1>Dashboard</h1></body></html>';
    const result = prepareInlineHtmlForSrcDoc(
      inputHtml,
      '/sites/TestSite2/SiteAssets/GPOzaurr/',
      'https://contoso.sharepoint.com/sites/TestSite2/SitePages/Dashboard.aspx',
    );

    expect(result).toContain('data-uhv-history-compat="1"');
    expect(result).toContain(
      '<base href="https://contoso.sharepoint.com/sites/TestSite2/SiteAssets/GPOzaurr/">',
    );
    expect(result.indexOf('data-uhv-history-compat="1"')).toBeLessThan(
      result.indexOf('<title>Report</title>'),
    );
  });

  it('does not add duplicate base tag when one already exists', () => {
    const inputHtml =
      '<html><head><base href="https://contoso.sharepoint.com/sites/TestSite2/SiteAssets/GPOzaurr/"><title>Report</title></head><body></body></html>';
    const result = prepareInlineHtmlForSrcDoc(
      inputHtml,
      '/sites/TestSite2/SiteAssets/GPOzaurr/',
      'https://contoso.sharepoint.com/sites/TestSite2/SitePages/Dashboard.aspx',
    );

    const baseMatches = result.match(/<base\s+/gi) || [];
    expect(baseMatches.length).toBe(1);
    expect(result).toContain('data-uhv-history-compat="1"');
  });

  it('does not add duplicate history shim when already present', () => {
    const inputHtml =
      '<html><head><script data-uhv-history-compat="1">(function(){return;})();</script></head><body></body></html>';
    const result = prepareInlineHtmlForSrcDoc(
      inputHtml,
      '/sites/TestSite2/SiteAssets/GPOzaurr/',
      'https://contoso.sharepoint.com/sites/TestSite2/SitePages/Dashboard.aspx',
    );

    const shimMatches = result.match(/data-uhv-history-compat="1"/gi) || [];
    expect(shimMatches.length).toBe(1);
  });

  it('neutralizes nested html iframe src to avoid direct browser download navigation', () => {
    const inputHtml =
      '<!DOCTYPE html><html><head><title>Report</title></head><body><iframe src="GPOzaurr/GPOBlockedInheritance_2021-04-15_184719.html"></iframe></body></html>';
    const result = prepareInlineHtmlForSrcDoc(
      inputHtml,
      '/sites/TestSite2/SiteAssets/',
      'https://contoso.sharepoint.com/sites/TestSite2/SitePages/Dashboard.aspx',
    );

    expect(result).toContain('data-uhv-inline-src="GPOzaurr/GPOBlockedInheritance_2021-04-15_184719.html"');
    expect(result).toContain('src="about:blank"');
  });
});

describe('loadSharePointFileContentForInline', () => {
  const webAbsoluteUrl = 'https://contoso.sharepoint.com/sites/TestSite2';
  const sourceUrl = '/sites/TestSite2/SiteAssets/Reports/index.html';
  const baseUrlForRelativeLinks = '/sites/TestSite2/SiteAssets/Reports/';
  const pageUrl = 'https://contoso.sharepoint.com/sites/TestSite2/SitePages/Dashboard.aspx';

  beforeEach(() => {
    clearInlineHtmlCacheForTests();
  });

  it('reuses cached HTML for repeated calls with identical inputs', async () => {
    const mockResponse = {
      ok: true,
      status: 200,
      statusText: 'OK',
      text: jest
        .fn()
        .mockResolvedValue('<html><head></head><body><h1>Cached</h1></body></html>'),
    };
    const mockGet = jest.fn().mockResolvedValue(mockResponse);
    const mockSpHttpClient = {
      get: mockGet,
    };

    const firstResult = await loadSharePointFileContentForInline(
      mockSpHttpClient as never,
      webAbsoluteUrl,
      sourceUrl,
      baseUrlForRelativeLinks,
      pageUrl,
    );
    const secondResult = await loadSharePointFileContentForInline(
      mockSpHttpClient as never,
      webAbsoluteUrl,
      sourceUrl,
      baseUrlForRelativeLinks,
      pageUrl,
    );

    expect(mockGet).toHaveBeenCalledTimes(1);
    expect(firstResult).toContain('<h1>Cached</h1>');
    expect(secondResult).toContain('<h1>Cached</h1>');
  });

  it('deduplicates concurrent in-flight requests for the same key', async () => {
    let resolveResponse: ((value: unknown) => void) | undefined;
    const getPromise = new Promise((resolve) => {
      resolveResponse = resolve;
    });
    const mockGet = jest.fn().mockReturnValue(getPromise);
    const mockSpHttpClient = {
      get: mockGet,
    };

    const pendingResultA = loadSharePointFileContentForInline(
      mockSpHttpClient as never,
      webAbsoluteUrl,
      sourceUrl,
      baseUrlForRelativeLinks,
      pageUrl,
    );
    const pendingResultB = loadSharePointFileContentForInline(
      mockSpHttpClient as never,
      webAbsoluteUrl,
      sourceUrl,
      baseUrlForRelativeLinks,
      pageUrl,
    );

    expect(mockGet).toHaveBeenCalledTimes(1);

    const mockResponse = {
      ok: true,
      status: 200,
      statusText: 'OK',
      text: jest
        .fn()
        .mockResolvedValue('<html><head></head><body><h1>InFlight</h1></body></html>'),
    };
    if (!resolveResponse) {
      throw new Error('Expected response resolver to be initialized.');
    }
    resolveResponse(mockResponse);

    const [resultA, resultB] = await Promise.all([pendingResultA, pendingResultB]);
    expect(resultA).toContain('<h1>InFlight</h1>');
    expect(resultB).toContain('<h1>InFlight</h1>');
  });

  it('does not cache failed fetch responses', async () => {
    const errorResponse = {
      ok: false,
      status: 503,
      statusText: 'Service Unavailable',
      text: jest.fn().mockResolvedValue(''),
    };
    const successResponse = {
      ok: true,
      status: 200,
      statusText: 'OK',
      text: jest
        .fn()
        .mockResolvedValue('<html><head></head><body><h1>Recovered</h1></body></html>'),
    };
    const mockGet = jest
      .fn()
      .mockResolvedValueOnce(errorResponse)
      .mockResolvedValueOnce(successResponse);
    const mockSpHttpClient = {
      get: mockGet,
    };

    await expect(
      loadSharePointFileContentForInline(
        mockSpHttpClient as never,
        webAbsoluteUrl,
        sourceUrl,
        baseUrlForRelativeLinks,
        pageUrl,
      ),
    ).rejects.toThrow('SharePoint API returned 503 Service Unavailable');

    const recovered = await loadSharePointFileContentForInline(
      mockSpHttpClient as never,
      webAbsoluteUrl,
      sourceUrl,
      baseUrlForRelativeLinks,
      pageUrl,
    );

    expect(mockGet).toHaveBeenCalledTimes(2);
    expect(recovered).toContain('<h1>Recovered</h1>');
  });

  it('treats cache-busted source URLs as distinct cache keys', async () => {
    const mockResponse = {
      ok: true,
      status: 200,
      statusText: 'OK',
      text: jest
        .fn()
        .mockResolvedValue('<html><head></head><body><h1>Versioned</h1></body></html>'),
    };
    const mockGet = jest.fn().mockResolvedValue(mockResponse);
    const mockSpHttpClient = {
      get: mockGet,
    };

    await loadSharePointFileContentForInline(
      mockSpHttpClient as never,
      webAbsoluteUrl,
      `${sourceUrl}?v=1`,
      baseUrlForRelativeLinks,
      pageUrl,
    );
    await loadSharePointFileContentForInline(
      mockSpHttpClient as never,
      webAbsoluteUrl,
      `${sourceUrl}?v=2`,
      baseUrlForRelativeLinks,
      pageUrl,
    );

    expect(mockGet).toHaveBeenCalledTimes(2);
  });

  it('bypasses cache when explicitly requested', async () => {
    const mockResponse = {
      ok: true,
      status: 200,
      statusText: 'OK',
      text: jest
        .fn()
        .mockResolvedValue('<html><head></head><body><h1>Bypass</h1></body></html>'),
    };
    const mockGet = jest.fn().mockResolvedValue(mockResponse);
    const mockSpHttpClient = {
      get: mockGet,
    };

    await loadSharePointFileContentForInline(
      mockSpHttpClient as never,
      webAbsoluteUrl,
      sourceUrl,
      baseUrlForRelativeLinks,
      pageUrl,
      undefined,
      {
        bypassCache: true,
      },
    );
    await loadSharePointFileContentForInline(
      mockSpHttpClient as never,
      webAbsoluteUrl,
      sourceUrl,
      baseUrlForRelativeLinks,
      pageUrl,
      undefined,
      {
        bypassCache: true,
      },
    );

    expect(mockGet).toHaveBeenCalledTimes(2);
  });

  it('disables response cache when TTL is set to zero', async () => {
    const mockResponse = {
      ok: true,
      status: 200,
      statusText: 'OK',
      text: jest
        .fn()
        .mockResolvedValue('<html><head></head><body><h1>NoCache</h1></body></html>'),
    };
    const mockGet = jest.fn().mockResolvedValue(mockResponse);
    const mockSpHttpClient = {
      get: mockGet,
    };

    await loadSharePointFileContentForInline(
      mockSpHttpClient as never,
      webAbsoluteUrl,
      sourceUrl,
      baseUrlForRelativeLinks,
      pageUrl,
      undefined,
      {
        cacheTtlMs: 0,
      },
    );
    await loadSharePointFileContentForInline(
      mockSpHttpClient as never,
      webAbsoluteUrl,
      sourceUrl,
      baseUrlForRelativeLinks,
      pageUrl,
      undefined,
      {
        cacheTtlMs: 0,
      },
    );

    expect(mockGet).toHaveBeenCalledTimes(2);
  });
});

