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

    expect(result).toContain('data-uhv-inline-csp="1"');
    expect(result).toContain('object-src &#39;none&#39;');
    expect(result).toContain('data-uhv-history-compat="1"');
    expect(result).toContain(
      '<base href="https://contoso.sharepoint.com/sites/TestSite2/SiteAssets/GPOzaurr/">',
    );
    expect(result.indexOf('data-uhv-inline-csp="1"')).toBeLessThan(
      result.indexOf('data-uhv-history-compat="1"'),
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

  it('does not add duplicate srcdoc CSP when one already exists', () => {
    const inputHtml =
      '<html><head><meta http-equiv="Content-Security-Policy" content="default-src \'self\'"><title>Report</title></head><body></body></html>';
    const result = prepareInlineHtmlForSrcDoc(
      inputHtml,
      '/sites/TestSite2/SiteAssets/GPOzaurr/',
      'https://contoso.sharepoint.com/sites/TestSite2/SitePages/Dashboard.aspx',
    );

    const cspMatches = result.match(/http-equiv="Content-Security-Policy"/gi) || [];
    expect(cspMatches.length).toBe(1);
    expect(result).not.toContain('data-uhv-inline-csp="1"');
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
        undefined,
        {
          maxRetryAttempts: 1,
        },
      ),
    ).rejects.toThrow('SharePoint API returned 503 Service Unavailable');

    const recovered = await loadSharePointFileContentForInline(
      mockSpHttpClient as never,
      webAbsoluteUrl,
      sourceUrl,
      baseUrlForRelativeLinks,
      pageUrl,
      undefined,
      {
        maxRetryAttempts: 1,
      },
    );

    expect(mockGet).toHaveBeenCalledTimes(2);
    expect(recovered).toContain('<h1>Recovered</h1>');
  });

  it('retries throttled responses and succeeds on a subsequent attempt', async () => {
    const throttledResponse = {
      ok: false,
      status: 503,
      statusText: 'Service Unavailable',
      headers: {
        get: jest.fn().mockReturnValue(undefined),
      },
      text: jest.fn().mockResolvedValue(''),
    };
    const successResponse = {
      ok: true,
      status: 200,
      statusText: 'OK',
      text: jest
        .fn()
        .mockResolvedValue('<html><head></head><body><h1>Retried</h1></body></html>'),
    };
    const mockGet = jest
      .fn()
      .mockResolvedValueOnce(throttledResponse)
      .mockResolvedValueOnce(successResponse);
    const mockSpHttpClient = {
      get: mockGet,
    };

    const result = await loadSharePointFileContentForInline(
      mockSpHttpClient as never,
      webAbsoluteUrl,
      sourceUrl,
      baseUrlForRelativeLinks,
      pageUrl,
      undefined,
      {
        maxRetryAttempts: 3,
        retryBaseDelayMs: 0,
        retryMaxDelayMs: 0,
      },
    );

    expect(mockGet).toHaveBeenCalledTimes(2);
    expect(result).toContain('<h1>Retried</h1>');
  });

  it('does not retry non-throttling errors', async () => {
    const notFoundResponse = {
      ok: false,
      status: 404,
      statusText: 'Not Found',
      headers: {
        get: jest.fn().mockReturnValue(undefined),
      },
      text: jest.fn().mockResolvedValue(''),
    };
    const mockGet = jest.fn().mockResolvedValue(notFoundResponse);
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
        undefined,
        {
          maxRetryAttempts: 4,
          retryBaseDelayMs: 0,
          retryMaxDelayMs: 0,
        },
      ),
    ).rejects.toThrow('SharePoint API returned 404 Not Found');

    expect(mockGet).toHaveBeenCalledTimes(1);
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

  it('retries transient thrown network errors before succeeding', async () => {
    const networkError = Object.assign(new Error('failed to fetch'), {
      name: 'TypeError',
    });
    const successResponse = {
      ok: true,
      status: 200,
      statusText: 'OK',
      text: jest
        .fn()
        .mockResolvedValue('<html><head></head><body><h1>RecoveredNetwork</h1></body></html>'),
    };
    const mockGet = jest
      .fn()
      .mockRejectedValueOnce(networkError)
      .mockResolvedValueOnce(successResponse);
    const mockSpHttpClient = {
      get: mockGet,
    };

    const result = await loadSharePointFileContentForInline(
      mockSpHttpClient as never,
      webAbsoluteUrl,
      sourceUrl,
      baseUrlForRelativeLinks,
      pageUrl,
      undefined,
      {
        maxRetryAttempts: 3,
        retryBaseDelayMs: 0,
        retryMaxDelayMs: 0,
      },
    );

    expect(mockGet).toHaveBeenCalledTimes(2);
    expect(result).toContain('<h1>RecoveredNetwork</h1>');
  });

  it('does not retry abort errors thrown by fetch', async () => {
    const abortError = Object.assign(new Error('request aborted'), {
      name: 'AbortError',
    });
    const mockGet = jest.fn().mockRejectedValue(abortError);
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
        undefined,
        {
          maxRetryAttempts: 3,
          retryBaseDelayMs: 0,
          retryMaxDelayMs: 0,
        },
      ),
    ).rejects.toThrow('request aborted');

    expect(mockGet).toHaveBeenCalledTimes(1);
  });

  it('reuses cache across page query/hash variants', async () => {
    const mockResponse = {
      ok: true,
      status: 200,
      statusText: 'OK',
      text: jest
        .fn()
        .mockResolvedValue('<html><head></head><body><h1>CacheNormalized</h1></body></html>'),
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
      'https://contoso.sharepoint.com/sites/TestSite2/SitePages/Dashboard.aspx?view=ops',
    );
    await loadSharePointFileContentForInline(
      mockSpHttpClient as never,
      webAbsoluteUrl,
      sourceUrl,
      baseUrlForRelativeLinks,
      'https://contoso.sharepoint.com/sites/TestSite2/SitePages/Dashboard.aspx?view=sales#section',
    );

    expect(mockGet).toHaveBeenCalledTimes(1);
  });
});

