import {
  clearInlineHtmlCacheForTests,
  loadSharePointFileContentForInline,
} from '../SharePointInlineContentHelper';

describe('loadSharePointFileContentForInline Retry-After handling', () => {
  const webAbsoluteUrl = 'https://contoso.sharepoint.com/sites/TestSite2';
  const sourceUrl = '/sites/TestSite2/SiteAssets/Reports/index.html';
  const baseUrlForRelativeLinks = '/sites/TestSite2/SiteAssets/Reports/';
  const pageUrl = 'https://contoso.sharepoint.com/sites/TestSite2/SitePages/Dashboard.aspx';

  beforeEach(() => {
    clearInlineHtmlCacheForTests();
  });

  it('caps Retry-After delays using retryMaxDelayMs', async () => {
    const throttledResponse = {
      ok: false,
      status: 503,
      statusText: 'Service Unavailable',
      headers: {
        get: jest.fn().mockImplementation((name: string) =>
          name.toLowerCase() === 'retry-after' ? '120' : undefined,
        ),
      },
      text: jest.fn().mockResolvedValue(''),
    };
    const successResponse = {
      ok: true,
      status: 200,
      statusText: 'OK',
      text: jest
        .fn()
        .mockResolvedValue('<html><head></head><body><h1>RetriedCapped</h1></body></html>'),
    };
    const mockGet = jest
      .fn()
      .mockResolvedValueOnce(throttledResponse)
      .mockResolvedValueOnce(successResponse);
    const mockSpHttpClient = {
      get: mockGet,
    };
    const setTimeoutSpy = jest.spyOn(global, 'setTimeout');

    try {
      const result = await loadSharePointFileContentForInline(
        mockSpHttpClient as never,
        webAbsoluteUrl,
        sourceUrl,
        baseUrlForRelativeLinks,
        pageUrl,
        undefined,
        {
          maxRetryAttempts: 2,
          retryBaseDelayMs: 750,
          retryMaxDelayMs: 1,
        },
      );

      expect(mockGet).toHaveBeenCalledTimes(2);
      expect(result).toContain('<h1>RetriedCapped</h1>');

      const retryTimeouts = setTimeoutSpy.mock.calls
        .map((call) => call[1])
        .filter((value): value is number => typeof value === 'number');
      expect(retryTimeouts).toContain(1);
    } finally {
      setTimeoutSpy.mockRestore();
    }
  });
});
