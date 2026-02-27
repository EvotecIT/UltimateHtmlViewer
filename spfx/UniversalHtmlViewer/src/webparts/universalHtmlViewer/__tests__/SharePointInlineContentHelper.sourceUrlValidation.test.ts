import {
  clearInlineHtmlCacheForTests,
  loadSharePointFileContentForInline,
} from '../SharePointInlineContentHelper';

describe('loadSharePointFileContentForInline source URL validation', () => {
  const webAbsoluteUrl = 'https://contoso.sharepoint.com/sites/TestSite2';
  const baseUrlForRelativeLinks = '/sites/TestSite2/SiteAssets/Reports/';
  const pageUrl = 'https://contoso.sharepoint.com/sites/TestSite2/SitePages/Dashboard.aspx';

  beforeEach(() => {
    clearInlineHtmlCacheForTests();
  });

  it('rejects protocol-relative source URLs', async () => {
    const mockGet = jest.fn();
    const mockSpHttpClient = {
      get: mockGet,
    };

    await expect(
      loadSharePointFileContentForInline(
        mockSpHttpClient as never,
        webAbsoluteUrl,
        '//contoso.sharepoint.com/sites/TestSite2/SiteAssets/Reports/index.html',
        baseUrlForRelativeLinks,
        pageUrl,
      ),
    ).rejects.toThrow('SharePoint file API mode requires a same-tenant URL or a site-relative URL.');

    expect(mockGet).not.toHaveBeenCalled();
  });

  it('rejects non-http absolute source URLs even when host matches', async () => {
    const mockGet = jest.fn();
    const mockSpHttpClient = {
      get: mockGet,
    };

    await expect(
      loadSharePointFileContentForInline(
        mockSpHttpClient as never,
        webAbsoluteUrl,
        'ftp://contoso.sharepoint.com/sites/TestSite2/SiteAssets/Reports/index.html',
        baseUrlForRelativeLinks,
        pageUrl,
      ),
    ).rejects.toThrow('SharePoint file API mode requires a same-tenant URL or a site-relative URL.');

    expect(mockGet).not.toHaveBeenCalled();
  });

  it('accepts same-host https absolute source URLs', async () => {
    const mockResponse = {
      ok: true,
      status: 200,
      statusText: 'OK',
      text: jest
        .fn()
        .mockResolvedValue('<html><head></head><body><h1>Absolute</h1></body></html>'),
    };
    const mockGet = jest.fn().mockResolvedValue(mockResponse);
    const mockSpHttpClient = {
      get: mockGet,
    };

    const result = await loadSharePointFileContentForInline(
      mockSpHttpClient as never,
      webAbsoluteUrl,
      'https://contoso.sharepoint.com/sites/TestSite2/SiteAssets/Reports/index.html?cache=1#anchor',
      baseUrlForRelativeLinks,
      pageUrl,
      undefined,
      {
        maxRetryAttempts: 1,
      },
    );

    expect(mockGet).toHaveBeenCalledTimes(1);
    expect(result).toContain('<h1>Absolute</h1>');
  });
});
