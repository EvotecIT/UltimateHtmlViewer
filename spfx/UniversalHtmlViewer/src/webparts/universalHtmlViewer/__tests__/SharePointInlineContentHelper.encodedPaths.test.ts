import {
  clearInlineHtmlCacheForTests,
  loadSharePointFileContentForInline,
} from '../SharePointInlineContentHelper';

describe('loadSharePointFileContentForInline encoded server-relative paths', () => {
  beforeEach(() => {
    clearInlineHtmlCacheForTests();
  });

  it('decodes server-relative encoded paths before calling SharePoint file API', async () => {
    const mockResponse = {
      ok: true,
      status: 200,
      statusText: 'OK',
      text: jest
        .fn()
        .mockResolvedValue('<html><head></head><body><h1>Encoded</h1></body></html>'),
    };
    const mockGet = jest.fn().mockResolvedValue(mockResponse);
    const mockSpHttpClient = {
      get: mockGet,
    };

    await loadSharePointFileContentForInline(
      mockSpHttpClient as never,
      'https://contoso.sharepoint.com/sites/TestSite2',
      '/sites/TestSite2/SiteAssets/My%20Reports/Quarter%20%231.html',
      '/sites/TestSite2/SiteAssets/My%20Reports/',
      'https://contoso.sharepoint.com/sites/TestSite2/SitePages/Dashboard.aspx',
    );

    const requestUrl = mockGet.mock.calls[0][0] as string;
    expect(requestUrl).toContain('My%20Reports');
    expect(requestUrl).toContain('Quarter%20%231.html');
    expect(requestUrl).not.toContain('%2520');
    expect(requestUrl).not.toContain('%2523');
  });
});
