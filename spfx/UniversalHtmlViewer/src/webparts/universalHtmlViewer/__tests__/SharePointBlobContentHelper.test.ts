import { prepareInlineHtmlForBlobUrl } from '../InlineHtmlTransformHelper';
import {
  clearInlineHtmlCacheForTests,
  loadSharePointFileContentForBlobUrl,
} from '../SharePointInlineContentHelper';

describe('prepareInlineHtmlForBlobUrl', () => {
  it('injects base href and history compatibility shim without default CSP', () => {
    const inputHtml = '<html><head><title>Report</title></head><body><h1>Dashboard</h1></body></html>';
    const result = prepareInlineHtmlForBlobUrl(
      inputHtml,
      '/sites/TestSite2/SiteAssets/GPOzaurr/',
      'https://contoso.sharepoint.com/sites/TestSite2/SitePages/Dashboard.aspx',
    );

    expect(result).not.toContain('data-uhv-inline-csp="1"');
    expect(result).toContain('data-uhv-history-compat="1"');
    expect(result).toContain(
      '<base href="https://contoso.sharepoint.com/sites/TestSite2/SiteAssets/GPOzaurr/">',
    );
  });
});

describe('loadSharePointFileContentForBlobUrl', () => {
  const webAbsoluteUrl = 'https://contoso.sharepoint.com/sites/TestSite2';
  const sourceUrl = '/sites/TestSite2/SiteAssets/Reports/index.html';
  const baseUrlForRelativeLinks = '/sites/TestSite2/SiteAssets/Reports/';
  const pageUrl = 'https://contoso.sharepoint.com/sites/TestSite2/SitePages/Dashboard.aspx';

  beforeEach(() => {
    clearInlineHtmlCacheForTests();
  });

  it('prepares blob iframe HTML without injecting the default inline CSP', async () => {
    const mockResponse = {
      ok: true,
      status: 200,
      statusText: 'OK',
      text: jest
        .fn()
        .mockResolvedValue('<html><head></head><body><h1>Blob</h1></body></html>'),
    };
    const mockGet = jest.fn().mockResolvedValue(mockResponse);
    const mockSpHttpClient = {
      get: mockGet,
    };

    const result = await loadSharePointFileContentForBlobUrl(
      mockSpHttpClient as never,
      webAbsoluteUrl,
      sourceUrl,
      baseUrlForRelativeLinks,
      pageUrl,
    );

    expect(result).toContain('<h1>Blob</h1>');
    expect(result).toContain('data-uhv-history-compat="1"');
    expect(result).not.toContain('data-uhv-inline-csp="1"');
  });
});
