import {
  clearInlineHtmlCacheForTests,
  loadSharePointFileContentForInline,
} from '../SharePointInlineContentHelper';

describe('SharePoint inline content cache context', () => {
  beforeEach(() => {
    clearInlineHtmlCacheForTests();
  });

  it('separates transformed HTML when sibling viewer state changes', async () => {
    const mockResponse = {
      ok: true,
      status: 200,
      statusText: 'OK',
      text: jest
        .fn()
        .mockResolvedValue(
          '<html><head></head><body><a href="Next.html">Next</a></body></html>',
        ),
    };
    const mockGet = jest.fn().mockResolvedValue(mockResponse);
    const mockSpHttpClient = { get: mockGet };
    const webAbsoluteUrl = 'https://contoso.sharepoint.com/sites/TestSite2';
    const sourceUrl = '/sites/TestSite2/SiteAssets/Reports/index.html';
    const baseUrl = '/sites/TestSite2/SiteAssets/Reports/';
    const rewriteOptions = {
      rewriteInlineAnchorHrefs: true,
      rewriteInlineAnchorAllowedFileExtensions: ['.html'],
      rewriteInlineAnchorAllowedPathPrefixes: ['/sites/TestSite2/SiteAssets/'],
      rewriteInlineAnchorDeepLinkQueryParamName: 'viewerOnePage',
    };

    const first = await loadSharePointFileContentForInline(
      mockSpHttpClient as never,
      webAbsoluteUrl,
      sourceUrl,
      baseUrl,
      'https://contoso.sharepoint.com/sites/TestSite2/SitePages/Dashboard.aspx?viewerTwoPage=first',
      undefined,
      rewriteOptions,
    );
    const second = await loadSharePointFileContentForInline(
      mockSpHttpClient as never,
      webAbsoluteUrl,
      sourceUrl,
      baseUrl,
      'https://contoso.sharepoint.com/sites/TestSite2/SitePages/Dashboard.aspx?viewerTwoPage=second',
      undefined,
      rewriteOptions,
    );

    expect(mockGet).toHaveBeenCalledTimes(2);
    expect(getAnchorQueryValue(first, 'viewerTwoPage')).toBe('first');
    expect(getAnchorQueryValue(second, 'viewerTwoPage')).toBe('second');
  });
});

function getAnchorQueryValue(html: string, queryParamName: string): string | undefined {
  const href = new DOMParser()
    .parseFromString(html, 'text/html')
    .querySelector('a')
    ?.getAttribute('href');
  return new URL(href || '').searchParams.get(queryParamName) || undefined;
}
