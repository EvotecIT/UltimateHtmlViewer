import { rewriteInlineNavigationAnchorHrefs } from '../InlineAnchorRewriteHelper';

describe('InlineAnchorRewriteHelper', () => {
  it('rewrites same-site html anchors to host-page deep links for browser context menu actions', () => {
    const inputHtml =
      '<!DOCTYPE html><html><head><title>Report</title></head><body><a href="ActiveDirectoryOverall_Computers.html">Computers</a></body></html>';
    const result = rewriteInlineNavigationAnchorHrefs(
      inputHtml,
      '/sites/TheDashboardPage/Shared Documents/',
      'https://knauf.sharepoint.com/sites/TheDashboardPage/SitePages/TheDashboardPage.aspx?isSPOFile=1',
    );

    expect(result).toContain(
      'data-uhv-inline-href="https://knauf.sharepoint.com/sites/TheDashboardPage/Shared%20Documents/ActiveDirectoryOverall_Computers.html"',
    );
    expect(result).toContain(
      'href="https://knauf.sharepoint.com/sites/TheDashboardPage/SitePages/TheDashboardPage.aspx?isSPOFile=1&amp;uhvPage=%2Fsites%2FTheDashboardPage%2FShared%2520Documents%2FActiveDirectoryOverall_Computers.html"',
    );
  });

  it('keeps external anchors unchanged', () => {
    const inputHtml =
      '<html><head><title>Report</title></head><body><a href="https://example.org/report.html">External</a></body></html>';
    const result = rewriteInlineNavigationAnchorHrefs(
      inputHtml,
      '/sites/TheDashboardPage/Shared Documents/',
      'https://knauf.sharepoint.com/sites/TheDashboardPage/SitePages/TheDashboardPage.aspx',
    );

    expect(result).toContain('href="https://example.org/report.html"');
    expect(result).not.toContain('data-uhv-inline-href="https://example.org/report.html"');
  });

  it('keeps same-site anchors outside the report base directory unchanged', () => {
    const inputHtml =
      '<html><body><a href="/sites/TheDashboardPage/SitePages/Other.aspx">Other page</a></body></html>';
    const result = rewriteInlineNavigationAnchorHrefs(
      inputHtml,
      '/sites/TheDashboardPage/Shared Documents/',
      'https://knauf.sharepoint.com/sites/TheDashboardPage/SitePages/TheDashboardPage.aspx',
    );

    expect(result).toContain('href="/sites/TheDashboardPage/SitePages/Other.aspx"');
    expect(result).not.toContain('data-uhv-inline-href=');
  });
});
