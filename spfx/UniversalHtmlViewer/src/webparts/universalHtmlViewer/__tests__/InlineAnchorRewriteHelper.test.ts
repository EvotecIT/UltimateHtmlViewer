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
    expect(result).toContain('data-uhv-inline-rewritten="1"');
    expect(result).toContain(
      'href="https://knauf.sharepoint.com/sites/TheDashboardPage/SitePages/TheDashboardPage.aspx?uhvPage=%2Fsites%2FTheDashboardPage%2FShared%2520Documents%2FActiveDirectoryOverall_Computers.html"',
    );
  });

  it('respects document base href when storing the original inline navigation target', () => {
    const inputHtml =
      '<html><head><base href="https://knauf.sharepoint.com/sites/TheDashboardPage/Shared%20Documents/Nested/"></head><body><a href="Computers.html">Computers</a></body></html>';
    const result = rewriteInlineNavigationAnchorHrefs(
      inputHtml,
      '/sites/TheDashboardPage/Shared Documents/',
      'https://knauf.sharepoint.com/sites/TheDashboardPage/SitePages/TheDashboardPage.aspx',
    );

    expect(result).toContain(
      'data-uhv-inline-href="https://knauf.sharepoint.com/sites/TheDashboardPage/Shared%20Documents/Nested/Computers.html"',
    );
    expect(result).toContain(
      'uhvPage=%2Fsites%2FTheDashboardPage%2FShared%2520Documents%2FNested%2FComputers.html',
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

  it('keeps eligible html downloads inside the viewer and preserves new-tab intent', () => {
    const inputHtml =
      '<html><body><a href="Download.html" download>Download</a><a href="NewTab.html" target="_blank">New tab</a><a href="Inline.html" target="_self">Inline</a></body></html>';
    const result = rewriteInlineNavigationAnchorHrefs(
      inputHtml,
      '/sites/TheDashboardPage/Shared Documents/',
      'https://knauf.sharepoint.com/sites/TheDashboardPage/SitePages/TheDashboardPage.aspx',
    );

    expect(result).not.toContain('download=""');
    expect(result).toContain('target="_blank"');
    expect(result).toContain('data-uhv-inline-href="https://knauf.sharepoint.com/sites/TheDashboardPage/Shared%20Documents/Download.html"');
    expect(result).toContain('data-uhv-inline-href="https://knauf.sharepoint.com/sites/TheDashboardPage/Shared%20Documents/NewTab.html"');
    expect(result).toContain('data-uhv-inline-href="https://knauf.sharepoint.com/sites/TheDashboardPage/Shared%20Documents/Inline.html"');
  });

  it('does not rewrite same-site html outside configured allowed path prefixes', () => {
    const inputHtml =
      '<html><body><a href="/sites/TheDashboardPage/Other/Outside.html">Outside</a></body></html>';
    const result = rewriteInlineNavigationAnchorHrefs(
      inputHtml,
      '/sites/TheDashboardPage/Shared Documents/',
      'https://knauf.sharepoint.com/sites/TheDashboardPage/SitePages/TheDashboardPage.aspx',
      {
        allowedPathPrefixes: ['/sites/TheDashboardPage/Shared Documents/'],
      },
    );

    expect(result).toContain('href="/sites/TheDashboardPage/Other/Outside.html"');
    expect(result).not.toContain('data-uhv-inline-href=');
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

  it('does not rewrite anchors whose deep-link query value would exceed the limit', () => {
    const longName = `${'a'.repeat(2100)}.html`;
    const inputHtml = `<html><body><a href="${longName}">Long report</a></body></html>`;
    const result = rewriteInlineNavigationAnchorHrefs(
      inputHtml,
      '/sites/TheDashboardPage/Shared Documents/',
      'https://knauf.sharepoint.com/sites/TheDashboardPage/SitePages/TheDashboardPage.aspx',
    );

    expect(result).toContain(`href="${longName}"`);
    expect(result).not.toContain('data-uhv-inline-href=');
  });

  it('uses configured allowed extensions when deciding whether anchors can be rewritten', () => {
    const inputHtml =
      '<html><body><a href="Report.aspx">Aspx</a><a href="Report.html">Html</a></body></html>';
    const result = rewriteInlineNavigationAnchorHrefs(
      inputHtml,
      '/sites/TheDashboardPage/Shared Documents/',
      'https://knauf.sharepoint.com/sites/TheDashboardPage/SitePages/TheDashboardPage.aspx',
      {
        allowedFileExtensions: ['.html'],
      },
    );

    expect(result).toContain('href="Report.aspx"');
    expect(result).toContain(
      'data-uhv-inline-href="https://knauf.sharepoint.com/sites/TheDashboardPage/Shared%20Documents/Report.html"',
    );
  });

  it('preserves configured host query parameters for dashboard-scoped deep links', () => {
    const inputHtml = '<html><body><a href="Computers.html">Computers</a></body></html>';
    const result = rewriteInlineNavigationAnchorHrefs(
      inputHtml,
      '/sites/TheDashboardPage/Shared Documents/Dashboards/ops/',
      'https://knauf.sharepoint.com/sites/TheDashboardPage/SitePages/TheDashboardPage.aspx?dashboard=ops&OR=Teams-HL&uhvPage=%2Fold',
      {
        preservedHostQueryParamNames: ['dashboard'],
      },
    );

    expect(result).toContain(
      'href="https://knauf.sharepoint.com/sites/TheDashboardPage/SitePages/TheDashboardPage.aspx?dashboard=ops&amp;uhvPage=%2Fsites%2FTheDashboardPage%2FShared%2520Documents%2FDashboards%2Fops%2FComputers.html"',
    );
    expect(result).not.toContain('OR=Teams-HL');
  });

  it('uses a viewer-specific deep-link query parameter', () => {
    const result = rewriteInlineNavigationAnchorHrefs(
      '<html><body><a href="Computers.html">Computers</a></body></html>',
      '/sites/TheDashboardPage/Shared Documents/',
      'https://knauf.sharepoint.com/sites/TheDashboardPage/SitePages/TheDashboardPage.aspx',
      {
        deepLinkQueryParamName: 'viewerTwoPage',
      },
    );

    expect(result).toContain('?viewerTwoPage=');
    expect(result).not.toContain('?uhvPage=');
  });

  it('removes untrusted original href data before rewriting authored anchors', () => {
    const inputHtml =
      '<html><body><a data-uhv-inline-href="https://knauf.sharepoint.com/sites/TheDashboardPage/Shared%20Documents/Hidden.html" href="Visible.html">Visible</a></body></html>';
    const result = rewriteInlineNavigationAnchorHrefs(
      inputHtml,
      '/sites/TheDashboardPage/Shared Documents/',
      'https://knauf.sharepoint.com/sites/TheDashboardPage/SitePages/TheDashboardPage.aspx',
    );

    expect(result).toContain(
      'data-uhv-inline-href="https://knauf.sharepoint.com/sites/TheDashboardPage/Shared%20Documents/Visible.html"',
    );
    expect(result).not.toContain('Hidden.html');
  });
});
