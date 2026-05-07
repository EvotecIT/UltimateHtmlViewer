import { prepareInlineHtmlForSrcDoc } from '../InlineHtmlTransformHelper';

describe('InlineHtmlTransformHelper anchor rewrite gating', () => {
  it('does not rewrite report anchors unless host deep-link anchors are enabled', () => {
    const result = prepareInlineHtmlForSrcDoc(
      '<html><body><a href="Computers.html">Computers</a></body></html>',
      '/sites/TheDashboardPage/Shared Documents/',
      'https://knauf.sharepoint.com/sites/TheDashboardPage/SitePages/TheDashboardPage.aspx',
    );

    expect(result).toContain('href="Computers.html"');
    expect(result).not.toContain('data-uhv-inline-href=');
  });

  it('rewrites report anchors when host deep-link anchors are enabled', () => {
    const result = prepareInlineHtmlForSrcDoc(
      '<html><body><a href="Computers.html">Computers</a></body></html>',
      '/sites/TheDashboardPage/Shared Documents/',
      'https://knauf.sharepoint.com/sites/TheDashboardPage/SitePages/TheDashboardPage.aspx?OR=Teams-HL',
      {
        rewriteInlineAnchorHrefs: true,
      },
    );

    expect(result).toContain('data-uhv-inline-href=');
    expect(result).toContain(
      'href="https://knauf.sharepoint.com/sites/TheDashboardPage/SitePages/TheDashboardPage.aspx?uhvPage=',
    );
    expect(result).not.toContain('OR=Teams-HL');
  });
});
