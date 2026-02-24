import { prepareInlineHtmlForSrcDoc } from '../InlineHtmlTransformHelper';

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

