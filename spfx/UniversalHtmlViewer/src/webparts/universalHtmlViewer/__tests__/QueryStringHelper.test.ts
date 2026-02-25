import { getQueryStringParam } from '../QueryStringHelper';

describe('QueryStringHelper', () => {
  it('returns decoded value for valid absolute URL input', () => {
    const value = getQueryStringParam(
      'https://contoso.sharepoint.com/sites/Reports/SitePages/Dashboard.aspx?uhvPage=%2Fsites%2FReports%2FSiteAssets%2Freport.html',
      'uhvPage',
    );

    expect(value).toBe('/sites/Reports/SiteAssets/report.html');
  });

  it('returns undefined when parameter is missing', () => {
    const value = getQueryStringParam(
      'https://contoso.sharepoint.com/sites/Reports/SitePages/Dashboard.aspx?foo=bar',
      'uhvPage',
    );

    expect(value).toBeUndefined();
  });

  it('supports site-relative URL parsing through fallback logic', () => {
    const value = getQueryStringParam(
      '/sites/Reports/SitePages/Dashboard.aspx?uhvPage=%2Fsites%2FReports%2FSiteAssets%2Fnested%2Freport.html',
      'uhvPage',
    );

    expect(value).toBe('/sites/Reports/SiteAssets/nested/report.html');
  });

  it('returns undefined when url or parameter name is empty', () => {
    expect(getQueryStringParam('', 'uhvPage')).toBeUndefined();
    expect(getQueryStringParam('https://contoso.sharepoint.com', '')).toBeUndefined();
  });
});
