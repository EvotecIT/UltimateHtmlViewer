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

  it('preserves empty parameter values for absolute URL input', () => {
    const valueWithEquals = getQueryStringParam(
      'https://contoso.sharepoint.com/sites/Reports/SitePages/Dashboard.aspx?uhvPage=',
      'uhvPage',
    );
    const valueWithoutEquals = getQueryStringParam(
      'https://contoso.sharepoint.com/sites/Reports/SitePages/Dashboard.aspx?uhvPage',
      'uhvPage',
    );

    expect(valueWithEquals).toBe('');
    expect(valueWithoutEquals).toBe('');
  });

  it('supports site-relative URL parsing through fallback logic', () => {
    const value = getQueryStringParam(
      '/sites/Reports/SitePages/Dashboard.aspx?uhvPage=%2Fsites%2FReports%2FSiteAssets%2Fnested%2Freport.html',
      'uhvPage',
    );

    expect(value).toBe('/sites/Reports/SiteAssets/nested/report.html');
  });

  it('preserves empty parameter values in fallback parser', () => {
    const valueWithEquals = getQueryStringParam(
      '/sites/Reports/SitePages/Dashboard.aspx?uhvPage=',
      'uhvPage',
    );
    const valueWithoutEquals = getQueryStringParam(
      '/sites/Reports/SitePages/Dashboard.aspx?uhvPage&foo=bar',
      'uhvPage',
    );

    expect(valueWithEquals).toBe('');
    expect(valueWithoutEquals).toBe('');
  });

  it('returns undefined when url or parameter name is empty', () => {
    expect(getQueryStringParam('', 'uhvPage')).toBeUndefined();
    expect(getQueryStringParam('https://contoso.sharepoint.com', '')).toBeUndefined();
  });

  it('returns undefined for malformed encoded query payloads in fallback parser', () => {
    const value = getQueryStringParam(
      '/sites/Reports/SitePages/Dashboard.aspx?uhvPage=%E0%A4%A',
      'uhvPage',
    );

    expect(value).toBeUndefined();
  });

  it('continues to later matching values when an earlier fallback value is malformed', () => {
    const value = getQueryStringParam(
      '/sites/Reports/SitePages/Dashboard.aspx?uhvPage=%E0%A4%A&uhvPage=%2Fsites%2FReports%2FSiteAssets%2Freport.html',
      'uhvPage',
    );

    expect(value).toBe('/sites/Reports/SiteAssets/report.html');
  });

  it('supports plus as space and keeps first matching value in fallback parser', () => {
    const value = getQueryStringParam(
      '/sites/Reports/SitePages/Dashboard.aspx?dashboard=report+name&dashboard=override',
      'dashboard',
    );

    expect(value).toBe('report name');
  });

  it('preserves equals signs in fallback parser values', () => {
    const value = getQueryStringParam(
      '/sites/Reports/SitePages/Dashboard.aspx?token=abc=def==&foo=bar',
      'token',
    );

    expect(value).toBe('abc=def==');
  });

  it('does not include hash fragments in fallback parser values', () => {
    const value = getQueryStringParam(
      '/sites/Reports/SitePages/Dashboard.aspx?uhvPage=%2Fsites%2FReports%2FSiteAssets%2Freport.html#section-a',
      'uhvPage',
    );

    expect(value).toBe('/sites/Reports/SiteAssets/report.html');
  });
});
