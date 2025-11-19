import { buildFinalUrl, isUrlAllowed, HtmlSourceMode } from '../UrlHelper';

describe('buildFinalUrl', () => {
  const currentPageUrl: string = 'https://contoso.sharepoint.com/sites/Reports/Pages/Dashboard.aspx';

  it('uses fullUrl as-is in FullUrl mode', () => {
    const mode: HtmlSourceMode = 'FullUrl';
    const url: string | null = buildFinalUrl({
      htmlSourceMode: mode,
      fullUrl: 'https://contoso.sharepoint.com/sites/Reports/Dashboards/sample.html',
    });

    expect(url).toBe(
      'https://contoso.sharepoint.com/sites/Reports/Dashboards/sample.html',
    );
  });

  it('returns null when fullUrl is empty in FullUrl mode', () => {
    const url: string | null = buildFinalUrl({
      htmlSourceMode: 'FullUrl',
      fullUrl: '   ',
    });

    expect(url).toBeNull();
  });

  it('joins basePath and relativePath in BasePathAndRelativePath mode', () => {
    const url: string | null = buildFinalUrl({
      htmlSourceMode: 'BasePathAndRelativePath',
      basePath: '/sites/Reports/Dashboards',
      relativePath: '/system1/index.html',
    });

    expect(url).toBe('/sites/Reports/Dashboards/system1/index.html');
  });

  it('builds URL using query string dashboard ID when present', () => {
    const url: string | null = buildFinalUrl({
      htmlSourceMode: 'BasePathAndDashboardId',
      basePath: '/sites/Reports/Dashboards',
      dashboardId: 'fallback-id',
      defaultFileName: 'index.html',
      queryStringParamName: 'dashboard',
      pageUrl: `${currentPageUrl}?dashboard=query-id`,
    });

    expect(url).toBe('/sites/Reports/Dashboards/query-id/index.html');
  });

  it('falls back to dashboardId when no query string parameter is present', () => {
    const url: string | null = buildFinalUrl({
      htmlSourceMode: 'BasePathAndDashboardId',
      basePath: '/sites/Reports/Dashboards',
      dashboardId: 'fallback-id',
      defaultFileName: 'index.html',
      queryStringParamName: 'dashboard',
      pageUrl: currentPageUrl,
    });

    expect(url).toBe('/sites/Reports/Dashboards/fallback-id/index.html');
  });

  it('uses default index.html when defaultFileName is empty', () => {
    const url: string | null = buildFinalUrl({
      htmlSourceMode: 'BasePathAndDashboardId',
      basePath: '/sites/Reports/Dashboards',
      dashboardId: 'abc',
      defaultFileName: '   ',
      queryStringParamName: 'dashboard',
      pageUrl: currentPageUrl,
    });

    expect(url).toBe('/sites/Reports/Dashboards/abc/index.html');
  });

  it('returns null when basePath is missing in BasePathAndRelativePath mode', () => {
    const url: string | null = buildFinalUrl({
      htmlSourceMode: 'BasePathAndRelativePath',
      basePath: '   ',
      relativePath: 'system1/index.html',
    });

    expect(url).toBeNull();
  });

  it('returns null when dashboard ID cannot be determined', () => {
    const url: string | null = buildFinalUrl({
      htmlSourceMode: 'BasePathAndDashboardId',
      basePath: '/sites/Reports/Dashboards',
      dashboardId: '   ',
      defaultFileName: 'index.html',
      queryStringParamName: 'dashboard',
      pageUrl: currentPageUrl,
    });

    expect(url).toBeNull();
  });
});

describe('isUrlAllowed', () => {
  const currentPageUrl: string = 'https://contoso.sharepoint.com/sites/Reports/Pages/Dashboard.aspx';

  it('allows site-relative URLs', () => {
    expect(isUrlAllowed('/sites/Reports/Dashboards/sample.html', currentPageUrl)).toBe(
      true,
    );
  });

  it('allows absolute URLs on the same host', () => {
    expect(
      isUrlAllowed(
        'https://contoso.sharepoint.com/sites/Reports/Dashboards/sample.html',
        currentPageUrl,
      ),
    ).toBe(true);
  });

  it('rejects absolute URLs on a different host', () => {
    expect(
      isUrlAllowed(
        'https://fabrikam.sharepoint.com/sites/Reports/Dashboards/sample.html',
        currentPageUrl,
      ),
    ).toBe(false);
  });

  it('rejects javascript: URLs', () => {
    expect(isUrlAllowed('javascript:alert("x")', currentPageUrl)).toBe(false);
  });

  it('rejects unsupported URL formats', () => {
    expect(isUrlAllowed('ftp://contoso.sharepoint.com/file.html', currentPageUrl)).toBe(
      false,
    );
  });
});

