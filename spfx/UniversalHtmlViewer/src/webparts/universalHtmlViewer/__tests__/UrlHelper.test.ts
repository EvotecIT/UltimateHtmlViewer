import {
  buildFinalUrl,
  isUrlAllowed,
  HtmlSourceMode,
  UrlSecurityMode,
} from '../UrlHelper';

describe('buildFinalUrl', () => {
  const currentPageUrl: string = 'https://contoso.sharepoint.com/sites/Reports/Pages/Dashboard.aspx';

  it('uses fullUrl as-is in FullUrl mode', () => {
    const mode: HtmlSourceMode = 'FullUrl';
    const url: string | undefined = buildFinalUrl({
      htmlSourceMode: mode,
      fullUrl: 'https://contoso.sharepoint.com/sites/Reports/Dashboards/sample.html',
    });

    expect(url).toBe(
      'https://contoso.sharepoint.com/sites/Reports/Dashboards/sample.html',
    );
  });

  it('returns null when fullUrl is empty in FullUrl mode', () => {
    const url: string | undefined = buildFinalUrl({
      htmlSourceMode: 'FullUrl',
      fullUrl: '   ',
    });

    expect(url).toBeUndefined();
  });

  it('joins basePath and relativePath in BasePathAndRelativePath mode', () => {
    const url: string | undefined = buildFinalUrl({
      htmlSourceMode: 'BasePathAndRelativePath',
      basePath: '/sites/Reports/Dashboards',
      relativePath: '/system1/index.html',
    });

    expect(url).toBe('/sites/Reports/Dashboards/system1/index.html');
  });

  it('builds URL using query string dashboard ID when present', () => {
    const url: string | undefined = buildFinalUrl({
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
    const url: string | undefined = buildFinalUrl({
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
    const url: string | undefined = buildFinalUrl({
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
    const url: string | undefined = buildFinalUrl({
      htmlSourceMode: 'BasePathAndRelativePath',
      basePath: '   ',
      relativePath: 'system1/index.html',
    });

    expect(url).toBeUndefined();
  });

  it('returns null when dashboard ID cannot be determined', () => {
    const url: string | undefined = buildFinalUrl({
      htmlSourceMode: 'BasePathAndDashboardId',
      basePath: '/sites/Reports/Dashboards',
      dashboardId: '   ',
      defaultFileName: 'index.html',
      queryStringParamName: 'dashboard',
      pageUrl: currentPageUrl,
    });

    expect(url).toBeUndefined();
  });
});

describe('isUrlAllowed', () => {
  const currentPageUrl: string = 'https://contoso.sharepoint.com/sites/Reports/Pages/Dashboard.aspx';
  const strictOptions = {
    securityMode: 'StrictTenant' as UrlSecurityMode,
    currentPageUrl,
  };

  it('allows site-relative URLs', () => {
    expect(isUrlAllowed('/sites/Reports/Dashboards/sample.html', strictOptions)).toBe(
      true,
    );
  });

  it('allows absolute URLs on the same host', () => {
    expect(
      isUrlAllowed(
        'https://contoso.sharepoint.com/sites/Reports/Dashboards/sample.html',
        strictOptions,
      ),
    ).toBe(true);
  });

  it('rejects absolute URLs on a different host', () => {
    expect(
      isUrlAllowed(
        'https://fabrikam.sharepoint.com/sites/Reports/Dashboards/sample.html',
        strictOptions,
      ),
    ).toBe(false);
  });

  it('rejects javascript: URLs', () => {
    const scriptUrl = ['java', 'script:alert("x")'].join('');
    expect(isUrlAllowed(scriptUrl, strictOptions)).toBe(false);
  });

  it('rejects unsupported URL formats', () => {
    expect(isUrlAllowed('ftp://contoso.sharepoint.com/file.html', strictOptions)).toBe(
      false,
    );
  });

  it('rejects protocol-relative URLs', () => {
    expect(isUrlAllowed('//evil.example.com/file.html', strictOptions)).toBe(false);
  });

  it('rejects data: URLs', () => {
    expect(isUrlAllowed('data:text/html;base64,AAA', strictOptions)).toBe(false);
  });

  it('allows allowlisted hosts when configured', () => {
    expect(
      isUrlAllowed('https://cdn.contoso.com/report.html', {
        securityMode: 'Allowlist',
        currentPageUrl,
        allowedHosts: ['cdn.contoso.com'],
      }),
    ).toBe(true);
  });

  it('rejects non-allowlisted hosts when in allowlist mode', () => {
    expect(
      isUrlAllowed('https://cdn.fabrikam.com/report.html', {
        securityMode: 'Allowlist',
        currentPageUrl,
        allowedHosts: ['cdn.contoso.com'],
      }),
    ).toBe(false);
  });

  it('supports wildcard suffix hosts in allowlist mode', () => {
    expect(
      isUrlAllowed('https://assets.sharepoint.com/report.html', {
        securityMode: 'Allowlist',
        currentPageUrl,
        allowedHosts: ['.sharepoint.com'],
      }),
    ).toBe(true);
  });

  it('allows any https in AnyHttps mode but rejects http', () => {
    expect(
      isUrlAllowed('https://external.example.com/report.html', {
        securityMode: 'AnyHttps',
        currentPageUrl,
      }),
    ).toBe(true);

    expect(
      isUrlAllowed('http://external.example.com/report.html', {
        securityMode: 'AnyHttps',
        currentPageUrl,
      }),
    ).toBe(false);
  });

  it('blocks http by default but allows it when explicitly enabled', () => {
    expect(
      isUrlAllowed('http://contoso.sharepoint.com/sites/Reports/Dashboards/a.html', {
        securityMode: 'StrictTenant',
        currentPageUrl,
      }),
    ).toBe(false);

    expect(
      isUrlAllowed('http://contoso.sharepoint.com/sites/Reports/Dashboards/a.html', {
        securityMode: 'StrictTenant',
        currentPageUrl,
        allowHttp: true,
      }),
    ).toBe(true);
  });

  it('enforces allowed path prefixes', () => {
    const options = {
      securityMode: 'StrictTenant' as UrlSecurityMode,
      currentPageUrl,
      allowedPathPrefixes: ['/sites/Reports/Dashboards'],
    };

    expect(isUrlAllowed('/sites/Reports/Dashboards/a.html', options)).toBe(true);
    expect(isUrlAllowed('/sites/Reports/Other/a.html', options)).toBe(false);
    expect(
      isUrlAllowed(
        'https://contoso.sharepoint.com/sites/Reports/Dashboards/a.html',
        options,
      ),
    ).toBe(true);
  });

  it('treats allowed path prefixes as case-insensitive', () => {
    const options = {
      securityMode: 'StrictTenant' as UrlSecurityMode,
      currentPageUrl,
      allowedPathPrefixes: ['/sites/reports/dashboards'],
    };

    expect(isUrlAllowed('/Sites/Reports/Dashboards/a.html', options)).toBe(true);
  });

  it('enforces allowed file extensions when configured', () => {
    const options = {
      securityMode: 'StrictTenant' as UrlSecurityMode,
      currentPageUrl,
      allowedFileExtensions: ['.html', '.htm', '.aspx'],
    };

    expect(isUrlAllowed('/sites/Reports/Dashboards/a.html', options)).toBe(true);
    expect(isUrlAllowed('/sites/Reports/Dashboards/a.HTM', options)).toBe(true);
    expect(isUrlAllowed('/sites/Reports/Dashboards/a.aspx', options)).toBe(true);
    expect(isUrlAllowed('/sites/Reports/Dashboards/a.js', options)).toBe(false);
    expect(isUrlAllowed('/sites/Reports/Dashboards/', options)).toBe(false);
  });

  it('rejects dot-segment paths', () => {
    expect(
      isUrlAllowed('/sites/Reports/Dashboards/../Secret/index.html', strictOptions),
    ).toBe(false);
  });

  it('rejects encoded dot-segment paths', () => {
    expect(
      isUrlAllowed('/sites/Reports/Dashboards/%2e%2e/Secret/index.html', strictOptions),
    ).toBe(false);

    expect(
      isUrlAllowed(
        'https://contoso.sharepoint.com/sites/Reports/Dashboards/%2E%2E/Secret/index.html',
        strictOptions,
      ),
    ).toBe(false);
  });

  it('rejects backslash dot-segment paths', () => {
    expect(
      isUrlAllowed('/sites/Reports/Dashboards/..\\Secret/index.html', strictOptions),
    ).toBe(false);

    expect(
      isUrlAllowed(
        'https://contoso.sharepoint.com/sites/Reports/Dashboards/..\\Secret/index.html',
        strictOptions,
      ),
    ).toBe(false);
  });

  it('rejects encoded slash and backslash in dot-segment traversal paths', () => {
    expect(
      isUrlAllowed(
        '/sites/Reports/Dashboards/%2e%2e%5CSecret/index.html',
        strictOptions,
      ),
    ).toBe(false);

    expect(
      isUrlAllowed(
        'https://contoso.sharepoint.com/sites/Reports/Dashboards/%2e%2e%2FSecret/index.html',
        strictOptions,
      ),
    ).toBe(false);
  });
});
