import {
  buildOpenInNewTabUrl,
  buildPageUrlWithoutInlineDeepLink,
  buildPageUrlWithInlineDeepLink,
  resolveInlineContentTarget,
  resolveInlineDeepLinkTarget,
} from '../InlineDeepLinkHelper';
import { UrlValidationOptions } from '../UrlHelper';

describe('InlineDeepLinkHelper', () => {
  const validationOptions: UrlValidationOptions = {
    securityMode: 'StrictTenant',
    currentPageUrl: 'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx',
    allowHttp: false,
    allowedPathPrefixes: ['/sites/testsite1/siteassets/reports/'],
    allowedFileExtensions: ['.html', '.htm', '.aspx'],
  };

  it('resolves same-site server-relative deep link value', () => {
    const result = resolveInlineDeepLinkTarget({
      pageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx?uhvPage=%2Fsites%2FTestSite1%2FSiteAssets%2FReports%2FGPO_Blocked_Inheritance.html',
      fallbackUrl: '/sites/TestSite1/SiteAssets/Reports/index.html',
      validationOptions,
    });

    expect(result).toBe(
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/GPO_Blocked_Inheritance.html',
    );
  });

  it('ignores deep links outside allowed prefixes', () => {
    const result = resolveInlineDeepLinkTarget({
      pageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx?uhvPage=%2Fsites%2FTestSite1%2FSiteAssets%2FOutside%2FReport.html',
      fallbackUrl: '/sites/TestSite1/SiteAssets/Reports/index.html',
      validationOptions,
    });

    expect(result).toBeUndefined();
  });

  it('rejects deep-link values with backslashes', () => {
    const result = resolveInlineDeepLinkTarget({
      pageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx?uhvPage=%5Csites%5CTestSite1%5CSiteAssets%5CReports%5CGPO_List.html',
      fallbackUrl: '/sites/TestSite1/SiteAssets/Reports/index.html',
      validationOptions,
    });

    expect(result).toBeUndefined();
  });

  it('rejects deep-link values above maximum supported length', () => {
    const oversizedValue = `/sites/TestSite1/SiteAssets/Reports/${'a'.repeat(2100)}.html`;
    const pageUrl = `https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx?uhvPage=${encodeURIComponent(
      oversizedValue,
    )}`;
    const result = resolveInlineDeepLinkTarget({
      pageUrl,
      fallbackUrl: '/sites/TestSite1/SiteAssets/Reports/index.html',
      validationOptions,
    });

    expect(result).toBeUndefined();
  });

  it('writes deep link param back to current page URL', () => {
    const result = buildPageUrlWithInlineDeepLink({
      pageUrl: 'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx',
      targetUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/GPO_Broken.html',
    });

    expect(result).toBe(
      'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx?uhvPage=%2Fsites%2FTestSite1%2FSiteAssets%2FReports%2FGPO_Broken.html',
    );
  });

  it('preserves existing page query parameters when writing deep links', () => {
    const result = buildPageUrlWithInlineDeepLink({
      pageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx?dashboard=ops&foo=bar',
      targetUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/GPO_List.html',
    });

    expect(result).toBe(
      'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx?dashboard=ops&foo=bar&uhvPage=%2Fsites%2FTestSite1%2FSiteAssets%2FReports%2FGPO_List.html',
    );
  });

  it('removes deep link parameter and keeps other query parameters', () => {
    const result = buildPageUrlWithoutInlineDeepLink({
      pageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx?dashboard=ops&uhvPage=%2Fsites%2FTestSite1%2FSiteAssets%2FReports%2FGPO_List.html&foo=bar',
    });

    expect(result).toBe(
      'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx?dashboard=ops&foo=bar',
    );
  });

  it('builds open-in-new-tab deep link from current target URL in inline mode', () => {
    const href = buildOpenInNewTabUrl({
      resolvedUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/Resolved.html',
      baseUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/Current.html',
      pageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx?legacy=1',
      currentPageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx?dashboard=ops',
      contentDeliveryMode: 'SharePointFileContent',
    });

    expect(href).toBe(
      'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx?dashboard=ops&uhvPage=%2Fsites%2FTestSite1%2FSiteAssets%2FReports%2FCurrent.html',
    );
  });

  it('returns resolved URL for direct mode open-in-new-tab links', () => {
    const href = buildOpenInNewTabUrl({
      resolvedUrl: 'https://external.example/report.html?v=1',
      baseUrl: 'https://external.example/report.html',
      pageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx?legacy=1',
      contentDeliveryMode: 'DirectUrl',
    });

    expect(href).toBe('https://external.example/report.html?v=1');
  });

  it('resolves initial content using deep link when allowed', () => {
    const result = resolveInlineContentTarget({
      pageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx?uhvPage=%2Fsites%2FTestSite1%2FSiteAssets%2FReports%2FOps.html',
      fallbackUrl: '/sites/TestSite1/SiteAssets/Reports/index.html',
      validationOptions,
    });

    expect(result.allowDeepLinkOverride).toBe(true);
    expect(result.hasRequestedDeepLink).toBe(true);
    expect(result.isRejectedRequestedDeepLink).toBe(false);
    expect(result.initialContentUrl).toBe(
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/Ops.html',
    );
  });

  it('ignores deep-link override in AnyHttps mode', () => {
    const anyHttpsOptions: UrlValidationOptions = {
      ...validationOptions,
      securityMode: 'AnyHttps',
      allowedPathPrefixes: undefined,
      allowedFileExtensions: undefined,
    };
    const result = resolveInlineContentTarget({
      pageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx?uhvPage=https%3A%2F%2Fexternal.example%2Freport.html',
      fallbackUrl: '/sites/TestSite1/SiteAssets/Reports/index.html',
      validationOptions: anyHttpsOptions,
    });

    expect(result.allowDeepLinkOverride).toBe(false);
    expect(result.hasRequestedDeepLink).toBe(true);
    expect(result.isRejectedRequestedDeepLink).toBe(false);
    expect(result.deepLinkedUrl).toBeUndefined();
    expect(result.initialContentUrl).toBe('/sites/TestSite1/SiteAssets/Reports/index.html');
  });

  it('supports explicitly disabling deep-link overrides from query string', () => {
    const result = resolveInlineContentTarget({
      pageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx?uhvPage=%2Fsites%2FTestSite1%2FSiteAssets%2FReports%2FOps.html',
      fallbackUrl: '/sites/TestSite1/SiteAssets/Reports/index.html',
      validationOptions,
      allowDeepLinkOverride: false,
    });

    expect(result.allowDeepLinkOverride).toBe(false);
    expect(result.hasRequestedDeepLink).toBe(true);
    expect(result.isRejectedRequestedDeepLink).toBe(false);
    expect(result.deepLinkedUrl).toBeUndefined();
    expect(result.initialContentUrl).toBe('/sites/TestSite1/SiteAssets/Reports/index.html');
  });
});

