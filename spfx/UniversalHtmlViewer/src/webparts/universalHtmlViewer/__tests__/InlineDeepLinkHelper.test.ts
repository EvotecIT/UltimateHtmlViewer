import {
  buildPageUrlWithoutInlineDeepLink,
  buildPageUrlWithInlineDeepLink,
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
});

