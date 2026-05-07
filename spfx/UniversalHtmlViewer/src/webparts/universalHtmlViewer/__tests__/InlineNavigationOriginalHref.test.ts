import { resolveInlineNavigationTarget } from '../InlineNavigationHelper';
import { UrlValidationOptions } from '../UrlHelper';

describe('InlineNavigationHelper original href data', () => {
  const validationOptions: UrlValidationOptions = {
    securityMode: 'StrictTenant',
    currentPageUrl: 'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Home.aspx',
    allowHttp: false,
    allowedPathPrefixes: ['/sites/testsite1/siteassets/'],
    allowedFileExtensions: ['.html', '.htm', '.aspx'],
  };

  it('prefers UHV original href data over host-page deep-link hrefs', () => {
    const anchor = document.createElement('a');
    anchor.href =
      'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Home.aspx?uhvPage=%2Fsites%2FTestSite1%2FSiteAssets%2FReports%2FComputers.html';
    anchor.setAttribute(
      'href',
      'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Home.aspx?uhvPage=%2Fsites%2FTestSite1%2FSiteAssets%2FReports%2FComputers.html',
    );
    anchor.setAttribute(
      'data-uhv-inline-href',
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/Computers.html',
    );
    anchor.setAttribute('data-uhv-inline-rewritten', '1');
    const clickEvent = new MouseEvent('click', { bubbles: true, button: 0 });
    Object.defineProperty(clickEvent, 'target', {
      value: anchor,
      configurable: true,
    });

    const result = resolveInlineNavigationTarget(clickEvent, {
      currentPageUrl: validationOptions.currentPageUrl,
      validationOptions,
      cacheBusterParamName: 'v',
    });

    expect(result).toBe(
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/Computers.html',
    );
  });

  it('ignores authored original href data that was not stamped by UHV rewriting', () => {
    const anchor = document.createElement('a');
    anchor.href = 'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/Visible.html';
    anchor.setAttribute(
      'href',
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/Visible.html',
    );
    anchor.setAttribute(
      'data-uhv-inline-href',
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/Hidden.html',
    );
    const clickEvent = new MouseEvent('click', { bubbles: true, button: 0 });
    Object.defineProperty(clickEvent, 'target', {
      value: anchor,
      configurable: true,
    });

    const result = resolveInlineNavigationTarget(clickEvent, {
      currentPageUrl: validationOptions.currentPageUrl,
      validationOptions,
      cacheBusterParamName: 'v',
    });

    expect(result).toBe(
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/Visible.html',
    );
  });
});
