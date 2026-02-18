import { resolveInlineNavigationTarget } from '../InlineNavigationHelper';
import { UrlValidationOptions } from '../UrlHelper';

describe('InlineNavigationHelper', () => {
  const validationOptions: UrlValidationOptions = {
    securityMode: 'StrictTenant',
    currentPageUrl: 'https://evotecpoland.sharepoint.com/sites/TestUHV1/SitePages/Home.aspx',
    allowHttp: false,
    allowedPathPrefixes: ['/sites/testuhv1/siteassets/'],
    allowedFileExtensions: ['.html', '.htm', '.aspx'],
  };

  it('resolves relative html links inside allowed path and strips cache param', () => {
    const anchor = document.createElement('a');
    anchor.href = 'https://evotecpoland.sharepoint.com/sites/TestUHV1/SiteAssets/Reports/index.html?v=123';
    anchor.setAttribute(
      'href',
      'https://evotecpoland.sharepoint.com/sites/TestUHV1/SiteAssets/Reports/index.html?v=123',
    );
    const clickEvent = new MouseEvent('click', { bubbles: true, button: 0 });
    Object.defineProperty(clickEvent, 'target', {
      value: anchor,
      configurable: true,
    });

    const result = resolveInlineNavigationTarget(clickEvent, {
      currentPageUrl:
        'https://evotecpoland.sharepoint.com/sites/TestUHV1/SiteAssets/Reports/start.html',
      validationOptions,
      cacheBusterParamName: 'v',
    });

    expect(result).toBe(
      'https://evotecpoland.sharepoint.com/sites/TestUHV1/SiteAssets/Reports/index.html',
    );
  });

  it('does not intercept external host links', () => {
    const anchor = document.createElement('a');
    anchor.href = 'https://example.org/report.html';
    anchor.setAttribute('href', 'https://example.org/report.html');
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

    expect(result).toBeUndefined();
  });

  it('does not intercept non-html links', () => {
    const anchor = document.createElement('a');
    anchor.href = 'https://evotecpoland.sharepoint.com/sites/TestUHV1/SiteAssets/Reports/file.csv';
    anchor.setAttribute(
      'href',
      'https://evotecpoland.sharepoint.com/sites/TestUHV1/SiteAssets/Reports/file.csv',
    );
    const clickEvent = new MouseEvent('click', { bubbles: true, button: 0 });
    Object.defineProperty(clickEvent, 'target', {
      value: anchor,
      configurable: true,
    });

    const result = resolveInlineNavigationTarget(clickEvent, {
      currentPageUrl:
        'https://evotecpoland.sharepoint.com/sites/TestUHV1/SiteAssets/Reports/start.html',
      validationOptions,
      cacheBusterParamName: 'v',
    });

    expect(result).toBeUndefined();
  });

  it('resolves .aspx links when extension is allowed', () => {
    const anchor = document.createElement('a');
    anchor.href = 'https://evotecpoland.sharepoint.com/sites/TestUHV1/SiteAssets/Reports/page.aspx?v=123';
    anchor.setAttribute(
      'href',
      'https://evotecpoland.sharepoint.com/sites/TestUHV1/SiteAssets/Reports/page.aspx?v=123',
    );
    const clickEvent = new MouseEvent('click', { bubbles: true, button: 0 });
    Object.defineProperty(clickEvent, 'target', {
      value: anchor,
      configurable: true,
    });

    const result = resolveInlineNavigationTarget(clickEvent, {
      currentPageUrl:
        'https://evotecpoland.sharepoint.com/sites/TestUHV1/SiteAssets/Reports/start.html',
      validationOptions,
      cacheBusterParamName: 'v',
    });

    expect(result).toBe(
      'https://evotecpoland.sharepoint.com/sites/TestUHV1/SiteAssets/Reports/page.aspx',
    );
  });

  it('does not intercept .aspx links when extension is not allowed', () => {
    const anchor = document.createElement('a');
    anchor.href = 'https://evotecpoland.sharepoint.com/sites/TestUHV1/SiteAssets/Reports/page.aspx';
    anchor.setAttribute(
      'href',
      'https://evotecpoland.sharepoint.com/sites/TestUHV1/SiteAssets/Reports/page.aspx',
    );
    const clickEvent = new MouseEvent('click', { bubbles: true, button: 0 });
    Object.defineProperty(clickEvent, 'target', {
      value: anchor,
      configurable: true,
    });

    const result = resolveInlineNavigationTarget(clickEvent, {
      currentPageUrl:
        'https://evotecpoland.sharepoint.com/sites/TestUHV1/SiteAssets/Reports/start.html',
      validationOptions: {
        ...validationOptions,
        allowedFileExtensions: ['.html', '.htm'],
      },
      cacheBusterParamName: 'v',
    });

    expect(result).toBeUndefined();
  });

  it('does not intercept links targeting a new tab', () => {
    const anchor = document.createElement('a');
    anchor.href = 'https://evotecpoland.sharepoint.com/sites/TestUHV1/SiteAssets/Reports/index.html';
    anchor.setAttribute(
      'href',
      'https://evotecpoland.sharepoint.com/sites/TestUHV1/SiteAssets/Reports/index.html',
    );
    anchor.setAttribute('target', '_blank');
    const clickEvent = new MouseEvent('click', { bubbles: true, button: 0 });
    Object.defineProperty(clickEvent, 'target', {
      value: anchor,
      configurable: true,
    });

    const result = resolveInlineNavigationTarget(clickEvent, {
      currentPageUrl:
        'https://evotecpoland.sharepoint.com/sites/TestUHV1/SiteAssets/Reports/start.html',
      validationOptions,
      cacheBusterParamName: 'v',
    });

    expect(result).toBeUndefined();
  });
});
