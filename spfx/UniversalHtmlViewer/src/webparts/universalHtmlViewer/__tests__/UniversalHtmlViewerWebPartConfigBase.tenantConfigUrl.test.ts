/* eslint-disable @typescript-eslint/no-explicit-any, @typescript-eslint/no-var-requires */
jest.mock('@microsoft/sp-http', () => ({
  SPHttpClient: {
    configurations: {
      v1: {},
    },
  },
  SPHttpClientResponse: class {},
}));
jest.mock('@microsoft/sp-webpart-base', () => ({
  BaseClientSideWebPart: class {},
}));

const {
  UniversalHtmlViewerWebPartConfigBase,
}: {
  UniversalHtmlViewerWebPartConfigBase: any;
} = require('../UniversalHtmlViewerWebPartConfigBase');

function createConfigHarness(): any {
  const configBase = Object.create(
    UniversalHtmlViewerWebPartConfigBase.prototype,
  ) as any;

  configBase.context = {
    pageContext: {
      web: {
        absoluteUrl: 'https://contoso.sharepoint.com/sites/TestSite1',
        serverRelativeUrl: '/sites/TestSite1',
      },
    },
  };

  return configBase;
}

describe('UniversalHtmlViewerWebPartConfigBase resolveTenantConfigUrl', () => {
  const pageUrl = 'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx';

  it('accepts mixed-case HTTPS tenant config URLs on the same host', () => {
    const configBase = createConfigHarness();

    const resolved = (configBase as any).resolveTenantConfigUrl(
      'HTTPS://contoso.sharepoint.com/sites/TestSite1/SiteAssets/uhv-config.json',
      pageUrl,
    );

    expect(resolved).toBe(
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/uhv-config.json',
    );
  });

  it('rejects mixed-case HTTP and unsupported absolute schemes', () => {
    const configBase = createConfigHarness();

    const httpResolved = (configBase as any).resolveTenantConfigUrl(
      'HTTP://contoso.sharepoint.com/sites/TestSite1/SiteAssets/uhv-config.json',
      pageUrl,
    );
    const ftpResolved = (configBase as any).resolveTenantConfigUrl(
      'ftp://contoso.sharepoint.com/sites/TestSite1/SiteAssets/uhv-config.json',
      pageUrl,
    );

    expect(httpResolved).toBeUndefined();
    expect(ftpResolved).toBeUndefined();
  });

  it('rejects cross-host HTTPS tenant config URLs', () => {
    const configBase = createConfigHarness();

    const resolved = (configBase as any).resolveTenantConfigUrl(
      'https://fabrikam.sharepoint.com/sites/TestSite1/SiteAssets/uhv-config.json',
      pageUrl,
    );

    expect(resolved).toBeUndefined();
  });

  it('resolves server-relative and web-relative tenant config paths', () => {
    const configBase = createConfigHarness();

    const serverRelativeResolved = (configBase as any).resolveTenantConfigUrl(
      '/sites/TestSite1/SiteAssets/uhv-config.json',
      pageUrl,
    );
    const webRelativeResolved = (configBase as any).resolveTenantConfigUrl(
      'SiteAssets/uhv-config.json',
      pageUrl,
    );

    expect(serverRelativeResolved).toBe(
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/uhv-config.json',
    );
    expect(webRelativeResolved).toBe(
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/uhv-config.json',
    );
  });
});
