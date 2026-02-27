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

function createConfigHarness(spHttpGet?: jest.Mock): any {
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
    spHttpClient: {
      get: spHttpGet || jest.fn(),
    },
  };

  return configBase;
}

describe('UniversalHtmlViewerWebPartConfigBase tenant config load errors', () => {
  const pageUrl = 'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Dashboard.aspx';

  it('tracks invalid tenant config URL resolution errors', async () => {
    const configBase = createConfigHarness();

    const result = await (configBase as any).tryLoadTenantConfig(pageUrl, {
      tenantConfigUrl: 'http://contoso.sharepoint.com/sites/TestSite1/SiteAssets/uhv-config.json',
    });

    expect(result).toBeUndefined();
    expect(configBase.lastTenantConfigLoadError).toBe(
      'Tenant config URL is invalid or not allowed.',
    );
  });

  it('tracks HTTP failures when loading tenant config', async () => {
    const spHttpGet = jest.fn().mockResolvedValue({
      ok: false,
      status: 403,
    });
    const configBase = createConfigHarness(spHttpGet);

    const result = await (configBase as any).tryLoadTenantConfig(pageUrl, {
      tenantConfigUrl: '/sites/TestSite1/SiteAssets/uhv-config.json',
    });

    expect(result).toBeUndefined();
    expect(configBase.lastTenantConfigLoadError).toBe('Tenant config request failed (403).');
  });

  it('tracks invalid JSON payload types', async () => {
    const spHttpGet = jest.fn().mockResolvedValue({
      ok: true,
      json: jest.fn().mockResolvedValue('not-an-object'),
    });
    const configBase = createConfigHarness(spHttpGet);

    const result = await (configBase as any).tryLoadTenantConfig(pageUrl, {
      tenantConfigUrl: '/sites/TestSite1/SiteAssets/uhv-config.json',
    });

    expect(result).toBeUndefined();
    expect(configBase.lastTenantConfigLoadError).toBe(
      'Tenant config payload is invalid. Expected a JSON object.',
    );
  });

  it('clears previous tenant config load error after successful load', async () => {
    const spHttpGet = jest
      .fn()
      .mockResolvedValueOnce({
        ok: false,
        status: 503,
      })
      .mockResolvedValueOnce({
        ok: true,
        json: jest.fn().mockResolvedValue({
          dashboardId: 'ops',
        }),
      });
    const configBase = createConfigHarness(spHttpGet);

    const failedResult = await (configBase as any).tryLoadTenantConfig(pageUrl, {
      tenantConfigUrl: '/sites/TestSite1/SiteAssets/uhv-config.json',
    });
    const successResult = await (configBase as any).tryLoadTenantConfig(pageUrl, {
      tenantConfigUrl: '/sites/TestSite1/SiteAssets/uhv-config.json',
    });

    expect(failedResult).toBeUndefined();
    expect(successResult).toEqual({
      dashboardId: 'ops',
    });
    expect(configBase.lastTenantConfigLoadError).toBeUndefined();
  });
});

export {};
