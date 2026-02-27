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
jest.mock('@microsoft/sp-lodash-subset', () => ({
  escape: (value: string): string => value,
}));
jest.mock('../UniversalHtmlViewerWebPart.module.scss', () => ({}));

const {
  UniversalHtmlViewerWebPartRuntimeBase,
}: {
  UniversalHtmlViewerWebPartRuntimeBase: any;
} = require('../UniversalHtmlViewerWebPartRuntimeBase');

function createRuntimeHarness(): any {
  const runtime = Object.create(
    UniversalHtmlViewerWebPartRuntimeBase.prototype,
  ) as any;

  runtime.properties = {
    configurationPreset: 'Custom',
    contentDeliveryMode: 'DirectUrl',
    securityMode: 'StrictTenant',
  };
  runtime.getContentDeliveryMode = jest.fn().mockReturnValue('DirectUrl');
  runtime.parseHosts = jest.fn().mockReturnValue([]);
  runtime.parsePathPrefixes = jest.fn().mockReturnValue([]);
  runtime.parseFileExtensions = jest.fn().mockReturnValue([]);

  return runtime;
}

describe('UniversalHtmlViewerWebPartRuntimeBase buildDiagnosticsData', () => {
  it('includes tenant config load error detail in diagnostics payload', () => {
    const runtime = createRuntimeHarness();
    runtime.lastTenantConfigLoadError = 'Tenant config request failed (503).';

    const data = runtime.buildDiagnosticsData({
      resolvedUrl: '/sites/TestSite1/SiteAssets/Reports/index.html',
    });

    expect(data.tenantConfigLoadError).toBe('Tenant config request failed (503).');
  });
});

export {};
