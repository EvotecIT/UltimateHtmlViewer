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

describe('UniversalHtmlViewerWebPartConfigBase mergeTenantConfig', () => {
  it('blocks prototype-polluting keys from tenant config', () => {
    const configBase = createConfigHarness();
    const props = {
      configurationPreset: 'Custom',
      lockPresetSettings: false,
      showLoadingIndicator: true,
    };
    const tenantConfig = JSON.parse(
      '{"__proto__":{"polluted":"yes"},"showLoadingIndicator":false}',
    );

    const merged = (configBase as any).mergeTenantConfig(props, tenantConfig, 'Override');

    expect(Object.getPrototypeOf(merged)).toBe(Object.prototype);
    expect((merged as any).polluted).toBeUndefined();
    expect(({} as any).polluted).toBeUndefined();
    expect(merged.showLoadingIndicator).toBe(false);
  });

  it('applies tenant numeric/boolean defaults in Merge mode when local values are defaults', () => {
    const configBase = createConfigHarness();
    const props = {
      configurationPreset: 'Custom',
      lockPresetSettings: false,
      fixedHeightPx: 800,
      refreshIntervalMinutes: 0,
      showLoadingIndicator: true,
      showChrome: true,
    };
    const tenantConfig = {
      fixedHeightPx: 1200,
      refreshIntervalMinutes: 10,
      showLoadingIndicator: false,
      showChrome: false,
    };

    const merged = (configBase as any).mergeTenantConfig(props, tenantConfig, 'Merge');

    expect(merged.fixedHeightPx).toBe(1200);
    expect(merged.refreshIntervalMinutes).toBe(10);
    expect(merged.showLoadingIndicator).toBe(false);
    expect(merged.showChrome).toBe(false);
  });

  it('keeps non-default local numeric/boolean values in Merge mode', () => {
    const configBase = createConfigHarness();
    const props = {
      configurationPreset: 'Custom',
      lockPresetSettings: false,
      fixedHeightPx: 950,
      refreshIntervalMinutes: 25,
      showLoadingIndicator: false,
      showChrome: false,
    };
    const tenantConfig = {
      fixedHeightPx: 1200,
      refreshIntervalMinutes: 10,
      showLoadingIndicator: true,
      showChrome: true,
    };

    const merged = (configBase as any).mergeTenantConfig(props, tenantConfig, 'Merge');

    expect(merged.fixedHeightPx).toBe(950);
    expect(merged.refreshIntervalMinutes).toBe(25);
    expect(merged.showLoadingIndicator).toBe(false);
    expect(merged.showChrome).toBe(false);
  });

  it('normalizes tenant dashboardList arrays and ignores unsupported keys', () => {
    const configBase = createConfigHarness();
    const props = {
      configurationPreset: 'Custom',
      lockPresetSettings: false,
      dashboardList: '',
    };
    const tenantConfig = {
      dashboardList: ['Sales|sales', 'Ops|ops'],
      unsupportedKey: 'ignored',
    };

    const merged = (configBase as any).mergeTenantConfig(props, tenantConfig, 'Override');

    expect(merged.dashboardList).toBe('Sales|sales,Ops|ops');
    expect((merged as any).unsupportedKey).toBeUndefined();
  });
});

export {};
