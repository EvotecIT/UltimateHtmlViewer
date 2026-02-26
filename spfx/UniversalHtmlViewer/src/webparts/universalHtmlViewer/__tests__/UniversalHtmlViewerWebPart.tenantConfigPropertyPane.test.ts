/* eslint-disable @typescript-eslint/no-explicit-any, @typescript-eslint/no-var-requires */
jest.mock('@microsoft/sp-core-library', () => ({
  Version: {
    parse: () => ({}),
  },
}));
jest.mock('@microsoft/sp-http', () => ({
  SPHttpClient: {
    configurations: {
      v1: {},
    },
  },
}));
jest.mock('@microsoft/sp-lodash-subset', () => ({
  escape: (value: string): string => value,
}));
jest.mock('@microsoft/sp-property-pane', () => ({
  PropertyPaneDropdown: (targetProperty: string, properties: Record<string, unknown>) => ({
    type: 'dropdown',
    targetProperty,
    properties,
  }),
  PropertyPaneSlider: (targetProperty: string, properties: Record<string, unknown>) => ({
    type: 'slider',
    targetProperty,
    properties,
  }),
  PropertyPaneTextField: (targetProperty: string, properties: Record<string, unknown>) => ({
    type: 'text',
    targetProperty,
    properties,
  }),
  PropertyPaneToggle: (targetProperty: string, properties: Record<string, unknown>) => ({
    type: 'toggle',
    targetProperty,
    properties,
  }),
}));
jest.mock('@microsoft/sp-webpart-base', () => ({
  BaseClientSideWebPart: class {
    protected onPropertyPaneFieldChanged(): void {
      return;
    }
  },
}));

const {
  default: UniversalHtmlViewerWebPart,
}: {
  default: any;
} = require('../UniversalHtmlViewerWebPart');

function createWebPartHarness(): any {
  const webPart = Object.create(UniversalHtmlViewerWebPart.prototype) as any;
  webPart.properties = {
    configurationPreset: 'Custom',
    lockPresetSettings: false,
    tenantConfigUrl: '',
    tenantConfigMode: 'Merge',
  };
  webPart.context = {
    propertyPane: {
      refresh: jest.fn(),
    },
    pageContext: {
      web: {
        absoluteUrl: 'https://contoso.sharepoint.com/sites/TestSite1',
      },
    },
  };
  webPart.render = jest.fn();
  webPart.applyPreset = jest.fn();
  return webPart;
}

function findPropertyPaneField(
  configuration: { pages?: Array<{ groups?: Array<{ groupFields?: unknown[] }> }> },
  targetProperty: string,
): { properties?: Record<string, unknown> } | undefined {
  const pages = configuration.pages || [];
  for (const page of pages) {
    const groups = page.groups || [];
    for (const group of groups) {
      const fields = group.groupFields || [];
      for (const field of fields as Array<{ targetProperty?: string }>) {
        if (field.targetProperty === targetProperty) {
          return field as { properties?: Record<string, unknown> };
        }
      }
    }
  }

  return undefined;
}

describe('UniversalHtmlViewerWebPart tenant config property pane behavior', () => {
  it('trims tenantConfigUrl and refreshes property pane on tenantConfigUrl changes', () => {
    const webPart = createWebPartHarness();
    webPart.properties.tenantConfigUrl = '  /sites/TestSite1/SiteAssets/uhv-config.json  ';
    webPart.properties.tenantConfigMode = 'Override';

    webPart.onPropertyPaneFieldChanged('tenantConfigUrl', '', webPart.properties.tenantConfigUrl);

    expect(webPart.properties.tenantConfigUrl).toBe(
      '/sites/TestSite1/SiteAssets/uhv-config.json',
    );
    expect(webPart.properties.tenantConfigMode).toBe('Override');
    expect(webPart.context.propertyPane.refresh).toHaveBeenCalled();
    expect(webPart.render).toHaveBeenCalled();
  });

  it('resets tenantConfigMode to Merge when tenantConfigUrl becomes empty', () => {
    const webPart = createWebPartHarness();
    webPart.properties.tenantConfigUrl = '   ';
    webPart.properties.tenantConfigMode = 'Override';

    webPart.onPropertyPaneFieldChanged('tenantConfigUrl', '/old.json', '   ');

    expect(webPart.properties.tenantConfigUrl).toBe('');
    expect(webPart.properties.tenantConfigMode).toBe('Merge');
    expect(webPart.context.propertyPane.refresh).toHaveBeenCalled();
    expect(webPart.render).toHaveBeenCalled();
  });

  it('disables tenantConfigMode dropdown when tenantConfigUrl is whitespace-only', () => {
    const webPart = createWebPartHarness();
    webPart.properties.tenantConfigUrl = '   ';

    const configuration = webPart.getPropertyPaneConfiguration();
    const tenantConfigModeField = findPropertyPaneField(configuration, 'tenantConfigMode');

    expect(tenantConfigModeField).toBeDefined();
    expect(tenantConfigModeField?.properties?.disabled).toBe(true);
  });

  it('enables tenantConfigMode dropdown when tenantConfigUrl is configured', () => {
    const webPart = createWebPartHarness();
    webPart.properties.tenantConfigUrl = '/sites/TestSite1/SiteAssets/uhv-config.json';

    const configuration = webPart.getPropertyPaneConfiguration();
    const tenantConfigModeField = findPropertyPaneField(configuration, 'tenantConfigMode');

    expect(tenantConfigModeField).toBeDefined();
    expect(tenantConfigModeField?.properties?.disabled).toBe(false);
  });
});
