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
        absoluteUrl: 'https://contoso.sharepoint.com/sites/TestSite2',
        serverRelativeUrl: '/sites/TestSite2',
      },
    },
  };

  return configBase;
}

describe('UniversalHtmlViewerWebPartConfigBase buildUrlValidationOptions', () => {
  const pageUrl = 'https://contoso.sharepoint.com/sites/TestSite2/SitePages/Dashboard.aspx';

  it('appends inferred source directory prefix for SharePoint inline mode when configured prefix is stale', () => {
    const configBase = createConfigHarness();

    const options = (configBase as any).buildUrlValidationOptions(pageUrl, {
      htmlSourceMode: 'FullUrl',
      fullUrl: 'https://contoso.sharepoint.com/sites/TestSite2/SiteAssets/Index.html',
      contentDeliveryMode: 'SharePointFileContent',
      securityMode: 'StrictTenant',
      allowHttp: false,
      allowedHosts: '',
      allowedPathPrefixes: '/sites/Reports/Dashboards/',
      allowedFileExtensions: '.html,.htm,.aspx',
    });

    expect(options.allowedPathPrefixes).toContain('/sites/Reports/Dashboards/');
    expect(options.allowedPathPrefixes).toContain('/sites/TestSite2/SiteAssets/');
  });

  it('does not append inferred source directory prefix for DirectUrl mode', () => {
    const configBase = createConfigHarness();

    const options = (configBase as any).buildUrlValidationOptions(pageUrl, {
      htmlSourceMode: 'FullUrl',
      fullUrl: 'https://contoso.sharepoint.com/sites/TestSite2/SiteAssets/Index.html',
      contentDeliveryMode: 'DirectUrl',
      securityMode: 'StrictTenant',
      allowHttp: false,
      allowedHosts: '',
      allowedPathPrefixes: '/sites/Reports/Dashboards/',
      allowedFileExtensions: '.html,.htm,.aspx',
    });

    expect(options.allowedPathPrefixes).toEqual(['/sites/Reports/Dashboards/']);
  });

  it('does not append inferred source directory prefix for cross-host full URL', () => {
    const configBase = createConfigHarness();

    const options = (configBase as any).buildUrlValidationOptions(pageUrl, {
      htmlSourceMode: 'FullUrl',
      fullUrl: 'https://fabrikam.sharepoint.com/sites/TestSite2/SiteAssets/Index.html',
      contentDeliveryMode: 'SharePointFileContent',
      securityMode: 'StrictTenant',
      allowHttp: false,
      allowedHosts: '',
      allowedPathPrefixes: '/sites/Reports/Dashboards/',
      allowedFileExtensions: '.html,.htm,.aspx',
    });

    expect(options.allowedPathPrefixes).toEqual(['/sites/Reports/Dashboards/']);
  });
});

export {};
