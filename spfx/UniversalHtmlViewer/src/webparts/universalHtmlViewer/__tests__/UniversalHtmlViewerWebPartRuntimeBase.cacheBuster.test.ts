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
  UniversalHtmlViewerWebPartRuntimeBase,
}: {
  UniversalHtmlViewerWebPartRuntimeBase: any;
} = require('../UniversalHtmlViewerWebPartRuntimeBase');

describe('UniversalHtmlViewerWebPartRuntimeBase FileLastModified cache buster', () => {
  it('passes a decoded server-relative path to the ResourcePath metadata API', async () => {
    const get = jest.fn().mockResolvedValue({
      ok: true,
      json: async () => ({ d: { ETag: '"etag-1"' } }),
    });
    const runtime = Object.create(UniversalHtmlViewerWebPartRuntimeBase.prototype) as any;
    runtime.context = {
      pageContext: {
        web: {
          absoluteUrl: 'https://contoso.sharepoint.com/sites/Test',
        },
      },
      spHttpClient: { get },
    };

    await runtime.resolveUrlWithCacheBuster(
      '/sites/Test/Shared%20Documents/report%23one.html',
      'FileLastModified',
      'v',
      'https://contoso.sharepoint.com/sites/Test/SitePages/Viewer.aspx',
    );

    const metadataUrl = get.mock.calls[0][0] as string;
    expect(metadataUrl).toContain(
      "%2Fsites%2FTest%2FShared%20Documents%2Freport%23one.html",
    );
    expect(metadataUrl).not.toContain('%2520');
    expect(metadataUrl).not.toContain('%2523');
  });
});

export {};
