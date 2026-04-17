import {
  getReportBrowserParentFolderPath,
  isPathInsideRoot,
  loadSharePointReportBrowserItems,
  normalizeSharePointReportBrowserRootPath,
} from '../SharePointReportBrowserHelper';

describe('SharePointReportBrowserHelper', () => {
  it('normalizes web-relative browser roots against the current web', () => {
    expect(
      normalizeSharePointReportBrowserRootPath(
        'SiteAssets/Reports/',
        'https://contoso.sharepoint.com/sites/TestSite1',
      ),
    ).toBe('/sites/TestSite1/SiteAssets/Reports');
  });

  it('keeps server-relative browser roots', () => {
    expect(
      normalizeSharePointReportBrowserRootPath(
        '/sites/TestSite1/SiteAssets/Reports/',
        'https://contoso.sharepoint.com/sites/TestSite1',
      ),
    ).toBe('/sites/TestSite1/SiteAssets/Reports');
  });

  it('detects root-contained paths case-insensitively', () => {
    expect(
      isPathInsideRoot(
        '/sites/TestSite1/SiteAssets/Reports/Global/index.html',
        '/sites/testsite1/siteassets/reports',
      ),
    ).toBe(true);
    expect(
      isPathInsideRoot(
        '/sites/TestSite1/SiteAssets/Other/index.html',
        '/sites/TestSite1/SiteAssets/Reports',
      ),
    ).toBe(false);
  });

  it('does not walk above the configured browser root', () => {
    expect(
      getReportBrowserParentFolderPath(
        '/sites/TestSite1/SiteAssets/Reports',
        '/sites/TestSite1/SiteAssets/Reports',
      ),
    ).toBe('/sites/TestSite1/SiteAssets/Reports');
    expect(
      getReportBrowserParentFolderPath(
        '/sites/TestSite1/SiteAssets/Reports',
        '/sites/TestSite1/SiteAssets/Reports/Global',
      ),
    ).toBe('/sites/TestSite1/SiteAssets/Reports');
  });

  it('loads folder view items with security-trimmed API results', async () => {
    const spHttpClient = {
      get: jest.fn((url: string) => {
        const value = url.includes('/Folders?')
          ? [
              {
                Name: 'Global',
                ServerRelativeUrl: '/sites/TestSite1/SiteAssets/Reports/Global',
              },
              {
                Name: 'Forms',
                ServerRelativeUrl: '/sites/TestSite1/SiteAssets/Reports/Forms',
              },
            ]
          : [
              {
                Name: 'Index.html',
                ServerRelativeUrl: '/sites/TestSite1/SiteAssets/Reports/Index.html',
              },
              {
                Name: 'Readme.txt',
                ServerRelativeUrl: '/sites/TestSite1/SiteAssets/Reports/Readme.txt',
              },
            ];
        return Promise.resolve({
          ok: true,
          json: () => Promise.resolve({ value }),
        });
      }),
    };

    const items = await loadSharePointReportBrowserItems({
      spHttpClient: spHttpClient as never,
      webAbsoluteUrl: 'https://contoso.sharepoint.com/sites/TestSite1',
      rootPath: '/sites/TestSite1/SiteAssets/Reports',
      allowedExtensions: ['.html'],
      view: 'Folders',
      maxItems: 100,
    });

    expect(items).toEqual([
      {
        kind: 'Folder',
        name: 'Global',
        relativePath: 'Global',
        serverRelativeUrl: '/sites/TestSite1/SiteAssets/Reports/Global',
      },
      {
        kind: 'File',
        name: 'Index.html',
        relativePath: 'Index.html',
        serverRelativeUrl: '/sites/TestSite1/SiteAssets/Reports/Index.html',
        timeLastModified: undefined,
      },
    ]);
    expect(spHttpClient.get).toHaveBeenCalledWith(
      expect.any(String),
      undefined,
      expect.objectContaining({
        headers: {
          Accept: 'application/json;odata=nometadata',
          'OData-Version': '',
        },
      }),
    );
  });

  it('follows SharePoint OData nextLink pages', async () => {
    const spHttpClient = {
      get: jest.fn((url: string) => {
        if (url === 'https://contoso.sharepoint.com/sites/TestSite1/_api/next-files') {
          return Promise.resolve({
            ok: true,
            json: () =>
              Promise.resolve({
                value: [
                  {
                    Name: 'Second.html',
                    ServerRelativeUrl:
                      '/sites/TestSite1/SiteAssets/Reports/Second.html',
                  },
                ],
              }),
          });
        }

        const value = url.includes('/Folders?')
          ? []
          : [
              {
                Name: 'First.html',
                ServerRelativeUrl: '/sites/TestSite1/SiteAssets/Reports/First.html',
              },
            ];
        return Promise.resolve({
          ok: true,
          json: () =>
            Promise.resolve({
              value,
              '@odata.nextLink': url.includes('/Files?')
                ? 'https://contoso.sharepoint.com/sites/TestSite1/_api/next-files'
                : undefined,
            }),
        });
      }),
    };

    const items = await loadSharePointReportBrowserItems({
      spHttpClient: spHttpClient as never,
      webAbsoluteUrl: 'https://contoso.sharepoint.com/sites/TestSite1',
      rootPath: '/sites/TestSite1/SiteAssets/Reports',
      allowedExtensions: ['.html'],
      view: 'Folders',
      maxItems: 100,
    });

    expect(items.map((item) => item.name)).toEqual(['First.html', 'Second.html']);
    expect(spHttpClient.get).toHaveBeenCalledWith(
      'https://contoso.sharepoint.com/sites/TestSite1/_api/next-files',
      undefined,
      expect.any(Object),
    );
  });

  it('does not double-encode URL-derived folder paths', async () => {
    const spHttpClient = {
      get: jest.fn((_url: string) =>
        Promise.resolve({
          ok: true,
          json: () => Promise.resolve({ value: [] }),
        }),
      ),
    };

    await loadSharePointReportBrowserItems({
      spHttpClient: spHttpClient as never,
      webAbsoluteUrl: 'https://contoso.sharepoint.com/sites/TestSite1',
      rootPath: 'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/My%20Reports',
      allowedExtensions: ['.html'],
      view: 'Folders',
      maxItems: 100,
    });

    const firstUrl = spHttpClient.get.mock.calls[0][0] as string;
    expect(firstUrl).toContain('My%20Reports');
    expect(firstUrl).not.toContain('My%2520Reports');
  });
});
