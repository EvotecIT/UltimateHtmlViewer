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

  it('treats tenant root as containing server-relative descendants', () => {
    expect(isPathInsideRoot('/sites/TestSite1/SiteAssets/Reports/index.html', '/')).toBe(
      true,
    );
    expect(isPathInsideRoot('/', '/')).toBe(true);
    expect(isPathInsideRoot('', '/')).toBe(false);
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

  it('stops paging when max item budget is reached', async () => {
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

        return Promise.resolve({
          ok: true,
          json: () =>
            Promise.resolve({
              value: url.includes('/Folders?')
                ? []
                : [
                    {
                      Name: 'First.html',
                      ServerRelativeUrl:
                        '/sites/TestSite1/SiteAssets/Reports/First.html',
                    },
                  ],
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
      view: 'Files',
      maxItems: 1,
    });

    expect(items.map((item) => item.name)).toEqual(['First.html']);
    expect(
      spHttpClient.get.mock.calls.some(
        ([url]) => url === 'https://contoso.sharepoint.com/sites/TestSite1/_api/next-files',
      ),
    ).toBe(false);
    expect(spHttpClient.get.mock.calls[0][0]).toContain('$top=5000');
  });

  it('continues file paging until enough allowed reports are found', async () => {
    const nextFilesLink = 'https://contoso.sharepoint.com/sites/TestSite1/_api/next-files';
    const spHttpClient = {
      get: jest.fn((url: string) => {
        if (url === nextFilesLink) {
          return Promise.resolve({
            ok: true,
            json: () =>
              Promise.resolve({
                value: [
                  {
                    Name: 'Report.html',
                    ServerRelativeUrl: '/sites/TestSite1/SiteAssets/Reports/Report.html',
                  },
                ],
              }),
          });
        }

        return Promise.resolve({
          ok: true,
          json: () =>
            Promise.resolve({
              value: url.includes('/Files?')
                ? [
                    {
                      Name: 'Data.csv',
                      ServerRelativeUrl: '/sites/TestSite1/SiteAssets/Reports/Data.csv',
                    },
                  ]
                : [],
              '@odata.nextLink': url.includes('/Files?') ? nextFilesLink : undefined,
            }),
        });
      }),
    };

    const items = await loadSharePointReportBrowserItems({
      spHttpClient: spHttpClient as never,
      webAbsoluteUrl: 'https://contoso.sharepoint.com/sites/TestSite1',
      rootPath: '/sites/TestSite1/SiteAssets/Reports',
      allowedExtensions: ['.html'],
      view: 'Files',
      maxItems: 1,
    });

    expect(items.map((item) => item.name)).toEqual(['Report.html']);
    expect(spHttpClient.get.mock.calls.some(([url]) => url === nextFilesLink)).toBe(true);
  });

  it('continues folder-view file paging after non-report files', async () => {
    const nextFilesLink = 'https://contoso.sharepoint.com/sites/TestSite1/_api/next-files';
    const spHttpClient = {
      get: jest.fn((url: string) => {
        if (url === nextFilesLink) {
          return Promise.resolve({
            ok: true,
            json: () =>
              Promise.resolve({
                value: [
                  {
                    Name: 'Report.html',
                    ServerRelativeUrl: '/sites/TestSite1/SiteAssets/Reports/Report.html',
                  },
                ],
              }),
          });
        }

        return Promise.resolve({
          ok: true,
          json: () =>
            Promise.resolve({
              value: url.includes('/Files?')
                ? [
                    {
                      Name: 'Data.csv',
                      ServerRelativeUrl: '/sites/TestSite1/SiteAssets/Reports/Data.csv',
                    },
                  ]
                : [],
              '@odata.nextLink': url.includes('/Files?') ? nextFilesLink : undefined,
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
      maxItems: 1,
    });

    expect(items.map((item) => item.name)).toEqual(['Report.html']);
    expect(spHttpClient.get.mock.calls.some(([url]) => url === nextFilesLink)).toBe(true);
  });

  it('preserves file rows when folder view reaches the item cap', async () => {
    const spHttpClient = {
      get: jest.fn((url: string) => {
        const value = url.includes('/Folders?')
          ? [
              {
                Name: 'Folder 1',
                ServerRelativeUrl: '/sites/TestSite1/SiteAssets/Reports/Folder1',
              },
              {
                Name: 'Folder 2',
                ServerRelativeUrl: '/sites/TestSite1/SiteAssets/Reports/Folder2',
              },
            ]
          : [
              {
                Name: 'Index.html',
                ServerRelativeUrl: '/sites/TestSite1/SiteAssets/Reports/Index.html',
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
      maxItems: 2,
    });

    expect(items.map((item) => item.kind)).toEqual(['Folder', 'File']);
    expect(items.map((item) => item.name)).toContain('Index.html');
  });

  it('does not cap recursive folder paging when searching for files', async () => {
    const folderNextLink = 'https://contoso.sharepoint.com/sites/TestSite1/_api/next-folders';
    const spHttpClient = {
      get: jest.fn((url: string) => {
        if (url === folderNextLink) {
          return Promise.resolve({
            ok: true,
            json: () =>
              Promise.resolve({
                value: [
                  {
                    Name: 'Folder 2',
                    ServerRelativeUrl: '/sites/TestSite1/SiteAssets/Reports/Folder2',
                  },
                ],
              }),
          });
        }

        if (url.includes('/Folders?')) {
          return Promise.resolve({
            ok: true,
            json: () =>
              Promise.resolve({
                value: [
                  {
                    Name: 'Folder 1',
                    ServerRelativeUrl: '/sites/TestSite1/SiteAssets/Reports/Folder1',
                  },
                ],
                '@odata.nextLink': folderNextLink,
              }),
          });
        }

        return Promise.resolve({
          ok: true,
          json: () =>
            Promise.resolve({
              value: url.includes('/Files?') && url.includes('Folder1')
                ? [
                    {
                      Name: 'Nested.html',
                      ServerRelativeUrl:
                        '/sites/TestSite1/SiteAssets/Reports/Folder1/Nested.html',
                    },
                  ]
                : [],
            }),
        });
      }),
    };

    const items = await loadSharePointReportBrowserItems({
      spHttpClient: spHttpClient as never,
      webAbsoluteUrl: 'https://contoso.sharepoint.com/sites/TestSite1',
      rootPath: '/sites/TestSite1/SiteAssets/Reports',
      allowedExtensions: ['.html'],
      view: 'Files',
      maxItems: 1,
    });

    expect(items.map((item) => item.name)).toEqual(['Nested.html']);
    expect(spHttpClient.get.mock.calls.some(([url]) => url === folderNextLink)).toBe(true);
  });

  it('traverses deeply nested report folders without a fixed depth cutoff', async () => {
    const rootPath = '/sites/TestSite1/SiteAssets/Reports';
    const finalFolderPath = `${rootPath}/${Array.from({ length: 14 }, (_, index) =>
      `Level${index + 1}`,
    ).join('/')}`;
    const spHttpClient = {
      get: jest.fn((url: string) => {
        const folderMatch = /@p1='([^']+)'/.exec(url);
        const folderPath = folderMatch ? decodeURIComponent(folderMatch[1]) : rootPath;
        if (url.includes('/Files?')) {
          return Promise.resolve({
            ok: true,
            json: () =>
              Promise.resolve({
                value:
                  folderPath === finalFolderPath
                    ? [
                        {
                          Name: 'Deep.html',
                          ServerRelativeUrl: `${finalFolderPath}/Deep.html`,
                        },
                      ]
                    : [],
              }),
          });
        }

        const currentDepth =
          folderPath === rootPath
            ? 0
            : folderPath.substring(rootPath.length + 1).split('/').length;
        const nextDepth = currentDepth + 1;
        return Promise.resolve({
          ok: true,
          json: () =>
            Promise.resolve({
              value:
                nextDepth <= 14
                  ? [
                      {
                        Name: `Level${nextDepth}`,
                        ServerRelativeUrl: `${folderPath}/Level${nextDepth}`,
                      },
                    ]
                  : [],
            }),
        });
      }),
    };

    const items = await loadSharePointReportBrowserItems({
      spHttpClient: spHttpClient as never,
      webAbsoluteUrl: 'https://contoso.sharepoint.com/sites/TestSite1',
      rootPath,
      allowedExtensions: ['.html'],
      view: 'Files',
      maxItems: 1,
    });

    expect(items.map((item) => item.name)).toEqual(['Deep.html']);
  });

  it('does not recurse forever when SharePoint returns a folder cycle', async () => {
    const rootPath = '/sites/TestSite1/SiteAssets/Reports';
    const childPath = `${rootPath}/Loop`;
    const spHttpClient = {
      get: jest.fn((url: string) => {
        const folderMatch = /@p1='([^']+)'/.exec(url);
        const folderPath = folderMatch ? decodeURIComponent(folderMatch[1]) : rootPath;
        if (url.includes('/Files?')) {
          return Promise.resolve({
            ok: true,
            json: () => Promise.resolve({ value: [] }),
          });
        }

        return Promise.resolve({
          ok: true,
          json: () =>
            Promise.resolve({
              value: [
                {
                  Name: 'Loop',
                  ServerRelativeUrl: folderPath === rootPath ? childPath : rootPath,
                },
              ],
            }),
        });
      }),
    };

    const items = await loadSharePointReportBrowserItems({
      spHttpClient: spHttpClient as never,
      webAbsoluteUrl: 'https://contoso.sharepoint.com/sites/TestSite1',
      rootPath,
      allowedExtensions: ['.html'],
      view: 'Files',
      maxItems: 10,
    });

    expect(items).toEqual([]);
    expect(spHttpClient.get).toHaveBeenCalledTimes(4);
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

  it('returns navigation-safe encoded server-relative URLs', async () => {
    const spHttpClient = {
      get: jest.fn((url: string) => {
        const value = url.includes('/Folders?')
          ? []
          : [
              {
                Name: 'Quarter #1.html',
                ServerRelativeUrl:
                  '/sites/TestSite1/SiteAssets/Reports/Quarter%20%231.html',
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

    expect(items[0].serverRelativeUrl).toBe(
      '/sites/TestSite1/SiteAssets/Reports/Quarter%20%231.html',
    );
    expect(items[0].relativePath).toBe('Quarter #1.html');
  });
});
