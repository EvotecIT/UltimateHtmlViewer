import { loadSharePointReportBrowserItems } from '../SharePointReportBrowserHelper';

describe('SharePoint report browser request budget', () => {
  it('bounds recursive Files view SharePoint requests', async () => {
    let folderNumber = 0;
    const spHttpClient = {
      get: jest.fn((url: string) => {
        const isFolderRequest = url.includes('/Folders?');
        const nextFolderPath = `/sites/TestSite1/SiteAssets/Reports/Folder${folderNumber++}`;
        return Promise.resolve({
          ok: true,
          json: () =>
            Promise.resolve({
              value: isFolderRequest
                ? [{ Name: `Folder${folderNumber}`, ServerRelativeUrl: nextFolderPath }]
                : [],
            }),
        });
      }),
    };

    await expect(
      loadSharePointReportBrowserItems({
        spHttpClient: spHttpClient as never,
        webAbsoluteUrl: 'https://contoso.sharepoint.com/sites/TestSite1',
        rootPath: '/sites/TestSite1/SiteAssets/Reports',
        allowedExtensions: ['.html'],
        view: 'Files',
        maxItems: 100,
        maxRequests: 3,
      }),
    ).rejects.toThrow('SharePoint report browser request limit reached.');
    expect(spHttpClient.get).toHaveBeenCalledTimes(3);
  });

  it('retains a shared folder-view budget until both parallel loads settle', async () => {
    let resolveFilePage: ((value: unknown) => void) | undefined;
    const filePage = new Promise((resolve) => {
      resolveFilePage = resolve;
    });
    const spHttpClient = {
      get: jest.fn((url: string) => {
        if (url.includes('/Folders?')) {
          return Promise.resolve({
            ok: false,
            status: 500,
            statusText: 'Server Error',
          });
        }

        return Promise.resolve({
          ok: true,
          json: () => filePage,
        });
      }),
    };

    const loadResult = loadSharePointReportBrowserItems({
      spHttpClient: spHttpClient as never,
      webAbsoluteUrl: 'https://contoso.sharepoint.com/sites/TestSite1',
      rootPath: '/sites/TestSite1/SiteAssets/Reports',
      allowedExtensions: ['.html'],
      view: 'Folders',
      maxItems: 100,
      maxRequests: 2,
    }).catch((error: unknown) => error);

    await Promise.resolve();
    await Promise.resolve();
    resolveFilePage?.({
      value: [],
      '@odata.nextLink': 'https://contoso.sharepoint.com/sites/TestSite1/_api/next-files',
    });

    const error = await loadResult;
    await new Promise((resolve) => setTimeout(resolve, 0));

    expect(error).toBeInstanceOf(Error);
    expect((error as Error).message).toContain('500 Server Error');
    expect(spHttpClient.get).toHaveBeenCalledTimes(2);
    expect(spHttpClient.get).not.toHaveBeenCalledWith(
      'https://contoso.sharepoint.com/sites/TestSite1/_api/next-files',
      expect.anything(),
      expect.anything(),
    );
  });
});
