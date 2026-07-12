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
});
