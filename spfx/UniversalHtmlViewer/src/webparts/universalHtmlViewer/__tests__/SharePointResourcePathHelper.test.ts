import {
  buildSharePointFileByPathApiUrl,
  buildSharePointFolderByPathApiUrl,
  decodeSharePointUrlPath,
} from '../SharePointResourcePathHelper';

describe('SharePointResourcePathHelper', () => {
  it('builds decoded-path file API aliases without treating percent or hash as URL syntax', () => {
    const url = buildSharePointFileByPathApiUrl(
      'https://contoso.sharepoint.com/sites/Test',
      "/sites/Test/Shared Documents/50% #1 O'Brien.html",
      '/$value',
    );

    expect(url).toContain('GetFileByServerRelativePath(decodedUrl=@p1)/$value');
    expect(url).toContain("50%25%20%231%20O''Brien.html");
  });

  it('builds decoded-path folder child collection queries', () => {
    const url = buildSharePointFolderByPathApiUrl(
      'https://contoso.sharepoint.com/sites/Test',
      '/sites/Test/Shared Documents/Reports',
      '/Files',
      '$select=Name',
    );

    expect(url).toContain('GetFolderByServerRelativePath(decodedUrl=@p1)/Files');
    expect(url).toContain("?@p1='%2Fsites%2FTest%2FShared%20Documents%2FReports'");
    expect(url).toContain('&$select=Name');
  });

  it('decodes one URL path layer without failing on a literal percent', () => {
    expect(
      decodeSharePointUrlPath('/sites/Test/Shared%20Documents/report%23one.html'),
    ).toBe('/sites/Test/Shared Documents/report#one.html');
    expect(decodeSharePointUrlPath('/sites/Test/50% complete.html')).toBe(
      '/sites/Test/50% complete.html',
    );
  });
});
