import {
  prepareInlineHtmlForSrcDoc,
} from '../InlineHtmlTransformHelper';

describe('inline CSP source allowlists', () => {
  it('adds configured external hosts to the generated inline CSP', () => {
    const inputHtml = '<html><head><title>Report</title></head><body><h1>Dashboard</h1></body></html>';
    const result = prepareInlineHtmlForSrcDoc(
      inputHtml,
      '/sites/TestSite2/SiteAssets/GPOzaurr/',
      'https://contoso.sharepoint.com/sites/TestSite2/SitePages/Dashboard.aspx',
      {
        additionalScriptSrcHosts: [
          'cdn.jsdelivr.net',
          'https://cdn.datatables.net/1.13.5/js/jquery.dataTables.min.js',
        ],
        additionalStyleSrcHosts: ['fonts.googleapis.com', 'cdn.datatables.net'],
        additionalImageSrcHosts: ['upload.wikimedia.org'],
      },
    );

    expect(result).toContain(
      'script-src &#39;self&#39; https://contoso.sharepoint.com blob: &#39;unsafe-inline&#39; &#39;unsafe-eval&#39; https://cdn.jsdelivr.net https://cdn.datatables.net',
    );
    expect(result).toContain(
      'style-src &#39;self&#39; https://contoso.sharepoint.com data: &#39;unsafe-inline&#39; https://fonts.googleapis.com https://cdn.datatables.net',
    );
    expect(result).toContain(
      'img-src &#39;self&#39; https://contoso.sharepoint.com data: blob: https://upload.wikimedia.org',
    );
  });
});
