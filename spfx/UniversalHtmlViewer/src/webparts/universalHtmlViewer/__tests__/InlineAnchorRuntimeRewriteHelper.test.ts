import { wireInlineAnchorRuntimeRewrite } from '../InlineAnchorRuntimeRewriteHelper';

describe('InlineAnchorRuntimeRewriteHelper', () => {
  it('rewrites FullCalendar-style anchors added after the report loads', async () => {
    const iframe = document.createElement('iframe');
    document.body.appendChild(iframe);
    const frameDocument = iframe.contentDocument as Document;
    frameDocument.head.innerHTML =
      '<base href="https://contoso.sharepoint.com/sites/Test/SiteAssets/">';

    const cleanup = wireInlineAnchorRuntimeRewrite({
      iframe,
      fallbackBaseUrl:
        'https://contoso.sharepoint.com/sites/Test/SiteAssets/Index.html',
      fallbackHostPageUrl:
        'https://contoso.sharepoint.com/sites/Test/SitePages/Dashboard.aspx',
      allowedFileExtensions: ['.html'],
      allowedPathPrefixes: ['/sites/Test/SiteAssets/'],
    });

    const calendarEvent = frameDocument.createElement('a');
    calendarEvent.className = 'fc-event';
    calendarEvent.href = 'GPOBroken_2021-04-05_230011.html';
    calendarEvent.setAttribute('href', 'GPOBroken_2021-04-05_230011.html');
    frameDocument.body.appendChild(calendarEvent);

    await new Promise((resolve) => window.setTimeout(resolve, 10));

    expect(calendarEvent.getAttribute('data-uhv-inline-href')).toBe(
      'https://contoso.sharepoint.com/sites/Test/SiteAssets/GPOBroken_2021-04-05_230011.html',
    );
    expect(calendarEvent.getAttribute('href')).toContain(
      '/sites/Test/SitePages/Dashboard.aspx?uhvPage=',
    );

    cleanup();
    iframe.remove();
  });

  it('removes download only from eligible html and leaves assets native', () => {
    const iframe = document.createElement('iframe');
    document.body.appendChild(iframe);
    const frameDocument = iframe.contentDocument as Document;
    frameDocument.head.innerHTML =
      '<base href="https://contoso.sharepoint.com/sites/Test/SiteAssets/">';
    frameDocument.body.innerHTML =
      '<a id="report" href="Report.html" download>Report</a>' +
      '<a id="csv" href="Report.csv" download>CSV</a>';

    const cleanup = wireInlineAnchorRuntimeRewrite({
      iframe,
      fallbackBaseUrl:
        'https://contoso.sharepoint.com/sites/Test/SiteAssets/Index.html',
      fallbackHostPageUrl:
        'https://contoso.sharepoint.com/sites/Test/SitePages/Dashboard.aspx',
      allowedFileExtensions: ['.html'],
      allowedPathPrefixes: ['/sites/Test/SiteAssets/'],
    });

    const report = frameDocument.getElementById('report') as HTMLAnchorElement;
    const csv = frameDocument.getElementById('csv') as HTMLAnchorElement;
    expect(report.hasAttribute('download')).toBe(false);
    expect(report.getAttribute('data-uhv-inline-rewritten')).toBe('1');
    expect(csv.hasAttribute('download')).toBe(true);
    expect(csv.getAttribute('href')).toBe('Report.csv');

    cleanup();
    iframe.remove();
  });

  it('uses a uniquely matching nested report directory for generated calendar filenames', () => {
    const iframe = document.createElement('iframe');
    document.body.appendChild(iframe);
    const frameDocument = iframe.contentDocument as Document;
    frameDocument.head.innerHTML =
      '<base href="https://contoso.sharepoint.com/sites/Test/SiteAssets/GPO_Broken.html">';
    frameDocument.body.innerHTML =
      '<iframe src="about:blank" data-uhv-inline-src="GPOzaurr/GPOBroken_2021-04-05_230011.html"></iframe>' +
      '<a id="event" class="fc-event" href="GPOBroken_2021-04-05_230011.html">GPO Broken</a>';

    const cleanup = wireInlineAnchorRuntimeRewrite({
      iframe,
      fallbackBaseUrl:
        'https://contoso.sharepoint.com/sites/Test/SiteAssets/GPO_Broken.html',
      fallbackHostPageUrl:
        'https://contoso.sharepoint.com/sites/Test/SitePages/Dashboard.aspx',
      allowedFileExtensions: ['.html'],
      allowedPathPrefixes: ['/sites/Test/SiteAssets/'],
    });

    const event = frameDocument.getElementById('event') as HTMLAnchorElement;
    expect(event.getAttribute('data-uhv-inline-href')).toBe(
      'https://contoso.sharepoint.com/sites/Test/SiteAssets/GPOzaurr/GPOBroken_2021-04-05_230011.html',
    );

    cleanup();
    iframe.remove();
  });
});
