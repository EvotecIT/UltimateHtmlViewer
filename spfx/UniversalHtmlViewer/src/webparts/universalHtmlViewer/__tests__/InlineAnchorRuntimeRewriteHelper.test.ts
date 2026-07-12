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
    const rewritten = waitForAttribute(
      calendarEvent,
      'data-uhv-inline-href',
      'https://contoso.sharepoint.com/sites/Test/SiteAssets/GPOBroken_2021-04-05_230011.html',
    );
    frameDocument.body.appendChild(calendarEvent);

    await rewritten;

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

  it('replaces stale UHV metadata when a calendar reuses an anchor for another report', async () => {
    const iframe = document.createElement('iframe');
    document.body.appendChild(iframe);
    const frameDocument = iframe.contentDocument as Document;
    frameDocument.head.innerHTML =
      '<base href="https://contoso.sharepoint.com/sites/Test/SiteAssets/">';
    frameDocument.body.innerHTML = '<a id="event" href="First.html">First</a>';

    const cleanup = wireInlineAnchorRuntimeRewrite({
      iframe,
      fallbackBaseUrl:
        'https://contoso.sharepoint.com/sites/Test/SiteAssets/Index.html',
      fallbackHostPageUrl:
        'https://contoso.sharepoint.com/sites/Test/SitePages/Dashboard.aspx',
      allowedFileExtensions: ['.html'],
      allowedPathPrefixes: ['/sites/Test/SiteAssets/'],
    });

    const event = frameDocument.getElementById('event') as HTMLAnchorElement;
    expect(event.getAttribute('data-uhv-inline-href')).toContain('/First.html');

    const rewritten = waitForAttribute(
      event,
      'data-uhv-inline-href',
      'https://contoso.sharepoint.com/sites/Test/SiteAssets/Second.html',
    );
    event.setAttribute('href', 'Second.html');
    await rewritten;

    expect(event.getAttribute('data-uhv-inline-href')).toBe(
      'https://contoso.sharepoint.com/sites/Test/SiteAssets/Second.html',
    );
    expect(event.getAttribute('href')).toContain('Second.html');
    expect(event.getAttribute('href')).not.toContain('First.html');

    cleanup();
    iframe.remove();
  });

  it('uses live sibling viewer state for anchors generated after host navigation', async () => {
    const originalUrl = `${window.location.pathname}${window.location.search}${window.location.hash}`;
    window.history.replaceState(
      null,
      '',
      '/sites/Test/SitePages/Dashboard.aspx?viewerTwoPage=old-report',
    );
    const iframe = document.createElement('iframe');
    document.body.appendChild(iframe);
    const frameDocument = iframe.contentDocument as Document;
    frameDocument.head.innerHTML =
      `<base href="${window.location.origin}/sites/Test/SiteAssets/">`;

    const cleanup = wireInlineAnchorRuntimeRewrite({
      iframe,
      fallbackBaseUrl: `${window.location.origin}/sites/Test/SiteAssets/Index.html`,
      fallbackHostPageUrl: window.location.href,
      allowedFileExtensions: ['.html'],
      allowedPathPrefixes: ['/sites/Test/SiteAssets/'],
    });

    window.history.pushState(
      null,
      '',
      '/sites/Test/SitePages/Dashboard.aspx?viewerTwoPage=current-report',
    );
    const calendarEvent = frameDocument.createElement('a');
    calendarEvent.setAttribute('href', 'Generated.html');
    const rewritten = waitForAttribute(
      calendarEvent,
      'data-uhv-inline-href',
      `${window.location.origin}/sites/Test/SiteAssets/Generated.html`,
    );
    frameDocument.body.appendChild(calendarEvent);

    await rewritten;

    const rewrittenUrl = new URL(calendarEvent.getAttribute('href') || '', window.location.href);
    expect(rewrittenUrl.searchParams.get('viewerTwoPage')).toBe('current-report');
    expect(rewrittenUrl.searchParams.get('viewerTwoPage')).not.toBe('old-report');

    cleanup();
    iframe.remove();
    window.history.replaceState(null, '', originalUrl || '/');
  });
});

function waitForAttribute(
  element: Element,
  attributeName: string,
  expectedValue: string,
): Promise<void> {
  if (element.getAttribute(attributeName) === expectedValue) {
    return Promise.resolve();
  }

  return new Promise((resolve, reject) => {
    let timeoutId = 0;
    const observer = new MutationObserver(() => {
      if (element.getAttribute(attributeName) !== expectedValue) {
        return;
      }

      window.clearTimeout(timeoutId);
      observer.disconnect();
      resolve();
    });
    timeoutId = window.setTimeout(() => {
      observer.disconnect();
      reject(
        new Error(
          `Timed out waiting for ${attributeName} to become ${expectedValue}.`,
        ),
      );
    }, 1000);
    observer.observe(element, {
      attributeFilter: [attributeName],
      attributes: true,
    });
  });
}
