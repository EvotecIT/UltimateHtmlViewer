import { wireInlineIframeNavigation } from '../InlineNavigationHelper';
import { getInlineNavigationBridgeScript } from '../InlineNavigationBridgeScript';

describe('single-entry report dropdown navigation', () => {
  const reportUrl = 'https://contoso.sharepoint.com/sites/Test/SiteAssets/Reports/only.html';
  const hostPageUrl = 'https://contoso.sharepoint.com/sites/Test/SitePages/Viewer.aspx';

  it.each(['host handler', 'iframe bridge'])(
    '%s leaves the heading interactive and navigates only when its report link is clicked',
    (route) => {
      const iframe = document.createElement('iframe');
      document.body.appendChild(iframe);
      const frameDocument = iframe.contentDocument;
      const frameWindow = iframe.contentWindow as (Window & typeof globalThis) | null;
      if (!frameDocument || !frameWindow) {
        throw new Error('The test requires an iframe document and window.');
      }
      frameDocument.head.innerHTML =
        '<base href="https://contoso.sharepoint.com/sites/Test/SiteAssets/Reports/start.html">';
      frameDocument.body.innerHTML =
        '<nav><ul><li><span id="heading" class="dropdown-heading"><span id="icon">Reports</span></span>' +
        `<ul id="menu" hidden><li><a id="report" href="${reportUrl}"><span id="label">Only report</span></a></li></ul>` +
        '</li></ul></nav>';
      const heading = frameDocument.getElementById('heading') as HTMLElement;
      const menu = frameDocument.getElementById('menu') as HTMLElement;
      const onHeadingClick = jest.fn(() => { menu.hidden = !menu.hidden; });
      heading.addEventListener('click', onHeadingClick);
      const onNavigate = jest.fn();
      const parentPostMessage = jest.spyOn(window, 'postMessage').mockImplementation(() => undefined);
      let cleanup = (): void => undefined;
      try {
        if (route === 'host handler') {
          cleanup = wireInlineIframeNavigation({
            iframe,
            currentPageUrl: hostPageUrl,
            validationOptions: {
              securityMode: 'StrictTenant',
              currentPageUrl: hostPageUrl,
              allowedPathPrefixes: ['/sites/Test/SiteAssets/Reports/'],
              allowedFileExtensions: ['.html'],
            },
            cacheBusterParamName: 'v',
            onNavigate,
          });
        } else {
          frameWindow.eval(getInlineNavigationBridgeScript(
            'https://contoso.sharepoint.com/sites/Test/SiteAssets/Reports/start.html',
            ['.html'],
            ['/sites/Test/SiteAssets/Reports/'],
            hostPageUrl,
          ));
        }
        parentPostMessage.mockClear();
        const icon = frameDocument.getElementById('icon') as HTMLElement;
        for (const type of ['pointerdown', 'mousedown', 'click']) {
          const event = new frameWindow.MouseEvent(type, { bubbles: true, cancelable: true, button: 0 });
          icon.dispatchEvent(event);
          expect(event.defaultPrevented).toBe(false);
        }
        expect(onHeadingClick).toHaveBeenCalledTimes(1);
        expect(menu.hidden).toBe(false);
        expect(onNavigate).not.toHaveBeenCalled();
        expect(parentPostMessage.mock.calls.filter(([payload]) => payload?.type === 'uhv-inline-nav')).toHaveLength(0);

        const label = frameDocument.getElementById('label') as HTMLElement;
        const reportClick = new frameWindow.MouseEvent('click', { bubbles: true, cancelable: true, button: 0 });
        label.dispatchEvent(reportClick);
        expect(reportClick.defaultPrevented).toBe(true);
        if (route === 'host handler') {
          expect(onNavigate).toHaveBeenCalledTimes(1);
          expect(onNavigate).toHaveBeenCalledWith(reportUrl);
        } else {
          const navigationMessages = parentPostMessage.mock.calls.filter(([payload]) => payload?.type === 'uhv-inline-nav');
          expect(navigationMessages).toHaveLength(1);
          expect(navigationMessages[0][0]).toEqual(expect.objectContaining({ targetUrl: reportUrl }));
        }
      } finally {
        cleanup();
        parentPostMessage.mockRestore();
        iframe.remove();
      }
    },
  );
});
