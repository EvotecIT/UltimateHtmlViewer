import { resolveInlineNavigationTarget } from '../InlineNavigationHelper';
import { UrlValidationOptions } from '../UrlHelper';

describe('InlineNavigationHelper', () => {
  const validationOptions: UrlValidationOptions = {
    securityMode: 'StrictTenant',
    currentPageUrl: 'https://contoso.sharepoint.com/sites/TestSite1/SitePages/Home.aspx',
    allowHttp: false,
    allowedPathPrefixes: ['/sites/testsite1/siteassets/'],
    allowedFileExtensions: ['.html', '.htm', '.aspx'],
  };

  it('resolves relative html links inside allowed path and strips cache param', () => {
    const anchor = document.createElement('a');
    anchor.href = 'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/index.html?v=123';
    anchor.setAttribute(
      'href',
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/index.html?v=123',
    );
    const clickEvent = new MouseEvent('click', { bubbles: true, button: 0 });
    Object.defineProperty(clickEvent, 'target', {
      value: anchor,
      configurable: true,
    });

    const result = resolveInlineNavigationTarget(clickEvent, {
      currentPageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/start.html',
      validationOptions,
      cacheBusterParamName: 'v',
    });

    expect(result).toBe(
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/index.html',
    );
  });

  it('does not intercept external host links', () => {
    const anchor = document.createElement('a');
    anchor.href = 'https://example.org/report.html';
    anchor.setAttribute('href', 'https://example.org/report.html');
    const clickEvent = new MouseEvent('click', { bubbles: true, button: 0 });
    Object.defineProperty(clickEvent, 'target', {
      value: anchor,
      configurable: true,
    });

    const result = resolveInlineNavigationTarget(clickEvent, {
      currentPageUrl: validationOptions.currentPageUrl,
      validationOptions,
      cacheBusterParamName: 'v',
    });

    expect(result).toBeUndefined();
  });

  it('does not intercept non-html links', () => {
    const anchor = document.createElement('a');
    anchor.href = 'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/file.csv';
    anchor.setAttribute(
      'href',
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/file.csv',
    );
    const clickEvent = new MouseEvent('click', { bubbles: true, button: 0 });
    Object.defineProperty(clickEvent, 'target', {
      value: anchor,
      configurable: true,
    });

    const result = resolveInlineNavigationTarget(clickEvent, {
      currentPageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/start.html',
      validationOptions,
      cacheBusterParamName: 'v',
    });

    expect(result).toBeUndefined();
  });

  it('does not intercept middle-click navigation', () => {
    const anchor = document.createElement('a');
    anchor.href = 'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/index.html';
    anchor.setAttribute(
      'href',
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/index.html',
    );
    const clickEvent = new MouseEvent('click', { bubbles: true, button: 1 });
    Object.defineProperty(clickEvent, 'target', {
      value: anchor,
      configurable: true,
    });

    const result = resolveInlineNavigationTarget(clickEvent, {
      currentPageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/start.html',
      validationOptions,
      cacheBusterParamName: 'v',
    });

    expect(result).toBeUndefined();
  });

  it('does not intercept modified-click navigation', () => {
    const anchor = document.createElement('a');
    anchor.href = 'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/index.html';
    anchor.setAttribute(
      'href',
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/index.html',
    );

    const modifierEvents: MouseEvent[] = [
      new MouseEvent('click', { bubbles: true, button: 0, ctrlKey: true }),
      new MouseEvent('click', { bubbles: true, button: 0, metaKey: true }),
      new MouseEvent('click', { bubbles: true, button: 0, shiftKey: true }),
      new MouseEvent('click', { bubbles: true, button: 0, altKey: true }),
    ];

    modifierEvents.forEach((clickEvent) => {
      Object.defineProperty(clickEvent, 'target', {
        value: anchor,
        configurable: true,
      });

      const result = resolveInlineNavigationTarget(clickEvent, {
        currentPageUrl:
          'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/start.html',
        validationOptions,
        cacheBusterParamName: 'v',
      });

      expect(result).toBeUndefined();
    });
  });

  it('does not intercept mailto and tel links', () => {
    const createClickEventForHref = (href: string): MouseEvent => {
      const anchor = document.createElement('a');
      anchor.href = href;
      anchor.setAttribute('href', href);
      const clickEvent = new MouseEvent('click', { bubbles: true, button: 0 });
      Object.defineProperty(clickEvent, 'target', {
        value: anchor,
        configurable: true,
      });
      return clickEvent;
    };

    const mailtoResult = resolveInlineNavigationTarget(
      createClickEventForHref('mailto:owner@contoso.com'),
      {
        currentPageUrl:
          'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/start.html',
        validationOptions,
        cacheBusterParamName: 'v',
      },
    );
    const telResult = resolveInlineNavigationTarget(createClickEventForHref('tel:+12065551212'), {
      currentPageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/start.html',
      validationOptions,
      cacheBusterParamName: 'v',
    });

    expect(mailtoResult).toBeUndefined();
    expect(telResult).toBeUndefined();
  });

  it('resolves .aspx links when extension is allowed', () => {
    const anchor = document.createElement('a');
    anchor.href = 'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/page.aspx?v=123';
    anchor.setAttribute(
      'href',
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/page.aspx?v=123',
    );
    const clickEvent = new MouseEvent('click', { bubbles: true, button: 0 });
    Object.defineProperty(clickEvent, 'target', {
      value: anchor,
      configurable: true,
    });

    const result = resolveInlineNavigationTarget(clickEvent, {
      currentPageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/start.html',
      validationOptions,
      cacheBusterParamName: 'v',
    });

    expect(result).toBe(
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/page.aspx',
    );
  });

  it('does not intercept .aspx links when extension is not allowed', () => {
    const anchor = document.createElement('a');
    anchor.href = 'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/page.aspx';
    anchor.setAttribute(
      'href',
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/page.aspx',
    );
    const clickEvent = new MouseEvent('click', { bubbles: true, button: 0 });
    Object.defineProperty(clickEvent, 'target', {
      value: anchor,
      configurable: true,
    });

    const result = resolveInlineNavigationTarget(clickEvent, {
      currentPageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/start.html',
      validationOptions: {
        ...validationOptions,
        allowedFileExtensions: ['.html', '.htm'],
      },
      cacheBusterParamName: 'v',
    });

    expect(result).toBeUndefined();
  });

  it('intercepts same-site html links even when target is _blank', () => {
    const anchor = document.createElement('a');
    anchor.href = 'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/index.html';
    anchor.setAttribute(
      'href',
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/index.html',
    );
    anchor.setAttribute('target', '_blank');
    const clickEvent = new MouseEvent('click', { bubbles: true, button: 0 });
    Object.defineProperty(clickEvent, 'target', {
      value: anchor,
      configurable: true,
    });

    const result = resolveInlineNavigationTarget(clickEvent, {
      currentPageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/start.html',
      validationOptions,
      cacheBusterParamName: 'v',
    });

    expect(result).toBe(
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/index.html',
    );
  });

  it('intercepts same-site html links even when download attribute is present', () => {
    const anchor = document.createElement('a');
    anchor.href = 'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/index.html';
    anchor.setAttribute(
      'href',
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/index.html',
    );
    anchor.setAttribute('download', 'report.html');
    const clickEvent = new MouseEvent('click', { bubbles: true, button: 0 });
    Object.defineProperty(clickEvent, 'target', {
      value: anchor,
      configurable: true,
    });

    const result = resolveInlineNavigationTarget(clickEvent, {
      currentPageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/start.html',
      validationOptions,
      cacheBusterParamName: 'v',
    });

    expect(result).toBe(
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/index.html',
    );
  });

  it('intercepts FullCalendar forced-url clicks when target is not the anchor', () => {
    const container = document.createElement('div');
    container.className = 'fc-event-forced-url';
    const anchor = document.createElement('a');
    anchor.href = 'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/calendar-item.html';
    anchor.setAttribute(
      'href',
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/calendar-item.html',
    );
    const inner = document.createElement('span');
    inner.textContent = 'Calendar event';
    container.appendChild(anchor);
    container.appendChild(inner);

    const clickEvent = new MouseEvent('click', { bubbles: true, button: 0 });
    Object.defineProperty(clickEvent, 'target', {
      value: inner,
      configurable: true,
    });

    const result = resolveInlineNavigationTarget(clickEvent, {
      currentPageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/start.html',
      validationOptions,
      cacheBusterParamName: 'v',
    });

    expect(result).toBe(
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/calendar-item.html',
    );
  });

  it('resolves link when click target is a text node inside anchor', () => {
    const anchor = document.createElement('a');
    anchor.href = 'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/index.html';
    anchor.setAttribute(
      'href',
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/index.html',
    );
    anchor.appendChild(document.createTextNode('Open report'));

    const clickEvent = new MouseEvent('click', { bubbles: true, button: 0 });
    Object.defineProperty(clickEvent, 'target', {
      value: anchor.firstChild,
      configurable: true,
    });

    const result = resolveInlineNavigationTarget(clickEvent, {
      currentPageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/start.html',
      validationOptions,
      cacheBusterParamName: 'v',
    });

    expect(result).toBe(
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/index.html',
    );
  });

  it('handles text-node targets in forced-url containers', () => {
    const container = document.createElement('div');
    container.className = 'fc-event-forced-url';
    const anchor = document.createElement('a');
    anchor.href = 'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/calendar-item-2.html';
    anchor.setAttribute(
      'href',
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/calendar-item-2.html',
    );
    const inner = document.createElement('span');
    inner.appendChild(document.createTextNode('Calendar event 2'));
    container.appendChild(anchor);
    container.appendChild(inner);

    const clickEvent = new MouseEvent('click', { bubbles: true, button: 0 });
    Object.defineProperty(clickEvent, 'target', {
      value: inner.firstChild,
      configurable: true,
    });

    const result = resolveInlineNavigationTarget(clickEvent, {
      currentPageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/start.html',
      validationOptions,
      cacheBusterParamName: 'v',
    });

    expect(result).toBe(
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/calendar-item-2.html',
    );
  });

  it('resolves links when click target is a non-anchor container with a single descendant anchor', () => {
    const listItem = document.createElement('li');
    const anchor = document.createElement('a');
    anchor.href = 'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/from-list-item.html';
    anchor.setAttribute(
      'href',
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/from-list-item.html',
    );
    const label = document.createElement('span');
    label.textContent = 'List entry';
    listItem.appendChild(anchor);
    listItem.appendChild(label);

    const clickEvent = new MouseEvent('click', { bubbles: true, button: 0 });
    Object.defineProperty(clickEvent, 'target', {
      value: listItem,
      configurable: true,
    });

    const result = resolveInlineNavigationTarget(clickEvent, {
      currentPageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/start.html',
      validationOptions,
      cacheBusterParamName: 'v',
    });

    expect(result).toBe(
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/from-list-item.html',
    );
  });

  it('does not guess a target when container has multiple descendant anchors', () => {
    const container = document.createElement('div');
    const firstAnchor = document.createElement('a');
    firstAnchor.href = 'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/one.html';
    firstAnchor.setAttribute(
      'href',
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/one.html',
    );
    const secondAnchor = document.createElement('a');
    secondAnchor.href = 'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/two.html';
    secondAnchor.setAttribute(
      'href',
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/two.html',
    );
    container.appendChild(firstAnchor);
    container.appendChild(secondAnchor);

    const clickEvent = new MouseEvent('click', { bubbles: true, button: 0 });
    Object.defineProperty(clickEvent, 'target', {
      value: container,
      configurable: true,
    });

    const result = resolveInlineNavigationTarget(clickEvent, {
      currentPageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/start.html',
      validationOptions,
      cacheBusterParamName: 'v',
    });

    expect(result).toBeUndefined();
  });

  it('resolves SVG anchor links that use xlink:href', () => {
    const svgNamespace = 'http://www.w3.org/2000/svg';
    const xlinkNamespace = 'http://www.w3.org/1999/xlink';
    const svgAnchor = document.createElementNS(svgNamespace, 'a');
    svgAnchor.setAttributeNS(
      xlinkNamespace,
      'xlink:href',
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/svg-target.html',
    );

    const innerShape = document.createElementNS(svgNamespace, 'rect');
    svgAnchor.appendChild(innerShape);
    const clickEvent = new MouseEvent('click', { bubbles: true, button: 0 });
    Object.defineProperty(clickEvent, 'target', {
      value: innerShape,
      configurable: true,
    });

    const result = resolveInlineNavigationTarget(clickEvent, {
      currentPageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/start.html',
      validationOptions,
      cacheBusterParamName: 'v',
    });

    expect(result).toBe(
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/svg-target.html',
    );
  });

  it('resolves links when events are retargeted and anchor is only available via composed path', () => {
    const host = document.createElement('div');
    const anchor = document.createElement('a');
    anchor.href = 'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/composed.html';
    anchor.setAttribute(
      'href',
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/composed.html',
    );
    const inner = document.createElement('span');
    anchor.appendChild(inner);
    host.appendChild(anchor);

    const clickEvent = new MouseEvent('click', { bubbles: true, button: 0 });
    Object.defineProperty(clickEvent, 'target', {
      value: host,
      configurable: true,
    });
    Object.defineProperty(clickEvent, 'composedPath', {
      value: () => [inner, anchor, host, document.body, document, window],
      configurable: true,
    });

    const result = resolveInlineNavigationTarget(clickEvent, {
      currentPageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/start.html',
      validationOptions,
      cacheBusterParamName: 'v',
    });

    expect(result).toBe(
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/composed.html',
    );
  });

  it('returns undefined for detached text-node targets', () => {
    const detachedTextNode = document.createTextNode('Detached');
    const clickEvent = new MouseEvent('click', { bubbles: true, button: 0 });
    Object.defineProperty(clickEvent, 'target', {
      value: detachedTextNode,
      configurable: true,
    });

    const result = resolveInlineNavigationTarget(clickEvent, {
      currentPageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/start.html',
      validationOptions,
      cacheBusterParamName: 'v',
    });

    expect(result).toBeUndefined();
  });

  it('still resolves when event is already defaultPrevented', () => {
    const anchor = document.createElement('a');
    anchor.href = 'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/index.html';
    anchor.setAttribute(
      'href',
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/index.html',
    );
    const clickEvent = new MouseEvent('click', { bubbles: true, button: 0 });
    Object.defineProperty(clickEvent, 'target', {
      value: anchor,
      configurable: true,
    });
    Object.defineProperty(clickEvent, 'defaultPrevented', {
      value: true,
      configurable: true,
    });

    const result = resolveInlineNavigationTarget(clickEvent, {
      currentPageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/start.html',
      validationOptions,
      cacheBusterParamName: 'v',
    });

    expect(result).toBe(
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/index.html',
    );
  });

  it('does not relax path-prefix checks for relative links', () => {
    const anchor = document.createElement('a');
    anchor.setAttribute('href', 'GPO_Blocked_Inheritance.html');
    Object.defineProperty(anchor, 'href', {
      value:
        'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/GPO_Blocked_Inheritance.html',
      configurable: true,
    });
    const clickEvent = new MouseEvent('click', { bubbles: true, button: 0 });
    Object.defineProperty(clickEvent, 'target', {
      value: anchor,
      configurable: true,
    });

    const result = resolveInlineNavigationTarget(clickEvent, {
      currentPageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/start.html',
      validationOptions: {
        ...validationOptions,
        allowedPathPrefixes: ['/sites/testsite1/siteassets/reports/subfolder/'],
      },
      cacheBusterParamName: 'v',
    });

    expect(result).toBeUndefined();
  });

  it('does not relax path-prefix checks for absolute links', () => {
    const anchor = document.createElement('a');
    anchor.href = 'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/GPO_Blocked_Inheritance.html';
    anchor.setAttribute(
      'href',
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/GPO_Blocked_Inheritance.html',
    );
    const clickEvent = new MouseEvent('click', { bubbles: true, button: 0 });
    Object.defineProperty(clickEvent, 'target', {
      value: anchor,
      configurable: true,
    });

    const result = resolveInlineNavigationTarget(clickEvent, {
      currentPageUrl:
        'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Reports/start.html',
      validationOptions: {
        ...validationOptions,
        allowedPathPrefixes: ['/sites/testsite1/siteassets/reports/subfolder/'],
      },
      cacheBusterParamName: 'v',
    });

    expect(result).toBeUndefined();
  });
});

