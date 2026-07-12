import { getInlineNavigationBridgeScript } from '../InlineNavigationBridgeScript';

describe('InlineNavigationBridgeScript', () => {
  it('produces executable JavaScript', () => {
    const script = getInlineNavigationBridgeScript(
      'https://contoso.sharepoint.com/sites/Test/SiteAssets/Reports/start.html',
      ['.html'],
      ['/sites/Test/SiteAssets/Reports/'],
      'https://contoso.sharepoint.com/sites/Test/SitePages/Viewer.aspx?dashboard=main&OR=Teams',
      'uhvPage',
      ['dashboard'],
    );

    // The bridge is shipped as an inline script string, so compilation is the
    // direct contract under test here rather than application-side evaluation.
    // eslint-disable-next-line no-new-func
    expect(() => new Function(script)).not.toThrow();
  });

  it('handles fragment-only anchors before emitting inline navigation', () => {
    const script = getInlineNavigationBridgeScript();
    const hashHandlerIndex = script.indexOf(
      'if (navigateToSamePageHash(rawHref, event)) { return; }',
    );
    const unresolvedHashIndex = script.indexOf('if (isSamePageHashHref(rawHref)) { return; }');
    const blockedProtocolIndex = script.indexOf('if (hasBlockedProtocol(rawHref)) { return; }');
    const emitIndex = script.indexOf('emit(absoluteTargetUrl, event);');

    expect(script).toContain('var navigateToSamePageHash = function(rawHref, event)');
    expect(script).toContain("if (hashHref === '#') { return null; }");
    expect(script).toContain('var target = findSamePageHashTarget(hashHref);');
    expect(script).toContain('if (!target) { return false; }');
    expect(script).toContain('window.location.hash = hashHref;');
    expect(script).toContain('target.scrollIntoView();');
    expect(hashHandlerIndex).toBeGreaterThan(-1);
    expect(unresolvedHashIndex).toBeGreaterThan(hashHandlerIndex);
    expect(blockedProtocolIndex).toBeGreaterThan(unresolvedHashIndex);
    expect(emitIndex).toBeGreaterThan(blockedProtocolIndex);
  });

  it('validates targets before suppressing native link behavior', () => {
    const script = getInlineNavigationBridgeScript(
      'https://contoso.sharepoint.com/sites/Test/SiteAssets/Reports/start.html',
      ['.html'],
      ['/sites/Test/SiteAssets/Reports/'],
    );

    const eligibilityIndex = script.indexOf(
      'if (!isEligibleTargetUrl(absoluteTargetUrl)) { return; }',
    );
    const nativeBehaviorIndex = script.indexOf(
      'if (shouldKeepNativeAnchorBehavior(anchor, event)) {',
    );
    const emitIndex = script.indexOf('emit(absoluteTargetUrl, event);');

    expect(script).toContain("anchor.removeAttribute('download')");
    expect(script).toContain("target !== '_self'");
    expect(script).toContain('event.metaKey || event.ctrlKey || event.shiftKey || event.altKey');
    expect(script).toContain('target.host.toLowerCase() !== parsedBase.host.toLowerCase()');
    expect(script).toContain("var allowedFileExtensions = [\".html\"];");
    expect(script).toContain("var configuredAllowedPathPrefixes = [\"/sites/Test/SiteAssets/Reports/\"];");
    expect(eligibilityIndex).toBeGreaterThan(-1);
    expect(nativeBehaviorIndex).toBeGreaterThan(eligibilityIndex);
    expect(emitIndex).toBeGreaterThan(nativeBehaviorIndex);
  });

  it('rewrites validated generated new-tab links through the viewer host page', () => {
    const script = getInlineNavigationBridgeScript(
      '/sites/Test/SiteAssets/Reports/start.html',
      ['.html'],
      ['/sites/Test/SiteAssets/Reports/'],
      'https://contoso.sharepoint.com/sites/Test/SitePages/Viewer.aspx?dashboard=main&OR=Teams&viewerTwoPage=%2Fsites%2FTest%2FSiteAssets%2FReports%2FSecond.html',
      'reportPage',
      ['dashboard'],
    );

    expect(script).toContain(
      'var configuredBaseUrl = "/sites/Test/SiteAssets/Reports/start.html";',
    );
    expect(script).toContain('new URL(configuredBase || fallbackBase, fallbackBase || undefined)');
    expect(script).toContain('var deepLinkQueryParamName = "reportPage";');
    expect(script).toContain('var preservedHostQueryParamNames = ["dashboard"];');
    expect(script).toContain('var rewriteNativeTargetAnchor = function(anchor, targetUrl)');
    expect(script).toContain("anchor.setAttribute('href', hostDeepLinkUrl)");
    expect(script).toContain('rewriteNativeTargetAnchor(anchor, absoluteTargetUrl);');
  });

  it('does not emit native host deep links when inbound deep links are disabled', () => {
    const script = getInlineNavigationBridgeScript(
      '/sites/Test/SiteAssets/Reports/start.html',
      ['.html'],
      ['/sites/Test/SiteAssets/Reports/'],
      'https://contoso.sharepoint.com/sites/Test/SitePages/Viewer.aspx',
      'uhvPage',
      [],
      false,
    );

    expect(script).toContain('var enableHostDeepLinkUrls = false;');
    expect(script).toContain('if (!enableHostDeepLinkUrls) { return ""; }');
  });

  it('rewrites a generated new-tab anchor before native navigation can download it', () => {
    const frame = document.createElement('iframe');
    document.body.appendChild(frame);
    const frameWindow = frame.contentWindow;
    const frameDocument = frame.contentDocument;
    expect(frameWindow).not.toBeNull();
    expect(frameDocument).not.toBeNull();
    if (!frameWindow || !frameDocument) {
      return;
    }
    const executableFrameWindow = frameWindow as Window & typeof globalThis;

    frameDocument.head.innerHTML =
      '<base href="https://contoso.sharepoint.com/sites/Test/SiteAssets/Reports/start.html">';
    frameDocument.body.innerHTML =
      '<a id="generated" href="next.html" target="_blank" download>Next</a>';
    executableFrameWindow.eval(
      getInlineNavigationBridgeScript(
        '/sites/Test/SiteAssets/Reports/start.html',
        ['.html'],
        ['/sites/Test/SiteAssets/Reports/'],
        'https://contoso.sharepoint.com/sites/Test/SitePages/Viewer.aspx?dashboard=main&OR=Teams&viewerTwoPage=%2Fsites%2FTest%2FSiteAssets%2FReports%2FSecond.html',
        'uhvPage',
        ['dashboard'],
      ),
    );

    const anchor = frameDocument.getElementById('generated') as HTMLAnchorElement;
    anchor.dispatchEvent(
      new executableFrameWindow.MouseEvent('mousedown', {
        bubbles: true,
        button: 0,
      }),
    );

    expect(anchor.hasAttribute('download')).toBe(false);
    expect(anchor.getAttribute('data-uhv-inline-href')).toBe(
      'https://contoso.sharepoint.com/sites/Test/SiteAssets/Reports/next.html',
    );
    expect(anchor.href).toContain(
      'https://contoso.sharepoint.com/sites/Test/SitePages/Viewer.aspx?dashboard=main&',
    );
    expect(anchor.href).toContain('uhvPage=');
    expect(anchor.href).not.toContain('OR=Teams');
    expect(anchor.href).toContain('viewerTwoPage=');
    frame.remove();
  });

  it('forwards validated navigation messages from nested frames', () => {
    const script = getInlineNavigationBridgeScript();

    expect(script).toContain('var onNestedNavigationMessage = function(event)');
    expect(script).toContain("window.addEventListener('message', onNestedNavigationMessage)");
    expect(script).toContain("payload.type !== 'uhv-inline-nav'");
  });
});
