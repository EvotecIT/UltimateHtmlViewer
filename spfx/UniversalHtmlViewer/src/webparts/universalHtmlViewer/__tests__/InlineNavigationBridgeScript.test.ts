import { getInlineNavigationBridgeScript } from '../InlineNavigationBridgeScript';

describe('InlineNavigationBridgeScript', () => {
  it('produces executable JavaScript', () => {
    const script = getInlineNavigationBridgeScript(
      'https://contoso.sharepoint.com/sites/Test/SiteAssets/Reports/start.html',
      ['.html'],
      ['/sites/Test/SiteAssets/Reports/'],
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

    const nativeBehaviorIndex = script.indexOf(
      'if (shouldKeepNativeAnchorBehavior(anchor)) { return; }',
    );
    const eligibilityIndex = script.indexOf(
      'if (!isEligibleTargetUrl(absoluteTargetUrl)) { return; }',
    );
    const emitIndex = script.indexOf('emit(absoluteTargetUrl, event);');

    expect(script).not.toContain("anchor.hasAttribute('download')");
    expect(script).toContain("target !== '_self'");
    expect(script).toContain('target.host.toLowerCase() !== parsedBase.host.toLowerCase()');
    expect(script).toContain("var allowedFileExtensions = [\".html\"];");
    expect(script).toContain("var configuredAllowedPathPrefixes = [\"/sites/Test/SiteAssets/Reports/\"];");
    expect(nativeBehaviorIndex).toBeGreaterThan(-1);
    expect(eligibilityIndex).toBeGreaterThan(nativeBehaviorIndex);
    expect(emitIndex).toBeGreaterThan(eligibilityIndex);
  });

  it('forwards validated navigation messages from nested frames', () => {
    const script = getInlineNavigationBridgeScript();

    expect(script).toContain('var onNestedNavigationMessage = function(event)');
    expect(script).toContain("window.addEventListener('message', onNestedNavigationMessage)");
    expect(script).toContain("payload.type !== 'uhv-inline-nav'");
  });
});
