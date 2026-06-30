import { getInlineNavigationBridgeScript } from '../InlineNavigationBridgeScript';

describe('InlineNavigationBridgeScript', () => {
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
});
