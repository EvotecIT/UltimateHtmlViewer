import { getInlineNavigationBridgeScript } from '../InlineNavigationBridgeScript';

describe('InlineNavigationBridgeScript', () => {
  it('handles fragment-only anchors before emitting inline navigation', () => {
    const script = getInlineNavigationBridgeScript();
    const hashHandlerIndex = script.indexOf(
      'if (navigateToSamePageHash(rawHref, event)) { return; }',
    );
    const blockedProtocolIndex = script.indexOf('if (hasBlockedProtocol(rawHref)) { return; }');
    const emitIndex = script.indexOf('emit(absoluteTargetUrl, event);');

    expect(script).toContain('var navigateToSamePageHash = function(rawHref, event)');
    expect(script).toContain('window.location.hash = hashHref;');
    expect(script).toContain('target.scrollIntoView();');
    expect(hashHandlerIndex).toBeGreaterThan(-1);
    expect(blockedProtocolIndex).toBeGreaterThan(hashHandlerIndex);
    expect(emitIndex).toBeGreaterThan(blockedProtocolIndex);
  });
});
