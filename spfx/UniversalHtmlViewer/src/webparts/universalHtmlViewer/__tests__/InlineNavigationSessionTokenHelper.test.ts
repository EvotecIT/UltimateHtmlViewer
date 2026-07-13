import {
  INLINE_NAVIGATION_TOKEN_ATTRIBUTE,
  prepareAndStageInlineNavigationSession,
  prepareInlineNavigationSessionHtml,
} from '../InlineNavigationSessionTokenHelper';

describe('InlineNavigationSessionTokenHelper', () => {
  const bridgeHtml =
    '<html><head><script data-uhv-inline-nav-bridge="1">/* bridge */</script></head></html>';

  it('creates a fresh per-frame capability without mutating the cached source HTML', () => {
    const first = prepareInlineNavigationSessionHtml(bridgeHtml);
    const second = prepareInlineNavigationSessionHtml(bridgeHtml);

    expect(first.navigationToken).toMatch(/^[a-f0-9]{32}$/);
    expect(second.navigationToken).toMatch(/^[a-f0-9]{32}$/);
    expect(second.navigationToken).not.toBe(first.navigationToken);
    expect(first.html).toContain(
      `${INLINE_NAVIGATION_TOKEN_ATTRIBUTE}="${first.navigationToken}"`,
    );
    expect(second.html).toContain(
      `${INLINE_NAVIGATION_TOKEN_ATTRIBUTE}="${second.navigationToken}"`,
    );
    expect(bridgeHtml).not.toContain(INLINE_NAVIGATION_TOKEN_ATTRIBUTE);
  });

  it('stages the exact capability that is embedded in the prepared frame document', () => {
    const iframe = document.createElement('iframe');
    const preparedHtml = prepareAndStageInlineNavigationSession(iframe, bridgeHtml);
    const stagedToken = iframe.getAttribute(INLINE_NAVIGATION_TOKEN_ATTRIBUTE);

    expect(stagedToken).toMatch(/^[a-f0-9]{32}$/);
    expect(preparedHtml).toContain(
      `${INLINE_NAVIGATION_TOKEN_ATTRIBUTE}="${stagedToken}"`,
    );
  });

  it('fails closed when no injected bridge can carry the capability', () => {
    const iframe = document.createElement('iframe');
    iframe.setAttribute(INLINE_NAVIGATION_TOKEN_ATTRIBUTE, 'stale-token');

    const preparedHtml = prepareAndStageInlineNavigationSession(
      iframe,
      '<html><body>No bridge</body></html>',
    );

    expect(preparedHtml).toBe('<html><body>No bridge</body></html>');
    expect(iframe.hasAttribute(INLINE_NAVIGATION_TOKEN_ATTRIBUTE)).toBe(false);
  });
});
