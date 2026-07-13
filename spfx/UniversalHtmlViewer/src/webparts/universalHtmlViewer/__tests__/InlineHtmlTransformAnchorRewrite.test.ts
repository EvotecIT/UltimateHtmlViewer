import { prepareInlineHtmlForSrcDoc } from '../InlineHtmlTransformHelper';

describe('InlineHtmlTransformHelper anchor rewrite gating', () => {
  it('does not rewrite report anchors unless host deep-link anchors are enabled', () => {
    const result = prepareInlineHtmlForSrcDoc(
      '<html><body><a href="Computers.html">Computers</a></body></html>',
      '/sites/TheDashboardPage/Shared Documents/',
      'https://knauf.sharepoint.com/sites/TheDashboardPage/SitePages/TheDashboardPage.aspx',
    );

    expect(result).toContain('href="Computers.html"');
    expect(result).not.toContain('data-uhv-inline-href=');
  });

  it('rewrites report anchors when host deep-link anchors are enabled', () => {
    const result = prepareInlineHtmlForSrcDoc(
      '<html><body><a href="Computers.html">Computers</a></body></html>',
      '/sites/TheDashboardPage/Shared Documents/',
      'https://knauf.sharepoint.com/sites/TheDashboardPage/SitePages/TheDashboardPage.aspx?OR=Teams-HL',
      {
        rewriteInlineAnchorHrefs: true,
      },
    );

    expect(result).toContain('data-uhv-inline-href=');
    expect(result).toContain(
      'href="https://knauf.sharepoint.com/sites/TheDashboardPage/SitePages/TheDashboardPage.aspx?uhvPage=',
    );
    expect(result).not.toContain('OR=Teams-HL');
  });

  it('keeps generated strict CSP nonces limited to injected UHV scripts', () => {
    const result = prepareInlineHtmlForSrcDoc(
      '<html><head><script>window.reportLoaded = true;</script></head><body></body></html>',
      '/sites/TheDashboardPage/Shared Documents/',
      'https://knauf.sharepoint.com/sites/TheDashboardPage/SitePages/TheDashboardPage.aspx',
      {
        enforceStrictInlineCsp: true,
      },
    );
    const parsed = new DOMParser().parseFromString(result, 'text/html');
    const reportScript = Array.from(parsed.querySelectorAll('script')).find((script) =>
      (script.textContent || '').includes('window.reportLoaded'),
    );
    const compatibilityShim = parsed.querySelector(
      'script[data-uhv-history-compat="1"]',
    );
    const navigationBridge = parsed.querySelector(
      'script[data-uhv-inline-nav-bridge="1"]',
    );
    const generatedCsp = parsed.querySelector('meta[data-uhv-inline-csp="1"]');
    const nonce = compatibilityShim?.getAttribute('nonce') || '';

    expect(reportScript?.hasAttribute('nonce')).toBe(false);
    expect(nonce).not.toBe('');
    expect(navigationBridge?.getAttribute('nonce')).toBe(nonce);
    expect(generatedCsp?.getAttribute('content')).toContain(`'nonce-${nonce}'`);
  });
});
