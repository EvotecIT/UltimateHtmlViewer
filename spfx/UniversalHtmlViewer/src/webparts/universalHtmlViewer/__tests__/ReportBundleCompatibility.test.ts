import * as fs from 'fs';
import * as path from 'path';
import {
  clearExternalScriptInliningCacheForTests,
  inlineAllowedExternalScripts,
} from '../ExternalScriptInliningHelper';
import { appendAdditionalCspHostSources } from '../InlineCspSourceHelper';
import { wireInlineAnchorRuntimeRewrite } from '../InlineAnchorRuntimeRewriteHelper';
import { resolveInlineDeepLinkTarget } from '../InlineDeepLinkHelper';
import {
  IPrepareInlineHtmlForSrcDocOptions,
  prepareInlineHtmlForBlobUrl,
  prepareInlineHtmlForSrcDoc,
} from '../InlineHtmlTransformHelper';

/**
 * Compatibility spike for TestimoX report bundles (index.html + page HTML,
 * assets/, data/<dataset>/<chunk>.js sidecars, manifest.json) hosted through
 * UHV. See docs/Report-Bundle-Compatibility-Spike.md for the findings.
 */
const SAMPLE_ROOT = path.resolve(__dirname, '../../../../../../samples/report-bundle');
const LIBRARY_FOLDER_URL =
  'https://contoso.sharepoint.com/sites/Reports/Shared%20Documents/TestimoX/';
const INDEX_URL = `${LIBRARY_FOLDER_URL}index.html`;
const PAGE2_URL = `${LIBRARY_FOLDER_URL}page2.html`;
const HOST_PAGE_URL = 'https://contoso.sharepoint.com/sites/Reports/SitePages/Reports.aspx';
const TENANT_ORIGIN = 'https://contoso.sharepoint.com';

// Mirrors UniversalHtmlViewerWebPart.getInlineContentOptions() for a host page
// with allowQueryStringPageOverride enabled and default extensions.
const HOST_OPTIONS: IPrepareInlineHtmlForSrcDocOptions = {
  rewriteInlineAnchorHrefs: true,
  rewriteInlineAnchorAllowedFileExtensions: [],
  rewriteInlineAnchorAllowedPathPrefixes: ['/sites/Reports/Shared Documents/TestimoX/'],
  rewriteInlineAnchorDeepLinkQueryParamName: 'uhvPage',
  rewriteInlineAnchorPreservedHostQueryParamNames: [],
};

function readSample(relativePath: string): string {
  return fs.readFileSync(path.join(SAMPLE_ROOT, relativePath), 'utf8');
}

function parse(html: string): Document {
  return new DOMParser().parseFromString(html, 'text/html');
}

function getInjectedCspDirectives(html: string): Map<string, string[]> {
  const meta = parse(html).querySelector('meta[data-uhv-inline-csp="1"]');
  const directives = new Map<string, string[]>();
  (meta?.getAttribute('content') || '')
    .split(';')
    .map((entry) => entry.trim())
    .filter((entry) => entry.length > 0)
    .forEach((entry) => {
      const [name, ...sources] = entry.split(/\s+/);
      directives.set(name.toLowerCase(), sources);
    });
  return directives;
}

function safeDecodePath(value: string): string {
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
}

/**
 * Minimal CSP host-source matcher for the source shapes UHV emits
 * (scheme + host origins, blob:, data:). Keyword sources other than scheme
 * sources are ignored, which is intentional: 'self' in an about:srcdoc
 * document depends on the sandbox preset and cannot be evaluated here.
 * Path parts follow CSP rules: a source path ending in "/" is a prefix match,
 * any other path must match exactly, both compared percent-decoded.
 */
function isUrlAllowedBySources(url: string, sources: string[]): boolean {
  const target = new URL(url);
  return sources.some((source) => {
    if (source === target.protocol) {
      return true;
    }
    if (!/^https?:\/\//i.test(source)) {
      return false;
    }
    let allowed: URL;
    try {
      allowed = new URL(source);
    } catch {
      return false;
    }
    if (allowed.origin.toLowerCase() !== target.origin.toLowerCase()) {
      return false;
    }
    if (allowed.pathname === '/') {
      return true;
    }
    const sourcePath = safeDecodePath(allowed.pathname);
    const targetPath = safeDecodePath(target.pathname);
    return sourcePath.endsWith('/') ? targetPath.startsWith(sourcePath) : targetPath === sourcePath;
  });
}

function getScriptAndStyleUrls(html: string, baseHref: string): { scripts: string[]; styles: string[] } {
  const document = parse(html);
  const scripts = Array.from(document.querySelectorAll('script[src]')).map((script) =>
    new URL(script.getAttribute('src') || '', baseHref).toString(),
  );
  const styles = Array.from(document.querySelectorAll('link[rel="stylesheet"][href]')).map((link) =>
    new URL(link.getAttribute('href') || '', baseHref).toString(),
  );
  return { scripts, styles };
}

function waitForAttribute(element: Element, attributeName: string): Promise<void> {
  if (element.hasAttribute(attributeName)) {
    return Promise.resolve();
  }

  return new Promise((resolve, reject) => {
    let timeoutId = 0;
    const observer = new MutationObserver(() => {
      if (element.hasAttribute(attributeName)) {
        window.clearTimeout(timeoutId);
        observer.disconnect();
        resolve();
      }
    });
    timeoutId = window.setTimeout(() => {
      observer.disconnect();
      reject(new Error(`Timed out waiting for ${attributeName}.`));
    }, 1000);
    observer.observe(element, { attributes: true, attributeFilter: [attributeName] });
  });
}

describe('Report bundle compatibility (TestimoX spike)', () => {
  const indexHtml = readSample('index.html');
  const page2Html = readSample('page2.html');

  describe('SharePointFileContent (srcdoc)', () => {
    const prepared = prepareInlineHtmlForSrcDoc(indexHtml, INDEX_URL, HOST_PAGE_URL, HOST_OPTIONS);

    it('injects a CSP whose script-src and style-src allow assets/*.js, data/*.js and assets/*.css from the library', () => {
      const directives = getInjectedCspDirectives(prepared);
      const { scripts, styles } = getScriptAndStyleUrls(indexHtml, INDEX_URL);

      expect(scripts).toEqual([
        `${LIBRARY_FOLDER_URL}assets/app.js`,
        `${LIBRARY_FOLDER_URL}data/probes/0.js`,
      ]);
      expect(styles).toEqual([`${LIBRARY_FOLDER_URL}assets/app.css`]);
      expect(directives.get('script-src')).toEqual(
        expect.arrayContaining([TENANT_ORIGIN, "'unsafe-inline'", 'blob:']),
      );
      scripts.forEach((scriptUrl) =>
        expect(isUrlAllowedBySources(scriptUrl, directives.get('script-src') || [])).toBe(true),
      );
      styles.forEach((styleUrl) =>
        expect(isUrlAllowedBySources(styleUrl, directives.get('style-src') || [])).toBe(true),
      );
      // Runtime-injected sidecars (page2 lazy load) resolve to the same folder.
      expect(
        isUrlAllowedBySources(`${LIBRARY_FOLDER_URL}data/probes/1.js`, directives.get('script-src') || []),
      ).toBe(true);
      // manifest.json via fetch() is governed by connect-src.
      expect(
        isUrlAllowedBySources(`${LIBRARY_FOLDER_URL}manifest.json`, directives.get('connect-src') || []),
      ).toBe(true);
    });

    it('scopes the injected CSP to the tenant origin, not the library path', () => {
      const directives = getInjectedCspDirectives(prepared);
      const scriptSources = directives.get('script-src') || [];

      // Any script anywhere on the tenant origin is allowed; the policy is not
      // path-scoped to the bundle folder.
      expect(
        isUrlAllowedBySources(`${TENANT_ORIGIN}/sites/Other/SiteAssets/any.js`, scriptSources),
      ).toBe(true);
      expect(scriptSources.some((source) => source.includes('/sites/'))).toBe(false);
      // A CDN script is not allowed unless the web part adds it explicitly.
      expect(isUrlAllowedBySources('https://cdn.jsdelivr.net/npm/x.js', scriptSources)).toBe(false);
      // The additional-host helper normalizes to origins, so it cannot express a
      // path-scoped source either.
      expect(
        appendAdditionalCspHostSources("'self'", [
          'https://contoso.sharepoint.com/sites/Reports/Shared%20Documents/TestimoX/',
        ]),
      ).toBe(`'self' ${TENANT_ORIGIN}`);
    });

    it('keeps the bundle loadable under the strict inline CSP because the sample has no inline report scripts', () => {
      const strict = prepareInlineHtmlForSrcDoc(indexHtml, INDEX_URL, HOST_PAGE_URL, {
        ...HOST_OPTIONS,
        enforceStrictInlineCsp: true,
      });
      const directives = getInjectedCspDirectives(strict);
      const scriptSources = directives.get('script-src') || [];
      const reportInlineScripts = Array.from(parse(strict).querySelectorAll('script:not([src])')).filter(
        (script) =>
          !script.hasAttribute('data-uhv-history-compat') &&
          !script.hasAttribute('data-uhv-inline-nav-bridge'),
      );

      expect(scriptSources).not.toContain("'unsafe-inline'");
      expect(scriptSources).not.toContain("'unsafe-eval'");
      expect(reportInlineScripts).toHaveLength(0);
      expect(isUrlAllowedBySources(`${LIBRARY_FOLDER_URL}assets/app.js`, scriptSources)).toBe(true);
      expect(isUrlAllowedBySources(`${LIBRARY_FOLDER_URL}data/probes/0.js`, scriptSources)).toBe(true);
    });

    it('skips its own CSP when the bundle ships a CSP meta tag', () => {
      const withOwnCsp = indexHtml.replace(
        '<meta charset="utf-8">',
        '<meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="script-src \'self\'">',
      );
      const result = prepareInlineHtmlForSrcDoc(withOwnCsp, INDEX_URL, HOST_PAGE_URL, HOST_OPTIONS);

      // 'self' in about:srcdoc is the parent origin only with allow-same-origin;
      // under the Strict sandbox preset it is an opaque origin and blocks the
      // bundle's own scripts. Bundles must not ship a CSP meta tag.
      expect(result).not.toContain('data-uhv-inline-csp="1"');
    });

    it('injects <base href> so relative asset and sidecar paths resolve to the library folder', () => {
      const base = parse(prepared).querySelector('base');

      expect(base?.getAttribute('href')).toBe(INDEX_URL);
      expect(new URL('assets/app.js', INDEX_URL).toString()).toBe(`${LIBRARY_FOLDER_URL}assets/app.js`);
      expect(new URL('data/probes/0.js', INDEX_URL).toString()).toBe(
        `${LIBRARY_FOLDER_URL}data/probes/0.js`,
      );
      expect(new URL('manifest.json', INDEX_URL).toString()).toBe(`${LIBRARY_FOLDER_URL}manifest.json`);
    });

    it('drops query and fragment from the base when the deep-linked source carries #section', () => {
      const result = prepareInlineHtmlForSrcDoc(
        page2Html,
        `${PAGE2_URL}?cb=123#section`,
        HOST_PAGE_URL,
        HOST_OPTIONS,
      );

      expect(parse(result).querySelector('base')?.getAttribute('href')).toBe(PAGE2_URL);
    });

    it('does not add a second <base> when the bundle already defines one', () => {
      const withOwnBase = indexHtml.replace('<meta charset="utf-8">', '<meta charset="utf-8"><base href="./">');
      const result = prepareInlineHtmlForSrcDoc(withOwnBase, INDEX_URL, HOST_PAGE_URL, HOST_OPTIONS);
      const bases = parse(result).querySelectorAll('base');

      // A relative bundle <base> would resolve against about:srcdoc and break
      // every relative asset. Bundles must not ship a <base> tag.
      expect(bases).toHaveLength(1);
      expect(bases[0].getAttribute('href')).toBe('./');
    });

    it('rewrites page2.html#section to the host deep-link form and keeps the fragment', () => {
      const anchor = parse(prepared).querySelector('#nav-page2-section');
      const rewrittenHref = anchor?.getAttribute('href') || '';
      const hostUrl = new URL(rewrittenHref);

      expect(`${hostUrl.origin}${hostUrl.pathname}`).toBe(HOST_PAGE_URL);
      expect(hostUrl.searchParams.get('uhvPage')).toBe(
        '/sites/Reports/Shared%20Documents/TestimoX/page2.html#section',
      );
      expect(hostUrl.hash).toBe('');
      expect(anchor?.getAttribute('data-uhv-inline-href')).toBe(`${PAGE2_URL}#section`);

      // Loading the host page with that URL resolves the deep link back to the
      // bundle page with the fragment intact.
      expect(
        resolveInlineDeepLinkTarget({
          pageUrl: rewrittenHref,
          fallbackUrl: INDEX_URL,
          validationOptions: {
            securityMode: 'StrictTenant',
            currentPageUrl: rewrittenHref,
            allowedPathPrefixes: ['/sites/Reports/Shared Documents/TestimoX/'],
          },
        }),
      ).toBe(`${PAGE2_URL}#section`);
    });

    it('rewrites the page2 back-link to index.html#summary', () => {
      const result = prepareInlineHtmlForSrcDoc(page2Html, PAGE2_URL, HOST_PAGE_URL, HOST_OPTIONS);
      const anchor = parse(result).querySelector('#nav-overview');

      expect(new URL(anchor?.getAttribute('href') || '').searchParams.get('uhvPage')).toBe(
        '/sites/Reports/Shared%20Documents/TestimoX/index.html#summary',
      );
    });

    it('leaves the sample free of links that UHV would not rewrite', () => {
      const anchors = Array.from(parse(prepared).querySelectorAll('a[href]'));

      anchors.forEach((anchor) => {
        expect(anchor.getAttribute('data-uhv-inline-href')).toBeTruthy();
      });
    });
  });

  describe('bundle runtime inside a frame with the injected <base>', () => {
    let iframe: HTMLIFrameElement;
    let frameWindow: Window & typeof globalThis;
    let frameDocument: Document;

    beforeEach(() => {
      iframe = document.createElement('iframe');
      document.body.appendChild(iframe);
      frameWindow = iframe.contentWindow as Window & typeof globalThis;
      frameDocument = iframe.contentDocument as Document;
      const prepared = parse(prepareInlineHtmlForSrcDoc(indexHtml, INDEX_URL, HOST_PAGE_URL, HOST_OPTIONS));
      frameDocument.head.innerHTML = prepared.querySelector('base')?.outerHTML || '';
      frameDocument.body.innerHTML = prepared.body.innerHTML;
      frameDocument.body.setAttribute('data-report-page', 'index');
      frameWindow.eval(readSample('assets/app.js'));
    });

    afterEach(() => {
      iframe.remove();
    });

    it('resolves runtime sidecar URLs against the injected base', () => {
      const hfxData = (frameWindow as unknown as { HfxData: { resolveChunkUrl: (id: string, chunk: number) => string } })
        .HfxData;

      expect(frameDocument.baseURI).toBe(INDEX_URL);
      expect(hfxData.resolveChunkUrl('probes', 0)).toBe(`${LIBRARY_FOLDER_URL}data/probes/0.js`);
    });

    it('injects a runtime sidecar script pointing at the library and accepts its registration', () => {
      const hfxData = (
        frameWindow as unknown as {
          HfxData: {
            load: (id: string, chunk: number, callback: (error: Error | null, payload?: { rows: string[][] }) => void) => void;
          };
        }
      ).HfxData;
      let loadedRows = 0;

      hfxData.load('probes', 0, (error, payload) => {
        loadedRows = error ? -1 : payload?.rows.length || 0;
      });
      const injected = frameDocument.querySelector('script[data-hfx-sidecar="probes/0"]') as HTMLScriptElement;
      // jsdom does not fetch; execute the sidecar body as the browser would.
      frameWindow.eval(readSample('data/probes/0.js'));

      expect(injected.src).toBe(`${LIBRARY_FOLDER_URL}data/probes/0.js`);
      expect(loadedRows).toBe(3);
    });

    it('rewrites the runtime-created page2.html#section link to the host deep-link form', async () => {
      const cleanup = wireInlineAnchorRuntimeRewrite({
        iframe,
        fallbackBaseUrl: INDEX_URL,
        fallbackHostPageUrl: HOST_PAGE_URL,
        allowedFileExtensions: HOST_OPTIONS.rewriteInlineAnchorAllowedFileExtensions,
        allowedPathPrefixes: HOST_OPTIONS.rewriteInlineAnchorAllowedPathPrefixes,
        deepLinkQueryParamName: 'uhvPage',
      });

      try {
        frameDocument.dispatchEvent(new frameWindow.Event('DOMContentLoaded'));
        const runtimeAnchor = frameDocument.getElementById('runtime-link-page2') as HTMLAnchorElement;
        await waitForAttribute(runtimeAnchor, 'data-uhv-inline-href');

        expect(runtimeAnchor.getAttribute('data-uhv-inline-href')).toBe(`${PAGE2_URL}#section`);
        expect(new URL(runtimeAnchor.getAttribute('href') || '').searchParams.get('uhvPage')).toBe(
          '/sites/Reports/Shared%20Documents/TestimoX/page2.html#section',
        );
      } finally {
        cleanup();
      }
    });

    function postThemeMessage(data: unknown, source: Window): void {
      frameWindow.dispatchEvent(new frameWindow.MessageEvent('message', { data, source }));
    }

    it('applies host theme tokens sent with the proposed uhv-host-theme message', () => {
      const root = frameDocument.documentElement;

      postThemeMessage(
        { type: 'uhv-host-theme', version: 1, isDark: true, tokens: { accent: '#c239b3', text: 'red', surface: '#12345' } },
        frameWindow.parent,
      );

      expect(root.getAttribute('data-report-theme')).toBe('dark');
      expect(root.style.getPropertyValue('--report-accent')).toBe('#c239b3');
      // 'red' is valid CSS but not a hex token; '#12345' is not a valid colour.
      expect(root.style.getPropertyValue('--report-text')).toBe('');
      expect(root.style.getPropertyValue('--report-surface')).toBe('');

      postThemeMessage({ type: 'uhv-host-theme', version: 1, isDark: false, tokens: {} }, frameWindow.parent);

      expect(root.getAttribute('data-report-theme')).toBe('light');
      expect(root.style.getPropertyValue('--report-accent')).toBe('');
    });

    it('ignores theme messages from other sources or versions', () => {
      const root = frameDocument.documentElement;

      postThemeMessage({ type: 'uhv-host-theme', version: 1, isDark: true, tokens: { accent: '#c239b3' } }, frameWindow);
      postThemeMessage(
        { type: 'uhv-host-theme', version: 2, isDark: true, tokens: { accent: '#c239b3' } },
        frameWindow.parent,
      );
      postThemeMessage({ type: 'uhv-host-theme', isDark: true, tokens: { accent: '#c239b3' } }, frameWindow.parent);

      expect(root.hasAttribute('data-report-theme')).toBe(false);
      expect(root.style.getPropertyValue('--report-accent')).toBe('');
    });
  });

  describe('SharePointFileBlobUrl (blob:)', () => {
    const prepared = prepareInlineHtmlForBlobUrl(indexHtml, INDEX_URL, HOST_PAGE_URL, HOST_OPTIONS);

    it('injects <base href> so relative URLs resolve to the library, but no UHV CSP', () => {
      const document = parse(prepared);

      expect(document.querySelector('base')?.getAttribute('href')).toBe(INDEX_URL);
      expect(document.querySelector('meta[data-uhv-inline-csp="1"]')).toBeNull();
      expect(document.querySelector('script[data-uhv-inline-nav-bridge="1"]')).not.toBeNull();
    });

    it('rewrites bundle links identically to srcdoc mode', () => {
      const srcDocHref = parse(prepareInlineHtmlForSrcDoc(indexHtml, INDEX_URL, HOST_PAGE_URL, HOST_OPTIONS))
        .querySelector('#nav-page2-section')
        ?.getAttribute('href');
      const blobHref = parse(prepared).querySelector('#nav-page2-section')?.getAttribute('href');

      expect(blobHref).toBe(srcDocHref);
    });

    it('has no UHV CSP in blob mode, so CSP host options have nowhere to go', () => {
      // SharePointInlineContentHelper also does not forward the host options
      // for BlobUrl; this documents the transform-level behavior.
      const withHosts = prepareInlineHtmlForBlobUrl(indexHtml, INDEX_URL, HOST_PAGE_URL, {
        ...HOST_OPTIONS,
        additionalScriptSrcHosts: ['cdn.jsdelivr.net'],
      } as IPrepareInlineHtmlForSrcDocOptions);

      expect(withHosts).not.toContain('cdn.jsdelivr.net');
    });
  });

  describe('inlineExternalScripts fallback', () => {
    const originalFetch = globalThis.fetch;

    beforeEach(() => {
      clearExternalScriptInliningCacheForTests();
    });

    afterEach(() => {
      Object.defineProperty(globalThis, 'fetch', { configurable: true, value: originalFetch });
    });

    it('inlines same-origin asset and static sidecar scripts with same-origin credentials', async () => {
      const mockFetch = jest.fn().mockImplementation((url: string) =>
        Promise.resolve({
          ok: true,
          text: () =>
            Promise.resolve(url.indexOf('/data/') >= 0 ? readSample('data/probes/0.js') : readSample('assets/app.js')),
        }),
      );
      Object.defineProperty(globalThis, 'fetch', { configurable: true, value: mockFetch });

      const inlined = await inlineAllowedExternalScripts(indexHtml, INDEX_URL, HOST_PAGE_URL, { enabled: true });
      const strict = prepareInlineHtmlForSrcDoc(inlined, INDEX_URL, HOST_PAGE_URL, {
        ...HOST_OPTIONS,
        enforceStrictInlineCsp: true,
      });
      const inlinedScripts = parse(strict).querySelectorAll('script[data-uhv-inlined-external-script]');

      expect(mockFetch).toHaveBeenCalledWith(`${LIBRARY_FOLDER_URL}assets/app.js`, {
        credentials: 'same-origin',
        mode: 'cors',
      });
      expect(mockFetch).toHaveBeenCalledWith(`${LIBRARY_FOLDER_URL}data/probes/0.js`, {
        credentials: 'same-origin',
        mode: 'cors',
      });
      expect(inlinedScripts).toHaveLength(2);
      // Under the strict inline CSP the inlined report scripts get no nonce
      // (outside SharePoint, where no page nonce exists), so this fallback and
      // enforceStrictInlineCsp do not combine.
      Array.from(inlinedScripts).forEach((script) => expect(script.hasAttribute('nonce')).toBe(false));
      expect(getInjectedCspDirectives(strict).get('script-src')).not.toContain("'unsafe-inline'");
    });

    it('cannot inline sidecars that the runtime injects after load', async () => {
      const mockFetch = jest.fn().mockResolvedValue({
        ok: true,
        text: () => Promise.resolve(readSample('assets/app.js')),
      });
      Object.defineProperty(globalThis, 'fetch', { configurable: true, value: mockFetch });

      const inlined = await inlineAllowedExternalScripts(page2Html, PAGE2_URL, HOST_PAGE_URL, { enabled: true });

      expect(mockFetch).toHaveBeenCalledTimes(1);
      expect(mockFetch).toHaveBeenCalledWith(`${LIBRARY_FOLDER_URL}assets/app.js`, expect.anything());
      expect(parse(inlined).querySelectorAll('script[data-uhv-inlined-external-script]')).toHaveLength(1);
      // The sidecar URL only exists at runtime (HfxData.load), so it stays a
      // network script load that the host CSP and SharePoint headers govern.
      expect(inlined).not.toContain('HfxData.register("probes"');
    });
  });

  describe('manifest.json', () => {
    it('lists every page, asset and sidecar in the sample bundle', () => {
      const manifest = JSON.parse(readSample('manifest.json')) as {
        entry: string;
        pages: Array<{ path: string; sections: string[] }>;
        assets: string[];
        datasets: Array<{ id: string; chunks: Array<{ path: string }> }>;
      };
      const listed = [
        ...manifest.pages.map((page) => page.path),
        ...manifest.assets,
        ...manifest.datasets.reduce<string[]>(
          (all, dataset) => all.concat(dataset.chunks.map((chunk) => chunk.path)),
          [],
        ),
      ];

      expect(manifest.entry).toBe('index.html');
      listed.forEach((relativePath) => expect(fs.existsSync(path.join(SAMPLE_ROOT, relativePath))).toBe(true));
      expect(manifest.pages.find((page) => page.path === 'page2.html')?.sections).toContain('section');
    });
  });
});
