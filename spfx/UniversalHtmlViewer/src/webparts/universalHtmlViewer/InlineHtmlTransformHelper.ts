import { appendAdditionalCspHostSources } from './InlineCspSourceHelper';
import {
  getCanonicalHostPageUrl,
  rewriteInlineNavigationAnchorHrefs,
} from './InlineAnchorRewriteHelper';
import { getInlineNavigationBridgeScript } from './InlineNavigationBridgeScript';

export interface IPrepareInlineHtmlForSrcDocOptions {
  enforceStrictInlineCsp?: boolean;
  additionalScriptSrcHosts?: string[];
  additionalStyleSrcHosts?: string[];
  additionalImageSrcHosts?: string[];
  rewriteInlineAnchorHrefs?: boolean;
  rewriteInlineAnchorAllowedFileExtensions?: string[];
  rewriteInlineAnchorAllowedPathPrefixes?: string[];
  rewriteInlineAnchorDeepLinkQueryParamName?: string;
  rewriteInlineAnchorPreservedHostQueryParamNames?: string[];
}

export function prepareInlineHtmlForSrcDoc(
  html: string,
  baseUrlForRelativeLinks: string,
  pageUrl: string,
  options?: IPrepareInlineHtmlForSrcDocOptions,
): string {
  return prepareInlineHtmlForFrameDocument(
    html,
    baseUrlForRelativeLinks,
    pageUrl,
    options,
    true,
  );
}

export function prepareInlineHtmlForBlobUrl(
  html: string,
  baseUrlForRelativeLinks: string,
  pageUrl: string,
  options?: Pick<
    IPrepareInlineHtmlForSrcDocOptions,
    | 'rewriteInlineAnchorHrefs'
    | 'rewriteInlineAnchorAllowedFileExtensions'
    | 'rewriteInlineAnchorAllowedPathPrefixes'
    | 'rewriteInlineAnchorDeepLinkQueryParamName'
    | 'rewriteInlineAnchorPreservedHostQueryParamNames'
  >,
): string {
  return prepareInlineHtmlForFrameDocument(
    html,
    baseUrlForRelativeLinks,
    pageUrl,
    options,
    false,
  );
}

function prepareInlineHtmlForFrameDocument(
  html: string,
  baseUrlForRelativeLinks: string,
  pageUrl: string,
  options: IPrepareInlineHtmlForSrcDocOptions | undefined,
  injectDefaultContentSecurityPolicy: boolean,
): string {
  const enforceStrictInlineCsp = options?.enforceStrictInlineCsp === true;
  const pageScriptNonce = tryGetCurrentPageScriptNonce();
  const htmlWithNeutralizedNestedFrames = neutralizeNestedIframeSources(html);
  const htmlWithHostDeepLinkedAnchors =
    options?.rewriteInlineAnchorHrefs === true
      ? rewriteInlineNavigationAnchorHrefs(
          htmlWithNeutralizedNestedFrames,
          baseUrlForRelativeLinks,
          pageUrl,
          {
            allowedFileExtensions: options?.rewriteInlineAnchorAllowedFileExtensions,
            allowedPathPrefixes: options?.rewriteInlineAnchorAllowedPathPrefixes,
            preservedHostQueryParamNames:
              options?.rewriteInlineAnchorPreservedHostQueryParamNames,
            deepLinkQueryParamName:
              options?.rewriteInlineAnchorDeepLinkQueryParamName,
          },
        )
      : htmlWithNeutralizedNestedFrames;
  const shouldInjectCompatibilityShim = !/data-uhv-history-compat=/i.test(
    htmlWithHostDeepLinkedAnchors,
  );
  const shouldInjectInlineNavigationBridge = !/data-uhv-inline-nav-bridge=/i.test(
    htmlWithHostDeepLinkedAnchors,
  );
  const inlineScriptNonce =
    pageScriptNonce ||
    (enforceStrictInlineCsp
      ? createStrictInlineHistoryShimNonce()
      : undefined);
  const htmlWithNonceStampedScripts = applyPageScriptNonceToInlineScripts(
    htmlWithHostDeepLinkedAnchors,
    inlineScriptNonce,
  );
  const srcDocCspTag =
    injectDefaultContentSecurityPolicy && !hasContentSecurityPolicyMetaTag(htmlWithNonceStampedScripts)
      ? `<meta data-uhv-inline-csp="1" http-equiv="Content-Security-Policy" content="${escapeHtmlAttribute(
          getDefaultSrcDocContentSecurityPolicy(
            pageUrl,
            baseUrlForRelativeLinks,
            enforceStrictInlineCsp,
            inlineScriptNonce,
            options,
          ),
        )}">`
      : '';
  const baseHref = getAbsoluteUrlWithoutQuery(baseUrlForRelativeLinks, pageUrl);
  const baseTag = /<base\s+/i.test(htmlWithNonceStampedScripts)
    ? ''
    : `<base href="${escapeHtmlAttribute(baseHref)}">`;
  const compatibilityShimTag = shouldInjectCompatibilityShim
    ? `<script data-uhv-history-compat="1"${
        inlineScriptNonce
          ? ` nonce="${escapeHtmlAttribute(inlineScriptNonce)}"`
          : ''
      }>${getHistoryCompatibilityShimScript()}</script>`
    : '';
  const inlineNavigationBridgeTag = shouldInjectInlineNavigationBridge
    ? `<script data-uhv-inline-nav-bridge="1"${
        inlineScriptNonce
          ? ` nonce="${escapeHtmlAttribute(inlineScriptNonce)}"`
          : ''
      }>${getInlineNavigationBridgeScript(
        baseUrlForRelativeLinks,
        options?.rewriteInlineAnchorAllowedFileExtensions,
        options?.rewriteInlineAnchorAllowedPathPrefixes,
        getCanonicalHostPageUrl(
          pageUrl,
          options?.rewriteInlineAnchorPreservedHostQueryParamNames,
        ),
        options?.rewriteInlineAnchorDeepLinkQueryParamName,
        options?.rewriteInlineAnchorPreservedHostQueryParamNames,
        options?.rewriteInlineAnchorHrefs === true,
      )}</script>`
    : '';
  const headInjectedMarkup = `${srcDocCspTag}${compatibilityShimTag}${inlineNavigationBridgeTag}${baseTag}`;

  if (!headInjectedMarkup) {
    return htmlWithNonceStampedScripts;
  }

  if (/<head[\s>]/i.test(htmlWithNonceStampedScripts)) {
    return htmlWithNonceStampedScripts.replace(
      /<head([^>]*)>/i,
      `<head$1>${headInjectedMarkup}`,
    );
  }

  if (/<html[\s>]/i.test(htmlWithNonceStampedScripts)) {
    return htmlWithNonceStampedScripts.replace(
      /<html([^>]*)>/i,
      `<html$1><head>${headInjectedMarkup}</head>`,
    );
  }

  return `<head>${headInjectedMarkup}</head>${htmlWithNonceStampedScripts}`;
}

function tryGetCurrentPageScriptNonce(): string | undefined {
  if (typeof document === 'undefined') {
    return undefined;
  }

  const scripts = Array.from(document.scripts || []);
  for (let index = 0; index < scripts.length; index += 1) {
    const script = scripts[index];
    const scriptNonce = (script.nonce || script.getAttribute('nonce') || '').trim();
    if (scriptNonce) {
      return scriptNonce;
    }
  }

  return undefined;
}

function applyPageScriptNonceToInlineScripts(
  html: string,
  pageScriptNonce?: string,
): string {
  if (!html || !pageScriptNonce || typeof DOMParser === 'undefined' || !/<script[\s>]/i.test(html)) {
    return html;
  }

  try {
    const hasDoctype: boolean = /^\s*<!doctype/i.test(html);
    const parser = new DOMParser();
    const parsed = parser.parseFromString(html, 'text/html');
    if (!parsed || !parsed.documentElement) {
      return html;
    }

    const scripts = parsed.querySelectorAll('script');
    scripts.forEach((script) => {
      const hasSrc = (script.getAttribute('src') || '').trim().length > 0;
      const hasNonce = (script.getAttribute('nonce') || '').trim().length > 0;
      if (!hasSrc && !hasNonce) {
        script.setAttribute('nonce', pageScriptNonce);
      }
    });

    const rebuiltHtml = parsed.documentElement.outerHTML;
    if (!rebuiltHtml) {
      return html;
    }

    return hasDoctype ? `<!DOCTYPE html>${rebuiltHtml}` : rebuiltHtml;
  } catch {
    return html;
  }
}

function hasContentSecurityPolicyMetaTag(html: string): boolean {
  return /<meta[^>]*http-equiv\s*=\s*["']?\s*content-security-policy\b/i.test(html);
}

function getAbsoluteUrlWithoutQuery(url: string, pageUrl: string): string {
  try {
    const current = new URL(pageUrl);
    const absolute = url.startsWith('/') ? new URL(url, current.origin) : new URL(url);
    absolute.search = '';
    absolute.hash = '';
    return absolute.toString();
  } catch {
    return url;
  }
}

function getDefaultSrcDocContentSecurityPolicy(
  pageUrl: string,
  baseUrlForRelativeLinks: string,
  enforceStrictInlineCsp: boolean,
  historyCompatibilityShimNonce?: string,
  options?: IPrepareInlineHtmlForSrcDocOptions,
): string {
  const allowedOrigins = getAllowedOriginsForInlineSrcDoc(pageUrl, baseUrlForRelativeLinks);
  const scriptSourcesBase = enforceStrictInlineCsp
    ? `${allowedOrigins} blob:${
        historyCompatibilityShimNonce
          ? ` 'nonce-${historyCompatibilityShimNonce}'`
          : ''
      }`
    : `${allowedOrigins} blob: 'unsafe-inline' 'unsafe-eval'`;
  const scriptSources = appendAdditionalCspHostSources(
    scriptSourcesBase,
    options?.additionalScriptSrcHosts,
  );
  const styleSources = appendAdditionalCspHostSources(
    `${allowedOrigins} data: 'unsafe-inline'`,
    options?.additionalStyleSrcHosts,
  );
  const frameSources = `${allowedOrigins} blob:`;
  const mediaSources = appendAdditionalCspHostSources(
    `${allowedOrigins} data: blob:`,
    options?.additionalImageSrcHosts,
  );

  return [
    `default-src ${allowedOrigins} data: blob:`,
    `script-src ${scriptSources}`,
    `style-src ${styleSources}`,
    `img-src ${mediaSources}`,
    `font-src ${styleSources}`,
    `connect-src ${allowedOrigins}`,
    `frame-src ${frameSources}`,
    `child-src ${frameSources}`,
    `media-src ${mediaSources}`,
    "object-src 'none'",
    `base-uri ${allowedOrigins}`,
    `form-action ${allowedOrigins}`,
  ].join('; ');
}

function createStrictInlineHistoryShimNonce(): string {
  const targetLength = 24;
  const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  const randomIndexes = getRandomUint8Array(targetLength);
  if (!randomIndexes) {
    return buildPseudoRandomNonce(targetLength, alphabet);
  }

  let nonce = '';
  for (let index = 0; index < randomIndexes.length; index += 1) {
    const charIndex = randomIndexes[index] % alphabet.length;
    nonce += alphabet.charAt(charIndex);
  }

  return nonce;
}

function getRandomUint8Array(length: number): Uint8Array | undefined {
  try {
    const cryptoApi =
      typeof globalThis !== 'undefined' && globalThis.crypto
        ? globalThis.crypto
        : undefined;
    if (!cryptoApi || typeof cryptoApi.getRandomValues !== 'function') {
      return undefined;
    }

    const values = new Uint8Array(length);
    cryptoApi.getRandomValues(values);
    return values;
  } catch {
    return undefined;
  }
}

function buildPseudoRandomNonce(length: number, alphabet: string): string {
  let nonce = '';
  for (let index = 0; index < length; index += 1) {
    const randomIndex = Math.floor(Math.random() * alphabet.length);
    nonce += alphabet.charAt(randomIndex);
  }

  return nonce;
}

function getAllowedOriginsForInlineSrcDoc(pageUrl: string, baseUrlForRelativeLinks: string): string {
  const sourceSet = new Set<string>(["'self'"]);
  const pageOrigin = tryGetOrigin(pageUrl);
  if (pageOrigin) {
    sourceSet.add(pageOrigin);
  }

  const baseOrigin = tryGetOrigin(baseUrlForRelativeLinks, pageOrigin);
  if (baseOrigin) {
    sourceSet.add(baseOrigin);
  }

  return Array.from(sourceSet.values()).join(' ');
}

function tryGetOrigin(value: string, fallbackOrigin?: string): string | undefined {
  const normalized = (value || '').trim();
  if (!normalized) {
    return undefined;
  }

  try {
    const parsedUrl = fallbackOrigin ? new URL(normalized, fallbackOrigin) : new URL(normalized);
    return parsedUrl.origin;
  } catch {
    return undefined;
  }
}

function escapeHtmlAttribute(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function neutralizeNestedIframeSources(html: string): string {
  if (!html || !/<iframe[\s\S]+src\s*=/i.test(html)) {
    return html;
  }

  if (typeof DOMParser === 'undefined') {
    return html;
  }

  try {
    const hasDoctype: boolean = /^\s*<!doctype/i.test(html);
    const parser = new DOMParser();
    const parsed = parser.parseFromString(html, 'text/html');
    if (!parsed || !parsed.documentElement) {
      return html;
    }

    const frames = parsed.querySelectorAll('iframe[src]');
    frames.forEach((frame) => {
      const rawSrc = (frame.getAttribute('src') || '').trim();
      if (!shouldNeutralizeNestedIframeSource(rawSrc)) {
        return;
      }

      frame.setAttribute('data-uhv-inline-src', rawSrc);
      frame.setAttribute('src', 'about:blank');
    });

    const rebuiltHtml = parsed.documentElement.outerHTML;
    if (!rebuiltHtml) {
      return html;
    }

    return hasDoctype ? `<!DOCTYPE html>${rebuiltHtml}` : rebuiltHtml;
  } catch {
    return html;
  }
}

function shouldNeutralizeNestedIframeSource(rawSrc: string): boolean {
  const normalized = (rawSrc || '').trim().toLowerCase();
  const javaScriptScheme = `java${'script:'}`;
  if (!normalized || normalized.startsWith('#') || normalized.startsWith('about:')) {
    return false;
  }

  if (
    normalized.startsWith(javaScriptScheme) ||
    normalized.startsWith('data:') ||
    normalized.startsWith('mailto:') ||
    normalized.startsWith('tel:')
  ) {
    return false;
  }

  try {
    const parsed = new URL(rawSrc, 'https://placeholder.local/');
    const extension = getPathExtension(parsed.pathname);
    return extension === '.html' || extension === '.htm' || extension === '.aspx';
  } catch {
    return false;
  }
}

function getPathExtension(pathname: string): string {
  const normalizedPath = (pathname || '').toLowerCase();
  const lastSlash = normalizedPath.lastIndexOf('/');
  const lastDot = normalizedPath.lastIndexOf('.');
  if (lastDot === -1 || lastDot < lastSlash) {
    return '';
  }

  return normalizedPath.substring(lastDot);
}

function getHistoryCompatibilityShimScript(): string {
  return [
    '(function(){',
    '  try {',
    "    var isSrcDoc = String(window.location && window.location.href || '').indexOf('about:srcdoc') === 0;",
    '    if (!isSrcDoc || !window.history) { return; }',
    '    var historyObject = window.history;',
    '    var originalPushState = typeof historyObject.pushState === "function"',
    '      ? historyObject.pushState.bind(historyObject)',
    '      : undefined;',
    '    var originalReplaceState = typeof historyObject.replaceState === "function"',
    '      ? historyObject.replaceState.bind(historyObject)',
    '      : undefined;',
    '    var isRecoverable = function(error) {',
    "      var name = error && error.name ? String(error.name) : '';",
    "      return name === 'SecurityError' || name === 'TypeError';",
    '    };',
    '    var canIgnoreHashStateUrl = function(url) {',
    "      if (typeof url !== 'string' || url.length === 0) { return false; }",
    "      var hashIndex = url.indexOf('#');",
    '      if (hashIndex < 0) { return false; }',
    '      var hashValue = url.substring(hashIndex);',
    '      return hashValue.length > 1;',
    '    };',
    '    var wrapState = function(originalMethod, methodName) {',
    '      if (!originalMethod) { return; }',
    '      try {',
    '        var wrapped = function(state, title, url) {',
    '          try {',
    '            return originalMethod(state, title, url);',
    '          } catch (error) {',
    '            if (isRecoverable(error) && canIgnoreHashStateUrl(typeof url === "string" ? url : "")) {',
    '              // A srcdoc base URL points at the SharePoint file. Calling',
    '              // location.replace("#fragment") here resolves against that',
    '              // base and downloads the raw HTML attachment. The report has',
    '              // already applied its UI state, so only the unsafe URL update',
    '              // is intentionally ignored.',
    '              return;',
    '            }',
    '            if (isRecoverable(error)) { return; }',
    '            throw error;',
    '          }',
    '        };',
    '        Object.defineProperty(historyObject, methodName, {',
    '          configurable: true,',
    '          writable: true,',
    '          value: wrapped',
    '        });',
    '      } catch (_overrideError) {',
    '        return;',
    '      }',
    '    };',
    '    wrapState(originalPushState, "pushState");',
    '    wrapState(originalReplaceState, "replaceState");',
    '  } catch (_error) {',
    '    return;',
    '  }',
    '})();',
  ].join('\n');
}
