export interface IPrepareInlineHtmlForSrcDocOptions {
  enforceStrictInlineCsp?: boolean;
}

export function prepareInlineHtmlForSrcDoc(
  html: string,
  baseUrlForRelativeLinks: string,
  pageUrl: string,
  options?: IPrepareInlineHtmlForSrcDocOptions,
): string {
  const enforceStrictInlineCsp = options?.enforceStrictInlineCsp === true;
  const htmlWithNeutralizedNestedFrames = neutralizeNestedIframeSources(html);
  const shouldInjectCompatibilityShim = !/data-uhv-history-compat=/i.test(
    htmlWithNeutralizedNestedFrames,
  );
  const shouldInjectInlineNavigationBridge = !/data-uhv-inline-nav-bridge=/i.test(
    htmlWithNeutralizedNestedFrames,
  );
  const inlineScriptNonce =
    enforceStrictInlineCsp && (shouldInjectCompatibilityShim || shouldInjectInlineNavigationBridge)
      ? createStrictInlineHistoryShimNonce()
      : undefined;
  const srcDocCspTag = hasContentSecurityPolicyMetaTag(htmlWithNeutralizedNestedFrames)
    ? ''
    : `<meta data-uhv-inline-csp="1" http-equiv="Content-Security-Policy" content="${escapeHtmlAttribute(
        getDefaultSrcDocContentSecurityPolicy(
          pageUrl,
          baseUrlForRelativeLinks,
          enforceStrictInlineCsp,
          inlineScriptNonce,
        ),
      )}">`;
  const baseHref = getAbsoluteUrlWithoutQuery(baseUrlForRelativeLinks, pageUrl);
  const baseTag = /<base\s+/i.test(htmlWithNeutralizedNestedFrames)
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
      }>${getInlineNavigationBridgeScript()}</script>`
    : '';
  const headInjectedMarkup = `${srcDocCspTag}${compatibilityShimTag}${inlineNavigationBridgeTag}${baseTag}`;

  if (!headInjectedMarkup) {
    return htmlWithNeutralizedNestedFrames;
  }

  if (/<head[\s>]/i.test(htmlWithNeutralizedNestedFrames)) {
    return htmlWithNeutralizedNestedFrames.replace(
      /<head([^>]*)>/i,
      `<head$1>${headInjectedMarkup}`,
    );
  }

  if (/<html[\s>]/i.test(htmlWithNeutralizedNestedFrames)) {
    return htmlWithNeutralizedNestedFrames.replace(
      /<html([^>]*)>/i,
      `<html$1><head>${headInjectedMarkup}</head>`,
    );
  }

  return `<head>${headInjectedMarkup}</head>${htmlWithNeutralizedNestedFrames}`;
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
): string {
  const allowedOrigins = getAllowedOriginsForInlineSrcDoc(pageUrl, baseUrlForRelativeLinks);
  const scriptSources = enforceStrictInlineCsp
    ? `${allowedOrigins} blob:${
        historyCompatibilityShimNonce
          ? ` 'nonce-${historyCompatibilityShimNonce}'`
          : ''
      }`
    : `${allowedOrigins} blob: 'unsafe-inline' 'unsafe-eval'`;
  const styleSources = `${allowedOrigins} data: 'unsafe-inline'`;
  const frameSources = `${allowedOrigins} blob:`;
  const mediaSources = `${allowedOrigins} data: blob:`;

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
    '    var trySetHash = function(url, replace) {',
    "      if (typeof url !== 'string' || url.length === 0) { return false; }",
    "      var hashIndex = url.indexOf('#');",
    '      if (hashIndex < 0) { return false; }',
    '      var hashValue = url.substring(hashIndex);',
    '      if (!hashValue) { return false; }',
    '      try {',
    '        if (replace && typeof window.location.replace === "function") {',
    '          window.location.replace(hashValue);',
    '        } else {',
    '          window.location.hash = hashValue;',
    '        }',
    '        return true;',
    '      } catch (_hashError) {',
    '        return false;',
    '      }',
    '    };',
    '    var wrapState = function(originalMethod, methodName, replace) {',
    '      if (!originalMethod) { return; }',
    '      try {',
    '        var wrapped = function(state, title, url) {',
    '          try {',
    '            return originalMethod(state, title, url);',
    '          } catch (error) {',
    '            if (isRecoverable(error) && trySetHash(typeof url === "string" ? url : "", replace)) {',
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
    '    wrapState(originalPushState, "pushState", false);',
    '    wrapState(originalReplaceState, "replaceState", true);',
    '  } catch (_error) {',
    '    return;',
    '  }',
    '})();',
  ].join('\n');
}

function getInlineNavigationBridgeScript(): string {
  return [
    '(function(){',
    '  try {',
    '    if (window.__uhvInlineNavBridgeInstalled) { return; }',
    '    window.__uhvInlineNavBridgeInstalled = true;',
    "    var interceptedEvents = ['pointerdown', 'mousedown', 'click'];",
    "    var blockedProtocols = ['javascript:', 'data:', 'mailto:', 'tel:'];",
    '    var lastTargetUrl = "";',
    '    var lastTargetAt = 0;',
    '    var isPrimaryEvent = function(event) {',
    '      if (!event) { return false; }',
    "      if (typeof event.button === 'number' && event.button !== 0) { return false; }",
    '      return !event.metaKey && !event.ctrlKey && !event.shiftKey && !event.altKey;',
    '    };',
    '    var getTargetElement = function(event) {',
    '      if (!event) { return null; }',
    '      var target = event.target;',
    '      if (!target) { return null; }',
    "      if (typeof Element !== 'undefined' && target instanceof Element) {",
    '        return target;',
    '      }',
    "      if (typeof Node !== 'undefined' && target instanceof Node) {",
    "        return target.parentElement || null;",
    '      }',
    '      return null;',
    '    };',
    '    var getComposedPathElements = function(event) {',
    "      if (!event || typeof event.composedPath !== 'function') { return []; }",
    '      var path;',
    '      try {',
    '        path = event.composedPath();',
    '      } catch (_error) {',
    '        return [];',
    '      }',
    '      if (!Array.isArray(path) || path.length === 0) { return []; }',
    '      var elements = [];',
    '      for (var index = 0; index < path.length; index += 1) {',
    '        var entry = path[index];',
    "        if (typeof Element !== 'undefined' && entry instanceof Element) {",
    '          elements.push(entry);',
    '          continue;',
    '        }',
    "        if (typeof Node !== 'undefined' && entry instanceof Node && entry.parentElement) {",
    '          elements.push(entry.parentElement);',
    '        }',
    '      }',
    '      return elements;',
    '    };',
    '    var findUniqueDescendantAnchor = function(element) {',
    "      if (!element || typeof element.querySelectorAll !== 'function') { return null; }",
    "      var anchors = element.querySelectorAll('a[href], a[xlink\\\\:href], a');",
    '      if (!anchors || anchors.length !== 1) { return null; }',
    '      var anchor = anchors[0];',
    "      return anchor && String(anchor.tagName || '').toLowerCase() === 'a' ? anchor : null;",
    '    };',
    '    var resolveAnchor = function(event) {',
    '      var targetElement = getTargetElement(event);',
    '      if (!targetElement) { return null; }',
    "      if (typeof targetElement.closest === 'function') {",
    "        var closestAnchor = targetElement.closest('a');",
    "        if (closestAnchor && String(closestAnchor.tagName || '').toLowerCase() === 'a') {",
    '          return closestAnchor;',
    '        }',
    '      }',
    '      var uniqueDescendant = findUniqueDescendantAnchor(targetElement);',
    '      if (uniqueDescendant) {',
    '        return uniqueDescendant;',
    '      }',
    '      var pathElements = getComposedPathElements(event);',
    '      for (var pathIndex = 0; pathIndex < pathElements.length; pathIndex += 1) {',
    '        var pathElement = pathElements[pathIndex];',
    '        if (!pathElement) { continue; }',
    "        if (typeof pathElement.closest === 'function') {",
    "          var pathAnchor = pathElement.closest('a');",
    "          if (pathAnchor && String(pathAnchor.tagName || '').toLowerCase() === 'a') {",
    '            return pathAnchor;',
    '          }',
    '        }',
    '        var pathDescendant = findUniqueDescendantAnchor(pathElement);',
    '        if (pathDescendant) {',
    '          return pathDescendant;',
    '        }',
    '      }',
    "      var forcedContainer = typeof targetElement.closest === 'function'",
    "        ? targetElement.closest('.fc-event-forced-url')",
    '        : null;',
    '      if (forcedContainer) {',
    "        var forcedAnchor = forcedContainer.querySelector('a[href], a[xlink\\\\:href], a');",
    "        if (forcedAnchor && String(forcedAnchor.tagName || '').toLowerCase() === 'a') {",
    '          return forcedAnchor;',
    '        }',
    '      }',
    '      return null;',
    '    };',
    '    var getAnchorHref = function(anchor) {',
    "      var hrefAttr = String((anchor && anchor.getAttribute && anchor.getAttribute('href')) || '').trim();",
    '      if (hrefAttr) { return hrefAttr; }',
    '      try {',
    "        if (anchor && typeof anchor.getAttributeNS === 'function') {",
    "          var namespacedHref = String(anchor.getAttributeNS('http://www.w3.org/1999/xlink', 'href') || '').trim();",
    '          if (namespacedHref) { return namespacedHref; }',
    '        }',
    '      } catch (_error) {',
    '        return "";',
    '      }',
    "      var hrefProp = anchor ? anchor.href : '';",
    "      if (typeof hrefProp === 'string') {",
    '        return hrefProp.trim();',
    '      }',
    "      if (hrefProp && typeof hrefProp === 'object') {",
    "        var baseVal = String(hrefProp.baseVal || '').trim();",
    '        if (baseVal) { return baseVal; }',
    "        var animVal = String(hrefProp.animVal || '').trim();",
    '        if (animVal) { return animVal; }',
    '      }',
    '      return "";',
    '    };',
    '    var hasBlockedProtocol = function(href) {',
    "      var normalized = String(href || '').trim().toLowerCase();",
    "      for (var index = 0; index < blockedProtocols.length; index += 1) {",
    '        if (normalized.indexOf(blockedProtocols[index]) === 0) {',
    '          return true;',
    '        }',
    '      }',
    '      return false;',
    '    };',
    '    var toAbsoluteUrl = function(rawHref) {',
    "      var base = String(document.baseURI || '').trim();",
    '      if (!base) {',
    "        base = String((window.location && window.location.href) || '').trim();",
    '      }',
    '      try {',
    '        return new URL(rawHref, base || undefined).toString();',
    '      } catch (_error) {',
    '        return String(rawHref || "").trim();',
    '      }',
    '    };',
    '    var suppress = function(event) {',
    "      if (!event || typeof event.preventDefault !== 'function') { return; }",
    '      event.preventDefault();',
    "      if (typeof event.stopPropagation === 'function') {",
    '        event.stopPropagation();',
    '      }',
    "      if (typeof event.stopImmediatePropagation === 'function') {",
    '        event.stopImmediatePropagation();',
    '      }',
    '      event.cancelBubble = true;',
    '      event.returnValue = false;',
    '    };',
    '    var emit = function(targetUrl, event) {',
    '      if (!targetUrl) { return; }',
    '      var now = Date.now();',
    '      if (lastTargetUrl === targetUrl && now - lastTargetAt < 500) {',
    '        suppress(event);',
    '        return;',
    '      }',
    '      lastTargetUrl = targetUrl;',
    '      lastTargetAt = now;',
    '      suppress(event);',
    '      try {',
    '        if (window.parent && window.parent !== window && typeof window.parent.postMessage === "function") {',
    "          window.parent.postMessage({ type: 'uhv-inline-nav', targetUrl: targetUrl }, '*');",
    '        }',
    '      } catch (_error) {',
    '        return;',
    '      }',
    '    };',
    '    var onInterceptedEvent = function(event) {',
    '      if (!isPrimaryEvent(event)) { return; }',
    '      var anchor = resolveAnchor(event);',
    '      if (!anchor) { return; }',
    '      var rawHref = getAnchorHref(anchor);',
    "      if (!rawHref || rawHref.charAt(0) === '#') { return; }",
    '      if (hasBlockedProtocol(rawHref)) { return; }',
    '      var absoluteTargetUrl = toAbsoluteUrl(rawHref);',
    "      if (!absoluteTargetUrl || absoluteTargetUrl.charAt(0) === '#') { return; }",
    '      emit(absoluteTargetUrl, event);',
    '    };',
    '    for (var eventIndex = 0; eventIndex < interceptedEvents.length; eventIndex += 1) {',
    '      var eventName = interceptedEvents[eventIndex];',
    '      document.addEventListener(eventName, onInterceptedEvent, true);',
    '    }',
    '  } catch (_error) {',
    '    return;',
    '  }',
    '})();',
  ].join('\n');
}
