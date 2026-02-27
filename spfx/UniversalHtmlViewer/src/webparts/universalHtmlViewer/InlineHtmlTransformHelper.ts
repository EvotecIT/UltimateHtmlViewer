export interface IPrepareInlineHtmlForSrcDocOptions {
  enforceStrictInlineCsp?: boolean;
}

export function prepareInlineHtmlForSrcDoc(
  html: string,
  baseUrlForRelativeLinks: string,
  pageUrl: string,
  options?: IPrepareInlineHtmlForSrcDocOptions,
): string {
  const htmlWithNeutralizedNestedFrames = neutralizeNestedIframeSources(html);
  const srcDocCspTag = hasContentSecurityPolicyMetaTag(htmlWithNeutralizedNestedFrames)
    ? ''
    : `<meta data-uhv-inline-csp="1" http-equiv="Content-Security-Policy" content="${escapeHtmlAttribute(
        getDefaultSrcDocContentSecurityPolicy(
          pageUrl,
          baseUrlForRelativeLinks,
          options?.enforceStrictInlineCsp === true,
        ),
      )}">`;
  const baseHref = getAbsoluteUrlWithoutQuery(baseUrlForRelativeLinks, pageUrl);
  const baseTag = /<base\s+/i.test(htmlWithNeutralizedNestedFrames)
    ? ''
    : `<base href="${escapeHtmlAttribute(baseHref)}">`;
  const compatibilityShimTag = /data-uhv-history-compat=/i.test(
    htmlWithNeutralizedNestedFrames,
  )
    ? ''
    : `<script data-uhv-history-compat="1">${getHistoryCompatibilityShimScript()}</script>`;
  const headInjectedMarkup = `${srcDocCspTag}${compatibilityShimTag}${baseTag}`;

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
): string {
  const allowedOrigins = getAllowedOriginsForInlineSrcDoc(pageUrl, baseUrlForRelativeLinks);
  const scriptSources = enforceStrictInlineCsp
    ? `${allowedOrigins} blob:`
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
