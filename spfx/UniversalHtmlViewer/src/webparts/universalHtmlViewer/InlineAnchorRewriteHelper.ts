import { buildPageUrlWithInlineDeepLink } from './InlineDeepLinkHelper';
import { INLINE_NAVIGATION_ORIGINAL_HREF_ATTRIBUTE } from './InlineNavigationAttributes';

export function rewriteInlineNavigationAnchorHrefs(
  html: string,
  baseUrlForRelativeLinks: string,
  pageUrl: string,
): string {
  if (!html || !/<a[\s\S]+href\s*=/i.test(html)) {
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

    const page = new URL(pageUrl);
    const baseUrl = getAbsoluteUrlWithoutQuery(baseUrlForRelativeLinks, pageUrl);
    const anchors = parsed.querySelectorAll('a[href]');
    anchors.forEach((anchor) => {
      if (anchor.hasAttribute(INLINE_NAVIGATION_ORIGINAL_HREF_ATTRIBUTE)) {
        return;
      }

      const rawHref = (anchor.getAttribute('href') || '').trim();
      const targetUrl = resolveInlineAnchorTargetUrl(rawHref, baseUrl, page);
      if (!targetUrl) {
        return;
      }

      const hostDeepLinkUrl = buildPageUrlWithInlineDeepLink({
        pageUrl,
        targetUrl,
      });
      if (!hostDeepLinkUrl || hostDeepLinkUrl === rawHref) {
        return;
      }

      anchor.setAttribute(INLINE_NAVIGATION_ORIGINAL_HREF_ATTRIBUTE, targetUrl);
      anchor.setAttribute('href', hostDeepLinkUrl);
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

function resolveInlineAnchorTargetUrl(
  rawHref: string,
  baseUrl: string,
  pageUrl: URL,
): string | undefined {
  const normalizedHref = (rawHref || '').trim();
  if (!normalizedHref || normalizedHref.startsWith('#')) {
    return undefined;
  }

  const javaScriptScheme = `java${'script:'}`;
  const lowerHref = normalizedHref.toLowerCase();
  if (
    lowerHref.startsWith(javaScriptScheme) ||
    lowerHref.startsWith('data:') ||
    lowerHref.startsWith('mailto:') ||
    lowerHref.startsWith('tel:')
  ) {
    return undefined;
  }

  let target: URL;
  try {
    target = new URL(normalizedHref, baseUrl || pageUrl.toString());
  } catch {
    return undefined;
  }

  if (target.host.toLowerCase() !== pageUrl.host.toLowerCase()) {
    return undefined;
  }

  if (!isInsideBaseDirectory(target, baseUrl)) {
    return undefined;
  }

  if (isCurrentPageDeepLink(target, pageUrl)) {
    return undefined;
  }

  const extension = getPathExtension(target.pathname);
  if (extension !== '.html' && extension !== '.htm' && extension !== '.aspx') {
    return undefined;
  }

  return target.toString();
}

function isInsideBaseDirectory(target: URL, baseUrl: string): boolean {
  try {
    const base = new URL(baseUrl);
    let basePath = base.pathname || '/';
    if (!basePath.endsWith('/')) {
      const lastSlash = basePath.lastIndexOf('/');
      basePath = lastSlash < 0 ? '/' : basePath.substring(0, lastSlash + 1);
    }

    return normalizePath(target.pathname).startsWith(normalizePath(basePath));
  } catch {
    return true;
  }
}

function isCurrentPageDeepLink(target: URL, pageUrl: URL): boolean {
  return (
    target.host.toLowerCase() === pageUrl.host.toLowerCase() &&
    target.pathname.toLowerCase() === pageUrl.pathname.toLowerCase() &&
    target.searchParams.has('uhvPage')
  );
}

function normalizePath(value: string): string {
  const normalized = (value || '').replace(/\\/g, '/');
  try {
    return decodeURIComponent(normalized).toLowerCase();
  } catch {
    return normalized.toLowerCase();
  }
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

function getPathExtension(pathname: string): string {
  const normalizedPath = (pathname || '').toLowerCase();
  const lastSlash = normalizedPath.lastIndexOf('/');
  const lastDot = normalizedPath.lastIndexOf('.');
  if (lastDot === -1 || lastDot < lastSlash) {
    return '';
  }

  return normalizedPath.substring(lastDot);
}
