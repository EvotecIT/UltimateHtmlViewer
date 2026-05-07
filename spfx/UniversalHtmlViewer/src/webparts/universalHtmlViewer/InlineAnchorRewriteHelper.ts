import {
  buildPageUrlWithInlineDeepLink,
  MAX_DEEP_LINK_QUERY_VALUE_LENGTH,
} from './InlineDeepLinkHelper';
import {
  INLINE_NAVIGATION_ORIGINAL_HREF_ATTRIBUTE,
  INLINE_NAVIGATION_REWRITTEN_ATTRIBUTE,
} from './InlineNavigationAttributes';

export interface IRewriteInlineNavigationAnchorHrefsOptions {
  allowedFileExtensions?: string[];
  preservedHostQueryParamNames?: string[];
}

export function rewriteInlineNavigationAnchorHrefs(
  html: string,
  baseUrlForRelativeLinks: string,
  pageUrl: string,
  options?: IRewriteInlineNavigationAnchorHrefsOptions,
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
    const fallbackBaseUrl = getAbsoluteUrlWithoutQuery(baseUrlForRelativeLinks, pageUrl);
    const baseUrl = getDocumentBaseUrl(parsed, fallbackBaseUrl);
    const hostPageUrl = getCanonicalHostPageUrl(
      page,
      options?.preservedHostQueryParamNames,
    );
    const anchors = parsed.querySelectorAll('a[href]');
    anchors.forEach((anchor) => {
      if (anchor.getAttribute(INLINE_NAVIGATION_REWRITTEN_ATTRIBUTE) === '1') {
        return;
      }

      if (anchor.hasAttribute(INLINE_NAVIGATION_ORIGINAL_HREF_ATTRIBUTE)) {
        anchor.removeAttribute(INLINE_NAVIGATION_ORIGINAL_HREF_ATTRIBUTE);
      }

      const rawHref = (anchor.getAttribute('href') || '').trim();
      const targetUrl = resolveInlineAnchorTargetUrl(
        rawHref,
        baseUrl,
        page,
        options?.allowedFileExtensions,
      );
      if (!targetUrl) {
        return;
      }

      if (getDeepLinkQueryValueLength(targetUrl, page) > MAX_DEEP_LINK_QUERY_VALUE_LENGTH) {
        return;
      }

      const hostDeepLinkUrl = buildPageUrlWithInlineDeepLink({
        pageUrl: hostPageUrl,
        targetUrl,
      });
      if (!hostDeepLinkUrl || hostDeepLinkUrl === rawHref) {
        return;
      }

      anchor.setAttribute(INLINE_NAVIGATION_ORIGINAL_HREF_ATTRIBUTE, targetUrl);
      anchor.setAttribute(INLINE_NAVIGATION_REWRITTEN_ATTRIBUTE, '1');
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

function getDocumentBaseUrl(parsed: Document, fallbackBaseUrl: string): string {
  const rawBaseHref = (parsed.querySelector('base[href]')?.getAttribute('href') || '').trim();
  if (!rawBaseHref) {
    return fallbackBaseUrl;
  }

  try {
    return new URL(rawBaseHref, fallbackBaseUrl).toString();
  } catch {
    return fallbackBaseUrl;
  }
}

function getCanonicalHostPageUrl(
  pageUrl: URL,
  preservedQueryParamNames?: string[],
): string {
  const canonical = new URL(pageUrl.toString());
  const preservedValues = new Map<string, string[]>();
  (preservedQueryParamNames || [])
    .map((entry) => (entry || '').trim())
    .filter((entry) => entry.length > 0)
    .forEach((paramName) => {
      const values = pageUrl.searchParams.getAll(paramName);
      if (values.length > 0) {
        preservedValues.set(paramName, values);
      }
    });

  canonical.search = '';
  preservedValues.forEach((values, paramName) => {
    values.forEach((value) => {
      canonical.searchParams.append(paramName, value);
    });
  });
  canonical.hash = '';
  return canonical.toString();
}

function resolveInlineAnchorTargetUrl(
  rawHref: string,
  baseUrl: string,
  pageUrl: URL,
  allowedFileExtensions?: string[],
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
  if (!isExtensionAllowed(extension, allowedFileExtensions)) {
    return undefined;
  }

  return target.toString();
}

function getDeepLinkQueryValueLength(targetUrl: string, pageUrl: URL): number {
  try {
    const target = new URL(targetUrl);
    if (target.host.toLowerCase() === pageUrl.host.toLowerCase()) {
      return `${target.pathname}${target.search}${target.hash}`.length;
    }

    return target.toString().length;
  } catch {
    return targetUrl.length;
  }
}

function isExtensionAllowed(extension: string, allowedFileExtensions?: string[]): boolean {
  if (!extension) {
    return false;
  }

  const normalizedAllowed = (allowedFileExtensions || [])
    .map((entry) => (entry || '').trim().toLowerCase())
    .filter((entry) => entry.length > 0)
    .map((entry) => (entry.startsWith('.') ? entry : `.${entry}`));

  if (normalizedAllowed.length > 0) {
    return normalizedAllowed.includes(extension);
  }

  return extension === '.html' || extension === '.htm' || extension === '.aspx';
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
