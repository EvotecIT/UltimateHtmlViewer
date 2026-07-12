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
  allowedPathPrefixes?: string[];
  preservedHostQueryParamNames?: string[];
  deepLinkQueryParamName?: string;
}

export const SHAREPOINT_TRANSIENT_HOST_QUERY_PARAM_NAMES: string[] = [
  'ct',
  'e',
  'isspofile',
  'or',
  'wdlor',
  'wdorigin',
  'web',
];

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
      rewriteInlineNavigationAnchorElement(
        anchor,
        baseUrl,
        hostPageUrl,
        options,
      );
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

export function rewriteInlineNavigationAnchorElement(
  anchor: Element,
  baseUrlForRelativeLinks: string,
  pageUrl: string,
  options?: IRewriteInlineNavigationAnchorHrefsOptions,
): boolean {
  if (
    !anchor ||
    anchor.tagName.toLowerCase() !== 'a' ||
    anchor.getAttribute(INLINE_NAVIGATION_REWRITTEN_ATTRIBUTE) === '1'
  ) {
    return false;
  }

  if (anchor.hasAttribute(INLINE_NAVIGATION_ORIGINAL_HREF_ATTRIBUTE)) {
    anchor.removeAttribute(INLINE_NAVIGATION_ORIGINAL_HREF_ATTRIBUTE);
  }

  let page: URL;
  try {
    page = new URL(pageUrl);
  } catch {
    return false;
  }

  const rawHref = (anchor.getAttribute('href') || '').trim();
  const targetUrl = resolveInlineAnchorTargetUrl(
    rawHref,
    baseUrlForRelativeLinks,
    page,
    options?.allowedFileExtensions,
    options?.allowedPathPrefixes,
  );
  if (!targetUrl) {
    return false;
  }

  if (getDeepLinkQueryValueLength(targetUrl, page) > MAX_DEEP_LINK_QUERY_VALUE_LENGTH) {
    return false;
  }

  const hostPageUrl = getCanonicalHostPageUrl(
    page,
    options?.preservedHostQueryParamNames,
  );
  const hostDeepLinkUrl = buildPageUrlWithInlineDeepLink({
    pageUrl: hostPageUrl,
    targetUrl,
    queryParamName: options?.deepLinkQueryParamName,
  });
  if (!hostDeepLinkUrl || hostDeepLinkUrl === rawHref) {
    return false;
  }

  // SharePoint commonly serves .html files as attachments. Eligible HTML links
  // are viewer navigation, even when a generated report added `download`.
  anchor.removeAttribute('download');
  anchor.setAttribute(INLINE_NAVIGATION_ORIGINAL_HREF_ATTRIBUTE, targetUrl);
  anchor.setAttribute(INLINE_NAVIGATION_REWRITTEN_ATTRIBUTE, '1');
  anchor.setAttribute('href', hostDeepLinkUrl);
  return true;
}

export function isInlineNavigationAnchorRewriteCurrent(
  anchor: Element,
  baseUrlForRelativeLinks: string,
  pageUrl: string,
  options?: IRewriteInlineNavigationAnchorHrefsOptions,
): boolean {
  const originalHref = (
    anchor.getAttribute(INLINE_NAVIGATION_ORIGINAL_HREF_ATTRIBUTE) || ''
  ).trim();
  if (
    anchor.getAttribute(INLINE_NAVIGATION_REWRITTEN_ATTRIBUTE) !== '1' ||
    !originalHref
  ) {
    return false;
  }

  const expectedAnchor = anchor.cloneNode(false) as Element;
  expectedAnchor.removeAttribute(INLINE_NAVIGATION_ORIGINAL_HREF_ATTRIBUTE);
  expectedAnchor.removeAttribute(INLINE_NAVIGATION_REWRITTEN_ATTRIBUTE);
  expectedAnchor.setAttribute('href', originalHref);
  if (
    !rewriteInlineNavigationAnchorElement(
      expectedAnchor,
      baseUrlForRelativeLinks,
      pageUrl,
      options,
    )
  ) {
    return false;
  }

  return expectedAnchor.getAttribute('href') === anchor.getAttribute('href');
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

export function getCanonicalHostPageUrl(
  pageUrl: URL | string,
  preservedQueryParamNames?: string[],
): string {
  const sourcePageUrl =
    typeof pageUrl === 'string' ? new URL(pageUrl) : pageUrl;
  const canonical = new URL(sourcePageUrl.toString());
  const explicitlyPreservedNames = new Set(
    (preservedQueryParamNames || [])
    .map((entry) => (entry || '').trim())
    .filter((entry) => entry.length > 0)
    .map((entry) => entry.toLowerCase()),
  );

  canonical.search = '';
  sourcePageUrl.searchParams.forEach((value, paramName) => {
    const normalizedName = paramName.toLowerCase();
    if (
      explicitlyPreservedNames.has(normalizedName) ||
      !SHAREPOINT_TRANSIENT_HOST_QUERY_PARAM_NAMES.includes(normalizedName)
    ) {
      canonical.searchParams.append(paramName, value);
    }
  });
  canonical.hash = '';
  return canonical.toString();
}

function resolveInlineAnchorTargetUrl(
  rawHref: string,
  baseUrl: string,
  pageUrl: URL,
  allowedFileExtensions?: string[],
  allowedPathPrefixes?: string[],
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

  if (!isInsideAllowedPath(target, baseUrl, allowedPathPrefixes)) {
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

function isInsideAllowedPath(
  target: URL,
  baseUrl: string,
  allowedPathPrefixes?: string[],
): boolean {
  const normalizedPrefixes = (allowedPathPrefixes || [])
    .map((prefix) => normalizePath(prefix).replace(/\/?$/, '/'))
    .filter((prefix) => prefix.length > 1);
  if (normalizedPrefixes.length === 0) {
    return isInsideBaseDirectory(target, baseUrl);
  }

  const targetPath = normalizePath(target.pathname);
  return normalizedPrefixes.some(
    (prefix) => targetPath === prefix.slice(0, -1) || targetPath.startsWith(prefix),
  );
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
