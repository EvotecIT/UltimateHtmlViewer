import { getQueryStringParam } from './QueryStringHelper';

export type HtmlSourceMode = 'FullUrl' | 'BasePathAndRelativePath' | 'BasePathAndDashboardId';

export type HeightMode = 'Fixed' | 'Viewport' | 'Auto';

export type UrlSecurityMode = 'StrictTenant' | 'Allowlist' | 'AnyHttps';

export type CacheBusterMode = 'None' | 'Timestamp' | 'FileLastModified';

export interface BuildUrlParams {
  htmlSourceMode: HtmlSourceMode;
  fullUrl?: string;
  basePath?: string;
  relativePath?: string;
  dashboardId?: string;
  defaultFileName?: string;
  queryStringParamName?: string;
  pageUrl?: string;
}

export interface UrlValidationOptions {
  securityMode: UrlSecurityMode;
  currentPageUrl: string;
  allowedHosts?: string[];
  allowedPathPrefixes?: string[];
  allowedFileExtensions?: string[];
  allowHttp?: boolean;
}

/**
 * Builds the final iframe URL based on configured properties and the current page URL.
 *
 * @param params Url building parameters.
 * @returns The computed URL string, or null if it cannot be determined.
 */
export function buildFinalUrl(params: BuildUrlParams): string | undefined {
  const mode: HtmlSourceMode = params.htmlSourceMode || 'FullUrl';

  if (mode === 'FullUrl') {
    const normalizedFullUrl: string = (params.fullUrl || '').trim();
    return normalizedFullUrl || undefined;
  }

  const basePathNormalized: string = normalizeBasePath(params.basePath);

  if (!basePathNormalized) {
    return undefined;
  }

  if (mode === 'BasePathAndRelativePath') {
    const relativePathNormalized: string = normalizeRelativePath(params.relativePath);

    if (!relativePathNormalized) {
      return undefined;
    }

    return `${basePathNormalized}${relativePathNormalized}`;
  }

  const queryParamName: string = (params.queryStringParamName || '').trim() || 'dashboard';
  const defaultFileName: string = (params.defaultFileName || '').trim() || 'index.html';

  const dashboardIdFromQuery: string | undefined = params.pageUrl
    ? getQueryStringParam(params.pageUrl, queryParamName)
    : undefined;

  const effectiveDashboardId: string = (dashboardIdFromQuery || params.dashboardId || '').trim();

  if (!effectiveDashboardId) {
    return undefined;
  }

  return `${basePathNormalized}${effectiveDashboardId}/${defaultFileName}`;
}

/**
 * Validates whether the provided URL is allowed to be used inside the iframe.
 *
 * Rules:
 * - Must be an https/http URL on the same host as currentPageUrl, OR
 * - A site-relative path starting with '/'.
 *
 * @param url The URL to validate.
 * @param currentPageUrl The current page URL used to determine the allowed host.
 */
export function isUrlAllowed(
  url: string | undefined,
  currentPageUrlOrOptions: string | UrlValidationOptions,
): boolean {
  if (!url) {
    return false;
  }

  const trimmedUrl: string = url.trim();

  if (!trimmedUrl) {
    return false;
  }

  const lowerUrl: string = trimmedUrl.toLowerCase();

  const blockedSchemes: string[] = ['javascript', 'data', 'vbscript'];
  if (blockedSchemes.some((scheme) => lowerUrl.startsWith(`${scheme}:`))) {
    return false;
  }

  if (trimmedUrl.startsWith('//') || trimmedUrl.startsWith('\\\\')) {
    return false;
  }

  const options: UrlValidationOptions =
    typeof currentPageUrlOrOptions === 'string'
      ? {
          securityMode: 'StrictTenant',
          currentPageUrl: currentPageUrlOrOptions,
        }
      : currentPageUrlOrOptions;

  if (trimmedUrl.startsWith('/')) {
    const pathOnly: string = stripQueryAndHash(trimmedUrl);
    return isPathAllowed(
      pathOnly,
      options.allowedPathPrefixes,
      options.allowedFileExtensions,
    );
  }

  if (lowerUrl.startsWith('http://') || lowerUrl.startsWith('https://')) {
    try {
      const target: URL = new URL(trimmedUrl);
      const current: URL = new URL(options.currentPageUrl);
      const rawAbsolutePath: string = getRawAbsolutePath(trimmedUrl);
      if (rawAbsolutePath && hasDotSegments(normalizePath(rawAbsolutePath))) {
        return false;
      }
      const targetHost: string = target.hostname.toLowerCase();
      const currentHost: string = current.hostname.toLowerCase();

      if (target.protocol === 'http:' && !options.allowHttp) {
        return false;
      }

      if (options.securityMode === 'AnyHttps') {
        if (target.protocol !== 'https:' && !(options.allowHttp && target.protocol === 'http:')) {
          return false;
        }
        return isPathAllowed(
          target.pathname,
          options.allowedPathPrefixes,
          options.allowedFileExtensions,
        );
      }

      if (targetHost === currentHost) {
        return isPathAllowed(
          target.pathname,
          options.allowedPathPrefixes,
          options.allowedFileExtensions,
        );
      }

      if (options.securityMode === 'Allowlist') {
        if (isHostAllowed(targetHost, options.allowedHosts)) {
          return isPathAllowed(
            target.pathname,
            options.allowedPathPrefixes,
            options.allowedFileExtensions,
          );
        }
      }

      return false;
    } catch {
      return false;
    }
  }

  return false;
}

function normalizeBasePath(basePath?: string): string {
  const value: string = (basePath || '').trim();

  if (!value) {
    return '';
  }

  let normalized: string = value;

  if (!normalized.startsWith('/')) {
    normalized = `/${normalized}`;
  }

  if (!normalized.endsWith('/')) {
    normalized = `${normalized}/`;
  }

  return normalized;
}

function normalizeRelativePath(relativePath?: string): string {
  const value: string = (relativePath || '').trim();

  if (!value) {
    return '';
  }

  let normalized: string = value;

  while (normalized.startsWith('/')) {
    normalized = normalized.substring(1);
  }

  return normalized;
}

function isPathAllowed(
  pathname: string,
  allowedPrefixes?: string[],
  allowedExtensions?: string[],
): boolean {
  const normalizedPath: string = normalizePath(pathname);

  if (hasDotSegments(normalizedPath)) {
    return false;
  }

  if (allowedExtensions && allowedExtensions.length > 0) {
    if (!isExtensionAllowed(normalizedPath, allowedExtensions)) {
      return false;
    }
  }

  if (!allowedPrefixes || allowedPrefixes.length === 0) {
    return true;
  }

  const prefixes: string[] = allowedPrefixes
    .map((prefix) => normalizePath(prefix))
    .filter((prefix) => prefix.length > 0);

  if (prefixes.length === 0) {
    return true;
  }

  return prefixes.some((prefix) => {
    if (normalizedPath === prefix) {
      return true;
    }

    if (prefix.endsWith('/')) {
      return normalizedPath.startsWith(prefix);
    }

    return normalizedPath.startsWith(`${prefix}/`);
  });
}

function normalizePath(pathname: string): string {
  const value: string = (pathname || '').trim();

  if (!value) {
    return '';
  }

  let normalized: string = value;

  if (!normalized.startsWith('/')) {
    normalized = `/${normalized}`;
  }

  while (normalized.includes('//')) {
    normalized = normalized.replace(/\/{2,}/g, '/');
  }

  return normalized.toLowerCase();
}

function hasDotSegments(pathname: string): boolean {
  const segments = pathname.split('/').filter((segment) => segment.length > 0);
  return segments.some((segment) => {
    const decodedSegment: string = decodePathSegment(segment);
    return decodedSegment === '.' || decodedSegment === '..';
  });
}

function decodePathSegment(segment: string): string {
  try {
    return decodeURIComponent(segment);
  } catch {
    return segment;
  }
}

function stripQueryAndHash(value: string): string {
  const hashIndex: number = value.indexOf('#');
  const queryIndex: number = value.indexOf('?');

  if (hashIndex === -1 && queryIndex === -1) {
    return value;
  }

  const cutIndex =
    hashIndex === -1
      ? queryIndex
      : queryIndex === -1
        ? hashIndex
        : Math.min(hashIndex, queryIndex);

  return value.substring(0, cutIndex);
}

function getRawAbsolutePath(url: string): string {
  const matched = url.match(/^[a-zA-Z][a-zA-Z\d+\-.]*:\/\/[^/]+(.*)$/);
  if (!matched) {
    return '';
  }

  const remainder: string = matched[1] || '/';
  if (remainder.startsWith('?') || remainder.startsWith('#')) {
    return '/';
  }

  return stripQueryAndHash(remainder);
}

function isExtensionAllowed(pathname: string, allowedExtensions: string[]): boolean {
  const normalized: string = pathname.toLowerCase();
  if (normalized.endsWith('/')) {
    return false;
  }

  const lastSlash = normalized.lastIndexOf('/');
  const lastDot = normalized.lastIndexOf('.');
  if (lastDot === -1 || lastDot < lastSlash) {
    return false;
  }

  const extension: string = normalized.substring(lastDot);
  const normalizedAllowed = allowedExtensions.map((ext) =>
    ext.startsWith('.') ? ext.toLowerCase() : `.${ext.toLowerCase()}`,
  );
  return normalizedAllowed.includes(extension);
}

function isHostAllowed(hostname: string, allowedHosts?: string[]): boolean {
  if (!allowedHosts || allowedHosts.length === 0) {
    return false;
  }

  const normalizedHost: string = hostname.toLowerCase();
  return allowedHosts.some((entry) => {
    let normalizedEntry: string = (entry || '').trim().toLowerCase();
    if (!normalizedEntry) {
      return false;
    }

    if (normalizedEntry.startsWith('*.')) {
      normalizedEntry = normalizedEntry.substring(1);
    }

    if (normalizedEntry.startsWith('.')) {
      return (
        normalizedHost.endsWith(normalizedEntry) &&
        normalizedHost.length > normalizedEntry.length
      );
    }

    return normalizedHost === normalizedEntry;
  });
}
