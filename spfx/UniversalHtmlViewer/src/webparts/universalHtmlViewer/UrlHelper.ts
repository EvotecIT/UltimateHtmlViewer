import { getQueryStringParam } from './QueryStringHelper';

export type HtmlSourceMode = 'FullUrl' | 'BasePathAndRelativePath' | 'BasePathAndDashboardId';

export type HeightMode = 'Fixed' | 'Viewport';

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
    return isPathAllowed(trimmedUrl, options.allowedPathPrefixes);
  }

  if (lowerUrl.startsWith('http://') || lowerUrl.startsWith('https://')) {
    try {
      const target: URL = new URL(trimmedUrl);
      const current: URL = new URL(options.currentPageUrl);
      const targetHost: string = target.host.toLowerCase();
      const currentHost: string = current.host.toLowerCase();

      if (options.securityMode === 'AnyHttps') {
        if (target.protocol !== 'https:') {
          return false;
        }
        return isPathAllowed(target.pathname, options.allowedPathPrefixes);
      }

      if (targetHost === currentHost) {
        return isPathAllowed(target.pathname, options.allowedPathPrefixes);
      }

      if (options.securityMode === 'Allowlist') {
        const allowedHosts: string[] = (options.allowedHosts || []).map((host) =>
          host.toLowerCase(),
        );
        if (allowedHosts.includes(targetHost)) {
          return isPathAllowed(target.pathname, options.allowedPathPrefixes);
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

function isPathAllowed(pathname: string, allowedPrefixes?: string[]): boolean {
  const normalizedPath: string = normalizePath(pathname);

  if (!allowedPrefixes || allowedPrefixes.length === 0) {
    return !hasDotSegments(normalizedPath);
  }

  const prefixes: string[] = allowedPrefixes
    .map((prefix) => normalizePath(prefix))
    .filter((prefix) => prefix.length > 0);

  if (prefixes.length === 0) {
    return true;
  }

  if (hasDotSegments(normalizedPath)) {
    return false;
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

  return normalized;
}

function hasDotSegments(pathname: string): boolean {
  const segments = pathname.split('/').filter((segment) => segment.length > 0);
  return segments.some((segment) => segment === '.' || segment === '..');
}
