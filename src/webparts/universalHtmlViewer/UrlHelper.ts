import { getQueryStringParam } from './QueryStringHelper';

export type HtmlSourceMode = 'FullUrl' | 'BasePathAndRelativePath' | 'BasePathAndDashboardId';

export type HeightMode = 'Fixed' | 'Viewport';

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

/**
 * Builds the final iframe URL based on configured properties and the current page URL.
 *
 * @param params Url building parameters.
 * @returns The computed URL string, or null if it cannot be determined.
 */
export function buildFinalUrl(params: BuildUrlParams): string | null {
  const mode: HtmlSourceMode = params.htmlSourceMode || 'FullUrl';

  if (mode === 'FullUrl') {
    const normalizedFullUrl: string = (params.fullUrl || '').trim();
    return normalizedFullUrl || null;
  }

  const basePathNormalized: string = normalizeBasePath(params.basePath);

  if (!basePathNormalized) {
    return null;
  }

  if (mode === 'BasePathAndRelativePath') {
    const relativePathNormalized: string = normalizeRelativePath(params.relativePath);

    if (!relativePathNormalized) {
      return null;
    }

    return `${basePathNormalized}${relativePathNormalized}`;
  }

  const queryParamName: string = (params.queryStringParamName || '').trim() || 'dashboard';
  const defaultFileName: string = (params.defaultFileName || '').trim() || 'index.html';

  const dashboardIdFromQuery: string | null = params.pageUrl
    ? getQueryStringParam(params.pageUrl, queryParamName)
    : null;

  const effectiveDashboardId: string = (dashboardIdFromQuery || params.dashboardId || '').trim();

  if (!effectiveDashboardId) {
    return null;
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
export function isUrlAllowed(url: string | null, currentPageUrl: string): boolean {
  if (!url) {
    return false;
  }

  const trimmedUrl: string = url.trim();

  if (!trimmedUrl) {
    return false;
  }

  const lowerUrl: string = trimmedUrl.toLowerCase();

  if (lowerUrl.startsWith('javascript:')) {
    return false;
  }

  if (trimmedUrl.startsWith('/')) {
    return true;
  }

  if (lowerUrl.startsWith('http://') || lowerUrl.startsWith('https://')) {
    try {
      const target: URL = new URL(trimmedUrl);
      const current: URL = new URL(currentPageUrl);

      return target.host.toLowerCase() === current.host.toLowerCase();
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

