import type { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { prepareInlineHtmlForSrcDoc } from './InlineHtmlTransformHelper';

interface IInlineHtmlCacheEntry {
  html: string;
  expiresAt: number;
}

const INLINE_HTML_CACHE_TTL_MS = 15000;
const INLINE_HTML_CACHE_MAX_ENTRIES = 120;
const inlineHtmlCache = new Map<string, IInlineHtmlCacheEntry>();
const inlineHtmlInFlightRequests = new Map<string, Promise<string>>();

export async function loadSharePointFileContentForInline(
  spHttpClient: SPHttpClient,
  webAbsoluteUrl: string,
  sourceUrl: string,
  baseUrlForRelativeLinks: string,
  pageUrl: string,
  spHttpClientConfiguration?: unknown,
): Promise<string> {
  const serverRelativePath = getServerRelativePathForSharePointFile(sourceUrl, pageUrl);

  if (!serverRelativePath) {
    throw new Error(
      'SharePoint file API mode requires a same-tenant URL or a site-relative URL.',
    );
  }

  const cacheKey = buildInlineHtmlCacheKey(
    webAbsoluteUrl,
    sourceUrl,
    baseUrlForRelativeLinks,
    pageUrl,
  );
  const cachedHtml = tryGetCachedInlineHtml(cacheKey);
  if (cachedHtml) {
    return cachedHtml;
  }

  const inFlightRequest = inlineHtmlInFlightRequests.get(cacheKey);
  if (inFlightRequest) {
    return inFlightRequest;
  }

  const loadRequest = loadSharePointInlineHtmlFromApi(
    spHttpClient,
    webAbsoluteUrl,
    serverRelativePath,
    baseUrlForRelativeLinks,
    pageUrl,
    spHttpClientConfiguration,
  )
    .then((preparedHtml) => {
      setCachedInlineHtml(cacheKey, preparedHtml);
      return preparedHtml;
    })
    .finally(() => {
      inlineHtmlInFlightRequests.delete(cacheKey);
    });
  inlineHtmlInFlightRequests.set(cacheKey, loadRequest);

  return loadRequest;
}

export function clearInlineHtmlCacheForTests(): void {
  inlineHtmlCache.clear();
  inlineHtmlInFlightRequests.clear();
}

async function loadSharePointInlineHtmlFromApi(
  spHttpClient: SPHttpClient,
  webAbsoluteUrl: string,
  serverRelativePath: string,
  baseUrlForRelativeLinks: string,
  pageUrl: string,
  spHttpClientConfiguration?: unknown,
): Promise<string> {
  const encodedPath = encodeURIComponent(serverRelativePath);
  const apiUrl = `${webAbsoluteUrl}/_api/web/GetFileByServerRelativeUrl(@p1)/$value?@p1='${encodedPath}'`;

  const response: SPHttpClientResponse = await spHttpClient.get(
    apiUrl,
    spHttpClientConfiguration as never,
    {
      headers: {
        Accept: 'text/html,*/*',
        'Cache-Control': 'no-cache',
        Pragma: 'no-cache',
      },
    },
  );

  if (!response.ok) {
    throw new Error(
      `SharePoint API returned ${response.status} ${response.statusText || ''}`.trim(),
    );
  }

  const html = await response.text();
  if (!html || html.trim().length === 0) {
    throw new Error('SharePoint API returned empty HTML content.');
  }

  return prepareInlineHtmlForSrcDoc(html, baseUrlForRelativeLinks, pageUrl);
}

function buildInlineHtmlCacheKey(
  webAbsoluteUrl: string,
  sourceUrl: string,
  baseUrlForRelativeLinks: string,
  pageUrl: string,
): string {
  const normalizedSourceUrl = (sourceUrl || '').trim();
  const normalizedBaseUrl = (baseUrlForRelativeLinks || '').trim();
  const normalizedPageUrl = (pageUrl || '').trim();
  return `${webAbsoluteUrl}|${normalizedSourceUrl}|${normalizedBaseUrl}|${normalizedPageUrl}`;
}

function tryGetCachedInlineHtml(cacheKey: string): string | undefined {
  const cached = inlineHtmlCache.get(cacheKey);
  if (!cached) {
    return undefined;
  }

  if (cached.expiresAt <= Date.now()) {
    inlineHtmlCache.delete(cacheKey);
    return undefined;
  }

  // Keep most-recently used entries at the tail.
  inlineHtmlCache.delete(cacheKey);
  inlineHtmlCache.set(cacheKey, cached);
  return cached.html;
}

function setCachedInlineHtml(cacheKey: string, html: string): void {
  if (!cacheKey || !html) {
    return;
  }

  inlineHtmlCache.delete(cacheKey);
  inlineHtmlCache.set(cacheKey, {
    html,
    expiresAt: Date.now() + INLINE_HTML_CACHE_TTL_MS,
  });
  trimInlineHtmlCache();
}

function trimInlineHtmlCache(): void {
  if (inlineHtmlCache.size <= INLINE_HTML_CACHE_MAX_ENTRIES) {
    return;
  }

  const now = Date.now();
  inlineHtmlCache.forEach((entry, key) => {
    if (entry.expiresAt <= now) {
      inlineHtmlCache.delete(key);
    }
  });

  while (inlineHtmlCache.size > INLINE_HTML_CACHE_MAX_ENTRIES) {
    const firstKey = inlineHtmlCache.keys().next().value as string | undefined;
    if (!firstKey) {
      break;
    }
    inlineHtmlCache.delete(firstKey);
  }
}

function getServerRelativePathForSharePointFile(
  sourceUrl: string,
  pageUrl: string,
): string | undefined {
  if (sourceUrl.startsWith('/')) {
    return stripQueryAndHashFromPath(sourceUrl);
  }

  try {
    const targetUrl = new URL(sourceUrl);
    const currentUrl = new URL(pageUrl);
    if (targetUrl.host.toLowerCase() !== currentUrl.host.toLowerCase()) {
      return undefined;
    }
    return decodeURIComponent(targetUrl.pathname);
  } catch {
    return undefined;
  }
}

function stripQueryAndHashFromPath(value: string): string {
  const hashIndex = value.indexOf('#');
  const queryIndex = value.indexOf('?');

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
