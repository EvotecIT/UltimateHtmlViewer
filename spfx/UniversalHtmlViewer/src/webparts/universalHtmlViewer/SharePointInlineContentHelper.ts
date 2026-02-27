import type { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { prepareInlineHtmlForSrcDoc } from './InlineHtmlTransformHelper';

interface IInlineHtmlCacheEntry {
  html: string;
  expiresAt: number;
}

export interface ILoadSharePointInlineContentOptions {
  cacheTtlMs?: number;
  bypassCache?: boolean;
  maxRetryAttempts?: number;
  retryBaseDelayMs?: number;
  retryMaxDelayMs?: number;
  enforceStrictInlineCsp?: boolean;
}

const DEFAULT_INLINE_HTML_CACHE_TTL_MS = 15000;
const INLINE_HTML_CACHE_MAX_ENTRIES = 120;
const DEFAULT_MAX_RETRY_ATTEMPTS = 3;
const DEFAULT_RETRY_BASE_DELAY_MS = 750;
const DEFAULT_RETRY_MAX_DELAY_MS = 8000;
const inlineHtmlCache = new Map<string, IInlineHtmlCacheEntry>();
const inlineHtmlInFlightRequests = new Map<string, Promise<string>>();

export async function loadSharePointFileContentForInline(
  spHttpClient: SPHttpClient,
  webAbsoluteUrl: string,
  sourceUrl: string,
  baseUrlForRelativeLinks: string,
  pageUrl: string,
  spHttpClientConfiguration?: unknown,
  options?: ILoadSharePointInlineContentOptions,
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
    options?.enforceStrictInlineCsp === true,
  );
  const bypassCache = options?.bypassCache === true;
  const cacheTtlMs = normalizeCacheTtlMs(options?.cacheTtlMs);
  const useResponseCache = !bypassCache && cacheTtlMs > 0;

  if (useResponseCache) {
    const cachedHtml = tryGetCachedInlineHtml(cacheKey);
    if (cachedHtml) {
      return cachedHtml;
    }
  }

  if (useResponseCache) {
    const inFlightRequest = inlineHtmlInFlightRequests.get(cacheKey);
    if (inFlightRequest) {
      return inFlightRequest;
    }
  }

  const loadRequest = loadSharePointInlineHtmlFromApi(
    spHttpClient,
    webAbsoluteUrl,
    serverRelativePath,
    baseUrlForRelativeLinks,
    pageUrl,
    spHttpClientConfiguration,
    options,
  )
    .then((preparedHtml) => {
      if (useResponseCache) {
        setCachedInlineHtml(cacheKey, preparedHtml, cacheTtlMs);
      }
      return preparedHtml;
    })
    .finally(() => {
      if (useResponseCache) {
        inlineHtmlInFlightRequests.delete(cacheKey);
      }
    });

  if (useResponseCache) {
    inlineHtmlInFlightRequests.set(cacheKey, loadRequest);
  }

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
  options?: ILoadSharePointInlineContentOptions,
): Promise<string> {
  const encodedPath = encodeURIComponent(serverRelativePath);
  const apiUrl = `${webAbsoluteUrl}/_api/web/GetFileByServerRelativeUrl(@p1)/$value?@p1='${encodedPath}'`;

  const response: SPHttpClientResponse = await getInlineHtmlResponseWithRetry(
    spHttpClient,
    apiUrl,
    spHttpClientConfiguration,
    options,
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

  return prepareInlineHtmlForSrcDoc(html, baseUrlForRelativeLinks, pageUrl, {
    enforceStrictInlineCsp: options?.enforceStrictInlineCsp === true,
  });
}

async function getInlineHtmlResponseWithRetry(
  spHttpClient: SPHttpClient,
  apiUrl: string,
  spHttpClientConfiguration?: unknown,
  options?: ILoadSharePointInlineContentOptions,
): Promise<SPHttpClientResponse> {
  const maxRetryAttempts = normalizeMaxRetryAttempts(options?.maxRetryAttempts);
  const retryBaseDelayMs = normalizeRetryDelayMs(
    options?.retryBaseDelayMs,
    DEFAULT_RETRY_BASE_DELAY_MS,
  );
  const retryMaxDelayMs = normalizeRetryDelayMs(
    options?.retryMaxDelayMs,
    DEFAULT_RETRY_MAX_DELAY_MS,
  );
  let response: SPHttpClientResponse | undefined;

  for (let attempt = 1; attempt <= maxRetryAttempts; attempt += 1) {
    try {
      response = await spHttpClient.get(apiUrl, spHttpClientConfiguration as never, {
        headers: {
          Accept: 'text/html,*/*',
          'Cache-Control': 'no-cache',
          Pragma: 'no-cache',
        },
      });
    } catch (error) {
      const isFinalAttempt = attempt >= maxRetryAttempts;
      if (isFinalAttempt || isAbortRequestError(error)) {
        throw error;
      }
      const computedBackoffMs = getRetryBackoffDelayMs(
        attempt,
        retryBaseDelayMs,
        retryMaxDelayMs,
      );
      await sleep(computedBackoffMs);
      continue;
    }

    if (response.ok) {
      return response;
    }

    const isFinalAttempt = attempt >= maxRetryAttempts;
    if (isFinalAttempt || !isRetryableStatusCode(response.status)) {
      return response;
    }

    const retryAfterMs = tryGetRetryAfterDelayMs(response);
    const computedBackoffMs = getRetryBackoffDelayMs(
      attempt,
      retryBaseDelayMs,
      retryMaxDelayMs,
    );
    const delayMs = retryAfterMs !== undefined ? retryAfterMs : computedBackoffMs;
    await sleep(delayMs);
  }

  if (!response) {
    throw new Error('SharePoint API call did not produce a response.');
  }

  return response;
}

function buildInlineHtmlCacheKey(
  webAbsoluteUrl: string,
  sourceUrl: string,
  baseUrlForRelativeLinks: string,
  pageUrl: string,
  enforceStrictInlineCsp: boolean,
): string {
  const normalizedSourceUrl = (sourceUrl || '').trim();
  const normalizedBaseUrl = (baseUrlForRelativeLinks || '').trim();
  const normalizedPageUrl = normalizePageUrlForCache(pageUrl);
  const normalizedStrictMode = enforceStrictInlineCsp ? 'strict-csp' : 'default-csp';
  return `${webAbsoluteUrl}|${normalizedSourceUrl}|${normalizedBaseUrl}|${normalizedPageUrl}|${normalizedStrictMode}`;
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

function setCachedInlineHtml(cacheKey: string, html: string, cacheTtlMs: number): void {
  if (!cacheKey || !html) {
    return;
  }

  inlineHtmlCache.delete(cacheKey);
  inlineHtmlCache.set(cacheKey, {
    html,
    expiresAt: Date.now() + cacheTtlMs,
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

function normalizeCacheTtlMs(cacheTtlMs?: number): number {
  if (typeof cacheTtlMs !== 'number' || !Number.isFinite(cacheTtlMs)) {
    return DEFAULT_INLINE_HTML_CACHE_TTL_MS;
  }
  if (cacheTtlMs <= 0) {
    return 0;
  }

  return Math.round(cacheTtlMs);
}

function normalizeMaxRetryAttempts(value?: number): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    return DEFAULT_MAX_RETRY_ATTEMPTS;
  }

  const rounded = Math.floor(value);
  if (rounded <= 1) {
    return 1;
  }

  if (rounded > 5) {
    return 5;
  }

  return rounded;
}

function normalizeRetryDelayMs(value: number | undefined, fallback: number): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    return fallback;
  }

  if (value <= 0) {
    return 0;
  }

  return Math.round(value);
}

function isRetryableStatusCode(statusCode: number): boolean {
  return statusCode === 429 || statusCode === 502 || statusCode === 503 || statusCode === 504;
}

function tryGetRetryAfterDelayMs(response: SPHttpClientResponse): number | undefined {
  const headers = response.headers;
  const retryAfterRaw = headers?.get?.('Retry-After');
  if (!retryAfterRaw) {
    return undefined;
  }

  const retryAfter = retryAfterRaw.trim();
  if (!retryAfter) {
    return undefined;
  }

  const retryAfterSeconds = Number(retryAfter);
  if (Number.isFinite(retryAfterSeconds) && retryAfterSeconds >= 0) {
    return Math.round(retryAfterSeconds * 1000);
  }

  const retryDateMs = Date.parse(retryAfter);
  if (!Number.isFinite(retryDateMs)) {
    return undefined;
  }

  const deltaMs = retryDateMs - Date.now();
  if (deltaMs <= 0) {
    return 0;
  }

  return Math.round(deltaMs);
}

function getRetryBackoffDelayMs(
  attempt: number,
  retryBaseDelayMs: number,
  retryMaxDelayMs: number,
): number {
  if (retryBaseDelayMs <= 0) {
    return 0;
  }

  const exponent = Math.max(0, attempt - 1);
  const computedDelay = retryBaseDelayMs * Math.pow(2, exponent);
  if (retryMaxDelayMs <= 0) {
    return Math.round(computedDelay);
  }

  return Math.round(Math.min(computedDelay, retryMaxDelayMs));
}

function isAbortRequestError(error: unknown): boolean {
  const source = error as
    | {
        name?: unknown;
        code?: unknown;
      }
    | undefined;
  const name = typeof source?.name === 'string' ? source.name.toLowerCase() : '';
  if (name === 'aborterror') {
    return true;
  }

  const code = typeof source?.code === 'string' ? source.code.toUpperCase() : '';
  return code === 'ABORT_ERR';
}

function sleep(delayMs: number): Promise<void> {
  if (delayMs <= 0) {
    return Promise.resolve();
  }

  return new Promise((resolve) => {
    setTimeout(resolve, delayMs);
  });
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

function normalizePageUrlForCache(pageUrl: string): string {
  const normalizedValue = (pageUrl || '').trim();
  if (!normalizedValue) {
    return '';
  }

  try {
    const parsed = new URL(normalizedValue);
    parsed.search = '';
    parsed.hash = '';
    return parsed.toString();
  } catch {
    return stripQueryAndHashFromPath(normalizedValue);
  }
}
