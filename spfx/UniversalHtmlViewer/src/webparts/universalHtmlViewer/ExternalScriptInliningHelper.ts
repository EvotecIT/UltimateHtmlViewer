export interface IInlineExternalScriptsOptions {
  enabled?: boolean;
  allowedHosts?: string[];
}

interface IExternalScriptLoadResult {
  scriptText: string;
  sourceUrl: string;
}

const DEFAULT_ALLOWED_HOSTS: string[] = [
  'code.jquery.com',
  'cdnjs.cloudflare.com',
  'cdn.datatables.net',
  'cdn.jsdelivr.net',
  'nightly.datatables.net',
  'unpkg.com',
];
const externalScriptCache = new Map<string, string>();
const externalScriptInFlightRequests = new Map<string, Promise<string>>();
const EXTERNAL_SCRIPT_CACHE_MAX_ENTRIES = 80;

export async function inlineAllowedExternalScripts(
  html: string,
  baseUrlForRelativeLinks: string,
  pageUrl: string,
  options?: IInlineExternalScriptsOptions,
): Promise<string> {
  if (
    options?.enabled !== true ||
    !html ||
    !/<script[\s\S]+src\s*=/i.test(html) ||
    typeof DOMParser === 'undefined' ||
    typeof fetch !== 'function'
  ) {
    return html;
  }

  try {
    const hasDoctype: boolean = /^\s*<!doctype/i.test(html);
    const parser = new DOMParser();
    const parsed = parser.parseFromString(html, 'text/html');
    if (!parsed || !parsed.documentElement) {
      return html;
    }

    const allowedHosts = getInlineExternalScriptAllowedHosts(options.allowedHosts);
    const scripts = Array.from(parsed.querySelectorAll('script[src]'));
    for (const script of scripts) {
      const rawSrc = (script.getAttribute('src') || '').trim();
      const absoluteSrc = resolveExternalScriptUrl(rawSrc, baseUrlForRelativeLinks, pageUrl);
      if (
        !absoluteSrc ||
        !isInlineableScriptElement(script) ||
        !isExternalScriptHostAllowed(absoluteSrc, pageUrl, allowedHosts)
      ) {
        continue;
      }

      const scriptLoadResult = await loadExternalScriptText(
        absoluteSrc,
        pageUrl,
        allowedHosts,
      );
      if (!scriptLoadResult.scriptText) {
        continue;
      }

      script.removeAttribute('src');
      script.removeAttribute('integrity');
      script.removeAttribute('crossorigin');
      script.setAttribute('data-uhv-inlined-external-script', scriptLoadResult.sourceUrl);
      script.textContent = `${escapeScriptTextForInlineHtml(
        scriptLoadResult.scriptText,
      )}\n//# sourceURL=${scriptLoadResult.sourceUrl}`;
    }

    const rebuiltHtml = parsed.documentElement.outerHTML;
    if (!rebuiltHtml) {
      return html;
    }

    return hasDoctype ? `<!DOCTYPE html>${rebuiltHtml}` : rebuiltHtml;
  } catch {
    return html;
  }
}

export function clearExternalScriptInliningCacheForTests(): void {
  externalScriptCache.clear();
  externalScriptInFlightRequests.clear();
}

export function getInlineExternalScriptAllowedHosts(configuredHosts?: string[]): string[] {
  const normalizedConfiguredHosts = (configuredHosts || [])
    .map((host) => normalizeAllowedHost(host))
    .filter((host) => host.length > 0);
  if (normalizedConfiguredHosts.length > 0) {
    return normalizedConfiguredHosts;
  }

  return DEFAULT_ALLOWED_HOSTS;
}

function resolveExternalScriptUrl(
  rawSrc: string,
  baseUrlForRelativeLinks: string,
  pageUrl: string,
): string {
  if (!rawSrc || rawSrc.startsWith('#')) {
    return '';
  }

  const normalizedSrc = rawSrc.trim().toLowerCase();
  const javaScriptScheme = `java${'script:'}`;
  if (
    normalizedSrc.startsWith(javaScriptScheme) ||
    normalizedSrc.startsWith('data:') ||
    normalizedSrc.startsWith('vbscript:')
  ) {
    return '';
  }

  try {
    const page = new URL(pageUrl);
    const base = baseUrlForRelativeLinks
      ? new URL(baseUrlForRelativeLinks, page.origin)
      : page;
    const absolute = new URL(rawSrc, base.toString());
    if (absolute.protocol !== 'https:' && absolute.protocol !== 'http:') {
      return '';
    }
    return absolute.toString();
  } catch {
    return '';
  }
}

function isInlineableScriptElement(script: Element): boolean {
  const scriptType = (script.getAttribute('type') || '').trim().toLowerCase();
  if (!scriptType) {
    return true;
  }

  return [
    'text/javascript',
    'application/javascript',
    'application/ecmascript',
    'text/ecmascript',
  ].includes(scriptType);
}

function isExternalScriptHostAllowed(
  scriptUrl: string,
  pageUrl: string,
  allowedHosts: string[],
): boolean {
  try {
    const parsedScriptUrl = new URL(scriptUrl);
    const parsedPageUrl = new URL(pageUrl);
    if (parsedScriptUrl.hostname.toLowerCase() === parsedPageUrl.hostname.toLowerCase()) {
      return true;
    }

    const scriptHost = parsedScriptUrl.hostname.toLowerCase();
    return allowedHosts.some((allowedHost) => {
      if (!allowedHost) {
        return false;
      }

      if (allowedHost.startsWith('.')) {
        return scriptHost.endsWith(allowedHost) && scriptHost.length > allowedHost.length;
      }

      return scriptHost === allowedHost;
    });
  } catch {
    return false;
  }
}

async function loadExternalScriptText(
  scriptUrl: string,
  pageUrl: string,
  allowedHosts: string[],
): Promise<IExternalScriptLoadResult> {
  const candidateUrls = getExternalScriptCandidateUrls(scriptUrl).filter((candidateUrl) =>
    isExternalScriptHostAllowed(candidateUrl, pageUrl, allowedHosts),
  );

  for (const candidateUrl of candidateUrls) {
    const scriptText = await loadSingleExternalScriptText(candidateUrl);
    if (scriptText) {
      return { scriptText, sourceUrl: candidateUrl };
    }
  }

  return { scriptText: '', sourceUrl: '' };
}

function getExternalScriptCandidateUrls(scriptUrl: string): string[] {
  try {
    const parsedUrl = new URL(scriptUrl);
    if (
      parsedUrl.hostname.toLowerCase() === 'nightly.datatables.net' &&
      parsedUrl.pathname.toLowerCase() === '/js/jquery.datatables.min.js'
    ) {
      return [
        'https://cdn.datatables.net/1.10.24/js/jquery.dataTables.min.js',
        scriptUrl,
      ];
    }
  } catch {
    return [scriptUrl];
  }

  return [scriptUrl];
}

async function loadSingleExternalScriptText(scriptUrl: string): Promise<string> {
  const cachedScript = externalScriptCache.get(scriptUrl);
  if (cachedScript) {
    externalScriptCache.delete(scriptUrl);
    externalScriptCache.set(scriptUrl, cachedScript);
    return cachedScript;
  }

  const inFlightRequest = externalScriptInFlightRequests.get(scriptUrl);
  if (inFlightRequest) {
    return inFlightRequest;
  }

  const request = fetch(scriptUrl, { credentials: 'omit', mode: 'cors' })
    .then(async (response) => {
      if (!response.ok) {
        return '';
      }
      const scriptText = await response.text();
      if (scriptText) {
        externalScriptCache.delete(scriptUrl);
        externalScriptCache.set(scriptUrl, scriptText);
        trimExternalScriptCache();
      }
      return scriptText;
    })
    .catch(() => '')
    .finally(() => {
      externalScriptInFlightRequests.delete(scriptUrl);
    });

  externalScriptInFlightRequests.set(scriptUrl, request);
  return request;
}

function trimExternalScriptCache(): void {
  while (externalScriptCache.size > EXTERNAL_SCRIPT_CACHE_MAX_ENTRIES) {
    const firstKey = externalScriptCache.keys().next().value as string | undefined;
    if (!firstKey) {
      break;
    }
    externalScriptCache.delete(firstKey);
  }
}

function normalizeAllowedHost(host: string): string {
  let normalized = (host || '').trim().toLowerCase();
  if (!normalized) {
    return '';
  }

  try {
    if (normalized.startsWith('http://') || normalized.startsWith('https://')) {
      normalized = new URL(normalized).hostname.toLowerCase();
    }
  } catch {
    return '';
  }

  if (normalized.startsWith('*.')) {
    normalized = normalized.substring(1);
  }

  return normalized.split(':')[0];
}

function escapeScriptTextForInlineHtml(scriptText: string): string {
  return scriptText.replace(/<\/script/gi, '<\\/script');
}
