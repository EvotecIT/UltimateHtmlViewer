import { getInlineExternalScriptAllowedHosts } from './ExternalScriptInliningHelper';

export type InlineHtmlRenderMode = 'SrcDoc' | 'BlobUrl';

export function buildInlineHtmlCacheKey(
  webAbsoluteUrl: string,
  sourceUrl: string,
  baseUrlForRelativeLinks: string,
  pageUrl: string,
  renderMode: InlineHtmlRenderMode,
  enforceStrictInlineCsp: boolean,
  inlineExternalScripts: boolean,
  inlineExternalScriptAllowedHosts: string[],
  inlineCspScriptAllowedHosts: string[],
  inlineCspStyleAllowedHosts: string[],
  inlineCspImageAllowedHosts: string[],
  rewriteInlineAnchorHrefs: boolean,
  rewriteInlineAnchorAllowedFileExtensions: string[],
  rewriteInlineAnchorAllowedPathPrefixes: string[],
  rewriteInlineAnchorDeepLinkQueryParamName: string,
  rewriteInlineAnchorPreservedHostQueryParamNames: string[],
): string {
  const normalizedSourceUrl = (sourceUrl || '').trim();
  const normalizedBaseUrl = (baseUrlForRelativeLinks || '').trim();
  const normalizedPageUrl = normalizePageUrlForCache(pageUrl);
  const normalizedRenderMode = renderMode === 'BlobUrl' ? 'blob-url' : 'srcdoc';
  const normalizedStrictMode = enforceStrictInlineCsp ? 'strict-csp' : 'default-csp';
  const normalizedExternalScriptMode = inlineExternalScripts
    ? `inline-scripts:${getInlineExternalScriptAllowedHosts(
        inlineExternalScriptAllowedHosts,
      ).join(',')}`
    : 'external-scripts';
  const normalizedInlineCspSources = [
    `script:${normalizeList(inlineCspScriptAllowedHosts, true).join(',')}`,
    `style:${normalizeList(inlineCspStyleAllowedHosts, true).join(',')}`,
    `img:${normalizeList(inlineCspImageAllowedHosts, true).join(',')}`,
  ].join('|');
  const normalizedAnchorRewriteMode = rewriteInlineAnchorHrefs
    ? `rewrite-anchor-hrefs:${normalizeExtensions(
        rewriteInlineAnchorAllowedFileExtensions,
      ).join(',')}:${normalizeList(
      rewriteInlineAnchorAllowedPathPrefixes,
      true,
      ).join(',')}:${normalizeList(
        [rewriteInlineAnchorDeepLinkQueryParamName],
        false,
      ).join(',')}:${normalizeList(
        rewriteInlineAnchorPreservedHostQueryParamNames,
        false,
      ).join(',')}`
    : 'raw-anchor-hrefs';
  return `${webAbsoluteUrl}|${normalizedSourceUrl}|${normalizedBaseUrl}|${normalizedPageUrl}|${normalizedRenderMode}|${normalizedStrictMode}|${normalizedExternalScriptMode}|${normalizedInlineCspSources}|${normalizedAnchorRewriteMode}`;
}

function normalizeExtensions(extensions: string[]): string[] {
  return normalizeList(extensions, true).map((extension) =>
    extension.startsWith('.') ? extension : `.${extension}`,
  );
}

function normalizeList(values: string[], lowerCase: boolean): string[] {
  return Array.from(
    new Set(
      (values || [])
        .map((value) => (lowerCase ? value.toLowerCase() : value).trim())
        .filter((value) => value.length > 0),
    ),
  ).sort();
}

function normalizePageUrlForCache(pageUrl: string): string {
  try {
    const parsed = new URL(pageUrl);
    parsed.search = '';
    parsed.hash = '';
    return parsed.toString();
  } catch {
    const normalized = (pageUrl || '').trim();
    const queryIndex = normalized.indexOf('?');
    const hashIndex = normalized.indexOf('#');
    const indexes = [queryIndex, hashIndex].filter((index) => index >= 0);
    return indexes.length > 0
      ? normalized.substring(0, Math.min(...indexes))
      : normalized;
  }
}
