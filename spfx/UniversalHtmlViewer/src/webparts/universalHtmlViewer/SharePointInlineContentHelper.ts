import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export async function loadSharePointFileContentForInline(
  spHttpClient: SPHttpClient,
  webAbsoluteUrl: string,
  sourceUrl: string,
  baseUrlForRelativeLinks: string,
  pageUrl: string,
): Promise<string> {
  const serverRelativePath = getServerRelativePathForSharePointFile(sourceUrl, pageUrl);

  if (!serverRelativePath) {
    throw new Error(
      'SharePoint file API mode requires a same-tenant URL or a site-relative URL.',
    );
  }

  const encodedPath = encodeURIComponent(serverRelativePath);
  const apiUrl = `${webAbsoluteUrl}/_api/web/GetFileByServerRelativeUrl(@p1)/$value?@p1='${encodedPath}'`;

  const response: SPHttpClientResponse = await spHttpClient.get(
    apiUrl,
    SPHttpClient.configurations.v1,
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

  return withBaseHrefForSrcDoc(html, baseUrlForRelativeLinks, pageUrl);
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

function withBaseHrefForSrcDoc(
  html: string,
  baseUrlForRelativeLinks: string,
  pageUrl: string,
): string {
  if (/<base\s+/i.test(html)) {
    return html;
  }

  const baseHref = getAbsoluteUrlWithoutQuery(baseUrlForRelativeLinks, pageUrl);
  const baseTag = `<base href="${escapeHtmlAttribute(baseHref)}">`;

  if (/<head[\s>]/i.test(html)) {
    return html.replace(/<head([^>]*)>/i, `<head$1>${baseTag}`);
  }

  if (/<html[\s>]/i.test(html)) {
    return html.replace(/<html([^>]*)>/i, `<html$1><head>${baseTag}</head>`);
  }

  return `<head>${baseTag}</head>${html}`;
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

function escapeHtmlAttribute(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}
