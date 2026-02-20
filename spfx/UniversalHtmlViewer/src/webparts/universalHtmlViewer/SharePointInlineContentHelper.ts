import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { prepareInlineHtmlForSrcDoc } from './InlineHtmlTransformHelper';

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

  return prepareInlineHtmlForSrcDoc(html, baseUrlForRelativeLinks, pageUrl);
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
