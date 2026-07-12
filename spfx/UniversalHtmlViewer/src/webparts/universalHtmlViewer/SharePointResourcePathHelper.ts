/**
 * Encodes a decoded SharePoint resource path for use as an OData string alias.
 * ResourcePath endpoints require decoded paths so literal percent and hash
 * characters remain unambiguous after URL decoding.
 */
export function encodeSharePointDecodedPathAlias(decodedPath: string): string {
  const escapedODataValue = (decodedPath || '').replace(/'/g, "''");
  return encodeURIComponent(escapedODataValue);
}

export function buildSharePointFileByPathApiUrl(
  webAbsoluteUrl: string,
  decodedServerRelativePath: string,
  resourceSuffix: string = '',
  query: string = '',
): string {
  const encodedPath = encodeSharePointDecodedPathAlias(decodedServerRelativePath);
  const normalizedSuffix = resourceSuffix
    ? resourceSuffix.startsWith('/')
      ? resourceSuffix
      : `/${resourceSuffix}`
    : '';
  const aliasQuery = `@p1='${encodedPath}'`;
  const normalizedQuery = (query || '').replace(/^[?&]+/, '');
  const combinedQuery = normalizedQuery ? `${aliasQuery}&${normalizedQuery}` : aliasQuery;
  return `${webAbsoluteUrl}/_api/web/GetFileByServerRelativePath(decodedUrl=@p1)${normalizedSuffix}?${combinedQuery}`;
}

export function buildSharePointFolderByPathApiUrl(
  webAbsoluteUrl: string,
  decodedServerRelativePath: string,
  resourceSuffix: string = '',
  query: string = '',
): string {
  const encodedPath = encodeSharePointDecodedPathAlias(decodedServerRelativePath);
  const normalizedSuffix = resourceSuffix
    ? resourceSuffix.startsWith('/')
      ? resourceSuffix
      : `/${resourceSuffix}`
    : '';
  const aliasQuery = `@p1='${encodedPath}'`;
  const normalizedQuery = (query || '').replace(/^[?&]+/, '');
  const combinedQuery = normalizedQuery ? `${aliasQuery}&${normalizedQuery}` : aliasQuery;
  return `${webAbsoluteUrl}/_api/web/GetFolderByServerRelativePath(decodedUrl=@p1)${normalizedSuffix}?${combinedQuery}`;
}
