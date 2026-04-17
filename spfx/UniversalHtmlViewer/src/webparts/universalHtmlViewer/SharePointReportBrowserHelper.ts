import type { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { ReportBrowserDefaultView } from './UniversalHtmlViewerTypes';

export interface ISharePointReportBrowserItem {
  name: string;
  serverRelativeUrl: string;
  kind: 'Folder' | 'File';
  relativePath: string;
  timeLastModified?: string;
}

export interface ILoadSharePointReportBrowserItemsOptions {
  spHttpClient: SPHttpClient;
  webAbsoluteUrl: string;
  rootPath: string;
  currentFolderPath?: string;
  allowedExtensions: string[];
  view: ReportBrowserDefaultView;
  maxItems: number;
  spHttpClientConfiguration?: unknown;
}

interface ISharePointFolderListEntry {
  Name?: string;
  ServerRelativeUrl?: string;
}

interface ISharePointFileListEntry {
  Name?: string;
  ServerRelativeUrl?: string;
  TimeLastModified?: string;
}

interface ISharePointListResponse<T> {
  value?: T[];
}

const DEFAULT_MAX_ITEMS = 300;
const MAX_RECURSION_DEPTH = 12;
const skippedFolderNames = new Set<string>(['forms']);

export async function loadSharePointReportBrowserItems(
  options: ILoadSharePointReportBrowserItemsOptions,
): Promise<ISharePointReportBrowserItem[]> {
  const rootPath = normalizeSharePointReportBrowserRootPath(
    options.rootPath,
    options.webAbsoluteUrl,
  );
  if (!rootPath) {
    return [];
  }

  const currentFolderPath =
    options.view === 'Folders'
      ? normalizeServerRelativeFolderPath(options.currentFolderPath || rootPath)
      : rootPath;
  const maxItems = normalizeMaxItems(options.maxItems);
  const allowedExtensions = normalizeAllowedExtensions(options.allowedExtensions);

  if (options.view === 'Files') {
    const files: ISharePointReportBrowserItem[] = [];
    await appendFolderFilesRecursive(
      options,
      rootPath,
      rootPath,
      allowedExtensions,
      maxItems,
      files,
      0,
    );
    return files;
  }

  const [folders, files] = await Promise.all([
    loadFolderEntries(options, currentFolderPath),
    loadFileEntries(options, currentFolderPath),
  ]);

  return [
    ...folders
      .filter((folder) => !skippedFolderNames.has(String(folder.Name || '').toLowerCase()))
      .map((folder) =>
        buildFolderItem(folder, rootPath),
      )
      .filter((item): item is ISharePointReportBrowserItem => !!item),
    ...files
      .filter((file) => isAllowedReportFile(file.ServerRelativeUrl, allowedExtensions))
      .map((file) => buildFileItem(file, rootPath))
      .filter((item): item is ISharePointReportBrowserItem => !!item),
  ].slice(0, maxItems);
}

export function normalizeSharePointReportBrowserRootPath(
  rootPath: string | undefined,
  webAbsoluteUrl: string,
): string {
  const trimmed = stripQueryAndHash(rootPath || '').trim();
  if (!trimmed) {
    return '';
  }

  const webPath = normalizeServerRelativeFolderPath(tryGetUrlPath(webAbsoluteUrl));

  if (trimmed.startsWith('http://') || trimmed.startsWith('https://')) {
    return normalizeServerRelativeFolderPath(tryGetUrlPath(trimmed));
  }

  if (trimmed.startsWith('/')) {
    return normalizeServerRelativeFolderPath(trimmed);
  }

  return normalizeServerRelativeFolderPath(`${webPath}/${trimmed}`);
}

export function getReportBrowserParentFolderPath(
  rootPath: string,
  currentFolderPath: string,
): string {
  const normalizedRoot = normalizeServerRelativeFolderPath(rootPath);
  const normalizedCurrent = normalizeServerRelativeFolderPath(currentFolderPath);
  if (!normalizedRoot || normalizedCurrent === normalizedRoot) {
    return normalizedRoot;
  }

  const lastSlash = normalizedCurrent.lastIndexOf('/');
  if (lastSlash <= 0) {
    return normalizedRoot;
  }

  const parentPath = normalizeServerRelativeFolderPath(normalizedCurrent.substring(0, lastSlash));
  if (!isPathInsideRoot(parentPath, normalizedRoot)) {
    return normalizedRoot;
  }

  return parentPath || normalizedRoot;
}

export function isPathInsideRoot(path: string, rootPath: string): boolean {
  const normalizedPath = normalizeServerRelativeFolderPath(path).toLowerCase();
  const normalizedRoot = normalizeServerRelativeFolderPath(rootPath).toLowerCase();
  return normalizedPath === normalizedRoot || normalizedPath.startsWith(`${normalizedRoot}/`);
}

async function appendFolderFilesRecursive(
  options: ILoadSharePointReportBrowserItemsOptions,
  rootPath: string,
  folderPath: string,
  allowedExtensions: string[],
  maxItems: number,
  result: ISharePointReportBrowserItem[],
  depth: number,
): Promise<void> {
  if (result.length >= maxItems || depth > MAX_RECURSION_DEPTH) {
    return;
  }

  const files = await loadFileEntries(options, folderPath);
  files
    .filter((file) => isAllowedReportFile(file.ServerRelativeUrl, allowedExtensions))
    .forEach((file) => {
      if (result.length < maxItems) {
        const item = buildFileItem(file, rootPath);
        if (item) {
          result.push(item);
        }
      }
    });

  if (result.length >= maxItems) {
    return;
  }

  const folders = await loadFolderEntries(options, folderPath);
  for (const folder of folders) {
    const folderName = String(folder.Name || '').toLowerCase();
    const childFolderPath = normalizeServerRelativeFolderPath(folder.ServerRelativeUrl);
    if (
      !childFolderPath ||
      skippedFolderNames.has(folderName) ||
      !isPathInsideRoot(childFolderPath, rootPath)
    ) {
      continue;
    }

    await appendFolderFilesRecursive(
      options,
      rootPath,
      childFolderPath,
      allowedExtensions,
      maxItems,
      result,
      depth + 1,
    );
  }
}

async function loadFolderEntries(
  options: ILoadSharePointReportBrowserItemsOptions,
  folderPath: string,
): Promise<ISharePointFolderListEntry[]> {
  const apiUrl = buildFolderChildrenApiUrl(
    options.webAbsoluteUrl,
    folderPath,
    'Folders',
    '$select=Name,ServerRelativeUrl&$orderby=Name',
  );
  const response = await loadSharePointJson<ISharePointListResponse<ISharePointFolderListEntry>>(
    options,
    apiUrl,
  );
  return response.value || [];
}

async function loadFileEntries(
  options: ILoadSharePointReportBrowserItemsOptions,
  folderPath: string,
): Promise<ISharePointFileListEntry[]> {
  const apiUrl = buildFolderChildrenApiUrl(
    options.webAbsoluteUrl,
    folderPath,
    'Files',
    '$select=Name,ServerRelativeUrl,TimeLastModified&$orderby=Name',
  );
  const response = await loadSharePointJson<ISharePointListResponse<ISharePointFileListEntry>>(
    options,
    apiUrl,
  );
  return response.value || [];
}

async function loadSharePointJson<T>(
  options: ILoadSharePointReportBrowserItemsOptions,
  apiUrl: string,
): Promise<T> {
  const response: SPHttpClientResponse = await options.spHttpClient.get(
    apiUrl,
    options.spHttpClientConfiguration as never,
    {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'OData-Version': '',
      },
    },
  );

  if (!response.ok) {
    throw new Error(
      `SharePoint report browser API returned ${response.status} ${
        response.statusText || ''
      }`.trim(),
    );
  }

  return (await response.json()) as T;
}

function buildFolderChildrenApiUrl(
  webAbsoluteUrl: string,
  folderPath: string,
  childCollectionName: 'Files' | 'Folders',
  query: string,
): string {
  const encodedPath = encodeURIComponent(normalizeServerRelativeFolderPath(folderPath));
  return `${webAbsoluteUrl}/_api/web/GetFolderByServerRelativeUrl(@p1)/${childCollectionName}?@p1='${encodedPath}'&${query}`;
}

function buildFolderItem(
  folder: ISharePointFolderListEntry,
  rootPath: string,
): ISharePointReportBrowserItem | undefined {
  const serverRelativeUrl = normalizeServerRelativeFolderPath(folder.ServerRelativeUrl);
  if (!serverRelativeUrl || !isPathInsideRoot(serverRelativeUrl, rootPath)) {
    return undefined;
  }

  const name = String(folder.Name || '').trim() || getLeafName(serverRelativeUrl);
  return {
    kind: 'Folder',
    name,
    serverRelativeUrl,
    relativePath: getRelativePath(rootPath, serverRelativeUrl),
  };
}

function buildFileItem(
  file: ISharePointFileListEntry,
  rootPath: string,
): ISharePointReportBrowserItem | undefined {
  const serverRelativeUrl = normalizeServerRelativeFolderPath(file.ServerRelativeUrl);
  if (!serverRelativeUrl || !isPathInsideRoot(serverRelativeUrl, rootPath)) {
    return undefined;
  }

  const name = String(file.Name || '').trim() || getLeafName(serverRelativeUrl);
  return {
    kind: 'File',
    name,
    serverRelativeUrl,
    relativePath: getRelativePath(rootPath, serverRelativeUrl),
    timeLastModified: file.TimeLastModified,
  };
}

function normalizeServerRelativeFolderPath(value?: string): string {
  const trimmed = stripQueryAndHash(value || '').trim().replace(/\\/g, '/');
  if (!trimmed) {
    return '';
  }

  let normalized = trimmed.startsWith('/') ? trimmed : `/${trimmed}`;
  while (normalized.includes('//')) {
    normalized = normalized.replace(/\/{2,}/g, '/');
  }
  return normalized.endsWith('/') && normalized.length > 1
    ? normalized.substring(0, normalized.length - 1)
    : normalized;
}

function normalizeAllowedExtensions(allowedExtensions: string[]): string[] {
  const normalized = allowedExtensions
    .map((extension) => extension.trim().toLowerCase())
    .filter((extension) => extension.length > 0)
    .map((extension) => (extension.startsWith('.') ? extension : `.${extension}`));
  return normalized.length > 0 ? normalized : ['.html', '.htm', '.aspx'];
}

function isAllowedReportFile(
  serverRelativeUrl: string | undefined,
  allowedExtensions: string[],
): boolean {
  const normalizedUrl = String(serverRelativeUrl || '').toLowerCase();
  return allowedExtensions.some((extension) => normalizedUrl.endsWith(extension));
}

function normalizeMaxItems(value: number): number {
  if (!Number.isFinite(value) || value <= 0) {
    return DEFAULT_MAX_ITEMS;
  }
  return Math.max(1, Math.min(Math.floor(value), 1000));
}

function getRelativePath(rootPath: string, serverRelativeUrl: string): string {
  const normalizedRoot = normalizeServerRelativeFolderPath(rootPath);
  const normalizedUrl = normalizeServerRelativeFolderPath(serverRelativeUrl);
  if (normalizedUrl.toLowerCase() === normalizedRoot.toLowerCase()) {
    return '';
  }
  return normalizedUrl.substring(normalizedRoot.length).replace(/^\/+/, '');
}

function getLeafName(serverRelativeUrl: string): string {
  const normalized = normalizeServerRelativeFolderPath(serverRelativeUrl);
  const lastSlash = normalized.lastIndexOf('/');
  return lastSlash === -1 ? normalized : normalized.substring(lastSlash + 1);
}

function stripQueryAndHash(value: string): string {
  const hashIndex = value.indexOf('#');
  const queryIndex = value.indexOf('?');
  if (hashIndex === -1 && queryIndex === -1) {
    return value;
  }
  if (hashIndex === -1) {
    return value.substring(0, queryIndex);
  }
  if (queryIndex === -1) {
    return value.substring(0, hashIndex);
  }
  return value.substring(0, Math.min(hashIndex, queryIndex));
}

function tryGetUrlPath(value: string): string {
  try {
    return new URL(value).pathname;
  } catch {
    return value;
  }
}
