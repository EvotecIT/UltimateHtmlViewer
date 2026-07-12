import type { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { ReportBrowserDefaultView } from './UniversalHtmlViewerTypes';
import { buildSharePointFolderByPathApiUrl } from './SharePointResourcePathHelper';

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
  maxRequests?: number;
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
  '@odata.nextLink'?: string;
  'odata.nextLink'?: string;
}

type SettledResult<T> =
  | { status: 'fulfilled'; value: T }
  | { status: 'rejected'; reason: unknown };

const DEFAULT_MAX_ITEMS = 300;
const DEFAULT_MAX_REQUESTS = 200;
const MAX_SHAREPOINT_PAGE_SIZE = 5000;
const skippedFolderNames = new Set<string>(['forms']);
const requestBudgets = new WeakMap<object, { remaining: number }>();

class ReportBrowserRequestBudgetExceededError extends Error {}

export async function loadSharePointReportBrowserItems(
  options: ILoadSharePointReportBrowserItemsOptions,
): Promise<ISharePointReportBrowserItem[]> {
  const requestOptions: ILoadSharePointReportBrowserItemsOptions = { ...options };
  const rootPath = normalizeSharePointReportBrowserRootPath(
    options.rootPath,
    options.webAbsoluteUrl,
  );
  if (!rootPath) {
    return [];
  }
  requestBudgets.set(requestOptions, {
    remaining: normalizeMaxRequests(options.maxRequests),
  });

  const currentFolderPath =
    options.view === 'Folders'
      ? normalizeServerRelativeFolderPath(options.currentFolderPath || rootPath)
      : rootPath;
  const maxItems = normalizeMaxItems(options.maxItems);
  const allowedExtensions = normalizeAllowedExtensions(options.allowedExtensions);

  if (options.view === 'Files') {
    const files: ISharePointReportBrowserItem[] = [];
    try {
      await appendFolderFilesRecursive(
        requestOptions,
        rootPath,
        rootPath,
        allowedExtensions,
        maxItems,
        files,
        new Set<string>(),
      );
      return files;
    } finally {
      requestBudgets.delete(requestOptions);
    }
  }

  try {
    const [folderResult, fileResult] = await Promise.all([
      settlePromise(loadFolderEntries(requestOptions, currentFolderPath, maxItems)),
      settlePromise(
        loadAllowedFileEntries(
          requestOptions,
          currentFolderPath,
          allowedExtensions,
          maxItems,
        ),
      ),
    ]);
    if (folderResult.status === 'rejected') {
      throw folderResult.reason;
    }
    if (fileResult.status === 'rejected') {
      throw fileResult.reason;
    }
    const folders = folderResult.value;
    const files = fileResult.value;

    const folderItems = folders
      .filter((folder) => !skippedFolderNames.has(String(folder.Name || '').toLowerCase()))
      .map((folder) =>
        buildFolderItem(folder, rootPath),
      )
      .filter((item): item is ISharePointReportBrowserItem => !!item);
    const fileItems = files
      .map((file) => buildFileItem(file, rootPath))
      .filter((item): item is ISharePointReportBrowserItem => !!item);

    return mergeFolderViewItems(folderItems, fileItems, maxItems);
  } finally {
    requestBudgets.delete(requestOptions);
  }
}

async function settlePromise<T>(promise: Promise<T>): Promise<SettledResult<T>> {
  try {
    return {
      status: 'fulfilled',
      value: await promise,
    };
  } catch (reason) {
    return {
      status: 'rejected',
      reason,
    };
  }
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
  if (normalizedRoot === '/') {
    return normalizedPath.startsWith('/');
  }
  return normalizedPath === normalizedRoot || normalizedPath.startsWith(`${normalizedRoot}/`);
}

async function appendFolderFilesRecursive(
  options: ILoadSharePointReportBrowserItemsOptions,
  rootPath: string,
  folderPath: string,
  allowedExtensions: string[],
  maxItems: number,
  result: ISharePointReportBrowserItem[],
  visitedFolderPaths: Set<string>,
): Promise<void> {
  const normalizedFolderPath = normalizeServerRelativeFolderPath(folderPath);
  const folderVisitKey = normalizedFolderPath.toLowerCase();
  if (
    result.length >= maxItems ||
    !normalizedFolderPath ||
    visitedFolderPaths.has(folderVisitKey)
  ) {
    return;
  }
  visitedFolderPaths.add(folderVisitKey);

  await appendAllowedFileItemsFromFolder(
    options,
    rootPath,
    normalizedFolderPath,
    allowedExtensions,
    maxItems,
    result,
  );

  if (result.length >= maxItems) {
    return;
  }

  await appendChildFolderFilesRecursive(
    options,
    rootPath,
    normalizedFolderPath,
    allowedExtensions,
    maxItems,
    result,
    visitedFolderPaths,
  );
}

async function appendChildFolderFilesRecursive(
  options: ILoadSharePointReportBrowserItemsOptions,
  rootPath: string,
  folderPath: string,
  allowedExtensions: string[],
  maxItems: number,
  result: ISharePointReportBrowserItem[],
  visitedFolderPaths: Set<string>,
): Promise<void> {
  let apiUrl = buildFolderChildrenApiUrl(
    options.webAbsoluteUrl,
    folderPath,
    'Folders',
    `$select=Name,ServerRelativeUrl&$orderby=Name&$top=${normalizeSharePointPageSize(
      maxItems - result.length,
    )}`,
  );
  const visitedUrls = new Set<string>();

  while (apiUrl && result.length < maxItems && !visitedUrls.has(apiUrl)) {
    visitedUrls.add(apiUrl);
    const response =
      await loadSharePointJson<ISharePointListResponse<ISharePointFolderListEntry>>(
        options,
        apiUrl,
      );

    for (const folder of response.value || []) {
      if (result.length >= maxItems) {
        return;
      }

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
        visitedFolderPaths,
      );
    }

    apiUrl = getNextLink(response);
  }
}

async function appendAllowedFileItemsFromFolder(
  options: ILoadSharePointReportBrowserItemsOptions,
  rootPath: string,
  folderPath: string,
  allowedExtensions: string[],
  maxItems: number,
  result: ISharePointReportBrowserItem[],
): Promise<void> {
  const files = await loadAllowedFileEntries(
    options,
    folderPath,
    allowedExtensions,
    maxItems - result.length,
  );

  for (const file of files) {
    if (result.length >= maxItems) {
      return;
    }

    const item = buildFileItem(file, rootPath);
    if (item) {
      result.push(item);
    }
  }
}

async function loadAllowedFileEntries(
  options: ILoadSharePointReportBrowserItemsOptions,
  folderPath: string,
  allowedExtensions: string[],
  maxAcceptedEntries: number,
): Promise<ISharePointFileListEntry[]> {
  if (maxAcceptedEntries <= 0) {
    return [];
  }

  let apiUrl = buildFolderChildrenApiUrl(
    options.webAbsoluteUrl,
    folderPath,
    'Files',
    `$select=Name,ServerRelativeUrl,TimeLastModified&$orderby=Name&$top=${normalizeSharePointPageSize(
      maxAcceptedEntries,
    )}`,
  );
  const files: ISharePointFileListEntry[] = [];
  const visitedUrls = new Set<string>();

  while (apiUrl && files.length < maxAcceptedEntries && !visitedUrls.has(apiUrl)) {
    visitedUrls.add(apiUrl);
    const response = await loadSharePointJson<ISharePointListResponse<ISharePointFileListEntry>>(
      options,
      apiUrl,
    );

    for (const file of response.value || []) {
      if (files.length >= maxAcceptedEntries) {
        break;
      }

      if (isAllowedReportFile(file.ServerRelativeUrl, allowedExtensions)) {
        files.push(file);
      }
    }

    apiUrl = getNextLink(response);
  }

  return files;
}

async function loadFolderEntries(
  options: ILoadSharePointReportBrowserItemsOptions,
  folderPath: string,
  maxEntries?: number,
): Promise<ISharePointFolderListEntry[]> {
  const apiUrl = buildFolderChildrenApiUrl(
    options.webAbsoluteUrl,
    folderPath,
    'Folders',
    `$select=Name,ServerRelativeUrl&$orderby=Name&$top=${normalizeSharePointPageSize(maxEntries)}`,
  );
  const response = await loadSharePointJson<ISharePointListResponse<ISharePointFolderListEntry>>(
    options,
    apiUrl,
  );
  return loadSharePointPagedListEntries(options, response, apiUrl, maxEntries);
}

async function loadSharePointPagedListEntries<T>(
  options: ILoadSharePointReportBrowserItemsOptions,
  firstPage: ISharePointListResponse<T>,
  firstPageUrl: string,
  maxEntries?: number,
): Promise<T[]> {
  const entryLimit = normalizeSharePointEntryLimit(maxEntries);
  const entries: T[] = [...(firstPage.value || [])];
  let nextLink = getNextLink(firstPage);
  const visitedUrls = new Set<string>([firstPageUrl]);

  while (nextLink && entries.length < entryLimit && !visitedUrls.has(nextLink)) {
    visitedUrls.add(nextLink);
    const nextPage = await loadSharePointJson<ISharePointListResponse<T>>(options, nextLink);
    entries.push(...(nextPage.value || []));
    nextLink = getNextLink(nextPage);
  }

  return entries.slice(0, entryLimit);
}

function getNextLink<T>(response: ISharePointListResponse<T>): string {
  return response['@odata.nextLink'] || response['odata.nextLink'] || '';
}

function mergeFolderViewItems(
  folders: ISharePointReportBrowserItem[],
  files: ISharePointReportBrowserItem[],
  maxItems: number,
): ISharePointReportBrowserItem[] {
  if (folders.length + files.length <= maxItems) {
    return [...folders, ...files];
  }

  if (maxItems <= 1) {
    return files.length > 0 ? files.slice(0, 1) : folders.slice(0, 1);
  }

  if (folders.length === 0) {
    return files.slice(0, maxItems);
  }

  if (files.length === 0) {
    return folders.slice(0, maxItems);
  }

  const fileBudget = Math.max(1, Math.min(files.length, Math.floor(maxItems / 2)));
  const folderBudget = Math.max(1, Math.min(folders.length, maxItems - fileBudget));
  const remainingBudget = maxItems - folderBudget - fileBudget;
  const extraFolders = Math.max(0, Math.min(remainingBudget, folders.length - folderBudget));
  const extraFiles = Math.max(
    0,
    Math.min(remainingBudget - extraFolders, files.length - fileBudget),
  );

  return [
    ...folders.slice(0, folderBudget + extraFolders),
    ...files.slice(0, fileBudget + extraFiles),
  ];
}

async function loadSharePointJson<T>(
  options: ILoadSharePointReportBrowserItemsOptions,
  apiUrl: string,
): Promise<T> {
  const requestBudget = requestBudgets.get(options);
  if (requestBudget) {
    if (requestBudget.remaining <= 0) {
      throw new ReportBrowserRequestBudgetExceededError(
        'SharePoint report browser request limit reached.',
      );
    }
    requestBudget.remaining -= 1;
  }
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
  return buildSharePointFolderByPathApiUrl(
    webAbsoluteUrl,
    normalizeServerRelativeFolderPath(folderPath),
    `/${childCollectionName}`,
    query,
  );
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
    serverRelativeUrl: encodeServerRelativeUrlForNavigation(serverRelativeUrl),
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
    serverRelativeUrl: encodeServerRelativeUrlForNavigation(serverRelativeUrl),
    relativePath: getRelativePath(rootPath, serverRelativeUrl),
    timeLastModified: file.TimeLastModified,
  };
}

function normalizeServerRelativeFolderPath(value?: string): string {
  const trimmed = tryDecodeUriComponent((value || '').trim()).replace(/\\/g, '/');
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

function encodeServerRelativeUrlForNavigation(value: string): string {
  const normalized = normalizeServerRelativeFolderPath(value);
  if (!normalized) {
    return '';
  }

  return normalized
    .split('/')
    .map((segment, index) => (index === 0 ? '' : encodeURIComponent(segment)))
    .join('/');
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

function normalizeMaxRequests(value?: number): number {
  if (!Number.isFinite(value) || !value || value <= 0) {
    return DEFAULT_MAX_REQUESTS;
  }
  return Math.max(1, Math.min(Math.floor(value), 1000));
}

function normalizeSharePointPageSize(value?: number): number {
  if (!Number.isFinite(value) || !value || value <= 0) {
    return MAX_SHAREPOINT_PAGE_SIZE;
  }

  return Math.max(1, Math.min(Math.floor(value), MAX_SHAREPOINT_PAGE_SIZE));
}

function normalizeSharePointEntryLimit(value?: number): number {
  if (!Number.isFinite(value) || !value || value <= 0) {
    return Number.POSITIVE_INFINITY;
  }

  return Math.max(1, Math.floor(value));
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

function tryDecodeUriComponent(value: string): string {
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
}
