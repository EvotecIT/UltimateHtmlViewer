import { isUrlAllowed, UrlValidationOptions } from './UrlHelper';

export interface IInlineNavigationOptions {
  iframe: HTMLIFrameElement;
  currentPageUrl: string;
  validationOptions: UrlValidationOptions;
  cacheBusterParamName: string;
  onNavigate: (targetUrl: string) => void;
}

export function wireInlineIframeNavigation(options: IInlineNavigationOptions): void {
  const attachHandler = (): void => {
    const iframeDocument: Document | undefined = tryGetIframeDocument(options.iframe);
    if (!iframeDocument) {
      return;
    }

    const rootElement: HTMLElement | undefined = iframeDocument.documentElement || undefined;
    if (!rootElement || rootElement.getAttribute('data-uhv-inline-nav') === '1') {
      return;
    }

    rootElement.setAttribute('data-uhv-inline-nav', '1');
    iframeDocument.addEventListener('click', (event) => {
      const targetUrl: string | undefined = resolveInlineNavigationTarget(event, options);
      if (!targetUrl) {
        return;
      }

      event.preventDefault();
      event.stopPropagation();
      options.onNavigate(targetUrl);
    }, true);
  };

  options.iframe.addEventListener('load', attachHandler);
  attachHandler();
}

export function resolveInlineNavigationTarget(
  event: MouseEvent,
  options: Pick<
    IInlineNavigationOptions,
    'currentPageUrl' | 'validationOptions' | 'cacheBusterParamName'
  >,
): string | undefined {
  if (!isPrimaryClick(event)) {
    return undefined;
  }

  const anchor = getAnchorFromEvent(event);
  if (!anchor) {
    return undefined;
  }

  const rawHref: string = (anchor.getAttribute('href') || '').trim();
  if (!rawHref || rawHref.startsWith('#')) {
    return undefined;
  }

  const protocolBlocked = isNonHttpProtocol(rawHref);
  if (protocolBlocked) {
    return undefined;
  }

  let absoluteUrl: URL;
  try {
    absoluteUrl = new URL(rawHref, anchor.href || options.currentPageUrl);
  } catch {
    return undefined;
  }

  if (!isSameHostAsCurrentPage(absoluteUrl, options.currentPageUrl)) {
    return undefined;
  }

  if (
    !isInlineNavigablePath(
      absoluteUrl.pathname,
      options.validationOptions.allowedFileExtensions,
    )
  ) {
    return undefined;
  }

  const normalizedAbsoluteUrl: string = stripQueryParam(
    absoluteUrl.toString(),
    options.cacheBusterParamName,
  );

  if (!isUrlAllowed(normalizedAbsoluteUrl, options.validationOptions)) {
    // Safe fallback for relative links inside trusted inline reports:
    // keep host and extension restrictions, but relax path-prefix checks
    // to avoid false negatives on generated relative navigation menus.
    if (
      !isRelativePathReference(rawHref) ||
      !isUrlAllowed(normalizedAbsoluteUrl, {
        ...options.validationOptions,
        allowedPathPrefixes: undefined,
      })
    ) {
      return undefined;
    }
  }

  return normalizedAbsoluteUrl;
}

function isPrimaryClick(event: MouseEvent): boolean {
  return (
    event.button === 0 &&
    !event.metaKey &&
    !event.ctrlKey &&
    !event.shiftKey &&
    !event.altKey
  );
}

function getAnchorFromEvent(event: MouseEvent): HTMLAnchorElement | undefined {
  const target = event.target as Element | undefined;
  if (!target) {
    return undefined;
  }

  const anchor = target.closest('a[href]');
  if (!anchor || anchor.tagName.toLowerCase() !== 'a') {
    return undefined;
  }

  return anchor as HTMLAnchorElement;
}

function isNonHttpProtocol(value: string): boolean {
  const normalized = value.trim().toLowerCase();
  const protocolMatch = normalized.match(/^([a-z][a-z0-9+\-.]*):/i);
  if (!protocolMatch) {
    return false;
  }

  const protocol: string = (protocolMatch[1] || '').toLowerCase();
  return protocol === 'javascript' || protocol === 'data' || protocol === 'mailto' || protocol === 'tel';
}

function isRelativePathReference(value: string): boolean {
  const normalized = (value || '').trim();
  if (!normalized) {
    return false;
  }
  if (
    normalized.startsWith('/') ||
    normalized.startsWith('#') ||
    normalized.startsWith('?') ||
    normalized.startsWith('//')
  ) {
    return false;
  }

  return !/^[a-z][a-z0-9+\-.]*:/i.test(normalized);
}

function isSameHostAsCurrentPage(targetUrl: URL, currentPageUrl: string): boolean {
  try {
    const current = new URL(currentPageUrl);
    return targetUrl.host.toLowerCase() === current.host.toLowerCase();
  } catch {
    return false;
  }
}

function isInlineNavigablePath(
  pathname: string,
  allowedExtensions?: string[],
): boolean {
  const extension = getPathExtension(pathname);
  if (!extension) {
    return false;
  }

  const normalizedAllowed: string[] = (allowedExtensions || [])
    .map((entry) => (entry.startsWith('.') ? entry.toLowerCase() : `.${entry.toLowerCase()}`))
    .filter((entry) => entry.length > 1);

  if (normalizedAllowed.length > 0) {
    return normalizedAllowed.includes(extension);
  }

  return extension === '.html' || extension === '.htm' || extension === '.aspx';
}

function getPathExtension(pathname: string): string {
  const normalized = (pathname || '').toLowerCase();
  const lastSlash = normalized.lastIndexOf('/');
  const lastDot = normalized.lastIndexOf('.');
  if (lastDot === -1 || lastDot < lastSlash) {
    return '';
  }

  return normalized.substring(lastDot);
}

function stripQueryParam(url: string, paramName: string): string {
  const normalizedName: string = (paramName || '').trim();
  if (!normalizedName) {
    return url;
  }

  try {
    const parsed = new URL(url);
    parsed.searchParams.delete(normalizedName);
    return parsed.toString();
  } catch {
    return url;
  }
}

function tryGetIframeDocument(iframe: HTMLIFrameElement): Document | undefined {
  try {
    return iframe.contentDocument || undefined;
  } catch {
    return undefined;
  }
}
