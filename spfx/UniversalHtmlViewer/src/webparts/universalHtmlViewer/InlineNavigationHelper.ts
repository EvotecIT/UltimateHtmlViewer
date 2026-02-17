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
    });
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

  const target = (anchor.getAttribute('target') || '').trim().toLowerCase();
  if (target && target !== '_self') {
    return undefined;
  }

  if (anchor.hasAttribute('download')) {
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

  if (!isHtmlFilePath(absoluteUrl.pathname)) {
    return undefined;
  }

  const normalizedAbsoluteUrl: string = stripQueryParam(
    absoluteUrl.toString(),
    options.cacheBusterParamName,
  );

  if (!isUrlAllowed(normalizedAbsoluteUrl, options.validationOptions)) {
    return undefined;
  }

  return normalizedAbsoluteUrl;
}

function isPrimaryClick(event: MouseEvent): boolean {
  return (
    !event.defaultPrevented &&
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

function isSameHostAsCurrentPage(targetUrl: URL, currentPageUrl: string): boolean {
  try {
    const current = new URL(currentPageUrl);
    return targetUrl.host.toLowerCase() === current.host.toLowerCase();
  } catch {
    return false;
  }
}

function isHtmlFilePath(pathname: string): boolean {
  const normalized = (pathname || '').toLowerCase();
  return normalized.endsWith('.html') || normalized.endsWith('.htm');
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
