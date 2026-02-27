import { isUrlAllowed, UrlValidationOptions } from './UrlHelper';

export interface IInlineNavigationOptions {
  iframe: HTMLIFrameElement;
  currentPageUrl: string;
  validationOptions: UrlValidationOptions;
  cacheBusterParamName: string;
  onNavigate: (targetUrl: string) => void;
}

export function wireInlineIframeNavigation(options: IInlineNavigationOptions): () => void {
  const clickHandlers = new Map<EventTarget, (event: Event) => void>();
  const handledEventMarkerKey = '__uhvInlineNavigationHandled';

  const attachHandler = (): void => {
    const iframeDocument: Document | undefined = tryGetIframeDocument(options.iframe);
    if (!iframeDocument) {
      return;
    }

    const rootElement: HTMLElement | undefined = iframeDocument.documentElement || undefined;
    if (!rootElement || rootElement.getAttribute('data-uhv-inline-nav') === '1') {
      return;
    }

    if (clickHandlers.has(iframeDocument)) {
      return;
    }

    rootElement.setAttribute('data-uhv-inline-nav', '1');
    const onClick = (event: Event): void => {
      const handledEvent = event as Event & {
        [handledEventMarkerKey]?: boolean;
      };
      if (handledEvent[handledEventMarkerKey]) {
        return;
      }

      const targetUrl: string | undefined = resolveInlineNavigationTarget(
        event as MouseEvent,
        options,
      );
      if (!targetUrl) {
        return;
      }

      event.preventDefault();
      event.stopPropagation();
      handledEvent[handledEventMarkerKey] = true;
      options.onNavigate(targetUrl);
    };
    iframeDocument.addEventListener('click', onClick, true);

    const iframeWindow: Window | undefined = tryGetIframeWindow(options.iframe);
    if (iframeWindow) {
      iframeWindow.addEventListener('click', onClick, true);
      clickHandlers.set(iframeWindow, onClick);
    }

    clickHandlers.set(iframeDocument, onClick);
  };

  options.iframe.addEventListener('load', attachHandler);
  attachHandler();

  return (): void => {
    options.iframe.removeEventListener('load', attachHandler);
    clickHandlers.forEach((handler, eventTarget) => {
      eventTarget.removeEventListener('click', handler, true);
      if (eventTarget instanceof Document) {
        const rootElement: HTMLElement | undefined = eventTarget.documentElement || undefined;
        if (rootElement && rootElement.getAttribute('data-uhv-inline-nav') === '1') {
          rootElement.removeAttribute('data-uhv-inline-nav');
        }
      }
    });
    clickHandlers.clear();
  };
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

  const rawHref: string = getAnchorNavigationHref(anchor);
  if (!rawHref || rawHref.startsWith('#')) {
    return undefined;
  }

  const protocolBlocked = isNonHttpProtocol(rawHref);
  if (protocolBlocked) {
    return undefined;
  }

  let absoluteUrl: URL;
  try {
    absoluteUrl = new URL(rawHref, getAnchorAbsoluteHref(anchor, options.currentPageUrl));
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
    return undefined;
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

function getAnchorFromEvent(event: MouseEvent): Element | undefined {
  const targetElement: Element | undefined = getEventTargetElement(event);
  if (!targetElement) {
    return undefined;
  }

  const anchor = findClosestAnchorElement(targetElement);
  if (anchor) {
    return anchor;
  }

  const pathElements = getEventComposedPathElements(event);
  for (let index = 0; index < pathElements.length; index += 1) {
    const pathAnchor = findClosestAnchorElement(pathElements[index]);
    if (pathAnchor) {
      return pathAnchor;
    }
  }

  const forcedUrlContainer = targetElement.closest('.fc-event-forced-url');
  if (!forcedUrlContainer) {
    return undefined;
  }

  const forcedAnchor = forcedUrlContainer.querySelector('a[href], a[xlink\\:href], a');
  if (!forcedAnchor || forcedAnchor.tagName.toLowerCase() !== 'a') {
    return undefined;
  }

  return forcedAnchor;
}

function findClosestAnchorElement(element: Element): Element | undefined {
  const anchor = element.closest('a');
  if (!anchor || anchor.tagName.toLowerCase() !== 'a') {
    return undefined;
  }

  return anchor;
}

function getEventTargetElement(event: MouseEvent): Element | undefined {
  const target = event.target as EventTarget | null;
  if (!target) {
    return undefined;
  }

  if (target instanceof Element) {
    return target;
  }

  if (target instanceof Node) {
    return target.parentElement || undefined;
  }

  return undefined;
}

function getEventComposedPathElements(event: MouseEvent): Element[] {
  type EventWithComposedPath = MouseEvent & {
    composedPath?: () => EventTarget[];
  };

  const eventWithComposedPath = event as EventWithComposedPath;
  if (typeof eventWithComposedPath.composedPath !== 'function') {
    return [];
  }

  let pathTargets: EventTarget[];
  try {
    pathTargets = eventWithComposedPath.composedPath();
  } catch {
    return [];
  }

  if (!Array.isArray(pathTargets) || pathTargets.length === 0) {
    return [];
  }

  const elements: Element[] = [];
  const seenElements = new Set<Element>();
  pathTargets.forEach((target) => {
    if (target instanceof Element) {
      if (!seenElements.has(target)) {
        elements.push(target);
        seenElements.add(target);
      }
      return;
    }

    if (target instanceof Node && target.parentElement && !seenElements.has(target.parentElement)) {
      elements.push(target.parentElement);
      seenElements.add(target.parentElement);
    }
  });

  return elements;
}

function getAnchorNavigationHref(anchor: Element): string {
  const attributeHref = (anchor.getAttribute('href') || '').trim();
  if (attributeHref) {
    return attributeHref;
  }

  const xlinkHref = readXLinkHref(anchor);
  if (xlinkHref) {
    return xlinkHref;
  }

  return getAnchorHrefFromProperty(anchor);
}

function getAnchorAbsoluteHref(anchor: Element, fallbackUrl: string): string {
  const hrefFromProperty = getAnchorHrefFromProperty(anchor);
  if (hrefFromProperty) {
    return hrefFromProperty;
  }

  const xlinkHref = readXLinkHref(anchor);
  if (xlinkHref) {
    return xlinkHref;
  }

  const attributeHref = (anchor.getAttribute('href') || '').trim();
  return attributeHref || fallbackUrl;
}

function getAnchorHrefFromProperty(anchor: Element): string {
  const anchorWithHref = anchor as Element & {
    href?: string | { baseVal?: string; animVal?: string };
  };
  const hrefValue = anchorWithHref.href;
  if (typeof hrefValue === 'string') {
    return hrefValue.trim();
  }

  if (hrefValue && typeof hrefValue === 'object') {
    const baseValue = (hrefValue.baseVal || '').trim();
    if (baseValue) {
      return baseValue;
    }

    const animatedValue = (hrefValue.animVal || '').trim();
    if (animatedValue) {
      return animatedValue;
    }
  }

  return '';
}

function readXLinkHref(anchor: Element): string {
  const xlinkNamespace = 'http://www.w3.org/1999/xlink';
  const anchorWithGetAttributeNs = anchor as Element & {
    getAttributeNS?: (namespace: string | null, localName: string) => string | null;
  };
  if (typeof anchorWithGetAttributeNs.getAttributeNS === 'function') {
    const namespacedHref = (anchorWithGetAttributeNs.getAttributeNS(xlinkNamespace, 'href') || '').trim();
    if (namespacedHref) {
      return namespacedHref;
    }
  }

  return '';
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

function tryGetIframeWindow(iframe: HTMLIFrameElement): Window | undefined {
  try {
    return iframe.contentWindow || undefined;
  } catch {
    return undefined;
  }
}
