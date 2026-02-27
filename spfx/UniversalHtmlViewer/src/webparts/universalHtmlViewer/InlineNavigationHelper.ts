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
  const interceptedEventNames: string[] = ['pointerdown', 'mousedown', 'click'];
  const handledEvents = new WeakSet<Event>();
  let lastNavigationTargetUrl = '';
  let lastNavigationTimestamp = 0;
  const emitNavigation = (targetUrl: string, event?: Event): void => {
    const now = Date.now();
    const isDuplicatedNavigation =
      lastNavigationTargetUrl === targetUrl && now - lastNavigationTimestamp < 500;

    if (event) {
      suppressInterceptedNavigationEvent(event);
      handledEvents.add(event);
    }
    if (isDuplicatedNavigation) {
      return;
    }

    lastNavigationTargetUrl = targetUrl;
    lastNavigationTimestamp = now;
    options.onNavigate(targetUrl);
  };
  const registerClickHandler = (
    eventTarget: EventTarget,
    handler: (event: Event) => void,
  ): void => {
    const existingHandler = clickHandlers.get(eventTarget);
    if (existingHandler) {
      interceptedEventNames.forEach((eventName) => {
        eventTarget.removeEventListener(eventName, existingHandler, true);
      });
    }
    interceptedEventNames.forEach((eventName) => {
      eventTarget.addEventListener(eventName, handler, true);
    });
    clickHandlers.set(eventTarget, handler);
  };

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
      if (handledEvents.has(event)) {
        return;
      }

      const targetUrl: string | undefined = resolveInlineNavigationTarget(
        event as MouseEvent,
        options,
      );
      if (!targetUrl) {
        return;
      }

      emitNavigation(targetUrl, event);
    };
    registerClickHandler(iframeDocument, onClick);

    const iframeWindow: Window | undefined = tryGetIframeWindow(options.iframe);
    if (iframeWindow) {
      registerClickHandler(iframeWindow, onClick);
    }
  };
  const onInlineNavigationBridgeMessage = (event: MessageEvent): void => {
    const iframeWindow = tryGetIframeWindow(options.iframe);
    if (iframeWindow && event.source && event.source !== iframeWindow) {
      return;
    }

    const bridgePayload = event.data as
      | {
          type?: unknown;
          targetUrl?: unknown;
        }
      | undefined;
    if (!bridgePayload || bridgePayload.type !== 'uhv-inline-nav') {
      return;
    }

    const rawTargetUrl =
      typeof bridgePayload.targetUrl === 'string' ? bridgePayload.targetUrl.trim() : '';
    if (!rawTargetUrl) {
      return;
    }

    const targetUrl = resolveInlineNavigationTargetFromRawHref(rawTargetUrl, {
      currentPageUrl: options.currentPageUrl,
      validationOptions: options.validationOptions,
      cacheBusterParamName: options.cacheBusterParamName,
    });
    if (!targetUrl) {
      return;
    }

    emitNavigation(targetUrl);
  };

  options.iframe.addEventListener('load', attachHandler);
  if (typeof window !== 'undefined') {
    window.addEventListener('message', onInlineNavigationBridgeMessage);
  }
  attachHandler();

  return (): void => {
    options.iframe.removeEventListener('load', attachHandler);
    if (typeof window !== 'undefined') {
      window.removeEventListener('message', onInlineNavigationBridgeMessage);
    }
    clickHandlers.forEach((handler, eventTarget) => {
      interceptedEventNames.forEach((eventName) => {
        eventTarget.removeEventListener(eventName, handler, true);
      });
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

  return resolveInlineNavigationTargetFromRawHref(rawHref, {
    currentPageUrl: options.currentPageUrl,
    validationOptions: options.validationOptions,
    cacheBusterParamName: options.cacheBusterParamName,
    baseUrl: getAnchorAbsoluteHref(anchor, options.currentPageUrl),
  });
}

interface IInlineNavigationTargetResolutionOptions
  extends Pick<
    IInlineNavigationOptions,
    'currentPageUrl' | 'validationOptions' | 'cacheBusterParamName'
  > {
  baseUrl?: string;
}

function resolveInlineNavigationTargetFromRawHref(
  rawHref: string,
  options: IInlineNavigationTargetResolutionOptions,
): string | undefined {
  const normalizedRawHref = (rawHref || '').trim();
  if (!normalizedRawHref || normalizedRawHref.startsWith('#')) {
    return undefined;
  }

  if (isNonHttpProtocol(normalizedRawHref)) {
    return undefined;
  }

  let absoluteUrl: URL;
  const baseUrl = (options.baseUrl || '').trim() || options.currentPageUrl;
  try {
    absoluteUrl = new URL(normalizedRawHref, baseUrl);
  } catch {
    try {
      absoluteUrl = new URL(normalizedRawHref, options.currentPageUrl);
    } catch {
      return undefined;
    }
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

  const descendantAnchor = findUniqueDescendantAnchorElement(targetElement);
  if (descendantAnchor) {
    return descendantAnchor;
  }

  const pathElements = getEventComposedPathElements(event);
  for (let index = 0; index < pathElements.length; index += 1) {
    const pathElement = pathElements[index];
    const pathAnchor =
      findClosestAnchorElement(pathElement) || findUniqueDescendantAnchorElement(pathElement);
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

function findUniqueDescendantAnchorElement(element: Element): Element | undefined {
  const anchors = element.querySelectorAll('a[href], a[xlink\\:href], a');
  if (anchors.length !== 1) {
    return undefined;
  }

  const anchor = anchors[0];
  if (anchor.tagName.toLowerCase() !== 'a') {
    return undefined;
  }

  return anchor;
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

  const ownerDocumentBaseUrl = (anchor.ownerDocument?.baseURI || '').trim();
  const attributeHref = (anchor.getAttribute('href') || '').trim();
  if (attributeHref) {
    try {
      return new URL(attributeHref, ownerDocumentBaseUrl || fallbackUrl).toString();
    } catch {
      return attributeHref;
    }
  }

  return ownerDocumentBaseUrl || fallbackUrl;
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

function suppressInterceptedNavigationEvent(event: Event): void {
  event.preventDefault();
  event.stopPropagation();
  if (typeof event.stopImmediatePropagation === 'function') {
    event.stopImmediatePropagation();
  }

  const mutableEvent = event as Event & {
    cancelBubble?: boolean;
    returnValue?: boolean;
  };
  mutableEvent.cancelBubble = true;
  mutableEvent.returnValue = false;
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
