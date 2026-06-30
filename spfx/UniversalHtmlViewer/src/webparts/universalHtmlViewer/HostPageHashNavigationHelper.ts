import {
  findSamePageHashTarget,
  isSamePageHashHref,
  navigateToSamePageHash,
} from './SamePageHashNavigationHelper';

const HOST_HASH_MARKER_ATTRIBUTE = 'data-uhv-host-hash';
const HOST_HASH_SCROLL_OFFSET_PX = 96;

export function tryApplyHostPageHashNavigation(
  iframe: HTMLIFrameElement,
  onHandled?: (iframeDocument: Document, hashHref: string) => void,
): boolean {
  if (typeof window === 'undefined' || !window.location) {
    return false;
  }

  const hashHref = (window.location.hash || '').trim();
  const iframeDocument = tryGetIframeDocument(iframe);
  const rootElement = iframeDocument?.documentElement || undefined;
  if (!iframeDocument || !rootElement) {
    return false;
  }

  if (!isSamePageHashHref(hashHref)) {
    clearHostPageHashNavigationMarker(rootElement);
    return false;
  }

  if (rootElement.getAttribute(HOST_HASH_MARKER_ATTRIBUTE) === hashHref) {
    return true;
  }

  const handled = navigateToSamePageHash(iframeDocument, hashHref, true);
  if (handled) {
    rootElement.setAttribute(HOST_HASH_MARKER_ATTRIBUTE, hashHref);
    onHandled?.(iframeDocument, hashHref);
  } else {
    clearHostPageHashNavigationMarker(rootElement);
  }

  return handled;
}

export function clearHostPageHashNavigationMarker(rootElement: HTMLElement): void {
  rootElement.removeAttribute(HOST_HASH_MARKER_ATTRIBUTE);
}

export function scrollHostPageToIframeHashTarget(
  iframe: HTMLIFrameElement,
  iframeDocument: Document,
  hashHref: string,
): boolean {
  const targetElement = findSamePageHashTarget(iframeDocument, hashHref);
  const hostWindow = iframe.ownerDocument?.defaultView || window;
  if (!targetElement || !hostWindow || typeof hostWindow.scrollTo !== 'function') {
    return false;
  }

  if (typeof iframe.getBoundingClientRect !== 'function') {
    return false;
  }

  const iframeRect = iframe.getBoundingClientRect();
  const targetRect = targetElement.getBoundingClientRect();
  const scrollContainer = findScrollableHostContainer(iframe);
  if (scrollContainer) {
    const containerRect = scrollContainer.getBoundingClientRect();
    const targetTop = Math.max(
      0,
      scrollContainer.scrollTop + iframeRect.top - containerRect.top + targetRect.top - HOST_HASH_SCROLL_OFFSET_PX,
    );

    return scrollElementToTop(scrollContainer, targetTop);
  }

  const currentScrollTop =
    hostWindow.pageYOffset ||
    hostWindow.document.documentElement.scrollTop ||
    hostWindow.document.body?.scrollTop ||
    0;
  const targetTop = Math.max(
    0,
    currentScrollTop + iframeRect.top + targetRect.top - HOST_HASH_SCROLL_OFFSET_PX,
  );

  try {
    hostWindow.scrollTo({
      top: targetTop,
      left: hostWindow.pageXOffset || 0,
      behavior: 'auto',
    });
    return true;
  } catch {
    try {
      hostWindow.scrollTo(hostWindow.pageXOffset || 0, targetTop);
      return true;
    } catch {
      return false;
    }
  }
}

function findScrollableHostContainer(iframe: HTMLIFrameElement): HTMLElement | undefined {
  const hostWindow = iframe.ownerDocument?.defaultView || window;
  let currentElement = iframe.parentElement;

  while (currentElement && currentElement !== iframe.ownerDocument?.body) {
    const overflowY = hostWindow.getComputedStyle(currentElement).overflowY || '';
    if (
      currentElement.scrollHeight > currentElement.clientHeight &&
      (overflowY === 'auto' || overflowY === 'scroll' || overflowY === 'overlay')
    ) {
      return currentElement;
    }

    currentElement = currentElement.parentElement;
  }

  return undefined;
}

function scrollElementToTop(element: HTMLElement, top: number): boolean {
  const scrollableElement = element as HTMLElement & {
    scrollTo?: (optionsOrX?: ScrollToOptions | number, y?: number) => void;
  };

  if (typeof scrollableElement.scrollTo === 'function') {
    try {
      scrollableElement.scrollTo({
        top,
        left: element.scrollLeft || 0,
        behavior: 'auto',
      });
      return true;
    } catch {
      try {
        scrollableElement.scrollTo(element.scrollLeft || 0, top);
        return true;
      } catch {
        // Fall back to direct scrollTop assignment below.
      }
    }
  }

  try {
    element.scrollTop = top;
    return element.scrollTop === top;
  } catch {
    return false;
  }
}

function tryGetIframeDocument(iframe: HTMLIFrameElement): Document | undefined {
  try {
    return iframe.contentDocument || undefined;
  } catch {
    return undefined;
  }
}
