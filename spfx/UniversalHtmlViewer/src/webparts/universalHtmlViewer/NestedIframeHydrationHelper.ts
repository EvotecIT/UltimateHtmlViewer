import { isUrlAllowed, UrlValidationOptions } from './UrlHelper';
import { resolveInlineNavigationTarget } from './InlineNavigationHelper';

export type NestedIframeHydrationDiagnosticEvent =
  | 'nestedHydrationStarted'
  | 'nestedHydrationSucceeded'
  | 'nestedHydrationFailed'
  | 'nestedHydrationStaleResultIgnored'
  | 'nestedNavigationStarted'
  | 'nestedNavigationSucceeded'
  | 'nestedNavigationFailed'
  | 'nestedNavigationStaleResultIgnored';

export interface INestedIframeHydrationOptions {
  iframe: HTMLIFrameElement;
  currentPageUrl: string;
  validationOptions: UrlValidationOptions;
  cacheBusterParamName: string;
  loadInlineHtml: (
    sourceUrl: string,
    baseUrlForRelativeLinks: string,
  ) => Promise<string | undefined>;
  onDiagnosticsEvent?: (eventName: NestedIframeHydrationDiagnosticEvent) => void;
}

export function wireNestedIframeHydration(
  options: INestedIframeHydrationOptions,
): () => void {
  let observer: MutationObserver | undefined;
  const frameCleanupMap = new Map<HTMLIFrameElement, () => void>();
  const mutationScanDebounceMs = 40;
  let scheduledScanTimeoutId: number | undefined;

  const clearScheduledScan = (): void => {
    if (scheduledScanTimeoutId === undefined) {
      return;
    }

    if (typeof window !== 'undefined' && typeof window.clearTimeout === 'function') {
      window.clearTimeout(scheduledScanTimeoutId);
    } else {
      clearTimeout(scheduledScanTimeoutId as unknown as ReturnType<typeof setTimeout>);
    }
    scheduledScanTimeoutId = undefined;
  };

  const scheduleScan = (): void => {
    if (scheduledScanTimeoutId !== undefined) {
      return;
    }

    const executeScan = (): void => {
      scheduledScanTimeoutId = undefined;
      scanFrames();
    };

    if (typeof window !== 'undefined' && typeof window.setTimeout === 'function') {
      scheduledScanTimeoutId = window.setTimeout(executeScan, mutationScanDebounceMs);
      return;
    }

    scheduledScanTimeoutId = setTimeout(
      executeScan,
      mutationScanDebounceMs,
    ) as unknown as number;
  };

  const pruneStaleFrameCleanup = (activeFrames?: Set<HTMLIFrameElement>): void => {
    frameCleanupMap.forEach((cleanup, frame) => {
      const isFrameActive = activeFrames
        ? activeFrames.has(frame)
        : frame.isConnected;
      if (isFrameActive) {
        return;
      }

      cleanup();
      frameCleanupMap.delete(frame);
    });
  };

  function scanFrames(): void {
    const iframeDocument = tryGetIframeDocument(options.iframe);
    if (!iframeDocument) {
      pruneStaleFrameCleanup();
      return;
    }

    const nestedFrames = iframeDocument.querySelectorAll(
      'iframe[src], iframe[data-uhv-inline-src]',
    );
    const activeFrames = new Set<HTMLIFrameElement>();
    nestedFrames.forEach((frame) => {
      activeFrames.add(frame as HTMLIFrameElement);
    });
    pruneStaleFrameCleanup(activeFrames);

    activeFrames.forEach((frame) => {
      ensureNestedFrameNavigationWired(
        frame,
        options,
        frameCleanupMap,
      );
      hydrateNestedFrame(
        frame,
        iframeDocument.baseURI || options.currentPageUrl,
        options,
      );
    });
  }

  const attachObserver = (): void => {
    const iframeDocument = tryGetIframeDocument(options.iframe);
    if (!iframeDocument || !iframeDocument.documentElement || typeof MutationObserver === 'undefined') {
      return;
    }

    observer = new MutationObserver((mutations) => {
      let shouldScan = false;

      mutations.forEach((mutation) => {
        if (mutation.type === 'attributes') {
          const target = mutation.target as Element;
          if (target && target.tagName.toLowerCase() === 'iframe') {
            const frame = target as HTMLIFrameElement;
            frame.removeAttribute('data-uhv-nested-state');
            frame.removeAttribute('data-uhv-nested-src');
            shouldScan = true;
          }
          return;
        }

        shouldScan = true;
      });

      if (shouldScan) {
        scheduleScan();
      }
    });
    observer.observe(iframeDocument.documentElement, {
      childList: true,
      subtree: true,
      attributes: true,
      attributeFilter: ['src', 'data-uhv-inline-src'],
    });
  };

  const onLoad = (): void => {
    clearScheduledScan();
    if (observer) {
      observer.disconnect();
      observer = undefined;
    }
    scanFrames();
    attachObserver();
  };

  options.iframe.addEventListener('load', onLoad);
  onLoad();

  return (): void => {
    clearScheduledScan();
    options.iframe.removeEventListener('load', onLoad);
    if (observer) {
      observer.disconnect();
      observer = undefined;
    }
    frameCleanupMap.forEach((cleanup) => {
      cleanup();
    });
    frameCleanupMap.clear();
  };
}

function hydrateNestedFrame(
  frame: HTMLIFrameElement,
  baseUrl: string,
  options: INestedIframeHydrationOptions,
): void {
  const rawSrc = getFrameSource(frame);
  if (!rawSrc || rawSrc.startsWith('#')) {
    return;
  }

  const lastSrc = frame.getAttribute('data-uhv-nested-src');
  if (lastSrc && lastSrc !== rawSrc) {
    frame.removeAttribute('data-uhv-nested-state');
  }
  frame.setAttribute('data-uhv-nested-src', rawSrc);

  const state = frame.getAttribute('data-uhv-nested-state');
  if (state === 'processing' || state === 'done') {
    return;
  }

  const normalizedUrl = resolveNestedFrameUrl(rawSrc, baseUrl, options);
  if (!normalizedUrl) {
    return;
  }

  frame.setAttribute('data-uhv-nested-state', 'processing');
  const hydrationSource = rawSrc;
  emitDiagnosticsEvent(options, 'nestedHydrationStarted');

  options
    .loadInlineHtml(normalizedUrl, normalizedUrl)
    .then((inlineHtml) => {
      if (!frame.isConnected) {
        return;
      }
      if (frame.getAttribute('data-uhv-nested-src') !== hydrationSource) {
        emitDiagnosticsEvent(options, 'nestedHydrationStaleResultIgnored');
        return;
      }
      if (frame.getAttribute('data-uhv-nested-state') !== 'processing') {
        return;
      }
      if (!inlineHtml || inlineHtml.trim().length === 0) {
        frame.setAttribute('data-uhv-nested-state', 'failed');
        emitDiagnosticsEvent(options, 'nestedHydrationFailed');
        return;
      }
      frame.srcdoc = inlineHtml;
      frame.setAttribute('data-uhv-nested-state', 'done');
      emitDiagnosticsEvent(options, 'nestedHydrationSucceeded');
    })
    .catch(() => {
      if (!frame.isConnected) {
        return;
      }
      if (frame.getAttribute('data-uhv-nested-src') !== hydrationSource) {
        emitDiagnosticsEvent(options, 'nestedHydrationStaleResultIgnored');
        return;
      }
      if (frame.getAttribute('data-uhv-nested-state') === 'processing') {
        frame.setAttribute('data-uhv-nested-state', 'failed');
        emitDiagnosticsEvent(options, 'nestedHydrationFailed');
      }
    });
}

function getFrameSource(frame: HTMLIFrameElement): string {
  const inlineSrc = (frame.getAttribute('data-uhv-inline-src') || '').trim();
  if (inlineSrc) {
    return inlineSrc;
  }

  return (frame.getAttribute('src') || '').trim();
}

function ensureNestedFrameNavigationWired(
  frame: HTMLIFrameElement,
  options: INestedIframeHydrationOptions,
  frameCleanupMap: Map<HTMLIFrameElement, () => void>,
): void {
  if (frameCleanupMap.has(frame)) {
    return;
  }

  const interceptedEventNames: string[] = ['pointerdown', 'mousedown', 'click'];
  let wiredDocument: Document | undefined;
  let wiredWindow: Window | undefined;
  let wiredClickHandler: ((event: Event) => void) | undefined;
  let lastNavigationTargetUrl = '';
  let lastNavigationTimestamp = 0;
  const pendingScrollResetTimeoutIds: number[] = [];
  const clearPendingScrollResetTimeouts = (): void => {
    if (pendingScrollResetTimeoutIds.length === 0) {
      return;
    }

    while (pendingScrollResetTimeoutIds.length > 0) {
      const timeoutId = pendingScrollResetTimeoutIds.pop();
      if (timeoutId === undefined) {
        continue;
      }

      if (typeof window !== 'undefined' && typeof window.clearTimeout === 'function') {
        window.clearTimeout(timeoutId);
      } else {
        clearTimeout(timeoutId as unknown as ReturnType<typeof setTimeout>);
      }
    }
  };
  const clearFrameClickHandlers = (): void => {
    if (wiredDocument && wiredClickHandler) {
      interceptedEventNames.forEach((eventName) => {
        wiredDocument?.removeEventListener(eventName, wiredClickHandler as EventListener, true);
      });
    }
    if (wiredWindow && wiredClickHandler) {
      interceptedEventNames.forEach((eventName) => {
        wiredWindow?.removeEventListener(eventName, wiredClickHandler as EventListener, true);
      });
    }
    if (wiredDocument?.documentElement?.getAttribute('data-uhv-inline-nav') === '1') {
      wiredDocument.documentElement.removeAttribute('data-uhv-inline-nav');
    }
    wiredDocument = undefined;
    wiredWindow = undefined;
    wiredClickHandler = undefined;
  };

  const onFrameLoad = (): void => {
    clearFrameClickHandlers();
    clearPendingScrollResetTimeouts();
    resetNestedFrameScrollPosition(frame);
    if (typeof window !== 'undefined') {
      const firstTimeoutId = window.setTimeout(() => {
        resetNestedFrameScrollPosition(frame);
      }, 80);
      const secondTimeoutId = window.setTimeout(() => {
        resetNestedFrameScrollPosition(frame);
      }, 260);
      pendingScrollResetTimeoutIds.push(firstTimeoutId, secondTimeoutId);
    }

    const frameDocument = tryGetIframeDocument(frame);
    if (!frameDocument || !frameDocument.documentElement) {
      return;
    }

    const root = frameDocument.documentElement;
    if (root.getAttribute('data-uhv-inline-nav') === '1') {
      return;
    }

    root.setAttribute('data-uhv-inline-nav', '1');
    const handledEvents = new WeakSet<Event>();
    const onClick = (event: Event): void => {
      if (handledEvents.has(event)) {
        return;
      }

      const currentPageUrl =
        frame.getAttribute('data-uhv-nested-src') ||
        frameDocument.baseURI ||
        options.currentPageUrl;

      const targetUrl: string | undefined = resolveInlineNavigationTarget(
        event as MouseEvent,
        {
          currentPageUrl,
          validationOptions: options.validationOptions,
          cacheBusterParamName: options.cacheBusterParamName,
        },
      );
      if (!targetUrl) {
        return;
      }

      const now = Date.now();
      const isDuplicatedNavigation =
        lastNavigationTargetUrl === targetUrl && now - lastNavigationTimestamp < 500;
      suppressInterceptedNavigationEvent(event);
      handledEvents.add(event);
      if (isDuplicatedNavigation) {
        return;
      }

      lastNavigationTargetUrl = targetUrl;
      lastNavigationTimestamp = now;
      frame.setAttribute('data-uhv-nested-state', 'processing');
      frame.setAttribute('data-uhv-nested-src', targetUrl);
      const navigationSource = targetUrl;
      emitDiagnosticsEvent(options, 'nestedNavigationStarted');

      options
        .loadInlineHtml(targetUrl, targetUrl)
        .then((inlineHtml) => {
          if (!frame.isConnected) {
            return;
          }
          if (frame.getAttribute('data-uhv-nested-src') !== navigationSource) {
            emitDiagnosticsEvent(options, 'nestedNavigationStaleResultIgnored');
            return;
          }
          if (frame.getAttribute('data-uhv-nested-state') !== 'processing') {
            return;
          }
          if (!inlineHtml || inlineHtml.trim().length === 0) {
            frame.setAttribute('data-uhv-nested-state', 'failed');
            emitDiagnosticsEvent(options, 'nestedNavigationFailed');
            return;
          }
          frame.srcdoc = inlineHtml;
          frame.setAttribute('data-uhv-nested-state', 'done');
          emitDiagnosticsEvent(options, 'nestedNavigationSucceeded');
        })
        .catch(() => {
          if (!frame.isConnected) {
            return;
          }
          if (frame.getAttribute('data-uhv-nested-src') !== navigationSource) {
            emitDiagnosticsEvent(options, 'nestedNavigationStaleResultIgnored');
            return;
          }
          frame.setAttribute('data-uhv-nested-state', 'failed');
          emitDiagnosticsEvent(options, 'nestedNavigationFailed');
        });
    };
    const frameWindow = tryGetIframeWindow(frame);
    if (frameWindow) {
      interceptedEventNames.forEach((eventName) => {
        frameWindow.addEventListener(eventName, onClick, true);
      });
      wiredWindow = frameWindow;
    }
    wiredDocument = frameDocument;
    wiredClickHandler = onClick;
    interceptedEventNames.forEach((eventName) => {
      frameDocument.addEventListener(eventName, onClick, true);
    });
  };

  frame.addEventListener('load', onFrameLoad);
  onFrameLoad();

  frameCleanupMap.set(frame, () => {
    clearPendingScrollResetTimeouts();
    frame.removeEventListener('load', onFrameLoad);
    clearFrameClickHandlers();
  });
}

function resolveNestedFrameUrl(
  rawSrc: string,
  baseUrl: string,
  options: INestedIframeHydrationOptions,
): string | undefined {
  let absoluteUrl: URL;
  try {
    absoluteUrl = new URL(rawSrc, baseUrl);
  } catch {
    try {
      absoluteUrl = new URL(rawSrc, options.currentPageUrl);
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

  const normalizedUrl = stripQueryParam(
    absoluteUrl.toString(),
    options.cacheBusterParamName,
  );

  if (!isUrlAllowed(normalizedUrl, options.validationOptions)) {
    return undefined;
  }

  return normalizedUrl;
}

function emitDiagnosticsEvent(
  options: INestedIframeHydrationOptions,
  eventName: NestedIframeHydrationDiagnosticEvent,
): void {
  if (!options.onDiagnosticsEvent) {
    return;
  }

  try {
    options.onDiagnosticsEvent(eventName);
  } catch {
    return;
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

function resetNestedFrameScrollPosition(frame: HTMLIFrameElement, depth: number = 0): void {
  if (depth > 2) {
    return;
  }

  try {
    const frameWindow = frame.contentWindow;
    if (frameWindow) {
      frameWindow.scrollTo(0, 0);
    }

    const frameDocument = frame.contentDocument;
    if (!frameDocument) {
      return;
    }
    if (frameDocument.documentElement) {
      frameDocument.documentElement.scrollTop = 0;
      frameDocument.documentElement.scrollLeft = 0;
    }
    if (frameDocument.body) {
      frameDocument.body.scrollTop = 0;
      frameDocument.body.scrollLeft = 0;
    }

    const nestedFrames = frameDocument.querySelectorAll('iframe');
    nestedFrames.forEach((nestedFrame) => {
      resetNestedFrameScrollPosition(nestedFrame as HTMLIFrameElement, depth + 1);
    });
  } catch {
    // Ignore cross-origin nested frame access issues.
  }
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
  const normalizedName = (paramName || '').trim();
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
