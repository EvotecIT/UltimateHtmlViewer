import { isUrlAllowed, UrlValidationOptions } from './UrlHelper';
import { resolveInlineNavigationTarget } from './InlineNavigationHelper';

export interface INestedIframeHydrationOptions {
  iframe: HTMLIFrameElement;
  currentPageUrl: string;
  validationOptions: UrlValidationOptions;
  cacheBusterParamName: string;
  loadInlineHtml: (
    sourceUrl: string,
    baseUrlForRelativeLinks: string,
  ) => Promise<string | undefined>;
}

export function wireNestedIframeHydration(
  options: INestedIframeHydrationOptions,
): () => void {
  let observer: MutationObserver | undefined;
  const frameCleanupMap = new Map<HTMLIFrameElement, () => void>();

  const scanFrames = (): void => {
    const iframeDocument = tryGetIframeDocument(options.iframe);
    if (!iframeDocument) {
      return;
    }

    const nestedFrames = iframeDocument.querySelectorAll(
      'iframe[src], iframe[data-uhv-inline-src]',
    );
    nestedFrames.forEach((frame) => {
      ensureNestedFrameNavigationWired(
        frame as HTMLIFrameElement,
        options,
        frameCleanupMap,
      );
      hydrateNestedFrame(
        frame as HTMLIFrameElement,
        iframeDocument.baseURI || options.currentPageUrl,
        options,
      );
    });
  };

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
        scanFrames();
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

  options
    .loadInlineHtml(normalizedUrl, normalizedUrl)
    .then((inlineHtml) => {
      if (frame.getAttribute('data-uhv-nested-state') !== 'processing') {
        return;
      }
      if (!inlineHtml || inlineHtml.trim().length === 0) {
        frame.setAttribute('data-uhv-nested-state', 'failed');
        return;
      }
      frame.srcdoc = inlineHtml;
      frame.setAttribute('data-uhv-nested-state', 'done');
    })
    .catch(() => {
      if (frame.getAttribute('data-uhv-nested-state') === 'processing') {
        frame.setAttribute('data-uhv-nested-state', 'failed');
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

  let wiredDocument: Document | undefined;
  let wiredClickHandler: ((event: Event) => void) | undefined;
  const clearDocumentHandler = (): void => {
    if (wiredDocument && wiredClickHandler) {
      wiredDocument.removeEventListener('click', wiredClickHandler, true);
    }
    if (wiredDocument?.documentElement?.getAttribute('data-uhv-inline-nav') === '1') {
      wiredDocument.documentElement.removeAttribute('data-uhv-inline-nav');
    }
    wiredDocument = undefined;
    wiredClickHandler = undefined;
  };

  const onFrameLoad = (): void => {
    clearDocumentHandler();
    resetNestedFrameScrollPosition(frame);
    if (typeof window !== 'undefined') {
      window.setTimeout(() => {
        resetNestedFrameScrollPosition(frame);
      }, 80);
      window.setTimeout(() => {
        resetNestedFrameScrollPosition(frame);
      }, 260);
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
    const onClick = (event: Event): void => {
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

      event.preventDefault();
      event.stopPropagation();
      frame.setAttribute('data-uhv-nested-state', 'processing');
      frame.setAttribute('data-uhv-nested-src', targetUrl);

      options
        .loadInlineHtml(targetUrl, targetUrl)
        .then((inlineHtml) => {
          if (!inlineHtml || inlineHtml.trim().length === 0) {
            frame.setAttribute('data-uhv-nested-state', 'failed');
            return;
          }
          frame.srcdoc = inlineHtml;
          frame.setAttribute('data-uhv-nested-state', 'done');
        })
        .catch(() => {
          frame.setAttribute('data-uhv-nested-state', 'failed');
        });
    };
    wiredDocument = frameDocument;
    wiredClickHandler = onClick;
    frameDocument.addEventListener('click', onClick, true);
  };

  frame.addEventListener('load', onFrameLoad);
  onFrameLoad();

  frameCleanupMap.set(frame, () => {
    frame.removeEventListener('load', onFrameLoad);
    clearDocumentHandler();
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

function tryGetIframeDocument(iframe: HTMLIFrameElement): Document | undefined {
  try {
    return iframe.contentDocument || undefined;
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
