import { isUrlAllowed, UrlValidationOptions } from './UrlHelper';

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

  const scanFrames = (): void => {
    const iframeDocument = tryGetIframeDocument(options.iframe);
    if (!iframeDocument) {
      return;
    }

    const nestedFrames = iframeDocument.querySelectorAll('iframe[src]');
    nestedFrames.forEach((frame) => {
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
      attributeFilter: ['src'],
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
  };
}

function hydrateNestedFrame(
  frame: HTMLIFrameElement,
  baseUrl: string,
  options: INestedIframeHydrationOptions,
): void {
  const rawSrc = (frame.getAttribute('src') || '').trim();
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

function resolveNestedFrameUrl(
  rawSrc: string,
  baseUrl: string,
  options: INestedIframeHydrationOptions,
): string | undefined {
  let absoluteUrl: URL;
  try {
    absoluteUrl = new URL(rawSrc, baseUrl);
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
