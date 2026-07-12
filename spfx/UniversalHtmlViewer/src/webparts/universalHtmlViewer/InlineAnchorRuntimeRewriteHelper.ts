import {
  IRewriteInlineNavigationAnchorHrefsOptions,
  rewriteInlineNavigationAnchorElement,
} from './InlineAnchorRewriteHelper';

export interface IInlineAnchorRuntimeRewriteOptions
  extends IRewriteInlineNavigationAnchorHrefsOptions {
  iframe: HTMLIFrameElement;
  fallbackBaseUrl: string;
  fallbackHostPageUrl: string;
}

/**
 * Rewrites anchors created after the report document was parsed (for example,
 * FullCalendar event links). The rewritten host-page URL is also a safe
 * fallback if another report listener prevents UHV's click interceptor from
 * cancelling the browser's native navigation.
 */
export function wireInlineAnchorRuntimeRewrite(
  options: IInlineAnchorRuntimeRewriteOptions,
): () => void {
  let observer: MutationObserver | undefined;
  let scanTimeoutId: number | undefined;

  const clearScheduledScan = (): void => {
    if (scanTimeoutId === undefined) {
      return;
    }

    getFrameWindow(options.iframe)?.clearTimeout(scanTimeoutId);
    scanTimeoutId = undefined;
  };

  const rewriteCurrentAnchors = (): void => {
    scanTimeoutId = undefined;
    const frameDocument = tryGetFrameDocument(options.iframe);
    if (!frameDocument) {
      return;
    }

    const baseUrl = frameDocument.baseURI || options.fallbackBaseUrl;
    const hostPageUrl =
      options.fallbackHostPageUrl ||
      options.iframe.ownerDocument?.defaultView?.location?.href ||
      '';
    frameDocument.querySelectorAll('a[href]').forEach((anchor) => {
      rewriteInlineNavigationAnchorElement(
        anchor,
        getRuntimeAnchorBaseUrl(anchor, frameDocument, baseUrl),
        hostPageUrl,
        options,
      );
    });
  };

  const scheduleScan = (): void => {
    if (scanTimeoutId !== undefined) {
      return;
    }

    const frameWindow = getFrameWindow(options.iframe);
    if (!frameWindow) {
      rewriteCurrentAnchors();
      return;
    }

    scanTimeoutId = frameWindow.setTimeout(rewriteCurrentAnchors, 0);
  };

  const attach = (): void => {
    clearScheduledScan();
    observer?.disconnect();
    observer = undefined;

    const frameDocument = tryGetFrameDocument(options.iframe);
    if (!frameDocument?.documentElement) {
      return;
    }

    rewriteCurrentAnchors();
    const Observer = frameDocument.defaultView?.MutationObserver;
    if (!Observer) {
      return;
    }

    observer = new Observer((mutations) => {
      if (mutations.some((mutation) => mutation.type === 'childList')) {
        scheduleScan();
      }
    });
    observer.observe(frameDocument.documentElement, {
      childList: true,
      subtree: true,
    });
  };

  options.iframe.addEventListener('load', attach);
  attach();

  return (): void => {
    clearScheduledScan();
    options.iframe.removeEventListener('load', attach);
    observer?.disconnect();
    observer = undefined;
  };
}

function getRuntimeAnchorBaseUrl(
  anchor: Element,
  frameDocument: Document,
  defaultBaseUrl: string,
): string {
  const rawHref = (anchor.getAttribute('href') || '').trim();
  if (
    !rawHref ||
    rawHref.startsWith('/') ||
    rawHref.startsWith('#') ||
    /^[a-z][a-z0-9+.-]*:/i.test(rawHref) ||
    rawHref.includes('/')
  ) {
    return defaultBaseUrl;
  }

  const rawFileName = getUrlFileName(rawHref, defaultBaseUrl);
  if (!rawFileName) {
    return defaultBaseUrl;
  }

  const matchingSources = new Set<string>();
  frameDocument
    .querySelectorAll('iframe[data-uhv-inline-src], iframe[data-uhv-nested-src]')
    .forEach((frame) => {
      const source =
        (frame.getAttribute('data-uhv-inline-src') || '').trim() ||
        (frame.getAttribute('data-uhv-nested-src') || '').trim();
      if (!source || getUrlFileName(source, defaultBaseUrl) !== rawFileName) {
        return;
      }

      try {
        matchingSources.add(new URL(source, defaultBaseUrl).toString());
      } catch {
        return;
      }
    });

  return matchingSources.size === 1
    ? Array.from(matchingSources)[0]
    : defaultBaseUrl;
}

function getUrlFileName(value: string, baseUrl: string): string {
  try {
    const parsed = new URL(value, baseUrl);
    const path = parsed.pathname.replace(/\\/g, '/');
    const lastSlash = path.lastIndexOf('/');
    return decodeURIComponent(path.substring(lastSlash + 1)).toLowerCase();
  } catch {
    return '';
  }
}

function tryGetFrameDocument(iframe: HTMLIFrameElement): Document | undefined {
  try {
    return iframe.contentDocument || undefined;
  } catch {
    return undefined;
  }
}

function getFrameWindow(iframe: HTMLIFrameElement): Window | undefined {
  try {
    return iframe.contentWindow || iframe.ownerDocument?.defaultView || undefined;
  } catch {
    return iframe.ownerDocument?.defaultView || undefined;
  }
}
