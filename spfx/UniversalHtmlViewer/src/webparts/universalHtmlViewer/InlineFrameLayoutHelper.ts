import { HeightMode } from './UrlHelper';

export interface IInlineFrameLayoutOptions {
  iframe: HTMLIFrameElement;
  heightMode: HeightMode;
  fixedHeightPx: number;
  fitContentWidth: boolean;
}

export function wireInlineFrameLayout(options: IInlineFrameLayoutOptions): () => void {
  let mutationObserver: MutationObserver | undefined;
  let resizeObserver: ResizeObserver | undefined;
  let resizeFallbackTimer: number | undefined;
  let rafId: number | undefined;

  const clearObservers = (): void => {
    if (mutationObserver) {
      mutationObserver.disconnect();
      mutationObserver = undefined;
    }
    if (resizeObserver) {
      resizeObserver.disconnect();
      resizeObserver = undefined;
    }
    if (resizeFallbackTimer && typeof window !== 'undefined') {
      window.clearInterval(resizeFallbackTimer);
      resizeFallbackTimer = undefined;
    }
  };

  const applyLayout = (): void => {
    const iframeDocument = tryGetIframeDocument(options.iframe);
    if (!iframeDocument || !iframeDocument.documentElement || !iframeDocument.body) {
      return;
    }

    const root = iframeDocument.documentElement;
    const body = iframeDocument.body;

    if (options.fitContentWidth) {
      root.style.removeProperty('zoom');
      const frameWidth = Math.max(0, Math.floor(options.iframe.getBoundingClientRect().width));
      const naturalWidth = Math.max(
        root.scrollWidth,
        body.scrollWidth,
        root.clientWidth,
        body.clientWidth,
        1,
      );
      if (frameWidth > 0 && naturalWidth > frameWidth) {
        const scale = frameWidth / naturalWidth;
        root.style.setProperty('zoom', scale.toString());
      }
      root.style.setProperty('overflow-x', 'hidden');
      body.style.setProperty('overflow-x', 'hidden');
    } else {
      root.style.removeProperty('zoom');
      root.style.removeProperty('overflow-x');
      body.style.removeProperty('overflow-x');
    }

    if (options.heightMode !== 'Auto') {
      return;
    }

    const minHeightPx = normalizeMinimumHeight(options.fixedHeightPx);
    const renderedHeight = Math.max(
      root.getBoundingClientRect().height,
      body.getBoundingClientRect().height,
      minHeightPx,
    );
    options.iframe.style.height = `${Math.ceil(renderedHeight)}px`;
  };

  const scheduleLayout = (): void => {
    if (typeof window === 'undefined') {
      applyLayout();
      return;
    }

    if (rafId) {
      window.cancelAnimationFrame(rafId);
    }

    rafId = window.requestAnimationFrame(() => {
      rafId = undefined;
      applyLayout();
    });
  };

  const attachObservers = (): void => {
    clearObservers();

    const iframeDocument = tryGetIframeDocument(options.iframe);
    if (!iframeDocument || !iframeDocument.documentElement || !iframeDocument.body) {
      return;
    }

    if (typeof MutationObserver !== 'undefined') {
      mutationObserver = new MutationObserver(() => {
        scheduleLayout();
      });
      mutationObserver.observe(iframeDocument.documentElement, {
        childList: true,
        subtree: true,
        characterData: true,
      });
    }

    if (typeof ResizeObserver !== 'undefined') {
      resizeObserver = new ResizeObserver(() => {
        scheduleLayout();
      });
      resizeObserver.observe(iframeDocument.documentElement);
      resizeObserver.observe(iframeDocument.body);
    } else if (typeof window !== 'undefined') {
      resizeFallbackTimer = window.setInterval(() => {
        scheduleLayout();
      }, 1500);
    }
  };

  const onLoad = (): void => {
    attachObservers();
    scheduleLayout();

    if (typeof window !== 'undefined') {
      window.setTimeout(scheduleLayout, 100);
      window.setTimeout(scheduleLayout, 400);
    }
  };

  const onWindowResize = (): void => {
    scheduleLayout();
  };

  options.iframe.addEventListener('load', onLoad);
  if (typeof window !== 'undefined') {
    window.addEventListener('resize', onWindowResize);
  }

  onLoad();

  return (): void => {
    options.iframe.removeEventListener('load', onLoad);
    if (typeof window !== 'undefined') {
      window.removeEventListener('resize', onWindowResize);
    }
    if (rafId && typeof window !== 'undefined') {
      window.cancelAnimationFrame(rafId);
      rafId = undefined;
    }
    clearObservers();
  };
}

function normalizeMinimumHeight(value: number): number {
  if (typeof value === 'number' && value > 0) {
    return Math.floor(value);
  }
  return 600;
}

function tryGetIframeDocument(iframe: HTMLIFrameElement): Document | undefined {
  try {
    return iframe.contentDocument || undefined;
  } catch {
    return undefined;
  }
}
