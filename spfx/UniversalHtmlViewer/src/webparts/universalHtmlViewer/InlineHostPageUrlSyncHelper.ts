export const INLINE_HOST_PAGE_URL_CHANGED_EVENT = 'uhv-inline-host-page-url-changed';
const INLINE_HOST_PAGE_URL_MESSAGE = 'uhv-inline-host-page-url';
const INLINE_NAVIGATION_READY_MESSAGE = 'uhv-inline-ready';

export function wireInlineHostPageUrlSync(iframe: HTMLIFrameElement): () => void {
  let isPreparedInlineDocument = hasPreparedInlineDocument(iframe);
  const postCurrentHostPageUrl = (): void => {
    if (typeof window === 'undefined' || !isPreparedInlineDocument) {
      return;
    }

    const iframeWindow = iframe.contentWindow || undefined;
    if (!iframeWindow || typeof iframeWindow.postMessage !== 'function') {
      return;
    }

    iframeWindow.postMessage(
      {
        type: INLINE_HOST_PAGE_URL_MESSAGE,
        hostPageUrl: window.location.href,
      },
      '*',
    );
  };
  const onFrameLoad = (): void => {
    isPreparedInlineDocument = hasPreparedInlineDocument(iframe);
    postCurrentHostPageUrl();
  };
  const onInlineReadyMessage = (event: MessageEvent): void => {
    const iframeWindow = iframe.contentWindow || undefined;
    if (!iframeWindow || event.source !== iframeWindow) {
      return;
    }

    const payload = event.data as { type?: unknown } | undefined;
    if (!payload || payload.type !== INLINE_NAVIGATION_READY_MESSAGE) {
      return;
    }

    isPreparedInlineDocument = true;
    postCurrentHostPageUrl();
  };

  iframe.addEventListener('load', onFrameLoad);
  if (typeof window !== 'undefined') {
    window.addEventListener('message', onInlineReadyMessage);
    window.addEventListener('popstate', postCurrentHostPageUrl);
    window.addEventListener(INLINE_HOST_PAGE_URL_CHANGED_EVENT, postCurrentHostPageUrl);
  }
  postCurrentHostPageUrl();

  return (): void => {
    iframe.removeEventListener('load', onFrameLoad);
    if (typeof window !== 'undefined') {
      window.removeEventListener('message', onInlineReadyMessage);
      window.removeEventListener('popstate', postCurrentHostPageUrl);
      window.removeEventListener(INLINE_HOST_PAGE_URL_CHANGED_EVENT, postCurrentHostPageUrl);
    }
  };
}

function hasPreparedInlineDocument(iframe: HTMLIFrameElement): boolean {
  try {
    return !!iframe.contentDocument?.querySelector(
      'script[data-uhv-inline-nav-bridge="1"]',
    );
  } catch {
    return false;
  }
}
