export const INLINE_HOST_PAGE_URL_CHANGED_EVENT = 'uhv-inline-host-page-url-changed';
const INLINE_HOST_PAGE_URL_MESSAGE = 'uhv-inline-host-page-url';

export function wireInlineHostPageUrlSync(iframe: HTMLIFrameElement): () => void {
  const postCurrentHostPageUrl = (): void => {
    if (typeof window === 'undefined') {
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

  iframe.addEventListener('load', postCurrentHostPageUrl);
  if (typeof window !== 'undefined') {
    window.addEventListener('popstate', postCurrentHostPageUrl);
    window.addEventListener(INLINE_HOST_PAGE_URL_CHANGED_EVENT, postCurrentHostPageUrl);
  }
  postCurrentHostPageUrl();

  return (): void => {
    iframe.removeEventListener('load', postCurrentHostPageUrl);
    if (typeof window !== 'undefined') {
      window.removeEventListener('popstate', postCurrentHostPageUrl);
      window.removeEventListener(INLINE_HOST_PAGE_URL_CHANGED_EVENT, postCurrentHostPageUrl);
    }
  };
}
