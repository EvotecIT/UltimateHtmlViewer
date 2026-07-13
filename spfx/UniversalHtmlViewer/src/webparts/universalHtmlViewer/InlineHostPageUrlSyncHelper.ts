import {
  getStagedInlineNavigationSessionToken,
  INLINE_NAVIGATION_SESSION_STAGED_EVENT,
  INLINE_NAVIGATION_TOKEN_ATTRIBUTE,
} from './InlineNavigationSessionTokenHelper';

export const INLINE_HOST_PAGE_URL_CHANGED_EVENT = 'uhv-inline-host-page-url-changed';
const INLINE_HOST_PAGE_URL_MESSAGE = 'uhv-inline-host-page-url';
const INLINE_NAVIGATION_READY_MESSAGE = 'uhv-inline-ready';

export interface IInlineHostPageUrlSyncHandle {
  cleanup: () => void;
  isAuthenticatedBridgeMessage: (event: MessageEvent) => boolean;
}

export function wireInlineHostPageUrlSync(
  iframe: HTMLIFrameElement,
): IInlineHostPageUrlSyncHandle {
  let expectedNavigationToken = getStagedInlineNavigationSessionToken(iframe);
  let activeNavigationToken = '';
  let expectedLoadObserved = false;
  let isPreparedInlineDocument =
    !expectedNavigationToken && hasPreparedInlineDocument(iframe);
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
  const onInlineNavigationSessionStaged = (): void => {
    expectedNavigationToken = getStagedInlineNavigationSessionToken(iframe);
    activeNavigationToken = '';
    expectedLoadObserved = false;
    isPreparedInlineDocument = false;
  };
  const onFrameLoad = (): void => {
    const stagedNavigationToken = getStagedInlineNavigationSessionToken(iframe);
    if (stagedNavigationToken) {
      if (stagedNavigationToken !== expectedNavigationToken) {
        expectedNavigationToken = stagedNavigationToken;
        activeNavigationToken = '';
      }
      expectedLoadObserved = true;
      isPreparedInlineDocument = activeNavigationToken === expectedNavigationToken;
      if (isPreparedInlineDocument) {
        iframe.removeAttribute(INLINE_NAVIGATION_TOKEN_ATTRIBUTE);
      }
    } else {
      expectedNavigationToken = '';
      activeNavigationToken = '';
      expectedLoadObserved = false;
      isPreparedInlineDocument = hasPreparedInlineDocument(iframe);
    }
    postCurrentHostPageUrl();
  };
  const onInlineReadyMessage = (event: MessageEvent): void => {
    const iframeWindow = iframe.contentWindow || undefined;
    if (!iframeWindow || event.source !== iframeWindow) {
      return;
    }

    const payload = event.data as
      | { type?: unknown; navigationToken?: unknown }
      | undefined;
    if (!payload || payload.type !== INLINE_NAVIGATION_READY_MESSAGE) {
      return;
    }

    const suppliedNavigationToken =
      typeof payload.navigationToken === 'string'
        ? payload.navigationToken.trim()
        : '';
    if (expectedNavigationToken) {
      if (suppliedNavigationToken !== expectedNavigationToken) {
        return;
      }
      activeNavigationToken = expectedNavigationToken;
      if (expectedLoadObserved) {
        iframe.removeAttribute(INLINE_NAVIGATION_TOKEN_ATTRIBUTE);
      }
    } else if (!hasPreparedInlineDocument(iframe)) {
      return;
    }

    isPreparedInlineDocument = true;
    postCurrentHostPageUrl();
  };
  const isAuthenticatedBridgeMessage = (event: MessageEvent): boolean => {
    const iframeWindow = iframe.contentWindow || undefined;
    if (!iframeWindow || event.source !== iframeWindow || !isPreparedInlineDocument) {
      return false;
    }

    const payload = event.data as { navigationToken?: unknown } | undefined;
    const suppliedNavigationToken =
      typeof payload?.navigationToken === 'string'
        ? payload.navigationToken.trim()
        : '';
    if (activeNavigationToken) {
      return suppliedNavigationToken === activeNavigationToken;
    }

    return hasPreparedInlineDocument(iframe);
  };

  iframe.addEventListener('load', onFrameLoad);
  iframe.addEventListener(
    INLINE_NAVIGATION_SESSION_STAGED_EVENT,
    onInlineNavigationSessionStaged,
  );
  if (typeof window !== 'undefined') {
    window.addEventListener('message', onInlineReadyMessage);
    window.addEventListener('popstate', postCurrentHostPageUrl);
    window.addEventListener(INLINE_HOST_PAGE_URL_CHANGED_EVENT, postCurrentHostPageUrl);
  }
  postCurrentHostPageUrl();

  const cleanup = (): void => {
    iframe.removeEventListener('load', onFrameLoad);
    iframe.removeEventListener(
      INLINE_NAVIGATION_SESSION_STAGED_EVENT,
      onInlineNavigationSessionStaged,
    );
    if (typeof window !== 'undefined') {
      window.removeEventListener('message', onInlineReadyMessage);
      window.removeEventListener('popstate', postCurrentHostPageUrl);
      window.removeEventListener(INLINE_HOST_PAGE_URL_CHANGED_EVENT, postCurrentHostPageUrl);
    }
  };

  return {
    cleanup,
    isAuthenticatedBridgeMessage,
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
