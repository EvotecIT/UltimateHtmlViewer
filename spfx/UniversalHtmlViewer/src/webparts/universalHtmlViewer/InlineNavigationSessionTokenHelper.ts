export const INLINE_NAVIGATION_TOKEN_ATTRIBUTE = 'data-uhv-inline-nav-token';
export const INLINE_NAVIGATION_SESSION_STAGED_EVENT = 'uhv-inline-nav-session-staged';

const INLINE_NAVIGATION_BRIDGE_SCRIPT_PATTERN =
  /<script\b[^>]*\bdata-uhv-inline-nav-bridge\s*=\s*(["'])1\1[^>]*>/i;

export interface IPreparedInlineNavigationSession {
  html: string;
  navigationToken?: string;
}

export function prepareInlineNavigationSessionHtml(
  inlineHtml: string,
): IPreparedInlineNavigationSession {
  const navigationToken = createInlineNavigationToken();
  if (!navigationToken) {
    return { html: inlineHtml };
  }

  let bridgeStamped = false;
  const preparedHtml = inlineHtml.replace(
    INLINE_NAVIGATION_BRIDGE_SCRIPT_PATTERN,
    (bridgeTag: string): string => {
      bridgeStamped = true;
      const tagWithoutExistingToken = bridgeTag.replace(
        /\sdata-uhv-inline-nav-token\s*=\s*(["'])[^"']*\1/gi,
        '',
      );
      return tagWithoutExistingToken.replace(
        />$/,
        ` ${INLINE_NAVIGATION_TOKEN_ATTRIBUTE}="${navigationToken}">`,
      );
    },
  );

  return bridgeStamped
    ? { html: preparedHtml, navigationToken }
    : { html: inlineHtml };
}

export function prepareAndStageInlineNavigationSession(
  iframe: HTMLIFrameElement,
  inlineHtml: string,
): string {
  const preparedSession = prepareInlineNavigationSessionHtml(inlineHtml);
  stageInlineNavigationSessionToken(iframe, preparedSession.navigationToken);
  return preparedSession.html;
}

export function stageInlineNavigationSessionToken(
  iframe: HTMLIFrameElement,
  navigationToken?: string,
): void {
  const normalizedToken = (navigationToken || '').trim();
  if (normalizedToken) {
    iframe.setAttribute(INLINE_NAVIGATION_TOKEN_ATTRIBUTE, normalizedToken);
  } else {
    iframe.removeAttribute(INLINE_NAVIGATION_TOKEN_ATTRIBUTE);
  }

  try {
    iframe.dispatchEvent(new Event(INLINE_NAVIGATION_SESSION_STAGED_EVENT));
  } catch {
    // Staging still succeeds for minimal test or legacy DOM implementations.
  }
}

export function getStagedInlineNavigationSessionToken(
  iframe: HTMLIFrameElement,
): string {
  try {
    return (iframe.getAttribute(INLINE_NAVIGATION_TOKEN_ATTRIBUTE) || '').trim();
  } catch {
    return '';
  }
}

function createInlineNavigationToken(): string | undefined {
  try {
    if (typeof crypto !== 'undefined' && typeof crypto.getRandomValues === 'function') {
      const values = new Uint32Array(4);
      crypto.getRandomValues(values);
      return Array.from(values)
        .map((value) => value.toString(16).padStart(8, '0'))
        .join('');
    }
  } catch {
    return undefined;
  }

  return undefined;
}
