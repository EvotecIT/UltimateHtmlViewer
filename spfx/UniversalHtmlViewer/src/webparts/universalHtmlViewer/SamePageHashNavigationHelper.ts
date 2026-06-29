export function isSamePageHashHref(value: string): boolean {
  return (value || '').trim().startsWith('#');
}

export function navigateToSamePageHash(
  ownerDocument: Document | undefined,
  hashHref: string,
  requireTarget: boolean,
): boolean {
  if (!ownerDocument) {
    return false;
  }

  const targetElement = findSamePageHashTarget(ownerDocument, hashHref);
  if (requireTarget && !targetElement) {
    return false;
  }

  const targetWindow = ownerDocument.defaultView || undefined;
  if (targetWindow?.location) {
    try {
      targetWindow.location.hash = hashHref;
      return true;
    } catch {
      // Fall back to direct element scrolling below.
    }
  }

  if (targetElement && typeof targetElement.scrollIntoView === 'function') {
    targetElement.scrollIntoView();
    return true;
  }

  if (hashHref === '#' && targetWindow && typeof targetWindow.scrollTo === 'function') {
    try {
      targetWindow.scrollTo(0, 0);
      return true;
    } catch {
      return false;
    }
  }

  return !requireTarget;
}

export function findSamePageHashTarget(ownerDocument: Document, hashHref: string): HTMLElement | undefined {
  if (hashHref === '#') {
    return ownerDocument.documentElement || ownerDocument.body || undefined;
  }

  const fragment = decodeHashFragment(hashHref.substring(1));
  if (!fragment) {
    return undefined;
  }

  return (
    ownerDocument.getElementById(fragment) ||
    (ownerDocument.getElementsByName(fragment)[0] as HTMLElement | undefined)
  );
}

function decodeHashFragment(fragment: string): string {
  try {
    return decodeURIComponent(fragment);
  } catch {
    return fragment;
  }
}
