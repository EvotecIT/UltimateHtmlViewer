interface IHistoryScrollRestorationState {
  originalValue?: 'auto' | 'manual';
  owners: Set<string>;
}

interface IWindowWithHistoryScrollRestorationState extends Window {
  __uhvHistoryScrollRestoration?: IHistoryScrollRestorationState;
}

/**
 * Acquires shared ownership of manual browser scroll restoration.
 * The original value is restored only after the final UHV instance releases it.
 */
export function acquireManualHistoryScrollRestoration(
  targetWindow: Window,
  ownerId: string,
): boolean {
  if (!targetWindow?.history || !ownerId) {
    return false;
  }

  const stateWindow = targetWindow as IWindowWithHistoryScrollRestorationState;
  const historyObject = targetWindow.history as History & {
    scrollRestoration?: 'auto' | 'manual';
  };
  const existingState = stateWindow.__uhvHistoryScrollRestoration;
  if (existingState) {
    existingState.owners.add(ownerId);
    return true;
  }

  const originalValue =
    typeof historyObject.scrollRestoration === 'string'
      ? historyObject.scrollRestoration
      : undefined;
  try {
    historyObject.scrollRestoration = 'manual';
  } catch {
    return false;
  }

  stateWindow.__uhvHistoryScrollRestoration = {
    originalValue,
    owners: new Set<string>([ownerId]),
  };
  return true;
}

export function releaseManualHistoryScrollRestoration(
  targetWindow: Window,
  ownerId: string,
): void {
  const stateWindow = targetWindow as IWindowWithHistoryScrollRestorationState;
  const state = stateWindow.__uhvHistoryScrollRestoration;
  if (!state || !ownerId) {
    return;
  }

  state.owners.delete(ownerId);
  if (state.owners.size > 0) {
    return;
  }

  const historyObject = targetWindow.history as History & {
    scrollRestoration?: 'auto' | 'manual';
  };
  try {
    if (state.originalValue) {
      historyObject.scrollRestoration = state.originalValue;
    }
  } catch {
    // Unsupported browser contexts do not need restoration bookkeeping.
  }
  delete stateWindow.__uhvHistoryScrollRestoration;
}
