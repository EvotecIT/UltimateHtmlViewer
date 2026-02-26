export interface IIframeLoadFallbackState {
  timeoutId?: number;
  iframe?: HTMLIFrameElement;
  loadHandler?: () => void;
}

export interface ISetupIframeLoadFallbackLifecycleOptions {
  state: IIframeLoadFallbackState;
  iframe?: HTMLIFrameElement;
  timeoutMs: number;
  onLoad: () => void;
  onTimeout: () => void;
  setTimeoutFn: (handler: () => void, timeoutMs: number) => number;
  clearTimeoutFn: (timeoutId: number) => void;
}

export function clearIframeLoadFallbackLifecycleState(
  state: IIframeLoadFallbackState,
  clearTimeoutFn: (timeoutId: number) => void,
): void {
  if (typeof state.timeoutId === 'number') {
    clearTimeoutFn(state.timeoutId);
  }

  if (state.iframe && state.loadHandler) {
    state.iframe.removeEventListener('load', state.loadHandler);
  }

  state.timeoutId = undefined;
  state.iframe = undefined;
  state.loadHandler = undefined;
}

export function setupIframeLoadFallbackLifecycleState(
  options: ISetupIframeLoadFallbackLifecycleOptions,
): void {
  clearIframeLoadFallbackLifecycleState(options.state, options.clearTimeoutFn);

  if (!options.iframe || options.timeoutMs <= 0) {
    return;
  }

  const iframe: HTMLIFrameElement = options.iframe;

  const onIframeLoad = (): void => {
    clearIframeLoadFallbackLifecycleState(options.state, options.clearTimeoutFn);
    options.onLoad();
  };
  iframe.addEventListener('load', onIframeLoad);
  options.state.iframe = iframe;
  options.state.loadHandler = onIframeLoad;

  options.state.timeoutId = options.setTimeoutFn(() => {
    clearIframeLoadFallbackLifecycleState(options.state, options.clearTimeoutFn);
    options.onTimeout();
  }, options.timeoutMs);
}
