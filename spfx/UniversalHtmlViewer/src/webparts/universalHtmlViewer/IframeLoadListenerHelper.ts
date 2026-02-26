export interface IIframeLoadListenerState {
  iframe?: HTMLIFrameElement;
  loadHandler?: () => void;
}

export interface ISetupIframeLoadListenerLifecycleOptions {
  state: IIframeLoadListenerState;
  iframe?: HTMLIFrameElement;
  onLoad: () => void;
}

export function clearIframeLoadListenerLifecycleState(
  state: IIframeLoadListenerState,
): void {
  if (state.iframe && state.loadHandler) {
    state.iframe.removeEventListener('load', state.loadHandler);
  }

  state.iframe = undefined;
  state.loadHandler = undefined;
}

export function setupIframeLoadListenerLifecycleState(
  options: ISetupIframeLoadListenerLifecycleOptions,
): void {
  clearIframeLoadListenerLifecycleState(options.state);

  if (!options.iframe) {
    return;
  }

  const iframe: HTMLIFrameElement = options.iframe;

  const onIframeLoad = (): void => {
    clearIframeLoadListenerLifecycleState(options.state);
    options.onLoad();
  };

  iframe.addEventListener('load', onIframeLoad);
  options.state.iframe = iframe;
  options.state.loadHandler = onIframeLoad;
}
