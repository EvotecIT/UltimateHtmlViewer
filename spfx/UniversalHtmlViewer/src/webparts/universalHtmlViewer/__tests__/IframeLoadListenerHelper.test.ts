import {
  clearIframeLoadListenerLifecycleState,
  IIframeLoadListenerState,
  setupIframeLoadListenerLifecycleState,
} from '../IframeLoadListenerHelper';

describe('IframeLoadListenerHelper', () => {
  it('wires a load listener and tracks lifecycle state', () => {
    const addEventListenerSpy = jest.fn();
    const iframe = {
      addEventListener: addEventListenerSpy,
      removeEventListener: jest.fn(),
    } as unknown as HTMLIFrameElement;
    const state: IIframeLoadListenerState = {};
    const onLoad = jest.fn();

    setupIframeLoadListenerLifecycleState({
      state,
      iframe,
      onLoad,
    });

    expect(addEventListenerSpy).toHaveBeenCalledWith('load', expect.any(Function));
    expect(state.iframe).toBe(iframe);
    expect(state.loadHandler).toEqual(expect.any(Function));
    expect(onLoad).not.toHaveBeenCalled();
  });

  it('clears existing listener before wiring a new one', () => {
    const firstAddEventListenerSpy = jest.fn();
    const firstRemoveEventListenerSpy = jest.fn();
    const firstIframe = {
      addEventListener: firstAddEventListenerSpy,
      removeEventListener: firstRemoveEventListenerSpy,
    } as unknown as HTMLIFrameElement;

    const secondAddEventListenerSpy = jest.fn();
    const secondIframe = {
      addEventListener: secondAddEventListenerSpy,
      removeEventListener: jest.fn(),
    } as unknown as HTMLIFrameElement;

    const state: IIframeLoadListenerState = {};
    setupIframeLoadListenerLifecycleState({
      state,
      iframe: firstIframe,
      onLoad: jest.fn(),
    });
    const firstHandler = state.loadHandler;
    if (!firstHandler) {
      throw new Error('Expected first load handler to be wired.');
    }

    setupIframeLoadListenerLifecycleState({
      state,
      iframe: secondIframe,
      onLoad: jest.fn(),
    });

    expect(firstRemoveEventListenerSpy).toHaveBeenCalledWith('load', firstHandler);
    expect(secondAddEventListenerSpy).toHaveBeenCalledWith('load', expect.any(Function));
    expect(state.iframe).toBe(secondIframe);
    expect(state.loadHandler).toEqual(expect.any(Function));
  });

  it('clears listener and state when load event fires', () => {
    const addEventListenerSpy = jest.fn();
    const removeEventListenerSpy = jest.fn();
    const iframe = {
      addEventListener: addEventListenerSpy,
      removeEventListener: removeEventListenerSpy,
    } as unknown as HTMLIFrameElement;
    const state: IIframeLoadListenerState = {};
    const onLoad = jest.fn();

    setupIframeLoadListenerLifecycleState({
      state,
      iframe,
      onLoad,
    });

    const loadHandler = state.loadHandler;
    if (!loadHandler) {
      throw new Error('Expected load handler to be wired.');
    }
    loadHandler();

    expect(removeEventListenerSpy).toHaveBeenCalledWith('load', loadHandler);
    expect(onLoad).toHaveBeenCalledTimes(1);
    expect(state.iframe).toBeUndefined();
    expect(state.loadHandler).toBeUndefined();
  });

  it('clears listener lifecycle state explicitly', () => {
    const removeEventListenerSpy = jest.fn();
    const loadHandler = jest.fn();
    const state: IIframeLoadListenerState = {
      iframe: {
        removeEventListener: removeEventListenerSpy,
      } as unknown as HTMLIFrameElement,
      loadHandler,
    };

    clearIframeLoadListenerLifecycleState(state);

    expect(removeEventListenerSpy).toHaveBeenCalledWith('load', loadHandler);
    expect(state.iframe).toBeUndefined();
    expect(state.loadHandler).toBeUndefined();
  });
});
