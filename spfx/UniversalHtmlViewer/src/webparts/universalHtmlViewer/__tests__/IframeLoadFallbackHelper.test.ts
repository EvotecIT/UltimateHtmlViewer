import {
  clearIframeLoadFallbackLifecycleState,
  IIframeLoadFallbackState,
  setupIframeLoadFallbackLifecycleState,
} from '../IframeLoadFallbackHelper';

describe('IframeLoadFallbackHelper', () => {
  it('wires load listener and timeout callback', () => {
    const addEventListenerSpy = jest.fn();
    const removeEventListenerSpy = jest.fn();
    const iframe = {
      addEventListener: addEventListenerSpy,
      removeEventListener: removeEventListenerSpy,
    } as unknown as HTMLIFrameElement;
    const state: IIframeLoadFallbackState = {};
    const onLoad = jest.fn();
    const onTimeout = jest.fn();
    const setTimeoutFn = jest.fn().mockReturnValue(123);
    const clearTimeoutFn = jest.fn();

    setupIframeLoadFallbackLifecycleState({
      state,
      iframe,
      timeoutMs: 1000,
      onLoad,
      onTimeout,
      setTimeoutFn,
      clearTimeoutFn,
    });

    expect(addEventListenerSpy).toHaveBeenCalledWith('load', expect.any(Function));
    expect(setTimeoutFn).toHaveBeenCalledWith(expect.any(Function), 1000);
    expect(state.timeoutId).toBe(123);
    expect(state.iframe).toBe(iframe);
    expect(state.loadHandler).toEqual(expect.any(Function));
    expect(onLoad).not.toHaveBeenCalled();
    expect(onTimeout).not.toHaveBeenCalled();
  });

  it('clears timeout and listener when iframe load fires', () => {
    const addEventListenerSpy = jest.fn();
    const removeEventListenerSpy = jest.fn();
    const iframe = {
      addEventListener: addEventListenerSpy,
      removeEventListener: removeEventListenerSpy,
    } as unknown as HTMLIFrameElement;
    const state: IIframeLoadFallbackState = {};
    const onLoad = jest.fn();
    const setTimeoutFn = jest.fn().mockReturnValue(456);
    const clearTimeoutFn = jest.fn();

    setupIframeLoadFallbackLifecycleState({
      state,
      iframe,
      timeoutMs: 500,
      onLoad,
      onTimeout: jest.fn(),
      setTimeoutFn,
      clearTimeoutFn,
    });

    const loadHandler = state.loadHandler;
    if (!loadHandler) {
      throw new Error('Expected load handler to be wired.');
    }
    loadHandler();

    expect(clearTimeoutFn).toHaveBeenCalledWith(456);
    expect(removeEventListenerSpy).toHaveBeenCalledWith('load', loadHandler);
    expect(onLoad).toHaveBeenCalledTimes(1);
    expect(state.timeoutId).toBeUndefined();
    expect(state.iframe).toBeUndefined();
    expect(state.loadHandler).toBeUndefined();
  });

  it('clears timeout and listener when timeout callback fires', () => {
    const addEventListenerSpy = jest.fn();
    const removeEventListenerSpy = jest.fn();
    const iframe = {
      addEventListener: addEventListenerSpy,
      removeEventListener: removeEventListenerSpy,
    } as unknown as HTMLIFrameElement;
    const state: IIframeLoadFallbackState = {};
    const onTimeout = jest.fn();
    let timeoutHandler: (() => void) | undefined;
    const setTimeoutFn = jest.fn().mockImplementation((handler: () => void) => {
      timeoutHandler = handler;
      return 789;
    });
    const clearTimeoutFn = jest.fn();

    setupIframeLoadFallbackLifecycleState({
      state,
      iframe,
      timeoutMs: 250,
      onLoad: jest.fn(),
      onTimeout,
      setTimeoutFn,
      clearTimeoutFn,
    });

    if (!timeoutHandler) {
      throw new Error('Expected timeout handler to be scheduled.');
    }
    timeoutHandler();

    expect(clearTimeoutFn).toHaveBeenCalledWith(789);
    expect(removeEventListenerSpy).toHaveBeenCalledWith('load', expect.any(Function));
    expect(onTimeout).toHaveBeenCalledTimes(1);
    expect(state.timeoutId).toBeUndefined();
    expect(state.iframe).toBeUndefined();
    expect(state.loadHandler).toBeUndefined();
  });

  it('clears existing lifecycle state explicitly', () => {
    const removeEventListenerSpy = jest.fn();
    const state: IIframeLoadFallbackState = {
      timeoutId: 321,
      iframe: {
        removeEventListener: removeEventListenerSpy,
      } as unknown as HTMLIFrameElement,
      loadHandler: jest.fn(),
    };
    const clearTimeoutFn = jest.fn();

    clearIframeLoadFallbackLifecycleState(state, clearTimeoutFn);

    expect(clearTimeoutFn).toHaveBeenCalledWith(321);
    expect(removeEventListenerSpy).toHaveBeenCalledWith('load', expect.any(Function));
    expect(state.timeoutId).toBeUndefined();
    expect(state.iframe).toBeUndefined();
    expect(state.loadHandler).toBeUndefined();
  });
});
