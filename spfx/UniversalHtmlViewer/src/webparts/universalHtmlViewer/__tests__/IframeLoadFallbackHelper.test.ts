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

  it('treats an already-complete iframe as loaded immediately', () => {
    const addEventListenerSpy = jest.fn();
    const removeEventListenerSpy = jest.fn();
    const iframe = {
      addEventListener: addEventListenerSpy,
      removeEventListener: removeEventListenerSpy,
      contentWindow: {
        location: {
          href: 'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Index.html',
        },
      },
      contentDocument: {
        readyState: 'complete',
      },
    } as unknown as HTMLIFrameElement;
    const state: IIframeLoadFallbackState = {};
    const onLoad = jest.fn();
    const onTimeout = jest.fn();
    const setTimeoutFn = jest.fn().mockReturnValue(147);
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

    expect(onLoad).toHaveBeenCalledTimes(1);
    expect(onTimeout).not.toHaveBeenCalled();
    expect(clearTimeoutFn).toHaveBeenCalledWith(147);
    expect(removeEventListenerSpy).toHaveBeenCalledWith('load', expect.any(Function));
    expect(state.timeoutId).toBeUndefined();
  });

  it('treats timeout on an already-complete iframe as a successful load', () => {
    const addEventListenerSpy = jest.fn();
    const removeEventListenerSpy = jest.fn();
    const iframe = {
      addEventListener: addEventListenerSpy,
      removeEventListener: removeEventListenerSpy,
      contentWindow: {
        location: {
          href: 'about:blank',
        },
      },
      contentDocument: {
        readyState: 'loading',
      },
    } as unknown as HTMLIFrameElement;
    const state: IIframeLoadFallbackState = {};
    const onLoad = jest.fn();
    const onTimeout = jest.fn();
    let timeoutHandler: (() => void) | undefined;
    const setTimeoutFn = jest.fn().mockImplementation((handler: () => void) => {
      timeoutHandler = handler;
      return 258;
    });
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

    (iframe as unknown as {
      contentWindow: { location: { href: string } };
      contentDocument: { readyState: string };
    }).contentWindow.location.href =
      'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Index.html';
    (iframe as unknown as {
      contentWindow: { location: { href: string } };
      contentDocument: { readyState: string };
    }).contentDocument.readyState = 'complete';

    if (!timeoutHandler) {
      throw new Error('Expected timeout handler to be scheduled.');
    }
    timeoutHandler();

    expect(onLoad).toHaveBeenCalledTimes(1);
    expect(onTimeout).not.toHaveBeenCalled();
    expect(clearTimeoutFn).toHaveBeenCalledWith(258);
    expect(removeEventListenerSpy).toHaveBeenCalledWith('load', expect.any(Function));
    expect(state.timeoutId).toBeUndefined();
  });

  it('does not treat a populated iframe URL as loaded while the document is still loading', () => {
    const addEventListenerSpy = jest.fn();
    const removeEventListenerSpy = jest.fn();
    const iframe = {
      addEventListener: addEventListenerSpy,
      removeEventListener: removeEventListenerSpy,
      contentWindow: {
        location: {
          href: 'https://contoso.sharepoint.com/sites/TestSite1/SiteAssets/Index.html',
        },
      },
      contentDocument: {
        readyState: 'loading',
      },
    } as unknown as HTMLIFrameElement;
    const state: IIframeLoadFallbackState = {};
    const onLoad = jest.fn();
    const onTimeout = jest.fn();
    let timeoutHandler: (() => void) | undefined;
    const setTimeoutFn = jest.fn().mockImplementation((handler: () => void) => {
      timeoutHandler = handler;
      return 369;
    });
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

    expect(onLoad).not.toHaveBeenCalled();

    if (!timeoutHandler) {
      throw new Error('Expected timeout handler to be scheduled.');
    }
    timeoutHandler();

    expect(onLoad).not.toHaveBeenCalled();
    expect(onTimeout).toHaveBeenCalledTimes(1);
    expect(clearTimeoutFn).toHaveBeenCalledWith(369);
    expect(removeEventListenerSpy).toHaveBeenCalledWith('load', expect.any(Function));
  });

  it('does not treat the initial about:blank document as a completed iframe load', () => {
    const addEventListenerSpy = jest.fn();
    const removeEventListenerSpy = jest.fn();
    const iframe = {
      addEventListener: addEventListenerSpy,
      removeEventListener: removeEventListenerSpy,
      contentWindow: {
        location: {
          href: 'about:blank',
        },
      },
      contentDocument: {
        readyState: 'complete',
      },
    } as unknown as HTMLIFrameElement;
    const state: IIframeLoadFallbackState = {};
    const onLoad = jest.fn();
    const onTimeout = jest.fn();
    let timeoutHandler: (() => void) | undefined;
    const setTimeoutFn = jest.fn().mockImplementation((handler: () => void) => {
      timeoutHandler = handler;
      return 741;
    });
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

    expect(onLoad).not.toHaveBeenCalled();

    if (!timeoutHandler) {
      throw new Error('Expected timeout handler to be scheduled.');
    }
    timeoutHandler();

    expect(onLoad).not.toHaveBeenCalled();
    expect(onTimeout).toHaveBeenCalledTimes(1);
    expect(clearTimeoutFn).toHaveBeenCalledWith(741);
    expect(removeEventListenerSpy).toHaveBeenCalledWith('load', expect.any(Function));
  });

  it('clears existing state and does not wire lifecycle when timeout is disabled', () => {
    const previousRemoveEventListenerSpy = jest.fn();
    const previousLoadHandler = jest.fn();
    const previousIframe = {
      removeEventListener: previousRemoveEventListenerSpy,
    } as unknown as HTMLIFrameElement;
    const nextAddEventListenerSpy = jest.fn();
    const nextIframe = {
      addEventListener: nextAddEventListenerSpy,
      removeEventListener: jest.fn(),
    } as unknown as HTMLIFrameElement;

    const state: IIframeLoadFallbackState = {
      timeoutId: 222,
      iframe: previousIframe,
      loadHandler: previousLoadHandler,
    };
    const setTimeoutFn = jest.fn().mockReturnValue(999);
    const clearTimeoutFn = jest.fn();

    setupIframeLoadFallbackLifecycleState({
      state,
      iframe: nextIframe,
      timeoutMs: 0,
      onLoad: jest.fn(),
      onTimeout: jest.fn(),
      setTimeoutFn,
      clearTimeoutFn,
    });

    expect(clearTimeoutFn).toHaveBeenCalledWith(222);
    expect(previousRemoveEventListenerSpy).toHaveBeenCalledWith('load', previousLoadHandler);
    expect(nextAddEventListenerSpy).not.toHaveBeenCalled();
    expect(setTimeoutFn).not.toHaveBeenCalled();
    expect(state.timeoutId).toBeUndefined();
    expect(state.iframe).toBeUndefined();
    expect(state.loadHandler).toBeUndefined();
  });

  it('clears existing state and does not wire lifecycle when iframe is missing', () => {
    const previousRemoveEventListenerSpy = jest.fn();
    const previousLoadHandler = jest.fn();
    const previousIframe = {
      removeEventListener: previousRemoveEventListenerSpy,
    } as unknown as HTMLIFrameElement;

    const state: IIframeLoadFallbackState = {
      timeoutId: 333,
      iframe: previousIframe,
      loadHandler: previousLoadHandler,
    };
    const setTimeoutFn = jest.fn().mockReturnValue(777);
    const clearTimeoutFn = jest.fn();

    setupIframeLoadFallbackLifecycleState({
      state,
      timeoutMs: 500,
      onLoad: jest.fn(),
      onTimeout: jest.fn(),
      setTimeoutFn,
      clearTimeoutFn,
    });

    expect(clearTimeoutFn).toHaveBeenCalledWith(333);
    expect(previousRemoveEventListenerSpy).toHaveBeenCalledWith('load', previousLoadHandler);
    expect(setTimeoutFn).not.toHaveBeenCalled();
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
