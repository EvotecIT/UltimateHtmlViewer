import {
  clearScheduledTimeoutState,
  IScheduledTimeoutState,
  scheduleTimeoutWithState,
} from '../ScheduledTimeoutHelper';

describe('ScheduledTimeoutHelper', () => {
  it('tracks scheduled timeout ids', () => {
    const state: IScheduledTimeoutState = {};
    const setTimeoutFn = jest.fn().mockReturnValue(101);

    const timeoutId = scheduleTimeoutWithState({
      state,
      timeoutMs: 25,
      handler: jest.fn(),
      setTimeoutFn,
    });

    expect(timeoutId).toBe(101);
    expect(setTimeoutFn).toHaveBeenCalledWith(expect.any(Function), 25);
    expect(state.timeoutIds).toEqual([101]);
  });

  it('removes timeout id from state after callback runs', () => {
    const state: IScheduledTimeoutState = {};
    let scheduledHandler: (() => void) | undefined;
    const setTimeoutFn = jest.fn().mockImplementation((handler: () => void) => {
      scheduledHandler = handler;
      return 202;
    });
    const handler = jest.fn();

    scheduleTimeoutWithState({
      state,
      timeoutMs: 10,
      handler,
      setTimeoutFn,
    });

    if (!scheduledHandler) {
      throw new Error('Expected timeout handler to be scheduled.');
    }
    scheduledHandler();

    expect(handler).toHaveBeenCalledTimes(1);
    expect(state.timeoutIds).toEqual([]);
  });

  it('clears all tracked timeout ids', () => {
    const state: IScheduledTimeoutState = {
      timeoutIds: [1, 2, 3],
    };
    const clearTimeoutFn = jest.fn();

    clearScheduledTimeoutState(state, clearTimeoutFn);

    expect(clearTimeoutFn).toHaveBeenCalledTimes(3);
    expect(clearTimeoutFn).toHaveBeenNthCalledWith(1, 1);
    expect(clearTimeoutFn).toHaveBeenNthCalledWith(2, 2);
    expect(clearTimeoutFn).toHaveBeenNthCalledWith(3, 3);
    expect(state.timeoutIds).toEqual([]);
  });

  it('does not schedule negative timeout values', () => {
    const state: IScheduledTimeoutState = {};
    const setTimeoutFn = jest.fn();

    const timeoutId = scheduleTimeoutWithState({
      state,
      timeoutMs: -1,
      handler: jest.fn(),
      setTimeoutFn,
    });

    expect(timeoutId).toBeUndefined();
    expect(setTimeoutFn).not.toHaveBeenCalled();
    expect(state.timeoutIds).toBeUndefined();
  });
});
