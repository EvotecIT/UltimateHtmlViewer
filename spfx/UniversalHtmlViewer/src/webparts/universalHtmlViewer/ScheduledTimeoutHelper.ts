export interface IScheduledTimeoutState {
  timeoutIds?: number[];
}

export interface IScheduleTimeoutWithStateOptions {
  state: IScheduledTimeoutState;
  timeoutMs: number;
  handler: () => void;
  setTimeoutFn: (handler: () => void, timeoutMs: number) => number;
}

export function clearScheduledTimeoutState(
  state: IScheduledTimeoutState,
  clearTimeoutFn: (timeoutId: number) => void,
): void {
  const timeoutIds = state.timeoutIds;
  if (!timeoutIds || timeoutIds.length === 0) {
    state.timeoutIds = [];
    return;
  }

  timeoutIds.forEach((timeoutId) => {
    clearTimeoutFn(timeoutId);
  });
  state.timeoutIds = [];
}

export function scheduleTimeoutWithState(
  options: IScheduleTimeoutWithStateOptions,
): number | undefined {
  if (options.timeoutMs < 0) {
    return undefined;
  }

  if (!options.state.timeoutIds) {
    options.state.timeoutIds = [];
  }

  let timeoutId = 0;
  const onTimeout = (): void => {
    const timeoutIds = options.state.timeoutIds || [];
    options.state.timeoutIds = timeoutIds.filter((id) => id !== timeoutId);
    options.handler();
  };

  timeoutId = options.setTimeoutFn(onTimeout, options.timeoutMs);
  options.state.timeoutIds.push(timeoutId);
  return timeoutId;
}
