import {
  acquireManualHistoryScrollRestoration,
  releaseManualHistoryScrollRestoration,
} from '../HistoryScrollRestorationHelper';

describe('HistoryScrollRestorationHelper', () => {
  it('restores browser history behavior only after the final viewer releases ownership', () => {
    const original = window.history.scrollRestoration;
    Object.defineProperty(window.history, 'scrollRestoration', {
      configurable: true,
      value: 'auto',
      writable: true,
    });

    try {
      expect(acquireManualHistoryScrollRestoration(window, 'viewer-a')).toBe(true);
      expect(acquireManualHistoryScrollRestoration(window, 'viewer-b')).toBe(true);
      expect(window.history.scrollRestoration).toBe('manual');

      releaseManualHistoryScrollRestoration(window, 'viewer-a');
      expect(window.history.scrollRestoration).toBe('manual');

      releaseManualHistoryScrollRestoration(window, 'viewer-b');
      expect(window.history.scrollRestoration).toBe('auto');
    } finally {
      Object.defineProperty(window.history, 'scrollRestoration', {
        configurable: true,
        value: original,
        writable: true,
      });
    }
  });
});
