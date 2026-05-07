import { extractTitleFromHtml, normalizePageTitle } from '../PageTitleHelper';

describe('PageTitleHelper', () => {
  it('extracts and normalizes html document titles', () => {
    const title = extractTitleFromHtml(
      '<html><head><title>  Active Directory   Overall - Computers  </title></head><body></body></html>',
    );

    expect(title).toBe('Active Directory Overall - Computers');
  });

  it('decodes basic title entities when DOMParser is unavailable', () => {
    const typedGlobal = globalThis as typeof globalThis & {
      DOMParser?: typeof DOMParser;
    };
    const originalDomParser = typedGlobal.DOMParser;
    try {
      Object.defineProperty(typedGlobal, 'DOMParser', {
        configurable: true,
        value: undefined,
      });

      expect(extractTitleFromHtml('<title>Users &amp; Computers</title>')).toBe(
        'Users & Computers',
      );
    } finally {
      Object.defineProperty(typedGlobal, 'DOMParser', {
        configurable: true,
        value: originalDomParser,
      });
    }
  });

  it('returns an empty title when html has no title element', () => {
    expect(extractTitleFromHtml('<html><body>No title</body></html>')).toBe('');
  });

  it('normalizes whitespace-only values to empty strings', () => {
    expect(normalizePageTitle(' \n\t ')).toBe('');
  });
});

export {};
