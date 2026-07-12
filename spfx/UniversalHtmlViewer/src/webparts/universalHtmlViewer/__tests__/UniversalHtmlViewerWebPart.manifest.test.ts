/* eslint-disable @typescript-eslint/no-var-requires */

describe('UniversalHtmlViewerWebPart manifest defaults', () => {
  it('keeps nested inline-frame hydration compatible by default', () => {
    const manifest = require('../UniversalHtmlViewerWebPart.manifest.json');
    const properties = manifest.preconfiguredEntries[0].properties;

    expect(properties.configurationPreset).toBe('SharePointLibraryRelaxed');
    expect(properties.sandboxPreset).toBe('Relaxed');
    expect(properties.enforceStrictInlineCsp).toBe(false);
  });
});

export {};
