/* eslint-disable @typescript-eslint/no-explicit-any, @typescript-eslint/no-var-requires */
jest.mock('@microsoft/sp-core-library', () => ({
  Version: {
    parse: () => ({}),
  },
}));
jest.mock('@microsoft/sp-http', () => ({
  SPHttpClient: {
    configurations: {
      v1: {},
    },
  },
}));
jest.mock('@microsoft/sp-lodash-subset', () => ({
  escape: (value: string): string => value,
}));
jest.mock('@microsoft/sp-property-pane', () => ({
  PropertyPaneDropdown: jest.fn(),
  PropertyPaneSlider: jest.fn(),
  PropertyPaneTextField: jest.fn(),
  PropertyPaneToggle: jest.fn(),
}));
jest.mock('@microsoft/sp-webpart-base', () => ({
  BaseClientSideWebPart: class {},
}));

const {
  default: UniversalHtmlViewerWebPart,
}: {
  default: any;
} = require('../UniversalHtmlViewerWebPart');

describe('UniversalHtmlViewerWebPart inline content options', () => {
  it('passes inferred validation prefixes to the in-frame navigation bridge', () => {
    const webPart = Object.create(UniversalHtmlViewerWebPart.prototype) as any;
    const options = webPart.getInlineContentOptions(
      {
        allowedFileExtensions: '.html,.htm',
        allowedPathPrefixes: '',
        allowQueryStringPageOverride: true,
      },
      false,
      ['/sites/Test/Shared Documents/Reports/'],
    );

    expect(options.rewriteInlineAnchorAllowedPathPrefixes).toEqual([
      '/sites/Test/Shared Documents/Reports/',
    ]);
  });
});

export {};
