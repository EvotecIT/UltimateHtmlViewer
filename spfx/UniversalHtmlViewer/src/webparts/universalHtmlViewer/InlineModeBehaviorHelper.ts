import { CacheBusterMode, HeightMode, UrlValidationOptions } from './UrlHelper';
import { ContentDeliveryMode } from './UniversalHtmlViewerTypes';
import { wireInlineFrameLayout } from './InlineFrameLayoutHelper';
import { wireInlineIframeNavigation } from './InlineNavigationHelper';
import { wireNestedIframeHydration } from './NestedIframeHydrationHelper';

export interface IInlineModeBehaviorOptions {
  contentDeliveryMode: ContentDeliveryMode;
  domElement: HTMLElement;
  pageUrl: string;
  validationOptions: UrlValidationOptions;
  cacheBusterParamName: string;
  cacheBusterMode: CacheBusterMode;
  heightMode: HeightMode;
  fixedHeightPx: number;
  fitContentWidth: boolean;
  onNavigate: (targetUrl: string, cacheBusterMode: CacheBusterMode) => void;
  loadInlineHtml: (
    sourceUrl: string,
    baseUrlForRelativeLinks: string,
  ) => Promise<string | undefined>;
}

export function applyInlineModeBehaviors(
  options: IInlineModeBehaviorOptions,
): (() => void) | undefined {
  if (options.contentDeliveryMode !== 'SharePointFileContent') {
    return undefined;
  }

  const iframe: HTMLIFrameElement | null = options.domElement.querySelector('iframe');
  if (!iframe) {
    return undefined;
  }

  wireInlineIframeNavigation({
    iframe,
    currentPageUrl: options.pageUrl,
    validationOptions: options.validationOptions,
    cacheBusterParamName: options.cacheBusterParamName,
    onNavigate: (targetUrl: string) => {
      options.onNavigate(targetUrl, options.cacheBusterMode);
    },
  });

  const hydrationCleanup = wireNestedIframeHydration({
    iframe,
    currentPageUrl: options.pageUrl,
    validationOptions: options.validationOptions,
    cacheBusterParamName: options.cacheBusterParamName,
    loadInlineHtml: options.loadInlineHtml,
  });

  const layoutCleanup = wireInlineFrameLayout({
    iframe,
    heightMode: options.heightMode,
    fixedHeightPx: options.fixedHeightPx,
    fitContentWidth: options.fitContentWidth,
  });

  return (): void => {
    hydrationCleanup();
    layoutCleanup();
  };
}
