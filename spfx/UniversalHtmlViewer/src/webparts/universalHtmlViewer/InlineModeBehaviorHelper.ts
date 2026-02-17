import { CacheBusterMode, UrlValidationOptions } from './UrlHelper';
import { ContentDeliveryMode } from './UniversalHtmlViewerTypes';
import { wireInlineIframeNavigation } from './InlineNavigationHelper';
import { wireNestedIframeHydration } from './NestedIframeHydrationHelper';

export interface IInlineModeBehaviorOptions {
  contentDeliveryMode: ContentDeliveryMode;
  domElement: HTMLElement;
  pageUrl: string;
  validationOptions: UrlValidationOptions;
  cacheBusterParamName: string;
  cacheBusterMode: CacheBusterMode;
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

  return wireNestedIframeHydration({
    iframe,
    currentPageUrl: options.pageUrl,
    validationOptions: options.validationOptions,
    cacheBusterParamName: options.cacheBusterParamName,
    loadInlineHtml: options.loadInlineHtml,
  });
}
