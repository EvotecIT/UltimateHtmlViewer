import { getQueryStringParam } from './QueryStringHelper';
import { isUrlAllowed, UrlValidationOptions } from './UrlHelper';
import { ContentDeliveryMode } from './UniversalHtmlViewerTypes';

export const DEFAULT_INLINE_DEEP_LINK_PARAM = 'uhvPage';

export interface IResolveInlineDeepLinkTargetOptions {
  pageUrl: string;
  fallbackUrl: string;
  queryParamName?: string;
  validationOptions: UrlValidationOptions;
}

export interface IResolvedInlineContentTarget {
  allowDeepLinkOverride: boolean;
  requestedDeepLinkValue: string;
  hasRequestedDeepLink: boolean;
  deepLinkedUrl?: string;
  initialContentUrl: string;
  isRejectedRequestedDeepLink: boolean;
}

export function resolveInlineDeepLinkTarget(
  options: IResolveInlineDeepLinkTargetOptions,
): string | undefined {
  const paramName: string =
    (options.queryParamName || '').trim() || DEFAULT_INLINE_DEEP_LINK_PARAM;
  const rawValue: string | undefined = getQueryStringParam(options.pageUrl, paramName);
  const value: string = (rawValue || '').trim();
  if (!value) {
    return undefined;
  }

  try {
    const fallbackAbsolute = toAbsoluteUrl(options.fallbackUrl, options.pageUrl);
    const target = new URL(value, fallbackAbsolute);
    const normalizedTarget = target.toString();
    if (!isUrlAllowed(normalizedTarget, options.validationOptions)) {
      return undefined;
    }
    return normalizedTarget;
  } catch {
    return undefined;
  }
}

export function resolveInlineContentTarget(
  options: IResolveInlineDeepLinkTargetOptions,
): IResolvedInlineContentTarget {
  const paramName: string =
    (options.queryParamName || '').trim() || DEFAULT_INLINE_DEEP_LINK_PARAM;
  const requestedDeepLinkValue: string =
    (getQueryStringParam(options.pageUrl, paramName) || '').trim();
  const hasRequestedDeepLink: boolean = requestedDeepLinkValue.length > 0;
  const allowDeepLinkOverride: boolean =
    options.validationOptions.securityMode !== 'AnyHttps';
  const deepLinkedUrl: string | undefined = allowDeepLinkOverride
    ? resolveInlineDeepLinkTarget(options)
    : undefined;
  const initialContentUrl: string = deepLinkedUrl || options.fallbackUrl;
  const isRejectedRequestedDeepLink: boolean =
    hasRequestedDeepLink && allowDeepLinkOverride && !deepLinkedUrl;

  return {
    allowDeepLinkOverride,
    requestedDeepLinkValue,
    hasRequestedDeepLink,
    deepLinkedUrl,
    initialContentUrl,
    isRejectedRequestedDeepLink,
  };
}

export interface IBuildPageUrlWithInlineDeepLinkOptions {
  pageUrl: string;
  targetUrl: string;
  queryParamName?: string;
}

export interface IBuildOpenInNewTabUrlOptions {
  resolvedUrl: string;
  baseUrl: string;
  pageUrl: string;
  currentPageUrl?: string;
  contentDeliveryMode: ContentDeliveryMode;
}

export function buildOpenInNewTabUrl(
  options: IBuildOpenInNewTabUrlOptions,
): string {
  if (options.contentDeliveryMode !== 'SharePointFileContent') {
    return options.resolvedUrl;
  }

  const effectivePageUrl: string = (options.currentPageUrl || '').trim() || options.pageUrl;
  if (!effectivePageUrl) {
    return options.resolvedUrl;
  }

  const pageDeepLinkUrl = buildPageUrlWithInlineDeepLink({
    pageUrl: effectivePageUrl,
    targetUrl: options.baseUrl,
    queryParamName: DEFAULT_INLINE_DEEP_LINK_PARAM,
  });
  if (!pageDeepLinkUrl) {
    return options.resolvedUrl;
  }

  return pageDeepLinkUrl;
}

export function buildPageUrlWithInlineDeepLink(
  options: IBuildPageUrlWithInlineDeepLinkOptions,
): string | undefined {
  const paramName: string =
    (options.queryParamName || '').trim() || DEFAULT_INLINE_DEEP_LINK_PARAM;
  if (!paramName) {
    return undefined;
  }

  try {
    const current = new URL(options.pageUrl);
    const target = new URL(options.targetUrl, current.toString());

    const encodedTarget =
      target.host.toLowerCase() === current.host.toLowerCase()
        ? `${target.pathname}${target.search}${target.hash}`
        : target.toString();

    current.searchParams.set(paramName, encodedTarget);
    return current.toString();
  } catch {
    return undefined;
  }
}

export interface IBuildPageUrlWithoutInlineDeepLinkOptions {
  pageUrl: string;
  queryParamName?: string;
}

export function buildPageUrlWithoutInlineDeepLink(
  options: IBuildPageUrlWithoutInlineDeepLinkOptions,
): string | undefined {
  const paramName: string =
    (options.queryParamName || '').trim() || DEFAULT_INLINE_DEEP_LINK_PARAM;
  if (!paramName) {
    return undefined;
  }

  try {
    const current = new URL(options.pageUrl);
    current.searchParams.delete(paramName);
    return current.toString();
  } catch {
    return undefined;
  }
}

function toAbsoluteUrl(value: string, pageUrl: string): string {
  if (!value) {
    throw new Error('Missing URL');
  }

  if (value.startsWith('/')) {
    const current = new URL(pageUrl);
    return new URL(value, current.origin).toString();
  }

  try {
    return new URL(value).toString();
  } catch {
    return new URL(value, pageUrl).toString();
  }
}
