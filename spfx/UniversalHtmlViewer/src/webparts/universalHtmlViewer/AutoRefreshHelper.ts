export interface IResolveAutoRefreshTargetOptions {
  baseUrl: string;
  pageUrl: string;
  currentBaseUrl?: string;
  currentPageUrl?: string;
}

export function resolveAutoRefreshTarget(
  options: IResolveAutoRefreshTargetOptions,
): { baseUrl: string; pageUrl: string } {
  const activeBaseUrl: string = (options.currentBaseUrl || '').trim() || options.baseUrl;
  const activePageUrl: string = (options.currentPageUrl || '').trim() || options.pageUrl;

  return {
    baseUrl: activeBaseUrl,
    pageUrl: activePageUrl,
  };
}
