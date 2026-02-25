export interface IResolveAutoRefreshTargetOptions {
  baseUrl: string;
  pageUrl: string;
  currentBaseUrl?: string;
  currentPageUrl?: string;
}

export interface IShouldExecuteAutoRefreshOptions {
  refreshInProgress: boolean;
  documentHidden?: boolean;
  pauseWhenHidden?: boolean;
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

export function shouldExecuteAutoRefresh(
  options: IShouldExecuteAutoRefreshOptions,
): boolean {
  if (options.refreshInProgress) {
    return false;
  }

  if (options.pauseWhenHidden !== false && options.documentHidden === true) {
    return false;
  }

  return true;
}
