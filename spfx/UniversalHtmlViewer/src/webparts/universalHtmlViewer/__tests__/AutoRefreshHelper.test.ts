import { resolveAutoRefreshTarget, shouldExecuteAutoRefresh } from '../AutoRefreshHelper';

describe('AutoRefreshHelper', () => {
  it('uses current navigation target when available', () => {
    const result = resolveAutoRefreshTarget({
      baseUrl: 'https://contoso.sharepoint.com/sites/Reports/SiteAssets/report-old.html',
      pageUrl: 'https://contoso.sharepoint.com/sites/Reports/SitePages/Dashboard.aspx?stale=1',
      currentBaseUrl: 'https://contoso.sharepoint.com/sites/Reports/SiteAssets/report-new.html',
      currentPageUrl: 'https://contoso.sharepoint.com/sites/Reports/SitePages/Dashboard.aspx',
    });

    expect(result).toEqual({
      baseUrl: 'https://contoso.sharepoint.com/sites/Reports/SiteAssets/report-new.html',
      pageUrl: 'https://contoso.sharepoint.com/sites/Reports/SitePages/Dashboard.aspx',
    });
  });

  it('falls back to configured values when current values are empty', () => {
    const result = resolveAutoRefreshTarget({
      baseUrl: 'https://contoso.sharepoint.com/sites/Reports/SiteAssets/report-default.html',
      pageUrl: 'https://contoso.sharepoint.com/sites/Reports/SitePages/Dashboard.aspx',
      currentBaseUrl: '   ',
      currentPageUrl: '',
    });

    expect(result).toEqual({
      baseUrl: 'https://contoso.sharepoint.com/sites/Reports/SiteAssets/report-default.html',
      pageUrl: 'https://contoso.sharepoint.com/sites/Reports/SitePages/Dashboard.aspx',
    });
  });

  it('skips auto-refresh while a refresh is already in progress', () => {
    const shouldExecute = shouldExecuteAutoRefresh({
      refreshInProgress: true,
      documentHidden: false,
    });

    expect(shouldExecute).toBe(false);
  });

  it('skips auto-refresh when page is hidden by default', () => {
    const shouldExecute = shouldExecuteAutoRefresh({
      refreshInProgress: false,
      documentHidden: true,
    });

    expect(shouldExecute).toBe(false);
  });

  it('allows auto-refresh when hidden pause is disabled', () => {
    const shouldExecute = shouldExecuteAutoRefresh({
      refreshInProgress: false,
      documentHidden: true,
      pauseWhenHidden: false,
    });

    expect(shouldExecute).toBe(true);
  });
});
