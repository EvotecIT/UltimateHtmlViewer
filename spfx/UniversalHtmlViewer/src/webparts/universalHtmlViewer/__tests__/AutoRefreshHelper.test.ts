import { resolveAutoRefreshTarget } from '../AutoRefreshHelper';

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
});
