import {
  applyImportedConfigToProps,
  buildConfigExport,
} from '../ConfigImportExportHelper';

describe('applyImportedConfigToProps', () => {
  it('applies valid values and normalizes enum casing', () => {
    const propsRecord: Record<string, unknown> = {};
    const result = applyImportedConfigToProps(propsRecord, {
      configurationPreset: 'sharepointlibraryfullpage',
      contentDeliveryMode: 'sharepointfilebloburl',
      htmlSourceMode: 'basepathandrelativepath',
      heightMode: 'auto',
      fixedHeightPx: '700',
      fitContentWidth: 'yes',
      enableExpertSecurityModes: 'true',
      refreshIntervalMinutes: 10,
      inlineContentCacheTtlSeconds: 20,
      showChrome: 'true',
      iframeLoading: 'LAZY',
      chromeDensity: 'compact',
      allowQueryStringPageOverride: 'false',
      enforceStrictInlineCsp: 'true',
      showReportBrowser: 'true',
      reportBrowserRootPath: '/sites/TestSite1/SiteAssets/Reports',
      reportBrowserDefaultView: 'files',
      reportBrowserMaxItems: '250',
    });

    expect(result.ignoredEntries).toHaveLength(0);
    expect(propsRecord.configurationPreset).toBe('SharePointLibraryFullPage');
    expect(propsRecord.contentDeliveryMode).toBe('SharePointFileBlobUrl');
    expect(propsRecord.htmlSourceMode).toBe('BasePathAndRelativePath');
    expect(propsRecord.heightMode).toBe('Auto');
    expect(propsRecord.fixedHeightPx).toBe(700);
    expect(propsRecord.fitContentWidth).toBe(true);
    expect(propsRecord.enableExpertSecurityModes).toBe(true);
    expect(propsRecord.refreshIntervalMinutes).toBe(10);
    expect(propsRecord.inlineContentCacheTtlSeconds).toBe(20);
    expect(propsRecord.showChrome).toBe(true);
    expect(propsRecord.iframeLoading).toBe('lazy');
    expect(propsRecord.chromeDensity).toBe('Compact');
    expect(propsRecord.allowQueryStringPageOverride).toBe(false);
    expect(propsRecord.enforceStrictInlineCsp).toBe(true);
    expect(propsRecord.showReportBrowser).toBe(true);
    expect(propsRecord.reportBrowserRootPath).toBe('/sites/TestSite1/SiteAssets/Reports');
    expect(propsRecord.reportBrowserDefaultView).toBe('Files');
    expect(propsRecord.reportBrowserMaxItems).toBe(250);
  });

  it('ignores unknown keys and invalid values', () => {
    const propsRecord: Record<string, unknown> = {};
    const result = applyImportedConfigToProps(propsRecord, {
      unknownKey: 'ignored',
      securityMode: 'invalid-mode',
      fixedHeightPx: 5000,
      showChrome: 'not-bool',
      refreshIntervalMinutes: '-2',
      inlineContentCacheTtlSeconds: 9999,
    });

    expect(result.appliedKeys).toHaveLength(0);
    expect(result.ignoredEntries).toHaveLength(6);
    expect(propsRecord.unknownKey).toBeUndefined();
    expect(propsRecord.securityMode).toBeUndefined();
    expect(propsRecord.fixedHeightPx).toBeUndefined();
    expect(propsRecord.showChrome).toBeUndefined();
    expect(propsRecord.refreshIntervalMinutes).toBeUndefined();
    expect(propsRecord.inlineContentCacheTtlSeconds).toBeUndefined();
  });
});

describe('buildConfigExport', () => {
  it('uses runtime-compatible defaults for unset timeout fields', () => {
    const exported = buildConfigExport({
      htmlSourceMode: 'FullUrl',
      heightMode: 'Fixed',
      fixedHeightPx: 800,
    });

    expect(exported.iframeLoadTimeoutSeconds).toBe(10);
    expect(exported.inlineContentCacheTtlSeconds).toBe(15);
    expect(exported.enforceStrictInlineCsp).toBe(false);
    expect(exported.contentDeliveryMode).toBe('SharePointFileContent');
    expect(exported.showReportBrowser).toBe(false);
    expect(exported.reportBrowserDefaultView).toBe('Folders');
    expect(exported.reportBrowserMaxItems).toBe(300);
  });
});
