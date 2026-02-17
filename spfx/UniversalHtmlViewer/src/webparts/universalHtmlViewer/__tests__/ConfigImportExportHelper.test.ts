import { applyImportedConfigToProps } from '../ConfigImportExportHelper';

describe('applyImportedConfigToProps', () => {
  it('applies valid values and normalizes enum casing', () => {
    const propsRecord: Record<string, unknown> = {};
    const result = applyImportedConfigToProps(propsRecord, {
      configurationPreset: 'sharepointlibraryfullpage',
      contentDeliveryMode: 'sharepointfilecontent',
      htmlSourceMode: 'basepathandrelativepath',
      fixedHeightPx: '700',
      refreshIntervalMinutes: 10,
      showChrome: 'true',
      iframeLoading: 'LAZY',
      chromeDensity: 'compact',
    });

    expect(result.ignoredEntries).toHaveLength(0);
    expect(propsRecord.configurationPreset).toBe('SharePointLibraryFullPage');
    expect(propsRecord.contentDeliveryMode).toBe('SharePointFileContent');
    expect(propsRecord.htmlSourceMode).toBe('BasePathAndRelativePath');
    expect(propsRecord.fixedHeightPx).toBe(700);
    expect(propsRecord.refreshIntervalMinutes).toBe(10);
    expect(propsRecord.showChrome).toBe(true);
    expect(propsRecord.iframeLoading).toBe('lazy');
    expect(propsRecord.chromeDensity).toBe('Compact');
  });

  it('ignores unknown keys and invalid values', () => {
    const propsRecord: Record<string, unknown> = {};
    const result = applyImportedConfigToProps(propsRecord, {
      unknownKey: 'ignored',
      securityMode: 'invalid-mode',
      fixedHeightPx: 5000,
      showChrome: 'not-bool',
      refreshIntervalMinutes: '-2',
    });

    expect(result.appliedKeys).toHaveLength(0);
    expect(result.ignoredEntries).toHaveLength(5);
    expect(propsRecord.unknownKey).toBeUndefined();
    expect(propsRecord.securityMode).toBeUndefined();
    expect(propsRecord.fixedHeightPx).toBeUndefined();
    expect(propsRecord.showChrome).toBeUndefined();
    expect(propsRecord.refreshIntervalMinutes).toBeUndefined();
  });
});
