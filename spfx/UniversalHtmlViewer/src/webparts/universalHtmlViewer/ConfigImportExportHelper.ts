import { IUniversalHtmlViewerWebPartProps } from './UniversalHtmlViewerTypes';

const enumValues: Record<string, string[]> = {
  configurationPreset: [
    'Custom',
    'SharePointLibraryRelaxed',
    'SharePointLibraryFullPage',
    'SharePointLibraryStrict',
    'AllowlistCDN',
    'AnyHttps',
  ],
  contentDeliveryMode: ['DirectUrl', 'SharePointFileContent'],
  htmlSourceMode: ['FullUrl', 'BasePathAndRelativePath', 'BasePathAndDashboardId'],
  heightMode: ['Fixed', 'Viewport', 'Auto'],
  securityMode: ['StrictTenant', 'Allowlist', 'AnyHttps'],
  tenantConfigMode: ['Merge', 'Override'],
  cacheBusterMode: ['None', 'Timestamp', 'FileLastModified'],
  sandboxPreset: ['None', 'Relaxed', 'Strict', 'Custom'],
  iframeLoading: ['', 'lazy', 'eager'],
  iframeReferrerPolicy: [
    '',
    'no-referrer',
    'no-referrer-when-downgrade',
    'origin',
    'origin-when-cross-origin',
    'same-origin',
    'strict-origin',
    'strict-origin-when-cross-origin',
    'unsafe-url',
  ],
  chromeDensity: ['Comfortable', 'Compact'],
};

const numberRanges: Record<string, { min: number; max: number }> = {
  fixedHeightPx: { min: 200, max: 2000 },
  iframeLoadTimeoutSeconds: { min: 0, max: 60 },
  refreshIntervalMinutes: { min: 0, max: 120 },
  inlineContentCacheTtlSeconds: { min: 0, max: 300 },
};

const booleanKeys = new Set<string>([
  'lockPresetSettings',
  'allowHttp',
  'enableExpertSecurityModes',
  'showDiagnostics',
  'showChrome',
  'fitContentWidth',
  'showOpenInNewTab',
  'showRefreshButton',
  'showStatus',
  'showLastUpdated',
  'showLoadingIndicator',
  'showConfigActions',
  'showDashboardSelector',
  'allowQueryStringPageOverride',
]);

const numberKeys = new Set<string>([
  'fixedHeightPx',
  'iframeLoadTimeoutSeconds',
  'refreshIntervalMinutes',
  'inlineContentCacheTtlSeconds',
]);

const stringKeys = new Set<string>([
  'configurationPreset',
  'contentDeliveryMode',
  'htmlSourceMode',
  'fullUrl',
  'basePath',
  'relativePath',
  'dashboardId',
  'dashboardList',
  'defaultFileName',
  'queryStringParamName',
  'heightMode',
  'securityMode',
  'allowedHosts',
  'allowedPathPrefixes',
  'allowedFileExtensions',
  'tenantConfigUrl',
  'tenantConfigMode',
  'cacheBusterMode',
  'cacheBusterParamName',
  'sandboxPreset',
  'iframeSandbox',
  'iframeAllow',
  'iframeReferrerPolicy',
  'iframeLoading',
  'iframeTitle',
  'chromeTitle',
  'chromeSubtitle',
  'chromeDensity',
]);

const allSupportedKeys = new Set<string>([
  ...Array.from(booleanKeys),
  ...Array.from(numberKeys),
  ...Array.from(stringKeys),
]);

const enumValueLookup: Record<string, Map<string, string>> = {};
Object.keys(enumValues).forEach((key) => {
  enumValueLookup[key] = new Map(
    enumValues[key].map((value) => [value.toLowerCase(), value]),
  );
});

export interface IConfigImportIssue {
  key: string;
  reason: string;
  value: unknown;
}

export interface IConfigImportResult {
  appliedKeys: string[];
  ignoredEntries: IConfigImportIssue[];
}

export function applyImportedConfigToProps(
  propsRecord: Record<string, unknown>,
  config: Record<string, unknown>,
): IConfigImportResult {
  const result: IConfigImportResult = {
    appliedKeys: [],
    ignoredEntries: [],
  };

  Object.entries(config).forEach(([key, value]) => {
    if (value === undefined || value === null) {
      return;
    }

    if (!allSupportedKeys.has(key)) {
      result.ignoredEntries.push({
        key,
        value,
        reason: 'Unsupported key.',
      });
      return;
    }

    if (booleanKeys.has(key)) {
      const parsedBoolean = parseBoolean(value);
      if (parsedBoolean === undefined) {
        result.ignoredEntries.push({
          key,
          value,
          reason: 'Expected a boolean value.',
        });
        return;
      }
      propsRecord[key] = parsedBoolean;
      result.appliedKeys.push(key);
      return;
    }

    if (numberKeys.has(key)) {
      const parsed = typeof value === 'number' ? value : Number(value);
      if (!Number.isFinite(parsed)) {
        result.ignoredEntries.push({
          key,
          value,
          reason: 'Expected a numeric value.',
        });
        return;
      }

      const range = numberRanges[key];
      if (range && (parsed < range.min || parsed > range.max)) {
        result.ignoredEntries.push({
          key,
          value,
          reason: `Value must be in range ${range.min}-${range.max}.`,
        });
        return;
      }

      propsRecord[key] = parsed;
      result.appliedKeys.push(key);
      return;
    }

    if (stringKeys.has(key)) {
      const enumLookup = enumValueLookup[key];
      if (enumLookup) {
        const normalizedInput = String(value).trim().toLowerCase();
        const canonical = enumLookup.get(normalizedInput);
        if (!canonical) {
          result.ignoredEntries.push({
            key,
            value,
            reason: 'Value is not in the allowed set.',
          });
          return;
        }
        propsRecord[key] = canonical;
        result.appliedKeys.push(key);
        return;
      }

      propsRecord[key] = String(value);
      result.appliedKeys.push(key);
    }
  });

  return result;
}

export function buildConfigExport(
  props: IUniversalHtmlViewerWebPartProps,
): Record<string, unknown> {
  return {
    configurationPreset: props.configurationPreset || 'Custom',
    lockPresetSettings: !!props.lockPresetSettings,
    contentDeliveryMode: props.contentDeliveryMode || 'DirectUrl',
    htmlSourceMode: props.htmlSourceMode,
    fullUrl: props.fullUrl || '',
    basePath: props.basePath || '',
    relativePath: props.relativePath || '',
    dashboardId: props.dashboardId || '',
    dashboardList: props.dashboardList || '',
    defaultFileName: props.defaultFileName || '',
    queryStringParamName: props.queryStringParamName || '',
    heightMode: props.heightMode,
    fixedHeightPx: props.fixedHeightPx,
    fitContentWidth: props.fitContentWidth === true,
    securityMode: props.securityMode || 'StrictTenant',
    enableExpertSecurityModes: props.enableExpertSecurityModes === true,
    allowHttp: !!props.allowHttp,
    allowedHosts: props.allowedHosts || '',
    allowedPathPrefixes: props.allowedPathPrefixes || '',
    allowedFileExtensions: props.allowedFileExtensions || '',
    tenantConfigUrl: props.tenantConfigUrl || '',
    tenantConfigMode: props.tenantConfigMode || 'Merge',
    cacheBusterMode: props.cacheBusterMode || 'None',
    cacheBusterParamName: props.cacheBusterParamName || 'v',
    inlineContentCacheTtlSeconds: props.inlineContentCacheTtlSeconds ?? 15,
    sandboxPreset: props.sandboxPreset || 'None',
    iframeSandbox: props.iframeSandbox || '',
    iframeAllow: props.iframeAllow || '',
    iframeReferrerPolicy: props.iframeReferrerPolicy || '',
    iframeLoading: props.iframeLoading || '',
    iframeTitle: props.iframeTitle || '',
    iframeLoadTimeoutSeconds: props.iframeLoadTimeoutSeconds ?? 0,
    refreshIntervalMinutes: props.refreshIntervalMinutes ?? 0,
    showDiagnostics: !!props.showDiagnostics,
    showChrome: props.showChrome !== false,
    chromeTitle: props.chromeTitle || '',
    chromeSubtitle: props.chromeSubtitle || '',
    showOpenInNewTab: props.showOpenInNewTab !== false,
    showRefreshButton: props.showRefreshButton !== false,
    showStatus: props.showStatus !== false,
    showLastUpdated: props.showLastUpdated !== false,
    chromeDensity: props.chromeDensity || 'Comfortable',
    showLoadingIndicator: props.showLoadingIndicator !== false,
    showConfigActions: props.showConfigActions === true,
    showDashboardSelector: props.showDashboardSelector === true,
    allowQueryStringPageOverride: props.allowQueryStringPageOverride === true,
  };
}

function parseBoolean(value: unknown): boolean | undefined {
  if (typeof value === 'boolean') {
    return value;
  }

  if (typeof value === 'number') {
    if (value === 1) {
      return true;
    }
    if (value === 0) {
      return false;
    }
    return undefined;
  }

  if (typeof value === 'string') {
    const normalized = value.trim().toLowerCase();
    if (normalized === 'true' || normalized === '1' || normalized === 'yes') {
      return true;
    }
    if (normalized === 'false' || normalized === '0' || normalized === 'no') {
      return false;
    }
  }

  return undefined;
}
