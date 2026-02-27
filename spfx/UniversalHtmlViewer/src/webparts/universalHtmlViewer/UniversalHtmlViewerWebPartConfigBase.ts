import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { getQueryStringParam } from './QueryStringHelper';
import { CacheBusterMode, UrlSecurityMode, UrlValidationOptions } from './UrlHelper';
import { applyImportedConfigToProps } from './ConfigImportExportHelper';
import {
  ConfigurationPreset,
  ITenantConfig,
  IUniversalHtmlViewerWebPartProps,
  TenantConfigMode,
} from './UniversalHtmlViewerTypes';

const blockedTenantConfigKeys = new Set<string>(['__proto__', 'prototype', 'constructor']);
const tenantMergeDefaultValues: Record<string, boolean | number> = {
  fixedHeightPx: 800,
  iframeLoadTimeoutSeconds: 10,
  refreshIntervalMinutes: 0,
  inlineContentCacheTtlSeconds: 15,
  lockPresetSettings: false,
  allowHttp: false,
  enableExpertSecurityModes: false,
  showDiagnostics: false,
  fitContentWidth: false,
  showConfigActions: false,
  showDashboardSelector: false,
  allowQueryStringPageOverride: false,
  showChrome: true,
  showOpenInNewTab: true,
  showRefreshButton: true,
  showStatus: true,
  showLastUpdated: true,
  showLoadingIndicator: true,
};

export abstract class UniversalHtmlViewerWebPartConfigBase extends BaseClientSideWebPart<IUniversalHtmlViewerWebPartProps> {
  protected refreshTimerId: number | undefined;
  protected iframeLoadTimeoutId: number | undefined;
  protected lastEffectiveProps: IUniversalHtmlViewerWebPartProps | undefined;
  protected lastTenantConfig: ITenantConfig | undefined;
  protected lastTenantConfigLoadError: string | undefined;
  protected lastValidationOptions: UrlValidationOptions | undefined;
  protected lastCacheBusterMode: CacheBusterMode | undefined;
  protected lastCacheLabel: string | undefined;
  protected currentBaseUrl: string | undefined;
  protected dashboardOptions: Array<{ id: string; label: string }> = [];

  protected buildUrlValidationOptions(
    currentPageUrl: string,
    effectiveProps: IUniversalHtmlViewerWebPartProps,
  ): UrlValidationOptions {
    const requestedSecurityMode: UrlSecurityMode =
      effectiveProps.securityMode || 'StrictTenant';
    const securityMode: UrlSecurityMode =
      requestedSecurityMode === 'AnyHttps' &&
      effectiveProps.enableExpertSecurityModes !== true
        ? 'StrictTenant'
        : requestedSecurityMode;
    const allowHttp: boolean = !!effectiveProps.allowHttp;
    const allowedHosts: string[] = this.parseHosts(effectiveProps.allowedHosts);
    const allowedPathPrefixes: string[] = this.parsePathPrefixes(
      effectiveProps.allowedPathPrefixes,
    );
    const allowedFileExtensions: string[] = this.parseFileExtensions(
      effectiveProps.allowedFileExtensions,
    );

    return {
      securityMode,
      currentPageUrl,
      allowHttp,
      allowedHosts,
      allowedPathPrefixes,
      allowedFileExtensions,
    };
  }

  protected async getEffectiveProperties(
    pageUrl: string,
  ): Promise<{ effectiveProps: IUniversalHtmlViewerWebPartProps; tenantConfig?: ITenantConfig }> {
    let effectiveProps: IUniversalHtmlViewerWebPartProps = {
      ...this.properties,
    };

    const preset: ConfigurationPreset = effectiveProps.configurationPreset || 'Custom';
    if (effectiveProps.lockPresetSettings && preset !== 'Custom') {
      this.applyPreset(preset, effectiveProps);
    }

    const tenantConfig: ITenantConfig | undefined = await this.tryLoadTenantConfig(
      pageUrl,
      effectiveProps,
    );

    if (tenantConfig) {
      const mode: TenantConfigMode = effectiveProps.tenantConfigMode || 'Merge';
      effectiveProps = this.mergeTenantConfig(effectiveProps, tenantConfig, mode);
    }

    return { effectiveProps, tenantConfig };
  }

  private mergeTenantConfig(
    props: IUniversalHtmlViewerWebPartProps,
    tenantConfig: ITenantConfig,
    mode: TenantConfigMode,
  ): IUniversalHtmlViewerWebPartProps {
    const nextProps: IUniversalHtmlViewerWebPartProps = { ...props };
    const applyIfEmpty = mode === 'Merge';

    const shouldApply = (key: string, currentValue: unknown): boolean => {
      if (!applyIfEmpty) {
        return true;
      }
      if (currentValue === undefined || currentValue === null) {
        return true;
      }
      if (typeof currentValue === 'string') {
        return currentValue.trim().length === 0;
      }
      if (Object.prototype.hasOwnProperty.call(tenantMergeDefaultValues, key)) {
        return currentValue === tenantMergeDefaultValues[key];
      }
      return false;
    };

    const nextPropsRecord = nextProps as unknown as Record<string, unknown>;
    const normalizedTenantConfig = this.sanitizeTenantConfigForMerge(tenantConfig);

    Object.entries(normalizedTenantConfig).forEach(([key, value]) => {
      if (value === undefined) {
        return;
      }
      const currentValue = nextPropsRecord[key];
      if (shouldApply(key, currentValue)) {
        nextPropsRecord[key] = value;
      }
    });

    const preset: ConfigurationPreset = nextProps.configurationPreset || 'Custom';
    if (nextProps.lockPresetSettings && preset !== 'Custom') {
      this.applyPreset(preset, nextProps);
    }

    return nextProps;
  }

  private sanitizeTenantConfigForMerge(tenantConfig: ITenantConfig): Record<string, unknown> {
    const normalizedInput: Record<string, unknown> = {};
    Object.entries(tenantConfig as Record<string, unknown>).forEach(([key, value]) => {
      if (value === undefined || blockedTenantConfigKeys.has(key)) {
        return;
      }
      if (key === 'dashboardList' && Array.isArray(value)) {
        normalizedInput[key] = value.join(',');
        return;
      }
      normalizedInput[key] = value;
    });

    const sanitizedConfig: Record<string, unknown> = Object.create(null) as Record<string, unknown>;
    applyImportedConfigToProps(sanitizedConfig, normalizedInput);
    return sanitizedConfig;
  }

  private async tryLoadTenantConfig(
    pageUrl: string,
    props: IUniversalHtmlViewerWebPartProps,
  ): Promise<ITenantConfig | undefined> {
    const rawUrl: string = (props.tenantConfigUrl || '').trim();
    if (!rawUrl) {
      this.lastTenantConfigLoadError = undefined;
      return undefined;
    }

    const resolvedUrl: string | undefined = this.resolveTenantConfigUrl(rawUrl, pageUrl);
    if (!resolvedUrl) {
      this.lastTenantConfigLoadError = 'Tenant config URL is invalid or not allowed.';
      return undefined;
    }

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        resolvedUrl,
        SPHttpClient.configurations.v1,
      );
      if (!response.ok) {
        this.lastTenantConfigLoadError = `Tenant config request failed (${response.status}).`;
        return undefined;
      }

      const data = await response.json();
      if (!data || typeof data !== 'object') {
        this.lastTenantConfigLoadError =
          'Tenant config payload is invalid. Expected a JSON object.';
        return undefined;
      }

      this.lastTenantConfigLoadError = undefined;
      return data as ITenantConfig;
    } catch {
      this.lastTenantConfigLoadError = 'Tenant config request failed due to a network or parsing error.';
      return undefined;
    }
  }

  private resolveTenantConfigUrl(rawUrl: string, pageUrl: string): string | undefined {
    const trimmed: string = rawUrl.trim();
    if (!trimmed) {
      return undefined;
    }

    if (trimmed.includes('\\')) {
      return undefined;
    }

    const trimmedLower: string = trimmed.toLowerCase();

    const currentUrl: URL = new URL(pageUrl || this.context.pageContext.web.absoluteUrl);
    const origin: string = currentUrl.origin;

    if (trimmedLower.startsWith('http://')) {
      return undefined;
    }

    if (trimmedLower.startsWith('https://')) {
      try {
        const target: URL = new URL(trimmed);
        if (target.hostname.toLowerCase() !== currentUrl.hostname.toLowerCase()) {
          return undefined;
        }
        return target.toString();
      } catch {
        return undefined;
      }
    }

    if (trimmedLower.includes('://')) {
      return undefined;
    }

    if (trimmed.startsWith('/')) {
      return `${origin}${trimmed}`;
    }

    const webServerRelative: string = this.context.pageContext.web.serverRelativeUrl || '/';
    const normalizedBase: string = webServerRelative.endsWith('/')
      ? webServerRelative.slice(0, -1)
      : webServerRelative;
    return `${origin}${normalizedBase}/${trimmed}`;
  }

  protected getEffectiveDashboardId(
    props: IUniversalHtmlViewerWebPartProps,
    pageUrl: string,
  ): string | undefined {
    if (props.htmlSourceMode !== 'BasePathAndDashboardId') {
      return undefined;
    }

    const queryParamName: string = (props.queryStringParamName || '').trim() || 'dashboard';
    const dashboardFromQuery: string | undefined = pageUrl
      ? getQueryStringParam(pageUrl, queryParamName)
      : undefined;
    const effectiveDashboardId: string = (dashboardFromQuery || props.dashboardId || '').trim();
    return effectiveDashboardId || undefined;
  }

  protected applyPreset(
    presetValue: ConfigurationPreset,
    targetProps?: IUniversalHtmlViewerWebPartProps,
  ): void {
    const preset: ConfigurationPreset = (presetValue || 'Custom') as ConfigurationPreset;
    const props: IUniversalHtmlViewerWebPartProps = targetProps || this.properties;
    props.configurationPreset = preset;

    if (preset === 'Custom') {
      return;
    }

    const basePathPrefix: string = this.normalizeBasePathForPrefix(props.basePath);

    props.allowHttp = false;
    props.enableExpertSecurityModes = false;
    props.allowedFileExtensions = '.html,.htm,.aspx';
    props.showChrome = true;
    props.showOpenInNewTab = true;
    props.showRefreshButton = true;
    props.showStatus = true;
    props.showLastUpdated = true;
    props.showLoadingIndicator = true;
    props.showConfigActions = true;
    props.showDashboardSelector = false;
    props.allowQueryStringPageOverride = false;
    props.fitContentWidth = false;
    props.chromeDensity = 'Comfortable';
    props.iframeLoadTimeoutSeconds = 10;
    props.inlineContentCacheTtlSeconds = 15;

    if (!props.chromeTitle || props.chromeTitle.trim().length === 0) {
      props.chromeTitle = 'Universal HTML Viewer';
    }

    if (basePathPrefix) {
      props.allowedPathPrefixes = basePathPrefix;
    }

    switch (preset) {
      case 'SharePointLibraryStrict':
        props.contentDeliveryMode = 'SharePointFileContent';
        props.securityMode = 'StrictTenant';
        props.cacheBusterMode = 'FileLastModified';
        props.sandboxPreset = 'Strict';
        props.fitContentWidth = true;
        break;
      case 'SharePointLibraryRelaxed':
        props.contentDeliveryMode = 'SharePointFileContent';
        props.securityMode = 'StrictTenant';
        props.cacheBusterMode = 'FileLastModified';
        props.sandboxPreset = 'Relaxed';
        props.fitContentWidth = true;
        break;
      case 'SharePointLibraryFullPage':
        props.contentDeliveryMode = 'SharePointFileContent';
        props.securityMode = 'StrictTenant';
        props.cacheBusterMode = 'FileLastModified';
        props.sandboxPreset = 'Relaxed';
        props.heightMode = 'Viewport';
        props.showChrome = false;
        props.fitContentWidth = true;
        break;
      case 'AllowlistCDN':
        props.contentDeliveryMode = 'DirectUrl';
        props.securityMode = 'Allowlist';
        props.cacheBusterMode = 'Timestamp';
        props.sandboxPreset = 'Relaxed';
        break;
      case 'AnyHttps':
        props.contentDeliveryMode = 'DirectUrl';
        props.securityMode = 'AnyHttps';
        props.enableExpertSecurityModes = true;
        props.cacheBusterMode = 'Timestamp';
        props.sandboxPreset = 'None';
        break;
      default:
        break;
    }
  }

  protected normalizeBasePathForPrefix(value?: string): string {
    const trimmed: string = (value || '').trim();
    if (!trimmed) {
      return '';
    }

    let normalized: string = trimmed;
    if (!normalized.startsWith('/')) {
      normalized = `/${normalized}`;
    }
    if (!normalized.endsWith('/')) {
      normalized = `${normalized}/`;
    }
    return normalized.toLowerCase();
  }

  protected parseHosts(value?: string): string[] {
    return (value || '')
      .split(/[,;\s]+/g)
      .map((entry) => entry.trim())
      .filter((entry) => entry.length > 0)
      .map((entry) => {
        let hostValue: string = entry;
        try {
          if (entry.includes('://')) {
            hostValue = new URL(entry).hostname;
          } else {
            hostValue = entry.split('/')[0];
          }
        } catch {
          hostValue = entry;
        }

        if (hostValue.startsWith('*.')) {
          hostValue = hostValue.substring(1);
        }

        const withoutPort: string = hostValue.split(':')[0];
        return withoutPort.toLowerCase();
      });
  }

  protected parsePathPrefixes(value?: string): string[] {
    return (value || '')
      .split(/[,;\s]+/g)
      .map((entry) => entry.trim())
      .filter((entry) => entry.length > 0)
      .map((entry) => {
        if (!entry.startsWith('/')) {
          return `/${entry}`;
        }
        return entry;
      });
  }

  protected parseFileExtensions(value?: string): string[] {
    return (value || '')
      .split(/[,;\s]+/g)
      .map((entry) => entry.trim())
      .filter((entry) => entry.length > 0)
      .map((entry) => {
        if (entry.startsWith('.')) {
          return entry.toLowerCase();
        }
        return `.${entry.toLowerCase()}`;
      });
  }

  protected normalizeCacheBusterParamName(value?: string): string {
    const trimmed: string = (value || '').trim();
    if (!trimmed) {
      return 'v';
    }
    if (!/^[a-zA-Z0-9_-]+$/.test(trimmed)) {
      return 'v';
    }
    return trimmed;
  }

  protected normalizeIframeLoading(value?: string): string {
    const normalized: string = (value || '').trim().toLowerCase();
    if (normalized === 'lazy' || normalized === 'eager') {
      return normalized;
    }
    return '';
  }

  protected normalizeIframeSandbox(value?: string, preset?: string): string {
    const normalizedPreset: string = (preset || '').trim().toLowerCase();
    if (normalizedPreset && normalizedPreset !== 'custom') {
      if (normalizedPreset === 'relaxed') {
        return 'allow-same-origin allow-scripts allow-forms allow-popups allow-downloads';
      }
      if (normalizedPreset === 'strict') {
        return 'allow-scripts';
      }
      return '';
    }

    const tokens: string[] = (value || '')
      .split(/\s+/g)
      .map((token) => token.trim())
      .filter((token) => token.length > 0);

    if (tokens.length === 0) {
      return '';
    }

    const allowedTokens = new Set<string>([
      'allow-downloads',
      'allow-downloads-without-user-activation',
      'allow-forms',
      'allow-modals',
      'allow-orientation-lock',
      'allow-pointer-lock',
      'allow-popups',
      'allow-popups-to-escape-sandbox',
      'allow-presentation',
      'allow-same-origin',
      'allow-scripts',
      'allow-storage-access-by-user-activation',
      'allow-top-navigation',
      'allow-top-navigation-by-user-activation',
      'allow-top-navigation-to-custom-protocols',
    ]);

    const sanitized = tokens.filter((token) => allowedTokens.has(token));
    return sanitized.join(' ');
  }

  protected normalizeIframeAllow(value?: string): string {
    const trimmed: string = (value || '').trim();
    if (!trimmed) {
      return '';
    }
    return trimmed.replace(/[^a-zA-Z0-9;=(),\s\-:'"*]/g, '');
  }

  protected normalizeReferrerPolicy(value?: string): string {
    const normalized: string = (value || '').trim().toLowerCase();
    const allowed = new Set<string>([
      'no-referrer',
      'no-referrer-when-downgrade',
      'origin',
      'origin-when-cross-origin',
      'same-origin',
      'strict-origin',
      'strict-origin-when-cross-origin',
      'unsafe-url',
    ]);
    if (allowed.has(normalized)) {
      return normalized;
    }
    return '';
  }

  protected getCurrentPageUrl(): string {
    if (typeof window !== 'undefined' && window.location && window.location.href) {
      return window.location.href;
    }

    try {
      return this.context.pageContext.web.absoluteUrl;
    } catch {
      return '';
    }
  }
}
