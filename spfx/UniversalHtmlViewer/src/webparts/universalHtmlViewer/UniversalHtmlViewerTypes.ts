import { CacheBusterMode, HeightMode, HtmlSourceMode, UrlSecurityMode } from './UrlHelper';

export type ConfigurationPreset =
  | 'Custom'
  | 'SharePointLibraryRelaxed'
  | 'SharePointLibraryFullPage'
  | 'SharePointLibraryStrict'
  | 'AllowlistCDN'
  | 'AnyHttps';

export type TenantConfigMode = 'Merge' | 'Override';

export type ChromeDensity = 'Comfortable' | 'Compact';
export type ContentDeliveryMode = 'DirectUrl' | 'SharePointFileContent';

export interface IUniversalHtmlViewerWebPartProps {
  configurationPreset?: ConfigurationPreset;
  lockPresetSettings?: boolean;
  contentDeliveryMode?: ContentDeliveryMode;
  htmlSourceMode: HtmlSourceMode;
  fullUrl?: string;
  basePath?: string;
  relativePath?: string;
  dashboardId?: string;
  dashboardList?: string | string[];
  defaultFileName?: string;
  queryStringParamName?: string;
  heightMode: HeightMode;
  fixedHeightPx: number;
  securityMode?: UrlSecurityMode;
  allowHttp?: boolean;
  allowedHosts?: string;
  allowedPathPrefixes?: string;
  allowedFileExtensions?: string;
  tenantConfigUrl?: string;
  tenantConfigMode?: TenantConfigMode;
  cacheBusterMode?: CacheBusterMode;
  cacheBusterParamName?: string;
  sandboxPreset?: string;
  iframeSandbox?: string;
  iframeAllow?: string;
  iframeReferrerPolicy?: string;
  iframeLoading?: string;
  iframeTitle?: string;
  iframeLoadTimeoutSeconds?: number;
  refreshIntervalMinutes?: number;
  showDiagnostics?: boolean;
  showChrome?: boolean;
  chromeTitle?: string;
  chromeSubtitle?: string;
  showOpenInNewTab?: boolean;
  showRefreshButton?: boolean;
  showStatus?: boolean;
  showLastUpdated?: boolean;
  chromeDensity?: ChromeDensity;
  showLoadingIndicator?: boolean;
  showConfigActions?: boolean;
  showDashboardSelector?: boolean;
}

export interface ITenantConfig {
  configurationPreset?: ConfigurationPreset;
  lockPresetSettings?: boolean;
  contentDeliveryMode?: ContentDeliveryMode;
  htmlSourceMode?: HtmlSourceMode;
  fullUrl?: string;
  basePath?: string;
  relativePath?: string;
  dashboardId?: string;
  dashboardList?: string;
  defaultFileName?: string;
  queryStringParamName?: string;
  heightMode?: HeightMode;
  fixedHeightPx?: number;
  securityMode?: UrlSecurityMode;
  allowHttp?: boolean;
  allowedHosts?: string;
  allowedPathPrefixes?: string;
  allowedFileExtensions?: string;
  tenantConfigUrl?: string;
  tenantConfigMode?: TenantConfigMode;
  cacheBusterMode?: CacheBusterMode;
  cacheBusterParamName?: string;
  sandboxPreset?: string;
  iframeSandbox?: string;
  iframeAllow?: string;
  iframeReferrerPolicy?: string;
  iframeLoading?: string;
  iframeTitle?: string;
  iframeLoadTimeoutSeconds?: number;
  refreshIntervalMinutes?: number;
  showDiagnostics?: boolean;
  showChrome?: boolean;
  chromeTitle?: string;
  chromeSubtitle?: string;
  showOpenInNewTab?: boolean;
  showRefreshButton?: boolean;
  showStatus?: boolean;
  showLastUpdated?: boolean;
  chromeDensity?: ChromeDensity;
  showLoadingIndicator?: boolean;
  showConfigActions?: boolean;
  showDashboardSelector?: boolean;
}
