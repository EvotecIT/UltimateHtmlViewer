import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneToggle,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import styles from './UniversalHtmlViewerWebPart.module.scss';
import { getQueryStringParam } from './QueryStringHelper';
import {
  buildFinalUrl,
  isUrlAllowed,
  HtmlSourceMode,
  HeightMode,
  UrlSecurityMode,
  CacheBusterMode,
  UrlValidationOptions,
} from './UrlHelper';

type ConfigurationPreset =
  | 'Custom'
  | 'SharePointLibraryRelaxed'
  | 'SharePointLibraryStrict'
  | 'AllowlistCDN'
  | 'AnyHttps';

type TenantConfigMode = 'Merge' | 'Override';

type ChromeDensity = 'Comfortable' | 'Compact';

export interface IUniversalHtmlViewerWebPartProps {
  configurationPreset?: ConfigurationPreset;
  lockPresetSettings?: boolean;
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

interface ITenantConfig {
  configurationPreset?: ConfigurationPreset;
  lockPresetSettings?: boolean;
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

export default class UniversalHtmlViewerWebPart extends BaseClientSideWebPart<IUniversalHtmlViewerWebPartProps> {
  private refreshTimerId: number | undefined;
  private iframeLoadTimeoutId: number | undefined;
  private lastEffectiveProps: IUniversalHtmlViewerWebPartProps | undefined;
  private lastTenantConfig: ITenantConfig | undefined;
  private lastValidationOptions: UrlValidationOptions | undefined;
  private lastCacheBusterMode: CacheBusterMode | undefined;
  private lastCacheLabel: string | undefined;
  private currentBaseUrl: string | undefined;
  private dashboardOptions: Array<{ id: string; label: string }> = [];

  public render(): void {
    this.renderAsync().catch((error) => {
      this.clearRefreshTimer();
      this.clearIframeLoadTimeout();
      this.domElement.innerHTML = this.buildMessageHtml(
        'UniversalHtmlViewer: Failed to render content.',
        this.buildDiagnosticsHtml({
          error: String(error),
        }),
      );
      // eslint-disable-next-line no-console
      console.error('UniversalHtmlViewer render failed', error);
    });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: unknown,
    newValue: unknown,
  ): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    const currentPreset: ConfigurationPreset =
      this.properties.configurationPreset || 'Custom';
    const isPresetLocked: boolean =
      !!this.properties.lockPresetSettings && currentPreset !== 'Custom';

    if (
      (propertyPath === 'configurationPreset' || propertyPath === 'lockPresetSettings') &&
      newValue !== oldValue
    ) {
      this.applyPreset(currentPreset);
      this.context.propertyPane.refresh();
      this.render();
      return;
    }

    if (propertyPath === 'basePath' && isPresetLocked) {
      this.applyPreset(currentPreset);
      this.context.propertyPane.refresh();
      this.render();
      return;
    }

    if (
      propertyPath === 'htmlSourceMode' ||
      propertyPath === 'securityMode' ||
      propertyPath === 'sandboxPreset' ||
      propertyPath === 'cacheBusterMode' ||
      propertyPath === 'showChrome' ||
      propertyPath === 'configurationPreset' ||
      propertyPath === 'lockPresetSettings' ||
      propertyPath === 'tenantConfigMode' ||
      propertyPath === 'showDashboardSelector'
    ) {
      this.context.propertyPane.refresh();
    }

    this.render();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const htmlSourceMode: HtmlSourceMode = this.properties.htmlSourceMode || 'FullUrl';
    const isFullUrl: boolean = htmlSourceMode === 'FullUrl';
    const isRelativePath: boolean = htmlSourceMode === 'BasePathAndRelativePath';
    const isDashboardId: boolean = htmlSourceMode === 'BasePathAndDashboardId';
    const securityMode: UrlSecurityMode = this.properties.securityMode || 'StrictTenant';
    const isAllowlistMode: boolean = securityMode === 'Allowlist';
    const cacheBusterMode: CacheBusterMode = this.properties.cacheBusterMode || 'None';
    const sandboxPreset: string = this.properties.sandboxPreset || 'None';
    const isCustomSandbox: boolean = sandboxPreset === 'Custom';
    const showChrome: boolean = this.properties.showChrome !== false;
    const preset: ConfigurationPreset = this.properties.configurationPreset || 'Custom';
    const isPresetLocked: boolean =
      !!this.properties.lockPresetSettings && preset !== 'Custom';
    const showDashboardSelector: boolean = this.properties.showDashboardSelector === true;

    return {
      pages: [
        {
          header: {
            description: 'Configure the HTML source and layout.',
          },
          groups: [
            {
              groupName: 'Presets & UX',
              groupFields: [
                PropertyPaneDropdown('configurationPreset', {
                  label: 'Configuration preset',
                  options: [
                    { key: 'Custom', text: 'Custom (manual settings)' },
                    { key: 'SharePointLibraryRelaxed', text: 'SharePoint library (relaxed)' },
                    { key: 'SharePointLibraryStrict', text: 'SharePoint library (strict)' },
                    { key: 'AllowlistCDN', text: 'Allowlist CDN' },
                    { key: 'AnyHttps', text: 'Any HTTPS (unsafe)' },
                  ],
                }),
                PropertyPaneToggle('lockPresetSettings', {
                  label: 'Lock preset settings',
                  onText: 'Locked',
                  offText: 'Editable',
                  disabled: preset === 'Custom',
                }),
                PropertyPaneToggle('showChrome', {
                  label: 'Show header chrome',
                  onText: 'On',
                  offText: 'Off',
                  disabled: isPresetLocked,
                }),
                PropertyPaneDropdown('chromeDensity', {
                  label: 'Chrome density',
                  options: [
                    { key: 'Comfortable', text: 'Comfortable' },
                    { key: 'Compact', text: 'Compact' },
                  ],
                  disabled: !showChrome || isPresetLocked,
                }),
                PropertyPaneTextField('chromeTitle', {
                  label: 'Chrome title',
                  description: 'Shown above the iframe.',
                  disabled: !showChrome,
                }),
                PropertyPaneTextField('chromeSubtitle', {
                  label: 'Chrome subtitle',
                  description: 'Optional helper text under the title.',
                  disabled: !showChrome,
                }),
                PropertyPaneToggle('showRefreshButton', {
                  label: 'Show refresh button',
                  onText: 'On',
                  offText: 'Off',
                  disabled: !showChrome || isPresetLocked,
                }),
                PropertyPaneToggle('showOpenInNewTab', {
                  label: 'Show "Open in new tab"',
                  onText: 'On',
                  offText: 'Off',
                  disabled: !showChrome || isPresetLocked,
                }),
                PropertyPaneToggle('showStatus', {
                  label: 'Show status pill',
                  onText: 'On',
                  offText: 'Off',
                  disabled: !showChrome || isPresetLocked,
                }),
                PropertyPaneToggle('showLastUpdated', {
                  label: 'Show last updated time',
                  onText: 'On',
                  offText: 'Off',
                  disabled: !showChrome || isPresetLocked,
                }),
                PropertyPaneToggle('showLoadingIndicator', {
                  label: 'Show loading indicator',
                  onText: 'On',
                  offText: 'Off',
                  disabled: isPresetLocked,
                }),
                PropertyPaneToggle('showConfigActions', {
                  label: 'Show config export/import',
                  onText: 'On',
                  offText: 'Off',
                  disabled: !showChrome,
                }),
                PropertyPaneToggle('showDashboardSelector', {
                  label: 'Show dashboard selector',
                  onText: 'On',
                  offText: 'Off',
                }),
                PropertyPaneTextField('dashboardList', {
                  label: 'Dashboard list (comma-separated)',
                  description: 'Optional list, e.g. Sales|sales, Ops|ops',
                  disabled: !showDashboardSelector,
                }),
              ],
            },
            {
              groupName: 'Source settings',
              groupFields: [
                PropertyPaneDropdown('htmlSourceMode', {
                  label: 'HTML source mode',
                  options: [
                    { key: 'FullUrl', text: 'Full URL' },
                    { key: 'BasePathAndRelativePath', text: 'Base path + relative path' },
                    { key: 'BasePathAndDashboardId', text: 'Base path + dashboard ID' },
                  ],
                }),
                PropertyPaneTextField('fullUrl', {
                  label: 'Full URL to HTML page',
                  description: 'Used when HTML source mode is "FullUrl".',
                  disabled: !isFullUrl,
                  onGetErrorMessage: this.validateFullUrl.bind(this),
                  deferredValidationTime: 200,
                }),
                PropertyPaneTextField('basePath', {
                  label: 'Base path (site-relative)',
                  description:
                    'Site-relative base path, used when HTML source mode is not "FullUrl". Example: /sites/Reports/Dashboards/',
                  disabled: isFullUrl,
                  onGetErrorMessage: this.validateBasePath.bind(this),
                  deferredValidationTime: 200,
                }),
                PropertyPaneTextField('relativePath', {
                  label: 'Relative path from base',
                  description:
                    'Used when HTML source mode is "BasePathAndRelativePath". Example: system1/index.html',
                  disabled: !isRelativePath,
                }),
                PropertyPaneTextField('dashboardId', {
                  label: 'Dashboard ID (fallback)',
                  description:
                    'Used when HTML source mode is "BasePathAndDashboardId" and no query string parameter is provided.',
                  disabled: !isDashboardId,
                }),
                PropertyPaneTextField('defaultFileName', {
                  label: 'Default file name',
                  description:
                    'Used when HTML source mode is "BasePathAndDashboardId". Defaults to "index.html" when left empty.',
                  disabled: !isDashboardId,
                }),
                PropertyPaneTextField('queryStringParamName', {
                  label: 'Query string parameter name',
                  description:
                    'Used when HTML source mode is "BasePathAndDashboardId" to read the dashboard ID from the page URL. Defaults to "dashboard" when left empty.',
                  disabled: !isDashboardId,
                }),
              ],
            },
            {
              groupName: 'Security',
              groupFields: [
                PropertyPaneDropdown('securityMode', {
                  label: 'Security mode',
                  options: [
                    { key: 'StrictTenant', text: 'Strict tenant (default)' },
                    { key: 'Allowlist', text: 'Tenant + allowlist' },
                    { key: 'AnyHttps', text: 'Any HTTPS (unsafe)' },
                  ],
                  disabled: isPresetLocked,
                }),
                PropertyPaneToggle('allowHttp', {
                  label: 'Allow HTTP (unsafe)',
                  onText: 'Allow',
                  offText: 'Block',
                  disabled: isPresetLocked,
                }),
                PropertyPaneTextField('allowedHosts', {
                  label: 'Allowed hosts (comma-separated)',
                  description:
                    'Used when security mode is "Allowlist". Example: cdn.contoso.com, files.contoso.net',
                  disabled: !isAllowlistMode,
                  onGetErrorMessage: this.validateAllowedHosts.bind(this),
                  deferredValidationTime: 200,
                }),
                PropertyPaneTextField('allowedPathPrefixes', {
                  label: 'Allowed path prefixes (comma-separated)',
                  description:
                    'Optional site-relative path prefixes that the URL must start with. Example: /sites/Reports/Dashboards/',
                  disabled: isPresetLocked,
                  onGetErrorMessage: this.validateAllowedPathPrefixes.bind(this),
                  deferredValidationTime: 200,
                }),
                PropertyPaneTextField('allowedFileExtensions', {
                  label: 'Allowed file extensions (comma-separated)',
                  description: 'Optional file extensions, e.g. .html, .htm',
                  disabled: isPresetLocked,
                  onGetErrorMessage: this.validateAllowedFileExtensions.bind(this),
                  deferredValidationTime: 200,
                }),
              ],
            },
            {
              groupName: 'Tenant configuration',
              groupFields: [
                PropertyPaneTextField('tenantConfigUrl', {
                  label: 'Tenant config JSON URL',
                  description:
                    'Optional JSON config (site-relative or absolute URL in this tenant).',
                  onGetErrorMessage: this.validateTenantConfigUrl.bind(this),
                  deferredValidationTime: 200,
                }),
                PropertyPaneDropdown('tenantConfigMode', {
                  label: 'Tenant config mode',
                  options: [
                    { key: 'Merge', text: 'Merge (use when fields are empty)' },
                    { key: 'Override', text: 'Override (config wins)' },
                  ],
                  disabled: !this.properties.tenantConfigUrl,
                }),
              ],
            },
            {
              groupName: 'Cache & refresh',
              groupFields: [
                PropertyPaneDropdown('cacheBusterMode', {
                  label: 'Cache-busting mode',
                  options: [
                    { key: 'None', text: 'None' },
                    { key: 'Timestamp', text: 'Timestamp query param' },
                    { key: 'FileLastModified', text: 'SharePoint file modified time' },
                  ],
                  disabled: isPresetLocked,
                }),
                PropertyPaneTextField('cacheBusterParamName', {
                  label: 'Cache-buster parameter name',
                  description: 'Defaults to "v" when empty.',
                  disabled: cacheBusterMode === 'None' || isPresetLocked,
                }),
                PropertyPaneSlider('refreshIntervalMinutes', {
                  label: 'Auto-refresh interval (minutes)',
                  min: 0,
                  max: 120,
                  step: 1,
                }),
              ],
            },
            {
              groupName: 'Layout',
              groupFields: [
                PropertyPaneDropdown('heightMode', {
                  label: 'Height mode',
                  options: [
                    { key: 'Fixed', text: 'Fixed' },
                    { key: 'Viewport', text: 'Viewport (100vh)' },
                  ],
                }),
                PropertyPaneSlider('fixedHeightPx', {
                  label: 'Fixed height (px)',
                  min: 200,
                  max: 2000,
                  step: 50,
                }),
              ],
            },
            {
              groupName: 'Iframe',
              groupFields: [
                PropertyPaneTextField('iframeTitle', {
                  label: 'Iframe title',
                  description: 'Used for accessibility. Defaults to "Universal HTML Viewer".',
                }),
                PropertyPaneDropdown('iframeLoading', {
                  label: 'Loading mode',
                  options: [
                    { key: '', text: 'Browser default' },
                    { key: 'lazy', text: 'Lazy' },
                    { key: 'eager', text: 'Eager' },
                  ],
                }),
                PropertyPaneDropdown('sandboxPreset', {
                  label: 'Sandbox preset',
                  options: [
                    { key: 'None', text: 'None (no sandbox)' },
                    { key: 'Relaxed', text: 'Relaxed' },
                    { key: 'Strict', text: 'Strict' },
                    { key: 'Custom', text: 'Custom (use tokens below)' },
                  ],
                  disabled: isPresetLocked,
                }),
                PropertyPaneTextField('iframeSandbox', {
                  label: 'Sandbox tokens',
                  description:
                    'Space-separated sandbox tokens used when Sandbox preset is "Custom". Example: allow-scripts allow-same-origin',
                  disabled: !isCustomSandbox || isPresetLocked,
                }),
                PropertyPaneTextField('iframeAllow', {
                  label: 'Permissions policy (allow)',
                  description:
                    'Optional iframe allow attribute. Example: fullscreen; clipboard-read; clipboard-write',
                }),
                PropertyPaneDropdown('iframeReferrerPolicy', {
                  label: 'Referrer policy',
                  options: [
                    { key: '', text: 'Browser default' },
                    { key: 'no-referrer', text: 'no-referrer' },
                    { key: 'no-referrer-when-downgrade', text: 'no-referrer-when-downgrade' },
                    { key: 'origin', text: 'origin' },
                    { key: 'origin-when-cross-origin', text: 'origin-when-cross-origin' },
                    { key: 'same-origin', text: 'same-origin' },
                    { key: 'strict-origin', text: 'strict-origin' },
                    { key: 'strict-origin-when-cross-origin', text: 'strict-origin-when-cross-origin' },
                    { key: 'unsafe-url', text: 'unsafe-url' },
                  ],
                }),
                PropertyPaneSlider('iframeLoadTimeoutSeconds', {
                  label: 'Iframe load timeout (seconds)',
                  min: 0,
                  max: 60,
                  step: 1,
                  disabled: isPresetLocked,
                }),
              ],
            },
            {
              groupName: 'Diagnostics',
              groupFields: [
                PropertyPaneToggle('showDiagnostics', {
                  label: 'Show diagnostics panel',
                  onText: 'On',
                  offText: 'Off',
                }),
              ],
            },
          ],
        },
      ],
    };
  }

  private getIframeHeightStyle(props: IUniversalHtmlViewerWebPartProps): string {
    const heightMode: HeightMode = props.heightMode || 'Fixed';

    if (heightMode === 'Viewport') {
      return 'height:100vh;';
    }

    const fixedHeightPx: number =
      typeof props.fixedHeightPx === 'number' && props.fixedHeightPx > 0
        ? props.fixedHeightPx
        : 800;

    return `height:${fixedHeightPx}px;`;
  }

  protected onDispose(): void {
    this.clearRefreshTimer();
    this.clearIframeLoadTimeout();
  }

  private async renderAsync(): Promise<void> {
    this.clearIframeLoadTimeout();
    const pageUrl: string = this.getCurrentPageUrl();
    const { effectiveProps, tenantConfig } = await this.getEffectiveProperties(pageUrl);
    this.lastEffectiveProps = effectiveProps;
    this.lastTenantConfig = tenantConfig;

    const htmlSourceMode: HtmlSourceMode = effectiveProps.htmlSourceMode || 'FullUrl';
    const currentDashboardId: string | undefined = this.getEffectiveDashboardId(
      effectiveProps,
      pageUrl,
    );

    const finalUrl: string | undefined = buildFinalUrl({
      htmlSourceMode,
      fullUrl: effectiveProps.fullUrl,
      basePath: effectiveProps.basePath,
      relativePath: effectiveProps.relativePath,
      dashboardId: effectiveProps.dashboardId,
      defaultFileName: effectiveProps.defaultFileName,
      queryStringParamName: effectiveProps.queryStringParamName,
      pageUrl,
    });

    if (!finalUrl) {
      this.clearRefreshTimer();
      this.clearIframeLoadTimeout();
      this.domElement.innerHTML = this.buildMessageHtml(
        'UniversalHtmlViewer: No URL configured. Please update the web part settings.',
        this.buildDiagnosticsHtml(
          this.buildDiagnosticsData({
            htmlSourceMode,
            pageUrl,
            finalUrl,
            tenantConfigLoaded: !!tenantConfig,
          }, effectiveProps),
          effectiveProps,
        ),
      );
      return;
    }

    const validationOptions: UrlValidationOptions = this.buildUrlValidationOptions(
      pageUrl,
      effectiveProps,
    );
    this.lastValidationOptions = validationOptions;

    if (!isUrlAllowed(finalUrl, validationOptions)) {
      this.clearRefreshTimer();
      this.clearIframeLoadTimeout();
      this.domElement.innerHTML = this.buildMessageHtml(
        'UniversalHtmlViewer: The target URL is invalid or not allowed.',
        this.buildDiagnosticsHtml(
          this.buildDiagnosticsData({
            htmlSourceMode,
            pageUrl,
            finalUrl,
            validationOptions,
            tenantConfigLoaded: !!tenantConfig,
          }, effectiveProps),
          effectiveProps,
        ),
      );
      return;
    }

    const iframeHeightStyle: string = this.getIframeHeightStyle(effectiveProps);
    const cacheBusterMode: CacheBusterMode = effectiveProps.cacheBusterMode || 'None';
    this.lastCacheBusterMode = cacheBusterMode;
    const cacheBusterParamName: string = this.normalizeCacheBusterParamName(
      effectiveProps.cacheBusterParamName,
    );
    const resolvedUrl: string = await this.resolveUrlWithCacheBuster(
      finalUrl,
      cacheBusterMode,
      cacheBusterParamName,
      pageUrl,
    );

    this.renderIframe(
      resolvedUrl,
      iframeHeightStyle,
      this.buildDiagnosticsHtml(
        this.buildDiagnosticsData({
          htmlSourceMode,
          pageUrl,
          finalUrl,
          resolvedUrl,
          validationOptions,
          cacheBusterMode,
          tenantConfigLoaded: !!tenantConfig,
        }, effectiveProps),
        effectiveProps,
      ),
      finalUrl,
      cacheBusterMode,
      cacheBusterParamName,
      pageUrl,
      validationOptions,
      effectiveProps,
      currentDashboardId,
    );
    this.currentBaseUrl = finalUrl;
    this.setupIframeLoadFallback(resolvedUrl, effectiveProps);
    this.setupAutoRefresh(
      finalUrl,
      cacheBusterMode,
      cacheBusterParamName,
      pageUrl,
      effectiveProps,
    );
  }

  private renderIframe(
    url: string,
    iframeHeightStyle: string,
    diagnosticsHtml: string,
    baseUrl: string,
    cacheBusterMode: CacheBusterMode,
    cacheBusterParamName: string,
    pageUrl: string,
    validationOptions: UrlValidationOptions,
    effectiveProps: IUniversalHtmlViewerWebPartProps,
    currentDashboardId?: string,
  ): void {
    const iframeTitle: string =
      (effectiveProps.iframeTitle || '').trim() || 'Universal HTML Viewer';
    const iframeLoading: string = this.normalizeIframeLoading(effectiveProps.iframeLoading);
    const iframeSandbox: string = this.normalizeIframeSandbox(
      effectiveProps.iframeSandbox,
      effectiveProps.sandboxPreset,
    );
    const iframeAllow: string = this.normalizeIframeAllow(effectiveProps.iframeAllow);
    const iframeReferrerPolicy: string = this.normalizeReferrerPolicy(
      effectiveProps.iframeReferrerPolicy,
    );
    const chromeHtml: string = this.buildChromeHtml(
      url,
      validationOptions,
      cacheBusterMode,
      effectiveProps,
      currentDashboardId,
    );
    const loadingHtml: string = this.buildLoadingHtml(effectiveProps);
    const iframeContainerClass: string = chromeHtml
      ? styles.iframeContainer
      : `${styles.iframeContainer} ${styles.iframeContainerNoChrome}`;

    const loadingAttribute: string = iframeLoading
      ? ` loading="${escape(iframeLoading)}"`
      : '';
    const sandboxAttribute: string = iframeSandbox
      ? ` sandbox="${escape(iframeSandbox)}"`
      : '';
    const allowAttribute: string = iframeAllow ? ` allow="${escape(iframeAllow)}"` : '';
    const referrerPolicyAttribute: string = iframeReferrerPolicy
      ? ` referrerpolicy="${escape(iframeReferrerPolicy)}"`
      : '';

    this.domElement.innerHTML = `
      <div class="${styles.universalHtmlViewer}">
        ${chromeHtml}
        <div class="${iframeContainerClass}">
          ${loadingHtml}
        <iframe class="${styles.iframe}"
          src="${escape(url)}"
          title="${escape(iframeTitle)}"
          style="${iframeHeightStyle}border:0;"
          width="100%"
          frameborder="0"${loadingAttribute}${sandboxAttribute}${allowAttribute}${referrerPolicyAttribute}
        ></iframe>
        </div>
      </div>${diagnosticsHtml}`;

    this.attachChromeHandlers(
      baseUrl,
      cacheBusterMode,
      cacheBusterParamName,
      pageUrl,
      effectiveProps,
    );
  }

  private buildUrlValidationOptions(
    currentPageUrl: string,
    effectiveProps: IUniversalHtmlViewerWebPartProps,
  ): UrlValidationOptions {
    const securityMode: UrlSecurityMode = effectiveProps.securityMode || 'StrictTenant';
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

  private async getEffectiveProperties(
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

    const shouldApply = (currentValue: unknown): boolean => {
      if (!applyIfEmpty) {
        return true;
      }
      if (currentValue === undefined || currentValue === null) {
        return true;
      }
      if (typeof currentValue === 'string') {
        return currentValue.trim().length === 0;
      }
      return false;
    };

    const nextPropsRecord = nextProps as unknown as Record<string, unknown>;

    Object.entries(tenantConfig).forEach(([key, value]) => {
      if (value === undefined) {
        return;
      }
      if (key === 'dashboardList' && Array.isArray(value)) {
        value = value.join(',');
      }
      const currentValue = nextPropsRecord[key];
      if (shouldApply(currentValue)) {
        nextPropsRecord[key] = value;
      }
    });

    const preset: ConfigurationPreset = nextProps.configurationPreset || 'Custom';
    if (nextProps.lockPresetSettings && preset !== 'Custom') {
      this.applyPreset(preset, nextProps);
    }

    return nextProps;
  }

  private async tryLoadTenantConfig(
    pageUrl: string,
    props: IUniversalHtmlViewerWebPartProps,
  ): Promise<ITenantConfig | undefined> {
    const rawUrl: string = (props.tenantConfigUrl || '').trim();
    if (!rawUrl) {
      return undefined;
    }

    const resolvedUrl: string | undefined = this.resolveTenantConfigUrl(rawUrl, pageUrl);
    if (!resolvedUrl) {
      return undefined;
    }

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        resolvedUrl,
        SPHttpClient.configurations.v1,
      );
      if (!response.ok) {
        return undefined;
      }

      const data = await response.json();
      if (!data || typeof data !== 'object') {
        return undefined;
      }

      return data as ITenantConfig;
    } catch {
      return undefined;
    }
  }

  private resolveTenantConfigUrl(rawUrl: string, pageUrl: string): string | undefined {
    const trimmed: string = rawUrl.trim();
    if (!trimmed) {
      return undefined;
    }

    const currentUrl: URL = new URL(pageUrl || this.context.pageContext.web.absoluteUrl);
    const origin: string = currentUrl.origin;

    if (trimmed.startsWith('http://')) {
      return undefined;
    }

    if (trimmed.startsWith('https://')) {
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

    if (trimmed.startsWith('/')) {
      return `${origin}${trimmed}`;
    }

    const webServerRelative: string = this.context.pageContext.web.serverRelativeUrl || '/';
    const normalizedBase: string = webServerRelative.endsWith('/')
      ? webServerRelative.slice(0, -1)
      : webServerRelative;
    return `${origin}${normalizedBase}/${trimmed}`;
  }

  private getEffectiveDashboardId(
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

  private applyPreset(
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
    props.allowedFileExtensions = '.html,.htm';
    props.showChrome = true;
    props.showOpenInNewTab = true;
    props.showRefreshButton = true;
    props.showStatus = true;
    props.showLastUpdated = true;
    props.showLoadingIndicator = true;
    props.showConfigActions = true;
    props.showDashboardSelector = false;
    props.chromeDensity = 'Comfortable';
    props.iframeLoadTimeoutSeconds = 10;

    if (!props.chromeTitle || props.chromeTitle.trim().length === 0) {
      props.chromeTitle = 'Universal HTML Viewer';
    }

    if (basePathPrefix) {
      props.allowedPathPrefixes = basePathPrefix;
    }

    switch (preset) {
      case 'SharePointLibraryStrict':
        props.securityMode = 'StrictTenant';
        props.cacheBusterMode = 'FileLastModified';
        props.sandboxPreset = 'Strict';
        break;
      case 'SharePointLibraryRelaxed':
        props.securityMode = 'StrictTenant';
        props.cacheBusterMode = 'FileLastModified';
        props.sandboxPreset = 'Relaxed';
        break;
      case 'AllowlistCDN':
        props.securityMode = 'Allowlist';
        props.cacheBusterMode = 'Timestamp';
        props.sandboxPreset = 'Relaxed';
        break;
      case 'AnyHttps':
        props.securityMode = 'AnyHttps';
        props.cacheBusterMode = 'Timestamp';
        props.sandboxPreset = 'None';
        break;
      default:
        break;
    }
  }

  private normalizeBasePathForPrefix(value?: string): string {
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

  private parseHosts(value?: string): string[] {
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

  private parsePathPrefixes(value?: string): string[] {
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

  private parseFileExtensions(value?: string): string[] {
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

  private normalizeCacheBusterParamName(value?: string): string {
    const trimmed: string = (value || '').trim();
    if (!trimmed) {
      return 'v';
    }
    if (!/^[a-zA-Z0-9_-]+$/.test(trimmed)) {
      return 'v';
    }
    return trimmed;
  }

  private normalizeIframeLoading(value?: string): string {
    const normalized: string = (value || '').trim().toLowerCase();
    if (normalized === 'lazy' || normalized === 'eager') {
      return normalized;
    }
    return '';
  }

  private normalizeIframeSandbox(value?: string, preset?: string): string {
    const normalizedPreset: string = (preset || '').trim().toLowerCase();
    if (normalizedPreset && normalizedPreset !== 'custom') {
      if (normalizedPreset === 'relaxed') {
        return 'allow-same-origin allow-scripts allow-forms allow-popups';
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

  private normalizeIframeAllow(value?: string): string {
    const trimmed: string = (value || '').trim();
    if (!trimmed) {
      return '';
    }
    return trimmed.replace(/[^a-zA-Z0-9;=(),\s\-:'"*]/g, '');
  }

  private normalizeReferrerPolicy(value?: string): string {
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

  private setupAutoRefresh(
    baseUrl: string,
    cacheBusterMode: CacheBusterMode,
    cacheBusterParamName: string,
    pageUrl: string,
    effectiveProps: IUniversalHtmlViewerWebPartProps,
  ): void {
    const refreshIntervalMs: number = this.getRefreshIntervalMs(effectiveProps);

    if (refreshIntervalMs <= 0) {
      this.clearRefreshTimer();
      return;
    }

    this.clearRefreshTimer();
    if (typeof window !== 'undefined') {
      this.refreshTimerId = window.setInterval(() => {
        this.refreshIframe(baseUrl, cacheBusterMode, cacheBusterParamName, pageUrl).catch(() => {
          return undefined;
        });
      }, refreshIntervalMs);
    }
  }

  private getRefreshIntervalMs(props: IUniversalHtmlViewerWebPartProps): number {
    const minutes: number = props.refreshIntervalMinutes || 0;
    if (typeof minutes !== 'number' || minutes <= 0) {
      return 0;
    }
    return minutes * 60 * 1000;
  }

  private async refreshIframe(
    baseUrl: string,
    cacheBusterMode: CacheBusterMode,
    cacheBusterParamName: string,
    pageUrl: string,
  ): Promise<void> {
    const iframe: HTMLIFrameElement | null = this.domElement.querySelector('iframe');
    if (!iframe) {
      return;
    }

    this.setLoadingVisible(true);

    const refreshedUrl: string = await this.resolveUrlWithCacheBuster(
      baseUrl,
      cacheBusterMode,
      cacheBusterParamName,
      pageUrl,
    );

    iframe.src = refreshedUrl;
    this.updateStatusBadge(this.lastValidationOptions, cacheBusterMode, this.lastEffectiveProps);
  }

  private clearRefreshTimer(): void {
    if (this.refreshTimerId && typeof window !== 'undefined') {
      window.clearInterval(this.refreshTimerId);
    }
    this.refreshTimerId = undefined;
  }

  private setupIframeLoadFallback(
    url: string,
    effectiveProps?: IUniversalHtmlViewerWebPartProps,
  ): void {
    this.clearIframeLoadTimeout();

    const props: IUniversalHtmlViewerWebPartProps =
      effectiveProps || this.lastEffectiveProps || this.properties;
    const timeoutSeconds: number = this.getIframeLoadTimeoutSeconds(props);
    if (timeoutSeconds <= 0) {
      return;
    }

    const iframe: HTMLIFrameElement | null = this.domElement.querySelector('iframe');
    if (!iframe || typeof window === 'undefined') {
      return;
    }

    iframe.addEventListener('load', () => {
      this.clearIframeLoadTimeout();
      this.setLoadingVisible(false);
    });

    this.iframeLoadTimeoutId = window.setTimeout(() => {
      this.clearRefreshTimer();
      this.domElement.innerHTML = this.buildMessageHtml(
        'UniversalHtmlViewer: The content did not load in time. It may be blocked by SharePoint security headers.',
        `${this.buildOpenInNewTabHtml(url)}${this.buildDiagnosticsHtml(
          this.buildDiagnosticsData({
            resolvedUrl: url,
            timeoutSeconds,
          }, props),
          props,
        )}`,
      );
    }, timeoutSeconds * 1000);
  }

  private getIframeLoadTimeoutSeconds(props: IUniversalHtmlViewerWebPartProps): number {
    const configuredSeconds = props.iframeLoadTimeoutSeconds;
    if (typeof configuredSeconds !== 'number') {
      return 10;
    }
    if (configuredSeconds <= 0) {
      return 0;
    }
    return configuredSeconds;
  }

  private clearIframeLoadTimeout(): void {
    if (this.iframeLoadTimeoutId && typeof window !== 'undefined') {
      window.clearTimeout(this.iframeLoadTimeoutId);
    }
    this.iframeLoadTimeoutId = undefined;
  }

  private async resolveUrlWithCacheBuster(
    baseUrl: string,
    cacheBusterMode: CacheBusterMode,
    cacheBusterParamName: string,
    pageUrl: string,
  ): Promise<string> {
    if (cacheBusterMode === 'Timestamp') {
      const timestamp = Date.now();
      this.lastCacheLabel = new Date(timestamp).toLocaleString();
      return this.appendQueryParam(baseUrl, cacheBusterParamName, `${timestamp}`);
    }

    if (cacheBusterMode === 'FileLastModified') {
      const cacheValue = await this.tryGetFileLastModifiedCacheValue(baseUrl, pageUrl);
      if (cacheValue) {
        this.lastCacheLabel = cacheValue.label;
        return this.appendQueryParam(baseUrl, cacheBusterParamName, cacheValue.value);
      }
      const fallbackTimestamp = Date.now();
      this.lastCacheLabel = new Date(fallbackTimestamp).toLocaleString();
      return this.appendQueryParam(baseUrl, cacheBusterParamName, `${fallbackTimestamp}`);
    }

    this.lastCacheLabel = undefined;
    return baseUrl;
  }

  private async tryGetFileLastModifiedCacheValue(
    url: string,
    pageUrl: string,
  ): Promise<{ value: string; label: string } | null> {
    const serverRelativePath: string | null = this.tryGetServerRelativePath(url, pageUrl);
    if (!serverRelativePath) {
      return null;
    }

    const encodedPath: string = encodeURIComponent(serverRelativePath);
    const apiUrl: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativeUrl(@p1)?@p1='${encodedPath}'&$select=TimeLastModified,ETag`;

    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        apiUrl,
        SPHttpClient.configurations.v1,
      );
      if (!response.ok) {
        return null;
      }

      const data = await response.json();
      const timeLastModified: string | undefined = data?.d?.TimeLastModified;
      const etag: string | undefined = data?.d?.ETag;

      if (etag) {
        const sanitized = etag.replace(/[^a-zA-Z0-9.-]/g, '');
        return {
          value: sanitized,
          label: `ETag ${sanitized.substring(0, 10)}`,
        };
      }

      if (timeLastModified) {
        const parsedDate = new Date(timeLastModified);
        return {
          value: parsedDate.getTime().toString(),
          label: parsedDate.toLocaleString(),
        };
      }
    } catch {
      return null;
    }

    return null;
  }

  private tryGetServerRelativePath(url: string, pageUrl: string): string | null {
    if (url.startsWith('/')) {
      return url;
    }

    try {
      const target = new URL(url);
      const current = new URL(pageUrl);
      if (target.host.toLowerCase() !== current.host.toLowerCase()) {
        return null;
      }
      return decodeURIComponent(target.pathname);
    } catch {
      return null;
    }
  }

  private appendQueryParam(url: string, name: string, value: string): string {
    const safeName: string = encodeURIComponent(name);
    const safeValue: string = encodeURIComponent(value);

    const [base, hash] = url.split('#');
    const separator: string = base.includes('?') ? '&' : '?';
    const combined: string = `${base}${separator}${safeName}=${safeValue}`;

    if (hash) {
      return `${combined}#${hash}`;
    }

    return combined;
  }

  private getCurrentPageUrl(): string {
    if (typeof window !== 'undefined' && window.location && window.location.href) {
      return window.location.href;
    }

    try {
      return this.context.pageContext.web.absoluteUrl;
    } catch {
      return '';
    }
  }

  private buildChromeHtml(
    resolvedUrl: string,
    validationOptions: UrlValidationOptions,
    cacheBusterMode: CacheBusterMode,
    props: IUniversalHtmlViewerWebPartProps,
    currentDashboardId?: string,
  ): string {
    if (props.showChrome === false) {
      return '';
    }

    const title: string =
      (props.chromeTitle || '').trim() ||
      (props.iframeTitle || '').trim() ||
      'Universal HTML Viewer';
    const subtitle: string = (props.chromeSubtitle || '').trim();
    const showOpenInNewTab: boolean = props.showOpenInNewTab !== false;
    const showRefreshButton: boolean = props.showRefreshButton !== false;
    const showStatus: boolean = props.showStatus !== false;
    const showConfigActions: boolean = props.showConfigActions === true;
    const showDashboardSelector: boolean =
      props.showDashboardSelector === true &&
      props.htmlSourceMode === 'BasePathAndDashboardId';

    const statusLabel: string = showStatus
      ? this.getStatusLabel(validationOptions, cacheBusterMode, props)
      : '';
    const statusHtml: string = statusLabel
      ? `<span class="${styles.status}" data-uhv-status>${escape(statusLabel)}</span>`
      : '';

    const openInNewTabHtml: string = showOpenInNewTab
      ? `<a class="${styles.actionLink}" href="${escape(resolvedUrl)}" target="_blank" rel="noopener noreferrer">
          Open in new tab
        </a>`
      : '';

    const refreshHtml: string = showRefreshButton
      ? `<button class="${styles.actionButton}" type="button" data-uhv-action="refresh">Refresh</button>`
      : '';

    const exportHtml: string = showConfigActions
      ? `<button class="${styles.actionButton}" type="button" data-uhv-action="export">Export</button>`
      : '';

    const importHtml: string = showConfigActions
      ? `<button class="${styles.actionButton}" type="button" data-uhv-action="import">Import</button>
         <input class="${styles.hiddenInput}" type="file" data-uhv-config-input accept="application/json" />`
      : '';

    const subtitleHtml: string = subtitle
      ? `<div class="${styles.chromeSubtitle}">${escape(subtitle)}</div>`
      : '';

    const chromeClass: string =
      (props.chromeDensity || 'Comfortable') === 'Compact'
        ? `${styles.chrome} ${styles.chromeCompact}`
        : styles.chrome;

    const dashboardHtml: string = showDashboardSelector
      ? this.buildDashboardSelectorHtml(props.dashboardList, currentDashboardId)
      : '';

    return `
      <div class="${chromeClass}">
        <div class="${styles.chromeLeft}">
          <div class="${styles.chromeTitle}">${escape(title)}</div>
          ${subtitleHtml}
        </div>
        <div class="${styles.chromeRight}">
          ${statusHtml}
          ${refreshHtml}
          ${exportHtml}
          ${importHtml}
          ${openInNewTabHtml}
        </div>
      </div>
      ${dashboardHtml}`;
  }

  private buildDashboardSelectorHtml(
    rawList: string | string[] | undefined,
    currentDashboardId?: string,
  ): string {
    const options = this.parseDashboardList(rawList);
    this.dashboardOptions = options;
    if (options.length === 0) {
      return '';
    }

    const optionsHtml = options
      .map((option) => {
        const isSelected = currentDashboardId === option.id;
        return `<option value="${escape(option.id)}"${isSelected ? ' selected' : ''}>${escape(
          option.label,
        )}</option>`;
      })
      .join('');

    return `
      <div class="${styles.dashboardBar}">
        <label class="${styles.dashboardLabel}">Dashboard</label>
        <input class="${styles.dashboardInput}" type="search" placeholder="Filter dashboards" data-uhv-dashboard-filter />
        <select class="${styles.dashboardSelect}" data-uhv-dashboard-select>
          ${optionsHtml}
        </select>
      </div>`;
  }

  private buildLoadingHtml(props: IUniversalHtmlViewerWebPartProps): string {
    if (props.showLoadingIndicator === false) {
      return '';
    }

    return `<div class="${styles.loading}" data-uhv-loading>Loading…</div>`;
  }

  private parseDashboardList(
    rawList?: string | string[],
  ): Array<{ id: string; label: string }> {
    const normalizedList: string = Array.isArray(rawList) ? rawList.join(',') : rawList || '';
    const entries = normalizedList
      .split(/[,;\n]+/g)
      .map((entry) => entry.trim())
      .filter((entry) => entry.length > 0);

    const result: Array<{ id: string; label: string }> = [];
    const seen = new Set<string>();

    for (const entry of entries) {
      let label = entry;
      let id = entry;

      if (entry.includes('|')) {
        const [left, right] = entry.split('|');
        label = (left || '').trim();
        id = (right || '').trim();
      } else if (entry.includes('=')) {
        const [left, right] = entry.split('=');
        label = (left || '').trim();
        id = (right || '').trim();
      }

      if (!id) {
        continue;
      }

      const normalizedId = id.toLowerCase();
      if (seen.has(normalizedId)) {
        continue;
      }

      seen.add(normalizedId);
      result.push({
        id,
        label: label || id,
      });
    }

    return result;
  }

  private attachChromeHandlers(
    baseUrl: string,
    cacheBusterMode: CacheBusterMode,
    cacheBusterParamName: string,
    pageUrl: string,
    effectiveProps: IUniversalHtmlViewerWebPartProps,
  ): void {
    const refreshButton: HTMLButtonElement | null = this.domElement.querySelector(
      '[data-uhv-action="refresh"]',
    );
    if (refreshButton) {
      refreshButton.addEventListener('click', () => {
        this.setLoadingVisible(true);
        const activeBaseUrl = this.currentBaseUrl || baseUrl;
        this.refreshIframe(activeBaseUrl, cacheBusterMode, cacheBusterParamName, pageUrl).catch(
          () => {
            return undefined;
          },
        );
      });
    }

    const exportButton: HTMLButtonElement | null = this.domElement.querySelector(
      '[data-uhv-action="export"]',
    );
    if (exportButton) {
      exportButton.addEventListener('click', () => {
        this.exportConfig(effectiveProps);
      });
    }

    const importButton: HTMLButtonElement | null = this.domElement.querySelector(
      '[data-uhv-action="import"]',
    );
    const importInput: HTMLInputElement | null = this.domElement.querySelector(
      '[data-uhv-config-input]',
    );
    if (importButton && importInput) {
      importButton.addEventListener('click', () => {
        importInput.value = '';
        importInput.click();
      });

      importInput.addEventListener('change', () => {
        const file: File | undefined = importInput.files?.[0];
        if (!file) {
          return;
        }
        this.importConfig(file);
      });
    }

    const dashboardSelect: HTMLSelectElement | null = this.domElement.querySelector(
      '[data-uhv-dashboard-select]',
    );
    const dashboardFilter: HTMLInputElement | null = this.domElement.querySelector(
      '[data-uhv-dashboard-filter]',
    );
    if (dashboardSelect) {
      dashboardSelect.addEventListener('change', () => {
        const selectedId: string = dashboardSelect.value;
        this.handleDashboardSelection(
          selectedId,
          effectiveProps,
          pageUrl,
          cacheBusterParamName,
        );
      });
    }
    if (dashboardFilter && dashboardSelect) {
      dashboardFilter.addEventListener('input', () => {
        this.filterDashboardOptions(dashboardFilter.value, dashboardSelect);
      });
    }

    const iframe: HTMLIFrameElement | null = this.domElement.querySelector('iframe');
    if (iframe) {
      iframe.addEventListener('load', () => {
        this.setLoadingVisible(false);
      });
    }
  }

  private setLoadingVisible(visible: boolean): void {
    const loadingElement: HTMLElement | null = this.domElement.querySelector('[data-uhv-loading]');
    if (!loadingElement) {
      return;
    }

    if (visible) {
      loadingElement.classList.remove(styles.loadingHidden);
    } else {
      loadingElement.classList.add(styles.loadingHidden);
    }
  }

  private filterDashboardOptions(
    filterValue: string,
    selectElement: HTMLSelectElement,
  ): void {
    const normalizedFilter = filterValue.trim().toLowerCase();
    const options = this.dashboardOptions;
    const filtered = normalizedFilter
      ? options.filter(
          (option) =>
            option.label.toLowerCase().includes(normalizedFilter) ||
            option.id.toLowerCase().includes(normalizedFilter),
        )
      : options;

    const currentValue = selectElement.value;
    selectElement.innerHTML = filtered
      .map((option) => {
        const isSelected = option.id === currentValue;
        return `<option value="${escape(option.id)}"${isSelected ? ' selected' : ''}>${escape(
          option.label,
        )}</option>`;
      })
      .join('');
  }

  private async handleDashboardSelection(
    dashboardId: string,
    props: IUniversalHtmlViewerWebPartProps,
    pageUrl: string,
    cacheBusterParamName: string,
  ): Promise<void> {
    const normalizedId: string = (dashboardId || '').trim();
    if (!normalizedId) {
      return;
    }

    const url = this.buildUrlFromDashboardId(props, normalizedId);
    if (!url) {
      return;
    }

    const validationOptions = this.buildUrlValidationOptions(pageUrl, props);
    this.lastValidationOptions = validationOptions;
    if (!isUrlAllowed(url, validationOptions)) {
      return;
    }

    const cacheBusterMode: CacheBusterMode = props.cacheBusterMode || 'None';
    this.lastCacheBusterMode = cacheBusterMode;

    this.setLoadingVisible(true);
    const resolvedUrl = await this.resolveUrlWithCacheBuster(
      url,
      cacheBusterMode,
      cacheBusterParamName,
      pageUrl,
    );

    const iframe: HTMLIFrameElement | null = this.domElement.querySelector('iframe');
    if (iframe) {
      iframe.src = resolvedUrl;
    }

    this.currentBaseUrl = url;
    this.setupIframeLoadFallback(resolvedUrl, props);
    this.setupAutoRefresh(url, cacheBusterMode, cacheBusterParamName, pageUrl, props);
    this.updateStatusBadge(validationOptions, cacheBusterMode, props);
  }

  private buildUrlFromDashboardId(
    props: IUniversalHtmlViewerWebPartProps,
    dashboardId: string,
  ): string | undefined {
    if (props.htmlSourceMode !== 'BasePathAndDashboardId') {
      return undefined;
    }

    const basePath: string = this.normalizeBasePathForPrefix(props.basePath);
    if (!basePath) {
      return undefined;
    }

    const fileName: string = (props.defaultFileName || '').trim() || 'index.html';
    return `${basePath}${dashboardId}/${fileName}`;
  }

  private exportConfig(props: IUniversalHtmlViewerWebPartProps): void {
    const exportData = this.buildConfigExport(props);
    const json = JSON.stringify(exportData, null, 2);
    const blob = new Blob([json], { type: 'application/json' });
    const url = URL.createObjectURL(blob);

    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = 'universal-html-viewer.config.json';
    anchor.click();

    URL.revokeObjectURL(url);
  }

  private importConfig(file: File): void {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const text = String(reader.result || '');
        const parsed = JSON.parse(text);
        if (!parsed || typeof parsed !== 'object') {
          return;
        }
        this.applyImportedConfig(parsed as Record<string, unknown>);
      } catch {
        return;
      }
    };
    reader.readAsText(file);
  }

  private applyImportedConfig(config: Record<string, unknown>): void {
    const propsRecord = this.properties as unknown as Record<string, unknown>;
    const booleanKeys = new Set<string>([
      'lockPresetSettings',
      'allowHttp',
      'showDiagnostics',
      'showChrome',
      'showOpenInNewTab',
      'showRefreshButton',
      'showStatus',
      'showLastUpdated',
      'showLoadingIndicator',
      'showConfigActions',
      'showDashboardSelector',
    ]);
    const numberKeys = new Set<string>(['fixedHeightPx', 'iframeLoadTimeoutSeconds', 'refreshIntervalMinutes']);
    const stringKeys = new Set<string>([
      'configurationPreset',
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

    Object.entries(config).forEach(([key, value]) => {
      if (value === undefined || value === null) {
        return;
      }
      if (booleanKeys.has(key)) {
        propsRecord[key] = value === true || value === 'true' || value === 1;
        return;
      }
      if (numberKeys.has(key)) {
        const parsed = typeof value === 'number' ? value : Number(value);
        if (!Number.isNaN(parsed)) {
          propsRecord[key] = parsed;
        }
        return;
      }
      if (stringKeys.has(key)) {
        propsRecord[key] = String(value);
      }
    });

    const preset: ConfigurationPreset = this.properties.configurationPreset || 'Custom';
    if (this.properties.lockPresetSettings && preset !== 'Custom') {
      this.applyPreset(preset);
    }

    this.context.propertyPane.refresh();
    this.render();
  }

  private buildConfigExport(
    props: IUniversalHtmlViewerWebPartProps,
  ): Record<string, unknown> {
    return {
      configurationPreset: props.configurationPreset || 'Custom',
      lockPresetSettings: !!props.lockPresetSettings,
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
      securityMode: props.securityMode || 'StrictTenant',
      allowHttp: !!props.allowHttp,
      allowedHosts: props.allowedHosts || '',
      allowedPathPrefixes: props.allowedPathPrefixes || '',
      allowedFileExtensions: props.allowedFileExtensions || '',
      tenantConfigUrl: props.tenantConfigUrl || '',
      tenantConfigMode: props.tenantConfigMode || 'Merge',
      cacheBusterMode: props.cacheBusterMode || 'None',
      cacheBusterParamName: props.cacheBusterParamName || 'v',
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
    };
  }

  private updateStatusBadge(
    validationOptions?: UrlValidationOptions,
    cacheBusterMode?: CacheBusterMode,
    props?: IUniversalHtmlViewerWebPartProps,
  ): void {
    const statusElement: HTMLElement | null = this.domElement.querySelector('[data-uhv-status]');
    if (!statusElement) {
      return;
    }

    const effectiveProps: IUniversalHtmlViewerWebPartProps =
      props || this.lastEffectiveProps || this.properties;
    const validation =
      validationOptions || this.lastValidationOptions || this.buildUrlValidationOptions(
        this.getCurrentPageUrl(),
        effectiveProps,
      );
    const cacheMode: CacheBusterMode =
      cacheBusterMode || this.lastCacheBusterMode || 'None';

    statusElement.textContent = this.getStatusLabel(validation, cacheMode, effectiveProps);
  }

  private getStatusLabel(
    validationOptions: UrlValidationOptions,
    cacheBusterMode: CacheBusterMode,
    props: IUniversalHtmlViewerWebPartProps,
  ): string {
    const parts: string[] = [];

    parts.push(validationOptions.securityMode);
    if (validationOptions.securityMode === 'Allowlist' && validationOptions.allowedHosts) {
      parts.push(`${validationOptions.allowedHosts.length} hosts`);
    }

    if (validationOptions.allowHttp) {
      parts.push('HTTP allowed');
    }

    if (cacheBusterMode !== 'None') {
      parts.push(`Cache: ${cacheBusterMode}`);
    }

    if (props.showLastUpdated !== false && this.lastCacheLabel) {
      parts.push(`Updated: ${this.lastCacheLabel}`);
    }

    return parts.join(' • ');
  }

  private buildMessageHtml(message: string, extraHtml?: string): string {
    const extra: string = extraHtml ? extraHtml : '';
    return `
      <div class="${styles.universalHtmlViewer}">
        <div class="${styles.message}">${escape(message)}</div>
      </div>${extra}`;
  }

  private buildOpenInNewTabHtml(url: string): string {
    const escapedUrl: string = escape(url);
    return `
      <div class="${styles.fallback}">
        <a class="${styles.fallbackLink}" href="${escapedUrl}" target="_blank" rel="noopener noreferrer">
          Open in new tab
        </a>
      </div>`;
  }

  private buildDiagnosticsHtml(
    data?: Record<string, unknown>,
    effectiveProps?: IUniversalHtmlViewerWebPartProps,
  ): string {
    const props: IUniversalHtmlViewerWebPartProps = effectiveProps || this.properties;
    if (!props.showDiagnostics || !data) {
      return '';
    }

    const json: string = JSON.stringify(data, null, 2) || '';
    const escaped: string = escape(json);

    return `
      <div class="${styles.diagnostics}">
        <div class="${styles.diagnosticsTitle}">Diagnostics</div>
        <pre class="${styles.diagnosticsBody}">${escaped}</pre>
      </div>`;
  }

  private buildDiagnosticsData(
    values: Record<string, unknown>,
    effectiveProps?: IUniversalHtmlViewerWebPartProps,
  ): Record<string, unknown> {
    const props: IUniversalHtmlViewerWebPartProps = effectiveProps || this.properties;
    return {
      timestamp: new Date().toISOString(),
      ...values,
      configurationPreset: props.configurationPreset || 'Custom',
      lockPresetSettings: !!props.lockPresetSettings,
      allowHttp: !!props.allowHttp,
      allowedHosts: this.parseHosts(props.allowedHosts),
      allowedPathPrefixes: this.parsePathPrefixes(props.allowedPathPrefixes),
      allowedFileExtensions: this.parseFileExtensions(props.allowedFileExtensions),
      securityMode: props.securityMode || 'StrictTenant',
      tenantConfigUrl: props.tenantConfigUrl || '',
      tenantConfigMode: props.tenantConfigMode || 'Merge',
      dashboardList: props.dashboardList || '',
      cacheBusterMode: props.cacheBusterMode || 'None',
      sandboxPreset: props.sandboxPreset || 'None',
      iframeSandbox: props.iframeSandbox || '',
      iframeLoadTimeoutSeconds: this.getIframeLoadTimeoutSeconds(props),
      showChrome: props.showChrome !== false,
      showLoadingIndicator: props.showLoadingIndicator !== false,
      showLastUpdated: props.showLastUpdated !== false,
      chromeDensity: props.chromeDensity || 'Comfortable',
      showConfigActions: props.showConfigActions === true,
      showDashboardSelector: props.showDashboardSelector === true,
    };
  }

  private validateFullUrl(value?: string): string {
    const trimmed: string = (value || '').trim();
    if (!trimmed) {
      return '';
    }

    const lower = trimmed.toLowerCase();
    const blockedSchemes = ['javascript', 'data', 'vbscript'];
    if (blockedSchemes.some((scheme) => lower.startsWith(`${scheme}:`))) {
      return 'Unsupported or unsafe URL scheme.';
    }

    if (trimmed.startsWith('/')) {
      return '';
    }

    if (lower.startsWith('https://')) {
      return '';
    }

    if (lower.startsWith('http://')) {
      return this.properties.allowHttp
        ? ''
        : 'HTTP is blocked by default. Enable "Allow HTTP" if required.';
    }

    return 'Enter a site-relative path or an absolute http/https URL.';
  }

  private validateBasePath(value?: string): string {
    const trimmed: string = (value || '').trim();
    if (!trimmed) {
      return '';
    }
    if (trimmed.includes('://')) {
      return 'Base path must be site-relative, e.g. /sites/Reports/Dashboards/.';
    }
    if (!trimmed.startsWith('/')) {
      return 'Base path must start with "/".';
    }
    if (trimmed.includes('?') || trimmed.includes('#')) {
      return 'Base path should not include query strings or fragments.';
    }
    if (this.hasDotSegments(trimmed)) {
      return 'Base path must not include "." or ".." segments.';
    }
    return '';
  }

  private validateAllowedHosts(value?: string): string {
    const entries = (value || '')
      .split(/[,;\s]+/g)
      .map((entry) => entry.trim())
      .filter((entry) => entry.length > 0);

    for (const entry of entries) {
      let host = entry;
      try {
        if (entry.includes('://')) {
          host = new URL(entry).hostname;
        } else {
          host = entry.split('/')[0];
        }
      } catch {
        return `Invalid host entry: "${entry}".`;
      }

      if (host.startsWith('*.')) {
        host = host.substring(1);
      }

      host = host.split(':')[0];
      if (host.startsWith('.')) {
        host = host.substring(1);
      }

      if (!/^[a-z0-9.-]+$/i.test(host) || host.length === 0) {
        return `Invalid host entry: "${entry}".`;
      }
    }

    return '';
  }

  private validateAllowedPathPrefixes(value?: string): string {
    const entries = (value || '')
      .split(/[,;\s]+/g)
      .map((entry) => entry.trim())
      .filter((entry) => entry.length > 0);

    for (const entry of entries) {
      if (!entry.startsWith('/')) {
        return `Path prefixes must start with "/": "${entry}".`;
      }
      if (entry.includes('://')) {
        return `Path prefixes must be site-relative: "${entry}".`;
      }
      if (entry.includes('?') || entry.includes('#')) {
        return `Path prefixes must not include query strings: "${entry}".`;
      }
      if (this.hasDotSegments(entry)) {
        return `Path prefixes must not include "." or "..": "${entry}".`;
      }
    }

    return '';
  }

  private validateAllowedFileExtensions(value?: string): string {
    const entries = (value || '')
      .split(/[,;\s]+/g)
      .map((entry) => entry.trim())
      .filter((entry) => entry.length > 0);

    for (const entry of entries) {
      const normalized = entry.startsWith('.') ? entry : `.${entry}`;
      if (!/^\.[a-z0-9]+$/i.test(normalized)) {
        return `Invalid extension: "${entry}". Use values like .html, .htm.`;
      }
    }

    return '';
  }

  private validateTenantConfigUrl(value?: string): string {
    const trimmed: string = (value || '').trim();
    if (!trimmed) {
      return '';
    }

    if (trimmed.startsWith('http://')) {
      return 'Tenant config should use HTTPS.';
    }

    if (trimmed.startsWith('https://')) {
      try {
        const target = new URL(trimmed);
        const current = new URL(this.getCurrentPageUrl() || this.context.pageContext.web.absoluteUrl);
        if (target.hostname.toLowerCase() !== current.hostname.toLowerCase()) {
          return 'Tenant config must be hosted in the same SharePoint tenant.';
        }
        return '';
      } catch {
        return 'Invalid tenant config URL.';
      }
    }

    if (trimmed.includes('://')) {
      return 'Tenant config must be site-relative or an absolute HTTPS URL.';
    }

    return '';
  }

  private hasDotSegments(pathname: string): boolean {
    const segments = pathname.split('/').filter((segment) => segment.length > 0);
    return segments.some((segment) => segment === '.' || segment === '..');
  }
}
