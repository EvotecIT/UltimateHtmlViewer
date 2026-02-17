import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle,
} from '@microsoft/sp-property-pane';

import styles from './UniversalHtmlViewerWebPart.module.scss';
import { buildMessageHtml } from './MarkupHelper';
import {
  buildFinalUrl,
  CacheBusterMode,
  HtmlSourceMode,
  isUrlAllowed,
  UrlSecurityMode,
  UrlValidationOptions,
} from './UrlHelper';
import {
  validateAllowedFileExtensions,
  validateAllowedHosts,
  validateAllowedPathPrefixes,
  validateBasePath,
  validateFullUrl,
  validateTenantConfigUrl,
} from './ValidationHelper';
import { UniversalHtmlViewerWebPartUiBase } from './UniversalHtmlViewerWebPartUiBase';
import { applyInlineModeBehaviors } from './InlineModeBehaviorHelper';
import { loadSharePointFileContentForInline } from './SharePointInlineContentHelper';
import {
  ConfigurationPreset,
  ContentDeliveryMode,
  IUniversalHtmlViewerWebPartProps,
} from './UniversalHtmlViewerTypes';

export default class UniversalHtmlViewerWebPart extends UniversalHtmlViewerWebPartUiBase {
  private nestedIframeHydrationCleanup: (() => void) | undefined;

  public render(): void {
    this.renderAsync().catch((error) => {
      this.clearRefreshTimer();
      this.clearIframeLoadTimeout();
      this.domElement.innerHTML = buildMessageHtml(
        'UniversalHtmlViewer: Failed to render content.',
        this.buildDiagnosticsHtml({
          error: String(error),
        }),
        styles.universalHtmlViewer,
        styles.message,
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
      propertyPath === 'contentDeliveryMode' ||
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
                    { key: 'SharePointLibraryFullPage', text: 'SharePoint library (full page)' },
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
                PropertyPaneDropdown('contentDeliveryMode', {
                  label: 'Content delivery mode',
                  options: [
                    { key: 'DirectUrl', text: 'Direct URL in iframe' },
                    {
                      key: 'SharePointFileContent',
                      text: 'SharePoint file API (inline iframe)',
                    },
                  ],
                }),
                PropertyPaneTextField('fullUrl', {
                  label: 'Full URL to HTML page',
                  description: 'Used when HTML source mode is "FullUrl".',
                  disabled: !isFullUrl,
                  onGetErrorMessage: (value?: string): string =>
                    validateFullUrl(value, !!this.properties.allowHttp),
                  deferredValidationTime: 200,
                }),
                PropertyPaneTextField('basePath', {
                  label: 'Base path (site-relative)',
                  description:
                    'Site-relative base path, used when HTML source mode is not "FullUrl". Example: /sites/Reports/Dashboards/',
                  disabled: isFullUrl,
                  onGetErrorMessage: (value?: string): string => validateBasePath(value),
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
                  onGetErrorMessage: (value?: string): string => validateAllowedHosts(value),
                  deferredValidationTime: 200,
                }),
                PropertyPaneTextField('allowedPathPrefixes', {
                  label: 'Allowed path prefixes (comma-separated)',
                  description:
                    'Optional site-relative path prefixes that the URL must start with. Example: /sites/Reports/Dashboards/',
                  disabled: isPresetLocked,
                  onGetErrorMessage: (value?: string): string =>
                    validateAllowedPathPrefixes(value),
                  deferredValidationTime: 200,
                }),
                PropertyPaneTextField('allowedFileExtensions', {
                  label: 'Allowed file extensions (comma-separated)',
                  description: 'Optional file extensions, e.g. .html, .htm',
                  disabled: isPresetLocked,
                  onGetErrorMessage: (value?: string): string =>
                    validateAllowedFileExtensions(value),
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
                  onGetErrorMessage: (value?: string): string =>
                    validateTenantConfigUrl(
                      value,
                      this.getCurrentPageUrl() || this.context.pageContext.web.absoluteUrl,
                    ),
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

  private async renderAsync(): Promise<void> {
    this.clearIframeLoadTimeout();
    this.clearNestedIframeHydration();
    const pageUrl: string = this.getCurrentPageUrl();
    const { effectiveProps, tenantConfig } = await this.getEffectiveProperties(pageUrl);
    this.lastEffectiveProps = effectiveProps;
    this.lastTenantConfig = tenantConfig;

    const htmlSourceMode: HtmlSourceMode = effectiveProps.htmlSourceMode || 'FullUrl';
    const contentDeliveryMode: ContentDeliveryMode = this.getContentDeliveryMode(
      effectiveProps,
    );
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
      this.domElement.innerHTML = buildMessageHtml(
        'UniversalHtmlViewer: No URL configured. Please update the web part settings.',
        this.buildDiagnosticsHtml(
          this.buildDiagnosticsData({
            htmlSourceMode,
            contentDeliveryMode,
            pageUrl,
            finalUrl,
            tenantConfigLoaded: !!tenantConfig,
          }, effectiveProps),
          effectiveProps,
        ),
        styles.universalHtmlViewer,
        styles.message,
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
      this.domElement.innerHTML = buildMessageHtml(
        'UniversalHtmlViewer: The target URL is invalid or not allowed.',
        this.buildDiagnosticsHtml(
          this.buildDiagnosticsData({
            htmlSourceMode,
            contentDeliveryMode,
            pageUrl,
            finalUrl,
            validationOptions,
            tenantConfigLoaded: !!tenantConfig,
          }, effectiveProps),
          effectiveProps,
        ),
        styles.universalHtmlViewer,
        styles.message,
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
    let inlineHtml: string | undefined;
    if (contentDeliveryMode === 'SharePointFileContent') {
      try {
        inlineHtml = await loadSharePointFileContentForInline(
          this.context.spHttpClient,
          this.context.pageContext.web.absoluteUrl,
          resolvedUrl,
          finalUrl,
          pageUrl,
        );
      } catch (error) {
        this.clearRefreshTimer();
        this.clearIframeLoadTimeout();
        this.domElement.innerHTML = buildMessageHtml(
          'UniversalHtmlViewer: Failed to load HTML from SharePoint file API.',
          this.buildDiagnosticsHtml(
            this.buildDiagnosticsData({
              htmlSourceMode,
              contentDeliveryMode,
              pageUrl,
              finalUrl,
              resolvedUrl,
              validationOptions,
              cacheBusterMode,
              tenantConfigLoaded: !!tenantConfig,
              error: String(error),
            }, effectiveProps),
            effectiveProps,
          ),
          styles.universalHtmlViewer,
          styles.message,
        );
        return;
      }
    }

    this.renderIframe(
      resolvedUrl,
      iframeHeightStyle,
      this.buildDiagnosticsHtml(
        this.buildDiagnosticsData({
          htmlSourceMode,
          contentDeliveryMode,
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
      inlineHtml,
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
    this.clearNestedIframeHydration();
    this.nestedIframeHydrationCleanup = applyInlineModeBehaviors({
      contentDeliveryMode,
      domElement: this.domElement,
      pageUrl,
      validationOptions,
      cacheBusterParamName,
      cacheBusterMode,
      onNavigate: (targetUrl: string, inlineCacheBusterMode: CacheBusterMode) => {
        this.currentBaseUrl = targetUrl;
        this.setLoadingVisible(true);
        this.lastCacheBusterMode = inlineCacheBusterMode;
        this.refreshIframe(
          targetUrl,
          inlineCacheBusterMode,
          cacheBusterParamName,
          pageUrl,
        ).catch(() => {
          return undefined;
        });
      },
      loadInlineHtml: async (
        sourceUrl: string,
        baseUrlForRelativeLinks: string,
      ): Promise<string | undefined> => {
        try {
          return await loadSharePointFileContentForInline(
            this.context.spHttpClient,
            this.context.pageContext.web.absoluteUrl,
            sourceUrl,
            baseUrlForRelativeLinks,
            pageUrl,
          );
        } catch { return undefined; }
      },
    });
  }
  protected async trySetIframeSrcDocFromSource(
    iframe: HTMLIFrameElement,
    sourceUrl: string,
    pageUrl: string,
    props: IUniversalHtmlViewerWebPartProps,
  ): Promise<boolean> {
    if (this.getContentDeliveryMode(props) !== 'SharePointFileContent') {
      return false;
    }
    try {
      const inlineHtml = await loadSharePointFileContentForInline(
        this.context.spHttpClient,
        this.context.pageContext.web.absoluteUrl,
        sourceUrl,
        this.currentBaseUrl || sourceUrl,
        pageUrl,
      );
      iframe.srcdoc = inlineHtml;
      return true;
    } catch { return false; }
  }
  private clearNestedIframeHydration(): void {
    if (this.nestedIframeHydrationCleanup) {
      this.nestedIframeHydrationCleanup();
      this.nestedIframeHydrationCleanup = undefined;
    }
  }
  protected onDispose(): void {
    this.clearNestedIframeHydration();
    super.onDispose();
  }
}
