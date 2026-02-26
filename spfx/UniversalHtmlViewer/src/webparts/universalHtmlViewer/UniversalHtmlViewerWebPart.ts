/* eslint-disable max-lines */
import { Version } from '@microsoft/sp-core-library';
import { SPHttpClient } from '@microsoft/sp-http';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle,
} from '@microsoft/sp-property-pane';

import styles from './UniversalHtmlViewerWebPart.module.scss';
import {
  buildActionLinkHtml,
  buildMessageHtml,
  buildOpenInNewTabHtml,
} from './MarkupHelper';
import {
  buildFinalUrl,
  CacheBusterMode,
  HeightMode,
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
  buildPageUrlWithoutInlineDeepLink,
  buildPageUrlWithInlineDeepLink,
  DEFAULT_INLINE_DEEP_LINK_PARAM,
  resolveInlineContentTarget,
} from './InlineDeepLinkHelper';
import {
  ConfigurationPreset,
  ContentDeliveryMode,
  IUniversalHtmlViewerWebPartProps,
} from './UniversalHtmlViewerTypes';

interface IInlineDeepLinkFrameMetrics {
  frameClientHeight: number;
  frameScrollHeight: number;
  documentHeight: number;
  pendingNestedFrames: number;
}

export default class UniversalHtmlViewerWebPart extends UniversalHtmlViewerWebPartUiBase {
  private nestedIframeHydrationCleanup: (() => void) | undefined;
  private initialDeepLinkScrollLockCleanup: (() => void) | undefined;
  private readonly inlineDeepLinkParamName: string = DEFAULT_INLINE_DEEP_LINK_PARAM;
  private isInlineDeepLinkPopStateWired: boolean = false;
  private isInlineDeepLinkScrollRestorationManaged: boolean = false;
  private previousInlineDeepLinkScrollRestoration: 'auto' | 'manual' | undefined;
  private hasLoggedAnyHttpsWarning: boolean = false;
  private readonly onInlineDeepLinkPopState = (): void => {
    this.render();
  };

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

    if (propertyPath === 'enableExpertSecurityModes' && newValue === false) {
      if (this.properties.securityMode === 'AnyHttps') {
        this.properties.securityMode = 'StrictTenant';
      }
      if (this.properties.configurationPreset === 'AnyHttps') {
        this.properties.configurationPreset = 'Custom';
        this.properties.lockPresetSettings = false;
      }
      this.context.propertyPane.refresh();
    }

    if (propertyPath === 'tenantConfigUrl') {
      const normalizedTenantConfigUrl: string = (this.properties.tenantConfigUrl || '').trim();
      this.properties.tenantConfigUrl = normalizedTenantConfigUrl;
      if (
        !normalizedTenantConfigUrl &&
        this.properties.tenantConfigMode &&
        this.properties.tenantConfigMode !== 'Merge'
      ) {
        this.properties.tenantConfigMode = 'Merge';
      }
      this.context.propertyPane.refresh();
    }

    if (
      propertyPath === 'htmlSourceMode' ||
      propertyPath === 'contentDeliveryMode' ||
      propertyPath === 'securityMode' ||
      propertyPath === 'sandboxPreset' ||
      propertyPath === 'cacheBusterMode' ||
      propertyPath === 'heightMode' ||
      propertyPath === 'showChrome' ||
      propertyPath === 'enableExpertSecurityModes' ||
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
    const enableExpertSecurityModes: boolean = this.properties.enableExpertSecurityModes === true;
    const securityModeOptions = [
      { key: 'StrictTenant', text: 'Strict tenant (default)' },
      { key: 'Allowlist', text: 'Tenant + allowlist' },
      ...(enableExpertSecurityModes || securityMode === 'AnyHttps'
        ? [{ key: 'AnyHttps', text: 'Any HTTPS (unsafe expert mode)' }]
        : []),
    ];
    const isAllowlistMode: boolean = securityMode === 'Allowlist';
    const heightMode: HeightMode = this.properties.heightMode || 'Fixed';
    const isInlineContentMode: boolean =
      (this.properties.contentDeliveryMode || 'DirectUrl') === 'SharePointFileContent';
    const cacheBusterMode: CacheBusterMode = this.properties.cacheBusterMode || 'None';
    const sandboxPreset: string = this.properties.sandboxPreset || 'None';
    const isCustomSandbox: boolean = sandboxPreset === 'Custom';
    const showChrome: boolean = this.properties.showChrome !== false;
    const preset: ConfigurationPreset = this.properties.configurationPreset || 'Custom';
    const configurationPresetOptions = [
      { key: 'Custom', text: 'Custom (manual settings)' },
      { key: 'SharePointLibraryRelaxed', text: 'SharePoint library (relaxed)' },
      { key: 'SharePointLibraryFullPage', text: 'SharePoint library (full page)' },
      { key: 'SharePointLibraryStrict', text: 'SharePoint library (strict)' },
      { key: 'AllowlistCDN', text: 'Allowlist CDN' },
      ...(enableExpertSecurityModes || preset === 'AnyHttps'
        ? [{ key: 'AnyHttps', text: 'Any HTTPS (unsafe expert mode)' }]
        : []),
    ];
    const isPresetLocked: boolean =
      !!this.properties.lockPresetSettings && preset !== 'Custom';
    const showDashboardSelector: boolean = this.properties.showDashboardSelector === true;

    return {
      pages: [
        {
          header: {
            description: 'Start with Quick setup. Advanced options are lower.',
          },
          groups: [
            {
              groupName: 'Quick setup (Most used)',
              groupFields: [
                PropertyPaneDropdown('configurationPreset', {
                  label: 'Configuration preset',
                  options: configurationPresetOptions,
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
              ],
            },
            {
              groupName: 'Source (Required)',
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
              groupName: 'Layout (Recommended)',
              groupFields: [
                PropertyPaneDropdown('heightMode', {
                  label: 'Height mode',
                  options: [
                    { key: 'Fixed', text: 'Fixed' },
                    { key: 'Viewport', text: 'Viewport (100vh)' },
                    { key: 'Auto', text: 'Auto (content height)' },
                  ],
                }),
                PropertyPaneToggle('fitContentWidth', {
                  label: 'Fit content to width (inline mode)',
                  onText: 'On',
                  offText: 'Off',
                  disabled: !isInlineContentMode,
                }),
                PropertyPaneSlider('fixedHeightPx', {
                  label:
                    heightMode === 'Auto' ? 'Minimum height (px)' : 'Fixed height (px)',
                  min: 200,
                  max: 2000,
                  step: 50,
                  disabled: heightMode === 'Viewport',
                }),
              ],
            },
            {
              groupName: 'Display & UX (Advanced)',
              groupFields: [
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
              groupName: 'Security (Advanced)',
              groupFields: [
                PropertyPaneToggle('enableExpertSecurityModes', {
                  label: 'Enable expert security modes (unsafe)',
                  onText: 'Enabled',
                  offText: 'Disabled',
                  disabled: isPresetLocked,
                }),
                PropertyPaneDropdown('securityMode', {
                  label: 'Security mode',
                  options: securityModeOptions,
                  disabled: isPresetLocked,
                }),
                PropertyPaneToggle('allowHttp', {
                  label: 'Allow HTTP (unsafe)',
                  onText: 'Allow',
                  offText: 'Block',
                  disabled: isPresetLocked,
                }),
                PropertyPaneToggle('allowQueryStringPageOverride', {
                  label: 'Allow uhvPage query override (inline mode)',
                  onText: 'On',
                  offText: 'Off',
                  disabled: !isInlineContentMode || isPresetLocked,
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
                  description: 'Optional file extensions, e.g. .html, .htm, .aspx',
                  disabled: isPresetLocked,
                  onGetErrorMessage: (value?: string): string =>
                    validateAllowedFileExtensions(value),
                  deferredValidationTime: 200,
                }),
              ],
            },
            {
              groupName: 'Tenant configuration (Advanced)',
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
                  disabled: !(this.properties.tenantConfigUrl || '').trim(),
                }),
              ],
            },
            {
              groupName: 'Cache & refresh (Advanced)',
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
                PropertyPaneSlider('inlineContentCacheTtlSeconds', {
                  label: 'Inline content cache TTL (seconds)',
                  min: 0,
                  max: 300,
                  step: 5,
                  disabled: !isInlineContentMode,
                }),
              ],
            },
            {
              groupName: 'Iframe (Advanced)',
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
    this.clearInitialDeepLinkScrollLock();
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
    this.configureInlineDeepLinkPopState(contentDeliveryMode === 'SharePointFileContent');
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
    const resolvedContentTarget = resolveInlineContentTarget({
      pageUrl,
      fallbackUrl: finalUrl,
      queryParamName: this.inlineDeepLinkParamName,
      validationOptions,
      allowDeepLinkOverride: effectiveProps.allowQueryStringPageOverride === true,
    });
    const requestedDeepLinkValue: string = resolvedContentTarget.requestedDeepLinkValue;
    const hasRequestedDeepLink: boolean = resolvedContentTarget.hasRequestedDeepLink;
    const shouldResetHostScrollToTopOnInitialDeepLink: boolean =
      contentDeliveryMode === 'SharePointFileContent' && hasRequestedDeepLink;
    if (this.isScrollTraceEnabled()) {
      this.emitScrollTrace(
        'deep-link-evaluation',
        {
          contentDeliveryMode,
          hasRequestedDeepLink,
          shouldResetHostScrollToTopOnInitialDeepLink,
          requestedDeepLinkValue,
        },
      );
    }
    if (shouldResetHostScrollToTopOnInitialDeepLink) {
      this.applyInitialDeepLinkScrollLock();
    }
    if (resolvedContentTarget.isRejectedRequestedDeepLink) {
      this.clearRefreshTimer();
      this.clearIframeLoadTimeout();
      const resetToDefaultHtml = this.buildResetToDefaultDashboardHtml(pageUrl);
      this.domElement.innerHTML = buildMessageHtml(
        'UniversalHtmlViewer: The requested deep link is invalid or not allowed.',
        `${resetToDefaultHtml}${this.buildDiagnosticsHtml(
          this.buildDiagnosticsData({
            htmlSourceMode,
            contentDeliveryMode,
            pageUrl,
            finalUrl,
            requestedDeepLinkValue,
            validationOptions,
            tenantConfigLoaded: !!tenantConfig,
          }, effectiveProps),
          effectiveProps,
        )}`,
        styles.universalHtmlViewer,
        styles.message,
      );
      return;
    }
    const initialContentUrl: string = resolvedContentTarget.initialContentUrl;
    if (validationOptions.securityMode === 'AnyHttps') {
      this.logAnyHttpsWarningOnce();
    }

    if (!isUrlAllowed(initialContentUrl, validationOptions)) {
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
            initialContentUrl,
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
    const inlineContentCacheTtlMs = this.getInlineContentCacheTtlMs(effectiveProps);
    const cacheBusterParamName: string = this.normalizeCacheBusterParamName(
      effectiveProps.cacheBusterParamName,
    );
    const resolvedUrl: string = await this.resolveUrlWithCacheBuster(
      initialContentUrl,
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
          initialContentUrl,
          pageUrl,
          SPHttpClient.configurations.v1,
          {
            cacheTtlMs: inlineContentCacheTtlMs,
          },
        );
      } catch (error) {
        this.clearRefreshTimer();
        this.clearIframeLoadTimeout();
        const statusCode = this.getInlineLoadErrorStatusCode(error);
        const isAccessDenied = statusCode === 401 || statusCode === 403;
        const resetToDefaultHtml = this.buildResetToDefaultDashboardHtml(pageUrl);
        const accessHelpHtml = `${buildOpenInNewTabHtml(
          initialContentUrl,
          styles.fallback,
          styles.fallbackLink,
          'Open file in SharePoint / Request access',
        )}${resetToDefaultHtml}<div class="${styles.fallback}">${
          isAccessDenied
            ? 'Access was denied. Use "Request access" on the opened SharePoint page.'
            : 'If access is denied, use "Request access" on the opened SharePoint page.'
        }</div>`;
        this.domElement.innerHTML = buildMessageHtml(
          isAccessDenied
            ? 'UniversalHtmlViewer: You do not have access to this report page.'
            : 'UniversalHtmlViewer: Failed to load HTML from SharePoint file API.',
          `${accessHelpHtml}${this.buildDiagnosticsHtml(
            this.buildDiagnosticsData({
              htmlSourceMode,
              contentDeliveryMode,
              pageUrl,
              finalUrl,
              initialContentUrl,
              resolvedUrl,
              statusCode,
              validationOptions,
              cacheBusterMode,
              tenantConfigLoaded: !!tenantConfig,
              error: String(error),
            }, effectiveProps),
            effectiveProps,
          )}`,
          styles.universalHtmlViewer,
          styles.message,
        );
        return;
      }
    }

    this.currentBaseUrl = initialContentUrl;
    this.renderIframe(
      resolvedUrl,
      iframeHeightStyle,
      this.buildDiagnosticsHtml(
        this.buildDiagnosticsData({
          htmlSourceMode,
          contentDeliveryMode,
          pageUrl,
          finalUrl,
          initialContentUrl,
          resolvedUrl,
          validationOptions,
          cacheBusterMode,
          tenantConfigLoaded: !!tenantConfig,
        }, effectiveProps),
        effectiveProps,
      ),
      initialContentUrl,
      cacheBusterMode,
      cacheBusterParamName,
      pageUrl,
      validationOptions,
      effectiveProps,
      currentDashboardId,
      inlineHtml,
    );
    this.setupIframeLoadFallback(resolvedUrl, effectiveProps);
    this.setupAutoRefresh(
      initialContentUrl,
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
      heightMode: effectiveProps.heightMode || 'Fixed',
      fixedHeightPx:
        typeof effectiveProps.fixedHeightPx === 'number' && effectiveProps.fixedHeightPx > 0
          ? effectiveProps.fixedHeightPx
          : 800,
      fitContentWidth: effectiveProps.fitContentWidth === true,
      onNavigate: (targetUrl: string, inlineCacheBusterMode: CacheBusterMode) => {
        this.currentBaseUrl = targetUrl;
        this.setLoadingVisible(true);
        this.lastCacheBusterMode = inlineCacheBusterMode;
        const updatedPageUrl = this.getCurrentPageUrl();
        this.onNavigatedToUrl(targetUrl, updatedPageUrl);
        const navigatedPageUrl = this.getCurrentPageUrl();
        if (contentDeliveryMode === 'SharePointFileContent') {
          this.applyInitialDeepLinkScrollLock();
        }
        this.setupAutoRefresh(
          targetUrl,
          inlineCacheBusterMode,
          cacheBusterParamName,
          navigatedPageUrl,
          this.lastEffectiveProps || this.properties,
        );
        this.updateOpenInNewTabLink(
          targetUrl,
          navigatedPageUrl,
          this.lastEffectiveProps || this.properties,
        );
        this.refreshIframe(
          targetUrl,
          inlineCacheBusterMode,
          cacheBusterParamName,
          navigatedPageUrl,
          true,
          false,
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
            this.getCurrentPageUrl(),
            SPHttpClient.configurations.v1,
            {
              cacheTtlMs: this.getInlineContentCacheTtlMs(
                this.lastEffectiveProps || this.properties,
              ),
            },
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
    bypassInlineContentCache: boolean = false,
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
        SPHttpClient.configurations.v1,
        {
          cacheTtlMs: this.getInlineContentCacheTtlMs(props),
          bypassCache: bypassInlineContentCache,
        },
      );
      iframe.srcdoc = inlineHtml;
      return true;
    } catch { return false; }
  }
  private getInlineContentCacheTtlMs(props: IUniversalHtmlViewerWebPartProps): number {
    const rawSeconds = props.inlineContentCacheTtlSeconds;
    if (typeof rawSeconds !== 'number' || !Number.isFinite(rawSeconds)) {
      return 15000;
    }
    if (rawSeconds <= 0) {
      return 0;
    }

    return Math.round(rawSeconds * 1000);
  }
  private logAnyHttpsWarningOnce(): void {
    if (this.hasLoggedAnyHttpsWarning) {
      return;
    }

    this.hasLoggedAnyHttpsWarning = true;
    // eslint-disable-next-line no-console
    console.warn(
      'UniversalHtmlViewer: AnyHttps mode is active. Restrict this mode to trusted and controlled scenarios.',
    );
  }
  protected onNavigatedToUrl(targetUrl: string, pageUrl: string): void {
    const effectiveProps: IUniversalHtmlViewerWebPartProps =
      this.lastEffectiveProps || this.properties;
    if (this.getContentDeliveryMode(effectiveProps) !== 'SharePointFileContent') {
      return;
    }

    const currentPageUrl: string = (pageUrl || '').trim() || this.getCurrentPageUrl();
    const updatedPageUrl = buildPageUrlWithInlineDeepLink({
      pageUrl: currentPageUrl,
      targetUrl,
      queryParamName: this.inlineDeepLinkParamName,
    });
    if (
      !updatedPageUrl ||
      updatedPageUrl === currentPageUrl ||
      typeof window === 'undefined' ||
      !window.history ||
      typeof window.history.pushState !== 'function'
    ) {
      return;
    }

    try {
      window.history.pushState(window.history.state, '', updatedPageUrl);
    } catch {
      return;
    }
  }
  private buildResetToDefaultDashboardHtml(pageUrl: string): string {
    const resetUrl = buildPageUrlWithoutInlineDeepLink({
      pageUrl,
      queryParamName: this.inlineDeepLinkParamName,
    });
    if (!resetUrl || resetUrl === pageUrl) {
      return '';
    }

    return buildActionLinkHtml(
      resetUrl,
      styles.fallback,
      styles.fallbackLink,
      'Reset to default dashboard',
      false,
    );
  }
  private getInlineLoadErrorStatusCode(error: unknown): number | undefined {
    const source = error as
      | {
          status?: unknown;
          response?: {
            status?: unknown;
          };
        }
      | undefined;
    const directStatus = source?.status;
    if (typeof directStatus === 'number' && directStatus >= 100 && directStatus <= 599) {
      return directStatus;
    }

    const nestedStatus = source?.response?.status;
    if (typeof nestedStatus === 'number' && nestedStatus >= 100 && nestedStatus <= 599) {
      return nestedStatus;
    }

    const message = String(error || '');
    const statusMatch = message.match(/\b([1-5]\d{2})\b/);
    if (!statusMatch) {
      return undefined;
    }

    const parsed = Number(statusMatch[1]);
    if (Number.isNaN(parsed)) {
      return undefined;
    }
    return parsed;
  }
  private configureInlineDeepLinkPopState(enabled: boolean): void {
    if (typeof window === 'undefined') {
      return;
    }

    this.configureInlineDeepLinkScrollRestoration(enabled);

    if (enabled && !this.isInlineDeepLinkPopStateWired) {
      window.addEventListener('popstate', this.onInlineDeepLinkPopState);
      this.isInlineDeepLinkPopStateWired = true;
      return;
    }

    if (!enabled && this.isInlineDeepLinkPopStateWired) {
      window.removeEventListener('popstate', this.onInlineDeepLinkPopState);
      this.isInlineDeepLinkPopStateWired = false;
    }
  }
  private configureInlineDeepLinkScrollRestoration(enabled: boolean): void {
    if (typeof window === 'undefined' || !window.history) {
      return;
    }

    const historyObject = window.history as History & {
      scrollRestoration?: 'auto' | 'manual';
    };

    if (enabled) {
      if (this.isInlineDeepLinkScrollRestorationManaged) {
        return;
      }

      const currentScrollRestoration =
        typeof historyObject.scrollRestoration === 'string'
          ? historyObject.scrollRestoration
          : undefined;
      this.previousInlineDeepLinkScrollRestoration = currentScrollRestoration;
      try {
        historyObject.scrollRestoration = 'manual';
        this.isInlineDeepLinkScrollRestorationManaged = true;
      } catch {
        this.previousInlineDeepLinkScrollRestoration = undefined;
      }
      return;
    }

    if (!this.isInlineDeepLinkScrollRestorationManaged) {
      return;
    }

    try {
      if (this.previousInlineDeepLinkScrollRestoration) {
        historyObject.scrollRestoration = this.previousInlineDeepLinkScrollRestoration;
      }
    } catch {
      // Ignore restoration failures in unsupported browser contexts.
    }
    this.isInlineDeepLinkScrollRestorationManaged = false;
    this.previousInlineDeepLinkScrollRestoration = undefined;
  }
  private clearNestedIframeHydration(): void {
    if (this.nestedIframeHydrationCleanup) {
      this.nestedIframeHydrationCleanup();
      this.nestedIframeHydrationCleanup = undefined;
    }
  }
  private applyInitialDeepLinkScrollLock(): void {
    if (typeof window === 'undefined') {
      return;
    }

    this.clearInitialDeepLinkScrollLock();

    const historyObject = window.history as History & {
      scrollRestoration?: 'auto' | 'manual';
    };
    const previousScrollRestoration =
      typeof historyObject.scrollRestoration === 'string'
        ? historyObject.scrollRestoration
        : undefined;
    try {
      if (previousScrollRestoration) {
        historyObject.scrollRestoration = 'manual';
      }
    } catch {
      // Continue even if scroll restoration cannot be toggled in this browser context.
    }

    let released = false;
    let intervalId = 0;
    let releaseTimerId = 0;
    let lastAppliedAt = 0;
    let lastTraceAt = 0;
    const lockStartedAt = Date.now();
    let lastDeviationAt = lockStartedAt;
    let lastInlineFrameMetrics: IInlineDeepLinkFrameMetrics | undefined =
      this.getInlineDeepLinkFrameMetrics();
    const hostScrollContainers = this.getPotentialHostScrollContainers();
    const scrollTraceEnabled = this.isScrollTraceEnabled();
    let release: () => void = (): void => {
      return;
    };
    const minLockDurationMs = 500;
    const stableReleaseDurationMs = 900;
    const maxLockDurationMs = 12000;
    const frameHeightShiftThresholdPx = 6;
    const documentHeightShiftThresholdPx = 12;
    const timeouts: number[] = [];
    const interactionEvents: Array<keyof WindowEventMap> = [
      'wheel',
      'touchstart',
      'mousedown',
      'pointerdown',
    ];
    const trace = (
      eventName: string,
      data?: Record<string, unknown>,
      force: boolean = false,
    ): void => {
      if (!scrollTraceEnabled) {
        return;
      }

      const now = Date.now();
      if (!force && now - lastTraceAt < 300) {
        return;
      }

      lastTraceAt = now;
      this.emitScrollTrace(eventName, data);
    };
    trace(
      'scroll-lock-start',
      {
        previousScrollRestoration: previousScrollRestoration || '(none)',
        hostContainers: hostScrollContainers.map((container) =>
          this.describeScrollElement(container),
        ),
      },
      true,
    );

    const restoreTop = (): void => {
      if (released) {
        return;
      }

      const now = Date.now();
      if (now - lastAppliedAt < 40) {
        return;
      }
      lastAppliedAt = now;
      this.forceHostScrollTop();
      this.restoreHostScrollPosition({ x: 0, y: 0 });
      this.resetInlineIframeScrollPositionForDeepLink();
      const currentHostScrollContainers = this.getPotentialHostScrollContainers();
      const scrollOffsets = this.getDeepLinkScrollOffsets(currentHostScrollContainers);
      const inlineFrameMetrics = this.getInlineDeepLinkFrameMetrics();
      const hasPendingNestedFrames = (inlineFrameMetrics?.pendingNestedFrames || 0) > 0;
      const hasFrameHeightShift =
        !!inlineFrameMetrics &&
        !!lastInlineFrameMetrics &&
        Math.abs(inlineFrameMetrics.frameClientHeight - lastInlineFrameMetrics.frameClientHeight) >=
          frameHeightShiftThresholdPx;
      const hasDocumentHeightShift =
        !!inlineFrameMetrics &&
        !!lastInlineFrameMetrics &&
        Math.abs(inlineFrameMetrics.documentHeight - lastInlineFrameMetrics.documentHeight) >=
          documentHeightShiftThresholdPx;
      const hasLayoutShift = hasFrameHeightShift || hasDocumentHeightShift;

      if (hasLayoutShift) {
        trace('iframe-layout-shift-detected', {
          frameClientHeight: inlineFrameMetrics?.frameClientHeight || 0,
          previousFrameClientHeight: lastInlineFrameMetrics?.frameClientHeight || 0,
          documentHeight: inlineFrameMetrics?.documentHeight || 0,
          previousDocumentHeight: lastInlineFrameMetrics?.documentHeight || 0,
        });
      }
      if (hasPendingNestedFrames) {
        trace('nested-frame-processing', {
          pendingNestedFrames: inlineFrameMetrics?.pendingNestedFrames || 0,
        });
      }
      if (inlineFrameMetrics) {
        lastInlineFrameMetrics = inlineFrameMetrics;
      }

      if (scrollOffsets.maxOffset > 2 || hasLayoutShift || hasPendingNestedFrames) {
        lastDeviationAt = now;
      } else if (
        now - lockStartedAt >= minLockDurationMs &&
        now - lastDeviationAt >= stableReleaseDurationMs
      ) {
        trace(
          'auto-release-stable',
          {
            stableForMs: now - lastDeviationAt,
            lockDurationMs: now - lockStartedAt,
            offsets: scrollOffsets,
          },
          true,
        );
        release();
        return;
      }
      trace('restore-top');
    };
    const onUserInteraction = (event: Event): void => {
      if (!event.isTrusted) {
        trace('user-interaction-ignored-not-trusted', { type: event.type });
        return;
      }
      if (Date.now() - lockStartedAt < 250) {
        trace('user-interaction-ignored-early', { type: event.type });
        return;
      }
      trace('user-interaction-release', { type: event.type }, true);
      release();
    };
    const onWindowScroll = (): void => {
      if (released) {
        return;
      }
      const windowScrollTop =
        window.scrollY ||
        document.documentElement?.scrollTop ||
        document.body?.scrollTop ||
        0;
      if (windowScrollTop > 2) {
        lastDeviationAt = Date.now();
        trace('window-scroll-detected', { windowScrollTop });
        restoreTop();
      }
    };
    const onHostScroll = (): void => {
      if (released) {
        return;
      }
      const hasOffsetContainer = hostScrollContainers.some(
        (hostScrollContainer) => hostScrollContainer.scrollTop > 2,
      );
      if (hasOffsetContainer) {
        lastDeviationAt = Date.now();
        trace(
          'host-scroll-detected',
          {
            containerOffsets: hostScrollContainers.map((hostScrollContainer) => ({
              element: this.describeScrollElement(hostScrollContainer),
              top: hostScrollContainer.scrollTop || 0,
              left: hostScrollContainer.scrollLeft || 0,
            })),
          },
          true,
        );
        restoreTop();
      }
    };

    release = (): void => {
      if (released) {
        return;
      }

      trace('scroll-lock-release', undefined, true);
      this.forceHostScrollTop();
      this.resetInlineIframeScrollPositionForDeepLink();
      released = true;
      window.clearInterval(intervalId);
      window.clearTimeout(releaseTimerId);
      timeouts.forEach((timeoutId) => {
        window.clearTimeout(timeoutId);
      });
      interactionEvents.forEach((eventName) => {
        window.removeEventListener(eventName, onUserInteraction);
      });
      window.removeEventListener('scroll', onWindowScroll, true);
      hostScrollContainers.forEach((hostScrollContainer) => {
        hostScrollContainer.removeEventListener('scroll', onHostScroll, true);
      });

      if (previousScrollRestoration) {
        try {
          historyObject.scrollRestoration = previousScrollRestoration;
        } catch {
          return;
        }
      }
      trace('scroll-lock-released', undefined, true);
    };

    interactionEvents.forEach((eventName) => {
      window.addEventListener(eventName, onUserInteraction);
    });
    window.addEventListener('scroll', onWindowScroll, true);
    hostScrollContainers.forEach((hostScrollContainer) => {
      hostScrollContainer.addEventListener('scroll', onHostScroll, true);
    });

    timeouts.push(
      window.setTimeout(restoreTop, 0),
      window.setTimeout(restoreTop, 120),
      window.setTimeout(restoreTop, 350),
      window.setTimeout(restoreTop, 800),
      window.setTimeout(restoreTop, 1400),
      window.setTimeout(restoreTop, 2200),
      window.setTimeout(restoreTop, 3200),
      window.setTimeout(restoreTop, 4500),
      window.setTimeout(restoreTop, 6200),
      window.setTimeout(restoreTop, 8200),
    );
    intervalId = window.setInterval(restoreTop, 80);
    releaseTimerId = window.setTimeout(() => {
      trace(
        'auto-release-timeout',
        {
          lockDurationMs: Date.now() - lockStartedAt,
        },
        true,
      );
      release();
    }, maxLockDurationMs);
    this.initialDeepLinkScrollLockCleanup = release;
  }
  private getDeepLinkScrollOffsets(
    hostScrollContainers: HTMLElement[],
  ): { windowTop: number; hostMaxTop: number; iframeTop: number; maxOffset: number } {
    const windowTop =
      window.scrollY || document.documentElement?.scrollTop || document.body?.scrollTop || 0;
    let hostMaxTop = 0;
    hostScrollContainers.forEach((container) => {
      if (container.scrollTop > hostMaxTop) {
        hostMaxTop = container.scrollTop;
      }
    });

    const iframeTop = this.getInlineIframeMaxScrollTop();
    const maxOffset = Math.max(windowTop, hostMaxTop, iframeTop);
    return {
      windowTop,
      hostMaxTop,
      iframeTop,
      maxOffset,
    };
  }
  private getInlineIframeMaxScrollTop(): number {
    const iframe = this.domElement.querySelector('iframe');
    if (!iframe) {
      return 0;
    }
    return this.getIframeDeepMaxScrollTop(iframe);
  }
  private getInlineDeepLinkFrameMetrics(): IInlineDeepLinkFrameMetrics | undefined {
    const iframe = this.domElement.querySelector('iframe');
    if (!iframe) {
      return undefined;
    }

    const metrics: IInlineDeepLinkFrameMetrics = {
      frameClientHeight: iframe.clientHeight || 0,
      frameScrollHeight: iframe.scrollHeight || 0,
      documentHeight: 0,
      pendingNestedFrames: 0,
    };

    try {
      const iframeDocument = iframe.contentDocument;
      if (!iframeDocument) {
        return metrics;
      }

      const root = iframeDocument.documentElement;
      const body = iframeDocument.body;
      metrics.documentHeight = Math.max(
        root?.scrollHeight || 0,
        body?.scrollHeight || 0,
        root?.offsetHeight || 0,
        body?.offsetHeight || 0,
        root?.clientHeight || 0,
        body?.clientHeight || 0,
      );
      metrics.pendingNestedFrames = iframeDocument.querySelectorAll(
        'iframe[data-uhv-nested-state="processing"]',
      ).length;
    } catch {
      // Ignore cross-origin iframe access issues.
    }

    return metrics;
  }
  private resetInlineIframeScrollPositionForDeepLink(): void {
    const iframe = this.domElement.querySelector('iframe');
    if (!iframe) {
      return;
    }

    this.resetIframeScrollPosition(iframe);
  }
  private clearInitialDeepLinkScrollLock(): void {
    if (!this.initialDeepLinkScrollLockCleanup) {
      return;
    }

    this.initialDeepLinkScrollLockCleanup();
    this.initialDeepLinkScrollLockCleanup = undefined;
  }
  protected onDispose(): void {
    this.configureInlineDeepLinkPopState(false);
    this.clearInitialDeepLinkScrollLock();
    this.clearNestedIframeHydration();
    super.onDispose();
  }
}
