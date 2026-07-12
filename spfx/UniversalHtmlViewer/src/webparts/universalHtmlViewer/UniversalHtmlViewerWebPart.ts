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
import { NestedIframeHydrationDiagnosticEvent } from './NestedIframeHydrationHelper';
import {
  ILoadSharePointInlineContentOptions,
  loadSharePointFileContentForBlobUrl,
  loadSharePointFileContentForInline,
} from './SharePointInlineContentHelper';
import { extractTitleFromHtml } from './PageTitleHelper';
import {
  buildPageUrlWithoutInlineDeepLink,
  buildPageUrlWithInlineDeepLink,
  DEFAULT_INLINE_DEEP_LINK_PARAM,
  IResolvedInlineContentTarget,
  resolveInlineContentTarget,
} from './InlineDeepLinkHelper';
import {
  createDefaultDeepLinkScrollLockDiagnosticsCounters,
  type DeepLinkScrollLockReleaseReason,
} from './UniversalHtmlViewerWebPartRuntimeBase';
import {
  ConfigurationPreset,
  ContentDeliveryMode,
  IUniversalHtmlViewerWebPartProps,
  isInlineContentDeliveryMode,
  isReportBrowserSourceMode,
} from './UniversalHtmlViewerTypes';
import {
  acquireManualHistoryScrollRestoration,
  releaseManualHistoryScrollRestoration,
} from './HistoryScrollRestorationHelper';
import { INLINE_HOST_PAGE_URL_CHANGED_EVENT } from './InlineHostPageUrlSyncHelper';

interface IInlineDeepLinkFrameMetrics {
  frameClientHeight: number;
  frameScrollHeight: number;
  documentHeight: number;
  pendingNestedFrames: number;
}

interface IPageTitleSyncEntry {
  ownerId: string;
  syncedTitle: string;
}

interface IPageTitleSyncState {
  originalTitle: string;
  entries: IPageTitleSyncEntry[];
}

interface IWindowWithPageTitleSync extends Window {
  __uhvPageTitleSync?: IPageTitleSyncState;
}

let nextPageTitleSyncOwnerId = 1;

export default class UniversalHtmlViewerWebPart extends UniversalHtmlViewerWebPartUiBase {
  private nestedIframeHydrationCleanup: (() => void) | undefined;
  private initialDeepLinkScrollLockCleanup:
    | ((reason?: DeepLinkScrollLockReleaseReason) => void)
    | undefined;
  private isInlineDeepLinkPopStateWired: boolean = false;
  private readonly historyScrollRestorationOwnerId: string = `uhv-history-${nextPageTitleSyncOwnerId++}`;
  private hasLoggedAnyHttpsWarning: boolean = false;
  private activeInlineBlobUrl: string | undefined;
  private readonly pageTitleSyncOwnerId: string = `uhv-title-${nextPageTitleSyncOwnerId++}`;
  private renderRequestId: number = 0;
  private renderDisposed: boolean = false;
  private readonly onInlineDeepLinkPopState = (): void => {
    this.render();
  };

  public render(): void {
    if (this.renderDisposed) {
      return;
    }
    const renderRequestId = ++this.renderRequestId;
    this.renderAsync(renderRequestId).catch((error) => {
      if (!this.isRenderRequestCurrent(renderRequestId)) {
        return;
      }
      this.clearRefreshTimer();
      this.clearIframeLoadTimeout();
      this.restoreOriginalDocumentTitle();
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
      propertyPath === 'showOpenInNewTab' &&
      newValue === true &&
      isInlineContentDeliveryMode(this.getContentDeliveryMode(this.properties))
    ) {
      this.properties.allowQueryStringPageOverride = true;
      this.context.propertyPane.refresh();
    }
    if (
      propertyPath === 'allowQueryStringPageOverride' &&
      newValue === false &&
      isInlineContentDeliveryMode(this.getContentDeliveryMode(this.properties))
    ) {
      this.properties.showOpenInNewTab = false;
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
      propertyPath === 'showDashboardSelector' ||
      propertyPath === 'showReportBrowser'
    ) {
      if (propertyPath === 'htmlSourceMode') {
        const nextSourceMode = (newValue || 'FullUrl') as HtmlSourceMode;
        const nextIsReportBrowserMode = isReportBrowserSourceMode(nextSourceMode);
        this.properties.showReportBrowser = nextIsReportBrowserMode;
        if (nextIsReportBrowserMode) {
          if (!isInlineContentDeliveryMode(this.properties.contentDeliveryMode)) {
            this.properties.contentDeliveryMode = 'SharePointFileContent';
          }
          if (!this.properties.reportBrowserRootPath && this.properties.basePath) {
            this.properties.reportBrowserRootPath = this.properties.basePath;
          }
        }
      }
      if (
        propertyPath === 'contentDeliveryMode' &&
        isReportBrowserSourceMode(this.properties.htmlSourceMode) &&
        !isInlineContentDeliveryMode(this.properties.contentDeliveryMode)
      ) {
        this.properties.contentDeliveryMode = 'SharePointFileContent';
      }
      this.context.propertyPane.refresh();
    }

    this.render();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const htmlSourceMode: HtmlSourceMode = this.properties.htmlSourceMode || 'FullUrl';
    const isFullUrl: boolean = htmlSourceMode === 'FullUrl';
    const isReportBrowserMode: boolean = isReportBrowserSourceMode(htmlSourceMode);
    const isRelativePath: boolean = htmlSourceMode === 'BasePathAndRelativePath';
    const isDashboardId: boolean = htmlSourceMode === 'BasePathAndDashboardId';
    const securityMode: UrlSecurityMode = this.properties.securityMode || 'StrictTenant';
    const enableExpertSecurityModes: boolean = this.properties.enableExpertSecurityModes === true;
    const currentContentDeliveryMode: ContentDeliveryMode = this.getContentDeliveryMode(
      this.properties,
    );
    const securityModeOptions = [
      { key: 'StrictTenant', text: 'Strict tenant (default)' },
      { key: 'Allowlist', text: 'Tenant + allowlist' },
      ...(enableExpertSecurityModes || securityMode === 'AnyHttps'
        ? [{ key: 'AnyHttps', text: 'Any HTTPS (unsafe expert mode)' }]
        : []),
    ];
    const isAllowlistMode: boolean = securityMode === 'Allowlist';
    const heightMode: HeightMode = this.properties.heightMode || 'Fixed';
    const isInlineContentMode: boolean = isInlineContentDeliveryMode(
      currentContentDeliveryMode,
    );
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
    const canUseReportBrowser: boolean = showChrome && isInlineContentMode;
    const showLegacyDirectUrlOption: boolean =
      enableExpertSecurityModes ||
      currentContentDeliveryMode === 'DirectUrl' ||
      preset === 'AllowlistCDN' ||
      preset === 'AnyHttps' ||
      securityMode === 'AnyHttps';
    const contentDeliveryModeOptions = [
      {
        key: 'SharePointFileContent',
        text: 'SharePoint file API (inline iframe)',
      },
      {
        key: 'SharePointFileBlobUrl',
        text: 'SharePoint file API (blob iframe)',
      },
      ...(showLegacyDirectUrlOption
        ? [{ key: 'DirectUrl', text: 'Direct URL in iframe (legacy / external only)' }]
        : []),
    ];

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
              groupName: 'Initial content (Required)',
              groupFields: [
                PropertyPaneDropdown('htmlSourceMode', {
                  label: 'Initial HTML source mode',
                  options: [
                    { key: 'FullUrl', text: 'Single page URL' },
                    {
                      key: 'SharePointReportBrowser',
                      text: 'SharePoint report browser folder',
                    },
                    { key: 'BasePathAndRelativePath', text: 'Base path + relative path' },
                    { key: 'BasePathAndDashboardId', text: 'Base path + dashboard ID' },
                  ],
                }),
                PropertyPaneDropdown('contentDeliveryMode', {
                  label: 'Content delivery mode',
                  options: contentDeliveryModeOptions,
                }),
                PropertyPaneTextField('fullUrl', {
                  label: 'HTML page URL',
                  description: 'Used only when initial HTML source mode is "Single page URL".',
                  disabled: !isFullUrl,
                  onGetErrorMessage: (value?: string): string =>
                    validateFullUrl(value, !!this.properties.allowHttp),
                  deferredValidationTime: 200,
                }),
                PropertyPaneTextField('basePath', {
                  label: 'Base path (site-relative)',
                  description:
                    'Site-relative base path for relative path or dashboard ID modes. Example: /sites/Reports/Dashboards/',
                  disabled: isFullUrl || isReportBrowserMode,
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
                  label: isReportBrowserMode
                    ? 'Initial/default report file name'
                    : 'Default file name',
                  description:
                    'Used by dashboard ID and report browser modes. Defaults to "index.html" when left empty.',
                  disabled: !isDashboardId && !isReportBrowserMode,
                }),
                PropertyPaneTextField('queryStringParamName', {
                  label: 'Query string parameter name',
                  description:
                    'Used when initial HTML source mode is "Base path + dashboard ID" to read the dashboard ID from the page URL. Defaults to "dashboard" when left empty.',
                  disabled: !isDashboardId,
                }),
              ],
            },
            {
              groupName: 'Report browser source',
              groupFields: [
                PropertyPaneTextField('reportBrowserRootPath', {
                  label: 'Browser root folder',
                  description:
                    'Folder UHV enumerates for the report picker. Choose "SharePoint report browser folder" above to use this.',
                  disabled: !isReportBrowserMode || !canUseReportBrowser,
                }),
                PropertyPaneDropdown('reportBrowserDefaultView', {
                  label: 'Default browser view',
                  options: [
                    { key: 'Folders', text: 'Folders' },
                    { key: 'Files', text: 'Files (recursive)' },
                  ],
                  disabled: !isReportBrowserMode || !canUseReportBrowser,
                }),
                PropertyPaneSlider('reportBrowserMaxItems', {
                  label: 'Maximum browser items',
                  min: 25,
                  max: 1000,
                  step: 25,
                  disabled: !isReportBrowserMode || !canUseReportBrowser,
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
                PropertyPaneToggle('syncPageTitle', {
                  label: 'Sync browser tab title from report',
                  onText: 'On',
                  offText: 'Off',
                  disabled: !isInlineContentMode || isPresetLocked,
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
                  label: 'Allow page query override (inline mode)',
                  onText: 'On',
                  offText: 'Off',
                  disabled: !isInlineContentMode || isPresetLocked,
                }),
                PropertyPaneTextField('inlineDeepLinkParamName', {
                  label: 'Page query parameter name',
                  description:
                    'Default: uhvPage. Use a unique name for each viewer when a page contains multiple viewer web parts.',
                  disabled:
                    !isInlineContentMode ||
                    this.properties.allowQueryStringPageOverride !== true ||
                    isPresetLocked,
                }),
                PropertyPaneToggle('enforceStrictInlineCsp', {
                  label: 'Enforce strict inline CSP (scripts)',
                  onText: 'On',
                  offText: 'Off',
                  disabled: !isInlineContentMode || isPresetLocked,
                }),
                PropertyPaneToggle('inlineExternalScripts', {
                  label: 'Inline external report scripts (compatibility)',
                  onText: 'On',
                  offText: 'Off',
                  disabled: !isInlineContentMode || isPresetLocked,
                }),
                PropertyPaneTextField('inlineExternalScriptAllowedHosts', {
                  label: 'External script hosts to inline',
                  description:
                    'Optional. Defaults to common PSWriteHTML CDNs: code.jquery.com, cdnjs.cloudflare.com, cdn.jsdelivr.net, cdn.datatables.net, nightly.datatables.net, unpkg.com',
                  disabled:
                    !isInlineContentMode ||
                    this.properties.inlineExternalScripts !== true ||
                    isPresetLocked,
                  onGetErrorMessage: (value?: string): string => validateAllowedHosts(value),
                  deferredValidationTime: 200,
                }),
                PropertyPaneTextField('inlineCspScriptAllowedHosts', {
                  label: 'Inline CSP script hosts',
                  description:
                    'Optional trusted hosts allowed by the generated inline CSP script-src directive. Example: cdn.jsdelivr.net, cdn.datatables.net, cdnjs.cloudflare.com',
                  disabled: !isInlineContentMode || isPresetLocked,
                  onGetErrorMessage: (value?: string): string => validateAllowedHosts(value),
                  deferredValidationTime: 200,
                }),
                PropertyPaneTextField('inlineCspStyleAllowedHosts', {
                  label: 'Inline CSP style hosts',
                  description:
                    'Optional trusted hosts allowed by the generated inline CSP style-src/font-src directives. Example: fonts.googleapis.com, cdn.datatables.net',
                  disabled: !isInlineContentMode || isPresetLocked,
                  onGetErrorMessage: (value?: string): string => validateAllowedHosts(value),
                  deferredValidationTime: 200,
                }),
                PropertyPaneTextField('inlineCspImageAllowedHosts', {
                  label: 'Inline CSP image hosts',
                  description:
                    'Optional trusted hosts allowed by the generated inline CSP img-src/media-src directives. Example: upload.wikimedia.org',
                  disabled: !isInlineContentMode || isPresetLocked,
                  onGetErrorMessage: (value?: string): string => validateAllowedHosts(value),
                  deferredValidationTime: 200,
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
                    {
                      key: 'Relaxed',
                      text: 'Relaxed (trusted content; same-origin access)',
                    },
                    { key: 'Strict', text: 'Strict (isolated origin)' },
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

  private async renderAsync(renderRequestId: number): Promise<void> {
    this.normalizeInlineDeepLinkConfiguration(this.properties);
    this.invalidateRefreshRequests();
    this.clearInitialDeepLinkScrollLock();
    this.resetDeepLinkScrollLockDiagnostics();
    this.clearIframeLoadTimeout();
    this.clearNestedIframeHydration();
    this.revokeActiveInlineBlobUrl();
    this.resetNestedIframeDiagnostics();
    const pageUrl: string = this.getCurrentPageUrl();
    const { effectiveProps, tenantConfig } = await this.getEffectiveProperties(pageUrl);
    if (!this.isRenderRequestCurrent(renderRequestId)) {
      return;
    }
    this.lastEffectiveProps = effectiveProps;
    this.lastTenantConfig = tenantConfig;
    const htmlSourceMode: HtmlSourceMode = effectiveProps.htmlSourceMode || 'FullUrl';
    const contentDeliveryMode: ContentDeliveryMode = this.getContentDeliveryMode(
      effectiveProps,
    );
    if (!this.shouldSyncPageTitle(effectiveProps, contentDeliveryMode)) {
      this.restoreOriginalDocumentTitle();
    }
    this.configureInlineDeepLinkPopState(
      isInlineContentDeliveryMode(contentDeliveryMode) &&
        this.shouldEnableInlineDeepLinks(effectiveProps),
    );
    const currentDashboardId: string | undefined = this.getEffectiveDashboardId(
      effectiveProps,
      pageUrl,
    );

    const finalUrl: string | undefined = buildFinalUrl({
      htmlSourceMode,
      fullUrl: effectiveProps.fullUrl,
      basePath: effectiveProps.basePath,
      reportBrowserRootPath: effectiveProps.reportBrowserRootPath,
      webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
      relativePath: effectiveProps.relativePath,
      dashboardId: effectiveProps.dashboardId,
      defaultFileName: effectiveProps.defaultFileName,
      queryStringParamName: effectiveProps.queryStringParamName,
      pageUrl,
    });

    if (!finalUrl) {
      this.clearRefreshTimer();
      this.clearIframeLoadTimeout();
      this.restoreOriginalDocumentTitle();
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
      queryParamName: this.getInlineDeepLinkParamName(effectiveProps),
      validationOptions,
      allowDeepLinkOverride: this.shouldAllowInlineDeepLinkOverride(effectiveProps),
    });
    const requestedDeepLinkValue: string = resolvedContentTarget.requestedDeepLinkValue;
    const hasRequestedDeepLink: boolean = resolvedContentTarget.hasRequestedDeepLink;
    const hasAppliedDeepLink: boolean = !!resolvedContentTarget.deepLinkedUrl;
    const shouldResetHostScrollToTopOnInitialDeepLink: boolean =
      this.shouldApplyInitialDeepLinkScrollLock(contentDeliveryMode, resolvedContentTarget);
    if (this.isScrollTraceEnabled()) {
      this.emitScrollTrace(
        'deep-link-evaluation',
        {
          contentDeliveryMode,
          hasRequestedDeepLink,
          hasAppliedDeepLink,
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
      this.clearInitialDeepLinkScrollLock();
      this.restoreOriginalDocumentTitle();
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
      this.clearInitialDeepLinkScrollLock();
      this.restoreOriginalDocumentTitle();
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
    const cacheBusterParamName: string = this.normalizeCacheBusterParamName(
      effectiveProps.cacheBusterParamName,
    );
    const resolvedUrl: string = await this.resolveUrlWithCacheBuster(
      initialContentUrl,
      cacheBusterMode,
      cacheBusterParamName,
      pageUrl,
    );
    if (!this.isRenderRequestCurrent(renderRequestId)) {
      return;
    }
    let inlineHtml: string | undefined;
    let iframeUrl: string = resolvedUrl;
    if (isInlineContentDeliveryMode(contentDeliveryMode)) {
      try {
        if (contentDeliveryMode === 'SharePointFileBlobUrl') {
          inlineHtml = await loadSharePointFileContentForBlobUrl(
            this.context.spHttpClient,
            this.context.pageContext.web.absoluteUrl,
            resolvedUrl,
            initialContentUrl,
            pageUrl,
            SPHttpClient.configurations.v1,
            this.getInlineContentOptions(
              effectiveProps,
              false,
              validationOptions.allowedPathPrefixes,
            ),
          );
          if (!this.isRenderRequestCurrent(renderRequestId)) {
            return;
          }
          iframeUrl = this.createInlineBlobUrl(inlineHtml);
        } else {
          inlineHtml = await loadSharePointFileContentForInline(
            this.context.spHttpClient,
            this.context.pageContext.web.absoluteUrl,
            resolvedUrl,
            initialContentUrl,
            pageUrl,
            SPHttpClient.configurations.v1,
            this.getInlineContentOptions(
              effectiveProps,
              false,
              validationOptions.allowedPathPrefixes,
            ),
          );
          if (!this.isRenderRequestCurrent(renderRequestId)) {
            return;
          }
        }
      } catch (error) {
        if (!this.isRenderRequestCurrent(renderRequestId)) {
          return;
        }
        this.clearRefreshTimer();
        this.clearIframeLoadTimeout();
        this.clearInitialDeepLinkScrollLock();
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
        this.restoreOriginalDocumentTitle();
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

    this.syncPageTitleFromHtml(inlineHtml, effectiveProps);
    this.currentBaseUrl = initialContentUrl;
    this.renderIframe(
      iframeUrl,
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
      contentDeliveryMode === 'SharePointFileContent' ? inlineHtml : undefined,
    );
    this.setupIframeLoadFallback(iframeUrl, effectiveProps);
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
      rewriteAnchorHrefs: this.shouldRewriteInlineAnchorHrefs(effectiveProps),
      deepLinkQueryParamName: this.getInlineDeepLinkParamName(effectiveProps),
      preservedHostQueryParamNames:
        this.getInlineAnchorPreservedHostQueryParamNames(effectiveProps),
      onNavigate: (targetUrl: string, inlineCacheBusterMode: CacheBusterMode) => {
        this.currentBaseUrl = targetUrl;
        this.setLoadingVisible(true);
        this.lastCacheBusterMode = inlineCacheBusterMode;
        const updatedPageUrl = this.getCurrentPageUrl();
        this.onNavigatedToUrl(targetUrl, updatedPageUrl);
        const navigatedPageUrl = this.getCurrentPageUrl();
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
          true,
        ).catch(() => {
          return undefined;
        });
      },
      loadInlineHtml: async (
        sourceUrl: string,
        baseUrlForRelativeLinks: string,
      ): Promise<string | undefined> => {
        try {
          const inlineHtml = await loadSharePointFileContentForInline(
            this.context.spHttpClient,
            this.context.pageContext.web.absoluteUrl,
            sourceUrl,
            baseUrlForRelativeLinks,
            this.getCurrentPageUrl(),
            SPHttpClient.configurations.v1,
            this.getInlineContentOptions(
              this.lastEffectiveProps || this.properties,
              false,
              this.lastValidationOptions?.allowedPathPrefixes,
            ),
          );
          this.lastInlineContentLoadError = '';
          return inlineHtml;
        } catch (error) {
          this.lastInlineContentLoadError = this.formatInlineContentLoadError(error);
          return undefined;
        }
      },
      onNestedDiagnosticsEvent: (eventName: NestedIframeHydrationDiagnosticEvent): void => {
        this.recordNestedIframeDiagnosticEvent(eventName);
      },
    });
  }
  protected async trySetIframeSrcDocFromSource(
    iframe: HTMLIFrameElement,
    sourceUrl: string,
    pageUrl: string,
    props: IUniversalHtmlViewerWebPartProps,
    bypassInlineContentCache: boolean = false,
    refreshRequestId?: number,
  ): Promise<boolean> {
    const contentDeliveryMode: ContentDeliveryMode = this.getContentDeliveryMode(props);
    if (!isInlineContentDeliveryMode(contentDeliveryMode)) {
      return false;
    }
    try {
      const baseUrlForRelativeLinks: string = this.currentBaseUrl || sourceUrl;
      if (contentDeliveryMode === 'SharePointFileBlobUrl') {
        const blobHtml = await loadSharePointFileContentForBlobUrl(
          this.context.spHttpClient,
          this.context.pageContext.web.absoluteUrl,
          sourceUrl,
          baseUrlForRelativeLinks,
          pageUrl,
          SPHttpClient.configurations.v1,
          this.getInlineContentOptions(
            props,
            bypassInlineContentCache,
            this.lastValidationOptions?.allowedPathPrefixes,
          ),
        );
        if (
          refreshRequestId !== undefined &&
          !this.isRefreshRequestCurrent(refreshRequestId)
        ) {
          return false;
        }
        this.lastInlineContentLoadError = '';
        this.syncPageTitleFromHtml(blobHtml, props);
        iframe.removeAttribute('srcdoc');
        this.replaceInlineBlobFrameLocation(
          iframe,
          this.createInlineBlobUrl(blobHtml),
        );
        return true;
      }

      const inlineHtml = await loadSharePointFileContentForInline(
        this.context.spHttpClient,
        this.context.pageContext.web.absoluteUrl,
        sourceUrl,
        baseUrlForRelativeLinks,
        pageUrl,
        SPHttpClient.configurations.v1,
        this.getInlineContentOptions(
          props,
          bypassInlineContentCache,
          this.lastValidationOptions?.allowedPathPrefixes,
        ),
      );
      if (
        refreshRequestId !== undefined &&
        !this.isRefreshRequestCurrent(refreshRequestId)
      ) {
        return false;
      }
      this.lastInlineContentLoadError = '';
      this.syncPageTitleFromHtml(inlineHtml, props);
      iframe.srcdoc = inlineHtml;
      return true;
    } catch (error) {
      if (
        refreshRequestId === undefined ||
        this.isRefreshRequestCurrent(refreshRequestId)
      ) {
        this.lastInlineContentLoadError = this.formatInlineContentLoadError(error);
      }
      return false;
    }
  }

  private getInlineContentOptions(
    props: IUniversalHtmlViewerWebPartProps,
    bypassCache: boolean = false,
    effectiveAllowedPathPrefixes?: string[],
  ): ILoadSharePointInlineContentOptions {
    return {
      cacheTtlMs: this.getInlineContentCacheTtlMs(props),
      bypassCache,
      enforceStrictInlineCsp: props.enforceStrictInlineCsp === true,
      inlineExternalScripts: props.inlineExternalScripts === true,
      inlineExternalScriptAllowedHosts: this.parseHosts(
        props.inlineExternalScriptAllowedHosts,
      ),
      inlineCspScriptAllowedHosts: this.parseHosts(props.inlineCspScriptAllowedHosts),
      inlineCspStyleAllowedHosts: this.parseHosts(props.inlineCspStyleAllowedHosts),
      inlineCspImageAllowedHosts: this.parseHosts(props.inlineCspImageAllowedHosts),
      rewriteInlineAnchorHrefs: this.shouldRewriteInlineAnchorHrefs(props),
      rewriteInlineAnchorAllowedFileExtensions: this.parseFileExtensions(
        props.allowedFileExtensions,
      ),
      rewriteInlineAnchorAllowedPathPrefixes:
        effectiveAllowedPathPrefixes || this.parsePathPrefixes(props.allowedPathPrefixes),
      rewriteInlineAnchorDeepLinkQueryParamName: this.getInlineDeepLinkParamName(props),
      rewriteInlineAnchorPreservedHostQueryParamNames:
        this.getInlineAnchorPreservedHostQueryParamNames(props),
    };
  }
  private getInlineAnchorPreservedHostQueryParamNames(
    props: IUniversalHtmlViewerWebPartProps,
  ): string[] {
    if ((props.htmlSourceMode || 'FullUrl') !== 'BasePathAndDashboardId') {
      return [];
    }

    return [(props.queryStringParamName || '').trim() || 'dashboard'];
  }
  private shouldRewriteInlineAnchorHrefs(props: IUniversalHtmlViewerWebPartProps): boolean {
    return this.shouldEnableInlineDeepLinks(props);
  }
  private shouldEnableInlineDeepLinks(
    props: IUniversalHtmlViewerWebPartProps,
  ): boolean {
    return (
      props.allowQueryStringPageOverride === true &&
      !(props.enableExpertSecurityModes === true && props.securityMode === 'AnyHttps')
    );
  }
  private shouldAllowInlineDeepLinkOverride(
    props: IUniversalHtmlViewerWebPartProps,
  ): boolean {
    return (
      props.allowQueryStringPageOverride === true &&
      !(props.enableExpertSecurityModes === true && props.securityMode === 'AnyHttps')
    );
  }
  private shouldSyncPageTitle(
    props: IUniversalHtmlViewerWebPartProps,
    contentDeliveryMode: ContentDeliveryMode,
  ): boolean {
    return props.syncPageTitle === true && isInlineContentDeliveryMode(contentDeliveryMode);
  }

  private createInlineBlobUrl(html: string): string {
    const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
    const nextBlobUrl = URL.createObjectURL(blob);
    const previousBlobUrl = this.activeInlineBlobUrl;
    this.activeInlineBlobUrl = nextBlobUrl;
    if (previousBlobUrl && previousBlobUrl !== nextBlobUrl) {
      URL.revokeObjectURL(previousBlobUrl);
    }
    return nextBlobUrl;
  }
  private replaceInlineBlobFrameLocation(
    iframe: HTMLIFrameElement,
    blobUrl: string,
  ): void {
    try {
      const iframeWindow = iframe.contentWindow;
      if (iframeWindow && typeof iframeWindow.location?.replace === 'function') {
        iframeWindow.location.replace(blobUrl);
        return;
      }
    } catch {
      // Fall back to src when the active sandbox does not expose frame navigation.
    }
    iframe.src = blobUrl;
  }
  private syncPageTitleFromHtml(
    html: string | undefined,
    props: IUniversalHtmlViewerWebPartProps,
  ): void {
    if (
      props.syncPageTitle !== true ||
      typeof window === 'undefined' ||
      typeof document === 'undefined'
    ) {
      return;
    }

    const reportTitle = extractTitleFromHtml(html);
    if (!reportTitle) {
      this.restoreOriginalDocumentTitle();
      return;
    }

    const syncWindow = window as IWindowWithPageTitleSync;
    const existingSyncState = syncWindow.__uhvPageTitleSync;
    const existingEntries = existingSyncState?.entries || [];
    const originalTitle = existingSyncState?.originalTitle ?? document.title ?? '';
    const nextEntries = existingEntries.filter(
      (entry) => entry.ownerId !== this.pageTitleSyncOwnerId,
    );
    nextEntries.push({
      ownerId: this.pageTitleSyncOwnerId,
      syncedTitle: reportTitle,
    });
    syncWindow.__uhvPageTitleSync = {
      originalTitle,
      entries: nextEntries,
    };
    document.title = reportTitle;
  }
  private restoreOriginalDocumentTitle(): void {
    if (typeof window === 'undefined' || typeof document === 'undefined') {
      return;
    }

    const syncWindow = window as IWindowWithPageTitleSync;
    const syncState = syncWindow.__uhvPageTitleSync;
    if (!syncState) {
      return;
    }

    const currentEntry = syncState.entries
      .slice()
      .reverse()
      .find((entry) => entry.ownerId === this.pageTitleSyncOwnerId);
    if (!currentEntry) {
      return;
    }

    const nextEntries = syncState.entries.filter(
      (entry) => entry.ownerId !== this.pageTitleSyncOwnerId,
    );
    const nextActiveEntry =
      nextEntries.length > 0 ? nextEntries[nextEntries.length - 1] : undefined;
    if (document.title === currentEntry.syncedTitle) {
      document.title = nextActiveEntry?.syncedTitle || syncState.originalTitle;
    }

    if (nextEntries.length > 0) {
      syncWindow.__uhvPageTitleSync = {
        originalTitle: syncState.originalTitle,
        entries: nextEntries,
      };
      return;
    }
    delete syncWindow.__uhvPageTitleSync;
  }
  private revokeActiveInlineBlobUrl(): void {
    if (!this.activeInlineBlobUrl) {
      return;
    }

    URL.revokeObjectURL(this.activeInlineBlobUrl);
    this.activeInlineBlobUrl = undefined;
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
    if (!isInlineContentDeliveryMode(this.getContentDeliveryMode(effectiveProps))) {
      return;
    }
    if (
      !this.shouldEnableInlineDeepLinks(effectiveProps)
    ) {
      return;
    }

    const currentPageUrl: string = (pageUrl || '').trim() || this.getCurrentPageUrl();
    const updatedPageUrl = buildPageUrlWithInlineDeepLink({
      pageUrl: currentPageUrl,
      targetUrl,
      queryParamName: this.getInlineDeepLinkParamName(effectiveProps),
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
      window.dispatchEvent(new Event(INLINE_HOST_PAGE_URL_CHANGED_EVENT));
    } catch {
      return;
    }
  }
  private buildResetToDefaultDashboardHtml(pageUrl: string): string {
    const effectiveProps = this.lastEffectiveProps || this.properties;
    const resetUrl = buildPageUrlWithoutInlineDeepLink({
      pageUrl,
      queryParamName: this.getInlineDeepLinkParamName(effectiveProps),
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
  private formatInlineContentLoadError(error: unknown): string {
    if (error instanceof Error && error.message) {
      return error.message;
    }

    const source = error as
      | {
          message?: unknown;
        }
      | undefined;
    const directMessage = typeof source?.message === 'string' ? source.message.trim() : '';
    if (directMessage) {
      return directMessage;
    }

    const statusCode = this.getInlineLoadErrorStatusCode(error);
    if (statusCode !== undefined) {
      return `Inline content load failed with status ${statusCode}.`;
    }

    const fallback = String(error || '').trim();
    return fallback || 'Inline content load failed.';
  }
  private resetNestedIframeDiagnostics(): void {
    this.nestedIframeDiagnostics = {
      hydrationStarted: 0,
      hydrationSucceeded: 0,
      hydrationFailed: 0,
      hydrationStaleResultIgnored: 0,
      navigationStarted: 0,
      navigationSucceeded: 0,
      navigationFailed: 0,
      navigationStaleResultIgnored: 0,
    };
  }
  private resetDeepLinkScrollLockDiagnostics(): void {
    this.deepLinkScrollLockDiagnostics = createDefaultDeepLinkScrollLockDiagnosticsCounters();
  }
  private recordNestedIframeDiagnosticEvent(
    eventName: NestedIframeHydrationDiagnosticEvent,
  ): void {
    switch (eventName) {
      case 'nestedHydrationStarted':
        this.nestedIframeDiagnostics.hydrationStarted += 1;
        break;
      case 'nestedHydrationSucceeded':
        this.nestedIframeDiagnostics.hydrationSucceeded += 1;
        break;
      case 'nestedHydrationFailed':
        this.nestedIframeDiagnostics.hydrationFailed += 1;
        break;
      case 'nestedHydrationStaleResultIgnored':
        this.nestedIframeDiagnostics.hydrationStaleResultIgnored += 1;
        break;
      case 'nestedNavigationStarted':
        this.nestedIframeDiagnostics.navigationStarted += 1;
        break;
      case 'nestedNavigationSucceeded':
        this.nestedIframeDiagnostics.navigationSucceeded += 1;
        break;
      case 'nestedNavigationFailed':
        this.nestedIframeDiagnostics.navigationFailed += 1;
        break;
      case 'nestedNavigationStaleResultIgnored':
        this.nestedIframeDiagnostics.navigationStaleResultIgnored += 1;
        break;
      default:
        break;
    }
  }
  private shouldApplyInitialDeepLinkScrollLock(
    contentDeliveryMode: ContentDeliveryMode,
    resolvedContentTarget: IResolvedInlineContentTarget,
  ): boolean {
    return isInlineContentDeliveryMode(contentDeliveryMode) && !!resolvedContentTarget.deepLinkedUrl;
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
    if (enabled) {
      acquireManualHistoryScrollRestoration(window, this.historyScrollRestorationOwnerId);
      return;
    }
    releaseManualHistoryScrollRestoration(window, this.historyScrollRestorationOwnerId);
  }
  private clearNestedIframeHydration(): void {
    if (this.nestedIframeHydrationCleanup) {
      this.nestedIframeHydrationCleanup();
      this.nestedIframeHydrationCleanup = undefined;
    }
  }
  private getInlineDeepLinkParamName(props: IUniversalHtmlViewerWebPartProps): string {
    return (props.inlineDeepLinkParamName || '').trim() || DEFAULT_INLINE_DEEP_LINK_PARAM;
  }
  private applyInitialDeepLinkScrollLock(): void {
    if (typeof window === 'undefined') {
      return;
    }

    this.clearInitialDeepLinkScrollLock('replace');
    this.deepLinkScrollLockDiagnostics.starts += 1;
    this.deepLinkScrollLockDiagnostics.active = true;

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
    let release: (reason?: DeepLinkScrollLockReleaseReason) => void = (): void => {
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
        release('auto-stable');
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
      release('user-interaction');
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

    release = (reason: DeepLinkScrollLockReleaseReason = 'manual'): void => {
      if (released) {
        return;
      }

      trace('scroll-lock-release', undefined, true);
      this.deepLinkScrollLockDiagnostics.releases += 1;
      this.deepLinkScrollLockDiagnostics.active = false;
      this.deepLinkScrollLockDiagnostics.lastReleaseReason = reason;
      this.deepLinkScrollLockDiagnostics.lastLockDurationMs = Date.now() - lockStartedAt;
      switch (reason) {
        case 'auto-stable':
          this.deepLinkScrollLockDiagnostics.releasedByAutoStable += 1;
          break;
        case 'user-interaction':
          this.deepLinkScrollLockDiagnostics.releasedByUserInteraction += 1;
          break;
        case 'timeout':
          this.deepLinkScrollLockDiagnostics.releasedByTimeout += 1;
          break;
        case 'replace':
          this.deepLinkScrollLockDiagnostics.releasedByReplace += 1;
          break;
        case 'dispose':
          this.deepLinkScrollLockDiagnostics.releasedByDispose += 1;
          break;
        default:
          this.deepLinkScrollLockDiagnostics.releasedByManual += 1;
          break;
      }
      this.forceHostScrollTop();
      this.resetInlineIframeScrollPositionForDeepLink();
      released = true;
      this.initialDeepLinkScrollLockCleanup = undefined;
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
      release('timeout');
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
  private clearInitialDeepLinkScrollLock(
    reason: DeepLinkScrollLockReleaseReason = 'manual',
  ): void {
    if (!this.initialDeepLinkScrollLockCleanup) {
      return;
    }

    this.initialDeepLinkScrollLockCleanup(reason);
    this.initialDeepLinkScrollLockCleanup = undefined;
  }
  protected onDispose(): void {
    this.renderDisposed = true;
    this.renderRequestId += 1;
    this.configureInlineDeepLinkPopState(false);
    this.clearInitialDeepLinkScrollLock('dispose');
    this.clearNestedIframeHydration();
    this.revokeActiveInlineBlobUrl();
    this.restoreOriginalDocumentTitle();
    super.onDispose();
  }

  private isRenderRequestCurrent(renderRequestId: number): boolean {
    return !this.renderDisposed && renderRequestId === this.renderRequestId;
  }
}
