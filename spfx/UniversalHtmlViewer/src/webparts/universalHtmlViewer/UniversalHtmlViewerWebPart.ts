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
import {
  buildFinalUrl,
  isUrlAllowed,
  HtmlSourceMode,
  HeightMode,
  UrlSecurityMode,
  CacheBusterMode,
  UrlValidationOptions,
} from './UrlHelper';

export interface IUniversalHtmlViewerWebPartProps {
  configurationPreset?: string;
  htmlSourceMode: HtmlSourceMode;
  fullUrl?: string;
  basePath?: string;
  relativePath?: string;
  dashboardId?: string;
  defaultFileName?: string;
  queryStringParamName?: string;
  heightMode: HeightMode;
  fixedHeightPx: number;
  securityMode?: UrlSecurityMode;
  allowHttp?: boolean;
  allowedHosts?: string;
  allowedPathPrefixes?: string;
  allowedFileExtensions?: string;
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
  showLoadingIndicator?: boolean;
}

export default class UniversalHtmlViewerWebPart extends BaseClientSideWebPart<IUniversalHtmlViewerWebPartProps> {
  private refreshTimerId: number | undefined;
  private iframeLoadTimeoutId: number | undefined;

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

    if (propertyPath === 'configurationPreset' && newValue !== oldValue) {
      this.applyPreset(String(newValue || 'Custom'));
      this.context.propertyPane.refresh();
      this.render();
      return;
    }

    if (
      propertyPath === 'htmlSourceMode' ||
      propertyPath === 'securityMode' ||
      propertyPath === 'sandboxPreset' ||
      propertyPath === 'cacheBusterMode' ||
      propertyPath === 'showChrome'
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
                PropertyPaneToggle('showChrome', {
                  label: 'Show header chrome',
                  onText: 'On',
                  offText: 'Off',
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
                  disabled: !showChrome,
                }),
                PropertyPaneToggle('showOpenInNewTab', {
                  label: 'Show "Open in new tab"',
                  onText: 'On',
                  offText: 'Off',
                  disabled: !showChrome,
                }),
                PropertyPaneToggle('showStatus', {
                  label: 'Show status pill',
                  onText: 'On',
                  offText: 'Off',
                  disabled: !showChrome,
                }),
                PropertyPaneToggle('showLoadingIndicator', {
                  label: 'Show loading indicator',
                  onText: 'On',
                  offText: 'Off',
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
                }),
                PropertyPaneToggle('allowHttp', {
                  label: 'Allow HTTP (unsafe)',
                  onText: 'Allow',
                  offText: 'Block',
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
                  onGetErrorMessage: this.validateAllowedPathPrefixes.bind(this),
                  deferredValidationTime: 200,
                }),
                PropertyPaneTextField('allowedFileExtensions', {
                  label: 'Allowed file extensions (comma-separated)',
                  description: 'Optional file extensions, e.g. .html, .htm',
                  onGetErrorMessage: this.validateAllowedFileExtensions.bind(this),
                  deferredValidationTime: 200,
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
                }),
                PropertyPaneTextField('cacheBusterParamName', {
                  label: 'Cache-buster parameter name',
                  description: 'Defaults to "v" when empty.',
                  disabled: cacheBusterMode === 'None',
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
                }),
                PropertyPaneTextField('iframeSandbox', {
                  label: 'Sandbox tokens',
                  description:
                    'Space-separated sandbox tokens used when Sandbox preset is "Custom". Example: allow-scripts allow-same-origin',
                  disabled: !isCustomSandbox,
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

  private getIframeHeightStyle(): string {
    const heightMode: HeightMode = this.properties.heightMode || 'Fixed';

    if (heightMode === 'Viewport') {
      return 'height:100vh;';
    }

    const fixedHeightPx: number =
      typeof this.properties.fixedHeightPx === 'number' && this.properties.fixedHeightPx > 0
        ? this.properties.fixedHeightPx
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
    const htmlSourceMode: HtmlSourceMode = this.properties.htmlSourceMode || 'FullUrl';

    const finalUrl: string | undefined = buildFinalUrl({
      htmlSourceMode,
      fullUrl: this.properties.fullUrl,
      basePath: this.properties.basePath,
      relativePath: this.properties.relativePath,
      dashboardId: this.properties.dashboardId,
      defaultFileName: this.properties.defaultFileName,
      queryStringParamName: this.properties.queryStringParamName,
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
          }),
        ),
      );
      return;
    }

    const validationOptions: UrlValidationOptions = this.buildUrlValidationOptions(pageUrl);

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
          }),
        ),
      );
      return;
    }

    const iframeHeightStyle: string = this.getIframeHeightStyle();
    const cacheBusterMode: CacheBusterMode = this.properties.cacheBusterMode || 'None';
    const cacheBusterParamName: string = this.normalizeCacheBusterParamName(
      this.properties.cacheBusterParamName,
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
        }),
      ),
      finalUrl,
      cacheBusterMode,
      cacheBusterParamName,
      pageUrl,
      validationOptions,
    );
    this.setupIframeLoadFallback(resolvedUrl);
    this.setupAutoRefresh(finalUrl, cacheBusterMode, cacheBusterParamName, pageUrl);
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
  ): void {
    const iframeTitle: string =
      (this.properties.iframeTitle || '').trim() || 'Universal HTML Viewer';
    const iframeLoading: string = this.normalizeIframeLoading(this.properties.iframeLoading);
    const iframeSandbox: string = this.normalizeIframeSandbox(
      this.properties.iframeSandbox,
      this.properties.sandboxPreset,
    );
    const iframeAllow: string = this.normalizeIframeAllow(this.properties.iframeAllow);
    const iframeReferrerPolicy: string = this.normalizeReferrerPolicy(
      this.properties.iframeReferrerPolicy,
    );
    const chromeHtml: string = this.buildChromeHtml(url, validationOptions, cacheBusterMode);
    const loadingHtml: string = this.buildLoadingHtml();
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

    this.attachChromeHandlers(baseUrl, cacheBusterMode, cacheBusterParamName, pageUrl);
  }

  private buildUrlValidationOptions(currentPageUrl: string): UrlValidationOptions {
    const securityMode: UrlSecurityMode = this.properties.securityMode || 'StrictTenant';
    const allowHttp: boolean = !!this.properties.allowHttp;
    const allowedHosts: string[] = this.parseHosts(this.properties.allowedHosts);
    const allowedPathPrefixes: string[] = this.parsePathPrefixes(
      this.properties.allowedPathPrefixes,
    );
    const allowedFileExtensions: string[] = this.parseFileExtensions(
      this.properties.allowedFileExtensions,
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

  private applyPreset(presetValue: string): void {
    const preset: string = (presetValue || 'Custom').trim();
    this.properties.configurationPreset = preset;

    if (preset === 'Custom') {
      return;
    }

    const basePathPrefix: string = this.normalizeBasePathForPrefix(this.properties.basePath);

    this.properties.allowHttp = false;
    this.properties.allowedFileExtensions = '.html,.htm';
    this.properties.showChrome = true;
    this.properties.showOpenInNewTab = true;
    this.properties.showRefreshButton = true;
    this.properties.showStatus = true;
    this.properties.showLoadingIndicator = true;
    this.properties.iframeLoadTimeoutSeconds = 10;

    if (!this.properties.chromeTitle || this.properties.chromeTitle.trim().length === 0) {
      this.properties.chromeTitle = 'Universal HTML Viewer';
    }

    if (basePathPrefix) {
      this.properties.allowedPathPrefixes = basePathPrefix;
    }

    switch (preset) {
      case 'SharePointLibraryStrict':
        this.properties.securityMode = 'StrictTenant';
        this.properties.cacheBusterMode = 'FileLastModified';
        this.properties.sandboxPreset = 'Strict';
        break;
      case 'SharePointLibraryRelaxed':
        this.properties.securityMode = 'StrictTenant';
        this.properties.cacheBusterMode = 'FileLastModified';
        this.properties.sandboxPreset = 'Relaxed';
        break;
      case 'AllowlistCDN':
        this.properties.securityMode = 'Allowlist';
        this.properties.cacheBusterMode = 'Timestamp';
        this.properties.sandboxPreset = 'Relaxed';
        break;
      case 'AnyHttps':
        this.properties.securityMode = 'AnyHttps';
        this.properties.cacheBusterMode = 'Timestamp';
        this.properties.sandboxPreset = 'None';
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
  ): void {
    const refreshIntervalMs: number = this.getRefreshIntervalMs();

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

  private getRefreshIntervalMs(): number {
    const minutes: number = this.properties.refreshIntervalMinutes || 0;
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
  }

  private clearRefreshTimer(): void {
    if (this.refreshTimerId && typeof window !== 'undefined') {
      window.clearInterval(this.refreshTimerId);
    }
    this.refreshTimerId = undefined;
  }

  private setupIframeLoadFallback(url: string): void {
    this.clearIframeLoadTimeout();

    const timeoutSeconds: number = this.getIframeLoadTimeoutSeconds();
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
          }),
        )}`,
      );
    }, timeoutSeconds * 1000);
  }

  private getIframeLoadTimeoutSeconds(): number {
    const configuredSeconds = this.properties.iframeLoadTimeoutSeconds;
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
      return this.appendQueryParam(baseUrl, cacheBusterParamName, `${Date.now()}`);
    }

    if (cacheBusterMode === 'FileLastModified') {
      const cacheValue = await this.tryGetFileLastModifiedCacheValue(baseUrl, pageUrl);
      if (cacheValue) {
        return this.appendQueryParam(baseUrl, cacheBusterParamName, cacheValue);
      }
      return this.appendQueryParam(baseUrl, cacheBusterParamName, `${Date.now()}`);
    }

    return baseUrl;
  }

  private async tryGetFileLastModifiedCacheValue(
    url: string,
    pageUrl: string,
  ): Promise<string | null> {
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
        return etag.replace(/[^a-zA-Z0-9.-]/g, '');
      }

      if (timeLastModified) {
        return new Date(timeLastModified).getTime().toString();
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
  ): string {
    if (!this.properties.showChrome) {
      return '';
    }

    const title: string =
      (this.properties.chromeTitle || '').trim() ||
      (this.properties.iframeTitle || '').trim() ||
      'Universal HTML Viewer';
    const subtitle: string = (this.properties.chromeSubtitle || '').trim();
    const showOpenInNewTab: boolean = this.properties.showOpenInNewTab !== false;
    const showRefreshButton: boolean = this.properties.showRefreshButton !== false;
    const showStatus: boolean = this.properties.showStatus !== false;

    const statusLabel: string = showStatus
      ? this.getStatusLabel(validationOptions, cacheBusterMode)
      : '';
    const statusHtml: string = statusLabel
      ? `<span class="${styles.status}">${escape(statusLabel)}</span>`
      : '';

    const openInNewTabHtml: string = showOpenInNewTab
      ? `<a class="${styles.actionLink}" href="${escape(resolvedUrl)}" target="_blank" rel="noopener noreferrer">
          Open in new tab
        </a>`
      : '';

    const refreshHtml: string = showRefreshButton
      ? `<button class="${styles.actionButton}" type="button" data-uhv-action="refresh">Refresh</button>`
      : '';

    const subtitleHtml: string = subtitle
      ? `<div class="${styles.chromeSubtitle}">${escape(subtitle)}</div>`
      : '';

    return `
      <div class="${styles.chrome}">
        <div class="${styles.chromeLeft}">
          <div class="${styles.chromeTitle}">${escape(title)}</div>
          ${subtitleHtml}
        </div>
        <div class="${styles.chromeRight}">
          ${statusHtml}
          ${refreshHtml}
          ${openInNewTabHtml}
        </div>
      </div>`;
  }

  private buildLoadingHtml(): string {
    if (this.properties.showLoadingIndicator === false) {
      return '';
    }

    return `<div class="${styles.loading}" data-uhv-loading>Loading…</div>`;
  }

  private attachChromeHandlers(
    baseUrl: string,
    cacheBusterMode: CacheBusterMode,
    cacheBusterParamName: string,
    pageUrl: string,
  ): void {
    const refreshButton: HTMLButtonElement | null = this.domElement.querySelector(
      '[data-uhv-action="refresh"]',
    );
    if (refreshButton) {
      refreshButton.addEventListener('click', () => {
        this.setLoadingVisible(true);
        this.refreshIframe(baseUrl, cacheBusterMode, cacheBusterParamName, pageUrl).catch(() => {
          return undefined;
        });
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

  private getStatusLabel(
    validationOptions: UrlValidationOptions,
    cacheBusterMode: CacheBusterMode,
  ): string {
    const parts: string[] = [];

    parts.push(validationOptions.securityMode);
    if (validationOptions.securityMode === 'Allowlist') {
      parts.push('Allowlist');
    }

    if (validationOptions.allowHttp) {
      parts.push('HTTP allowed');
    }

    if (cacheBusterMode !== 'None') {
      parts.push(`Cache: ${cacheBusterMode}`);
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

  private buildDiagnosticsHtml(data?: Record<string, unknown>): string {
    if (!this.properties.showDiagnostics || !data) {
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

  private buildDiagnosticsData(values: Record<string, unknown>): Record<string, unknown> {
    return {
      timestamp: new Date().toISOString(),
      ...values,
      configurationPreset: this.properties.configurationPreset || 'Custom',
      allowHttp: !!this.properties.allowHttp,
      allowedHosts: this.parseHosts(this.properties.allowedHosts),
      allowedPathPrefixes: this.parsePathPrefixes(this.properties.allowedPathPrefixes),
      allowedFileExtensions: this.parseFileExtensions(this.properties.allowedFileExtensions),
      securityMode: this.properties.securityMode || 'StrictTenant',
      cacheBusterMode: this.properties.cacheBusterMode || 'None',
      sandboxPreset: this.properties.sandboxPreset || 'None',
      iframeSandbox: this.properties.iframeSandbox || '',
      iframeLoadTimeoutSeconds: this.getIframeLoadTimeoutSeconds(),
      showChrome: this.properties.showChrome !== false,
      showLoadingIndicator: this.properties.showLoadingIndicator !== false,
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

  private hasDotSegments(pathname: string): boolean {
    const segments = pathname.split('/').filter((segment) => segment.length > 0);
    return segments.some((segment) => segment === '.' || segment === '..');
  }
}
