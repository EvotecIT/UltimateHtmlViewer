import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField,
  PropertyPaneSlider,
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
  allowedHosts?: string;
  allowedPathPrefixes?: string;
  cacheBusterMode?: CacheBusterMode;
  cacheBusterParamName?: string;
  iframeSandbox?: string;
  iframeAllow?: string;
  iframeReferrerPolicy?: string;
  iframeLoading?: string;
  iframeTitle?: string;
  refreshIntervalMinutes?: number;
}

export default class UniversalHtmlViewerWebPart extends BaseClientSideWebPart<IUniversalHtmlViewerWebPartProps> {
  private refreshTimerId: number | undefined;

  public render(): void {
    this.renderAsync().catch((error) => {
      this.clearRefreshTimer();
      this.domElement.innerHTML = this.buildMessageHtml(
        'UniversalHtmlViewer: Failed to render content.',
      );
      // eslint-disable-next-line no-console
      console.error('UniversalHtmlViewer render failed', error);
    });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Configure the HTML source and layout.',
          },
          groups: [
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
                }),
                PropertyPaneTextField('basePath', {
                  label: 'Base path (site-relative)',
                  description:
                    'Site-relative base path, used when HTML source mode is not "FullUrl". Example: /sites/Reports/Dashboards/',
                }),
                PropertyPaneTextField('relativePath', {
                  label: 'Relative path from base',
                  description:
                    'Used when HTML source mode is "BasePathAndRelativePath". Example: system1/index.html',
                }),
                PropertyPaneTextField('dashboardId', {
                  label: 'Dashboard ID (fallback)',
                  description:
                    'Used when HTML source mode is "BasePathAndDashboardId" and no query string parameter is provided.',
                }),
                PropertyPaneTextField('defaultFileName', {
                  label: 'Default file name',
                  description:
                    'Used when HTML source mode is "BasePathAndDashboardId". Defaults to "index.html" when left empty.',
                }),
                PropertyPaneTextField('queryStringParamName', {
                  label: 'Query string parameter name',
                  description:
                    'Used when HTML source mode is "BasePathAndDashboardId" to read the dashboard ID from the page URL. Defaults to "dashboard" when left empty.',
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
                PropertyPaneTextField('allowedHosts', {
                  label: 'Allowed hosts (comma-separated)',
                  description:
                    'Used when security mode is "Allowlist". Example: cdn.contoso.com, files.contoso.net',
                }),
                PropertyPaneTextField('allowedPathPrefixes', {
                  label: 'Allowed path prefixes (comma-separated)',
                  description:
                    'Optional site-relative path prefixes that the URL must start with. Example: /sites/Reports/Dashboards/',
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
                PropertyPaneTextField('iframeSandbox', {
                  label: 'Sandbox tokens',
                  description:
                    'Space-separated sandbox tokens. Leave empty for no sandbox. Example: allow-scripts allow-same-origin',
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
  }

  private async renderAsync(): Promise<void> {
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
      this.domElement.innerHTML = this.buildMessageHtml(
        'UniversalHtmlViewer: No URL configured. Please update the web part settings.',
      );
      return;
    }

    const validationOptions: UrlValidationOptions = this.buildUrlValidationOptions(pageUrl);

    if (!isUrlAllowed(finalUrl, validationOptions)) {
      this.clearRefreshTimer();
      this.domElement.innerHTML = this.buildMessageHtml(
        'UniversalHtmlViewer: The target URL is invalid or not allowed.',
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

    this.renderIframe(resolvedUrl, iframeHeightStyle);
    this.setupAutoRefresh(finalUrl, cacheBusterMode, cacheBusterParamName, pageUrl);
  }

  private renderIframe(url: string, iframeHeightStyle: string): void {
    const iframeTitle: string =
      (this.properties.iframeTitle || '').trim() || 'Universal HTML Viewer';
    const iframeLoading: string = this.normalizeIframeLoading(this.properties.iframeLoading);
    const iframeSandbox: string = this.normalizeIframeSandbox(this.properties.iframeSandbox);
    const iframeAllow: string = this.normalizeIframeAllow(this.properties.iframeAllow);
    const iframeReferrerPolicy: string = this.normalizeReferrerPolicy(
      this.properties.iframeReferrerPolicy,
    );

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
        <iframe class="${styles.iframe}"
          src="${escape(url)}"
          title="${escape(iframeTitle)}"
          style="${iframeHeightStyle}border:0;"
          width="100%"
          frameborder="0"${loadingAttribute}${sandboxAttribute}${allowAttribute}${referrerPolicyAttribute}
        ></iframe>
      </div>`;
  }

  private buildUrlValidationOptions(currentPageUrl: string): UrlValidationOptions {
    const securityMode: UrlSecurityMode = this.properties.securityMode || 'StrictTenant';
    const allowedHosts: string[] = this.parseHosts(this.properties.allowedHosts);
    const allowedPathPrefixes: string[] = this.parsePathPrefixes(
      this.properties.allowedPathPrefixes,
    );

    return {
      securityMode,
      currentPageUrl,
      allowedHosts,
      allowedPathPrefixes,
    };
  }

  private parseHosts(value?: string): string[] {
    return (value || '')
      .split(/[,;\s]+/g)
      .map((entry) => entry.trim())
      .filter((entry) => entry.length > 0)
      .map((entry) => {
        try {
          if (entry.includes('://')) {
            return new URL(entry).host.toLowerCase();
          }
          const withoutPath: string = entry.split('/')[0];
          return withoutPath.toLowerCase();
        } catch {
          return entry.toLowerCase();
        }
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

  private normalizeIframeSandbox(value?: string): string {
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

  private buildMessageHtml(message: string): string {
    return `
      <div class="${styles.universalHtmlViewer}">
        <div class="${styles.message}">${escape(message)}</div>
      </div>`;
  }
}
