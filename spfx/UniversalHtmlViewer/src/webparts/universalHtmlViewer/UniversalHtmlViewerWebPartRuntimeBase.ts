import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './UniversalHtmlViewerWebPart.module.scss';
import { buildOpenInNewTabHtml, buildMessageHtml } from './MarkupHelper';
import { CacheBusterMode, UrlValidationOptions } from './UrlHelper';
import {
  ContentDeliveryMode,
  IUniversalHtmlViewerWebPartProps,
} from './UniversalHtmlViewerTypes';
import { UniversalHtmlViewerWebPartConfigBase } from './UniversalHtmlViewerWebPartConfigBase';

export abstract class UniversalHtmlViewerWebPartRuntimeBase extends UniversalHtmlViewerWebPartConfigBase {
  protected setupAutoRefresh(
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

  protected async refreshIframe(
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

    const effectiveProps: IUniversalHtmlViewerWebPartProps =
      this.lastEffectiveProps || this.properties;
    const contentDeliveryMode: ContentDeliveryMode = this.getContentDeliveryMode(
      effectiveProps,
    );

    if (contentDeliveryMode === 'SharePointFileContent') {
      const updatedFromContent: boolean = await this.trySetIframeSrcDocFromSource(
        iframe,
        refreshedUrl,
        pageUrl,
        effectiveProps,
      );
      if (updatedFromContent) {
        this.updateStatusBadge(this.lastValidationOptions, cacheBusterMode, this.lastEffectiveProps);
        return;
      }
    }

    iframe.src = refreshedUrl;
    this.updateStatusBadge(this.lastValidationOptions, cacheBusterMode, this.lastEffectiveProps);
  }

  protected getContentDeliveryMode(
    props: IUniversalHtmlViewerWebPartProps,
  ): ContentDeliveryMode {
    return props.contentDeliveryMode || 'DirectUrl';
  }

  protected async trySetIframeSrcDocFromSource(
    _iframe: HTMLIFrameElement,
    _sourceUrl: string,
    _pageUrl: string,
    _props: IUniversalHtmlViewerWebPartProps,
  ): Promise<boolean> {
    return false;
  }

  protected clearRefreshTimer(): void {
    if (this.refreshTimerId && typeof window !== 'undefined') {
      window.clearInterval(this.refreshTimerId);
    }
    this.refreshTimerId = undefined;
  }

  protected setupIframeLoadFallback(
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
      this.domElement.innerHTML = buildMessageHtml(
        'UniversalHtmlViewer: The content did not load in time. It may be blocked by SharePoint security headers.',
        `${buildOpenInNewTabHtml(url, styles.fallback, styles.fallbackLink)}${this.buildDiagnosticsHtml(
          this.buildDiagnosticsData({
            resolvedUrl: url,
            timeoutSeconds,
          }, props),
          props,
        )}`,
        styles.universalHtmlViewer,
        styles.message,
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

  protected clearIframeLoadTimeout(): void {
    if (this.iframeLoadTimeoutId && typeof window !== 'undefined') {
      window.clearTimeout(this.iframeLoadTimeoutId);
    }
    this.iframeLoadTimeoutId = undefined;
  }

  protected async resolveUrlWithCacheBuster(
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
      const pathOnly = this.stripQueryAndHash(url);
      return pathOnly || null;
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

  private stripQueryAndHash(value: string): string {
    const hashIndex = value.indexOf('#');
    const queryIndex = value.indexOf('?');

    if (hashIndex === -1 && queryIndex === -1) {
      return value;
    }

    const cutIndex =
      hashIndex === -1
        ? queryIndex
        : queryIndex === -1
          ? hashIndex
          : Math.min(hashIndex, queryIndex);

    return value.substring(0, cutIndex);
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

  protected updateStatusBadge(
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

  protected getStatusLabel(
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

  protected buildDiagnosticsHtml(
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

  protected buildDiagnosticsData(
    values: Record<string, unknown>,
    effectiveProps?: IUniversalHtmlViewerWebPartProps,
  ): Record<string, unknown> {
    const props: IUniversalHtmlViewerWebPartProps = effectiveProps || this.properties;
    return {
      timestamp: new Date().toISOString(),
      ...values,
      configurationPreset: props.configurationPreset || 'Custom',
      lockPresetSettings: !!props.lockPresetSettings,
      contentDeliveryMode: this.getContentDeliveryMode(props),
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
      fitContentWidth: props.fitContentWidth === true,
      showLastUpdated: props.showLastUpdated !== false,
      chromeDensity: props.chromeDensity || 'Comfortable',
      showConfigActions: props.showConfigActions === true,
      showDashboardSelector: props.showDashboardSelector === true,
    };
  }

  protected setLoadingVisible(visible: boolean): void {
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

  protected onDispose(): void {
    this.clearRefreshTimer();
    this.clearIframeLoadTimeout();
  }
}
