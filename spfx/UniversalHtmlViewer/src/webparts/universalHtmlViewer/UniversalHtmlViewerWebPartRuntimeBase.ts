/* eslint-disable max-lines */
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './UniversalHtmlViewerWebPart.module.scss';
import { buildOpenInNewTabHtml, buildMessageHtml } from './MarkupHelper';
import { getQueryStringParam } from './QueryStringHelper';
import { resolveAutoRefreshTarget, shouldExecuteAutoRefresh } from './AutoRefreshHelper';
import { CacheBusterMode, UrlValidationOptions } from './UrlHelper';
import {
  clearIframeLoadFallbackLifecycleState,
  IIframeLoadFallbackState,
  setupIframeLoadFallbackLifecycleState,
} from './IframeLoadFallbackHelper';
import {
  ContentDeliveryMode,
  IUniversalHtmlViewerWebPartProps,
} from './UniversalHtmlViewerTypes';
import { UniversalHtmlViewerWebPartConfigBase } from './UniversalHtmlViewerWebPartConfigBase';

export abstract class UniversalHtmlViewerWebPartRuntimeBase extends UniversalHtmlViewerWebPartConfigBase {
  protected refreshInProgress: boolean = false;
  private readonly iframeLoadFallbackState: IIframeLoadFallbackState = {};

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
        const shouldRunRefresh = shouldExecuteAutoRefresh({
          refreshInProgress: this.refreshInProgress,
          documentHidden: typeof document !== 'undefined' ? document.hidden : false,
        });
        if (!shouldRunRefresh) {
          return;
        }

        const activeTarget = resolveAutoRefreshTarget({
          baseUrl,
          pageUrl,
          currentBaseUrl: this.currentBaseUrl,
          currentPageUrl: this.getCurrentPageUrl(),
        });
        this.refreshIframe(
          activeTarget.baseUrl,
          cacheBusterMode,
          cacheBusterParamName,
          activeTarget.pageUrl,
        ).catch(() => {
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
    resetInlineScrollToTop: boolean = false,
    preserveHostScrollPosition: boolean = false,
    bypassInlineContentCache: boolean = false,
  ): Promise<void> {
    if (this.refreshInProgress) {
      return;
    }

    this.refreshInProgress = true;
    try {
      const iframe: HTMLIFrameElement | null = this.domElement.querySelector('iframe');
      if (!iframe) {
        return;
      }

      const hostScrollPosition = preserveHostScrollPosition
        ? this.captureHostScrollPosition()
        : undefined;

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
          bypassInlineContentCache,
        );
        if (updatedFromContent) {
          if (resetInlineScrollToTop) {
            this.resetIframeScrollPosition(iframe, refreshedUrl);
          }
          if (hostScrollPosition) {
            this.restoreHostScrollPosition(hostScrollPosition);
          }
          this.updateStatusBadge(this.lastValidationOptions, cacheBusterMode, this.lastEffectiveProps);
          return;
        }

        // Keep the current inline content when refresh fails to avoid direct SharePoint file
        // navigation, which can trigger browser downloads instead of rendering.
        this.setLoadingVisible(false);
        if (hostScrollPosition) {
          this.restoreHostScrollPosition(hostScrollPosition);
        }
        return;
      }

      const activeIframe: HTMLIFrameElement | null = this.domElement.querySelector('iframe');
      if (!activeIframe || activeIframe !== iframe) {
        return;
      }

      if (hostScrollPosition) {
        const restoreAfterLoad = (): void => {
          activeIframe.removeEventListener('load', restoreAfterLoad);
          this.restoreHostScrollPosition(hostScrollPosition);
        };
        activeIframe.addEventListener('load', restoreAfterLoad);
      }

      activeIframe.src = refreshedUrl;
      this.updateStatusBadge(this.lastValidationOptions, cacheBusterMode, this.lastEffectiveProps);
    } finally {
      this.refreshInProgress = false;
    }
  }

  protected resetIframeScrollPosition(iframe: HTMLIFrameElement, targetUrl?: string): void {
    if (hasUrlHash(targetUrl)) {
      return;
    }

    const applyReset = (): void => {
      try {
        const iframeWindow = iframe.contentWindow;
        if (iframeWindow) {
          iframeWindow.scrollTo(0, 0);
        }

        const iframeDocument = iframe.contentDocument;
        if (iframeDocument?.documentElement) {
          iframeDocument.documentElement.scrollTop = 0;
          iframeDocument.documentElement.scrollLeft = 0;
        }
        if (iframeDocument?.body) {
          iframeDocument.body.scrollTop = 0;
          iframeDocument.body.scrollLeft = 0;
        }
        this.resetIframeDeepScrollPosition(iframe);
      } catch {
        return;
      }
    };

    applyReset();
    if (typeof window !== 'undefined') {
      window.setTimeout(applyReset, 0);
      window.setTimeout(applyReset, 120);
      window.setTimeout(applyReset, 350);
      window.setTimeout(applyReset, 800);
    }
  }
  protected getIframeDeepMaxScrollTop(iframe: HTMLIFrameElement): number {
    let maxTop = 0;

    try {
      const iframeWindow = iframe.contentWindow;
      const windowTop = iframeWindow?.scrollY || 0;
      if (windowTop > maxTop) {
        maxTop = windowTop;
      }
    } catch {
      return maxTop;
    }

    try {
      const iframeDocument = iframe.contentDocument;
      const documentMaxTop = this.getDocumentDeepMaxScrollTop(iframeDocument);
      if (documentMaxTop > maxTop) {
        maxTop = documentMaxTop;
      }
    } catch {
      return maxTop;
    }

    return maxTop;
  }
  protected resetIframeDeepScrollPosition(iframe: HTMLIFrameElement): void {
    try {
      this.resetDocumentDeepScrollPosition(iframe.contentDocument);
    } catch {
      return;
    }
  }

  protected captureHostScrollPosition(): { x: number; y: number } | undefined {
    if (typeof window === 'undefined' || typeof document === 'undefined') {
      return undefined;
    }

    const hostScrollContainers = this.getPotentialHostScrollContainers();
    const activeHostScrollContainer = hostScrollContainers.find(
      (container) => container.scrollTop !== 0 || container.scrollLeft !== 0,
    );
    const hostScrollContainer = activeHostScrollContainer || hostScrollContainers[0];
    if (hostScrollContainer) {
      return {
        x: hostScrollContainer.scrollLeft || 0,
        y: hostScrollContainer.scrollTop || 0,
      };
    }

    const scrollingElement = document.scrollingElement as HTMLElement | null;
    if (scrollingElement) {
      return {
        x: scrollingElement.scrollLeft || 0,
        y: scrollingElement.scrollTop || 0,
      };
    }

    return {
      x: window.scrollX || 0,
      y: window.scrollY || 0,
    };
  }

  protected restoreHostScrollPosition(position: { x: number; y: number }): void {
    if (!position || typeof window === 'undefined' || typeof document === 'undefined') {
      return;
    }

    const applyRestore = (): void => {
      const hostScrollContainers = this.getPotentialHostScrollContainers();
      hostScrollContainers.forEach((hostScrollContainer) => {
        hostScrollContainer.scrollLeft = position.x;
        hostScrollContainer.scrollTop = position.y;
      });

      const scrollingElement = document.scrollingElement as HTMLElement | null;
      if (scrollingElement) {
        scrollingElement.scrollLeft = position.x;
        scrollingElement.scrollTop = position.y;
      }

      if (document.documentElement) {
        document.documentElement.scrollLeft = position.x;
        document.documentElement.scrollTop = position.y;
      }
      if (document.body) {
        document.body.scrollLeft = position.x;
        document.body.scrollTop = position.y;
      }

      window.scrollTo(position.x, position.y);
    };

    applyRestore();
    window.setTimeout(applyRestore, 0);
    window.setTimeout(applyRestore, 120);
    window.setTimeout(applyRestore, 350);
    window.setTimeout(applyRestore, 800);
  }
  protected forceHostScrollTop(): void {
    if (typeof window === 'undefined' || typeof document === 'undefined') {
      return;
    }

    const touched = new Set<HTMLElement>();
    const setTop = (element?: HTMLElement | null): void => {
      if (!element || touched.has(element)) {
        return;
      }

      touched.add(element);
      element.scrollTop = 0;
      element.scrollLeft = 0;
    };

    setTop(document.scrollingElement as HTMLElement | null);
    setTop(document.documentElement);
    setTop(document.body);

    this.getPotentialHostScrollContainers().forEach((container) => {
      setTop(container);
    });

    // Try common SharePoint scroll containers explicitly.
    const sharePointScrollSelectors = [
      '#spPageChromeAppDiv',
      '[data-automationid="contentScrollRegion"]',
      '.SPPageChrome',
      '.CanvasZoneContainer',
      '.CanvasZone',
      '.CanvasComponent',
      '[role="main"]',
      'main',
    ];
    const sharePointCandidates = document.querySelectorAll<HTMLElement>(
      sharePointScrollSelectors.join(','),
    );
    sharePointCandidates.forEach((element) => {
      setTop(element);
    });

    // Active element ancestry can include hidden scroll wrappers.
    let current = document.activeElement as HTMLElement | null;
    while (current) {
      setTop(current);
      current = current.parentElement;
    }

    const topMarker = this.ensureHostTopMarker();
    if (topMarker) {
      try {
        topMarker.scrollIntoView({
          block: 'start',
          inline: 'nearest',
          behavior: 'auto',
        });
      } catch {
        try {
          topMarker.scrollIntoView(true);
        } catch {
          return;
        }
      }
    }

    window.scrollTo(0, 0);
  }
  protected ensureHostTopMarker(): HTMLElement | undefined {
    if (typeof document === 'undefined' || !document.body) {
      return undefined;
    }

    const markerId = 'uhv-scroll-top-marker';
    let marker = document.getElementById(markerId) as HTMLElement | undefined;
    if (!marker) {
      marker = document.createElement('div');
      marker.id = markerId;
      marker.setAttribute('aria-hidden', 'true');
      marker.style.position = 'absolute';
      marker.style.top = '0';
      marker.style.left = '0';
      marker.style.width = '1px';
      marker.style.height = '1px';
      marker.style.pointerEvents = 'none';
      marker.style.opacity = '0';
      document.body.insertBefore(marker, document.body.firstChild);
    }

    return marker;
  }

  protected getPotentialHostScrollContainers(): HTMLElement[] {
    if (typeof document === 'undefined') {
      return [];
    }

    const candidates = new Set<HTMLElement>();
    const scrollingElement = document.scrollingElement;
    if (scrollingElement instanceof HTMLElement) {
      candidates.add(scrollingElement);
    }
    if (document.documentElement) {
      candidates.add(document.documentElement);
    }
    if (document.body) {
      candidates.add(document.body);
    }

    let current: HTMLElement | null = this.domElement;
    while (current) {
      candidates.add(current);
      current = current.parentElement;
    }

    const result: HTMLElement[] = [];
    candidates.forEach((candidate) => {
      const overflowDelta = candidate.scrollHeight - candidate.clientHeight;
      if (overflowDelta > 2) {
        result.push(candidate);
      }
    });

    return result;
  }

  protected getHostScrollContainer(): HTMLElement | undefined {
    const candidates = this.getPotentialHostScrollContainers();
    let best: HTMLElement | undefined;
    let bestOverflowDelta = 0;
    candidates.forEach((candidate) => {
      const overflowDelta = candidate.scrollHeight - candidate.clientHeight;
      if (overflowDelta <= 2) {
        return;
      }

      const isDocumentRoot =
        candidate === document.documentElement ||
        candidate === document.body ||
        candidate === document.scrollingElement;
      if (!isDocumentRoot) {
        const style = window.getComputedStyle(candidate);
        const overflowY = (style.overflowY || '').toLowerCase();
        const isScrollableByStyle =
          overflowY === 'auto' || overflowY === 'scroll' || overflowY === 'overlay';
        if (!isScrollableByStyle) {
          return;
        }
      }

      if (overflowDelta > bestOverflowDelta) {
        best = candidate;
        bestOverflowDelta = overflowDelta;
      }
    });

    return best;
  }
  protected isScrollTraceEnabled(): boolean {
    if (typeof window === 'undefined') {
      return false;
    }

    const pageUrl = window.location.href || '';
    const queryValue = (getQueryStringParam(pageUrl, 'uhvTraceScroll') || '').trim();
    return isEnabledDebugValue(queryValue);
  }
  protected emitScrollTrace(
    eventName: string,
    data?: Record<string, unknown>,
  ): void {
    if (!this.isScrollTraceEnabled() || typeof window === 'undefined') {
      return;
    }

    const payload: Record<string, unknown> = {
      event: eventName,
      isoTime: new Date().toISOString(),
      ...data,
      snapshot: this.buildScrollTraceSnapshot(),
    };
    const traceWindow = window as Window & {
      __uhvScrollTrace?: Array<Record<string, unknown>>;
    };
    const traceBuffer = traceWindow.__uhvScrollTrace || [];
    traceBuffer.push(payload);
    if (traceBuffer.length > 300) {
      traceBuffer.splice(0, traceBuffer.length - 300);
    }
    traceWindow.__uhvScrollTrace = traceBuffer;

    // eslint-disable-next-line no-console
    console.warn('[UHV scroll trace]', payload);
  }
  protected buildScrollTraceSnapshot(): Record<string, unknown> {
    if (typeof window === 'undefined' || typeof document === 'undefined') {
      return {};
    }

    const containers = this.getPotentialHostScrollContainers();
    const containerState = containers.map((container) => ({
      element: this.describeScrollElement(container),
      top: container.scrollTop,
      left: container.scrollLeft,
      height: container.clientHeight,
      scrollHeight: container.scrollHeight,
    }));

    return {
      window: {
        top: window.scrollY || 0,
        left: window.scrollX || 0,
      },
      iframe: this.getInlineIframeScrollState(),
      scrollingElement: describeElementScrollState(document.scrollingElement),
      documentElement: describeElementScrollState(document.documentElement),
      body: describeElementScrollState(document.body),
      activeElement: this.describeScrollElement(
        document.activeElement instanceof HTMLElement ? document.activeElement : undefined,
      ),
      hostContainers: containerState,
    };
  }
  protected getInlineIframeScrollState(): Record<string, unknown> | undefined {
    const iframe = this.domElement.querySelector('iframe');
    if (!iframe) {
      return undefined;
    }

    const result: Record<string, unknown> = {
      element: this.describeScrollElement(iframe),
    };
    try {
      const iframeWindow = iframe.contentWindow;
      result.windowTop = iframeWindow?.scrollY || 0;
      result.windowLeft = iframeWindow?.scrollX || 0;
    } catch {
      result.windowAccess = 'blocked';
    }

    try {
      const iframeDocument = iframe.contentDocument;
      if (iframeDocument?.documentElement) {
        result.documentTop = iframeDocument.documentElement.scrollTop || 0;
      }
      if (iframeDocument?.body) {
        result.bodyTop = iframeDocument.body.scrollTop || 0;
      }
      result.deepMaxTop = this.getIframeDeepMaxScrollTop(iframe);
    } catch {
      result.documentAccess = 'blocked';
    }

    return result;
  }
  protected describeScrollElement(element?: HTMLElement | null): string {
    if (!element) {
      return '(none)';
    }

    const parts: string[] = [element.tagName.toLowerCase()];
    if (element.id) {
      parts.push(`#${element.id}`);
    }
    const className = (element.className || '').toString().trim();
    if (className) {
      const firstClass = className.split(/\s+/g)[0];
      if (firstClass) {
        parts.push(`.${firstClass}`);
      }
    }

    return parts.join('');
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
    _bypassInlineContentCache: boolean = false,
  ): Promise<boolean> {
    return false;
  }

  protected onNavigatedToUrl(_targetUrl: string, _pageUrl: string): void {
    return;
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
    this.clearIframeLoadFallbackLifecycle();

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

    setupIframeLoadFallbackLifecycleState({
      state: this.iframeLoadFallbackState,
      iframe,
      timeoutMs: timeoutSeconds * 1000,
      setTimeoutFn: (handler: () => void, timeoutMs: number): number =>
        window.setTimeout(handler, timeoutMs),
      clearTimeoutFn: (timeoutId: number): void => {
        window.clearTimeout(timeoutId);
      },
      onLoad: (): void => {
        this.iframeLoadTimeoutId = undefined;
        this.setLoadingVisible(false);
      },
      onTimeout: (): void => {
        this.iframeLoadTimeoutId = undefined;
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
      },
    });
    this.iframeLoadTimeoutId = this.iframeLoadFallbackState.timeoutId;
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
    this.clearIframeLoadFallbackLifecycle();
  }
  protected clearIframeLoadFallbackListener(): void {
    this.clearIframeLoadFallbackLifecycle();
  }
  private clearIframeLoadFallbackLifecycle(): void {
    clearIframeLoadFallbackLifecycleState(
      this.iframeLoadFallbackState,
      (timeoutId: number): void => {
        if (typeof window !== 'undefined' && typeof window.clearTimeout === 'function') {
          window.clearTimeout(timeoutId);
          return;
        }

        clearTimeout(timeoutId as unknown as ReturnType<typeof setTimeout>);
      },
    );
    this.iframeLoadTimeoutId = this.iframeLoadFallbackState.timeoutId;
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
      enableExpertSecurityModes: props.enableExpertSecurityModes === true,
      allowHttp: !!props.allowHttp,
      allowedHosts: this.parseHosts(props.allowedHosts),
      allowedPathPrefixes: this.parsePathPrefixes(props.allowedPathPrefixes),
      allowedFileExtensions: this.parseFileExtensions(props.allowedFileExtensions),
      securityMode: props.securityMode || 'StrictTenant',
      tenantConfigUrl: props.tenantConfigUrl || '',
      tenantConfigMode: props.tenantConfigMode || 'Merge',
      dashboardList: props.dashboardList || '',
      cacheBusterMode: props.cacheBusterMode || 'None',
      inlineContentCacheTtlSeconds: props.inlineContentCacheTtlSeconds ?? 15,
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
    this.clearIframeLoadFallbackListener();
  }
  private getDocumentDeepMaxScrollTop(
    iframeDocument?: Document,
    depth: number = 0,
  ): number {
    if (!iframeDocument || depth > 2) {
      return 0;
    }

    let maxTop = 0;
    const documentElementTop = iframeDocument.documentElement?.scrollTop || 0;
    if (documentElementTop > maxTop) {
      maxTop = documentElementTop;
    }
    const bodyTop = iframeDocument.body?.scrollTop || 0;
    if (bodyTop > maxTop) {
      maxTop = bodyTop;
    }

    const queryRoot = iframeDocument.body || iframeDocument.documentElement;
    if (!queryRoot) {
      return maxTop;
    }

    const allElements = queryRoot.querySelectorAll<HTMLElement>('*');
    const maxElements = 2500;
    const scanLength = Math.min(allElements.length, maxElements);
    for (let index = 0; index < scanLength; index += 1) {
      const element = allElements[index];
      const elementTop = element.scrollTop || 0;
      if (elementTop > maxTop) {
        maxTop = elementTop;
      }

      if (element instanceof HTMLIFrameElement) {
        let nestedDocument: Document | undefined;
        try {
          nestedDocument = element.contentDocument || undefined;
        } catch {
          nestedDocument = undefined;
        }
        const nestedMax = this.getDocumentDeepMaxScrollTop(nestedDocument, depth + 1);
        if (nestedMax > maxTop) {
          maxTop = nestedMax;
        }
      }
    }

    return maxTop;
  }
  private resetDocumentDeepScrollPosition(
    iframeDocument?: Document,
    depth: number = 0,
  ): void {
    if (!iframeDocument || depth > 2) {
      return;
    }

    if (iframeDocument.documentElement) {
      iframeDocument.documentElement.scrollTop = 0;
      iframeDocument.documentElement.scrollLeft = 0;
    }
    if (iframeDocument.body) {
      iframeDocument.body.scrollTop = 0;
      iframeDocument.body.scrollLeft = 0;
    }

    const queryRoot = iframeDocument.body || iframeDocument.documentElement;
    if (!queryRoot) {
      return;
    }

    const allElements = queryRoot.querySelectorAll<HTMLElement>('*');
    const maxElements = 2500;
    const scanLength = Math.min(allElements.length, maxElements);
    for (let index = 0; index < scanLength; index += 1) {
      const element = allElements[index];
      if (element.scrollTop || element.scrollLeft) {
        element.scrollTop = 0;
        element.scrollLeft = 0;
      }

      if (element instanceof HTMLIFrameElement) {
        let nestedDocument: Document | undefined;
        try {
          const nestedWindow = element.contentWindow;
          if (nestedWindow) {
            nestedWindow.scrollTo(0, 0);
          }
          nestedDocument = element.contentDocument || undefined;
        } catch {
          // Ignore nested iframe cross-origin errors.
          nestedDocument = undefined;
        }
        this.resetDocumentDeepScrollPosition(nestedDocument, depth + 1);
      }
    }
  }
}

function hasUrlHash(url?: string): boolean {
  const value = (url || '').trim();
  if (!value) {
    return false;
  }

  const hashIndex = value.indexOf('#');
  if (hashIndex < 0) {
    return false;
  }

  return hashIndex < value.length - 1;
}

function isEnabledDebugValue(value?: string): boolean {
  const normalized = (value || '').trim().toLowerCase();
  return normalized === '1' || normalized === 'true' || normalized === 'yes' || normalized === 'on';
}

function describeElementScrollState(
  element?: Element,
): Record<string, unknown> | undefined {
  if (!(element instanceof HTMLElement)) {
    return undefined;
  }

  return {
    element: `${element.tagName.toLowerCase()}${element.id ? `#${element.id}` : ''}`,
    top: element.scrollTop || 0,
    left: element.scrollLeft || 0,
    height: element.clientHeight,
    scrollHeight: element.scrollHeight,
  };
}
