/* eslint-disable max-lines */
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient } from '@microsoft/sp-http';
import styles from './UniversalHtmlViewerWebPart.module.scss';
import {
  applyImportedConfigToProps,
  buildConfigExport,
  IConfigImportResult,
} from './ConfigImportExportHelper';
import { CacheBusterMode, HeightMode, isUrlAllowed, UrlValidationOptions } from './UrlHelper';
import {
  ConfigurationPreset,
  ContentDeliveryMode,
  IUniversalHtmlViewerWebPartProps,
  isInlineContentDeliveryMode,
  isReportBrowserSourceMode,
  ReportBrowserDefaultView,
} from './UniversalHtmlViewerTypes';
import {
  buildOpenInNewTabUrl,
} from './InlineDeepLinkHelper';
import { UniversalHtmlViewerWebPartRuntimeBase } from './UniversalHtmlViewerWebPartRuntimeBase';
import {
  getReportBrowserParentFolderPath,
  ISharePointReportBrowserItem,
  isPathInsideRoot,
  loadSharePointReportBrowserItems,
  normalizeSharePointReportBrowserRootPath,
} from './SharePointReportBrowserHelper';

export abstract class UniversalHtmlViewerWebPartUiBase extends UniversalHtmlViewerWebPartRuntimeBase {
  private reportBrowserView: ReportBrowserDefaultView | undefined;
  private reportBrowserFolderPath: string | undefined;
  private reportBrowserRootPath: string | undefined;
  private reportBrowserDefaultView: ReportBrowserDefaultView | undefined;
  private reportBrowserLoadRequestId: number = 0;

  protected getIframeHeightStyle(props: IUniversalHtmlViewerWebPartProps): string {
    const heightMode: HeightMode = props.heightMode || 'Fixed';

    const fixedHeightPx: number =
      typeof props.fixedHeightPx === 'number' && props.fixedHeightPx > 0
        ? props.fixedHeightPx
        : 800;

    if (heightMode === 'Auto') {
      return `height:${fixedHeightPx}px;`;
    }

    if (heightMode === 'Viewport') {
      return 'height:100vh;';
    }

    return `height:${fixedHeightPx}px;`;
  }

  protected renderIframe(
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
    srcDocHtml?: string,
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
      baseUrl,
      pageUrl,
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
          src="${escape(srcDocHtml ? 'about:blank' : url)}"
          title="${escape(iframeTitle)}"
          style="${iframeHeightStyle}border:0;"
          width="100%"
          frameborder="0"${loadingAttribute}${sandboxAttribute}${allowAttribute}${referrerPolicyAttribute}
        ></iframe>
        </div>
      </div>${diagnosticsHtml}`;

    if (srcDocHtml) {
      const iframe: HTMLIFrameElement | null = this.domElement.querySelector('iframe');
      if (iframe) {
        iframe.srcdoc = srcDocHtml;
      }
    }

    this.attachChromeHandlers(
      baseUrl,
      cacheBusterMode,
      cacheBusterParamName,
      pageUrl,
      effectiveProps,
    );
  }

  private buildChromeHtml(
    resolvedUrl: string,
    baseUrl: string,
    pageUrl: string,
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
    const openInNewTabUrl: string = this.getOpenInNewTabUrl(
      resolvedUrl,
      baseUrl,
      pageUrl,
      props,
    );

    const statusLabel: string = showStatus
      ? this.getStatusLabel(validationOptions, cacheBusterMode, props)
      : '';
    const statusHtml: string = statusLabel
      ? `<span class="${styles.status}" data-uhv-status>${escape(statusLabel)}</span>`
      : '';

    const openInNewTabHtml: string = showOpenInNewTab
      ? `<a class="${styles.actionLink}" href="${escape(openInNewTabUrl)}" target="_blank" rel="noopener noreferrer" data-uhv-action="open-in-new-tab">
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
    const anyHttpsWarningHtml: string =
      validationOptions.securityMode === 'AnyHttps'
        ? `<div class="${styles.anyHttpsWarning}">
            Warning: Any HTTPS mode is enabled. Restrict usage to trusted, controlled scenarios.
          </div>`
        : '';

    const chromeClass: string =
      (props.chromeDensity || 'Comfortable') === 'Compact'
        ? `${styles.chrome} ${styles.chromeCompact}`
        : styles.chrome;

    const dashboardHtml: string = showDashboardSelector
      ? this.buildDashboardSelectorHtml(props.dashboardList, currentDashboardId)
      : '';
    const reportBrowserHtml: string = this.buildReportBrowserHtml(props);

    return `
      <div class="${chromeClass}">
        <div class="${styles.chromeLeft}">
          <div class="${styles.chromeTitle}">${escape(title)}</div>
          ${subtitleHtml}
          ${anyHttpsWarningHtml}
        </div>
        <div class="${styles.chromeRight}">
          ${statusHtml}
          ${refreshHtml}
          ${exportHtml}
          ${importHtml}
          ${openInNewTabHtml}
        </div>
      </div>
      ${dashboardHtml}
      ${reportBrowserHtml}`;
  }
  private getOpenInNewTabUrl(
    resolvedUrl: string,
    baseUrl: string,
    pageUrl: string,
    props: IUniversalHtmlViewerWebPartProps,
  ): string {
    const contentDeliveryMode: ContentDeliveryMode = this.getContentDeliveryMode(props);
    return buildOpenInNewTabUrl({
      resolvedUrl,
      baseUrl,
      pageUrl,
      currentPageUrl: this.getCurrentPageUrl(),
      contentDeliveryMode,
    });
  }
  protected updateOpenInNewTabLink(
    baseUrl: string,
    pageUrl: string,
    props: IUniversalHtmlViewerWebPartProps,
  ): void {
    const openInNewTabLink: HTMLAnchorElement | null = this.domElement.querySelector(
      '[data-uhv-action="open-in-new-tab"]',
    );
    if (!openInNewTabLink) {
      return;
    }

    const currentHref: string = (openInNewTabLink.getAttribute('href') || '').trim();
    const resolvedUrlForFallback: string = currentHref || baseUrl;
    const updatedHref = this.getOpenInNewTabUrl(
      resolvedUrlForFallback,
      baseUrl,
      pageUrl,
      props,
    );
    openInNewTabLink.setAttribute('href', updatedHref);
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

  private buildReportBrowserHtml(props: IUniversalHtmlViewerWebPartProps): string {
    const contentDeliveryMode: ContentDeliveryMode = this.getContentDeliveryMode(props);
    if (
      !isReportBrowserSourceMode(props.htmlSourceMode) ||
      !isInlineContentDeliveryMode(contentDeliveryMode) ||
      props.showChrome === false
    ) {
      return '';
    }

    const defaultView = this.normalizeReportBrowserView(props.reportBrowserDefaultView);
    const currentView = this.reportBrowserView || defaultView;
    const folderButtonClass =
      currentView === 'Folders'
        ? `${styles.reportBrowserViewButton} ${styles.reportBrowserViewButtonActive}`
        : styles.reportBrowserViewButton;
    const filesButtonClass =
      currentView === 'Files'
        ? `${styles.reportBrowserViewButton} ${styles.reportBrowserViewButtonActive}`
        : styles.reportBrowserViewButton;

    return `
      <div class="${styles.reportBrowser}" data-uhv-report-browser>
        <div class="${styles.reportBrowserToolbar}">
          <div class="${styles.reportBrowserTitle}">Reports</div>
          <button class="${folderButtonClass}" type="button" data-uhv-report-view="Folders">Folders</button>
          <button class="${filesButtonClass}" type="button" data-uhv-report-view="Files">Files</button>
          <input class="${styles.reportBrowserSearch}" type="search" placeholder="Filter reports" data-uhv-report-filter />
        </div>
        <div class="${styles.reportBrowserStatus}" data-uhv-report-status>Loading reports...</div>
        <div class="${styles.reportBrowserList}" data-uhv-report-list></div>
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
        const activePageUrl: string = this.getCurrentPageUrl() || pageUrl;
        const activeCacheBusterMode: CacheBusterMode =
          this.lastCacheBusterMode || cacheBusterMode;
        this.refreshIframe(
          activeBaseUrl,
          activeCacheBusterMode,
          cacheBusterParamName,
          activePageUrl,
          false,
          false,
          true,
        ).catch(() => {
          return undefined;
        });
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
        ).catch(() => {
          return undefined;
        });
      });
    }
    if (dashboardFilter && dashboardSelect) {
      dashboardFilter.addEventListener('input', () => {
        this.filterDashboardOptions(dashboardFilter.value, dashboardSelect);
      });
    }

    this.attachReportBrowserHandlers(
      baseUrl,
      cacheBusterMode,
      cacheBusterParamName,
      pageUrl,
      effectiveProps,
    );

    const iframe: HTMLIFrameElement | null = this.domElement.querySelector('iframe');
    if (iframe) {
      iframe.addEventListener('load', () => {
        this.setLoadingVisible(false);
      });
    }
  }

  private attachReportBrowserHandlers(
    baseUrl: string,
    cacheBusterMode: CacheBusterMode,
    cacheBusterParamName: string,
    pageUrl: string,
    effectiveProps: IUniversalHtmlViewerWebPartProps,
  ): void {
    const browserElement: HTMLElement | null = this.domElement.querySelector(
      '[data-uhv-report-browser]',
    );
    const contentDeliveryMode: ContentDeliveryMode = this.getContentDeliveryMode(effectiveProps);
    if (
      !browserElement ||
      !isReportBrowserSourceMode(effectiveProps.htmlSourceMode) ||
      !isInlineContentDeliveryMode(contentDeliveryMode)
    ) {
      return;
    }

    const rootPath = this.getEffectiveReportBrowserRootPath(
      effectiveProps,
      baseUrl,
    );
    if (!rootPath) {
      this.updateReportBrowserStatus('Configure a report browser root folder.');
      return;
    }

    const defaultView = this.normalizeReportBrowserView(
      effectiveProps.reportBrowserDefaultView,
    );
    if (
      this.reportBrowserRootPath !== rootPath ||
      this.reportBrowserDefaultView !== defaultView
    ) {
      this.reportBrowserRootPath = rootPath;
      this.reportBrowserFolderPath = rootPath;
      this.reportBrowserView = defaultView;
      this.reportBrowserDefaultView = defaultView;
    }

    const viewButtons = this.domElement.querySelectorAll<HTMLButtonElement>(
      '[data-uhv-report-view]',
    );
    viewButtons.forEach((button) => {
      button.addEventListener('click', () => {
        const nextView = this.normalizeReportBrowserView(
          button.getAttribute('data-uhv-report-view') || undefined,
        );
        this.reportBrowserView = nextView;
        if (nextView === 'Files') {
          this.reportBrowserFolderPath = rootPath;
        }
        this.updateReportBrowserViewButtons(nextView);
        this.loadAndRenderReportBrowser(effectiveProps, rootPath).catch(() => {
          return undefined;
        });
      });
    });

    const filterInput: HTMLInputElement | null = this.domElement.querySelector(
      '[data-uhv-report-filter]',
    );
    if (filterInput) {
      filterInput.addEventListener('input', () => {
        this.filterReportBrowserRows(filterInput.value);
      });
    }

    const listElement: HTMLElement | null = this.domElement.querySelector(
      '[data-uhv-report-list]',
    );
    if (listElement) {
      listElement.addEventListener('click', (event) => {
        const target = event.target;
        if (!(target instanceof Element)) {
          return;
        }

        const folderButton = target.closest<HTMLButtonElement>(
          '[data-uhv-report-folder]',
        );
        if (folderButton) {
          const folderPath = folderButton.getAttribute('data-uhv-report-folder') || '';
          if (folderPath && isPathInsideRoot(folderPath, rootPath)) {
            this.reportBrowserFolderPath = folderPath;
            this.loadAndRenderReportBrowser(effectiveProps, rootPath).catch(() => {
              return undefined;
            });
          }
          return;
        }

        const fileButton = target.closest<HTMLButtonElement>('[data-uhv-report-file]');
        if (!fileButton) {
          return;
        }

        const filePath = fileButton.getAttribute('data-uhv-report-file') || '';
        if (!filePath) {
          return;
        }

        this.handleReportBrowserFileSelection(
          filePath,
          effectiveProps,
          pageUrl,
          cacheBusterMode,
          cacheBusterParamName,
        ).catch(() => {
          return undefined;
        });
      });
    }

    this.loadAndRenderReportBrowser(effectiveProps, rootPath).catch(() => {
      return undefined;
    });
  }

  private async loadAndRenderReportBrowser(
    props: IUniversalHtmlViewerWebPartProps,
    rootPath: string,
  ): Promise<void> {
    const view = this.reportBrowserView || this.normalizeReportBrowserView(
      props.reportBrowserDefaultView,
    );
    const currentFolderPath = this.reportBrowserFolderPath || rootPath;
    const requestId = this.reportBrowserLoadRequestId + 1;
    this.reportBrowserLoadRequestId = requestId;
    this.updateReportBrowserStatus('Loading reports...');

    try {
      const items = await loadSharePointReportBrowserItems({
        spHttpClient: this.context.spHttpClient,
        webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
        rootPath,
        currentFolderPath,
        allowedExtensions: this.parseFileExtensions(props.allowedFileExtensions),
        view,
        maxItems: props.reportBrowserMaxItems ?? 300,
        spHttpClientConfiguration: SPHttpClient.configurations.v1,
      });

      if (requestId !== this.reportBrowserLoadRequestId) {
        return;
      }

      this.renderReportBrowserItems(items, rootPath, currentFolderPath, view);
    } catch (error) {
      if (requestId !== this.reportBrowserLoadRequestId) {
        return;
      }

      const message = error instanceof Error ? error.message : 'Unable to load reports.';
      this.updateReportBrowserStatus(message);
      this.updateReportBrowserList('');
    }
  }

  private renderReportBrowserItems(
    items: ISharePointReportBrowserItem[],
    rootPath: string,
    currentFolderPath: string,
    view: ReportBrowserDefaultView,
  ): void {
    const rows: string[] = [];
    if (view === 'Folders' && currentFolderPath !== rootPath) {
      const parentPath = getReportBrowserParentFolderPath(rootPath, currentFolderPath);
      rows.push(this.buildReportBrowserFolderRow('..', parentPath, 'Parent folder'));
    }

    items.forEach((item) => {
      if (item.kind === 'Folder') {
        rows.push(
          this.buildReportBrowserFolderRow(
            item.name,
            item.serverRelativeUrl,
            item.relativePath || item.name,
          ),
        );
      } else {
        rows.push(this.buildReportBrowserFileRow(item, view));
      }
    });

    this.updateReportBrowserList(rows.join(''));
    const visibleItemCount = items.length;
    const scopeLabel =
      view === 'Files'
        ? 'all accessible files'
        : this.getReportBrowserFolderLabel(rootPath, currentFolderPath);
    this.updateReportBrowserStatus(`${visibleItemCount} item(s) in ${scopeLabel}`);
    this.filterReportBrowserRows(this.getReportBrowserFilterValue());
  }

  private buildReportBrowserFolderRow(
    label: string,
    folderPath: string,
    title: string,
  ): string {
    return `
      <button class="${styles.reportBrowserRow}" type="button" data-uhv-report-folder="${escape(folderPath)}" data-uhv-report-row>
        <span class="${styles.reportBrowserIcon}">Folder</span>
        <span class="${styles.reportBrowserName}">${escape(label)}</span>
        <span class="${styles.reportBrowserPath}">${escape(title)}</span>
      </button>`;
  }

  private buildReportBrowserFileRow(
    item: ISharePointReportBrowserItem,
    view: ReportBrowserDefaultView,
  ): string {
    const label = view === 'Files' && item.relativePath ? item.relativePath : item.name;
    const modifiedLabel = item.timeLastModified
      ? new Date(item.timeLastModified).toLocaleString()
      : '';
    return `
      <button class="${styles.reportBrowserRow}" type="button" data-uhv-report-file="${escape(item.serverRelativeUrl)}" data-uhv-report-row>
        <span class="${styles.reportBrowserIcon}">HTML</span>
        <span class="${styles.reportBrowserName}">${escape(label)}</span>
        <span class="${styles.reportBrowserPath}">${escape(modifiedLabel || item.relativePath)}</span>
      </button>`;
  }

  private updateReportBrowserViewButtons(view: ReportBrowserDefaultView): void {
    const viewButtons = this.domElement.querySelectorAll<HTMLButtonElement>(
      '[data-uhv-report-view]',
    );
    viewButtons.forEach((button) => {
      const buttonView = this.normalizeReportBrowserView(
        button.getAttribute('data-uhv-report-view') || undefined,
      );
      if (buttonView === view) {
        button.classList.add(styles.reportBrowserViewButtonActive);
      } else {
        button.classList.remove(styles.reportBrowserViewButtonActive);
      }
    });
  }

  private filterReportBrowserRows(filterValue: string): void {
    const normalizedFilter = filterValue.trim().toLowerCase();
    const rows = this.domElement.querySelectorAll<HTMLElement>('[data-uhv-report-row]');
    rows.forEach((row) => {
      const text = row.textContent || '';
      row.style.display =
        !normalizedFilter || text.toLowerCase().includes(normalizedFilter) ? '' : 'none';
    });
  }

  private getReportBrowserFilterValue(): string {
    const filterInput = this.domElement.querySelector<HTMLInputElement>(
      '[data-uhv-report-filter]',
    );
    return filterInput?.value || '';
  }

  private async handleReportBrowserFileSelection(
    filePath: string,
    props: IUniversalHtmlViewerWebPartProps,
    pageUrl: string,
    cacheBusterMode: CacheBusterMode,
    cacheBusterParamName: string,
  ): Promise<void> {
    const currentPageUrl: string = this.getCurrentPageUrl() || pageUrl;
    const validationOptions = this.buildUrlValidationOptions(currentPageUrl, props);
    this.lastValidationOptions = validationOptions;
    if (!isUrlAllowed(filePath, validationOptions)) {
      this.updateReportBrowserStatus('Selected report is not allowed by UHV security settings.');
      return;
    }

    this.lastCacheBusterMode = cacheBusterMode;
    this.setLoadingVisible(true);
    this.currentBaseUrl = filePath;
    this.onNavigatedToUrl(filePath, currentPageUrl);
    const updatedPageUrl: string = this.getCurrentPageUrl() || currentPageUrl;
    this.updateOpenInNewTabLink(filePath, updatedPageUrl, props);
    this.setupIframeLoadFallback(filePath, props);
    await this.refreshIframe(
      filePath,
      cacheBusterMode,
      cacheBusterParamName,
      updatedPageUrl,
      true,
      true,
    );
    this.setupAutoRefresh(filePath, cacheBusterMode, cacheBusterParamName, updatedPageUrl, props);
    this.updateStatusBadge(validationOptions, cacheBusterMode, props);
  }

  private getEffectiveReportBrowserRootPath(
    props: IUniversalHtmlViewerWebPartProps,
    baseUrl: string,
  ): string {
    const configuredRootPath = (props.reportBrowserRootPath || '').trim();
    if (configuredRootPath) {
      return normalizeSharePointReportBrowserRootPath(
        configuredRootPath,
        this.context.pageContext.web.absoluteUrl,
      );
    }

    if (props.basePath) {
      const rootFromBasePath = normalizeSharePointReportBrowserRootPath(
        props.basePath,
        this.context.pageContext.web.absoluteUrl,
      );
      if (rootFromBasePath) {
        return rootFromBasePath;
      }
    }

    const sourceDirectory = this.getDirectoryPathFromUrl(baseUrl);
    return normalizeSharePointReportBrowserRootPath(
      sourceDirectory,
      this.context.pageContext.web.absoluteUrl,
    );
  }

  private getDirectoryPathFromUrl(value: string): string {
    const trimmed = (value || '').trim();
    if (!trimmed) {
      return '';
    }

    let path: string;
    try {
      path = new URL(trimmed, this.getCurrentPageUrl()).pathname;
    } catch {
      path = trimmed;
    }

    const queryIndex = path.indexOf('?');
    if (queryIndex !== -1) {
      path = path.substring(0, queryIndex);
    }
    const hashIndex = path.indexOf('#');
    if (hashIndex !== -1) {
      path = path.substring(0, hashIndex);
    }

    const lastSlashIndex = path.lastIndexOf('/');
    return lastSlashIndex <= 0 ? path : path.substring(0, lastSlashIndex);
  }

  private getReportBrowserFolderLabel(rootPath: string, currentFolderPath: string): string {
    if (rootPath === currentFolderPath) {
      return 'root folder';
    }

    const relativePath = currentFolderPath.substring(rootPath.length).replace(/^\/+/, '');
    return relativePath || 'root folder';
  }

  private normalizeReportBrowserView(
    value?: string,
  ): ReportBrowserDefaultView {
    return value === 'Files' ? 'Files' : 'Folders';
  }

  private updateReportBrowserStatus(message: string): void {
    const statusElement: HTMLElement | null = this.domElement.querySelector(
      '[data-uhv-report-status]',
    );
    if (statusElement) {
      statusElement.textContent = message;
    }
  }

  private updateReportBrowserList(html: string): void {
    const listElement: HTMLElement | null = this.domElement.querySelector(
      '[data-uhv-report-list]',
    );
    if (listElement) {
      listElement.innerHTML = html;
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

    const currentPageUrl: string = this.getCurrentPageUrl() || pageUrl;
    const validationOptions = this.buildUrlValidationOptions(currentPageUrl, props);
    this.lastValidationOptions = validationOptions;
    if (!isUrlAllowed(url, validationOptions)) {
      return;
    }

    const cacheBusterMode: CacheBusterMode = props.cacheBusterMode || 'None';
    this.lastCacheBusterMode = cacheBusterMode;

    this.setLoadingVisible(true);
    this.currentBaseUrl = url;
    this.onNavigatedToUrl(url, currentPageUrl);
    const updatedPageUrl: string = this.getCurrentPageUrl() || currentPageUrl;
    this.updateOpenInNewTabLink(url, updatedPageUrl, props);
    this.setupIframeLoadFallback(url, props);
    await this.refreshIframe(
      url,
      cacheBusterMode,
      cacheBusterParamName,
      updatedPageUrl,
      true,
      true,
    );
    this.setupAutoRefresh(url, cacheBusterMode, cacheBusterParamName, updatedPageUrl, props);
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
    const exportData = buildConfigExport(props);
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
    const importResult: IConfigImportResult = applyImportedConfigToProps(propsRecord, config);
    if (importResult.ignoredEntries.length > 0) {
      // eslint-disable-next-line no-console
      console.warn(
        'UniversalHtmlViewer: ignored invalid configuration entries during import.',
        importResult.ignoredEntries,
      );
    }

    const preset: ConfigurationPreset = this.properties.configurationPreset || 'Custom';
    if (this.properties.lockPresetSettings && preset !== 'Custom') {
      this.applyPreset(preset);
    }

    this.context.propertyPane.refresh();
    this.render();
  }
}
