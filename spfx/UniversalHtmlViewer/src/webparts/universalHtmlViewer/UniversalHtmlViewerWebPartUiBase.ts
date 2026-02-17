import { escape } from '@microsoft/sp-lodash-subset';
import styles from './UniversalHtmlViewerWebPart.module.scss';
import {
  applyImportedConfigToProps,
  buildConfigExport,
  IConfigImportResult,
} from './ConfigImportExportHelper';
import { CacheBusterMode, HeightMode, isUrlAllowed, UrlValidationOptions } from './UrlHelper';
import {
  ConfigurationPreset,
  IUniversalHtmlViewerWebPartProps,
} from './UniversalHtmlViewerTypes';
import { UniversalHtmlViewerWebPartRuntimeBase } from './UniversalHtmlViewerWebPartRuntimeBase';

export abstract class UniversalHtmlViewerWebPartUiBase extends UniversalHtmlViewerWebPartRuntimeBase {
  protected getIframeHeightStyle(props: IUniversalHtmlViewerWebPartProps): string {
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

    const iframe: HTMLIFrameElement | null = this.domElement.querySelector('iframe');
    if (iframe) {
      iframe.addEventListener('load', () => {
        this.setLoadingVisible(false);
      });
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
