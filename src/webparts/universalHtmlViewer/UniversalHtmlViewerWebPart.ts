import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField,
  PropertyPaneSlider,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './UniversalHtmlViewerWebPart.module.scss';
import {
  buildFinalUrl,
  isUrlAllowed,
  HtmlSourceMode,
  HeightMode,
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
}

export default class UniversalHtmlViewerWebPart extends BaseClientSideWebPart<IUniversalHtmlViewerWebPartProps> {
  public render(): void {
    const pageUrl: string = this.getCurrentPageUrl();
    const htmlSourceMode: HtmlSourceMode = this.properties.htmlSourceMode || 'FullUrl';

    const finalUrl: string | null = buildFinalUrl({
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
      this.domElement.innerHTML = this.buildMessageHtml(
        'UniversalHtmlViewer: No URL configured. Please update the web part settings.',
      );
      return;
    }

    if (!isUrlAllowed(finalUrl, pageUrl)) {
      this.domElement.innerHTML = this.buildMessageHtml(
        'UniversalHtmlViewer: The target URL is invalid or not allowed.',
      );
      return;
    }

    const iframeHeightStyle: string = this.getIframeHeightStyle();

    this.domElement.innerHTML = `
      <div class="${styles.universalHtmlViewer}">
        <iframe class="${styles.iframe}"
          src="${escape(finalUrl)}"
          style="${iframeHeightStyle}border:0;"
          width="100%"
          frameborder="0"
        ></iframe>
      </div>`;
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

