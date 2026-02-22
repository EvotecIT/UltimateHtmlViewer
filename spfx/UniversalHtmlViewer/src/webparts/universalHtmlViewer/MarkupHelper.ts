import { escape } from '@microsoft/sp-lodash-subset';

export function buildMessageHtml(
  message: string,
  extraHtml: string | undefined,
  viewerClassName: string,
  messageClassName: string,
): string {
  const extra: string = extraHtml ? extraHtml : '';
  return `<div class="${viewerClassName}"><div class="${messageClassName}">${escape(message)}</div></div>${extra}`;
}

export function buildOpenInNewTabHtml(
  url: string,
  fallbackClassName: string,
  fallbackLinkClassName: string,
  linkText?: string,
): string {
  return buildActionLinkHtml(
    url,
    fallbackClassName,
    fallbackLinkClassName,
    linkText || 'Open in new tab',
    true,
  );
}

export function buildActionLinkHtml(
  url: string,
  fallbackClassName: string,
  fallbackLinkClassName: string,
  linkText: string,
  openInNewTab?: boolean,
): string {
  const escapedUrl: string = escape(url);
  const label: string = escape((linkText || '').trim() || 'Open');
  const targetAttributes = openInNewTab
    ? ' target="_blank" rel="noopener noreferrer"'
    : '';
  return `<div class="${fallbackClassName}"><a class="${fallbackLinkClassName}" href="${escapedUrl}"${targetAttributes}>${label}</a></div>`;
}
