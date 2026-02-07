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
): string {
  const escapedUrl: string = escape(url);
  return `<div class="${fallbackClassName}"><a class="${fallbackLinkClassName}" href="${escapedUrl}" target="_blank" rel="noopener noreferrer">Open in new tab</a></div>`;
}
