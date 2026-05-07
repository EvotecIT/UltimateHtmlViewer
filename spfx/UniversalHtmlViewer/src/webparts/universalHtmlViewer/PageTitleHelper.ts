export function extractTitleFromHtml(html?: string): string {
  const value = (html || '').trim();
  if (!value) {
    return '';
  }

  if (typeof DOMParser !== 'undefined') {
    try {
      const parsed = new DOMParser().parseFromString(value, 'text/html');
      const title = normalizePageTitle(parsed?.querySelector('title')?.textContent || '');
      if (title) {
        return title;
      }
    } catch {
      return '';
    }
  }

  const titleMatch = value.match(/<title\b[^>]*>([\s\S]*?)<\/title>/i);
  if (!titleMatch) {
    return '';
  }

  return normalizePageTitle(decodeBasicHtmlEntities(titleMatch[1] || ''));
}

export function normalizePageTitle(value?: string): string {
  return (value || '').replace(/\s+/g, ' ').trim();
}

function decodeBasicHtmlEntities(value: string): string {
  return value
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'");
}
