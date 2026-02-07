export function validateFullUrl(value?: string, allowHttp?: boolean): string {
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
    return allowHttp
      ? ''
      : 'HTTP is blocked by default. Enable "Allow HTTP" if required.';
  }

  return 'Enter a site-relative path or an absolute http/https URL.';
}

export function validateBasePath(value?: string): string {
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
  if (hasDotSegments(trimmed)) {
    return 'Base path must not include "." or ".." segments.';
  }
  return '';
}

export function validateAllowedHosts(value?: string): string {
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

export function validateAllowedPathPrefixes(value?: string): string {
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
    if (hasDotSegments(entry)) {
      return `Path prefixes must not include "." or "..": "${entry}".`;
    }
  }

  return '';
}

export function validateAllowedFileExtensions(value?: string): string {
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

export function validateTenantConfigUrl(value?: string, currentPageUrl?: string): string {
  const trimmed: string = (value || '').trim();
  if (!trimmed) {
    return '';
  }

  if (trimmed.startsWith('http://')) {
    return 'Tenant config should use HTTPS.';
  }

  if (trimmed.startsWith('https://')) {
    try {
      const target = new URL(trimmed);
      const current = new URL(currentPageUrl || 'https://invalid.local/');
      if (target.hostname.toLowerCase() !== current.hostname.toLowerCase()) {
        return 'Tenant config must be hosted in the same SharePoint tenant.';
      }
      return '';
    } catch {
      return 'Invalid tenant config URL.';
    }
  }

  if (trimmed.includes('://')) {
    return 'Tenant config must be site-relative or an absolute HTTPS URL.';
  }

  return '';
}

function hasDotSegments(pathname: string): boolean {
  const segments = pathname.split('/').filter((segment) => segment.length > 0);
  return segments.some((segment) => segment === '.' || segment === '..');
}
