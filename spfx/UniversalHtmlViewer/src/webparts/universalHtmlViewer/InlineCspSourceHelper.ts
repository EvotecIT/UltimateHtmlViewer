export function appendAdditionalCspHostSources(baseSources: string, hosts?: string[]): string {
  const additionalSources = normalizeAdditionalCspHostSources(hosts);
  if (additionalSources.length === 0) {
    return baseSources;
  }

  return `${baseSources} ${additionalSources.join(' ')}`;
}

function normalizeAdditionalCspHostSources(hosts?: string[]): string[] {
  if (!hosts || hosts.length === 0) {
    return [];
  }

  const sourceSet = new Set<string>();
  hosts
    .join(',')
    .split(/[,;\s]+/g)
    .map((entry) => entry.trim())
    .filter((entry) => entry.length > 0)
    .forEach((entry) => {
      const source = normalizeAdditionalCspHostSource(entry);
      if (source) {
        sourceSet.add(source);
      }
    });

  return Array.from(sourceSet);
}

function normalizeAdditionalCspHostSource(entry: string): string | undefined {
  const lowered = entry.toLowerCase();
  if (lowered === "'self'" || lowered === 'data:' || lowered === 'blob:') {
    return lowered;
  }

  try {
    if (/^https?:\/\//i.test(entry)) {
      const url = new URL(entry);
      if (url.protocol !== 'https:') {
        return undefined;
      }
      return url.origin.toLowerCase();
    }
  } catch {
    return undefined;
  }

  if (lowered.startsWith('*.')) {
    return `https://${lowered}`;
  }

  if (lowered.startsWith('.')) {
    return `https://*${lowered}`;
  }

  if (/^[a-z0-9.-]+$/i.test(lowered)) {
    return `https://${lowered}`;
  }

  return undefined;
}
