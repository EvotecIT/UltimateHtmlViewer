/**
 * Returns the value of a query string parameter from a URL.
 *
 * @param url The URL to parse.
 * @param paramName The name of the query string parameter.
 */
export function getQueryStringParam(url: string, paramName: string): string | undefined {
  if (!url || !paramName) {
    return undefined;
  }

  try {
    const parsed: URL = new URL(url);
    const value: string | null = parsed.searchParams.get(paramName);

    if (!value) {
      return undefined;
    }

    return value;
  } catch {
    const questionMarkIndex: number = url.indexOf('?');

    if (questionMarkIndex === -1 || questionMarkIndex === url.length - 1) {
      return undefined;
    }

    const queryPart: string = url.substring(questionMarkIndex + 1);
    const pairs: string[] = queryPart.split('&');

    for (const pair of pairs) {
      const [key, value] = pair.split('=');

      if (decodeURIComponent(key) === paramName) {
        return value ? decodeURIComponent(value) : undefined;
      }
    }

    return undefined;
  }
}
