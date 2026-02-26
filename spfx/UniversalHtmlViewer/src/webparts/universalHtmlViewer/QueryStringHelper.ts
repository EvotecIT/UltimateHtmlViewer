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

    if (value === null) {
      return undefined;
    }

    return value;
  } catch {
    const questionMarkIndex: number = url.indexOf('?');

    if (questionMarkIndex === -1 || questionMarkIndex === url.length - 1) {
      return undefined;
    }

    const queryPartWithFragment: string = url.substring(questionMarkIndex + 1);
    const hashIndex: number = queryPartWithFragment.indexOf('#');
    const queryPart: string =
      hashIndex === -1
        ? queryPartWithFragment
        : queryPartWithFragment.substring(0, hashIndex);
    const pairs: string[] = queryPart.split('&');

    for (const pair of pairs) {
      if (!pair) {
        continue;
      }

      const equalsIndex: number = pair.indexOf('=');
      const key: string = equalsIndex === -1 ? pair : pair.substring(0, equalsIndex);
      const value: string = equalsIndex === -1 ? '' : pair.substring(equalsIndex + 1);
      const decodedKey: string | undefined = tryDecodeQueryComponent(key);
      if (!decodedKey) {
        continue;
      }

      if (decodedKey === paramName) {
        const decodedValue: string | undefined = tryDecodeQueryComponent(value);
        if (decodedValue !== undefined) {
          return decodedValue;
        }
        continue;
      }
    }

    return undefined;
  }
}

function tryDecodeQueryComponent(value?: string): string | undefined {
  if (value === undefined) {
    return undefined;
  }

  try {
    return decodeURIComponent(value.replace(/\+/g, '%20'));
  } catch {
    return undefined;
  }
}
