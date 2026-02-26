import {
  validateAllowedFileExtensions,
  validateAllowedHosts,
  validateAllowedPathPrefixes,
  validateBasePath,
  validateFullUrl,
  validateTenantConfigUrl,
} from '../ValidationHelper';

describe('ValidationHelper', () => {
  describe('validateFullUrl', () => {
    it('rejects unsafe schemes', () => {
      const javaScriptUrl = ['java', 'script:alert(1)'].join('');
      expect(validateFullUrl(javaScriptUrl)).toBe('Unsupported or unsafe URL scheme.');
      expect(validateFullUrl('data:text/html;base64,AAA')).toBe('Unsupported or unsafe URL scheme.');
    });

    it('rejects http unless explicitly enabled', () => {
      expect(validateFullUrl('http://contoso.sharepoint.com/sites/reports/index.html')).toBe(
        'HTTP is blocked by default. Enable "Allow HTTP" if required.',
      );
      expect(
        validateFullUrl('http://contoso.sharepoint.com/sites/reports/index.html', true),
      ).toBe('');
    });
  });

  describe('validateBasePath', () => {
    it('enforces site-relative structure', () => {
      expect(validateBasePath('sites/reports')).toBe('Base path must start with "/".');
      expect(validateBasePath('https://contoso.sharepoint.com/sites/reports')).toBe(
        'Base path must be site-relative, e.g. /sites/Reports/Dashboards/.',
      );
    });

    it('rejects traversal-like segments', () => {
      expect(validateBasePath('/sites/reports/../secret/')).toBe(
        'Base path must not include "." or ".." segments.',
      );
    });

    it('rejects encoded and backslash traversal-like segments', () => {
      expect(validateBasePath('/sites/reports/%2e%2e/secret/')).toBe(
        'Base path must not include "." or ".." segments.',
      );
      expect(validateBasePath('/sites/reports/..\\secret/')).toBe(
        'Base path must not include "." or ".." segments.',
      );
      expect(validateBasePath('/sites/reports/%2e%2e%5Csecret/')).toBe(
        'Base path must not include "." or ".." segments.',
      );
    });
  });

  describe('validateAllowedHosts', () => {
    it('accepts standard and wildcard host entries', () => {
      expect(validateAllowedHosts('cdn.contoso.com, *.contoso.net')).toBe('');
    });

    it('rejects invalid host values', () => {
      expect(validateAllowedHosts('cdn[dot]contoso.com')).toContain('Invalid host entry');
    });
  });

  describe('validateAllowedPathPrefixes', () => {
    it('enforces site-relative prefixes without query strings', () => {
      expect(validateAllowedPathPrefixes('/sites/reports/dashboard/')).toBe('');
      expect(validateAllowedPathPrefixes('sites/reports/dashboard/')).toContain(
        'Path prefixes must start with "/"',
      );
      expect(validateAllowedPathPrefixes('/sites/reports/dashboard/?foo=1')).toContain(
        'must not include query strings',
      );
    });

    it('rejects encoded and backslash traversal-like segments in path prefixes', () => {
      expect(validateAllowedPathPrefixes('/sites/reports/%2e%2e/secret/')).toContain(
        'must not include "." or ".."',
      );
      expect(validateAllowedPathPrefixes('/sites/reports/..\\secret/')).toContain(
        'must not include "." or ".."',
      );
      expect(validateAllowedPathPrefixes('/sites/reports/%2e%2e%2Fsecret/')).toContain(
        'must not include "." or ".."',
      );
    });
  });

  describe('validateAllowedFileExtensions', () => {
    it('accepts common extensions and rejects malformed values', () => {
      expect(validateAllowedFileExtensions('.html,htm,.aspx')).toBe('');
      expect(validateAllowedFileExtensions('.ht*m')).toContain('Invalid extension');
    });
  });

  describe('validateTenantConfigUrl', () => {
    const currentPageUrl =
      'https://contoso.sharepoint.com/sites/Reports/SitePages/Dashboard.aspx';

    it('allows same-host https and site-relative values', () => {
      expect(
        validateTenantConfigUrl(
          'https://contoso.sharepoint.com/sites/Reports/SiteAssets/uhv-config.json',
          currentPageUrl,
        ),
      ).toBe('');
      expect(validateTenantConfigUrl('/sites/Reports/SiteAssets/uhv-config.json', currentPageUrl)).toBe(
        '',
      );
    });

    it('rejects http and cross-host tenant config URLs', () => {
      expect(
        validateTenantConfigUrl(
          'http://contoso.sharepoint.com/sites/Reports/SiteAssets/uhv-config.json',
          currentPageUrl,
        ),
      ).toBe('Tenant config should use HTTPS.');
      expect(
        validateTenantConfigUrl(
          'https://fabrikam.sharepoint.com/sites/Reports/SiteAssets/uhv-config.json',
          currentPageUrl,
        ),
      ).toBe('Tenant config must be hosted in the same SharePoint tenant.');
    });
  });
});
