# Copy this file to: .\ignore\UHV.LocalProfile.ps1
# Then update values for your tenant.

$env:UHV_CLIENT_ID = "<entra-app-client-id-guid>"
$env:UHV_TENANT = "<tenant>.onmicrosoft.com"

$Global:UhvTenantName = "<tenant>"
$Global:UhvAppCatalogUrl = "https://<tenant>.sharepoint.com/sites/appcatalog"
$Global:UhvTenantAdminUrl = "https://<tenant>-admin.sharepoint.com"

# Optional helper values for frequent site operations
$Global:UhvSiteUrl = "https://<tenant>.sharepoint.com/sites/Reports"
$Global:UhvDashboardPath = "SiteAssets/Index.html"
