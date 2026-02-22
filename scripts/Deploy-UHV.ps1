[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$AppCatalogUrl,

    [string]$SppkgPath = (Join-Path (Resolve-Path (Join-Path $PSScriptRoot "..")) "spfx/UniversalHtmlViewer/sharepoint/solution/universal-html-viewer.sppkg"),

    [ValidateSet("Tenant", "Site")]
    [string]$Scope = "Tenant",

    [switch]$TenantWide,

    # PnP.PowerShell 3.x requires a ClientId for Interactive/DeviceLogin auth.
    [string]$ClientId,

    # Use tenant name (e.g. contoso.onmicrosoft.com) or tenant id (GUID).
    [string]$Tenant,

    # More reliable than -Interactive in headless/locked-down environments.
    [switch]$DeviceLogin,

    # Reuse cached login across script runs (recommended for repeated operations).
    [bool]$PersistLogin = $true,

    # Force a fresh sign-in prompt even if cached login exists.
    [switch]$ForceAuthentication,

    # Optional explicit admin URL, e.g. https://contoso-admin.sharepoint.com
    [string]$TenantAdminUrl
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not (Test-Path $SppkgPath)) {
    throw "Package not found at $SppkgPath. Run Build-UHV.ps1 first."
}

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    throw "PnP.PowerShell module not found. Install-Module PnP.PowerShell -Scope CurrentUser"
}

$connectUrl = $AppCatalogUrl
if ($Scope -eq "Tenant") {
    if ($TenantAdminUrl) {
        $connectUrl = $TenantAdminUrl
    } elseif ($AppCatalogUrl -match '^https://([^.]+)\.sharepoint\.com/?') {
        $connectUrl = "https://$($Matches[1])-admin.sharepoint.com"
    }
}

if ($Scope -eq "Site" -and $TenantWide.IsPresent) {
    throw "-TenantWide is only valid with -Scope Tenant."
}

Write-Host "Connecting to:" $connectUrl
if ($DeviceLogin.IsPresent) {
    if (-not $ClientId) {
        throw "ClientId is required for -DeviceLogin. Create an Entra ID app registration and pass -ClientId <GUID> (see docs/Deploy-SharePointOnline.md)."
    }
    if (-not $Tenant) {
        throw "Tenant is required for -DeviceLogin. Pass -Tenant <tenant>.onmicrosoft.com or the tenant GUID."
    }
    if ($ForceAuthentication.IsPresent) {
        Write-Warning "-ForceAuthentication is only supported with -Interactive login; ignoring it for -DeviceLogin."
    }
    Connect-PnPOnline -Url $connectUrl -DeviceLogin -ClientId $ClientId -Tenant $Tenant -PersistLogin:$PersistLogin
} else {
    if (-not $ClientId) {
        throw "ClientId is required for -Interactive with PnP.PowerShell 3.x. Create an Entra ID app registration and pass -ClientId <GUID> (see docs/Deploy-SharePointOnline.md)."
    }

    if ($Tenant) {
        if ($TenantAdminUrl) {
            Connect-PnPOnline -Url $connectUrl -Interactive -ClientId $ClientId -Tenant $Tenant -TenantAdminUrl $TenantAdminUrl -PersistLogin:$PersistLogin -ForceAuthentication:$ForceAuthentication.IsPresent
        } else {
            Connect-PnPOnline -Url $connectUrl -Interactive -ClientId $ClientId -Tenant $Tenant -PersistLogin:$PersistLogin -ForceAuthentication:$ForceAuthentication.IsPresent
        }
    } else {
        if ($TenantAdminUrl) {
            Connect-PnPOnline -Url $connectUrl -Interactive -ClientId $ClientId -TenantAdminUrl $TenantAdminUrl -PersistLogin:$PersistLogin -ForceAuthentication:$ForceAuthentication.IsPresent
        } else {
            Connect-PnPOnline -Url $connectUrl -Interactive -ClientId $ClientId -PersistLogin:$PersistLogin -ForceAuthentication:$ForceAuthentication.IsPresent
        }
    }
}

if ($Scope -eq "Tenant") {
    $tenantCatalogUrl = Get-PnPTenantAppCatalogUrl
    if ([string]::IsNullOrWhiteSpace($tenantCatalogUrl)) {
        throw @"
No tenant app catalog is configured for this tenant.

Run (while connected to https://<tenant>-admin.sharepoint.com):
  Register-PnPAppCatalogSite -Url '$AppCatalogUrl' -Owner '<admin@tenant>' -TimeZoneId <id>

Or if the app catalog site already exists:
  Set-PnPTenantAppCatalogUrl -Url '$AppCatalogUrl'
"@
    }

    if ($tenantCatalogUrl.TrimEnd('/') -ne $AppCatalogUrl.TrimEnd('/')) {
        Write-Warning "Tenant app catalog is currently '$tenantCatalogUrl', but -AppCatalogUrl was '$AppCatalogUrl'. Using tenant configured URL."
    }
} else {
    $siteCatalog = Get-PnPSiteCollectionAppCatalog -CurrentSite
    if (-not $siteCatalog) {
        throw "No site collection app catalog is enabled on '$AppCatalogUrl'. Enable it first or deploy with -Scope Tenant."
    }
}

Write-Host "Uploading package:" $SppkgPath
$app = Add-PnPApp -Path $SppkgPath -Overwrite -Scope $Scope -Force

Write-Host "Publishing package (Scope=$Scope, TenantWide=$($TenantWide.IsPresent))"
if ($Scope -eq "Tenant" -and $TenantWide.IsPresent) {
    try {
        Publish-PnPApp -Identity $app.Id -Scope $Scope -SkipFeatureDeployment -Force
    } catch {
        $errorText = $_.Exception.Message
        if ($errorText -match "SkipFeatureDeployment set to true") {
            Write-Warning "Package does not support tenant-wide deployment (skipFeatureDeployment=false). Publishing without -SkipFeatureDeployment."
            Publish-PnPApp -Identity $app.Id -Scope $Scope -Force
        } else {
            throw
        }
    }
} else {
    Publish-PnPApp -Identity $app.Id -Scope $Scope -Force
}

Write-Host "Deployment completed."
