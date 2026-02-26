[CmdletBinding()]
param(
    [string]$ClientId,
    [string]$Tenant,
    [string]$AppCatalogUrl,
    [string]$TenantAdminUrl,
    [string[]]$SiteUrls = @(),
    [switch]$DeviceLogin,
    [bool]$PersistLogin = $true,
    [switch]$ForceAuthentication,
    [switch]$SkipBuild,
    [switch]$SkipSiteUpdate,
    [switch]$ForceBootstrap
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Test-IsHttpsAbsoluteUrl {
    param(
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $false
    }

    $parsed = $null
    if (-not [Uri]::TryCreate($Value, [UriKind]::Absolute, [ref]$parsed)) {
        return $false
    }

    return $parsed.Scheme -eq "https"
}

function ConvertTo-HttpsUri {
    param(
        [string]$Value,
        [string]$ParameterName
    )

    if (-not (Test-IsHttpsAbsoluteUrl -Value $Value)) {
        throw "$ParameterName must be an absolute HTTPS URL."
    }

    return [Uri]$Value
}

if ([string]::IsNullOrWhiteSpace($ClientId)) {
    $ClientId = $env:UHV_CLIENT_ID
}
if ([string]::IsNullOrWhiteSpace($Tenant)) {
    $Tenant = $env:UHV_TENANT
}

if ([string]::IsNullOrWhiteSpace($ClientId)) {
    throw "ClientId is required. Pass -ClientId or set UHV_CLIENT_ID."
}
if ([string]::IsNullOrWhiteSpace($Tenant)) {
    throw "Tenant is required. Pass -Tenant or set UHV_TENANT."
}
if ([string]::IsNullOrWhiteSpace($AppCatalogUrl)) {
    throw "AppCatalogUrl is required. Pass -AppCatalogUrl 'https://<tenant>.sharepoint.com/sites/appcatalog'."
}
if ([string]::IsNullOrWhiteSpace($TenantAdminUrl)) {
    throw "TenantAdminUrl is required. Pass -TenantAdminUrl 'https://<tenant>-admin.sharepoint.com'."
}
if (-not $SkipSiteUpdate.IsPresent -and (-not $SiteUrls -or $SiteUrls.Count -eq 0)) {
    throw "SiteUrls are required unless -SkipSiteUpdate is used."
}

$appCatalogUri = ConvertTo-HttpsUri -Value $AppCatalogUrl -ParameterName "AppCatalogUrl"
$tenantAdminUri = ConvertTo-HttpsUri -Value $TenantAdminUrl -ParameterName "TenantAdminUrl"
$appCatalogHost = $appCatalogUri.Host.ToLowerInvariant()
$tenantAdminHost = $tenantAdminUri.Host.ToLowerInvariant()

$expectedTenantAdminHost = ""
$appCatalogHostLabels = $appCatalogHost.Split('.')
if ($appCatalogHostLabels.Length -ge 2) {
    $tenantLabel = $appCatalogHostLabels[0]
    if ($tenantLabel.EndsWith("-admin", [System.StringComparison]::OrdinalIgnoreCase)) {
        $tenantLabel = $tenantLabel.Substring(0, $tenantLabel.Length - 6)
    }

    if (-not [string]::IsNullOrWhiteSpace($tenantLabel)) {
        $suffix = ($appCatalogHostLabels[1..($appCatalogHostLabels.Length - 1)] -join ".")
        $expectedTenantAdminHost = "$tenantLabel-admin.$suffix"
    }
}
if ($expectedTenantAdminHost -and $tenantAdminHost -ne $expectedTenantAdminHost) {
    throw "TenantAdminUrl host '$tenantAdminHost' does not match AppCatalogUrl tenant host '$appCatalogHost'. Expected: '$expectedTenantAdminHost'."
}

if (-not $SkipSiteUpdate.IsPresent) {
    foreach ($siteUrl in $SiteUrls) {
        $siteUri = ConvertTo-HttpsUri -Value $siteUrl -ParameterName "SiteUrls"
        $siteHost = $siteUri.Host.ToLowerInvariant()
        if ($siteHost -ne $appCatalogHost) {
            throw "Each SiteUrls host must match AppCatalogUrl host '$appCatalogHost'. Invalid value: $siteUrl"
        }
    }
}

$scriptRoot = $PSScriptRoot
if (-not $scriptRoot) {
    $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}
if (-not $scriptRoot) {
    throw "Unable to determine script root."
}

Write-Host "Starting UHV deploy pipeline..."
Write-Host "App catalog: $AppCatalogUrl"
if ($SkipSiteUpdate.IsPresent) {
    Write-Host "Target sites: (skipped)"
} else {
    Write-Host "Target sites:"
    $SiteUrls | ForEach-Object { Write-Host " - $_" }
}

if (-not $SkipBuild.IsPresent) {
    & (Join-Path $scriptRoot "Build-UHV.ps1") -ForceBootstrap:$ForceBootstrap.IsPresent
}

& (Join-Path $scriptRoot "Deploy-UHV-Wrapper.ps1") `
    -AppCatalogUrl $AppCatalogUrl `
    -Scope Tenant `
    -ClientId $ClientId `
    -Tenant $Tenant `
    -DeviceLogin:$DeviceLogin.IsPresent `
    -PersistLogin $PersistLogin `
    -ForceAuthentication:$ForceAuthentication.IsPresent `
    -TenantAdminUrl $TenantAdminUrl `
    -SkipBuild

if (-not $SkipSiteUpdate.IsPresent) {
    & (Join-Path $scriptRoot "Update-UHVSiteApp.ps1") `
        -SiteUrls $SiteUrls `
        -InstallIfMissing `
        -ClientId $ClientId `
        -Tenant $Tenant `
        -DeviceLogin:$DeviceLogin.IsPresent `
        -PersistLogin $PersistLogin `
        -ForceAuthentication:$ForceAuthentication.IsPresent
}

Write-Host "UHV deploy pipeline completed."
