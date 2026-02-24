[CmdletBinding()]
param(
    [string]$ClientId,
    [string]$Tenant,
    [string]$AppCatalogUrl = "https://evotecpoland.sharepoint.com/sites/appcatalog",
    [string]$TenantAdminUrl = "https://evotecpoland-admin.sharepoint.com",
    [string[]]$SiteUrls = @(
        "https://evotecpoland.sharepoint.com/sites/TestUHV1",
        "https://evotecpoland.sharepoint.com/sites/TestUHV2",
        "https://evotecpoland.sharepoint.com"
    ),
    [switch]$DeviceLogin,
    [bool]$PersistLogin = $true,
    [switch]$ForceAuthentication,
    [switch]$SkipBuild,
    [switch]$SkipSiteUpdate,
    [switch]$ForceBootstrap
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

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

$scriptRoot = $PSScriptRoot
if (-not $scriptRoot) {
    $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}
if (-not $scriptRoot) {
    throw "Unable to determine script root."
}

Write-Host "Starting UHV deploy pipeline..."
Write-Host "App catalog: $AppCatalogUrl"
Write-Host "Target sites:"
$SiteUrls | ForEach-Object { Write-Host " - $_" }

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
