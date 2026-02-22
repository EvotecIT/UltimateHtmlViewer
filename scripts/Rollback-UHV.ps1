[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$AppCatalogUrl,

    [Parameter(Mandatory = $true)]
    [string]$RollbackSppkgPath,

    [ValidateSet("Tenant", "Site")]
    [string]$Scope = "Tenant",

    [switch]$TenantWide,

    [string[]]$SiteUrls,

    [ValidateSet("Tenant", "Site")]
    [string]$AppCatalogScope = "Tenant",

    [switch]$InstallIfMissing,

    [switch]$SkipSiteUpdate,

    [string]$ClientId,
    [string]$Tenant,
    [switch]$DeviceLogin,
    [bool]$PersistLogin = $true,
    [switch]$ForceAuthentication,
    [string]$TenantAdminUrl
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not (Test-Path $RollbackSppkgPath)) {
    throw "Rollback package not found: $RollbackSppkgPath"
}

$scriptRoot = $PSScriptRoot
if (-not $scriptRoot) {
    $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}

$deployScript = Join-Path $scriptRoot "Deploy-UHV.ps1"
if (-not (Test-Path $deployScript)) {
    throw "Deploy script not found at $deployScript"
}

$updateScript = Join-Path $scriptRoot "Update-UHVSiteApp.ps1"
if (-not (Test-Path $updateScript)) {
    throw "Update script not found at $updateScript"
}

Write-Host "Rolling back app catalog package to: $RollbackSppkgPath"
& $deployScript `
    -AppCatalogUrl $AppCatalogUrl `
    -SppkgPath $RollbackSppkgPath `
    -Scope $Scope `
    -TenantWide:$TenantWide.IsPresent `
    -ClientId $ClientId `
    -Tenant $Tenant `
    -DeviceLogin:$DeviceLogin.IsPresent `
    -PersistLogin $PersistLogin `
    -ForceAuthentication:$ForceAuthentication.IsPresent `
    -TenantAdminUrl $TenantAdminUrl

if ($SkipSiteUpdate.IsPresent) {
    Write-Host "Skipping site app update (-SkipSiteUpdate)."
    return
}

if (-not $SiteUrls -or $SiteUrls.Count -eq 0) {
    Write-Warning "Rollback package published, but no -SiteUrls provided. Existing site app instances may still run a newer version until updated."
    return
}

Write-Host "Updating app on target sites after rollback..."
& $updateScript `
    -SiteUrls $SiteUrls `
    -AppCatalogScope $AppCatalogScope `
    -InstallIfMissing:$InstallIfMissing.IsPresent `
    -ClientId $ClientId `
    -Tenant $Tenant `
    -DeviceLogin:$DeviceLogin.IsPresent `
    -PersistLogin $PersistLogin `
    -ForceAuthentication:$ForceAuthentication.IsPresent
