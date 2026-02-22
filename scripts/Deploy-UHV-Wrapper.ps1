[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$AppCatalogUrl,

    [string]$SppkgPath,

    [ValidateSet("Tenant", "Site")]
    [string]$Scope = "Tenant",

    [switch]$TenantWide,

    [string]$ClientId,

    [string]$Tenant,

    [switch]$DeviceLogin,

    [bool]$PersistLogin = $true,

    [switch]$ForceAuthentication,

    [string]$TenantAdminUrl,

    [switch]$SkipBuild,

    [switch]$NoDeploy,

    [switch]$ForceBootstrap
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$scriptRoot = $PSScriptRoot
if (-not $scriptRoot) {
    $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}
if (-not $scriptRoot) {
    throw "Unable to determine script root path."
}

$repoRoot = Resolve-Path (Join-Path $scriptRoot "..")

if (-not $SppkgPath) {
    $SppkgPath = Join-Path $repoRoot "spfx/UniversalHtmlViewer/sharepoint/solution/universal-html-viewer.sppkg"
}

if (-not $SkipBuild.IsPresent) {
    $buildScript = Join-Path $repoRoot "scripts/Build-UHV.ps1"
    if (-not (Test-Path $buildScript)) {
        throw "Build script not found at $buildScript"
    }

    & $buildScript -ForceBootstrap:$ForceBootstrap.IsPresent
}

if ($NoDeploy.IsPresent) {
    Write-Host "Skipping deployment (-NoDeploy). Package path: $SppkgPath"
    return
}

$deployScript = Join-Path $repoRoot "scripts/Deploy-UHV.ps1"
if (-not (Test-Path $deployScript)) {
    throw "Deploy script not found at $deployScript"
}

& $deployScript `
    -AppCatalogUrl $AppCatalogUrl `
    -SppkgPath $SppkgPath `
    -Scope $Scope `
    -TenantWide:$TenantWide.IsPresent `
    -ClientId $ClientId `
    -Tenant $Tenant `
    -DeviceLogin:$DeviceLogin.IsPresent `
    -PersistLogin $PersistLogin `
    -ForceAuthentication:$ForceAuthentication.IsPresent `
    -TenantAdminUrl $TenantAdminUrl
