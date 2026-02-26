[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [string]$FullUrl,
    [string]$SiteRelativeDashboardPath = "SiteAssets/Index.html",

    [string]$PageName = "Dashboard",
    [string]$PageTitle = "Dashboard",

    [ValidateSet("Article", "Home", "SingleWebPartAppPage")]
    [string]$PageLayoutType = "Article",

    [ValidateSet("SharePointLibraryRelaxed", "SharePointLibraryFullPage", "SharePointLibraryStrict", "Custom")]
    [string]$ConfigurationPreset = "SharePointLibraryRelaxed",

    [ValidateSet("SharePointFileContent", "DirectUrl")]
    [string]$ContentDeliveryMode = "SharePointFileContent",

    [bool]$Publish = $true,
    [bool]$ForceOverwrite = $true,
    [bool]$EnsureSitePagesForceCheckout = $false,
    [bool]$SetAsHomePage = $false,

    [switch]$InstallOnly,

    [ValidateSet("Tenant", "Site")]
    [string]$AppCatalogScope = "Tenant",

    [string]$ClientId,
    [string]$Tenant,
    [switch]$DeviceLogin,
    [bool]$PersistLogin = $true,
    [switch]$ForceAuthentication
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-FullUrl {
    param(
        [string]$TargetSiteUrl,
        [string]$ExplicitFullUrl,
        [string]$RelativePath
    )

    if (-not [string]::IsNullOrWhiteSpace($ExplicitFullUrl)) {
        return $ExplicitFullUrl.Trim()
    }

    $baseSiteUrl = $TargetSiteUrl.TrimEnd('/')
    $siteUri = [Uri]"$baseSiteUrl/"
    $normalized = [string]$RelativePath
    $normalized = $normalized.Trim()
    if (-not $normalized) {
        $normalized = "SiteAssets/Index.html"
    }

    return ([Uri]::new($siteUri, $normalized)).AbsoluteUri
}

$scriptRoot = $PSScriptRoot
if (-not $scriptRoot) {
    $scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
}

if ($InstallOnly.IsPresent) {
    $updateScript = Join-Path $scriptRoot "Update-UHVSiteApp.ps1"
    if (-not (Test-Path $updateScript)) {
        throw "Update script not found at $updateScript"
    }

    Write-Host "Ensuring UHV app is installed/updated on site: $SiteUrl"
    & $updateScript `
        -SiteUrls @($SiteUrl) `
        -AppCatalogScope $AppCatalogScope `
        -InstallIfMissing `
        -ClientId $ClientId `
        -Tenant $Tenant `
        -DeviceLogin:$DeviceLogin.IsPresent `
        -PersistLogin $PersistLogin `
        -ForceAuthentication:$ForceAuthentication.IsPresent
    return
}

$addPageScript = Join-Path $scriptRoot "Add-UHVPage.ps1"
if (-not (Test-Path $addPageScript)) {
    throw "Add page script not found at $addPageScript"
}

$effectiveFullUrl = Resolve-FullUrl `
    -TargetSiteUrl $SiteUrl `
    -ExplicitFullUrl $FullUrl `
    -RelativePath $SiteRelativeDashboardPath

Write-Host "Running one-command UHV site setup..."
Write-Host "Site: $SiteUrl"
Write-Host "Dashboard URL: $effectiveFullUrl"

& $addPageScript `
    -SiteUrl $SiteUrl `
    -FullUrl $effectiveFullUrl `
    -PageName $PageName `
    -PageTitle $PageTitle `
    -PageLayoutType $PageLayoutType `
    -ConfigurationPreset $ConfigurationPreset `
    -ContentDeliveryMode $ContentDeliveryMode `
    -Publish:$Publish `
    -ForceOverwrite:$ForceOverwrite `
    -SetAsHomePage:$SetAsHomePage `
    -EnsureSitePagesForceCheckout:$EnsureSitePagesForceCheckout `
    -AppCatalogScope $AppCatalogScope `
    -ClientId $ClientId `
    -Tenant $Tenant `
    -DeviceLogin:$DeviceLogin.IsPresent `
    -PersistLogin $PersistLogin `
    -ForceAuthentication:$ForceAuthentication.IsPresent
