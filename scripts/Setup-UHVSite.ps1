[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
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

function ConvertTo-HttpsUri {
    param(
        [string]$Value,
        [string]$ParameterName
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        throw "$ParameterName is required."
    }

    $parsed = $null
    if (-not [Uri]::TryCreate($Value, [UriKind]::Absolute, [ref]$parsed) -or $parsed.Scheme -ne "https") {
        throw "$ParameterName must be an absolute HTTPS URL."
    }

    return $parsed
}

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
    if (-not $PSCmdlet.ShouldProcess($SiteUrl, "Install or update UHV app on site")) {
        return
    }

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

$siteUri = ConvertTo-HttpsUri -Value $SiteUrl -ParameterName "SiteUrl"
$effectiveFullUri = ConvertTo-HttpsUri -Value $effectiveFullUrl -ParameterName "FullUrl"

if ($ContentDeliveryMode -eq "SharePointFileContent") {
    if ($effectiveFullUri.Host.ToLowerInvariant() -ne $siteUri.Host.ToLowerInvariant()) {
        throw "For SharePointFileContent mode, FullUrl must use the same host as SiteUrl. SiteUrl host='$($siteUri.Host)', FullUrl host='$($effectiveFullUri.Host)'."
    }
}
if (-not $PSCmdlet.ShouldProcess($SiteUrl, "Provision UHV page '$PageName'")) {
    return
}

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
