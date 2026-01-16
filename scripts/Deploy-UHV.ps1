param(
    [Parameter(Mandatory = $true)]
    [string]$AppCatalogUrl,

    [string]$SppkgPath = (Join-Path (Resolve-Path (Join-Path $PSScriptRoot "..")) "spfx/UniversalHtmlViewer/sharepoint/solution/universal-html-viewer.sppkg"),

    [switch]$TenantWide
)

if (-not (Test-Path $SppkgPath)) {
    throw "Package not found at $SppkgPath. Run Build-UHV.ps1 first."
}

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    throw "PnP.PowerShell module not found. Install-Module PnP.PowerShell -Scope CurrentUser"
}

Write-Host "Connecting to App Catalog:" $AppCatalogUrl
Connect-PnPOnline -Url $AppCatalogUrl -Interactive

Write-Host "Uploading package:" $SppkgPath
$app = Add-PnPApp -Path $SppkgPath -Overwrite

Write-Host "Publishing package (TenantWide=$($TenantWide.IsPresent))"
if ($TenantWide.IsPresent) {
    Publish-PnPApp -Identity $app.Id -SkipFeatureDeployment
} else {
    Publish-PnPApp -Identity $app.Id
}

Write-Host "Deployment completed."
