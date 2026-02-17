[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string[]]$SiteUrls,

    [ValidateSet("Tenant", "Site")]
    [string]$AppCatalogScope = "Tenant",

    [string]$UhvAppId = "15fe766e-03e4-4543-8129-c5e260b0b9e9",
    [string]$UhvAppNamePattern = "universal-html-viewer-client-side-solution",
    [switch]$InstallIfMissing,

    [string]$ClientId,
    [string]$Tenant,
    [switch]$DeviceLogin
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    throw "PnP.PowerShell module not found. Install-Module PnP.PowerShell -Scope CurrentUser"
}

function Get-StringPropertyValue {
    param(
        [Parameter(Mandatory = $true)]
        [object]$InputObject,
        [Parameter(Mandatory = $true)]
        [string[]]$PropertyNames
    )

    foreach ($propertyName in $PropertyNames) {
        if ($InputObject.PSObject.Properties[$propertyName]) {
            $value = $InputObject.PSObject.Properties[$propertyName].Value
            if ($null -ne $value) {
                return [string]$value
            }
        }
    }

    return ""
}

function Normalize-Id {
    param(
        [string]$Value
    )

    if (-not $Value) {
        return ""
    }

    return $Value.Trim().TrimStart("{").TrimEnd("}").ToLowerInvariant()
}

function Find-UhvCatalogApp {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("Tenant", "Site")]
        [string]$PreferredScope
    )

    $scopesToCheck = @($PreferredScope)
    if ($PreferredScope -eq "Tenant") {
        $scopesToCheck += "Site"
    } else {
        $scopesToCheck += "Tenant"
    }

    foreach ($scopeToCheck in $scopesToCheck) {
        $apps = Get-PnPApp -Scope $scopeToCheck
        $matched = $apps | Where-Object {
            $idValue = Normalize-Id (Get-StringPropertyValue -InputObject $_ -PropertyNames @("Id", "AppId"))
            $titleValue = (Get-StringPropertyValue -InputObject $_ -PropertyNames @("Title", "Name")).Trim().ToLowerInvariant()

            $idMatch = $false
            if ($UhvAppId) {
                $idMatch = $idValue -eq (Normalize-Id $UhvAppId)
            }

            $nameMatch = $false
            if ($UhvAppNamePattern) {
                $nameMatch = $titleValue -like "*$($UhvAppNamePattern.Trim().ToLowerInvariant())*"
            }

            return $idMatch -or $nameMatch
        } | Select-Object -First 1

        if ($matched) {
            return @{
                App = $matched
                Scope = $scopeToCheck
            }
        }
    }

    return $null
}

function Test-IsInstalled {
    param(
        [Parameter(Mandatory = $true)]
        [object]$App
    )

    if ($App.PSObject.Properties["Installed"]) {
        return [bool]$App.Installed
    }

    $installedVersion = Get-StringPropertyValue -InputObject $App -PropertyNames @("InstalledVersion")
    return -not [string]::IsNullOrWhiteSpace($installedVersion)
}

$results = @()

foreach ($siteUrl in $SiteUrls) {
    $status = "Unknown"
    $details = ""
    $resolvedScope = ""
    $appTitle = ""

    try {
        Write-Host "Connecting to site: $siteUrl"
        if ($DeviceLogin.IsPresent) {
            if (-not $ClientId) {
                throw "ClientId is required with -DeviceLogin."
            }
            if (-not $Tenant) {
                throw "Tenant is required with -DeviceLogin."
            }
            Connect-PnPOnline -Url $siteUrl -DeviceLogin -ClientId $ClientId -Tenant $Tenant
        } else {
            if (-not $ClientId) {
                throw "ClientId is required with -Interactive on PnP.PowerShell 3.x."
            }
            if ($Tenant) {
                Connect-PnPOnline -Url $siteUrl -Interactive -ClientId $ClientId -Tenant $Tenant
            } else {
                Connect-PnPOnline -Url $siteUrl -Interactive -ClientId $ClientId
            }
        }

        $resolved = Find-UhvCatalogApp -PreferredScope $AppCatalogScope
        if (-not $resolved) {
            $status = "AppNotFound"
            $details = "App package not found in tenant/site app catalog."
        } else {
            $app = $resolved.App
            $resolvedScope = $resolved.Scope
            $appIdText = Normalize-Id (Get-StringPropertyValue -InputObject $app -PropertyNames @("Id", "AppId"))
            $appTitle = (Get-StringPropertyValue -InputObject $app -PropertyNames @("Title", "Name")).Trim()
            if (-not $appTitle) {
                $appTitle = "Universal HTML Viewer"
            }

            $isInstalled = Test-IsInstalled -App $app
            if (-not $isInstalled) {
                if ($InstallIfMissing.IsPresent) {
                    $installIdentity = if ($appIdText) { $appIdText } else { $app }
                    Install-PnPApp -Identity $installIdentity -Scope $resolvedScope -Wait | Out-Null
                    $status = "Installed"
                    $details = "App installed on site."
                } else {
                    $status = "NotInstalled"
                    $details = "App is available but not installed. Use -InstallIfMissing."
                }
            }

            if ($status -ne "NotInstalled") {
                try {
                    $updateIdentity = if ($appIdText) { $appIdText } else { $app }
                    Update-PnPApp -Identity $updateIdentity -Scope $resolvedScope | Out-Null
                    if ($status -eq "Installed") {
                        $details = "App installed and update check completed."
                    } else {
                        $status = "UpdatedOrCurrent"
                        $details = "Update command completed (app updated or already current)."
                    }
                } catch {
                    $message = [string]$_.Exception.Message
                    if ($message -match "No updates|current version|no upgrade|already latest") {
                        if ($status -eq "Installed") {
                            $details = "App installed and already on latest version."
                        } else {
                            $status = "Current"
                            $details = "App already on latest version."
                        }
                    } else {
                        $status = "UpdateWarning"
                        $details = $message
                    }
                }
            }
        }
    } catch {
        $status = "Error"
        $details = [string]$_.Exception.Message
    }

    $results += [PSCustomObject]@{
        SiteUrl = $siteUrl
        ScopeUsed = $resolvedScope
        AppTitle = $appTitle
        Status = $status
        Details = $details
    }
}

$results | Format-Table -AutoSize
