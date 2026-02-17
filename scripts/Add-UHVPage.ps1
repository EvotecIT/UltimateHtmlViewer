[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,

    [string]$FullUrl,

    [string]$PageName = "UHV-Dashboard",
    [string]$PageTitle = "UHV Dashboard",
    [ValidateSet("Article", "Home", "SingleWebPartAppPage")]
    [string]$PageLayoutType = "Article",

    [ValidateSet("SharePointLibraryRelaxed", "SharePointLibraryFullPage", "SharePointLibraryStrict", "Custom")]
    [string]$ConfigurationPreset = "SharePointLibraryRelaxed",

    [ValidateSet("SharePointFileContent", "DirectUrl")]
    [string]$ContentDeliveryMode = "SharePointFileContent",

    [switch]$Publish,
    [switch]$SetAsHomePage,
    [switch]$ForceOverwrite,
    [switch]$EnsureSitePagesForceCheckout,
    [switch]$SkipAddWebPart,
    [switch]$SkipConfigureWebPartProperties,
    [switch]$SkipEnsureUhvAppOnSite,
    [ValidateSet("Tenant", "Site")]
    [string]$AppCatalogScope = "Tenant",
    [string]$UhvAppId = "15fe766e-03e4-4543-8129-c5e260b0b9e9",
    [string]$UhvAppNamePattern = "universal-html-viewer-client-side-solution",
    [int]$ComponentRetryCount = 8,
    [int]$ComponentRetryDelaySeconds = 3,

    [string]$ClientId,
    [string]$Tenant,
    [switch]$DeviceLogin
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (
    -not $SkipAddWebPart.IsPresent -and
    -not $SkipConfigureWebPartProperties.IsPresent -and
    [string]::IsNullOrWhiteSpace($FullUrl)
) {
    throw "FullUrl is required unless -SkipAddWebPart is specified."
}

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

function ConvertTo-PnPPropertiesJsonString {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Properties
    )

    $pairs = @()
    foreach ($key in $Properties.Keys) {
        $keyText = [string]$key
        $valueText = [string]$Properties[$key]
        $escapedValue = $valueText.Replace('"', '\"')
        $pairs += "`"$keyText`"=`"$escapedValue`""
    }

    return ($pairs -join ",")
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

function Ensure-UhvAppInstalledOnSite {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("Tenant", "Site")]
        [string]$PreferredScope
    )

    $resolved = Find-UhvCatalogApp -PreferredScope $PreferredScope
    if (-not $resolved) {
        throw "Universal HTML Viewer app package was not found in tenant/site app catalog. Deploy .sppkg first."
    }

    $app = $resolved.App
    $scopeToUse = $resolved.Scope
    $appIdText = Normalize-Id (Get-StringPropertyValue -InputObject $app -PropertyNames @("Id", "AppId"))
    $appTitle = (Get-StringPropertyValue -InputObject $app -PropertyNames @("Title", "Name")).Trim()
    if (-not $appTitle) {
        $appTitle = "Universal HTML Viewer"
    }

    Write-Host "Ensuring app is installed on site: $appTitle (Scope=$scopeToUse, Id=$appIdText)"

    try {
        $installIdentity = if ($appIdText) { $appIdText } else { $app }
        Install-PnPApp -Identity $installIdentity -Scope $scopeToUse -Wait | Out-Null
    } catch {
        $message = [string]$_.Exception.Message
        if ($message -match "already installed|already exists") {
            Write-Host "App is already installed on site."
        } else {
            throw
        }
    }

    try {
        $updateIdentity = if ($appIdText) { $appIdText } else { $app }
        Update-PnPApp -Identity $updateIdentity -Scope $scopeToUse | Out-Null
    } catch {
        $message = [string]$_.Exception.Message
        if ($message -match "No updates|current version|no upgrade|already latest") {
            Write-Host "App is already on latest version for this site."
        } else {
            Write-Warning "App update check returned: $message"
        }
    }
}

function Get-UhvComponentFromPage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PageIdentity
    )

    $uhvManifestId = "b5e51af1-1d0c-4b57-9b90-4f2af5120a4d"
    $normalizedManifestId = Normalize-Id $uhvManifestId
    $availableComponents = Get-PnPPageComponent -Page $PageIdentity -ListAvailable

    return $availableComponents | Where-Object {
        $idCandidates = @(
            Normalize-Id (Get-StringPropertyValue -InputObject $_ -PropertyNames @("Id")),
            Normalize-Id (Get-StringPropertyValue -InputObject $_ -PropertyNames @("ComponentId")),
            Normalize-Id (Get-StringPropertyValue -InputObject $_ -PropertyNames @("ClientSideComponentId")),
            Normalize-Id (Get-StringPropertyValue -InputObject $_ -PropertyNames @("ManifestId"))
        ) | Where-Object { $_ }
        return ($idCandidates -contains $normalizedManifestId)
    } | Select-Object -First 1
}

function Get-UhvComponentsOnPage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PageIdentity
    )

    $normalizedManifestId = Normalize-Id "b5e51af1-1d0c-4b57-9b90-4f2af5120a4d"
    $pageComponents = Get-PnPPageComponent -Page $PageIdentity

    return $pageComponents | Where-Object {
        $componentId = Normalize-Id (Get-StringPropertyValue -InputObject $_ -PropertyNames @("ComponentId", "ClientSideComponentId", "Id"))
        return $componentId -eq $normalizedManifestId
    }
}

if ($DeviceLogin.IsPresent) {
    if (-not $ClientId) {
        throw "ClientId is required with -DeviceLogin."
    }
    if (-not $Tenant) {
        throw "Tenant is required with -DeviceLogin."
    }
    Connect-PnPOnline -Url $SiteUrl -DeviceLogin -ClientId $ClientId -Tenant $Tenant
} else {
    if (-not $ClientId) {
        throw "ClientId is required with -Interactive on PnP.PowerShell 3.x."
    }
    if ($Tenant) {
        Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId -Tenant $Tenant
    } else {
        Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
    }
}

if ($EnsureSitePagesForceCheckout.IsPresent) {
    Set-PnPList -Identity "Site Pages" -ForceCheckout:$true
}

if (-not $SkipEnsureUhvAppOnSite.IsPresent) {
    Ensure-UhvAppInstalledOnSite -PreferredScope $AppCatalogScope
}

$pageFileName = $PageName.Trim()
if (-not $pageFileName) {
    throw "PageName cannot be empty."
}
if (-not $pageFileName.EndsWith(".aspx", [System.StringComparison]::OrdinalIgnoreCase)) {
    $pageFileName = "$pageFileName.aspx"
}
$pageNameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($pageFileName)

$existingPage = $null
try {
    $existingPage = Get-PnPPage -Identity $pageFileName -ErrorAction Stop
} catch {
    $existingPage = $null
}

if ($existingPage) {
    if (-not $ForceOverwrite.IsPresent) {
        throw "Page '$pageFileName' already exists. Use -ForceOverwrite to recreate it."
    }
    Remove-PnPPage -Identity $pageFileName -Force
}

Write-Host "Creating page $pageFileName on $SiteUrl"
Add-PnPPage -Name $pageNameWithoutExtension -Title $PageTitle -LayoutType $PageLayoutType | Out-Null

if (-not $SkipAddWebPart.IsPresent -and $PageLayoutType -ne "SingleWebPartAppPage") {
    Add-PnPPageSection -Page $pageFileName -SectionTemplate OneColumn -Order 1 | Out-Null
}

if ($SkipAddWebPart.IsPresent) {
    if ($Publish.IsPresent) {
        Set-PnPPage -Identity $pageFileName -Publish | Out-Null
    }

    if ($SetAsHomePage.IsPresent) {
        Set-PnPHomePage -RootFolderRelativeUrl "SitePages/$pageFileName"
    }

    $pageUrl = "$($SiteUrl.TrimEnd('/'))/SitePages/$pageFileName"
    Write-Host "Done (page only, no web part). Page URL: $pageUrl"
    return
}

$uhvComponent = $null
for ($attempt = 1; $attempt -le $ComponentRetryCount; $attempt++) {
    $uhvComponent = Get-UhvComponentFromPage -PageIdentity $pageFileName
    if ($uhvComponent) {
        break
    }

    if ($attempt -lt $ComponentRetryCount) {
        Start-Sleep -Seconds $ComponentRetryDelaySeconds
    }
}

$uhvManifestId = "b5e51af1-1d0c-4b57-9b90-4f2af5120a4d"
if (-not $uhvComponent) {
    Write-Warning "Universal HTML Viewer component was not found in page toolbox after waiting. Trying direct add using manifest id."
}

$webPartProperties = @{
    configurationPreset = $ConfigurationPreset
    contentDeliveryMode = $ContentDeliveryMode
    htmlSourceMode = "FullUrl"
    fullUrl = $FullUrl
}
$webPartPropertiesJson = ConvertTo-PnPPropertiesJsonString -Properties $webPartProperties

Write-Host "Adding Universal HTML Viewer web part with FullUrl=$FullUrl"
$webPartAdded = $false
$addedWebPartInstanceId = ""
if ($uhvComponent) {
    try {
        if ($PageLayoutType -eq "SingleWebPartAppPage") {
            $addedControl = Add-PnPPageWebPart `
                -Page $pageFileName `
                -Component $uhvComponent `
                -Order 1
        } else {
            $addedControl = Add-PnPPageWebPart `
                -Page $pageFileName `
                -Component $uhvComponent `
                -Section 1 `
                -Column 1 `
                -Order 1
        }
        $addedWebPartInstanceId = (Get-StringPropertyValue -InputObject $addedControl -PropertyNames @("InstanceId", "Id")).Trim()
        $webPartAdded = $true
    } catch {
        Write-Warning "Adding UHV via discovered component object failed."
    }
}

if (-not $webPartAdded) {
    if ($PageLayoutType -eq "SingleWebPartAppPage") {
        $addedControl = Add-PnPPageWebPart `
            -Page $pageFileName `
            -Component $uhvManifestId `
            -Order 1
    } else {
        $addedControl = Add-PnPPageWebPart `
            -Page $pageFileName `
            -Component $uhvManifestId `
            -Section 1 `
            -Column 1 `
            -Order 1
    }
    $addedWebPartInstanceId = (Get-StringPropertyValue -InputObject $addedControl -PropertyNames @("InstanceId", "Id")).Trim()
    $webPartAdded = $true
}

if (-not $SkipConfigureWebPartProperties.IsPresent) {
    $setSucceeded = $false
    if ($addedWebPartInstanceId) {
        try {
            Set-PnPPageWebPart -Page $pageFileName -Identity $addedWebPartInstanceId -PropertiesJson $webPartPropertiesJson | Out-Null
            $setSucceeded = $true
        } catch {
            Write-Warning "Setting web part properties by returned instance id failed. Trying discovery fallback."
        }
    }

    if (-not $setSucceeded) {
        $uhvControls = Get-UhvComponentsOnPage -PageIdentity $pageFileName
        if (-not $uhvControls -or $uhvControls.Count -eq 0) {
            throw "UHV web part was added but instance id could not be resolved to set properties."
        }

        $targetControl = $uhvControls | Select-Object -Last 1
        $targetInstanceId = (Get-StringPropertyValue -InputObject $targetControl -PropertyNames @("InstanceId", "Id")).Trim()
        if (-not $targetInstanceId) {
            throw "UHV web part was added but target instance id is missing; cannot set properties safely."
        }

        Set-PnPPageWebPart -Page $pageFileName -Identity $targetInstanceId -PropertiesJson $webPartPropertiesJson | Out-Null
    }
} else {
    Write-Host "Skipping UHV property configuration (-SkipConfigureWebPartProperties)."
}

if ($Publish.IsPresent) {
    Set-PnPPage -Identity $pageFileName -Publish | Out-Null
}

if ($SetAsHomePage.IsPresent) {
    Set-PnPHomePage -RootFolderRelativeUrl "SitePages/$pageFileName"
}

$pageUrl = "$($SiteUrl.TrimEnd('/'))/SitePages/$pageFileName"
Write-Host "Done. Page URL: $pageUrl"
