param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")),
    [string]$RequiredNodeVersion = "22.14.0",
    [switch]$ForceBootstrap
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$solutionPath = Join-Path $ProjectRoot "spfx/UniversalHtmlViewer"

if (-not (Test-Path $solutionPath)) {
    throw "SPFx solution not found at $solutionPath"
}

function Test-SupportedNodeVersion {
    param(
        [string]$NodeVersionText
    )

    if (-not $NodeVersionText) {
        return $false
    }

    $normalized = $NodeVersionText.Trim().TrimStart("v")
    if ($normalized -notmatch "^\d+\.\d+\.\d+$") {
        return $false
    }

    $version = [Version]$normalized
    return $version.Major -eq 22 -and $version -ge ([Version]"22.14.0")
}

function Get-LocalNodeRuntime {
    param(
        [string]$RepoRoot,
        [string]$NodeVersion,
        [bool]$ForceDownload = $false
    )

    if ($PSVersionTable.PSVersion.Major -lt 5) {
        throw "PowerShell 5+ is required to bootstrap the local Node.js runtime."
    }

    $toolsPath = Join-Path $RepoRoot ".tools"
    $runtimeFolderName = "node-v$NodeVersion-win-x64"
    $runtimeFolderPath = Join-Path $toolsPath $runtimeFolderName
    $nodeExePath = Join-Path $runtimeFolderPath "node.exe"
    $npmCmdPath = Join-Path $runtimeFolderPath "npm.cmd"

    if (-not $ForceDownload -and (Test-Path $nodeExePath) -and (Test-Path $npmCmdPath)) {
        return @{
            NodeExe = $nodeExePath
            NpmCmd = $npmCmdPath
            Source = "cached"
        }
    }

    if (-not (Test-Path $toolsPath)) {
        New-Item -ItemType Directory -Path $toolsPath | Out-Null
    }

    $zipPath = Join-Path $toolsPath "$runtimeFolderName.zip"
    $downloadUrl = "https://nodejs.org/dist/v$NodeVersion/$runtimeFolderName.zip"
    $checksumsPath = Join-Path $toolsPath "SHASUMS256-v$NodeVersion.txt"
    $checksumsUrl = "https://nodejs.org/dist/v$NodeVersion/SHASUMS256.txt"

    Write-Host "Downloading Node.js $NodeVersion from $downloadUrl"
    Invoke-WebRequest -Uri $downloadUrl -OutFile $zipPath
    Write-Host "Downloading Node.js checksums from $checksumsUrl"
    Invoke-WebRequest -Uri $checksumsUrl -OutFile $checksumsPath

    $expectedChecksum = Get-ExpectedArchiveChecksum -ChecksumsPath $checksumsPath -ArchiveFileName "$runtimeFolderName.zip"
    Assert-ArchiveChecksum -ArchivePath $zipPath -ExpectedChecksum $expectedChecksum

    if (Test-Path $runtimeFolderPath) {
        Remove-Item -Path $runtimeFolderPath -Recurse -Force
    }

    Expand-Archive -Path $zipPath -DestinationPath $toolsPath -Force
    Remove-Item -Path $zipPath -Force

    if (-not (Test-Path $nodeExePath) -or -not (Test-Path $npmCmdPath)) {
        throw "Failed to bootstrap local Node.js runtime at $runtimeFolderPath"
    }

    return @{
        NodeExe = $nodeExePath
        NpmCmd = $npmCmdPath
        Source = "downloaded"
    }
}

function Get-ExpectedArchiveChecksum {
    param(
        [string]$ChecksumsPath,
        [string]$ArchiveFileName
    )

    if (-not (Test-Path $ChecksumsPath)) {
        throw "Checksums file not found at $ChecksumsPath"
    }

    $matchingLine = Get-Content -Path $ChecksumsPath |
        Where-Object { $_ -match "\s+$([Regex]::Escape($ArchiveFileName))$" } |
        Select-Object -First 1

    if (-not $matchingLine) {
        throw "Unable to find checksum for $ArchiveFileName in $ChecksumsPath"
    }

    $parts = ($matchingLine -split "\s+", 2)
    $checksum = ""
    if ($parts.Length -gt 0 -and $parts[0]) {
        $checksum = $parts[0].Trim().ToLowerInvariant()
    }
    if ($checksum -notmatch "^[a-f0-9]{64}$") {
        throw "Invalid checksum format for $ArchiveFileName in $ChecksumsPath"
    }

    return $checksum
}

function Assert-ArchiveChecksum {
    param(
        [string]$ArchivePath,
        [string]$ExpectedChecksum
    )

    if (-not (Test-Path $ArchivePath)) {
        throw "Archive not found at $ArchivePath"
    }

    $actualChecksum = (Get-FileHash -Path $ArchivePath -Algorithm SHA256).Hash.Trim().ToLowerInvariant()
    if ($actualChecksum -ne $ExpectedChecksum) {
        throw "Checksum validation failed for $ArchivePath. Expected $ExpectedChecksum, got $actualChecksum."
    }

    Write-Host "Checksum verified for $ArchivePath"
}

$globalNodeVersion = ""
try {
    $globalNodeVersion = (& node -v 2>$null).Trim()
} catch {
    $globalNodeVersion = ""
}

$runtimeNodeExe = "node"
$runtimeNpmCmd = "npm"

if (Test-SupportedNodeVersion -NodeVersionText $globalNodeVersion) {
    Write-Host "Using Node.js from PATH: $globalNodeVersion"
} else {
    if ($globalNodeVersion) {
        Write-Host "Detected unsupported Node.js on PATH: $globalNodeVersion"
    } else {
        Write-Host "Node.js not found on PATH."
    }

    $localRuntime = Get-LocalNodeRuntime -RepoRoot $ProjectRoot -NodeVersion $RequiredNodeVersion -ForceDownload $ForceBootstrap.IsPresent
    $runtimeNodeExe = $localRuntime.NodeExe
    $runtimeNpmCmd = $localRuntime.NpmCmd

    $runtimePath = Split-Path -Parent $runtimeNodeExe
    $env:PATH = "$runtimePath;$env:PATH"
    Write-Host "Using local Node.js runtime ($($localRuntime.Source)): $runtimePath"
}

$activeNodeVersion = (& $runtimeNodeExe -v).Trim()
Write-Host "Active Node.js version: $activeNodeVersion"

Push-Location $solutionPath
try {
    if (Test-Path (Join-Path $solutionPath "package-lock.json")) {
        & $runtimeNpmCmd ci
    } else {
        & $runtimeNpmCmd install
    }
    if ($LASTEXITCODE -ne 0) {
        throw "Dependency installation failed."
    }

    & $runtimeNpmCmd run bundle:ship
    if ($LASTEXITCODE -ne 0) {
        throw "bundle:ship failed."
    }

    & $runtimeNpmCmd run package-solution:ship
    if ($LASTEXITCODE -ne 0) {
        throw "package-solution:ship failed."
    }

    $packagePath = Join-Path $solutionPath "sharepoint/solution/universal-html-viewer.sppkg"
    Write-Host "Package created:" $packagePath
} finally {
    Pop-Location
}
