param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")),
    [string]$OutputDirectory = (Join-Path (Resolve-Path (Join-Path $PSScriptRoot "..")) "release"),
    [switch]$RunBuild
)

$solutionPath = Join-Path $ProjectRoot "spfx/UniversalHtmlViewer"
$packagePath = Join-Path $solutionPath "sharepoint/solution/universal-html-viewer.sppkg"

if ($RunBuild.IsPresent) {
    & (Join-Path $ProjectRoot "scripts/Build-UHV.ps1")
}

if (-not (Test-Path $packagePath)) {
    throw "Package not found at $packagePath. Run Build-UHV.ps1 or use -RunBuild."
}

if (-not (Test-Path $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
}

$timestamp = Get-Date -Format "yyyyMMdd-HHmm"
$zipName = "UniversalHtmlViewer-$timestamp.zip"
$zipPath = Join-Path $OutputDirectory $zipName

$files = @(
    $packagePath,
    (Join-Path $ProjectRoot "README.md"),
    (Join-Path $ProjectRoot "scripts/Deploy-UHV.ps1")
)

Compress-Archive -Path $files -DestinationPath $zipPath -Force
Write-Host "Release package created:" $zipPath
