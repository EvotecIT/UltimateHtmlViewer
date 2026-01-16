param(
    [string]$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot ".."))
)

$solutionPath = Join-Path $ProjectRoot "spfx/UniversalHtmlViewer"

if (-not (Test-Path $solutionPath)) {
    throw "SPFx solution not found at $solutionPath"
}

Push-Location $solutionPath
try {
    npm install
    npm run bundle:ship
    npm run package-solution:ship

    $packagePath = Join-Path $solutionPath "sharepoint/solution/universal-html-viewer.sppkg"
    Write-Host "Package created:" $packagePath
} finally {
    Pop-Location
}
