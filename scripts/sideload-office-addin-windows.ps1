param(
    [string]$ManifestPath = ""
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoDir = Split-Path -Parent $scriptDir

if ([string]::IsNullOrWhiteSpace($ManifestPath)) {
    $distManifest = Join-Path $repoDir "dist\manifest.xml"
    $devManifest = Join-Path $repoDir "manifest.xml"

    if (Test-Path -LiteralPath $distManifest) {
        $ManifestPath = $distManifest
    }
    elseif (Test-Path -LiteralPath $devManifest) {
        $ManifestPath = $devManifest
    }
    else {
        throw "Manifest file not found. Pass -ManifestPath explicitly."
    }
}

$resolvedManifestPath = (Resolve-Path -LiteralPath $ManifestPath).Path
$devSettingsCmd = Join-Path $repoDir "node_modules\.bin\office-addin-dev-settings.cmd"

if (-not (Test-Path -LiteralPath $devSettingsCmd)) {
    throw "office-addin-dev-settings is not installed. Run npm install first."
}

Write-Host "Clearing stale PowerPoint sideload config..."
& (Join-Path $scriptDir "remove-office-addin-windows.ps1") -ManifestPath $resolvedManifestPath *> $null

Write-Host "Registering add-in and launching PowerPoint with sideload document: $resolvedManifestPath"
& $devSettingsCmd sideload $resolvedManifestPath desktop --app powerpoint
