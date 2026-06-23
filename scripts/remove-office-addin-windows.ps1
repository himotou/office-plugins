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
[xml]$manifestXml = Get-Content -LiteralPath $resolvedManifestPath

$addinIdNode = $manifestXml.SelectSingleNode("/*[local-name()='OfficeApp']/*[local-name()='Id']")
if (-not $addinIdNode -or [string]::IsNullOrWhiteSpace($addinIdNode.InnerText)) {
    throw "Unable to read add-in Id from manifest: $resolvedManifestPath"
}

$addinId = $addinIdNode.InnerText.Trim()
$targetDir = Join-Path $env:LOCALAPPDATA "link-bind-office-addin"
$registryPath = "HKCU:\Software\Microsoft\Office\16.0\Wef\Developer"
$removed = $false

if (Test-Path -LiteralPath $targetDir) {
    Get-ChildItem -Path $targetDir -Filter "$addinId.*" -File -ErrorAction SilentlyContinue | ForEach-Object {
        Remove-Item -LiteralPath $_.FullName -Force
        Write-Host "Removed manifest: $($_.FullName)"
        $removed = $true
    }
}

if (Test-Path -LiteralPath $registryPath) {
    $registryValue = Get-ItemProperty -Path $registryPath -Name $addinId -ErrorAction SilentlyContinue
    if ($null -ne $registryValue) {
        Remove-ItemProperty -Path $registryPath -Name $addinId -Force
        Write-Host "Removed registry entry: $registryPath -> $addinId"
        $removed = $true
    }
}

if (-not $removed) {
    Write-Host "No local sideload config found for add-in $addinId"
}

Write-Host "Restart PowerPoint to refresh the local add-in list."
