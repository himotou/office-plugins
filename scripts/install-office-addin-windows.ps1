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

$sourceLocationNode = $manifestXml.SelectSingleNode("/*[local-name()='OfficeApp']/*[local-name()='DefaultSettings']/*[local-name()='SourceLocation']")
$sourceLocation = ""
if ($sourceLocationNode) {
    $sourceLocation = $sourceLocationNode.GetAttribute("DefaultValue")
}

$addinId = $addinIdNode.InnerText.Trim()
$targetDir = Join-Path $env:LOCALAPPDATA "link-bind-office-addin"
$targetFileName = "$addinId.$([System.IO.Path]::GetFileName($resolvedManifestPath))"
$targetPath = Join-Path $targetDir $targetFileName

New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
Get-ChildItem -Path $targetDir -Filter "$addinId.*" -File -ErrorAction SilentlyContinue | Remove-Item -Force
Copy-Item -LiteralPath $resolvedManifestPath -Destination $targetPath -Force

$registryPath = "HKCU:\Software\Microsoft\Office\16.0\Wef\Developer"
New-Item -Path $registryPath -Force | Out-Null
New-ItemProperty -Path $registryPath -Name $addinId -Value $targetPath -PropertyType String -Force | Out-Null

Write-Host "Installed manifest: $targetPath"
Write-Host "Registered add-in: $registryPath -> $addinId"
Write-Host "Manifest source: $sourceLocation"

if ($sourceLocation -match '^https?://(localhost|127\.0\.0\.1)(:|/)') {
    Write-Warning "This manifest points to localhost. Use dist/manifest.xml for online deployment."
}
elseif (-not [string]::IsNullOrWhiteSpace($sourceLocation)) {
    Write-Warning "This manifest points to a remote URL. The add-in will only appear if that URL is reachable from PowerPoint."
}

Write-Host "Restart PowerPoint, open a presentation, then use Insert > My Add-ins next to the drop-down arrow to load the add-in."
