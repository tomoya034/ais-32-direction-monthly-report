param(
    [string]$Python = "python",
    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$version = (Get-Content -LiteralPath (Join-Path $projectRoot "VERSION") -Raw).Trim()
$distPath = Join-Path $projectRoot "dist"
$exe = Join-Path $distPath "AIS_32方位月報工具.exe"

if (-not $SkipBuild) {
    & (Join-Path $PSScriptRoot "build_windows.ps1") -Python $Python
}
if (-not (Test-Path -LiteralPath $exe)) {
    throw "找不到建置結果：$exe"
}

$packageName = "AIS_32方位月報工具_v$version"
$packageDirectory = Join-Path $distPath $packageName
$zipPath = Join-Path $distPath "$packageName.zip"
New-Item -ItemType Directory -Path $packageDirectory -Force | Out-Null
Copy-Item -LiteralPath $exe -Destination (Join-Path $packageDirectory "AIS_32方位月報工具.exe") -Force
Copy-Item -LiteralPath (Join-Path $projectRoot "使用說明.txt") -Destination $packageDirectory -Force
Compress-Archive -LiteralPath $packageDirectory -DestinationPath $zipPath -CompressionLevel Optimal -Force

$exeHash = Get-FileHash -LiteralPath $exe -Algorithm SHA256
$zipHash = Get-FileHash -LiteralPath $zipPath -Algorithm SHA256
$checksums = @(
    "$($exeHash.Hash)  AIS_32方位月報工具.exe",
    "$($zipHash.Hash)  $packageName.zip"
)
$checksumsPath = Join-Path $distPath "SHA256SUMS.txt"
Set-Content -LiteralPath $checksumsPath -Value $checksums -Encoding UTF8

Write-Host "Release 套件：$zipPath"
Write-Host "Checksum：$checksumsPath"
