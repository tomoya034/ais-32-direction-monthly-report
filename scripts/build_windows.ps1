param(
    [string]$Python = "python"
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$buildRoot = Join-Path $projectRoot ".build"
$buildTools = Join-Path $buildRoot "tools"
$tclRuntime = Join-Path $buildRoot "tcl_runtime"
$pyinstallerWork = Join-Path $buildRoot "pyinstaller"
$distPath = Join-Path $projectRoot "dist"

New-Item -ItemType Directory -Path $buildRoot -Force | Out-Null
New-Item -ItemType Directory -Path $buildTools -Force | Out-Null
New-Item -ItemType Directory -Path $tclRuntime -Force | Out-Null

& $Python -m pip install --upgrade --target $buildTools `
    -r (Join-Path $projectRoot "requirements.txt") `
    -r (Join-Path $projectRoot "requirements-build.txt")
if ($LASTEXITCODE -ne 0) { throw "建置相依套件安裝失敗。" }

$pythonPrefixes = @(
    & $Python -c "import sys; print(sys.prefix); print(sys.base_prefix)"
) | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique

$tclSource = $null
$tkSource = $null
foreach ($pythonPrefix in $pythonPrefixes) {
    $candidateTcl = Join-Path $pythonPrefix "tcl\tcl8.6"
    $candidateTk = Join-Path $pythonPrefix "tcl\tk8.6"
    if (
        (Test-Path -LiteralPath (Join-Path $candidateTcl "init.tcl")) -and
        (Test-Path -LiteralPath (Join-Path $candidateTk "tk.tcl"))
    ) {
        $tclSource = $candidateTcl
        $tkSource = $candidateTk
        break
    }
}
if (-not $tclSource -or -not $tkSource) {
    throw "找不到 Tcl/Tk 8.6；已檢查：$($pythonPrefixes -join ', ')"
}

Copy-Item -LiteralPath $tclSource -Destination $tclRuntime -Recurse -Force
Copy-Item -LiteralPath $tkSource -Destination $tclRuntime -Recurse -Force

$env:PYTHONPATH = $buildTools
$env:TCL_LIBRARY = Join-Path $tclRuntime "tcl8.6"
$env:TK_LIBRARY = Join-Path $tclRuntime "tk8.6"

& $Python -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --name "AIS_32方位月報工具" `
    --version-file (Join-Path $projectRoot "version_info.txt") `
    --distpath $distPath `
    --workpath $pyinstallerWork `
    --specpath $buildRoot `
    (Join-Path $projectRoot "ais_monthly_app.py")
if ($LASTEXITCODE -ne 0) { throw "EXE 建置失敗。" }

$exe = Join-Path $distPath "AIS_32方位月報工具.exe"
$hash = Get-FileHash -LiteralPath $exe -Algorithm SHA256
Write-Host "建置完成：$exe"
Write-Host "SHA-256：$($hash.Hash)"
