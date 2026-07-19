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

$prefixArguments = @("-c", "import sys; print(sys.prefix)")
$pythonPrefix = (& $Python @prefixArguments).Trim()
$tclSource = Join-Path $pythonPrefix "tcl\tcl8.6"
$tkSource = Join-Path $pythonPrefix "tcl\tk8.6"
if (-not (Test-Path -LiteralPath (Join-Path $tclSource "init.tcl"))) {
    throw "找不到 Tcl 8.6：$tclSource"
}
if (-not (Test-Path -LiteralPath (Join-Path $tkSource "tk.tcl"))) {
    throw "找不到 Tk 8.6：$tkSource"
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
