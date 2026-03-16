param(
    [string]$PythonExe = "python",
    [string]$EntryScript = "app-V10-2.py",
    [string]$ExeName = "TEMPO-CR",
    [ValidateSet("FastAPI", "Script")]
    [string]$Mode = "FastAPI",
    [string]$AppVarName = "app",
    [string]$Host = "127.0.0.1",
    [int]$Port = 8090,
    [switch]$OpenBrowser = $true,
    [switch]$OneFile = $true,
    [switch]$Clean = $true,
    [switch]$NoConsole = $false
)

$ErrorActionPreference = "Stop"

function Step([string]$msg) {
    Write-Host "`n==> $msg" -ForegroundColor Cyan
}

if (-not (Test-Path $EntryScript)) {
    throw "Script d'entrée introuvable: $EntryScript"
}

Step "Vérification de PyInstaller"
& $PythonExe -m PyInstaller --version 2>$null
if ($LASTEXITCODE -ne 0) {
    Step "Installation de PyInstaller"
    & $PythonExe -m pip install --upgrade pyinstaller
}

$flags = @()
if ($OneFile) { $flags += "--onefile" }
if ($Clean) { $flags += "--clean" }
if ($NoConsole) { $flags += "--noconsole" }

$buildTarget = $EntryScript
$tempLauncher = ""

if ($Mode -eq "FastAPI") {
    Step "Génération d'un launcher FastAPI"

    $entryAbs = (Resolve-Path $EntryScript).Path
    $entryEscaped = $entryAbs.Replace("\\", "\\\\")

    $launcherCode = @"
import importlib.util
import pathlib
import sys
import uvicorn

ENTRY_PATH = pathlib.Path(r"$entryEscaped")
APP_VAR = "$AppVarName"
HOST = "$Host"
PORT = $Port
OPEN_BROWSER = $([bool]$OpenBrowser)

spec = importlib.util.spec_from_file_location("tempo_app_module", str(ENTRY_PATH))
module = importlib.util.module_from_spec(spec)
sys.modules["tempo_app_module"] = module
spec.loader.exec_module(module)

if not hasattr(module, APP_VAR):
    raise RuntimeError(f"Variable FastAPI '{APP_VAR}' introuvable dans {ENTRY_PATH}")

app = getattr(module, APP_VAR)

if OPEN_BROWSER:
    try:
        import threading
        import webbrowser
        def _open_browser():
            webbrowser.open(f"http://{HOST}:{PORT}")
        threading.Timer(1.2, _open_browser).start()
    except Exception:
        pass

uvicorn.run(app, host=HOST, port=PORT, log_level="info")
"@

    $tempLauncher = Join-Path (Get-Location) "_temp_launcher_build.py"
    Set-Content -Path $tempLauncher -Value $launcherCode -Encoding UTF8
    $buildTarget = $tempLauncher

    # Dépendances utiles au runtime FastAPI/uvicorn
    $flags += @(
        "--hidden-import=uvicorn",
        "--hidden-import=fastapi",
        "--hidden-import=starlette",
        "--hidden-import=anyio",
        "--hidden-import=pydantic",
        "--hidden-import=pandas"
    )
}

try {
    Step "Build EXE"
    & $PythonExe -m PyInstaller @flags --name $ExeName $buildTarget
    if ($LASTEXITCODE -ne 0) {
        throw "Échec de compilation PyInstaller"
    }
}
finally {
    if ($tempLauncher -and (Test-Path $tempLauncher)) {
        Remove-Item $tempLauncher -Force -ErrorAction SilentlyContinue
    }
}

$exePath = Join-Path (Join-Path (Get-Location) "dist") "$ExeName.exe"
if (Test-Path $exePath) {
    Step "Build terminé"
    Write-Host "EXE généré: $exePath" -ForegroundColor Green
    if ($Mode -eq "FastAPI") {
        Write-Host "Au lancement, l'EXE démarre le serveur sur http://$Host`:$Port" -ForegroundColor Green
    }
} else {
    Write-Warning "Build terminé mais EXE non trouvé à l'emplacement attendu: $exePath"
}

Write-Host "`nExemples:" -ForegroundColor Yellow
Write-Host "  .\build_exe.ps1" -ForegroundColor Yellow
Write-Host "  .\build_exe.ps1 -Mode Script -EntryScript .\mon_script.py -ExeName MonOutil" -ForegroundColor Yellow
Write-Host "  .\build_exe.ps1 -Host 0.0.0.0 -Port 8090 -OpenBrowser:$true" -ForegroundColor Yellow
