param(
    [string]$PythonExe = "python",
    [string]$EntryScript = "app-V10-2.py",
    [string]$ExeName = "TEMPO-CR",
    [ValidateSet("FastAPI", "Script")]
    [string]$Mode = "FastAPI",
    [string]$AppVarName = "app",
    [Alias("Host")]
    [string]$BindHost = "127.0.0.1",
    [int]$Port = 8090,
    [bool]$OpenBrowser = $true,
    [switch]$OneFile = $true,
    [switch]$Clean = $true,
    [switch]$NoConsole = $false,
    [switch]$PauseOnExit = $true
)

$ErrorActionPreference = "Stop"
$script:ExitCode = 0

function Step([string]$msg) {
    Write-Host "`n==> $msg" -ForegroundColor Cyan
}

function Fail([string]$msg) {
    Write-Host "`n[ERREUR] $msg" -ForegroundColor Red
}

function Resolve-EntryPath([string]$PathValue) {
    if ([System.IO.Path]::IsPathRooted($PathValue)) {
        return (Resolve-Path $PathValue).Path
    }
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $candidate = Join-Path $scriptDir $PathValue
    return (Resolve-Path $candidate).Path
}

try {
    $entryAbs = Resolve-EntryPath $EntryScript
    Step "Script d'entrée: $entryAbs"

    Step "Vérification Python"
    & $PythonExe --version
    if ($LASTEXITCODE -ne 0) {
        throw "Python introuvable. Vérifiez -PythonExe (actuel: $PythonExe)."
    }

    Step "Vérification de PyInstaller"
    & $PythonExe -m PyInstaller --version 2>$null
    if ($LASTEXITCODE -ne 0) {
        Step "Installation de PyInstaller"
        & $PythonExe -m pip install --upgrade pyinstaller
        if ($LASTEXITCODE -ne 0) {
            throw "Impossible d'installer PyInstaller (pip a échoué)."
        }
    }

    $flags = @()
    if ($OneFile) { $flags += "--onefile" }
    if ($Clean) { $flags += "--clean" }
    if ($NoConsole) { $flags += "--noconsole" }

    $buildTarget = $entryAbs
    $tempLauncher = ""

    if ($Mode -eq "FastAPI") {
        Step "Génération d'un launcher FastAPI"

        $entryEscaped = $entryAbs.Replace("\\", "\\\\")
        $launcherCode = @"
import importlib.util
import pathlib
import sys
import uvicorn

ENTRY_PATH = pathlib.Path(r"$entryEscaped")
APP_VAR = "$AppVarName"
HOST = "$BindHost"
PORT = $Port
OPEN_BROWSER = $OpenBrowser

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
            Write-Host "Au lancement, l'EXE démarre le serveur sur http://$BindHost`:$Port" -ForegroundColor Green
        }
    }
    else {
        throw "Build terminé mais EXE non trouvé: $exePath"
    }

    Write-Host "`nExemples:" -ForegroundColor Yellow
    Write-Host "  .\build_exe.ps1" -ForegroundColor Yellow
    Write-Host "  .\build_exe.ps1 -Mode Script -EntryScript .\mon_script.py -ExeName MonOutil" -ForegroundColor Yellow
    Write-Host "  .\build_exe.ps1 -BindHost 0.0.0.0 -Port 8090 -OpenBrowser:$true" -ForegroundColor Yellow
}
catch {
    Fail $_.Exception.Message
    Write-Host "Détail:" -ForegroundColor DarkYellow
    Write-Host $_.Exception.ToString()
    $script:ExitCode = 1
}
finally {
    if ($PauseOnExit) {
        Write-Host "`nAppuyez sur Entrée pour fermer..." -ForegroundColor DarkGray
        [void](Read-Host)
    }
}

if ($script:ExitCode -ne 0) {
    exit $script:ExitCode
}
