param(
    [string]$PythonExe = "python",
    [string]$EntryScript = "app-V10-2.py",
    [string]$ExeName = "TEMPO-CR",
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
$pyiCheck = & $PythonExe -m PyInstaller --version 2>$null
if ($LASTEXITCODE -ne 0) {
    Step "Installation de PyInstaller"
    & $PythonExe -m pip install --upgrade pyinstaller
}

$flags = @()
if ($OneFile) { $flags += "--onefile" }
if ($Clean) { $flags += "--clean" }
if ($NoConsole) { $flags += "--noconsole" }

Step "Build EXE"
& $PythonExe -m PyInstaller @flags --name $ExeName $EntryScript
if ($LASTEXITCODE -ne 0) {
    throw "Echec de compilation PyInstaller"
}

$exePath = Join-Path (Join-Path (Get-Location) "dist") "$ExeName.exe"
if (Test-Path $exePath) {
    Step "Build terminé"
    Write-Host "EXE généré: $exePath" -ForegroundColor Green
} else {
    Write-Warning "Build terminé mais EXE non trouvé à l'emplacement attendu: $exePath"
}

Write-Host "`nAstuce: utilisez -NoConsole pour masquer la console et -OneFile:$false pour un build dossier." -ForegroundColor Yellow
