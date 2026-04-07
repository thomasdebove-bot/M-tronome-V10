@echo off
setlocal
cd /d "%~dp0"

echo ==============================================
echo  TEMPO build EXE launcher (PowerShell)
echo ==============================================
echo.
echo Si vous voyez l'avertissement de securite une seule fois,
echo cliquez sur "O" (Executer une fois).
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -NoExit -File "%~dp0build_exe.ps1" %*

if errorlevel 1 (
  echo.
  echo Le build a retourne une erreur. Verifiez les lignes [ERREUR] ci-dessus.
)

echo.
echo Appuyez sur une touche pour fermer cette fenetre...
pause >nul
endlocal
