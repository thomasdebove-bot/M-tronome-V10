# M-tronome-V10

## Build Windows (EXE)

Le projet inclut deux scripts pour créer l'exécutable :

- `build_exe.ps1` : script PowerShell principal (PyInstaller).
- `run_build_exe.bat` : lanceur recommandé sous Windows.

### Pourquoi utiliser `run_build_exe.bat` ?

Si PowerShell rencontre une erreur, la fenêtre peut se fermer trop vite quand on lance
directement le `.ps1`. Le `.bat` lance PowerShell avec `-NoExit` pour garder la fenêtre
ouverte et permettre de lire le message `[ERREUR]` et le détail.

### Utilisation

Depuis l'explorateur Windows :

1. Double-cliquer sur `run_build_exe.bat`.
2. Lire le résultat affiché (succès ou erreur).

Depuis un terminal :

```bat
run_build_exe.bat
```

Paramètres transmis au script PowerShell (exemple) :

```bat
run_build_exe.bat -Mode Script -EntryScript .\mon_script.py -ExeName MonOutil
```
