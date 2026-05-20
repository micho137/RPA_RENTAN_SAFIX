param(
    [switch]$Clean
)

$ErrorActionPreference = "Stop"

if ($Clean) {
    if (Test-Path ".\build") { Remove-Item ".\build" -Recurse -Force }
    if (Test-Path ".\dist") { Remove-Item ".\dist" -Recurse -Force }
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$venvPython = Join-Path $repoRoot "venv\Scripts\python.exe"
$pythonExe = if (Test-Path $venvPython) { $venvPython } else { "python" }

$tkCheck = @"
import tkinter
import _tkinter
print(tkinter.TkVersion)
"@

& $pythonExe -c $tkCheck | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "Tkinter no esta disponible en $pythonExe. Reinstala Python con Tcl/Tk y vuelve a crear el venv."
}

$tkDataCheck = @"
import pathlib
import sys
base = pathlib.Path(sys.base_prefix)
tcl = base / 'tcl' / 'tcl8.6'
tk = base / 'tcl' / 'tk8.6'
ok = tcl.exists() and tk.exists()
print(str(tcl), str(tk), ok)
raise SystemExit(0 if ok else 1)
"@

& $pythonExe -c $tkDataCheck | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "No se encontraron tcl8.6/tk8.6 en el Python base de $pythonExe."
}

& $pythonExe -m PyInstaller --noconfirm --clean .\AIVO.spec
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller fallo con codigo $LASTEXITCODE."
}

Write-Host ""
Write-Host "Build listo: .\dist\AIVO\AIVO.exe"
