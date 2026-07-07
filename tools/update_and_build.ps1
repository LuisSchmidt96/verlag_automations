<#
    update_and_build.ps1
    --------------------
    Holt die neuesten Aenderungen aus dem Git-Repo und baut alle Tools neu.

    Jedes Tool hat eine eigene *.spec-Datei in seinem Unterordner. Das Skript
    findet automatisch ALLE *.spec-Dateien im Repo und baut sie einzeln.
    Neue Automations werden also automatisch mitgebaut, sobald sie eine
    .spec-Datei mitbringen.

    Aufruf (aus PowerShell):
        .\tools\update_and_build.ps1

    Voraussetzung: Python 3.12 (inkl. tkinter) und git im PATH.
#>

$ErrorActionPreference = 'Stop'

# --- Ins Repo-Wurzelverzeichnis wechseln (ein Ordner ueber diesem Skript) ---
$RepoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $RepoRoot
Write-Host "Repo: $RepoRoot" -ForegroundColor Cyan

# --- 1) Neueste Aenderungen holen ------------------------------------------
Write-Host "`n[1/4] git pull ..." -ForegroundColor Cyan
git pull --ff-only

# --- 2) Virtuelle Umgebung sicherstellen -----------------------------------
Write-Host "`n[2/4] Virtuelle Umgebung / Abhaengigkeiten ..." -ForegroundColor Cyan
$VenvPython = Join-Path $RepoRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $VenvPython)) {
    Write-Host "  Erstelle .venv ..."
    python -m venv .venv
}
& $VenvPython -m pip install --upgrade pip --quiet
& $VenvPython -m pip install -r requirements.txt --quiet

# --- 3) Alle Tools bauen ----------------------------------------------------
Write-Host "`n[3/4] Tools bauen ..." -ForegroundColor Cyan
$Specs = Get-ChildItem -Path $RepoRoot -Recurse -Filter *.spec |
         Where-Object { $_.FullName -notmatch '\\(build|dist|\.venv)\\' }

if (-not $Specs) {
    Write-Warning "Keine *.spec-Dateien gefunden - nichts zu bauen."
    return
}

foreach ($Spec in $Specs) {
    Write-Host "  -> $($Spec.Name)" -ForegroundColor Yellow
    & $VenvPython -m PyInstaller --noconfirm $Spec.FullName
}

# --- 4) Ergebnis ------------------------------------------------------------
Write-Host "`n[4/4] Fertige Programme unter dist\ :" -ForegroundColor Green
Get-ChildItem -Path (Join-Path $RepoRoot 'dist') -Directory |
    ForEach-Object { Write-Host "  $($_.FullName)" }

Write-Host "`nHinweis: Jedes Tool erwartet neben der .exe einen Ordner 'data\'" -ForegroundColor DarkGray
Write-Host "(config.json, paketnr.txt, kommliste.xlsx). Beim ersten Start"      -ForegroundColor DarkGray
Write-Host "wird eine Standard-config.json angelegt, falls noch keine da ist."   -ForegroundColor DarkGray
