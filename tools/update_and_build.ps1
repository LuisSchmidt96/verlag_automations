<#
    update_and_build.ps1
    --------------------
    Holt die neuesten Aenderungen aus dem Git-Repo und baut alle Tools neu.

    Ergebnis-Layout im gemeinsamen VR-Tools-Ordner (eine Ebene UEBER dem Repo),
    damit alle fertigen Programme ordentlich nebeneinander liegen:

        VR-Tools\
        |-- repo\                 <- dieses Git-Repo (Quellcode)
        |-- BooxpressEtiketten\   <- fertiges Tool  (Doppelklick auf die .exe)
        |-- PiBiGenerator\        <- fertiges Tool
        `-- CoverPreviews\        <- fertiges Tool  (inkl. _NEU_Vorlage\)

    Jedes Tool hat eine eigene *.spec-Datei in seinem Unterordner. Das Skript
    findet automatisch ALLE *.spec-Dateien im Repo und baut sie einzeln - direkt
    als Ordner in VR-Tools ("--distpath"). Neue Automations werden automatisch
    mitgebaut, sobald sie eine .spec-Datei mitbringen.

    Jeder fertige Ordner ist fuer sich lauffaehig: kopieren, auf einen anderen
    PC legen, .exe starten. Liegt im Tool-Ordner eine Anleitung.txt, wird sie
    mitgeliefert. Grosse Nutzdaten, die nicht in die .exe gehoeren, stehen in
    $Beigaben - fuer CoverPreviews sind das die Mockup-Vorlagen (_NEU_Vorlage,
    rund 460 MB), die neben der .exe liegen muessen.

    Die virtuelle Umgebung und die PyInstaller-Zwischendateien liegen bewusst
    LOKAL (%LOCALAPPDATA%), nicht auf dem Netzlaufwerk: schneller, portabel
    (eine venv haengt am Python-Pfad des jeweiligen Rechners) und haelt den
    gemeinsamen VR-Tools-Ordner sauber.

    Aufruf (aus PowerShell):
        .\tools\update_and_build.ps1

    Voraussetzung: Python 3.12 (inkl. tkinter) und git im PATH.
#>

$ErrorActionPreference = 'Stop'

# --- Pfade bestimmen --------------------------------------------------------
$RepoRoot = Split-Path -Parent $PSScriptRoot   # ...\VR-Tools\repo
$OutRoot  = Split-Path -Parent $RepoRoot       # ...\VR-Tools  (Ziel der Tools)
Write-Host "Repo:    $RepoRoot" -ForegroundColor Cyan
Write-Host "Ausgabe: $OutRoot"  -ForegroundColor Cyan

# --- 1) Neueste Aenderungen holen ------------------------------------------
Write-Host "`n[1/4] git pull ..." -ForegroundColor Cyan
git -C $RepoRoot pull --ff-only

# --- 2) Virtuelle Umgebung sicherstellen (LOKAL, nicht auf dem Share) -------
Write-Host "`n[2/4] Virtuelle Umgebung / Abhaengigkeiten ..." -ForegroundColor Cyan
$BuildHome  = Join-Path $env:LOCALAPPDATA 'verlag_automations_build'
$VenvDir    = Join-Path $BuildHome '.venv'
$WorkPath   = Join-Path $BuildHome 'build'
$VenvPython = Join-Path $VenvDir 'Scripts\python.exe'

if (-not (Test-Path $VenvPython)) {
    Write-Host "  Erstelle .venv (Python 3.12) unter $VenvDir ..."
    if (Get-Command py -ErrorAction SilentlyContinue) {
        py -3.12 -m venv $VenvDir
    } else {
        python -m venv $VenvDir
    }
}
& $VenvPython -m pip install --upgrade pip --quiet
& $VenvPython -m pip install -r (Join-Path $RepoRoot 'requirements.txt') --quiet

# --- 3) Alle Tools bauen (LOKAL) und nach VR-Tools spiegeln -----------------
# ACHTUNG: PyInstaller loescht bei --noconfirm den KOMPLETTEN Ziel-Ordner.
# Wuerde man direkt auf den Share bauen, waeren die neben der .exe liegenden
# Nutzdaten (kommliste.xlsx, config.json, paketnr.txt, etiketten_output\) bei
# jedem Update WEG. Darum: lokal in einen Stage-Ordner bauen und dann nur die
# Programmteile uebertragen:
#   * _internal\  = reine PyInstaller-Ausgabe  -> spiegeln (/MIR)
#   * *.exe       = das Programm selbst        -> ueberschreiben
# Alles andere (die Nutzdaten des Anwenders) bleibt auf dem Share unangetastet.
Write-Host "`n[3/4] Tools bauen (lokal) und nach $OutRoot spiegeln ..." -ForegroundColor Cyan
Set-Location $BuildHome
$StageDir = Join-Path $BuildHome 'dist'
$Specs = Get-ChildItem -Path $RepoRoot -Recurse -Filter *.spec |
         Where-Object { $_.FullName -notmatch '\\(build|dist|\.venv)\\' }

if (-not $Specs) {
    Write-Warning "Keine *.spec-Dateien gefunden - nichts zu bauen."
    return
}

# Grosse Nutzdaten, die NICHT in die .exe gehoeren, aber im fertigen Ordner
# liegen muessen, damit er sich einfach kopieren laesst. Es gewinnt die erste
# Quelle, die es gibt.
$Beigaben = @{
    'CoverPreviews' = @{
        Ordner  = '_NEU_Vorlage'                 # muss neben der .exe liegen
        Quellen = @(
            '\\C019\d\Online\Webseite\Artikeldaten\_NEU_Vorlage',   # Original
            (Join-Path $RepoRoot 'cover_previews\_NEU_Vorlage')     # lokale Kopie
        )
    }
}

foreach ($Spec in $Specs) {
    $Name = $Spec.BaseName                       # == COLLECT-Name in der .spec
    Write-Host "  -> $($Spec.Name)" -ForegroundColor Yellow
    & $VenvPython -m PyInstaller --noconfirm --log-level WARN `
        --distpath $StageDir --workpath $WorkPath $Spec.FullName
    if ($LASTEXITCODE -ne 0) { throw "PyInstaller-Build fehlgeschlagen: $Name" }

    $Src = Join-Path $StageDir $Name
    $Dst = Join-Path $OutRoot  $Name
    New-Item -ItemType Directory -Force -Path $Dst | Out-Null
    # Programmteile spiegeln; robocopy-Exitcodes 0-7 sind Erfolg (kein throw).
    robocopy (Join-Path $Src '_internal') (Join-Path $Dst '_internal') /MIR /NFL /NDL /NJH /NJS /NP | Out-Null
    Copy-Item (Join-Path $Src '*.exe') $Dst -Force

    # Anleitung fuer den Anwender mitliefern, falls das Tool eine hat.
    $Anleitung = Join-Path $Spec.Directory 'Anleitung.txt'
    if (Test-Path $Anleitung) { Copy-Item $Anleitung $Dst -Force }

    # Beigaben (z. B. die Mockup-Vorlagen) in den fertigen Ordner spiegeln.
    $B = $Beigaben[$Name]
    if ($B) {
        $Quelle = $B.Quellen | Where-Object { Test-Path $_ } | Select-Object -First 1
        if ($Quelle) {
            $Ziel = Join-Path $Dst $B.Ordner
            Write-Host "     $($B.Ordner) aus $Quelle ..." -ForegroundColor DarkGray
            # /MIR: nur geaenderte Dateien laufen ueber die Leitung, der zweite
            # Build ist also schnell. Rund 460 MB beim ersten Mal.
            robocopy $Quelle $Ziel /MIR /NFL /NDL /NJH /NJS /NP | Out-Null
            if ($LASTEXITCODE -ge 8) { throw "Konnte $($B.Ordner) nicht kopieren." }
        } else {
            Write-Warning "  $Name : $($B.Ordner) nicht gefunden - der Ordner muss von Hand neben die .exe gelegt werden."
        }
    }
}

# --- 4) Ergebnis ------------------------------------------------------------
Write-Host "`n[4/4] Fertige Tools unter $OutRoot :" -ForegroundColor Green
foreach ($Spec in $Specs) {
    $ToolDir = Join-Path $OutRoot $Spec.BaseName   # COLLECT-Name == .spec-Basisname
    if (Test-Path $ToolDir) { Write-Host "  $ToolDir" }
}

Write-Host "`nHinweis: Die Tools legen config.json & Co. direkt neben der .exe an"  -ForegroundColor DarkGray
Write-Host "(kein data-Unterordner). BooxpressEtiketten braucht dort zusaetzlich"    -ForegroundColor DarkGray
Write-Host "die kommliste.xlsx (Stammdaten), CoverPreviews den Ordner _NEU_Vorlage"  -ForegroundColor DarkGray
Write-Host "(wird oben automatisch mitkopiert)."                                     -ForegroundColor DarkGray
