<#
    update_and_build.ps1
    --------------------
    Holt die neuesten Aenderungen aus dem Git-Repo und baut nur die Tools neu,
    an denen sich seit dem letzten Build etwas geaendert hat.

    Welche Tools gebaut werden, entscheidet ein Baustempel je Tool: die Git-SHA,
    aus der es zuletzt erfolgreich gebaut wurde (unter $StampDir, lokal neben der
    venv). Ein Tool wird neu gebaut, wenn seit seinem Stempel etwas in seinem
    eigenen Paket ODER in einem Paket, das es benutzt, ODER in requirements.txt
    geaendert wurde - oder wenn sein fertiger Ordner fehlt. Buchdurchgang haengt
    an cover_previews, pi_bi_generator und shopware_publisher (siehe $Deps);
    aendert sich eines davon, wird Buchdurchgang mitgebaut.

    Der Baustempel misst gegen HEAD NACH dem Pull, nicht nur gegen das, was der
    Pull gebracht hat - so werden auch lokale Commits erfasst, die schon da
    waren. Mit -Force werden alle Tools gebaut, egal was sich geaendert hat.

    Ergebnis-Layout im MASTER-Ordner auf dem NAS ($MasterRoot). Von dort holen
    die Clients ihre Fassung ab; einen zweiten Ablageort gibt es nicht mehr:

        \\VR-Archiv\VR-Austausch\VR-Tools\
        |-- repo\                 <- dieses Git-Repo (Quellcode)
        |-- launch.ps1            <- Launcher, den die Verknuepfungen aufrufen
        |-- Einrichten.cmd        <- legt die Verknuepfungen an
        |-- _NEU_Vorlage\         <- Mockup-Vorlagen, EINE Kopie fuer alle
        |-- BooxpressEtiketten\   <- fertiges Tool
        |-- PiBiGenerator\        <- fertiges Tool
        `-- CoverPreviews\        <- fertiges Tool

    Jedes Tool hat eine eigene *.spec-Datei in seinem Unterordner. Das Skript
    findet automatisch ALLE *.spec-Dateien im Repo und baut sie einzeln - direkt
    als Ordner in VR-Tools ("--distpath"). Neue Automations werden automatisch
    mitgebaut, sobald sie eine .spec-Datei mitbringen.

    Jeder fertige Ordner ist fuer sich lauffaehig: kopieren, auf einen anderen
    PC legen, .exe starten. Liegt im Tool-Ordner eine Anleitung.txt, wird sie
    mitgeliefert. Grosse Nutzdaten, die nicht in die .exe gehoeren, stehen in
    $Beigaben - fuer CoverPreviews sind das die Mockup-Vorlagen (_NEU_Vorlage,
    rund 460 MB), die neben der .exe liegen muessen.

    Der Launcher (launch.ps1 / einrichten.ps1 / Einrichten.cmd) wird mit
    veroeffentlicht und liegt neben den Tool-Ordnern. Die Kollegen kopieren
    nichts von Hand: ihre Verknuepfung ruft launch.ps1 auf, das die neue
    Fassung beim Start abholt und dann startet. Veroeffentlichen = dieses
    Skript laufen lassen, mehr nicht.

    FRUEHER lag der Master auf \\C019\d\VR-Tools und jedes Tool wurde zweimal
    uebertragen (einmal neben das Repo, einmal auf C019). Beides faellt weg:
    Master ist der NAS. Verknuepfungen, die noch auf C019 zeigen, bekommen
    KEINE Aktualisierungen mehr - dort muss einmal Einrichten.cmd vom NAS
    laufen.

    Uebertragen werden dabei nur die PROGRAMMTEILE (.exe, _internal\,
    Anleitung.txt, Beigaben) - nicht die Nutzdaten daneben (config.json,
    kommliste.xlsx, paketnr.txt, Ausgabeordner). Die gehoeren dem jeweiligen
    Anwender und bleiben unangetastet; launch.ps1 zieht dieselbe Grenze in die
    andere Richtung.

    Die virtuelle Umgebung und die PyInstaller-Zwischendateien liegen bewusst
    LOKAL (%LOCALAPPDATA%), nicht auf dem Netzlaufwerk: schneller, portabel
    (eine venv haengt am Python-Pfad des jeweiligen Rechners) und haelt den
    gemeinsamen VR-Tools-Ordner sauber.

    Aufruf (aus PowerShell):
        .\tools\update_and_build.ps1

    Voraussetzung: Python 3.12 (inkl. tkinter) und git im PATH.
#>

param(
    # Alle Tools bauen, auch die unveraenderten.
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

# --- Pfade bestimmen --------------------------------------------------------
$RepoRoot = Split-Path -Parent $PSScriptRoot    # ...\VR-Tools\repo

# Wohin veroeffentlicht wird. FEST verdrahtet und bewusst NICHT aus $RepoRoot
# abgeleitet: es liegen mehrere Klone herum (u. a. noch einer auf C019), und
# ein Lauf aus dem falschen Klon wuerde sonst neben DIESEN Klon legen - die
# Kollegen holten weiter die alte Fassung ab. Muss mit SHARE_VORLAGEN in
# cover_previews/core.py und buchdurchgang/core.py uebereinstimmen; dort steht
# derselbe Pfad fest im Code.
$MasterRoot = '\\VR-Archiv\VR-Austausch\VR-Tools'

Write-Host "Repo:   $RepoRoot"   -ForegroundColor Cyan
Write-Host "Master: $MasterRoot" -ForegroundColor Cyan

if (-not (Test-Path $MasterRoot)) {
    throw ("Der Master-Ordner $MasterRoot ist nicht erreichbar. " +
           "Ohne ihn hat das Veroeffentlichen keinen Zweck - bitte erst mit " +
           "dem Netzlaufwerk verbinden.")
}

# Gebaut wird immer der Stand DIESES Repos. Wer aus einem anderen Klon baut,
# soll das wenigstens sehen.
$RepoErwartet = Join-Path $MasterRoot 'repo'
if ($RepoRoot.TrimEnd('\') -ne $RepoErwartet.TrimEnd('\')) {
    Write-Warning "Gebaut wird aus $RepoRoot, veroeffentlicht nach $MasterRoot."
    Write-Warning "Das ist nicht der Klon neben dem Master ($RepoErwartet)."
}

# --- 1) Neueste Aenderungen holen ------------------------------------------
Write-Host "`n[1/4] git pull ..." -ForegroundColor Cyan
$VorPull = (git -C $RepoRoot rev-parse HEAD).Trim()
git -C $RepoRoot pull --ff-only
$NachPull = (git -C $RepoRoot rev-parse HEAD).Trim()
if ($VorPull -ne $NachPull) {
    Write-Host "  $($VorPull.Substring(0,7)) -> $($NachPull.Substring(0,7))" -ForegroundColor DarkGray
} else {
    Write-Host "  schon aktuell ($($NachPull.Substring(0,7)))" -ForegroundColor DarkGray
}

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
Write-Host "`n[3/4] Tools bauen (lokal) und nach $MasterRoot spiegeln ..." -ForegroundColor Cyan
Set-Location $BuildHome
$StageDir = Join-Path $BuildHome 'dist'
$Specs = Get-ChildItem -Path $RepoRoot -Recurse -Filter *.spec |
         Where-Object { $_.FullName -notmatch '\\(build|dist|\.venv)\\' }

if (-not $Specs) {
    Write-Warning "Keine *.spec-Dateien gefunden - nichts zu bauen."
    return
}

# Welches Tool haengt an welchen fremden Paketen? Nur Buchdurchgang benutzt
# andere Werkzeuge (es importiert cover_previews, pi_bi_generator und
# shopware_publisher); alle uebrigen sind fuer sich. Aendert sich ein
# benutztes Paket, muss der Abhaengige mitgebaut werden. Schluessel ist der
# Paket-Ordnername (== $Spec.Directory.Name), wie er in den Git-Pfaden steht.
$Deps = @{
    'buchdurchgang' = @('cover_previews', 'pi_bi_generator', 'shopware_publisher')
}

# Baustempel je Tool (Git-SHA des letzten erfolgreichen Builds), lokal neben
# der venv. Bleibt ein Build aus, behaelt das Tool seinen alten Stempel und
# wird beim naechsten Mal erneut geprueft.
$StampDir = Join-Path $BuildHome 'build_stamps'
New-Item -ItemType Directory -Force -Path $StampDir | Out-Null

# git diff je Ausgangs-SHA nur einmal rechnen (die Tools teilen sich meist
# denselben Stempel). $null bedeutet: SHA unbekannt (z. B. durch gc) -> neu bauen.
$DiffCache = @{}
function Get-GeaenderteDateien([string]$Von, [string]$Bis) {
    if (-not $DiffCache.ContainsKey($Von)) {
        $d = git -C $RepoRoot diff --name-only $Von $Bis 2>$null
        if ($LASTEXITCODE -ne 0) { $DiffCache[$Von] = $null }
        else { $DiffCache[$Von] = @($d | Where-Object { $_ }) }
    }
    return $DiffCache[$Von]
}

# Entscheidet (und begruendet), ob ein Tool gebaut werden muss.
function Get-BauGrund {
    param([string]$PkgName, [string]$OutDir, [string]$Stempel)

    if ($Force)                   { return 'erzwungen (-Force)' }
    if (-not (Test-Path $OutDir)) { return 'noch nicht veroeffentlicht' }
    if (-not $Stempel)            { return 'kein Baustempel' }

    $diff = Get-GeaenderteDateien $Stempel $NachPull
    if ($null -eq $diff)            { return 'Baustempel unbekannt' }
    if ($diff.Count -eq 0)          { return $null }        # nichts geaendert
    if ($diff -contains 'requirements.txt') { return 'requirements.txt geaendert' }

    # Beobachtet: das eigene Paket + alle benutzten Pakete.
    $beobachtet = @($PkgName) + $Deps[$PkgName]
    $treffer = $diff |
        ForEach-Object { ($_ -split '/')[0] } |
        Where-Object { $beobachtet -contains $_ } |
        Sort-Object -Unique
    if ($treffer) { return "geaendert: $($treffer -join ', ')" }
    return $null
}

# Die Mockup-Vorlagen (~480 MB) gehoeren nicht in die .exe, muessen aber
# erreichbar sein. Sie liegen EINMAL auf dem Share neben den Tool-Ordnern -
# nicht mehr in jedem Werkzeugordner und damit auch nicht in jeder lokalen
# Kopie jedes Anwenders. CoverPreviews und Buchdurchgang kennen diesen Pfad
# fest im Code (SHARE_VORLAGEN); ueber die config.json ginge es nicht, die
# wird bewusst nicht gespiegelt und gaelte nur auf einem Rechner.
# Seit der Master der NAS ist, ist das derselbe Ordner - frueher zeigte
# $ShareRoot auf C019, und die Vorlagen mussten getrennt verdrahtet werden.
# Muss weiterhin mit SHARE_VORLAGEN in cover_previews/core.py und
# buchdurchgang/core.py uebereinstimmen: dort steht derselbe Pfad fest im
# Code, und der wird NICHT von hier abgeleitet.
$VorlagenWurzel  = $MasterRoot
$VorlagenZiel    = Join-Path $VorlagenWurzel '_NEU_Vorlage'
$VorlagenQuellen = @(
    '\\C019\d\Online\Webseite\Artikeldaten\_NEU_Vorlage',       # Original
    (Join-Path $RepoRoot 'cover_previews\_NEU_Vorlage'),        # lokale Kopie
    # Umzugshilfe: bis hierher lagen die Vorlagen IM CoverPreviews-Ordner.
    # So holt der erste Lauf sie von dort an den gemeinsamen Ort, auch wenn
    # weder C019 verbunden noch eine lokale Kopie da ist.
    (Join-Path $VorlagenWurzel 'CoverPreviews\_NEU_Vorlage')
)

# Legt NUR die Programmteile ab: .exe, _internal\, Anleitung.txt, Beigaben.
# Bewusst NICHT die Nutzdaten daneben (config.json, kommliste.xlsx, paketnr.txt,
# Ausgabeordner): die gehoeren dem jeweiligen Anwender. Wuerde man den ganzen
# Ordner spiegeln, landete die eigene config.json bei allen anderen - und /MIR
# wuerde die kommliste.xlsx auf dem Share loeschen.
function Copy-Programmteile {
    param([string]$Src, [string]$Dst, [string]$Anleitung)

    New-Item -ItemType Directory -Force -Path $Dst | Out-Null
    # robocopy-Exitcodes 0-7 sind Erfolg, erst ab 8 ist es ein Fehler.
    robocopy (Join-Path $Src '_internal') (Join-Path $Dst '_internal') /MIR /NFL /NDL /NJH /NJS /NP | Out-Null
    if ($LASTEXITCODE -ge 8) { throw "Konnte _internal nicht nach $Dst kopieren." }
    Copy-Item (Join-Path $Src '*.exe') $Dst -Force

    if ($Anleitung -and (Test-Path $Anleitung)) { Copy-Item $Anleitung $Dst -Force }
}

$Gebaut        = @()
$Uebersprungen = @()

foreach ($Spec in $Specs) {
    $Name    = $Spec.BaseName                    # == COLLECT-Name in der .spec
    $PkgName = $Spec.Directory.Name              # Paket-Ordner, wie in Git-Pfaden
    $OutDir    = Join-Path $MasterRoot $Name
    $StampFile = Join-Path $StampDir "$Name.sha"
    $Stempel   = if (Test-Path $StampFile) { (Get-Content $StampFile -Raw).Trim() } else { '' }

    $Grund = Get-BauGrund -PkgName $PkgName -OutDir $OutDir -Stempel $Stempel
    if (-not $Grund) {
        Write-Host "  = $($Spec.Name) unveraendert - uebersprungen" -ForegroundColor DarkGray
        $Uebersprungen += $Name
        continue
    }

    Write-Host "  -> $($Spec.Name)  ($Grund)" -ForegroundColor Yellow
    & $VenvPython -m PyInstaller --noconfirm --log-level WARN `
        --distpath $StageDir --workpath $WorkPath $Spec.FullName
    if ($LASTEXITCODE -ne 0) { throw "PyInstaller-Build fehlgeschlagen: $Name" }

    $Src       = Join-Path $StageDir $Name
    $Anleitung = Join-Path $Spec.Directory 'Anleitung.txt'

    # Einmal uebertragen, nicht zweimal: Master und Ablageort sind derselbe
    # Ordner. Von hier holt der Launcher die neue Fassung ab.
    Copy-Programmteile -Src $Src -Dst $OutDir -Anleitung $Anleitung

    # Erst nach erfolgreichem Build + Kopieren stempeln: bricht etwas vorher ab,
    # bleibt der alte Stempel stehen und der naechste Lauf versucht es erneut.
    Set-Content -Path $StampFile -Value $NachPull -Encoding ASCII
    $Gebaut += $Name
}

if (-not $Gebaut) {
    Write-Host "  Nichts zu bauen - alle Tools sind aktuell." -ForegroundColor Green
}

# --- 3a) Mockup-Vorlagen: EINE Kopie fuer alle Werkzeuge --------------------
if (Test-Path $VorlagenWurzel) {
    $VorlagenQuelle = $VorlagenQuellen | Where-Object { Test-Path $_ } | Select-Object -First 1
    if ($VorlagenQuelle) {
        Write-Host "  -> _NEU_Vorlage -> $VorlagenZiel" -ForegroundColor Yellow
        # /MIR uebertraegt nur Geaendertes; nur der erste Lauf kostet Zeit.
        robocopy $VorlagenQuelle $VorlagenZiel /MIR /NFL /NDL /NJH /NJS /NP | Out-Null
        if ($LASTEXITCODE -ge 8) { throw "Konnte _NEU_Vorlage nicht nach $VorlagenZiel kopieren." }
    } elseif (-not (Test-Path $VorlagenZiel)) {
        # Nur warnen, wenn auf dem Share auch nichts liegt: sind die Vorlagen
        # dort schon aktuell, ist eine fehlende Quelle kein Problem.
        Write-Warning "  _NEU_Vorlage nicht gefunden - der 3D-Zweig findet keine Vorlage."
        Write-Warning "  Erwartet unter: $VorlagenZiel"
    }
} else {
    Write-Warning "  $VorlagenWurzel nicht erreichbar - Mockup-Vorlagen nicht veroeffentlicht."
}

# --- 3b) Launcher veroeffentlichen -----------------------------------------
# launch.ps1 haelt die Tools auf den Client-PCs aktuell, Einrichten.cmd legt
# dort einmalig die Verknuepfungen an. Beide muessen NEBEN den Tool-Ordnern auf
# dem Share liegen: launch.ps1 leitet den Master-Pfad aus dem eigenen Ort ab
# ($PSScriptRoot), damit der Servername nirgends fest verdrahtet ist.
Write-Host "  -> Launcher (launch.ps1, Einrichten.cmd)" -ForegroundColor Yellow
foreach ($Datei in @('launch.ps1', 'einrichten.ps1', 'Einrichten.cmd')) {
    Copy-Item (Join-Path $PSScriptRoot $Datei) $MasterRoot -Force
}

# --- 3c) Veroeffentlichten Stand festhalten --------------------------------
# launch.ps1 vergleicht das auf JEDEM Rechner mit dem HEAD des Repos und weist
# hin, wenn jemand committet, aber nicht veroeffentlicht hat. Ohne diese Datei
# faellt so etwas erst auf, wenn sich jemand ueber eine fehlende Aenderung
# wundert.
#
# Bewusst EIN Eintrag fuer den ganzen Lauf und keiner je Tool: die Baustempel
# ueberspringen unveraenderte Werkzeuge, deren Stempel bleibt also absichtlich
# alt. Ein Vergleich je Tool schluege deshalb dauerhaft an, auch wenn alles
# veroeffentlicht ist. Dieser Eintrag heisst nur: "bis hierher wurde alles
# betrachtet".
#
# Geschrieben wird erst hier, nach Bauen und Kopieren: bricht vorher etwas ab,
# bleibt der alte Stand stehen und der Hinweis erscheint weiter - richtig so.
Set-Content -Path (Join-Path $MasterRoot '.veroeffentlicht') `
            -Value $NachPull -Encoding ASCII

# --- 4) Ergebnis ------------------------------------------------------------
Write-Host "`n[4/4] Ergebnis:" -ForegroundColor Green
if ($Gebaut) {
    Write-Host "  Neu gebaut: $($Gebaut -join ', ')" -ForegroundColor Green
}
if ($Uebersprungen) {
    Write-Host "  Unveraendert: $($Uebersprungen -join ', ')" -ForegroundColor DarkGray
}
if ($Gebaut) {
    Write-Host "`nVeroeffentlicht auf $MasterRoot :" -ForegroundColor Green
    foreach ($Name in $Gebaut) {
        $ToolDir = Join-Path $MasterRoot $Name
        if (Test-Path $ToolDir) { Write-Host "  $ToolDir" }
    }
    Write-Host "`nDie Kollegen bekommen das beim naechsten Start automatisch," -ForegroundColor Green
    Write-Host "SOFERN ihre Verknuepfung auf $MasterRoot zeigt." -ForegroundColor Green
    Write-Host "Neuer PC (und jeder, dessen Verknuepfung noch auf C019 zeigt):" -ForegroundColor Green
    Write-Host "  einmal $MasterRoot\Einrichten.cmd doppelklicken."            -ForegroundColor Green
}

Write-Host "`nHinweis: Die Tools legen config.json & Co. direkt neben der .exe an"  -ForegroundColor DarkGray
Write-Host "(kein data-Unterordner). BooxpressEtiketten braucht dort zusaetzlich"    -ForegroundColor DarkGray
Write-Host "die kommliste.xlsx (Stammdaten). Die Mockup-Vorlagen liegen EINMAL"     -ForegroundColor DarkGray
Write-Host "unter $VorlagenZiel - CoverPreviews und Buchdurchgang finden sie dort"   -ForegroundColor DarkGray
Write-Host "von selbst, wenn kein _NEU_Vorlage neben der .exe liegt."                -ForegroundColor DarkGray
