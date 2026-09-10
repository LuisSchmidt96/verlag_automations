<#
    launch.ps1
    ----------
    Startet ein VR-Tool und bringt es dabei auf den neuesten Stand.

    Aufgerufen wird das Skript von den Verknuepfungen, die Einrichten.cmd auf
    dem Client anlegt:

        powershell -NoProfile -ExecutionPolicy Bypass -File \\<NAS>\VR-Tools\launch.ps1 CoverPreviews

    Ablauf:
      0. NUR auf einem Baurechner (Marke $BauMarke): update_and_build.ps1
         laufen lassen - das holt den neuesten Stand aus dem Git-Repo und baut
         die Tools neu, an denen sich etwas geaendert hat.
      1. Programmteile vom NAS in die lokale Kopie spiegeln (nur Geaendertes).
      2. Lokale .exe starten.

    Schritt 0 ist bewusst NICHT fuer alle: Bauen braucht git, Python 3.12 und
    PyInstaller, dauert Minuten und schreibt in den Master-Ordner. Wuerden das
    mehrere Rechner gleichzeitig tun, kaemen sich PyInstaller (loescht bei
    --noconfirm den Ziel-Ordner) und robocopy /MIR ins Gehege. Die Kollegen
    holen sich das Ergebnis wie bisher in Schritt 1 ab.

    Warum ueberhaupt lokal und nicht direkt vom Share starten:
      * Ein onedir-Build ueber SMB startet langsam.
      * Die Nutzdaten (config.json, kommliste.xlsx, paketnr.txt, Ausgabeordner)
        liegen neben der .exe und gehoeren dem jeweiligen Anwender - auf dem
        Share haetten alle dieselben.
      * Der Build wuerde Dateien austauschen, waehrend jemand sie offen hat.

    Der Master-Pfad steht NICHT im Skript: launch.ps1 liegt auf dem NAS neben
    den Tool-Ordnern, $PSScriptRoot ist damit das VR-Tools-Verzeichnis. Zieht
    der Share um, aendert sich nur das Ziel der Verknuepfungen (einmal
    Einrichten.cmd vom neuen Ort ausfuehren).
#>

param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Tool
)

$ErrorActionPreference = 'Stop'

$Master = Join-Path $PSScriptRoot $Tool
$Lokal  = Join-Path $env:LOCALAPPDATA "VR-Tools\$Tool"
$Exe    = Join-Path $Lokal "$Tool.exe"

# Was zum PROGRAMM gehoert und darum gespiegelt wird. Alles andere im
# Tool-Ordner gehoert dem Anwender und wird nie angefasst. Muss zu
# Copy-Programmteile in tools/update_and_build.ps1 passen - dort wird
# dieselbe Trennung beim Veroeffentlichen gezogen, hier beim Abholen.
$ProgrammOrdner  = @('_internal', '_NEU_Vorlage')
$ProgrammDateien = @('*.exe', 'Anleitung.txt')

function Show-Fehler([string]$Text) {
    Write-Host ""
    Write-Host $Text -ForegroundColor Red
    # Zusaetzlich als Fenster: die Konsole ist beim Doppelklick schnell wieder
    # weg, die Meldung soll aber ankommen.
    try {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show(
            $Text, 'VR-Tools', 'OK', 'Error') | Out-Null
    } catch { }
}

# --- 0) Auf einem Baurechner: pullen und neu bauen -------------------------
# Die Marke ist eine leere Datei und liegt LOKAL, nicht im Master-Ordner: so
# entscheidet jeder Rechner fuer sich, und ein erneutes Einrichten.cmd aendert
# nichts daran.
#
#   Anlegen:   New-Item "$env:LOCALAPPDATA\VR-Tools\bauen" -ItemType File
#   Entfernen: Remove-Item "$env:LOCALAPPDATA\VR-Tools\bauen"
#
# WAS gebaut werden muss, entscheidet update_and_build.ps1 selbst: es fuehrt je
# Tool einen Baustempel (die Git-SHA des letzten Baus) und ueberspringt alles
# Unveraenderte. Hier wird deshalb nichts nachgerechnet - einmal aufrufen
# genuegt, auch wenn sich nichts geaendert hat.
#
# Der Pfad wird aus dem eigenen Ort abgeleitet, wie $Master weiter oben: das
# Repo liegt neben den Tool-Ordnern im Master-Ordner. So bleibt auch hier kein
# Servername fest verdrahtet.
$BauMarke  = Join-Path $env:LOCALAPPDATA 'VR-Tools\bauen'
$BauSkript = Join-Path $PSScriptRoot 'repo\tools\update_and_build.ps1'

if (Test-Path -LiteralPath $BauMarke) {
    if (Test-Path -LiteralPath $BauSkript) {
        Write-Host "Baurechner - hole Aenderungen und baue neu ..." -ForegroundColor Cyan
        $Hier = Get-Location
        try {
            & $BauSkript
        } catch {
            # Ein fehlgeschlagener Bau darf die Arbeit NICHT aufhalten: dann
            # startet eben die vorhandene Fassung. Haeufigster Fall: das Tool
            # laeuft schon und seine .exe ist gesperrt.
            Write-Warning "Bauen fehlgeschlagen - starte die vorhandene Fassung."
            Write-Warning $_.Exception.Message
        } finally {
            # update_and_build.ps1 wechselt das Arbeitsverzeichnis.
            Set-Location $Hier
        }
    } else {
        Write-Warning "Baumarke liegt, aber $BauSkript ist nicht erreichbar."
        Write-Warning "Es wird nur gespiegelt und gestartet."
    }
}

try {
    $MasterDa = Test-Path -LiteralPath $Master

    if ($MasterDa) {
        if (-not (Test-Path -LiteralPath $Lokal)) {
            Write-Host "Erste Einrichtung von $Tool - das kann etwas dauern." -ForegroundColor Yellow
        }
        Write-Host "Aktualisiere $Tool ..." -ForegroundColor Cyan
        New-Item -ItemType Directory -Force -Path $Lokal | Out-Null

        foreach ($Ordner in $ProgrammOrdner) {
            $Quelle = Join-Path $Master $Ordner
            if (-not (Test-Path -LiteralPath $Quelle)) { continue }
            # /MIR haelt den Ordner deckungsgleich und uebertraegt nur
            # Geaendertes; beim CoverPreviews-Ordner _NEU_Vorlage (~460 MB)
            # kostet nur der erste Start Zeit.
            robocopy $Quelle (Join-Path $Lokal $Ordner) /MIR /NFL /NDL /NJH /NJS /NP | Out-Null
            if ($LASTEXITCODE -ge 8) {   # 0-7 sind bei robocopy Erfolg
                throw "Konnte $Ordner nicht aktualisieren (robocopy $LASTEXITCODE)."
            }
        }

        foreach ($Muster in $ProgrammDateien) {
            Get-ChildItem -LiteralPath $Master -Filter $Muster -File -ErrorAction SilentlyContinue |
                ForEach-Object {
                    # Datei festhalten: in catch ist $_ der Fehler, nicht mehr
                    # das Element aus der Pipeline.
                    $Datei = $_
                    try {
                        Copy-Item -LiteralPath $Datei.FullName -Destination $Lokal -Force
                    } catch {
                        # Laeuft das Tool gerade, ist die .exe gesperrt. Kein
                        # Grund abzubrechen: die vorhandene Fassung startet,
                        # die neue kommt beim naechsten Mal.
                        Write-Warning "$($Datei.Name) ist in Benutzung - bleibt vorerst wie es ist."
                    }
                }
        }
    } else {
        Write-Warning "NAS nicht erreichbar ($Master) - starte die vorhandene Fassung."
    }

    if (-not (Test-Path -LiteralPath $Exe)) {
        if ($MasterDa) {
            throw "Im Ordner $Master liegt keine $Tool.exe. Bitte Luis Bescheid geben."
        }
        throw ("$Tool ist auf diesem PC noch nicht eingerichtet, und der Ordner " +
               "$Master ist nicht erreichbar.`n`nBitte mit dem Netzlaufwerk " +
               "verbinden und es noch einmal versuchen.")
    }

    Start-Process -FilePath $Exe -WorkingDirectory $Lokal
} catch {
    Show-Fehler $_.Exception.Message
    exit 1
}
