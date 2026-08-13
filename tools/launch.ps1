<#
    launch.ps1
    ----------
    Startet ein VR-Tool und bringt es dabei auf den neuesten Stand.

    Aufgerufen wird das Skript von den Verknuepfungen, die Einrichten.cmd auf
    dem Client anlegt:

        powershell -NoProfile -ExecutionPolicy Bypass -File \\<NAS>\VR-Tools\launch.ps1 CoverPreviews

    Ablauf:
      1. Programmteile vom NAS in die lokale Kopie spiegeln (nur Geaendertes).
      2. Lokale .exe starten.

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
