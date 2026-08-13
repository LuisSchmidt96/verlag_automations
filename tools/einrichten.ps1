<#
    einrichten.ps1
    --------------
    Legt auf DIESEM PC die Verknuepfungen zu den VR-Tools an - einmal pro PC.
    Aufgerufen wird es ueber Einrichten.cmd (Doppelklick), das die
    PowerShell-Ausfuehrungsrichtlinie umgeht.

    Angelegt wird je Tool eine Verknuepfung auf dem Desktop und im Startmenue.
    Sie ruft nicht die .exe auf, sondern launch.ps1 auf dem NAS - das haelt das
    Tool bei jedem Start aktuell.

    Auf dem Client bleibt damit NUR die Verknuepfung liegen. Die ganze Logik
    steht auf dem NAS und laesst sich fuer alle auf einmal aendern.
#>

$ErrorActionPreference = 'Stop'

# Das Skript liegt auf dem NAS neben den Tool-Ordnern - von dort kommt der
# Master-Pfad, der in die Verknuepfungen geschrieben wird.
$Nas    = $PSScriptRoot
$Launch = Join-Path $Nas 'launch.ps1'

if (-not (Test-Path -LiteralPath $Launch)) {
    throw "launch.ps1 liegt nicht neben dieser Datei ($Nas) - bitte Einrichten.cmd direkt vom Netzlaufwerk starten."
}

# Ein Tool ist ein Unterordner, in dem eine gleichnamige .exe liegt
# (VR-Tools\CoverPreviews\CoverPreviews.exe) - genau das Layout, das
# update_and_build.ps1 veroeffentlicht.
$Tools = Get-ChildItem -LiteralPath $Nas -Directory |
         Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName "$($_.Name).exe") }

if (-not $Tools) {
    throw "In $Nas liegen keine Tool-Ordner. Wurde schon veroeffentlicht?"
}

$Desktop   = [Environment]::GetFolderPath('Desktop')
$Startmenu = Join-Path ([Environment]::GetFolderPath('Programs')) 'VR-Tools'
New-Item -ItemType Directory -Force -Path $Startmenu | Out-Null

$Shell = New-Object -ComObject WScript.Shell

foreach ($T in $Tools) {
    $Name = $T.Name
    foreach ($Ort in @($Desktop, $Startmenu)) {
        $Lnk = $Shell.CreateShortcut((Join-Path $Ort "$Name.lnk"))
        $Lnk.TargetPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
        $Lnk.Arguments  = "-NoProfile -ExecutionPolicy Bypass -File `"$Launch`" $Name"
        # Symbol von der .exe auf dem NAS, damit die Verknuepfung aussieht wie
        # das Programm (Windows merkt sich das Bild).
        $Lnk.IconLocation    = (Join-Path $T.FullName "$Name.exe") + ',0'
        $Lnk.WorkingDirectory = $T.FullName
        $Lnk.Description      = "$Name - wird beim Start automatisch aktualisiert"
        $Lnk.Save()
    }
    Write-Host "  $Name" -ForegroundColor Green
}

Write-Host ""
Write-Host "Fertig. Die Verknuepfungen liegen auf dem Desktop und im Startmenue" -ForegroundColor Cyan
Write-Host "unter 'VR-Tools'. Beim ersten Start holt sich jedes Tool seine" -ForegroundColor Cyan
Write-Host "Dateien vom Netzlaufwerk - das dauert einmalig etwas." -ForegroundColor Cyan
