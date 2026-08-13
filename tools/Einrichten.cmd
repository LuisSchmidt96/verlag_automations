@echo off
rem  Einrichten.cmd
rem  -------------------------------------------------------------------
rem  Einmal pro PC doppelklicken (direkt hier vom Netzlaufwerk aus).
rem  Legt die Verknuepfungen zu den VR-Tools auf Desktop und Startmenue an.
rem
rem  Der Umweg ueber die .cmd ist Absicht: eine .ps1 laesst sich nicht
rem  doppelklicken (bzw. oeffnet den Editor), und -ExecutionPolicy Bypass
rem  umgeht die Richtlinie fuer diesen einen Aufruf, ohne am PC etwas zu
rem  verstellen.
rem  -------------------------------------------------------------------

echo.
echo VR-Tools einrichten
echo ===================
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0einrichten.ps1"

echo.
pause
