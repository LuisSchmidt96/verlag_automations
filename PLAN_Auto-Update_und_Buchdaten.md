# VR-Tools: Auto-Update der Programme + Buchdaten-Architektur

Plan-Dokument (Brainstorm-Ergebnis vom 29.07.2026). Umsetzung noch offen.

## Context

Zwei operative Schmerzpunkte im Verlag, beide brauchen Architektur statt Einzelfixes:

1. **Verteilung der Tools ist Handarbeit.** Heute baut `tools/update_and_build.ps1`
   die PyInstaller-onedir-Ordner und spiegelt sie nach `\\<Share>\VR-Tools\<Tool>\`.
   Die Kollegen kopieren den Ordner von Hand auf ihren PC. Jedes Update = Luis geht
   an jeden PC. Soll wegfallen.
2. **Buchdaten liegen chaotisch verstreut.** Master sollen die `Artikeldaten\
   <Kurzcode>_<Titel>\`-Ordner sein, aber Editieren auf dem Share ist langsam,
   also legen alle ad-hoc lokale Kopien an (z. B. `\\C024\D_Buch\…`), die
   auseinanderlaufen — keine klare Wahrheit.

**Entscheidungen (mit dem Nutzer abgestimmt):**
- Update-Weg: **Launcher (sync-then-run)** — lokale Kopie, beim Start vom Master
  aktualisiert.
- Buchdaten: **Check-out/Check-in-Tool** — Master = Wahrheit, lokal schnell
  arbeiten, sichtbare Sperre.
- Infrastruktur: es gibt einen **NAS als Share (Master)** und einen **Backup-NAS**.
  Der NAS ist damit der zentrale Master für BEIDES (Tools + Buchdaten).

**Zielbild:** NAS = einzige Quelle der Wahrheit. Client-PCs starten lokale Kopien,
die der Launcher frisch hält. Bücher werden zum Bearbeiten lokal ausgecheckt
(Sperre auf dem NAS) und zurückgespielt. Luis' Aufwand bei Updates: nur noch
`bauen + veröffentlichen`; kein PC-Besuch.

---

## Phase 0 — NAS-Pfade festnageln (Voraussetzung, klein)

Die Pfade C019/C018/C024 sind laut `WIP_TIF-Freistellung.md` unklar (C019 antwortet
nicht, C018 schon, C024 ist ein PC). Vor allem anderen den **echten NAS-Pfad**
bestätigen — einmal, per `dir \\<NAS>\...`:
- Master-Share für die Tool-Verteilung (heute `\\C019\d\VR-Tools`).
- Master-Ordner der Buchdaten (heute `\\C019\d\Online\Webseite\Artikeldaten`).

Dann die Vorgabe an **einer** Stelle je Belang korrigieren:
- `cover_previews/core.py` → `DEFAULT_CONFIG["artikeldaten_dir"]`
- `shopware_publisher/core.py` → dieselbe Vorgabe
- `tools/update_and_build.ps1` → `$ShareRoot` und die `_NEU_Vorlage`-Quelle in `$Beigaben`
- README + `Anleitung.txt`

Das ist der offene Punkt aus der WIP-Datei; hier wird er erledigt.

---

## Phase 1 — Launcher (Auto-Update der Tools) — GEBAUT, ungetestet

Umgesetzt am 13.08.2026 als `tools/launch.ps1`, `tools/Einrichten.cmd` +
`tools/einrichten.ps1`; `tools/update_and_build.ps1` veröffentlicht sie nach
`$ShareRoot`. **Auf einem Windows-PC noch nicht ausprobiert** — hier gibt es
kein PowerShell, die Skripte sind also nicht einmal syntaktisch geprüft. Die
Prüfliste unten unter „Verifikation / Phase 1“ abarbeiten.

Abweichung vom Plan (zum Guten): Phase 1 hängt **nicht** an Phase 0. `launch.ps1`
liegt auf dem Share neben den Tool-Ordnern und leitet den Master-Pfad aus
`$PSScriptRoot` ab; `Einrichten.cmd` aus `%~dp0`. Der Servername steht damit
nirgends im Code — nur im Ziel der Verknüpfungen, und das entsteht beim
Einrichten aus dem Ort, von dem aus die .cmd gestartet wurde. Zieht der Share
um, reicht ein erneutes `Einrichten.cmd` vom neuen Ort.

### Ursprünglicher Plan

**Prinzip:** Client-PCs führen eine LOKALE Kopie aus; ein Launcher-Skript auf dem
NAS synct beim Start die Programmteile delta-weise herunter und startet dann die
lokale .exe. Das heutige config-neben-der-.exe-Muster bleibt (die config ist
pro-Nutzer und lokal) — die Tools selbst ändern sich **nicht**.

**Neue Dateien auf dem NAS (unter `\\<NAS>\…\VR-Tools\`):**
- `launch.ps1` — nimmt einen Tool-Namen als Argument und macht:
  1. Quelle `…\VR-Tools\<Tool>\`, Ziel `%LOCALAPPDATA%\VR-Tools\<Tool>\`.
  2. **Nur Programmteile** synchronisieren, Nutzerdaten schonen — exakt die
     Trennung aus `Copy-Programmteile` in `tools/update_and_build.ps1`
     (Zeile 120–142) in Pull-Richtung wiederverwenden: `_internal\` per
     `robocopy /MIR`, `*.exe` + `Anleitung.txt` + Beigaben (`_NEU_Vorlage\`, auch
     `/MIR`) kopieren; `config.json` und andere lose Nutzerdateien **nie**
     anfassen (liegen top-level, werden von den `/MIR`-Unterordnern nicht berührt).
  3. NAS nicht erreichbar → Sync überspringen, vorhandene lokale Kopie starten
     (Offline-Resilienz). Keine lokale Kopie **und** kein NAS → klare Meldung.
  4. `%LOCALAPPDATA%\VR-Tools\<Tool>\<Tool>.exe` starten.
- `Einrichten.cmd` — **einmalig pro PC** doppelklicken: legt Desktop-/Startmenü-
  Verknüpfungen für alle Tools an. Jede Verknüpfung zeigt auf
  `powershell -NoProfile -ExecutionPolicy Bypass -File \\<NAS>\…\VR-Tools\launch.ps1 <Tool>`
  (`-ExecutionPolicy Bypass` umgeht Policy-Fragen ohne Maschinen-Policy zu ändern).

**Client-Fußabdruck = nur die Verknüpfungen** (einmal ausgerollt, danach nie wieder
angefasst). Launcher-Logik liegt zentral auf dem NAS → auch die kann Luis für alle
auf einmal ändern. Update-Fluss bleibt: `update_and_build.ps1` bauen +
veröffentlichen; alle bekommen die neue Version beim nächsten Start.

**Notizen:**
- `_NEU_Vorlage` (~460 MB): der erste Start auf einem neuen PC zieht es einmal,
  danach überträgt `/MIR` nur Geändertes (fast nichts).
- Alte, verstreute Kopien (`D:\CoverPreviews\…` o. Ä.) können nach der Umstellung
  gelöscht werden; kanonisch ist `%LOCALAPPDATA%\VR-Tools\<Tool>\`.
- Kein Code der Tools ändert sich in dieser Phase — rein additiv.

---

## Phase 2 — Check-out/Check-in-Tool für Buchdaten

Ein **fünftes Tool** im gleichen Muster (`core.py`/`app.py`/`main.py`/`.spec`,
config neben der .exe, verteilt über den Launcher aus Phase 1). Arbeitstitel
`buch_manager` / GUI „Buchdaten".

**Begriffe:**
- **Master** = NAS `…\Artikeldaten\<Kurzcode>_<Titel>\` (einzige Wahrheit).
- **Arbeitskopie** = lokaler Ordner (konfigurierbar, wie `artikeldaten_dir`), z. B.
  `D:\Buch_Arbeit\<Kurzcode>_<Titel>\`.
- **Sperre** = Marker-Datei im Master-Buchordner `.in_bearbeitung.json`
  `{user, host, seit}` — weiche, sichtbare Sperre (kein OS-Lock), überstimmbar.

**Operationen (GUI):**
1. **Auschecken** — Buch wählen (Liste vom NAS, Suche nach Kurzcode/Titel),
   `robocopy` Master→lokal, Sperre schreiben. Schon gesperrt → wer/seit anzeigen,
   „schreibgeschützt öffnen" oder „übernehmen" (steal) anbieten.
2. **Einchecken / Zurückspielen** — `robocopy` lokal→Master. **Konfliktschutz:**
   beim Auschecken einen Schnappschuss (Dateiliste + Größe + mtime) mitschreiben;
   beim Einchecken den Master dagegen prüfen. Hat sich der Master seither geändert
   (jemand anders hat geschrieben), warnen und den Master vorher nach
   `_alt/<Zeitstempel>/` sichern — das **`sichere_weg()`/`_alt`-Muster aus
   `cover_previews/core.py`** wiederverwenden. Danach Sperre lösen.
3. **Übersicht** — welche Bücher sind gerade ausgecheckt und von wem (Master nach
   `.in_bearbeitung.json` scannen), welche lokalen Arbeitskopien habe ich.
4. **Neues Buch anlegen** — `<Kurzcode>_<Titel>` anlegen (Namensregel wie
   `ordner_name()` in cover_previews).

**Wiederverwendung (Muster kopieren, NICHT importieren — „kein geteilter Code"):**
- Kurzcode-/Ordnerauflösung: `artikeldaten_dir()`, `finde_artikel_ordner()`,
  `ordner_name()`, `ziel_ordner()` aus `cover_previews/core.py`.
- Backup-vor-Überschreiben: `sichere_weg()` / `kollisionen()` ebenda.
- `robocopy` für die Transfers (schnelles Delta).
- GUI-Muster inkl. „Ablageort"-Feld (der lokale Arbeitsordner) analog zum
  Artikeldaten-Feld, das gerade in `cover_previews/app.py` gebaut wurde.

**Konfliktmodell:** weiche Sperren (Sichtbarkeit) + „Master seit Auschecken
geändert"-Erkennung beim Einchecken (fängt den gefährlichen Fall, kein
Zwei-Personen-Merge nötig, da Bücher praktisch immer von einer Person bearbeitet
werden). Stale Locks (vergessenes Einchecken) über die Übersicht + „übernehmen".

---

## Verifikation

**Phase 0:** `dir \\<NAS>\…` bestätigt beide Master-Pfade; Tools starten und lösen
den Zielordner korrekt auf dem NAS auf (Statuszeile „vorhandener Ordner" statt
„Share nicht erreichbar").

**Phase 1:**
- Auf einem Test-PC `Einrichten.cmd` doppelklicken → Verknüpfungen da.
- Verknüpfung starten → lokale Kopie unter `%LOCALAPPDATA%\VR-Tools\<Tool>\`
  entsteht, Tool startet, `config.json` bleibt lokal/pro-Nutzer.
- Eine neue Version bauen+veröffentlichen → erneuter Start zieht nur die geänderten
  Dateien (robocopy-Log), Version aktuell; laufende config.json unangetastet.
- NAS trennen → Start nutzt die lokale Kopie (Offline-Resilienz).

**Phase 2:**
- Auschecken eines Buchs → lokale Kopie + `.in_bearbeitung.json` auf dem NAS.
- Zweiter „Auschecken"-Versuch (anderer Nutzer/Host simuliert) → Sperre erkannt.
- Datei lokal ändern → Einchecken → Master aktualisiert, Sperre weg.
- Master zwischenzeitlich von Hand ändern, dann Einchecken → Konfliktwarnung +
  `_alt/<Zeitstempel>/`-Sicherung greift.

---

## Reihenfolge & Nutzen

- **Phase 0** zuerst (klein, entblockt alles, erledigt den offenen WIP-Punkt).
- **Phase 1** bringt sofort Nutzen und ist risikoarm (kein Tool-Code, rein additiv)
  → Handarbeit beim Verteilen fällt weg.
- **Phase 2** ist der größere Neubau (eigenes Tool) → löst das Buchdaten-Chaos,
  wird selbst über Phase 1 verteilt.

## Offen / später

**Zum Launcher (bewusst zurückgestellt am 13.08.2026 — erst mal Wichtigeres):**

- **Phase 1 auf Windows durchtesten.** Die Skripte sind hier nicht einmal
  syntaktisch geprüft (kein PowerShell auf dem Linux-Rechner). Prüfliste unter
  „Verifikation / Phase 1“.
- **`git push` landet nicht auf dem Share.** Bauen muss auf Windows passieren
  (PyInstaller kann nicht für Windows querbauen), und `update_and_build.ps1`
  startet bisher jemand von Hand. Kleinster Weg: eine Aufgabenplanung-Aufgabe
  auf einem PC, der ohnehin läuft, die genau dieses Skript aufruft — kein Code
  nötig. Dann aber: SSH-Key des Task-Benutzers ohne Passphrase (sonst hängt der
  `git pull` still), Bauen überspringen wenn nichts Neues kam, und vor allem
  **Fehler sichtbar machen** (Log + `_Stand.txt` auf dem Share mit Zeitstempel
  und Commit) — unbeaufsichtigt merkt sonst niemand, dass seit Wochen nichts
  mehr durchläuft. GitHub Actions bringt hier nichts: der Runner kommt nicht an
  den Share, es bräuchte trotzdem etwas on-prem.
- **Nutzdaten ziehen beim Umstieg nicht mit.** Kanonisch wird
  `%LOCALAPPDATA%\VR-Tools\<Tool>\`; wer heute `D:\CoverPreviews\` o. Ä. nutzt,
  lässt seine `config.json` dort zurück — und **BooxpressEtiketten** seine
  `kommliste.xlsx` und `paketnr.txt`, ohne die es nicht arbeitet. Beim Ausrollen
  je PC einmal übernehmen (von Hand oder `Einrichten.cmd` beibringen, aber dann
  muss es raten, wo die alten Ordner liegen).
- Verknüpfungen und `%LOCALAPPDATA%` sind **pro Windows-Benutzer**: bei mehreren
  Konten auf einem PC muss jeder einmal `Einrichten.cmd` klicken.

**Sonst:**

- Genauer NAS-Name/Pfad (Phase 0 klärt das).
- Ob die Tools-.exe signiert werden sollen (SmartScreen/AV-Ruhe) — separat, nicht
  Teil dieses Plans.
- Optional später: eine `version.txt` je Tool für eine sichtbare Versionsanzeige
  im Launcher/GUI (nicht nötig für die Funktion, robocopy-Delta reicht).
