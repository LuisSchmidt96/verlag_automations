# WIP — freigestelltes CMYK-TIF & Artikeldaten-Share

Arbeitsstand vom **28.07.2026**, angefangen auf dem Linux-Rechner. Zwei Aufträge,
beide noch nicht fertig. Die Befunde unten sind **gemessen**, nicht vermutet —
sie sparen die halbe Arbeit, wenn hier woanders weitergemacht wird.

---

## 1. Freigestelltes CMYK-TIF zum 3D-Mockup  —  ERLEDIGT

Umgesetzt am 29.07.2026, gegen echtes Photoshop (27.3.1) verifiziert. Ergebnis-TIF
wie die Referenz `3D_300_05-627-8.tif`: CMYK, flach, weisser Grund, Buch-only,
Vektor-Beschneidungspfad „Pfad 1“, ICC eingebettet, 300 dpi, LZW, KEIN Alpha; kein
weisser Strich an den Kanten (gegen Magenta geprueft). Details siehe README
(Abschnitt „Freigestelltes CMYK-TIF“).

Beruehrte Stellen: `psd_analyse.py` (Ebenenbaum -> `spiegelung` je Vorlage),
`vorlagen_map.json` (`spiegelung`), `core.py` (`_tif_block`, `_baue_jsx`,
`erzeuge_3d_photoshop`, `ausgabe_namen`), `app.py` (Ankreuzfeld), README + Anleitung.
Die Referenz-TIFF `3D_300_05-627-8.tif` liegt weiter im Ordner (Vergleichsbild), ist
per .gitignore aussen vor.

---

## 2. „Findet den Share-Ausgabeordner nicht“

### Befund

`DEFAULT_CONFIG["artikeldaten_dir"]` (und dieselbe Voreinstellung in
`shopware_publisher/core.py`) steht auf

```
\\C019\d\Online\Webseite\Artikeldaten
```

**`C019` antwortet nicht.** Der Name löst zwar auf (`c019.internal` →
`192.168.2.145`), aber `ping` bleibt still und `smbclient -L //c019.internal`
sagt `NT_STATUS_HOST_UNREACHABLE`. `C024` dagegen antwortet
(`NT_STATUS_ACCESS_DENIED` = Anmeldung fehlt, Rechner ist da).

`core.artikeldaten_dir()` gibt in dem Fall `None` zurück und `ziel_ordner()`
fällt still auf `cover_output/` neben der .exe zurück — genau das ist am
27.07. passiert: die Ergebnisse liegen in
`\\c024\d_buch\CoverPreviews\cover_output\` statt im Artikelordner.

Die restlichen Pfade in der ausgelieferten `config.json`
(`\\c024\d_buch\CoverPreviews\config.json`) zeigen alle auf **C024**:
`last_input_dir = \\C024\D_Buch\Buch\…`, `vorlagen_dir = D:\CoverPreviews\_NEU_Vorlage`.
Auf `\\C024\D_Buch` gibt es aber **kein** `Online\Webseite\Artikeldaten`.

**Der Anwender hat inzwischen `smb://c018/d` genannt** — sehr wahrscheinlich ist
der richtige Server also **C018**, nicht C019 (Zahlendreher). Das ist noch **nicht
bestätigt**; zu prüfen ist, ob dort `Online\Webseite\Artikeldaten` liegt.

Auf dem Windows-Rechner prüfen:

```powershell
hostname          # bzw. echo %COMPUTERNAME%
net use           # worauf zeigen die verbundenen Laufwerke
net view \\C018   # welche Freigaben hat der Rechner
dir \\C018\d\Online\Webseite\Artikeldaten
```

### Was zu tun ist

* [ ] Richtigen Pfad bestätigen und die Voreinstellung in **beiden** Tools
      korrigieren (`cover_previews/core.py`, `shopware_publisher/core.py`),
      dazu README und `Anleitung.txt`. `tools/update_and_build.ps1` zieht die
      Vorlagen ebenfalls von `\\C019\d\…\_NEU_Vorlage` und spiegelt nach
      `\\C019\d\VR-Tools` — beides mitziehen.
* [ ] **Der eigentliche Bedienfehler:** der Ordner ist in der GUI gar nicht
      einstellbar. Steht der Pfad falsch, sieht der Anwender nur
      „⚠ Share nicht erreichbar — lokaler Ordner“ und kann nichts dagegen tun,
      ohne die `config.json` von Hand zu ändern. Also im Rahmen „Zielordner
      (Artikeldaten)“ ein Eingabefeld + „…“-Knopf für `artikeldaten_dir`
      ergänzen, und in der Statuszeile den **konfigurierten Pfad mit ausgeben**,
      damit man sieht, *worauf* er zeigt.
