# WIP — freigestelltes CMYK-TIF & Artikeldaten-Share

Arbeitsstand vom **28.07.2026**, angefangen auf dem Linux-Rechner. Zwei Aufträge,
beide noch nicht fertig. Die Befunde unten sind **gemessen**, nicht vermutet —
sie sparen die halbe Arbeit, wenn hier woanders weitergemacht wird.

---

## 1. Freigestelltes CMYK-TIF zum 3D-Mockup

### Anforderung (O-Ton)

> „to also create a .tif of the 3D_300dpi version where only the book without the
> shadow and background is in an alpha ebene or sth (and make sure there's no
> white lines on the edges, should be only book) and also in CMYK colorspace“

Also zusätzlich zu `{sc}.png`, `3D_300_{sc}.jpg`, `3D_72_{sc}.jpg` eine Datei
`3D_300_{sc}.tif` mit:

* **nur dem Buch** — ohne Hintergrund *und* ohne Spiegelung,
* **Transparenz + Alphakanal** (beides, damit es sowohl InDesign als auch
  Programme bedienen, die nur einen Alphakanal auswerten),
* **CMYK**,
* **keinem weißen Strich an den Kanten**.

### Befunde (gemessen, nicht geraten)

**Aufbau der Vorlagen.** Jede Mockup-PSD enthält das Buch **zweimal**: einmal
aufrecht, einmal als Spiegelung darunter — je als Gruppe aus Cover-, Rücken- und
Falz-Ebene. Dazu der Hintergrund (steht schon als `hide_bg` in
`vorlagen_map.json`). Nur `17x24` benennt die Gruppen ehrlich (`Buch` /
`Spiegelung`), alle anderen heißen `Gruppe 1` — **über den Namen ist das also
nicht zu erkennen**, wohl aber über die Lage auf der Leinwand (die Spiegelung
liegt tiefer).

Ausgeblendet werden muss je Vorlage diese Gruppe (mit `psd-tools` aus den
lokalen PSDs ausgelesen):

| Vorlage | Spiegelung | Vorlage | Spiegelung |
|---|---|---|---|
| `12x19.psd`   | 42 | `21x21.psd`  | 43 (schon unsichtbar) |
| `12x22.psd`   | 42 | `22x24.psd`  | 42 |
| `15,5x22.psd` | 42 | `28x21.psd`  | 42 |
| `16x16.psd`   | 43 | `29,7x21.psd`| 33 |
| `17x24.psd`   | 99 | `29x22.psd`  | 33 |
| `21x13,5.psd` | 45 | `EBOOK.psd`  | — (keine Spiegelung) |
| `21x14.psd`   | 47 | `MaGeBl.psd` | — (keine Smart-Objekte) |

**Woher der weiße Strich kommt.** Die **Falz-Ebene ragt oben und unten über die
Buchkante hinaus** und trägt dort einen praktisch unsichtbaren weißen Schein:
Alpha **1/255**, RGB **255,255,255**. Nachgewiesen im echten Photoshop-Ergebnis
`\\c024\d_buch\CoverPreviews\cover_output\05-627-8.png`:

```
Zeile 177–180: nur Pixel mit Alpha 1, RGB 255/255/255, x 283–290 (= Falz-Bereich)
Zeile 181:     das Buch fängt an (Alpha bis 241)
```

`doc.trim(TrimType.TRANSPARENT)` hängt sich an diese Pixel — der Zuschnitt liegt
also 4–11 px neben dem Buch, und beim Platzieren bzw. beim Verrechnen auf einen
Hintergrund wird daraus der **weiße Strich**. In `17x24` sind es 5 px oben und
11 px unten.

→ **Fix:** alles unter Alpha ~8 auf 0 setzen, **danach** eng zuschneiden.

**Kein Defringe nötig.** Die Buchkanten selbst sind **hart** (Alpha entweder 0
oder 255, gemessen an Ober-, Unter-, Linkskante und rechter Kante). In den
Halbtonpixeln steckt also *kein* eingerechnetes Weiß — „Remove White Matte“ /
Defringe würde die Farben nur unnötig verfälschen. Im ganzen Buch-Composite von
`17x24` gibt es 2002 halbdurchsichtige Pixel, davon 373 mit Alpha < 8; alle
gehören zum Falz-Schein.

### Geplanter Ablauf

Warum zwei Photoshop-Durchgänge mit Python dazwischen: **CMYK muss über
Photoshop** laufen (Verlags-ICC-Profil), aber **PIL kann kein CMYK + Alpha**
schreiben (Modus `CMYK` hat genau 4 Kanäle, `CMYKA` gibt es nicht). Die
Alpha-Säuberung wiederum ist in Python drei Zeilen und prüfbar, in ExtendScript
dagegen Gefummel (Levels auf einen Alphakanal). Also:

1. **JSX 1** (Erweiterung des bestehenden `_baue_jsx`), nach den JPEGs:
   Spiegelung ausblenden → `doc.trim(TrimType.TRANSPARENT)` → „Auf eine Ebene
   reduziert kopieren“ in ein frisches transparentes Dokument → als
   `_buch_{sc}.png` exportieren (Save-for-Web, PNG-24, Transparenz).
   Der Hintergrund ist zu diesem Zeitpunkt schon aus (Schritt 3).
2. **Python** (`core.freistellen()`): Alpha < `tif_alpha_min` (8) → 0, dann auf
   die gesäuberte Bounding-Box zuschneiden, wieder als PNG sichern.
3. **JSX 2** (neu, `_baue_tif_jsx`): PNG öffnen → `resizeImage(…, 300,
   ResampleMethod.NONE)` (nur DPI-Eintrag, wie bei den 3D-JPEGs) →
   `changeMode(ChangeMode.CMYK)` → Transparenz als Auswahl laden und zusätzlich
   als Alphakanal ablegen → `saveAs` TIFF (LZW, `alphaChannels=true`,
   `transparency=true`, `layers=false`, Profil einbetten).

`ps.DoJavaScript()` ist synchron, die Reihenfolge trägt also.

### Was schon im Code steht

* `core.py` → `DEFAULT_CONFIG`: `muster_3d_tif`, `tif_erzeugen`, `tif_alpha_min`.

### Was noch fehlt

* [ ] `vorlagen_map.json`: Schlüssel `"spiegelung": [id]` je Vorlage (Tabelle oben).
* [ ] `psd_analyse.py`: die IDs selbst herleiten, damit es eine PSD-Revision
      übersteht. Regel: alle Slots + Falz-Ebenen nach y-Mitte in „oben = Buch“ und
      „unten = Spiegelung“ teilen; je Spiegelungs-Ebene die **oberste Gruppe
      suchen, die nicht auch das Buch enthält** — das ist die Gruppe, die
      ausgeblendet wird. Dafür muss das Analyse-JSX zusätzlich den Ebenenbaum
      liefern (`id | parent-id | ist-Gruppe | name | top | bottom`); die
      Falz-Erkennung kann dann aus dem Baum kommen statt aus dem zweiten Walk.
* [ ] `core.py`: `freistellen()`, `_baue_tif_jsx()`, `_baue_jsx()` um den
      Buch-Durchgang erweitern, zweiter `DoJavaScript`-Aufruf in
      `erzeuge_3d_photoshop()`, TIF in die Fehlend-Prüfung und in
      `ausgabe_namen()`, `raeume_auf()` um `_buch_{sc}*.png` / `_tif_{sc}.jsx`
      ergänzen.
* [ ] `app.py`: Ankreuzfeld „3D zusätzlich als CMYK-TIF (Buch freigestellt)“,
      gekoppelt an die 3D-Ausgabe, Stand in `config.json`.
* [ ] README + `Anleitung.txt`.
* [ ] **Auf einem Windows-Rechner mit Photoshop gegenprüfen** — der ganze
      3D-Zweig läuft hier nur als Dry-Run.

### Zu klären

* Der Anwender hat eine **Referenz-TIFF unter `smb://c018/d`** erwähnt („you can
  check at smb://c018/d, there's a tiff“) — **noch nicht angesehen**. Vorher
  anschauen: Kanäle, Profil, ob flach oder mit Ebene, wie der Alphakanal heißt.
  Danach entscheiden, ob Transparenz *und* Alphakanal sinnvoll sind (6 Kanäle
  können manche RIPs stören) oder nur eines von beidem.
* Der bisherige `3D_300_05-627-8.tif` im Ausgabeordner ist eine **von Hand
  gespeicherte RGB-Kopie des JPEGs** (gleiche Pixel, kein Alpha) — der wird vom
  neuen TIF abgelöst, gleicher Dateiname ist also richtig.

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
