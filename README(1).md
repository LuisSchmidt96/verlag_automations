# BOOXpress-Etiketten-Generator

Erzeugt BOOXpress-Versandetiketten aus einem Lexware-Aufträge-Export.

## Funktion

1. Mitarbeiterin exportiert in Lexware die Tagesaufträge als xlsx
2. Klick in der App → Datei auswählen → "Etiketten erstellen"
3. App erzeugt eine docx mit Etiketten für alle Aufträge, deren Kunden in der BOOXpress-Komm-Liste auftauchen
4. Übersprungene Aufträge (Endkunden, Kunden ohne BOOXpress-Match wie G. Umbreit) werden mit Begründung gelistet — das ist gleichzeitig die "über GLS versenden"-Liste
5. Paketnummer wird automatisch in `paketnr.txt` hochgezählt

## Dateien im App-Ordner (neben der .exe)

| Datei | Zweck |
|---|---|
| `BooxpressEtiketten.exe` | die App selbst |
| `config.json` | Konstanten (Verlag-K-Nr, Absender, Libri-Adresse) — wird beim ersten Start angelegt |
| `paketnr.txt` | aktueller Paketnr-Counter |
| `kommliste.xlsx` | aktuelle BOOXpress-Komm-Liste (muss manuell aktualisiert werden, wenn BOOXpress eine neue Liste schickt) |
| `etiketten_output/` | hier landen die erzeugten Etiketten-docx mit Zeitstempel |

## Build auf Windows (einmalig)

```cmd
:: Voraussetzung: Python 3.10+ installiert (https://python.org)
pip install pandas openpyxl python-docx python-barcode pillow pyinstaller

:: .exe bauen (von dem Ordner aus, in dem booxpress_app.py + booxpress_core.py liegen)
pyinstaller --windowed --noconfirm --name BooxpressEtiketten ^
    --collect-all barcode ^
    booxpress_app.py
```

`--collect-all barcode` ist wichtig — die `python-barcode`-Library hat eingebettete Font-Dateien und ein Spezial-CSS, das PyInstaller sonst nicht mitnimmt; ohne dieses Flag bricht der Barcode-Schritt zur Laufzeit ab.

Nach dem Build liegt die App unter `dist\BooxpressEtiketten\BooxpressEtiketten.exe`. Den gesamten `BooxpressEtiketten`-Ordner auf den Verlag-PC kopieren, dort `kommliste.xlsx` reinlegen, fertig.

## Einrichtung beim ersten Start

1. Ordner mit der .exe auf dem Verlag-PC platzieren (z.B. `C:\BOOXpress`)
2. Aktuelle BOOXpress-Komm-Liste als `kommliste.xlsx` in den Ordner kopieren
3. App starten → es wird automatisch `config.json` und `paketnr.txt` angelegt
4. Falls die Verlag-K-Nr bei BOOXpress nicht `068249` ist: `config.json` öffnen und anpassen
5. Falls die Paketnr nicht bei 1 starten soll: in der App auf "Zurücksetzen..." klicken und Startwert eingeben (oder direkt `paketnr.txt` bearbeiten)

## config.json — Felder

| Schlüssel | Default | Bedeutung |
|---|---|---|
| `verlag_knr_booxpress` | `"068249"` | K-Nr des Verlags bei BOOXpress (Konstante in jedem Barcode) |
| `kommliste_pfad` | `"kommliste.xlsx"` | Pfad zur Komm-Liste relativ zum App-Ordner |
| `absender_name` | `"verlag regionalkultur"` | erste Absenderzeile |
| `absender_adresse` | `"Bahnhofstr. 2, 76698 Ubstadt-Weiher"` | zweite Absenderzeile |
| `libri_pattern` | `"^Libri"` | Regex: wenn Name1 in der Komm-Liste matched, wird Libri-Sonderlogik aktiv |
| `libri_strasse` | `"Europaallee 1"` | Strasse für Libri-Etiketten |
| `libri_plz_ort` | `"36244 Bad Hersfeld"` | PLZ + Ort für Libri-Etiketten |
| `output_dir` | `"etiketten_output"` | Ausgabeordner |

## Matching-Logik

- Endkunden (Lexware Kd.-Nr. ohne führende `1`) werden übersprungen (Vermerk in Log: "Endkunde — kein BOOXpress")
- Geschäftskunden: führende `1` wird gestrippt, der Rest wird gegen die `VD`-Spalte der Komm-Liste gematched
  - Match → Etikett wird erzeugt
  - Kein Match → übersprungen (Vermerk: "VD xxx nicht in Komm-Liste — vmtl. GLS")
- Libri-Spezialfall: wenn `Name1` in der Komm-Liste mit "Libri" anfängt (Regex), wird das Etikett als `Libri_<VD>` mit der Bad-Hersfeld-Adresse beschriftet, statt mit Name1 und der Komm-Listen-Adresse

## Barcode

- Format: ITF (Code 2 of 5 Interleaved)
- 18 Stellen: `<VerlagKNr 6-stellig><Paketnr 6-stellig><VD 6-stellig>`
- Beispiel: `068249002037025700` = Verlag 068249 + Paketnr 2037 + Libri FA 25700

## Bekannte Stolperstellen

- **Komm-Liste muss aktuell sein.** Wenn BOOXpress eine neue Komm-Liste schickt, muss diese die alte `kommliste.xlsx` im App-Ordner ersetzen. Sonst werden neue/umgezogene Buchhandlungen falsch behandelt.
- **Lexware-Export-Format**: das Skript erwartet Zeile 1 = Titel ("Aufträge Verkauf ..."), Zeile 2 = Header. Wenn Lexware sein Export-Format ändert (z.B. ab Lexware 2027), könnte das brechen — dann muss in `booxpress_core.py` die Zeile `header=1` in `lade_auftraege` angepasst werden.
- **Paketnummer-Counter persistiert in `paketnr.txt`**. Bei Backup/Restore oder PC-Wechsel diese Datei mitnehmen, sonst beginnt der Counter wieder bei 1.

## Wenn etwas nicht funktioniert

- Im Log-Bereich der App steht eine vollständige Fehlermeldung — Screenshot davon machen und mir schicken
- Wenn die App gar nicht startet: aus der Konsole (cmd, in den App-Ordner cd'en und `BooxpressEtiketten.exe` ausführen) sieht man evtl. mehr Fehler-Output
