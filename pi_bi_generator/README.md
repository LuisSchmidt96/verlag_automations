# PI/BI-Generator

Erzeugt aus einem **VLB-ONIX-XML**-Datensatz die **Presseinformation (PI)** und
**Buchinformation (BI)** – je als `.docx` und `.html` – in einem Ausgabeordner.
Der PDF-Export erfolgt weiterhin manuell aus Word.

## Bedienung

1. **VLB-ONIX-XML** auswählen (die aus dem VLB heruntergeladene `{ISBN}.xml`).
   Die erkannten Daten erscheinen sofort in der Vorschau.
2. **Coverbild** wählen (lokale `.jpg`/`.png`). Liegt neben der XML eine Datei wie
   `2D_300_{Kurzcode}.jpg`, wird sie automatisch vorgeschlagen. Alternativ
   *Vom Webserver* laden (`…/newsletter_/{Kurzcode}.png`).
3. Bei Bedarf **Ordnername** (Standard: Kurzcode, z. B. `05-559-2`) und
   **Produkt-Seite** (Webshop) anpassen.
4. **Dokumente erstellen** – es entstehen:
   `PI_{Kurzcode}.docx`, `BI_{Kurzcode}.docx`,
   `PI {Kurzcode}.html`, `BI {Kurzcode}.html`, `cover_{ISBN}.…`.

Der Werbetext wird **vollständig** aus der XML übernommen; die prüfende Person
kürzt ihn anschließend in Word, bis er auf eine Seite passt.

## Wichtig

- **Web-Assets manuell hochladen:** Die HTML-Fassung verweist auf
  `…/newsletter_/{Kurzcode}.png` (Cover-Thumbnail) und
  `…/presse/bib/bib_{Kurzcode}.pdf` (Blick ins Buch). Diese Dateien müssen
  separat auf den Webserver geladen werden.
- **Produkt-Seite:** Die Webshop-Produktseite steht nicht in der XML – Standard ist
  die Startseite; den echten Link ggf. im Feld eintragen.
- Die VLB-Cover-API (`api.vlb.de`) benötigt Zugangsdaten und wird **nicht**
  automatisch geladen; nutze eine lokale Datei oder *Vom Webserver*.

## Vorlagen (`vorlagen/`)

Die `*_vorlage.docx` sind aus den Muster-Dokumenten in `beispiele/` abgeleitet
(Briefkopf-Grafik, BI-Faxformular und Cover-Platzierung bleiben erhalten; jede
buchspezifische Zeile ist ein Platzhalter `{{…}}`). Neu erzeugen mit:

```bash
python -c "from pi_bi_generator import core; core.baue_docx_vorlagen()"
```

Die HTML-Vorlagen sind direkt in `vorlagen/` hinterlegt (UTF-8).

## Als Windows-.exe bauen

Aus dem Repo-Wurzelordner:

```cmd
pip install -r requirements.txt
pyinstaller pi_bi_generator/PiBiGenerator.spec
```

Ergebnis in `dist\PiBiGenerator\`. Der komplette Ordner wird auf den Ziel-PC
kopiert. Die Vorlagen sind in die `.exe` gebündelt; `config.json` wird beim
ersten Start direkt daneben angelegt (ebenso `pi_bi_output/` und `cover_cache/`).
