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

Der Einband kommt aus ONIX `b012` über `einband_map`: `BB` → „fester Einband“,
`BC`/`PB` → „**Broschur**“ (nicht „kartoniert“ — das Wort benutzt der Verlag
nicht), `BE` → „Klappenbroschur“, `BZ` → „Leinen“.

## Was wo steht

Die Dokumente haben zwei Stellen, an denen dieselben Angaben auftauchen — den
**Kopf** über dem Werbetext und den **Kasten** darunter. Sie sind bewusst
verschieden gesetzt:

| Zeile | Kopf | Kasten |
|-------|------|--------|
| Herausgeber | `A, B und C (Hrsg.)` | `A, B, C (Hrsg.)` |
| Autoren | `A, B und C` | `A, B und C` |
| Titel | nur der Titel | `Titel. Untertitel.` |
| darunter | Untertitel, sonst `Band 5` | Reihenzeile, `Mit Beiträgen von …`, `Herausgegeben vom … Fotografien von …` |

### Autor oder Herausgeber — wer steht oben?

Gebaut war das Werkzeug an einem Sammelband (Bausteine Bruchsal Bd. 5):
Herausgeber oben, **jeder** Autor (`A01`) als „Mit Beiträgen von“. Von 16 PIs
auf `vr-intern/VLB` sind aber 13 Autorenbücher. Beim PZN-Führer (05-194: Autor
`A01` + Körperschaft `B01`) wurde daraus „Psychiatrisches Zentrum Nordbaden
(PZN) (Hrsg.)“ und „Mit Beiträgen von Hansi Rau.“ Von Hand gesetzt war „Hansi
Rau“ oben und „Herausgegeben vom Psychiatrischen Zentrum …“ im Kasten. Jetzt:

| Im VLB | oben | im Kasten darunter |
|--------|------|--------------------|
| Personen als Herausgeber (`B01`) | Herausgeber mit „(Hrsg.)“ | Autoren als „Mit Beiträgen von …“ |
| Autoren (`A01`), keine Personen-Hrsg. | Autoren, ohne „(Hrsg.)“ | Körperschaft in der Reihenzeile oder als „Herausgegeben vom …“ |
| nur Körperschaft(en) | Körperschaft mit „(Hrsg.)“ | — |

**Weitere Beteiligte** (Vorwort, Einführung, Einleitung, Texte, Fotografien,
Illustrationen, Zeichnungen, Karten, Übersetzung, Nachwort; `A32` „Beiträge
von“ geht zu „Mit Beiträgen von“) stehen in einer Zeile mit „Herausgegeben
vom …“, wie in PI 05-217: „Herausgegeben vom Nationalpark Schwarzwald.
Fotografien von Burkhard Riegels.“ Rollen und Wortlaut stehen in
`NEBENROLLEN`; eine Rolle, die dort fehlt, wird nicht gesetzt, aber in der
Ausgabe gemeldet.

**Format** hinter dem Einband: höchstens 12,5 cm breit → „im handlichen
Taschenformat“ (nur die Breite zählt), sonst breiter als 17 **oder** höher als
24 cm → „im repräsentativen Großformat“. Dazwischen kein Zusatz. Das
„attraktive quadratische Format“ der alten PIs gibt es nicht mehr; ein
21 × 21 ist Großformat.

Dazu die Regeln, die man sonst wieder herausnimmt:

* **Körperschaften sind Herausgeber.** ONIX `b047` („Deutsche
  Waldenservereinigung e.V. Ötisheim-Schönenberg“) wurde früher gar nicht
  gelesen — ein solcher Herausgeber fiel ersatzlos aus PI und BI heraus.
* **Die Reihe steht nur im Kasten, wenn sie nicht der Buchtitel selbst ist.**
  Bei den „Bausteinen“ heisst die Reihe wie das Buch; dort stünde sonst zweimal
  dasselbe. Ohne Reihenzeile bleibt die Bandangabe in der Titelzeile
  (`… und ihres Umlands. Band 5.`), mit Reihenzeile wandert sie dorthin
  (`Waldenserstudien, hrsg. von …, Bd. 9`).
* **Steht oben schon jemand, kommt die herausgebende Körperschaft in die
  Reihenzeile** — es sei denn, es gibt keine Reihe; dann steht sie bei den
  Herausgebern (in VLB-Reihenfolge: „Landkreis Rastatt, Toni Huber (Hrsg.)“)
  bzw. bei einem Autorenbuch als „Herausgegeben vom …“.
* **`koerperschaft_dativ` in der `config.json`** hält die gebeugte Form für
  „hrsg. von …“: `"Deutsche Waldenservereinigung e.V. …"` →
  `"der Deutschen Waldenservereinigung e.V. …"`. Ob es *der*, *dem* oder *den*
  heisst, hängt am Geschlecht des Namens und steht in keinen Daten — geraten
  wird es deshalb nicht. „von dem“ wird zu „vom“. Fehlt ein Name in der
  Tabelle, erscheint die ungebeugte Form und das Werkzeug sagt es in der
  Ausgabe. Dieselbe Tabelle steht im `shopware_publisher` — ein Eintrag
  gehört in beide.

**Was das Werkzeug nicht wissen kann:** Es hält sich an die Rollen im VLB.
Wer dort als Herausgeber steht, aber auf dem Umschlag als Autor erscheint (oder
umgekehrt), steht im PI falsch — korrigiert werden muss es im VLB oder von
Hand in Word. Ob ein Band mit Personen-Herausgebern *und* einem einzigen
Autor ein Sammelband ist oder eine edierte Monografie, steht ebenfalls nicht
in den Daten; das Werkzeug behandelt ihn als Sammelband.

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
