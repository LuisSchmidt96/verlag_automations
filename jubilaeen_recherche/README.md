# jubilaeen_recherche

Recherche-Skript (Kommandozeile, keine .exe): holt Unternehmen, Vereine,
Gemeinden und Ortsteile in BW, RLP, Saarland und Hessen sowie in den an BW
grenzenden bayerischen Kreisen (`ZUSATZ_KREISE` im Skript) mit Gründungs- bzw.
Ersterwähnungsdatum aus Wikidata und rechnet die runden Jubiläen aus.

```
python jubilaeen.py --kreis Q8178 --jahr 2029   # Probelauf Landkreis Karlsruhe
python jubilaeen.py --jahr 2029                 # alle vier Länder
python jubilaeen.py --nur-auswerten --jahr 2030 # nur die vorhandene DB lesen
```

Ergebnis neben dem Skript: `jubilaeen*.sqlite`, `*_items.csv`, `*_daten.csv`,
`*_jahr_<X>.csv` (Semikolon, UTF-8 mit BOM → öffnet direkt in Excel).
`cache/` hält jede Rohantwort; ein zweiter Lauf fragt nichts doppelt ab.

## Tabellen

* `items` – ein Eintrag je QID, Hauptklasse, Land/Kreis/Gemeinde, `datum_status`
  (`wikidata` / `label` / `ungenau` / `keins`), `mehrdeutig`
* `item_klassen` – alle Klassen eines Items
* `daten` – **jeder** Wert von P571 und P1249 (mit Präzision und Rang) plus
  das Jahr aus dem Vereinsnamen (`quelle = 'label'`). `bevorzugt` markiert,
  woran die Jubiläen in erster Linie hängen (Orte: P1249 vor P571).
* `jubilaeen` / View `jubilaeen_in_jahr` – `SELECT * FROM jubilaeen_in_jahr WHERE jahr = 2029`
* `abfragen` – Protokoll: welche Region auf welcher Ebene, Timeouts, Fehlschläge

## Abfragestrategie

Pro Land und Klasse eine Abfrage; reißt sie die 60 s von WDQS, wird auf Kreise
geteilt, danach auf Gemeinden. Items, die direkt am Land, an einem
Regierungsbezirk oder einer Verbandsgemeinde hängen, holt eine eigene
„Zwischenebenen“-Abfrage, sonst fielen sie beim Teilen durchs Raster.

## Nachtragen aus Wikipedia

Wikidata kennt die Ersterwähnung fast keines Ortes; der Wikipedia-Artikel
nennt sie meist im Fließtext. `--pruefliste` schreibt für jedes Item ohne
Datum die Kandidatensätze aus dem Artikel in `*_pruefliste.md`. Die Liste
wird nach [pruefanleitung.md](pruefanleitung.md) abgearbeitet (von Hand oder
von Claude), das Ergebnis kommt mit wörtlichem Belegsatz in
[nachtraege.csv](nachtraege.csv). Der nächste Lauf liest die Datei ein
(`quelle = 'wikipedia'`), schon geprüfte Items stehen nicht mehr auf der Liste.

`nachtraege.csv` ist eingecheckt — sie ist Handarbeit, kein Cache.

Bei großen Mengen wird die Prüfliste in `pakete/paket_<nr>.md` geteilt und von
Agenten abgearbeitet (Ergebnis `pakete/ergebnis_<nr>.jsonl`).
`python ergebnisse_pruefen.py` kontrolliert die Ergebnisse (QID im Paket, Jahr im
Beleg, Beleg wörtlich im Paket), `--uebernehmen` hängt die gültigen an
`nachtraege.csv` an; `--modell=claude-haiku` vermerkt in `geprueft_am`, wer
geprüft hat (BW und Ring: Sonnet, Bayern: Haiku).

Stand November 2026: nachgetragen sind ganz BW, die Pfalz, Südhessen (Bergstraße,
Darmstadt-Dieburg, Odenwald, Groß-Gerau, Darmstadt), Rheinhessen/Nahe und die
bayerischen Grenzkreise (Lindau bis Miltenberg, dazu Memmingen). Offen:
Saarland, übriges RLP, übriges Hessen.

Nach dem Übernehmen `python jubilaeen.py` ohne `--nur-auswerten` laufen lassen –
nur so wird die Datenbank mit den Nachträgen neu gebaut (alles aus dem Cache).

## Empfehlungsliste

`python empfehlungen.py 2029` schreibt `empfehlungen_2029.csv`: alle runden
Jubiläen des Jahres, nach Punkten sortiert (Art, Jubiläumszahl, Einwohner,
Verlagsgebiet, Abzüge für unsichere Jahre), mit Kommentar und Beleg. Die
Gewichte stehen oben im Skript und sind eine Faustregel.

## Lead-Listen 2028–2030

`python leads.py 2028 2029 2030` schreibt je Jahr

* `leads_<jahr>.csv` – alle Treffer mit Stufe S/A/B/C/X
* `leads_<jahr>_verlagsgebiet.csv` – S, A, B in BW und der Pfalz, mit Kommentar
  und Kontaktdaten
* `leads_<jahr>_nachbargebiete.csv` – dasselbe für Südhessen, Rheinhessen/Nahe
  und die bayerischen Grenzkreise

Stufen und Kommentare sind von Hand gesetzt und stehen in
`einschaetzung_<jahr>.py` (nach QID, keine Punkteformel). Treffer, die dort noch
fehlen, meldet `leads.py` laut und setzt sie vorläufig auf C.

Kontaktdaten kommen aus `kontakte/` (lokal, nicht eingecheckt): Agenten haben
Impressum und Kontaktseiten gelesen; `python kontakte_holen.py` holt fehlende
Adressen ohne Agenten direkt von den Seiten (nur Adressen der eigenen Domain).
`python kontakte_pruefen.py` ruft jede Fundstelle selbst ab und prüft, ob die
E-Mail-Adresse dort wirklich steht (`kontakte/pruefung.json`).

Werbe-E-Mails an Firmen ohne Einwilligung sind nach § 7 UWG heikel – für den
ersten Kontakt eher Brief oder Anruf, die Adresse dann für die Antwort.
