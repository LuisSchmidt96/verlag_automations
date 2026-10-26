# jubilaeen_recherche

Recherche-Skript (Kommandozeile, keine .exe): holt Unternehmen, Vereine,
Gemeinden und Ortsteile in BW, RLP, Saarland und Hessen mit Gründungs- bzw.
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

## Empfehlungsliste

`python empfehlungen.py 2029` schreibt `empfehlungen_2029.csv`: alle runden
Jubiläen des Jahres, nach Punkten sortiert (Art, Jubiläumszahl, Einwohner,
Verlagsgebiet, Abzüge für unsichere Jahre), mit Kommentar und Beleg. Die
Gewichte stehen oben im Skript und sind eine Faustregel.
