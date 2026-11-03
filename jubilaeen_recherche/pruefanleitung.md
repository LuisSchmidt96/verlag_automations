# Prüfanleitung: Jahre aus Wikipedia nachtragen

Gilt für jeden, der eine `*_pruefliste.md` abarbeitet — Mensch, Claude oder
ein Hilfs-Agent. Ergebnis: je Item **eine Zeile** in `nachtraege.csv`
(Semikolon, UTF-8). Spalten:

`qid;label;klasse;art;jahr;praezision;unsicher;beleg;bemerkung;geprueft_am`

## Grundregel

**Kein Jahr ohne Beleg.** `beleg` ist der Satz aus der Prüfliste (oder aus dem
Artikel), in dem das Jahr steht — wörtlich, höchstens gekürzt. Jahre aus dem
eigenen Gedächtnis werden nicht eingetragen, auch wenn sie „bekannt“ sind:
genau diese Fehler sieht man später keinem Eintrag mehr an.

## Was eingetragen wird

* `art`: `ersterwaehnung` bei Gemeinden und Ortsteilen, `gruendung` bei
  Vereinen und Unternehmen.
* **Orte:** die erste Erwähnung **des Ortes selbst**. Nicht: Kirche, Burg,
  Pfarrei, Kloster, jüdische Gemeinde, Stadtrecht („1254 erstmals als Stadt“),
  Schultheiß, Wappen, Vereine im Ort.
* **Zusammengelegte Gemeinden** (Gebietsreform, Name ≠ einer der alten Orte,
  z. B. Walzbachtal, Bad Schönborn, Pfinztal, Doppelnamen wie
  Eggenstein-Leopoldshafen): das Jahr eines Ortsteils **nicht** der Gemeinde
  geben — sonst feiert „Bad Schönborn“ 1250 Jahre, die Mingolsheim gehören.
  `jahr` leer, in `bemerkung` die Ortsteile mit Jahr: „Zusammenschluss;
  Mingolsheim 773, Langenbrücken 1269“. Hat der Ortsteil ein eigenes Item auf
  der Liste, bekommt er dort sein Jahr.
  Ausnahme: nennt der Artikel für **alle** Ortsteile dasselbe Jahr („Alle drei
  Ortsteile wurden 1292 erstmals urkundlich erwähnt“, Waldbronn), bekommt es
  die Gemeinde.
* **Gemeinde = Kernort** (Ettlingen, Bretten, Forst): das Jahr des Kernorts
  gilt für die Gemeinde. Auch bei früherem Namen („Philippsburg, bis 1623
  Udenheim“) — die Einleitung steht deshalb mit auf der Liste.
* **Ortsteil ohne eigenen Text:** Viele Ortsteil-Artikel leiten auf die
  Gemeinde weiter oder sind Stubs. Dann gilt der Satz aus dem Gemeindeartikel
  unter der Überschrift des Ortsteils (steht im selben Block der Prüfliste);
  `bemerkung`: „Beleg aus Artikel <Gemeinde>“. Nur Sätze, die eindeutig unter
  diesem Ortsteil stehen — ein Satz über den Nachbarort ist kein Beleg.
* **Gegründete Orte** (Hugenottendorf, Neuanlage nach Hochwasser, Hofgut):
  `art` = `gruendung` mit dem Gründungsjahr.
* **Kein Unternehmen/Verein** (Baudenkmal, historische Schule, die als
  „business“ in Wikidata steht): kein Jahr, `bemerkung` sagt, was es ist.
* **Vereine/Unternehmen:** Gründungsjahr aus Infobox oder Text. Bei
  Fusionen das Jahr, auf das sich das Unternehmen selbst beruft („gegründet
  1876 als …“), die Fusion in `bemerkung`. Steht ein Jahr im Vereinsnamen,
  bestätigen oder korrigieren.

## Präzision und Unsicherheit

* `praezision` 9 = Jahr (auch bei vollem Datum).
* „um 770“, „767/768“, „zwischen 765 und 772“: das frühere Jahr, `unsicher` = 1.
* Nennt der Artikel eine frühere Erwähnung als zweifelhaft und eine spätere als
  „gesichert“: das gesicherte Jahr, die zweifelhafte in `bemerkung`.
* Nur Jahrhundert o. ä. („2. Hälfte des 11. Jahrhunderts“): `jahr` leer,
  `praezision` 7, Zeitraum in `bemerkung`.

## Nichts gefunden

Trotzdem eine Zeile: `jahr` leer, `bemerkung` „kein Beleg im Artikel“. Dann
taucht das Item nicht wieder auf der Prüfliste auf.

`geprueft_am`: Datum und wer, z. B. `2026-10-21 claude`.
