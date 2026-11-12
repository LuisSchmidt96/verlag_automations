# WIP: Was die alten Presselisten zeigen

Stand 04.11.2026. Das alte Modell („wofür ist jeder Kontakt zuständig“) ist durch
die Listen widerlegt und aus dem Tool entfernt. Das Tool schlägt jetzt aus
ähnlichen früheren Büchern vor. Das Ergebnis steht unter *Blindtest des fertigen
Tools*. Vor Änderungen am Vorschlag hier nachlesen.

## Kurzfassung

* Die Auswahl hängt nicht am Verbreitungsgebiet eines Kontakts. Die jetzige
  Regel mit der Agenten-Einschätzung ist schlechter als eine feste Liste der
  60 häufigsten Kontakte.
* Am besten trifft eine sortierte Vorschlagsliste aus **ähnlichen früheren
  Büchern**.
* Eine genaue Nachbildung seiner Listen ist kein sinnvolles Ziel: Zwei Ausgaben
  derselben Zeitschrift stimmen untereinander weniger überein als der Vorschlag
  mit einer neuen Liste.
* Die 1.903 Kontakte von Hand prüfen zu lassen, lohnt sich deshalb nicht.

## Datenlage

* `training_data/Presselisten für Rezensionsversand/`: 360 xlsx, 2020–2026.
  349 sind Auswahllisten, 11 sind die ganze Presseliste (1.620–1.760 Zeilen).
* Die Listen sind Access-Exporte mit Spalte `ID`, lassen sich also exakt
  zuordnen. Es gibt keine Filter und keine Durchstreichungen; ausgeblendet ist
  nur eine Zeile in einer Datei. Jede Zeile ist ein Empfänger.
* 23.300 Listenzeilen: 21.056 in den 1.903 Kontakten, 2.107 mit einer ID, die
  es in Access nicht mehr gibt (2020: 21 %, 2026: 1 %), 117 ohne ID, 20 außerhalb
  des Pressefilters.
* Bücher aus dem **öffentlichen** Shop (1.551 Titel). Die Sitemap endet 2022,
  vollständig ist nur `/Buecher/?p=N`. Die API braucht das Master-Passwort.
* Die Listen entstehen fast immer im Erscheinungsmonat. Damit ließen sich die
  Listen den Büchern zuordnen (36 von 349 unsicher).
* Region, Thema, Gattung und Reichweite je Buch hat Sonnet vergeben, **ohne die
  Empfänger zu sehen**, damit der Test sich nicht selbst bestätigt.

## Wie die Listen entstehen

* **Arbeitsweise:** Die Gesamtliste (`Presse_Stand_…`) wird kopiert, dann wird
  gelöscht, was nicht passt. Darauf deuten die Erstelldaten der Dateien.
* **Größe:** Median 79 (2020), 72 (2022), 58 (2025), 45 (2026). Am größten bei
  Wanderführern und Sachbüchern (rund 65), am kleinsten bei Ortschroniken,
  Jahrbüchern und Lyrik (rund 35).
* **Wenige Aktive:** 802 der 1.903 Kontakte wurden nie gewählt. Seit 2023 wurden
  956 mindestens einmal gewählt, seit 2025 noch 721.
* **Geschwister-Bücher:** Rund 40 Listen sind fast identisch mit einer anderen
  aus demselben Zeitraum (MA-Neckarstadt und MA-Euthanasie: dieselben 71).

Eine Liste besteht ungefähr aus diesen Schichten:

1. **Kern.** 40 Kontakte stehen in mindestens 30 % aller Listen seit 2023. Das
   sind die Kulturredaktionen (BNN, SWR 2, Pforzheimer Zeitung, Mannheimer
   Morgen, RNZ, Rheinpfalz, BZ), ekz, einige VIP-Einzelpersonen, DAMALS, die
   ZGO. Die 30 Häufigsten machen rund 30 % einer Liste aus, die 100 Häufigsten
   rund 60 %.
2. **Großraum.** Kulturredaktionen haben weite, weiche Einzugsgebiete.
   Rheinpfalz-Kultur: Pfalz 100 %, Rhein-Neckar 81 %, Karlsruhe 69 %,
   Württemberg 21 %. BZ-Kultur: Südbaden 85 %, sonst 30–60 %. Schöner Südwesten:
   Württemberg 79 %, Südbaden 85 %, Pfalz 12 %.
3. **Fachblöcke.** Wissenschaft: H-Soz-Kult 93 % (sonst 33 %), ZGO 85 %.
   Wanderführer: Das Wandermagazin 80 % (sonst 1 %). Lyrik: Am Erker, Allmende,
   Chaussée. Biografie: FAZ, Welt, Spiegel-Kultur. Als eigenes Feld braucht es
   die Gattung trotzdem nicht, die Themen tragen dieselbe Information (siehe
   Blindtest).
4. **Lokales.** Der schwer vorhersagbare Rest. Etwa ein Drittel der Empfänger
   einer neuen Liste wurde zuvor weniger als 25-mal gewählt, und genau diese
   Kontakte verpasst der Vorschlag fast immer.
5. **Zeitungsfamilien.** Südkurier, Schwarzwälder Bote und Schwäbische Zeitung
   gehen breiter raus als andere: Wenn gewählt, dann 40–46 % der Familie
   (2023–2024), selten alle. Bei den nahen Blättern wird gezielt gewählt: RNZ
   meist 1, Rheinpfalz 2–3 von 20.

**VIPs:** In 315 von 349 Listen, aber nur 24 Personen, fast alles Journalisten
und Zeitschriften mit zusätzlichem VIP-Haken. 21 davon sind im Pressefilter,
zwei gibt es in Access nicht mehr, einer ist kein Pressesatz (2 Listen). Die
übrigen 4.100 VIPs in Access spielen keine Rolle.

## Regeltest

Gelernt bis Ende 2024, geprüft an 89 Listen ab 2025. *Gefunden* = Anteil seiner
Liste, den das Verfahren trifft. *Richtig* = Anteil des Vorschlags, der wirklich
in seiner Liste stand. *F1* verbindet beides: 0 = nichts gemeinsam, 1 = identisch.

| Verfahren | Empfänger | gefunden | richtig | F1 |
|---|---|---|---|---|
| Jetzige Regel + Agenten-Einschätzung | 121 | 33 % | 13 % | 0,17 |
| Dasselbe, streng UND | 4 | 2 % | 16 % | 0,04 |
| Dasselbe, ODER | 196 | 41 % | 10 % | 0,16 |
| Immer die 60 Häufigsten | 60 | 41 % | 35 % | 0,36 |
| Gelernte Profile, UND | 52 | 34 % | 31 % | 0,31 |
| Gelernte Profile, ODER | 57 | 50 % | 41 % | 0,42 |
| 25 Bausteine, optimal angekreuzt (geschönt) | 60 | 60 % | 48 % | 0,52 |
| **Ähnliche Bücher, je Buch alle Listen davor** | 56 | 52 % | 49 % | **0,48** |
| Beste von 5 vorgeschlagenen Vorlagen (geschönt) | 54 | 54 % | 52 % | 0,51 |
| Beste aller früheren Listen (Obergrenze) | 56 | 63 % | 58 % | 0,59 |

„Geschönt“ heißt: Das Verfahren kennt die Lösung und wählt das Beste aus.

**Beständigkeit als Maßstab.** Ausgaben derselben Reihe mit mehr als 2 Monaten
Abstand stimmen untereinander zu F1 0,33 überein (21 Paare; seit 2023: 0,35 bei
8 Paaren). Identisch sind Listen nur im selben Zeitraum.

**Sortierter Vorschlag:** Unter den ersten 50 stehen 49 % seiner Liste, unter
100 stehen 67 %, unter 150 stehen 75 %, unter 200 stehen 81 %. 98 % einer neuen
Liste standen schon in früheren Listen. Die Einschätzung bringt hier nur 1–2
Punkte. Am schwächsten ist der Vorschlag bei Romanen (0,39) und Wanderführern
(0,37). Ohne die unsicheren Bücher und ohne Geschwister-Listen bleibt er bei
0,46–0,48.

**UND oder ODER:** Die jetzige Regel ist schon UND über die ausgefüllten Achsen.
Sie wirkt wie ODER, weil die meisten Kontakte nur eine Achse haben. Strenges UND
lässt 4 Empfänger übrig. Auch mit gelernten Profilen, bei denen jeder Kontakt
beide Achsen hat, gewinnt ODER (0,42 gegen 0,31). Die Listen sind eine Summe aus
Blöcken, keine Schnittmenge.

## Was widerlegt ist

* **„Engstes passendes Gebiet“** (`einschaetzanleitung.md`). 60 % der echten
  Empfänger passen laut Einschätzung weder nach Region noch nach Thema.
* **„Überregional im Zweifel leer“.** Die Welt 85 Listen, taz 68, Freitag 61.
  Insgesamt wurden 118 der 381 leer eingeschätzten Kontakte gewählt, das sind
  8 % aller Listenzeilen, darunter ekz (244) und VIP-Einzelpersonen (208, 104).
* **Access-Region als Baustein.** Selbst die beste Liste enthält nur 18 % der
  „Odenwald“-Kontakte. Das Feld ist kein Arbeitsfilter.
* **Bausteine zum Ankreuzen** bringen trotz geschönter Annahme kaum mehr als die
  ähnlichen Bücher.
* **Die Hierarchie als Abbild seiner Sicht.** Östringen: 35 Empfänger statt 227,
  niemand aus dem Rhein-Neckar-Kreis, dafür Karlsruher Stadtmagazine.

## Blindtest des fertigen Tools

`python -m presseverteiler.blindtest` prüft `core.vorschlag`, also den Code der
Oberfläche. Je Buch sieht er nur Listen mit früherem Datum. Beleg: In 25
Stichproben wurden die eigene und alle späteren Listen durch Zufall ersetzt, der
Vorschlag blieb gleich. Eingestellt ist an 2023–2024 (Nachbarn 15, Schwelle 0,30,
Alterung 2 Jahre), gemessen an den 89 Büchern ab 2025.

| | Empfänger (Median) | gefunden | richtig | F1 |
|---|---|---|---|---|
| **Vorschlag** | 58 | 53 % | 47 % | **0,47** |
| Immer die 60 bisher Häufigsten | 60 | 42 % | 36 % | 0,37 |
| Alte Regel + Einschätzung (vor dem Umbau) | 121 | 33 % | 13 % | 0,17 |
| Echte Listen | 48 | | | |

* Durch die echte Oberfläche gefahren (Häkchen setzen, Haken ablesen): 89 von 89
  Büchern gleich wie der Kern.
* Ohne fast gleiche frühere Listen: 0,45. 2025 und 2026 getrennt: 0,48 und 0,47.
* F1 je Buch: schlechtestes Viertel bis 0,38, bestes Viertel ab 0,57.
* Unter den ersten 50 stehen 50 % seiner Liste, unter 100 stehen 67 %, unter 150
  stehen 77 %, unter 200 stehen 82 %.
* Gut: Biografie 0,53, Wissenschaft 0,52, Sachbuch 0,50. Schwach: Wanderführer
  0,33, Lyrik 0,34, Roman 0,38. Die schlechtesten Fälle sind Gegenden ohne
  ähnliche Vorgänger (Rüsselsheim, Heuberg) und sehr kleine Listen.
* Ohne Themen 0,36, ohne Regionen 0,40: Es braucht beides.

**Gattung gestrichen.** Der Plan sah sie als neue Achse vor. Gemessen bringt sie
nichts, solange Themen da sind (mit 0,47, ohne 0,48, an 2023–2024 genauso). Beim
Buch kreuzt man also nur Regionen und Themen an. Stärker gewichtete Nachbarn
machen es schlechter (0,481 → 0,451), mehr Nachbarn ändern nichts.

**Wie blind ist das?** Die Bücher ab 2025 waren schon in der Auswertung sichtbar,
und auch die Entscheidung gegen die Gattung fiel mit Blick auf sie. Die Werte sind
aber an 2023–2024 eingestellt, und dort ergibt sich dasselbe Bild (0,48). Ganz
frisch wird es erst mit den nächsten echten Listen.

## Weitere Haken: geprüft, keiner lohnt sich

Geprüft am 04.11.2026, Skripte und Etiketten in
`presseverteiler/einschaetzung/lernen/haken/` (nicht eingecheckt). Ausgesucht wurde
nur an 2023–2024 (105 Bücher, F1 0,481), die Sieger gingen danach an die 89 Bücher
ab 2025. Der Versuchsstand bildet `core.vorschlag` ohne Abweichung nach. Die neuen
Merkmale hat Sonnet aus den Shop-Texten vergeben, ohne die Empfänger zu sehen.

| Kandidat | Änderung F1 (2023–2024) |
|---|---|
| Ort genau, Reichweite, gleicher Autor, Umfang | ±0,00 |
| Träger (Gemeinde/Verein/Firma …), Anlass, Reihe, Hauptperson, feine Stichworte | ±0,00 |
| Herausgeber, Epoche, Autorenart, Bildanteil, Zielgruppe | höchstens +0,005 |
| Lange nicht Gewählte streichen, Grundrate einmischen, Zeitnähe | ±0,00 |
| Klappentext zusätzlich zu den Schlagworten | +0,003 |
| PLZ-Nähe zum Buchort (Reihenfolge) | Top 100: +0,002 |
| Presse am Wohnort des Autors | ±0,00 |
| Deckel je Redaktion | −0,014 |
| Listengröße in 3 Stufen, **perfekt gesetzt** | +0,01 |
| Haken je Zeitungsfamilie, **perfekt gesetzt** | −0,02 bis −0,05 als „alle oder keiner“, +0,02 als reines Abwählen |
| Haken je Medienart, **perfekt gesetzt** | ±0,00 |

* **Gegenprobe ab 2025:** Herausgeber −0,003, Herausgeber + Epoche + Autorenart
  −0,002. An 2023–2024 waren das +0,005 und +0,007, also vermutlich Zufall aus 33
  Varianten.
* **Vermutlich keine Frage der Etikettenqualität:** 156 von 348 Etiketten sind als
  unsicher markiert. Nur mit den sicheren (54 Bücher) kommt nichts über +0,006.
* **Obergrenze:** Selbst die im Nachhinein besten Nachbarn ergeben nur F1 rund 0,60.
  Auch die genaue Listengröße brächte nur 0,50.
* **Nähe ist ein schwaches Signal:** Früher gewählte, nicht vorgeschlagene Kontakte
  aus der PLZ-Gegend (3 Stellen) stimmen zu 20 % (Buchort) und 18 % (Autor, Zufall
  8 %). Das sind rund 10–12 je Buch, zu wenig zum Vorankreuzen. Die Autorenadresse
  ließ sich nur bei 31 von 105 Büchern zuordnen, im Shop steht sie fast nie.
* **Klappentext allein** erreicht ohne Schlagworte 0,43 (ab 2025: 0,41). Er trägt
  dieselbe Information wie die Schlagworte, nichts Zusätzliches.

Was noch fehlt, steht also weder in den Buchdaten noch im Access-Export. Neue Haken
erst wieder prüfen, wenn es eine neue Datenquelle gibt, siehe die ersten beiden
offenen Fragen.

## Umsetzung

1. Beim Buch Regionen und Themen ankreuzen.
2. Das Tool zeigt die 15 ähnlichsten Bücher und eine sortierte Liste. Vorangekreuzt
   ist, wer bei mindestens 30 % davon dabei war.
3. Er prüft und ergänzt: Die Liste eines ähnlichen Buchs lässt sich übernehmen,
   und die Suche nimmt mit *Alle Treffer hinzufügen* ganze Zeitungsfamilien auf.
4. Jede gespeicherte Liste kommt in `Pressehistorie.xlsx` und zählt ab dem
   nächsten Tag.

Der Reiter *Kontakte* (vormals *Einordnen*) ist keine Pflicht mehr. Die
Einschätzung bleibt als Suchhilfe und rückt Kontakte aus der Gegend in der Liste
etwas nach oben.

## Offene Fragen an den Pressereferenten

* Nennen Autoren oder Auftraggeber selbst Pressekontakte, etwa im Autorenfragebogen?
  Gibt es das schriftlich? Das wäre die naheliegendste Quelle für die lokalen
  Empfänger, die der Vorschlag verpasst.
* Wird festgehalten, wer tatsächlich rezensiert hat? Dann ließe sich an der Wirkung
  messen statt an der Übereinstimmung mit seinen Listen.
* Warum unterscheiden sich Listen derselben Reihe so stark? Absicht oder Zufall?
* Was heißt `Direktversand`? Fast nur 2022, rund 10 % der Zeilen, vor allem
  große Kulturredaktionen.
* Was heißt Gelb? In 7 von 12 Listen stammen die gelben Zeilen aus einer Gegend,
  sonst sind sie gemischt.
* Seit 2023 sind alle Listen unter demselben Benutzerkonto gespeichert, davor
  unter mehreren. Stammen die älteren von jemand anderem?
* Sollen die 802 nie gewählten Kontakte überhaupt noch angezeigt werden?
* Ist „Freiburger Geschichtsblätter“ (Fribourg, Schweiz) gewollt? 35 Listen,
  davon 11 mit Bezug zu Freiburg im Breisgau. Eine Verwechslung ist möglich.

## Vorbehalte

* Der Test misst Übereinstimmung mit seinen Listen, nicht deren Qualität. Er
  nennt seine Arbeit selbst fehleranfällig. Wer nicht auf der Liste stand, war
  nicht unbedingt falsch.
* Die Beständigkeit beruht auf 21 Paaren aus 4 Reihen. Einzelne Ausgaben können
  bewusst andere Themen haben.
* Geprüft sind 89 Bücher. Bei Wanderführern (5) und Ortschroniken (5) ist das
  wenig. Die Etiketten stammen von einem Modell, nicht vom Verlag.
* Die Zahlen beziehen sich auf den Access-Export von August 2026.

## Nebenbei geklärt

* **IfU** heißt Institut für Unternehmensgeschichte (Seite im Shop).
* Die Frage nach „Landkreis Tuttlingen“ erübrigt sich mit dem neuen Vorschlag.

## Wiederholen

Der Blindtest läuft mit `python -m presseverteiler.blindtest`. Die Auswertung
davor liegt in `presseverteiler/einschaetzung/lernen/` (nicht eingecheckt, `listen.pkl`
enthält personenbezogene Daten). Die Skripte dort sind der Stand vor dem Umbau und
rufen teils die alte Regel auf. Die Skripte laufen im
Repo-Wurzelordner und bekommen den Ordner als Argument, `regeltest.py`,
`warum.py`, `bausteine.py`, `grenze.py` und `gegencheck.py` zum Beispiel so:

```
python3 presseverteiler/einschaetzung/lernen/regeltest.py presseverteiler/einschaetzung/lernen
```

`verhalten.py`, `luecken.py`, `gegenprobe.py`, `struktur2.py`, `vorlagen.py` und
`vip.py` bekommen stattdessen den Pfad zu `listen.pkl`.
