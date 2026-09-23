# Honorar-Abrechner

Rechnet einmal im Jahr die Autorenhonorare ab und erzeugt Briefe,
Zahlungsliste und die Meldung an die Künstlersozialkasse.

Ersetzt die Excel-Mappe `Honorare_<Jahr>.xlsx`, in der die Geschäftslogik
bisher in von Hand getippten Zellformeln steckte.

## Ablauf

Die vier Reiter sind die vier Schritte, und ihre Beschriftung zeigt den
Stand: ✓ fertig, ◐ angefangen, · noch nichts. Darunter steht immer **eine**
Zeile „Als Nächstes: …". Ein fünfter Reiter *Meldungen* sammelt alles, was
das Werkzeug zu sagen hat.

1. **Einmalig:** *1. Autoren und Bücher → Daten aus der alten Excel-Tabelle
   holen* — liest `Honorare_2025.xlsx` mit allen Blättern ein, inklusive der
   Jahreshistorie ab 2003. Danach ist die Altmappe Archiv.
2. **Jeden Januar:**
   1. *2. Zahlen eintragen* — verkaufte Exemplare und Eigenkauf je Buch.
      Wer lieber in Excel tippt, nimmt das Blatt `Abrechnung <Jahr>` im
      Bestand; das Werkzeug liest die Zahlen beim Laden von dort.
      Doppelklick öffnet das Feld, die Eingabetaste übernimmt und öffnet
      sofort die nächste Zeile, Tabulator springt zum Eigenkauf, Esc
      verwirft. „Zur ersten Zeile ohne Zahl" findet die Stelle wieder.
   2. *3. Nachrechnen → Beträge berechnen* — zeigt je Empfänger den Betrag.
      Grau = kein Brief, rot = bitte ansehen. **Doppelklick auf eine Zeile
      zeigt in ganzen Sätzen, wie der Betrag zustande kommt.**
   3. *4. Briefe und Listen* — Briefe, dann (unter Windows) PDF, dann die
      beiden Listen.
   4. *1. Autoren und Bücher → Speichern.* Beim Schließen wird gefragt, falls
      es vergessen wurde.

Neue Autoren und Bücher legt man auf Reiter 1 an („Neuer Autor …", „Neues
Buch …"). Das Detailfenster eines Autors — Doppelklick auf seine Zeile —
zeigt seine Bücher als Liste; Doppelklick auf ein Buch ändert dessen
Angaben, „Angaben ändern …" die Anschrift.

**Ein Titel, der nicht mehr abgerechnet wird**, gehört *stillgelegt*, nicht
gelöscht: im Buchdialog der Haken „Wird nicht mehr abgerechnet" plus ein
Grund („Titel vergriffen", „verstorben"). Das Buch verschwindet dann aus
Reiter 2 und aus jeder Berechnung, aber die erfassten Jahre bleiben — in
fünf Jahren lässt sich noch nachsehen, was 2025 gezahlt wurde. Der zweite
Haken, „Wird gesondert abgerechnet", ist für die Verträge aus
`Zahlung ab XX Ex.`: sie zählen für Staffel und Freimenge mit, lösen aber
keine Auszahlung aus und erscheinen im Blatt `Offene Sonderfälle`.

**Wirklich entfernen** — „Autor entfernen …" auf Reiter 1, „Buch entfernen …"
im Detailfenster — ist nur für das, was versehentlich angelegt wurde. Die
Rückfrage zählt auf, wie viele Jahre dabei verloren gehen, und nennt das
Stilllegen als Alternative.

Der Buchdialog ist nach der Frage gebaut, die tatsächlich zu beantworten
ist — **was bekommt der Autor je Exemplar?** — und nicht als Feldliste:

```
Was der Autor je Exemplar bekommt
  ( ) Ein fester Betrag je Exemplar        [      ] €
  (•) Ein Anteil am Verlagsabgabepreis     Ladenpreis [ 16,90 ] €
                                           darin MwSt [ 7 % ]
                                           Verlagsrabatt [ 40 ] %
Honorarsatz — gleichbleibend oder nach Menge gestaffelt
  ( ) Immer derselbe Satz                  [ 12 ] %
  (•) Gestaffelt                    bis [ 1500 ] Ex.  [ 12 ] %
                                    bis [ 2500 ] Ex.  [ 13 ] %
                                    bis [      ] Ex.  [ 14 ] %   ← nach oben offen
Ergibt je Exemplar — 1–1500: 1,14 €, 1501–2500: 1,23 €, ab 2501: 1,33 €
```

**Die Staffel ist kein eigenes Feld, sondern der Honorarsatz selbst** — einer,
der sich mit der Menge ändert. Deshalb stehen beide an derselben Stelle und
schließen einander aus.

Die Zeile darunter rechnet bei jeder Eingabe mit und zeigt **Eurobeträge, nicht
Prozente**. Das ist der eigentliche Prüfstein: ob 12 % vom Nettoabgabepreis
richtig ist, sieht niemand — ob 1,14 € je Exemplar stimmt, schon.

Beim Import bleibt das Deuten des Vertragstextes („bis 2500 Ex. 12 %, ab 2501
Ex. 13 %") natürlich, denn dort *existiert* der Text bereits. Beim Eingeben
wird nichts geraten.

## Was gerechnet wird

```
Vergütungsexemplare = verkauft − Eigenkauf + Korrektur
Betrag netto        = Vergütungsexemplare × Betrag je Exemplar
MwSt                = 7 % (vereinzelt 19 %), nur bei MwSt-pflichtigen
Auszahlung          = netto + MwSt − offene Vorauszahlung
```

Der **Betrag je Exemplar** ergibt sich entweder aus einem festen Wert oder
aus den Bestandteilen des Vertrags:

```
Ladenpreis → MwSt herausrechnen → minus Verlagsrabatt (40 %)
           → mal Honorarsatz → geteilt durch Zahl der Mitautoren
```

**Gerundet wird kaufmännisch**, also die Hälfte immer auf: aus 5,635 € wird
5,64 €. Pythons eingebautes Runden würde 5,63 € ergeben (es rundet zur geraden
Ziffer) — bei fünfhundert Posten und einer Excel zum Gegenrechnen ist das ein
Cent, den man sucht. Wo im Altbestand ein `ROUND(...)` um die Satzformel steht,
wird der Betrag je Exemplar vor der Multiplikation gerundet, sonst nicht; der
Unterschied macht bei 20 Exemplaren acht Cent aus und wird je Buch mitgeführt.

Dazu kommen drei Regeln, die die Altmappe kannte, aber nicht anwandte:

* **Schwelle.** Ist am Buch „Keine Berechnung bei 10 o. weniger“ gesetzt und
  liegen die Vergütungsexemplare **unter** zehn, entfällt das Honorar. Bei
  genau zehn wird gezahlt.
* **Staffel.** Steht im Vertrag „bis 2500 Ex. 12 %, ab 2501 Ex. 13 %“, wählt
  das Werkzeug den Satz nach dem kumulierten Stand **zu Jahresbeginn** und
  wendet diesen einen Satz auf die ganze Jahresmenge an — so, wie es bisher
  von Hand gemacht wurde. Ob stattdessen tranchenweise zu rechnen wäre, ist
  eine offene Frage (s. u.).
* **Freimenge.** „Ab dem 201. verkauften Exemplar“ heißt: die ersten 200
  Exemplare sind frei. Der Saldo wird über die Jahre fortgeschrieben und darf
  jahrelang negativ bleiben; gezahlt wird erst, wenn er ins Plus dreht.

**Kein Brief** geht hinaus, wenn der Auszahlungsbetrag null oder negativ ist
oder kein Buch erfasst wurde. Diese Empfänger stehen mit Begründung im
`protokoll.xlsx` — damit sichtbar bleibt, dass sie nicht vergessen wurden.

## Der Bestand ist eine Excel-Mappe

`Honorarbestand.xlsx` liegt **neben der .exe** und darf von Hand geöffnet und
geändert werden. Sieben Blätter, jedes mit genau einer Aufgabe:

| Blatt | Inhalt |
|---|---|
| `Hinweise` | die Regeln im Klartext — steht absichtlich an erster Stelle |
| **`Abrechnung <Jahr>`** | **das Arbeitsblatt**: eine Zeile je Buch, nach Autor sortiert. Links wer und was, in der Mitte die drei **gelb hinterlegten** Spalten zum Ausfüllen, rechts das Ergebnis der letzten Berechnung und eine Spalte `Besonderheit` |
| `Historie` | alle Jahre davor — vollständig, aber aus dem Weg |
| `Regeln` | **nur** die Bücher mit einer Besonderheit, nach Tragweite geordnet: erst Vorauszahlung und Staffel, zuletzt die stillgelegten. Je Zeile ein Satz, der erklärt, was gilt |
| `Empfänger` | Anschrift, IBAN, Aktenzeichen, E-Mail |
| `Bücher` | die technischen Konditionsfelder — hier rechnet das Werkzeug, hier muss man normalerweise nichts tun |
| `Staffeln` | die Stufen der Staffelverträge |

Im Arbeitsblatt wird **nur in drei Spalten** getippt: `verkaufte Ex.`,
`Eigenkauf` und `Korrektur`. Sie sind gelb hinterlegt, alles andere ist
Anzeige und wird beim nächsten Speichern neu berechnet. Getippt werden kann
wahlweise dort oder in der Eingabemaske des Werkzeugs — beides landet an
derselben Stelle.

Die Spalte **`Stand bis Vorjahr`** nimmt einem das Nachschlagen ab: sie zeigt,
wie viele Vergütungsexemplare bis zum Vorjahr aufgelaufen sind. Daran hängen
Staffel und Freimenge. Gerechnet wird der Wert aus dem Blatt `Historie`; man
muss ihn nicht pflegen.

Im Blatt `Regeln` steht bewusst **nicht**, wie gewöhnlich gerechnet wird —
Ladenpreis, Rabatt, Satz stehen im Blatt `Bücher` und im Werkzeug hinter dem
Doppelklick. Stünden sie auch dort, wären es 412 statt 275 Zeilen, und die 24
Staffelverträge lägen darin begraben.

Stillgelegte Bücher stehen **nicht** im Arbeitsblatt — sonst wäre es um ein
Fünftel länger, ohne dass je etwas einzutragen wäre. Ihre Zahlen liegen
vollständig in der `Historie`, und warum sie stillgelegt sind, steht in
`Regeln`.

Verknüpft wird über die Kennungen (`E0042`, `B0137`). Die dürfen **nicht**
geändert werden — sonst verliert ein Buch seinen Autor.

Zwei Dinge, die man wissen muss:

* **Leer ist nicht null.** Eine leere Zelle bei `verkaufte Ex.` heißt „noch
  nicht eingetragen“. Hat sich ein Buch wirklich nicht verkauft, gehört dort
  eine `0` hinein — sonst fehlt es in der Abrechnung.
* **Die Datei schließen, bevor das Werkzeug speichert.** Solange sie in Excel
  offen ist, kann nicht geschrieben werden; das Werkzeug sagt das dann auch,
  statt abzustürzen. Vor jedem Speichern entsteht `Honorarbestand.bak.xlsx`.
* **Die Arbeit geht nicht verloren.** Alle zwanzig eingetippten Zahlen
  schreibt das Werkzeug still einen Zwischenstand nach
  `Honorarbestand.wiederherstellung.xlsx`. Wurde das Programm zuletzt ohne
  Speichern beendet, bietet es beim Start an, diese Eingaben zurückzuholen;
  nach einem richtigen Speichern verschwindet die Datei. Beim Schließen wird
  ohnehin gefragt.

## Ausgaben

Alles nach `honorar_output/<Jahr>/` neben der .exe:

| Datei | Inhalt |
|---|---|
| `briefe/<Name>_<Art>-<Jahr>.docx` | je Empfänger ein Brief, Seite 1 hoch, Seite 2 quer |
| `briefe/….pdf` | dieselben Briefe als PDF — **nur unter Windows mit Word** |
| `Zahlungsliste_<Jahr>.xlsx` | **drei Blätter in einer Datei** (s. u.) |
| `import_protokoll.xlsx` | nur beim Import: jede Zeile und jede Unklarheit |

Die drei Listen gehören zusammen — dieselbe Abrechnung, dasselbe Jahr,
derselbe Arbeitsgang — und stehen deshalb als drei Blätter in **einer** Datei:

| Blatt | Inhalt |
|---|---|
| `Zahlungsliste <Jahr>` | für die Überweisungen, Aufbau wie bisher |
| `Künstlersozialkasse <Jahr>` | Buchungsformat, drei Abschnitte (19 % / 7 % / ohne) |
| `Protokoll <Jahr>` | wer bekommt einen Brief, wer nicht, und warum |
| `Offene Sonderfälle <Jahr>` | die Beträge, die auf eine Entscheidung warten — mit Buch, Betrag, Bankverbindung und der Herkunftszeile in der Altmappe |

Für die KSK zählen **nur** Zeilen mit Vergütungsart „Honorar“ — Rückflüsse,
Erlösanteile und Darlehensrückzahlungen sind keine Honorare im Sinne der KSK.

## Die Briefvorlage

`vorlagen/brief_vorlage.docx` trägt Platzhalter (`{{ANSCHRIFT}}`, `{{BETRAG}}`,
`{{IBAN}}` …), jeder in genau **einem** Run — sonst zerfällt er beim Ersetzen,
weil Word Text gern mitten im Wort aufteilt. Die Titeltabelle hat eine
Musterzeile, die je Buch geklont wird; die MwSt-Spalte entfällt bei
Empfängern ohne Steuerpflicht.

> **Die Vorlage ist ein Nachbau** aus den beiden Muster-PDFs, weil die echte
> Word-Vorlage des Verlags noch nicht vorliegt. Sobald sie da ist, wird sie an
> dieselbe Stelle gelegt — die Platzhalter müssen dann nur gleich heißen.
> `core.baue_briefvorlage()` erzeugt den Nachbau neu, falls er verlorengeht.
>
> Der Nachbau trägt eine Kennzeichnung in den Dokumenteigenschaften. Solange
> sie da ist, schreibt das Werkzeug bei jedem Brieflauf einen Hinweis ins
> Protokoll; mit der echten Vorlage verschwindet er von selbst, ohne dass
> jemand daran denken muss.

## Eigenheiten des Altbestands

Der Import rechnet die alten Zellformeln in Konditionen zurück. Von 252
Satzformeln gehen 246 vollständig auf; der Rest wird als fester Betrag je
Exemplar übernommen und zur Nachpflege vermerkt. Dabei stolpert man über
Dinge, die man kennen sollte:

* **Rundung hängt an der Formel.** Meist steht `ROUND(...)` um die Satzformel,
  bei einigen Zeilen aber nicht — dort rechnet Excel mit dem vollen Bruch
  weiter. Bei 20 Exemplaren macht das acht Cent aus. Das Werkzeug führt die
  Unterscheidung in der Spalte `Satz runden` mit.
* **Die Altmappe rechnet uneinheitlich.** Manche Zeilen ziehen die MwSt aus dem
  Ladenpreis heraus, andere rechnen auf den Bruttopreis, eine schlägt sie
  sogar auf. Das wurde 1:1 übernommen, nicht vereinheitlicht.
* **Wer zu einem Brief gehört, stand nirgends sauber.** Die einzige Quelle war
  die von Hand gebaute Summenkette `=X108+Z108+X109+…`. Eine davon fasst zwei
  verschiedene Personen zusammen — der Import meldet es und übernimmt es so,
  wie die Mappe es tat.
* **Einzelne Zellen sind kaputt.** Bei „Landau in der Pfalz“ ist die
  Vergütungsexemplar-Zelle leer, obwohl −6 Exemplare dastehen; Excel rechnet
  deshalb 0, das Werkzeug −10,02 €. Solche Fälle stehen im Protokoll.

## Was beim Abgleich mit der Altmappe herauskam

Die Gegenprobe deckt drei Dinge ab: jede Zeile, jede Auszahlung **und beide
Ausgabelisten**. Die letzten beiden kamen erst spät dazu — und genau dort
steckten die Fehler, die eine reine Rechenprobe nie gefunden hätte:

* **Der Empfänger des LUBW-Betrags hieß falsch.** Das LUBW-Blatt nennt nur
  Titel und Beträge, nicht den Zahlungsempfänger. Der steht in der
  Zahlungsliste: *Stiftung Naturschutzfond BW*, Vergütungsart *Spende*, mit
  eigener Bankverbindung. Der Import holt ihn jetzt von dort.
* **Die Anthologien hätten 15 Zahlungen doppelt ausgelöst.** Der Topf von
  „Tödliche Häppchen" (44,02 €) wurde am 27.01.2022 ausgeschüttet. Ein Import,
  der daraus Beträge für das laufende Jahr macht, verschickt dasselbe Geld ein
  zweites Mal. Beteiligte und Bankverbindungen werden übernommen, der Anteil
  aber **nicht** automatisch verteilt.
* **Ein Autor verlor seine Auszahlung durch eine Freimenge.** Wenn die
  Freimenge noch nicht erreicht ist, stand dort ein negativer Betrag — bei
  einem Empfänger −1.045,28 €, was seine 6,69 € auffraß. Das ist falsch: man
  schuldet dem Verlag nichts für Exemplare, die nie vergütet wurden. Der Saldo
  wird vorgetragen, der Betrag ist null.
* **Ein Brief ging an „verschiedene Autoren".** So heißen in der Altmappe die
  Sammelzeilen der Anthologien. Solche Posten bekommen jetzt weder Brief noch
  Überweisung; der Betrag steht mit Begründung im Protokoll und ist von Hand
  zu verteilen.
* **Das Blatt „Zahlung ab XX Ex." ist keine Zahlungsquelle.** Es ist ein
  Laufzettel: ein Zähler läuft auf die vereinbarte Freimenge zu, und die
  Betragsspalte zeigt den Stand mal Satz — bei einer Autorin 1540 Exemplare
  × 1 € = 1.540 €, jedes Jahr aufs Neue. Die Summe der ganzen Spalte ist
  **−9.954 €**, und nur vier der dort geführten 54 Personen stehen überhaupt
  in der alten Zahlungsliste, mit anderen Beträgen. Der Import übernimmt
  diese 68 Bücher deshalb mit dem Haken **„Gesondert abrechnen"**: sie zählen
  für Historie und Staffel, lösen aber keine Überweisung aus.
* **Der kumulierte Vortrag wurde gar nicht gelesen.** In jenem Blatt steht in
  der Spalte „Vergütungsexemplare" nicht die Jahresmenge, sondern der
  aufgelaufene Stand. Ohne ihn wurde aus einem Jahr mit Rückgaben („−69 Ex.")
  ein negativer Saldo, obwohl 1540 Exemplare aufgelaufen waren.
* **Der KSK-Meldung fehlte ein ganzes Konto.** Getrennt wird nicht nach der
  Höhe des Steuersatzes, sondern danach, **ob** der Autor
  mehrwertsteuerpflichtig ist. Wer es nicht ist, gehört auf Konto 4782
  („Honorare", ohne Umsatzsteuer) — im Altbestand 54 der 74 Buchungen.

Stand nach diesen Korrekturen, für 2025:

| | |
|---|---|
| Zeilen und Auszahlungen | stimmen, bis auf eine kaputte Zelle in der Quelle |
| Zahlungsliste | **88 von 91 Beträgen identisch** |
| KSK-Meldung | **69 von 72 Personen identisch** mit der KSK-Spalte der Altmappe |
| unerklärte Abweichungen | **3** — und die sind ein einziger Sachverhalt (s. u.) |

## Nach jeder Änderung: der Probelauf

```
python honorar_abrechner/probelauf.py
```

Fährt den ganzen Weg gegen die Beispieldaten und prüft 30 Werte gegen
festgeschriebene Sollzahlen — Import, Rundlauf, Berechnung, Gegenprobe,
Briefe, Listen und das Lesen eingetippter Zahlen. Weicht etwas ab, sagt das
Skript was, und der Rückgabewert ist 1. Ohne die Beispieldaten
(personenbezogen, nicht eingecheckt) überspringt es und sagt das.

Das ist der Wächter: ohne ihn bricht die nächste Änderung eine der sechs
mühsam gefundenen Regeln, ohne dass es jemand merkt.

## Gegenprobe

*Durchlauf → Gegen die Altmappe prüfen* rechnet jede Zeile und jede
Auszahlung des Jahres gegen die alte Mappe nach. Stand der letzten Prüfung
für 2025: **alle 275 Zeilen und alle 137 Auszahlungen stimmen auf den Cent**,
bis auf den einen oben genannten Fall mit der kaputten Zelle.

Prüfwerte zum Nachschlagen: Dieter Buck 2.368,98 € brutto / 2.214,01 € netto
(entspricht dem Muster-PDF und dem KSK-Blatt).

**Drei Abweichungen bleiben, und sie gehen alle auf dieselbe Ursache zurück —
ein Rechenfehler ist keine davon:**

* **Möller / Schenk** (46,63 € verschoben) — die Summenkette `AA245` der
  Altmappe fasst zwei verschiedene Personen zusammen. Welche der beiden das
  Geld bekommen soll, muss der Verlag entscheiden.
Der Fall **Stefan Schaupp** wird inzwischen automatisch erkannt und mit
Begründung ausgewiesen, statt als offene Abweichung dazustehen: dort ist die
**Altmappe falsch** —
  gemeldet wurde sein Auszahlungsbetrag statt seiner Honorarsumme, also
  36,92 € Rückfluss, die der Künstlersozialkasse nicht zu melden sind. Die
  Excel selbst rechnet in ihrer eigenen KSK-Spalte 202,94 € — wie dieses
  Werkzeug.

Die Gesamtsumme liegt bei **19.142,18 € für 92 Briefe**. Der Unterschied zur
alten Zahlungsliste sind im Wesentlichen die 128 € für die Stiftung
Naturschutzfond BW, die vorher auf dem LUBW-Blatt standen. Empfänger, deren
Bücher auf mehreren Blättern der Altmappe verteilt waren, bekommen jetzt
einen einzigen Brief; das Werkzeug weist das bei ihnen aus.

## Offene Punkte für den Verlag

1. **Die Verträge aus „Zahlung ab XX Ex." (68 Bücher, 741,67 € an 20
   Empfänger).** Zu finden sind sie an drei Stellen: im Reiter *3.
   Nachrechnen* über das Häkchen **„nur die offenen Sonderfälle"** (braun
   markiert), im Blatt `Offene Sonderfälle <Jahr>` der Listendatei — dort
   mit Buch, Betrag, Bankverbindung und Herkunftszeile — und im Blatt
   `Regeln` des Bestands. Das ist nötig, weil **19 der 20 gar keinen Brief
   bekommen** und sonst zwischen 250 grauen Zeilen verschwänden. Sie sind als
   „Gesondert abrechnen" übernommen und lösen derzeit **keine** Zahlung aus —
   die sichere Annahme, weil die Altmappe dort nur einen Stand führt und fast
   niemand davon in der Zahlungsliste auftaucht. Wie diese Verträge wirklich
   abgerechnet werden, muss der Verlag sagen; danach den Haken in der Spalte
   „Gesondert abrechnen" entfernen. **Das Werkzeug weist den Betrag bei
   jedem Durchlauf aus** („20 Empfänger bekommen zusammen 741,67 € NICHT"),
   damit die offene Entscheidung nicht in einem Protokoll versandet.
2. **Staffel — ein Satz oder tranchenweise?** Wird beim Überschreiten der
   Grenze der höhere Satz nur auf die darüberliegenden Exemplare angewandt
   oder auf die ganze Jahresmenge? Umgesetzt ist derzeit „ein Satz je Jahr“.
3. **Die kumulierten Stände.** In den Notizen steht „Stand 2023: 1460 Ex.“ —
   für 2024 und 2025 fehlt die Fortschreibung bei den Büchern, die nicht im
   Historienblatt stehen.
4. **Die echte Word-Vorlage** mit dem Briefkopf.
5. **Die Summenkette `AA245`** fasst Schenk und Möller zusammen. Fehler?
6. **Vertauschte Namen.** In der Altmappe steht bei „Mühlacker+“ der Nachname
   im Vornamensfeld; der Brief heißt deshalb `Sandra, Schuster_…` statt
   `Schuster, Sandra_…`. Das Werkzeug übernimmt die Felder so, wie sie
   dastehen — wer sie tauscht, tut es im Bestand.
7. **Zwei Empfänger ohne Bankverbindung**, zwei mit **mehreren
   Vergütungsarten** in einem Brief — wie soll damit umgegangen werden?
8. **Unterzeichnerin.** Die Musterbriefe zeichnet Silke Freitag; im
   `pi_bi_generator` wurde der Absender zwischenzeitlich geändert.
9. **Lexware-Artikelexport.** Die Zahlen gibt es — nur noch nicht als Datei.
   Einzelheiten unter „Der Weg zu den Stückzahlen".
10. **Die Anthologie-Töpfe.** Soll in diesem Jahr etwas ausgeschüttet werden?
    Das Werkzeug verteilt nicht von selbst; die Beteiligten und ihre
    Bankverbindungen liegen aber bereit.
11. **Stefan Schaupp und die Künstlersozialkasse.** Für 2025 wurden 36,92 €
    Rückfluss mitgemeldet, die dort nicht hingehören. Ob das zu berichtigen
    ist, entscheidet der Verlag.

## Der Weg zu den Stückzahlen

Derzeit werden `verk. Ex.` und `Eigenkauf` von Hand erfasst. Was dabei über
Lexware herausgefunden wurde, damit es beim nächsten Anlauf nicht noch einmal
gesucht werden muss:

* **Der Auftragsexport taugt nicht.** „Aufträge Verkauf <Jahr>" liegt auf
  Belegebene: 4639 Zeilen, 4639 verschiedene Belegnummern, keine ISBN, keine
  Mengenspalte. Die einzige Zahl ist der Rechnungsbetrag.
* **„Aufträge zum Artikel" auch nicht.** Dieser Bericht filtert zwar auf einen
  Titel, zeigt aber die Summen des **ganzen Belegs**: alle acht Beträge des
  Musterexports sind identisch mit dem Belegbetrag im Jahresexport. Eine
  Amazon-Sammelrechnung über 8.757,03 € ist eben nicht ein Titel.
* **Die Mengen stehen im Artikelstamm.** In der Artikelansicht, unten im
  Reiter **„Kunden"**, steht bei gesetztem Jahresfilter genau das Gewünschte:
  `Kunden-Nr. | Matchcode | Menge` — also die Stückzahl je Kunde für diesen
  Titel. Damit wäre auch der **Eigenkauf** automatisch abgrenzbar, nämlich
  über die Kundennummer des Autors. Ein direkter Export dieser Ansicht scheint
  es nicht zu geben; zu prüfen wäre der Nachbarreiter „Umsätze" und ob sich
  die Ansicht über einen Bericht für **alle** Artikel auf einmal ausgeben
  lässt. Je Titel einzeln wäre bei rund 490 Büchern keine Erleichterung.
* **Der Matchcode ist die Verlagskurzform.** In der Artikelliste steht als
  Matchcode `05-565`, `71-1`, `18-9` — dasselbe Format wie die Spalte `ISBN`
  im Honorarbestand. Ein späterer Import kann also darüber zuordnen und
  braucht die vollständige ISBN nicht.

## Bekannte Grenzen

* Die **PDF-Umwandlung braucht Windows mit Word**. Unter Linux wird der
  Schritt übersprungen und vermerkt; die .docx sind trotzdem fertig.
* Die **Anthologien** („Tödliche Häppchen“) sind als ein Empfänger je
  Beitragender mit einem Buch der Menge 1 abgebildet, dessen Betrag der
  Anteil am Topf ist. Das rechnet richtig, ist aber ein Behelf — wenn der
  Topf sich ändert, müssen die Anteile von Hand nachgezogen werden.
* Der Bestand kennt **keine gleichzeitige Bearbeitung**. Wer ihn öffnet,
  hinterlässt `honorarbestand.sperre`; ein Zweiter wird gewarnt, mehr nicht.

## Bauen

```
pyinstaller honorar_abrechner/HonorarAbrechner.spec
```

Die Briefvorlage wird per `datas` mitgebündelt und zur Laufzeit über
`core._vorlagen_dir()` gefunden.
