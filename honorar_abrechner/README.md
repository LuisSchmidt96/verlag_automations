# Honorar-Abrechner

Rechnet einmal im Jahr die Autorenhonorare ab und erzeugt Briefe,
Zahlungsliste und die Meldung an die Künstlersozialkasse.

Ersetzt die Excel-Mappe `Honorare_<Jahr>.xlsx`, in der die Geschäftslogik
bisher in von Hand getippten Zellformeln steckte.

## Ablauf

1. **Einmalig:** *Stammdaten → Altmappe importieren* — liest die alte
   `Honorare_2025.xlsx` mit allen Blättern ein, inklusive der Jahreshistorie
   ab 2003. Danach ist die Altmappe Archiv.
2. **Jeden Januar:**
   1. *Jahreserfassung* — `verk. Ex.` und `Eigenkauf` je Buch eintragen.
      Doppelklick oder Eingabetaste öffnet das Feld, Eingabetaste übernimmt
      und springt eine Zeile weiter.
   2. *Durchlauf → Jetzt durchrechnen* — zeigt je Empfänger den Betrag.
      Grau = kein Brief, rot = etwas nachzusehen.
   3. *Ausgaben* — Briefe, dann (unter Windows) PDF, dann die Listen.
   4. *Stammdaten → Bestand speichern.*

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
geändert werden. Fünf Blätter:

| Blatt | Inhalt |
|---|---|
| `Hinweise` | die Regeln im Klartext — steht absichtlich an erster Stelle |
| `Empfänger` | Anschrift, IBAN, Aktenzeichen, E-Mail |
| `Bücher` | Titel, ISBN, Vergütungsart und die ganze Kondition |
| `Jahreswerte` | eine Zeile je Buch **und Jahr** (Langformat, nicht eine Spalte je Jahr) |
| `Staffeln` | die Stufen der Staffelverträge |

Verknüpft wird über die Kennungen (`E0042`, `B0137`). Die dürfen **nicht**
geändert werden — sonst verliert ein Buch seinen Autor.

Zwei Dinge, die man wissen muss:

* **Leer ist nicht null.** In `Jahreswerte` heißt eine leere Zelle bei
  `verkauft` „noch nicht erfasst“. Hat sich ein Buch wirklich nicht verkauft,
  gehört dort eine `0` hinein.
* **Die Datei schließen, bevor das Werkzeug speichert.** Solange sie in Excel
  offen ist, kann nicht geschrieben werden; das Werkzeug sagt das dann auch,
  statt abzustürzen. Vor jedem Speichern entsteht `Honorarbestand.bak.xlsx`.

## Ausgaben

Alles nach `honorar_output/<Jahr>/` neben der .exe:

| Datei | Inhalt |
|---|---|
| `briefe/<Name>_<Art>-<Jahr>.docx` | je Empfänger ein Brief, Seite 1 hoch, Seite 2 quer |
| `briefe/….pdf` | dieselben Briefe als PDF — **nur unter Windows mit Word** |
| `Zahlungsliste_<Jahr>.xlsx` | für die Überweisungen, Aufbau wie bisher |
| `Kuenstlersozialkasse_<Jahr>.xlsx` | Buchungsformat, zwei Abschnitte (7 % / 19 %) |
| `protokoll.xlsx` | wer bekommt einen Brief, wer nicht, und warum |
| `import_protokoll.xlsx` | nur beim Import: jede Zeile und jede Unklarheit |

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

## Gegenprobe

*Durchlauf → Gegen die Altmappe prüfen* rechnet jede Zeile und jede
Auszahlung des Jahres gegen die alte Mappe nach. Stand der letzten Prüfung
für 2025: **alle 275 Zeilen und alle 137 Auszahlungen stimmen auf den Cent**,
bis auf den einen oben genannten Fall mit der kaputten Zelle.

Prüfwerte zum Nachschlagen: Dieter Buck 2.368,98 € brutto / 2.214,01 € netto
(entspricht dem Muster-PDF und dem KSK-Blatt).

Die Gesamtsumme liegt mit 19.923,02 € um 911,52 € über der Altmappe. Das ist
kein Rechenfehler: 738,82 € stammen aus dem Blatt „Zahlung ab XX Ex.“ und
166,09 € aus LUBW und den Anthologien — Geld, das bisher auf getrennten
Blättern stand und jetzt beim selben Empfänger im selben Brief landet. Sieben
Empfänger sind davon betroffen; das Werkzeug weist es bei ihnen aus.

## Offene Punkte für den Verlag

1. **Staffel — ein Satz oder tranchenweise?** Wird beim Überschreiten der
   Grenze der höhere Satz nur auf die darüberliegenden Exemplare angewandt
   oder auf die ganze Jahresmenge? Umgesetzt ist derzeit „ein Satz je Jahr“.
2. **Die kumulierten Stände.** In den Notizen steht „Stand 2023: 1460 Ex.“ —
   für 2024 und 2025 fehlt die Fortschreibung bei den Büchern, die nicht im
   Historienblatt stehen.
3. **Die echte Word-Vorlage** mit dem Briefkopf.
4. **Die Summenkette `AA245`** fasst Schenk und Möller zusammen. Fehler?
5. **Vertauschte Namen.** In der Altmappe steht bei „Mühlacker+“ der Nachname
   im Vornamensfeld; der Brief heißt deshalb `Sandra, Schuster_…` statt
   `Schuster, Sandra_…`. Das Werkzeug übernimmt die Felder so, wie sie
   dastehen — wer sie tauscht, tut es im Bestand.
6. **Zwei Empfänger ohne Bankverbindung**, zwei mit **mehreren
   Vergütungsarten** in einem Brief — wie soll damit umgegangen werden?
7. **Unterzeichnerin.** Die Musterbriefe zeichnet Silke Freitag; im
   `pi_bi_generator` wurde der Absender zwischenzeitlich geändert.
8. **Lexware-Artikelexport.** Die Zahlen gibt es — nur noch nicht als Datei.
   Einzelheiten unter „Der Weg zu den Stückzahlen".

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
