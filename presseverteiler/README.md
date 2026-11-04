# Presseverteiler

Wählt für ein neues Buch die Pressekontakte aus, die es bekommen sollen.
Bisher sucht der Pressereferent sie je Buch von Hand aus der Access-Liste
heraus. Das dauert rund 30 Minuten und geht leicht schief.

**Stand: Prototyp.** Die Schlagworte und die Auswahlregel sind ein
begründeter Vorschlag. Sie werden erst mit dem Pressereferenten festgezurrt
(siehe *Offene Punkte*).

## Wie es auswählt

Für jeden Pressekontakt ist festgehalten, wofür er zuständig ist: **Regionen**,
**Themen** und gegebenenfalls **immer**. Für das Buch kreuzt man dieselben
Schlagworte an. Ein Kontakt bekommt das Buch, wenn **jede ausgefüllte Achse
passt**. Eine leere Achse zählt als „egal“.

| Kontakt eingetragen mit | bekommt |
|---|---|
| nur Regionen (Lokalzeitung) | alles aus diesen Regionen |
| nur Themen (Eisenbahn-Fachblatt) | alles zu diesen Themen, egal woher |
| Regionen und Themen | nur, was in beiden passt |
| immer | jedes Buch |
| nichts | nie etwas |

Diese Verknüpfung (UND zwischen den Achsen, ODER innerhalb einer Achse,
leer = egal) kann der AutoFilter von Excel nicht. Deshalb reicht eine reine
Excel-Liste nicht aus.

**Ober- und Teilgebiete.** Im Blatt *Schlagworte* steht, wozu ein
Schlagwort gehört. Daraus folgt zweierlei:

* Wer für ein Gebiet oder Thema zuständig ist, bekommt auch alles darunter.
  Ein Buch über den Pfinzgau geht deshalb auch an alle, die für den
  Landkreis Karlsruhe, die Region Karlsruhe, Baden oder Baden-Württemberg
  eingetragen sind. Ein Buch zur Revolution 1848/49 geht an die
  Geschichtsredaktionen, aber nicht an die für Kirchengeschichte.
* Kreuzt man beim Buch eine **Region** an, sind ihre Teilgebiete gleich mit
  angekreuzt: Landkreis Karlsruhe nimmt Kraichgau und Pfinzgau mit, denn das
  Buch betrifft auch Bretten. Einzelne Teilgebiete lassen sich wieder
  abwählen. Bei **Themen** gilt das nicht: Ein Geschichtsbuch ist nicht
  automatisch Kirchengeschichte.

Ein Teilgebiet, das nur mitgenommen wurde, zieht seinerseits nichts nach
oben. Der Kraichgau liegt auch im Landkreis Heilbronn. Ein Buch über den
Landkreis Karlsruhe geht trotzdem nicht an die Redaktionen für Heilbronn
oder Württemberg.

Die Startliste der Schlagworte stammt aus den Kategorien des Webshops
(Herkunft „Shop“). Dazu kommen Ergänzungen, die die Presse braucht und der
Shop nicht hat (Herkunft „Ergänzung“): „Geschichte“ als Oberbegriff,
„Pfalz“ neben „Kurpfalz“, Stadtkreise und fehlende Landkreise. Mit der
ersten Einschätzung (November 2026) kamen die Gegenden dazu, in denen
Lokalpresse sitzt, die bisher nur unter „Baden“, „Württemberg“ oder „Pfalz“
gepasst hätte: Freiburg/Breisgau, Region Stuttgart, Tübingen, Reutlingen,
Oberschwaben, Ulm, Südpfalz, Westpfalz, Rheinhessen, Unterfranken, Saarland
u. a. Dazu kamen die Themen „Sport“ und „Architektur / Denkmalpflege“, die
häufigsten Access-Fachgebiete ohne Shop-Kategorie.
*Ortsgeschichte* hängt bewusst nicht unter *Geschichte*: Sonst bekäme jede
Geschichtszeitschrift jede Ortschronik.

## Dateien

| Datei | was | wer pflegt |
|---|---|---|
| Access-Export (xlsx) | Anschriften; die ganze Kundentabelle wie für den Mailinglisten-Abgleich | Access, das Tool liest nur |
| `Presseverteiler.xlsx` | wer wofür zuständig ist, dazu die Schlagwortliste | das Tool; in Excel lesbar und notfalls von Hand änderbar |
| `Presseversand/<Datum>_<ISBN>_<Titel>.xlsx` | je Buch die Empfänger (taugt als Serienbrief-Quelle) und ein Protokoll | das Tool |
| `einschaetzung/` (nur im Repo-Ordner) | Pakete und Ergebnisse der ersten Einschätzung, Sicherung der Mappe davor | `einschaetzung.py` |

Alles liegt neben der .exe. Die Mappe enthält Namen und Zuständigkeiten von
Journalisten und ist deshalb nicht eingecheckt.

Als **Presse** gilt in Access, wer bei `Presse/ZS` „Presse“ oder „ZS“ stehen
hat **oder** die Mailing-Kategorie `P` trägt. Die beiden Felder decken sich
nicht ganz, und wer nur eines prüft, verliert Kontakte. Sätze mit einem
Vermerk unter *Datensatz gelöscht* (verzogen, verstorben …) werden übergangen.

## Ablauf

1. **Einmalig:** Access-Export wählen und auf *Laden* klicken. Beim ersten
   Mal legt das Tool die Mappe an. Die alten Access-Felder `Region 1–3` und
   `Fachgebiet 1–3` werden, soweit eindeutig, als **ungeprüfte Vorschläge**
   übernommen (im Aug. 2026: 923 von 1.903 Kontakten). Was nicht zu übersetzen
   war, steht im Blatt *Altwerte aus Access*.
2. **Einordnen:** Im Reiter *Einordnen* zuerst die Ansicht „ungeprüft und
   unsicher“ wählen (rot), danach „ungeprüft“ (braun), und durchklicken.
   Rechts stehen:
   * die Begründung des Vorschlags („Lokalredaktion der BNN in Bretten“),
   * die alten Access-Werte und die Bemerkungen. Dort stehen oft Regeln, die
     es sonst nirgends gibt.

   Regionen und Themen stehen nebeneinander, darunter Medienart und Ressort.
   *Übernehmen* speichert sofort, markiert den Kontakt als geprüft und springt
   zum nächsten. Ein Kontakt ohne Schlagworte bekommt nie etwas. Auch das ist
   ein Vorschlag, den man bestätigt.
3. **Je Buch:** Titel und ISBN eintragen, Regionen und Themen ankreuzen.
   Die Liste prüfen: braun bedeutet *noch ungeprüft*, blau bedeutet *von Hand
   hinzugefügt*. Mit dem Haken wählt man für dieses Buch ab. Dann:
   * *E-Mail-Adressen kopieren* und in Outlook unter **Bcc** einfügen.
     Stehen bei einem Kontakt mehrere Adressen, gehen alle mit. Doppelte
     werden zusammengefasst.
   * *Liste speichern* legt die Datei unter `Presseversand/` ab. Das
     Blatt *Empfänger* dient als Quelle für Serienbrief oder Etiketten,
     die Blätter *Buch* und *Abgewählt* sind das Protokoll.

Wer die Mappe in Excel bearbeitet, muss dort speichern, Excel schließen und
im Tool auf *Laden* klicken. Bemerkt das Tool beim Speichern, dass die Mappe
seit dem Laden von außen geändert wurde, speichert es nicht, statt fremde
Änderungen zu überschreiben.

## Erste Einschätzung

Die Access-Felder ergaben nur für knapp die Hälfte der Kontakte einen
Vorschlag, und der war grob. Deshalb wurde jeder Kontakt einmal von Agenten
eingeschätzt, wie bei der Jubiläums-Recherche. Die Fragen dabei: Was ist das
für ein Medium, wo erscheint es, worüber berichtet es? Das Ergebnis steht als
**ungeprüfter** Vorschlag in der Mappe, mit Begründung, Medienart, Ressort
und dem Vermerk „unsicher“.

```
python -m presseverteiler.einschaetzung pakete       # einschaetzung/paket_NN.md
#   je Paket ein Agent nach einschaetzanleitung.md → ergebnis_NN.jsonl
python -m presseverteiler.einschaetzung pruefen      # vollständig? Schlagworte gültig?
python -m presseverteiler.einschaetzung uebernehmen  # in die Mappe, mit Sicherung
```

`uebernehmen` rührt geprüfte Einträge nicht an, gleicht das Blatt
*Schlagworte* an die Startliste im Code an und legt vorher eine Sicherung der
Mappe in `einschaetzung/` ab. Die Regeln, nach denen eingeschätzt wird, stehen
in [einschaetzanleitung.md](einschaetzanleitung.md): engstes passendes
Gebiet, Themen nur bei echter Beschränkung, „immer“ nur auf ausdrücklichen
Wunsch. Die Anleitung sagt auch, dass die Access-Region oft eine alte
Selektion ist („Odenwald“ beim Amtsblatt Mannheim).

**Medienart und Ressort** werden erfasst, gehen aber noch nicht in die
Auswahl ein. Sie werden erst dann ein Filter, wenn sich zeigt, dass der
Pressereferent danach auswählt.

## Entwickeln

```
python -m presseverteiler.main
python -m presseverteiler.probelauf [mailing_list_updater/Access_Export.xlsx]
pyinstaller presseverteiler/Presseverteiler.spec
```

`probelauf.py` prüft die Auswahlregel gegen feste Sollwerte. Mit einem
Access-Export legt es den Verteiler zusätzlich in einem temporären Ordner
neu an und zählt Startwerte, Altwerte und Empfänger für drei Testbücher.

## Offene Punkte

* Aus welcher Liste wählt der Pressereferent heute: aus der Access-Tabelle
  oder aus einer eigenen Abfrage bzw. Datei?
* Reichen die Shop-Kategorien? Kritisch sind die Ortsgeschichte (im Shop
  alphabetisch und ohne Region) und fehlende Landkreise.
* Wonach entscheidet er außer Region und Thema noch: Medium, Ressort,
  Wünsche einzelner Redakteure?
* Was braucht er nach der Auswahl: Mail, Serienbrief, Etiketten,
  Rezensionsexemplare?
* Wohin, wenn bei einem Kontakt mehrere Adressen stehen: alle (so ist es
  jetzt) oder nur die erste?
* Wählt er auch nach **Medienart oder Ressort** aus? Zum Beispiel bei der
  RNZ nur die Kultur, Rundfunk nur bei Autoreninterviews. Beides ist erfasst
  und müsste nur noch als Filter an die Auswahl.
* **Rezensionsexemplar oder nur Pressemitteilung:** Wer bekommt ein Buch,
  wer nur eine Mail?
* **Wohnort des Autors:** Soll ihn das Buch als zweites Regionsfeld
  bekommen? Darüber würden nur Lokalmedien ausgewählt, denn die Zeitung am
  Wohnort berichtet auch über ein Buch, das woanders spielt.
* **Gemeinde statt Kreis** für Amts- und Mitteilungsblätter: Die Ortschronik
  Östringen gehört ins Amtsblatt Östringen, nicht in alle Amtsblätter des
  Landkreises.
* **Vorlauf:** Jahrbücher und Vierteljahreshefte brauchen die Meldung Monate
  vorher. Muss das Programm daran erinnern?
* **Wünsche nach Autor:** Einzelne Kontakte möchten laut Bemerkung alle
  Bücher eines bestimmten Autors (Alpenverein Tübingen, Württemberg
  Reporter). Das lässt sich mit Region und Thema nicht abbilden.
* **Mehrere Leute je Redaktion:** Mit der Einschätzung bekäme die
  Ortsgeschichte Östringen 227 Empfänger, darunter 33 beim SWR und
  15 bei der RNZ. Schreibt er wirklich alle an, oder je Medium nur das
  passende Ressort?
* **Später:** aus alten Selektionen lernen (Shop-Kategorien per ISBN
  gegen die damaligen Empfänger abgleichen) und die Schlagworte eines Buchs
  aus dem Shop vorbelegen. Sobald es Versandprotokolle gibt, lässt sich auch
  festhalten, wer tatsächlich berichtet hat.
