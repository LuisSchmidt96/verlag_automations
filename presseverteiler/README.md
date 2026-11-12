# Presseverteiler

Schlägt für ein neues Buch die Pressekontakte vor, die es bekommen sollen.
Bisher sucht der Pressereferent sie je Buch von Hand aus der Access-Liste
heraus. Das dauert rund 30 Minuten und geht leicht schief.

**Stand: Prototyp.** Unter Linux geprüft, unter Windows noch nicht gebaut.

## Wie es vorschlägt

Beim Buch kreuzt man Regionen und Themen an. Das Programm sucht die 15
ähnlichsten früheren Bücher aus der Historie. Deren Listen stimmen ab:

* **Angekreuzt** ist, wer bei mindestens 30 % der ähnlichen Bücher dabei war
  (gewichtet nach Ähnlichkeit und Alter der Liste).
* **Grau darunter** stehen die übrigen Kontakte aus diesen Listen, sortiert nach
  Stimmen, dazu Kontakte aus der Gegend ohne Vorgeschichte.

Oben stehen die ähnlichen Bücher selbst. Wer die Liste eines davon übernimmt,
fängt mit genau dieser Liste an. Zwei Bücher derselben Reihe oder desselben Orts
bekamen bisher oft dieselbe Liste.

Jede gespeicherte Liste kommt in die Historie, ab dem nächsten Tag zählt sie für
neue Vorschläge. Der Vorschlag lernt also mit.

Warum nicht nach Zuständigkeit je Kontakt? Das war der erste Entwurf, und die
alten Listen haben ihn widerlegt: Er traf deutlich schlechter als eine feste
Liste der 60 häufigsten Kontakte. Die Befunde stehen in
[WIP_Presselisten-Auswertung.md](WIP_Presselisten-Auswertung.md).

**Was man erwarten darf** (blind geprüft an 89 Büchern ab 2025): Der Vorschlag
trifft etwa die Hälfte der echten Liste, unter den ersten 100 stehen zwei Drittel,
unter den ersten 150 drei Viertel. Er ersetzt das Prüfen also nicht. Mehr ist
kaum zu erwarten, denn zwei Ausgaben derselben Zeitschrift stimmen untereinander
noch weniger überein.

## Dateien

| Datei | was | wer pflegt |
|---|---|---|
| `_Presseverteiler/Access_Export.xlsx` | Anschriften; die ganze Kundentabelle wie für den Mailinglisten-Abgleich | Access, dorthin exportiert; das Tool liest nur |
| `_Presseverteiler/Pressehistorie.xlsx` | wer welches Buch bekommen hat | das Tool; in Excel lesbar und von Hand änderbar |
| `_Presseverteiler/Presseverteiler.xlsx` | Regionen und Themen je Kontakt, dazu die Schlagwortliste | das Tool; in Excel lesbar und von Hand änderbar |
| `<Titel>/` | je Buch ein Ordner mit den Listen für E-Mail und Post und dem Protokoll | das Tool |
| `config.json` neben der .exe | nur wenn ein Rechner einen anderen Export oder eine andere Ablage nimmt | das Tool, je Rechner |

Alles außer der `config.json` liegt auf dem NAS unter
`vr-intern/VR-Akquise/MK/Presse/Presselisten für Rezensionsversand`, bei den von
Hand angelegten Listen seit 2020. Anders als bei den übrigen Tools liegen die Mappen
nicht neben der .exe: Dort wären sie in der lokalen Kopie des Launchers
ungesichert, und jeder Rechner lernte für sich. Ein neuer Rechner braucht deshalb
nichts weiter, er klickt nur auf *Laden*. Ist das NAS nicht erreichbar, fragt das Tool beim Laden und
beim Ausgeben nach und weicht nie still aus. Wer dann einen anderen Ordner wählt,
etwa denselben über einen Laufwerksbuchstaben, dem merkt es ihn sich. Die Mappen
nennen Journalisten und sind deshalb nicht eingecheckt.

**Access-Export auffrischen:** In Access die *komplette Kundentabelle* (nicht die
Mailing-Abfrage) nach Excel exportieren und `_Presseverteiler/Access_Export.xlsx`
überschreiben. Wer im letzten Schritt des Assistenten *Exportschritte speichern*
anhakt, kann das später unter *Externe Daten → Gespeicherte Exporte* mit einem
Klick wiederholen. Neue Pressekontakte kennt das Tool erst nach dem nächsten
Export und *Laden*. Vorgeschlagen werden sie erst, wenn er sie einmal über die
Suche dazugenommen hat, weil der Vorschlag aus früheren Listen lernt.

Als **Presse** gilt in Access, wer bei `Presse/ZS` „Presse“ oder „ZS“ stehen
hat **oder** die Mailing-Kategorie `P` trägt. Die beiden Felder decken sich
nicht ganz, und wer nur eines prüft, verliert Kontakte. Sätze mit einem
Vermerk unter *Datensatz gelöscht* (verzogen, verstorben …) werden übergangen.
Die 24 VIPs, die in Listen vorkommen, sind fast alle auch als Presse markiert.

## Ablauf

1. **Laden** klicken. Den Access-Export, die Historie und die Mappe findet das Tool
   auf dem NAS.
2. **Je Buch:** Titel und ISBN eintragen, Regionen und Themen ankreuzen. Beide
   sind wichtig: Ohne Regionen oder ohne Themen trifft der Vorschlag spürbar
   schlechter.
3. **Prüfen:** Die Liste von oben nach unten durchsehen. Mit dem Haken an- oder
   abwählen. Die grauen Zeilen lohnen den Blick, die Hälfte der richtigen Liste
   steht dort oder fehlt ganz. Ein Klick auf eine Spaltenüberschrift sortiert,
   Texte nach dem Alphabet, *Quote* und *warum* nach der Stimmenzahl. Das Feld
   *suchen* filtert beim Tippen. Beides ändert nur die Anzeige: Wer ausgeblendet
   ist, bleibt angekreuzt.
4. **Ergänzen:** *Kontakte suchen und hinzufügen*. Die Suche kennt auch Regionen
   und Themen. Eine ganze Zeitungsfamilie geht mit „Südkurier“ suchen und *Alle
   Treffer hinzufügen*.
5. **Ausgeben:** *Fertig – alles ausgeben* erledigt alles auf einmal. Es legt
   in der Ablage einen Ordner mit dem Buchtitel an, legt die E-Mail-Adressen in die
   Zwischenablage (in Outlook unter **Bcc** einfügen), nimmt die Liste in die
   Historie auf und öffnet den Ordner. Stehen bei einem Kontakt mehrere Adressen,
   gehen alle mit, doppelte werden zusammengefasst. Am selben Tag ersetzt erneutes
   Ausgeben die Listen. Gibt es den Titel von einem anderen Tag, bekommt der neue
   Ordner das Datum, die alte Liste bleibt. Im Ordner liegen:
     * `Empfänger.xlsx`: alle Empfänger, *Buch* und *Abgewählt* sind das Protokoll.
     * `E-Mail.xlsx`: jede Adresse einmal, eine je Zeile.
     * `E-Mail-Adressen.txt`: dieselben Adressen, fertig für das Feld **Bcc**.
     * `Post.xlsx`: alle mit Anschrift, Quelle für Serienbrief und Etiketten. Die
       Spalte *auch per E-Mail* zeigt, wer nur per Post erreichbar ist.

Der Reiter **Kontakte** ist keine Pflicht. Dort stehen Regionen, Themen, Medienart
und Ressort je Kontakt als Suchhilfe, die Spalte *Listen* zeigt, wie oft jemand
gewählt wurde, und die Ansicht „nie gewählt“ die Kontakte ohne Vorgeschichte.

Wer eine Mappe in Excel bearbeitet, muss dort speichern, Excel schließen und im
Tool auf *Laden* klicken.

## Schlagworte

Die Startliste stammt aus den Kategorien des Webshops (Herkunft „Shop“). Dazu
kommen Ergänzungen, die der Shop nicht hat (Herkunft „Ergänzung“): „Geschichte“
als Oberbegriff, „Pfalz“ neben „Kurpfalz“, Stadtkreise, fehlende Landkreise und
die Gegenden, in denen Lokalpresse sitzt (Freiburg/Breisgau, Region Stuttgart,
Südpfalz, Rheinhessen, Unterfranken u. a.). „gehört zu“ macht Bücher vergleichbar:
Ein Buch über den Pfinzgau gilt auch als ähnlich zu Büchern über den Landkreis
Karlsruhe. Kreuzt man beim Buch eine Region an, sind ihre Teilgebiete mit
angekreuzt, verglichen wird aber nur das oberste Häkchen.

## Entwickeln

```
python -m presseverteiler.main
python -m presseverteiler.probelauf [mailing_list_updater/Access_Export.xlsx]
python -m presseverteiler.blindtest [--abstimmen] [--ab 2023] [--buecher]
python -m presseverteiler.historie_aufbauen [--ersetzen]
pyinstaller presseverteiler/Presseverteiler.spec
```

* `probelauf.py` prüft den Vorschlag an erfundenen Daten gegen feste Sollwerte.
* `blindtest.py` tut für jedes Buch so, als erschiene es gerade: Es kennt nur die
  Listen von davor und wird erst danach mit der echten Liste verglichen.
  `--abstimmen` schaut nur auf 2023–2024. Wer `NACHBARN`, `SCHWELLE` oder
  `ALTERUNG_JAHRE` in `core.py` ändert, stellt sie dort ein und misst dann blind.
* `historie_aufbauen.py` baut die Historie einmalig aus `training_data/` und den
  Buch-Etiketten in `einschaetzung/lernen/`. Beides ist nicht eingecheckt.
* `einschaetzung.py` und [einschaetzanleitung.md](einschaetzanleitung.md) sind die
  Agenten-Einschätzung der Kontakte. Ihr Ergebnis ist nur noch Suchhilfe.

## Offene Punkte

* Warum unterscheiden sich Listen derselben Reihe so stark: Absicht oder Zufall?
  Wenn Absicht, fehlt dem Vorschlag noch ein Kriterium.
* Was bedeuten `Direktversand` und die gelben Zeilen in den alten Listen?
* **Rezensionsexemplar oder nur Pressemitteilung:** Wer bekommt ein Buch, wer
  nur eine Mail? Das Tool unterscheidet das nicht.
* **Wohnort des Autors:** Lokalzeitungen dort berichten oft auch über ein Buch,
  das woanders spielt.
* **Wünsche nach Autor:** Einzelne Kontakte möchten laut Bemerkung alle Bücher
  eines Autors (Alpenverein Tübingen, Württemberg Reporter).
* Regionen und Themen aus dem Shop vorbelegen (per ISBN).
* Sollen die 802 nie gewählten Kontakte aus der Suche verschwinden?
