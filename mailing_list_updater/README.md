# Mailinglisten-Abgleich

Gleicht den Access-Adressstamm einmal im Jahr gegen die Lexware-Exporte ab:
Wer gekauft hat, bekommt ein aktuelles `Bestelldatum`, echte Neukunden werden
angelegt. Bisher war das Handarbeit, weil die beiden Systeme **keinen
gemeinsamen Schlüssel** haben — in Access steht nirgends eine Lexware-Kd.-Nr.

Das Tool schreibt **nicht** in die Access-Datei. Es erzeugt zwei Excel-Dateien,
die in Access importiert werden. Wer den Brief bekommt, entscheidet weiterhin
der Filter in Access; daran rührt das Tool nicht.

## Einmalig einrichten

1. In der Access-Kundentabelle ein Feld **`Lexware-Kd-Nr`** anlegen (Text, 20).

   Das ist der eigentliche Ausweg aus der Handarbeit: Das Tool trägt dort bei
   jedem bestätigten Treffer die Kd.-Nr. ein. Ab dem Folgejahr ist der Abgleich
   exakt statt geraten, und die Rate-Logik greift nur noch bei echten Neukunden.

2. Die UPDATE-Abfrage in Access anlegen — den fertigen SQL-Text erzeugt das
   Tool bei jedem Lauf als `access_import.txt`.

## Die vier Exporte

| Datei | woher | worauf achten |
|---|---|---|
| **Access-Kundentabelle** | Access, Export nach xlsx | die **komplette Tabelle**, nicht die Mailing-Abfrage |
| **Lexware-Kunden** | Verwaltung → Kunden, im Register auf **Alle**, dann Export → *Ansicht nach MS® Excel exportieren* | am besten ohne Einschränkung auf neue Kunden |
| **Lexware-Aufträge, laufendes Jahr** | Verwaltung → Aufträge, im Fenster *Auswahlkriterien* den Zeitraum setzen, dann Export | s. u. |
| **Lexware-Aufträge, Vorjahr** | dasselbe für das Vorjahr | **nicht vergessen** |

### Warum das Vorjahr dazugehört

Der Kundenexport reicht bis zum letzten Mailing zurück, ein Auftragsexport
umfasst aber nur ein Kalenderjahr. Wer im Spätjahr angelegt wurde, hätte sonst
kein Bestelljahr und sähe aus wie ein Kunde, der nie bestellt hat. Faustregel:
**die Aufträge müssen mindestens so weit zurückreichen wie der Kundenexport.**
Zu viel geladen schadet nie — das Bestelldatum wird nur angehoben, nie gesenkt.

Das Tool erkennt diese Lücke selbst und sagt, ab welcher Kd.-Nr. sie beginnt.

### Spalten in der Auftragsliste

Unter *Ansicht → Listeneinstellungen* lassen sich Spalten zuschalten. Nützlich
sind `Name`, `Vorname`, `Firma`, `Plz, Ort`, `Straße`, `Hausnummer` — damit
lassen sich Bestandskunden zuordnen, die nicht im Kundenexport stehen. Eine
E-Mail bietet die Auftragsliste nicht an; wer die braucht, muss den
Kundenvollexport ziehen.

Pflicht sind nur `Datum`, `Art`, `Belegnr.`, `Kd.-Nr.` und `Gesamt`. Reihenfolge
und Zusatzspalten sind egal, die Kopfzeile wird gesucht.

## Ablauf

1. Tool starten, die Dateien wählen (Aufträge dürfen mehrere sein), auf
   **Abgleichen**.
2. Warnungen oben lesen. Sie stehen da nicht zur Zierde — beide bisher
   eingebauten Prüfungen haben in der Erprobung echte Lücken gefunden.
3. Die fünf Reiter durchgehen:

   | Reiter | was zu tun ist |
   |---|---|
   | **Neu anlegen** | Liste sichten, Auffälligkeiten sind orange hervorgehoben |
   | **Aktualisieren** | je Feld ankreuzen, was aus Lexware übernommen wird; das Bestelldatum wird immer gesetzt. Über *Stattdessen* lässt sich der Fall auch als Zweitadresse neu anlegen oder ganz übergehen |
   | **Unklar** | je Fall entscheiden: *diesen aktualisieren* · *neu anlegen (Zweitadresse)* · *Ignorieren*. Danach springt die Ansicht in den Reiter, in dem der Fall gelandet ist |
   | ↳ Spalte *Folge* | **wichtigste Spalte des Reiters.** Steht dort „bekommt Post ohnehin", ist die Entscheidung folgenlos — der Empfänger ist über seine Kategorie dauerhaft im Mailing. Solche Fälle stehen unten und dürfen liegen bleiben. Im Probelauf: 262 von 344 folgenlos, **82 entscheiden wirklich etwas** |
   | **Ohne Auftrag** | sollte fast leer sein; ist er voll, fehlt eine Aufträge-Datei |
   | **Ignoriert** | was du ausdrücklich beiseitegelegt hast — damit du siehst, was du durch hast. Die Vorbelegung der unklaren Fälle zählt **nicht** dazu, sonst wäre der Reiter gleich nach dem Abgleich voll |

   Die Listen sind so vorsortiert, dass **die fragwürdigen Fälle oben stehen**:
   im Reiter „Aktualisieren" die mit Abweichungen und schwachem Treffer, in
   „Unklar" die gesperrten und die mit mehreren gleich guten Kandidaten, in
   „Neu anlegen" die auffälligen. Über die Spaltenköpfe lässt sich jederzeit
   anders sortieren.

   Unter jeder Liste steht der Satz so, wie er hinterher in Access aussieht —
   **jedes Feld änderbar, auch die leeren.** Gerade die leeren sind oft die
   wichtigen: steht in Lexware die Straße im PLZ-Feld, sind `Straße` und
   `Hausnummer` leer und müssen hier gefüllt werden. Oben stehen die Felder,
   die man tatsächlich anfasst, darunter der Rest der 50 Access-Spalten. Eine
   Handkorrektur hat Vorrang vor den Übernahme-Häkchen.

   **Die Liste oben zeigt immer das Ergebnis, nicht die Rohdaten.** Ändert man
   unten einen Namen oder legt ein Häkchen um, ändert sich die Zeile sofort
   mit; eine behobene Auffälligkeit verliert ihre Färbung. Die Spalte
   *wird in Access geändert* nennt genau die Felder, die der Import anfasst —
   nach Häkchen und Handkorrekturen. Steht dort „nur Bestelldatum", bleibt die
   Anschrift, wie sie ist.

   **Jede Entscheidung wird vorher gezeigt.** *Diesen aktualisieren*,
   *neu anlegen* und *zusammenlegen* öffnen ein Fenster mit dem fertigen Satz —
   änderbar, mit der Angabe, welche Felder sich ändern. Danach bleibt man im
   Reiter „Unklar" stehen und die Auswahl rückt auf den nächsten Fall vor;
   man muss also nicht in einen anderen Reiter springen und die Stelle
   wiederfinden.

   Beim Aktualisieren lassen sich dort auch `VIP/W/K/X`, `Autor`, `Presse/ZS`
   und `Titel 2` setzen — wer beim Durchsehen merkt, dass jemand als `K` statt
   `BUHA` geführt wird, kann das gleich richten. Aus Lexware **vorgeschlagen**
   werden diese Felder nie: die dortige Kundengruppe ist gröber als die über
   Jahre gepflegte Einordnung in Access.

   **Jede Entscheidung ist umkehrbar.** Ein als Zweitadresse angelegter Fall
   zeigt im Reiter „Neu anlegen" weiter seine Access-Kandidaten und lässt sich
   mit einem Klick zurück auf *aktualisieren* stellen — ohne den Abgleich neu
   zu fahren.

   Eine weit abweichende Anschrift ist übrigens oft kein Umzug, sondern die
   zweite Adresse desselben Menschen: privat in Access, dienstlich in Lexware.
   Beim Anlegen einer Zweitadresse bleibt der vorhandene Access-Satz
   **unverändert** — auch sein Bestelldatum. Soll er weiter Post bekommen,
   muss er selbst noch einen aktuellen Kauf vorweisen oder ein Merkmal tragen.

   **Mehrere Access-Sätze zusammenlegen.** Im Reiter „Unklar" lassen sich
   Kandidaten ankreuzen und über *Ausgewählte zusammenlegen …* zu einem Satz
   vereinen — für Dubletten, die sich in einer Kleinigkeit unterscheiden und
   deshalb nicht automatisch erkannt werden („Ellwangen (Jagst)" gegen
   „Ellwangen"). Der gehaltvollste Satz bleibt, E-Mail und Telefon behalten
   **alle** Werte mit Komma verbunden, alles ist im Fenster noch änderbar.

   In `kunden_komplett.xlsx` sind die aufgelösten Sätze entfernt — nur über
   die Gesamttabelle ist ein Zusammenlegen überhaupt möglich, ein
   Anfüge-Import kann nichts löschen. Wer den Weg über die Einzeldateien geht,
   findet die zu löschenden IDs in `zusammenlegen.xlsx`.

   **Abweichende Adressfelder sind vorgehakt** (`UEBERNAHME_VORGABE` in
   `core.py`): Lexware ist das System, in dem täglich gearbeitet wird, und eine
   Anschrift, unter der gerade geliefert wurde, ist aktueller als der
   Access-Stand — aber **nur, wenn nichts verlorengeht**. Wo
   die Übernahme Angaben löschen würde, steht ein ⚠ und das Häkchen fehlt;
   anhaken kann man es trotzdem. Als Verlust gilt, wenn der Lexware-Wert im
   Access-Wert schon steckt oder ihn abkürzt:

   ```
   Vaihingen/Enz                    ->  Vaihingen
   Württembergische Landesbibliothek->  Württemb. Landesbibliothek
   zzz_keine Werbeanrufe! 07224 …   ->  07224 40133
   ```

   Ein bloßer Längenvergleich reichte dafür nicht: „Mannheim" statt
   „Neckargemünd" ist kürzer, aber ein Umzug und kein Verlust. Ebenso wird
   eine Sammelanrede nicht über eine persönliche vorgeschlagen — kennt Access
   „Frau" und Margot Zander, ist „Damen und Herren" aus Lexware die ärmere
   Angabe.

   Keine Abweichung sind unterschiedliche **Schreibweisen**: `Kirchstraße`
   und `Kirchstr.` gelten als dasselbe, ebenso `Untere Hauptstr. 54` gegen
   `Untere Hauptstr.` + Hausnummer `54` — Lexware führt die Hausnummer mal
   getrennt, mal am Ende der Straße. Mannheimer Quadrate wie `C 5` bleiben
   dabei unangetastet.

   Im Reiter **Aktualisieren** werden genau die Felder gezeigt, welche die
   UPDATE-Abfrage schreibt: was dort steht, steht hinterher in Access, nicht
   mehr und nicht weniger. Die `ID` ist schreibgeschützt — sie ist der
   Schlüssel, über den verbunden wird.

4. **Dateien schreiben**. Alles landet in `mailing_output/<Jahr>/`.
5. Vor dem Schreiben zeigt das Tool **alle Änderungen am Stück**: sortierbar,
   scrollbar, eine Zeile je Satz mit dem, was geschieht. Die Reiter zeigen
   jeweils nur einen Ausschnitt — hier steht es zum ersten Mal beisammen.
6. Nach Access übertragen — wie, steht in `access_import.txt`.
7. `pflege.txt` an die Kolleginnen geben, die Lexware pflegen.

Entscheidungen werden laufend in `entscheidungen.json` gesichert und beim
nächsten Start zurückgeladen. Die Durchsicht von über tausend Fällen überlebt
so eine Mittagspause.

## Was herauskommt

| Datei | Inhalt |
|---|---|
| **`kunden_komplett.xlsx`** | **die ganze Kundentabelle, fertig aktualisiert und ergänzt — der einfache Weg** |
| `neu_anlegen.xlsx` | nur die neuen Sätze, Access-Spalten ohne `ID` |
| `aktualisieren.xlsx` | nur die geänderten, `ID` plus alle Felder mit dem Endwert |
| `protokoll.xlsx` | jede Entscheidung mit Punktzahl und Begründung |
| `zuordnung.xlsx` | Kd.-Nr. ↔ Access-ID, Sicherung neben dem Access-Feld |
| `zusammenlegen.xlsx` | welche Access-Sätze in welchen aufgehen — in `kunden_komplett.xlsx` schon entfernt, beim zweiten Weg von Hand zu löschen |
| `access_import.txt` | beide Wege nach Access, Schritt für Schritt |
| **`pflege.txt`** | **was in Lexware und Access aufzuräumen wäre** — Dubletten, Hausnummern im Straßenfeld, fehlende PLZ. Ändert nichts, zeigt nur; behoben wird es in der Quelle, sonst steht es nächstes Jahr wieder da |

### Der einfache Weg: eine Jahresdatei

Access-Datei kopieren und die Kopie `Adressen_<Jahr>.accdb` nennen → in der
Kopie die Tabelle leeren → `kunden_komplett.xlsx` **anfügen** → komprimieren
und reparieren. Kein SQL, keine Importtabelle. Die bisherige Datei wird nicht
angefasst und bleibt als Stand des Vorjahres liegen; geht etwas schief, löscht
man die Kopie und fängt neu an.

> **Beim Import „an die Tabelle anfügen" wählen, nicht „in eine neue Tabelle
> importieren".** Sonst rät Access die Feldtypen — und aus der PLZ `04103`
> wird die Zahl 4103. Betroffen wären 225 Postleitzahlen und 9 913
> Telefonnummern, und die `ID` verlöre ihre AutoWert-Eigenschaft.

Wer lieber gar nichts löscht, nimmt den zweiten Weg (`aktualisieren.xlsx` +
gespeicherte UPDATE-Abfrage + `neu_anlegen.xlsx`); beide stehen in
`access_import.txt`.

`kunden_komplett.xlsx` wird **nur** geschrieben, wenn die geladene Access-Datei
nachweislich die volle Tabelle ist. Ist es die Mailing-Abfrage, würde die Datei
beim Import alles löschen, was der Filter nicht zeigt — dann bleibt sie aus und
die Anleitung sagt warum.

`aktualisieren.xlsx` trägt in jedem Feld schon den Endwert (unveränderte Felder
den bisherigen Access-Wert). Dadurch bleibt die UPDATE-Abfrage stumpf und die
Datei lesbar: man sieht der Zeile an, wie der Datensatz hinterher aussieht.

## Wie der Abgleich arbeitet

Drei Stufen, in dieser Reihenfolge:

1. **Kd.-Nr.** — steht sie in Access, ist der Treffer exakt. Ab dem zweiten
   Jahr der Regelfall.
2. **Blocken** — nur wer PLZ, E-Mail oder Straße teilt, wird überhaupt
   verglichen.
3. **Punkten** — gewichtete Ähnlichkeit über Name, Vorname, Institution,
   Straße, Hausnummer, PLZ und E-Mail. Die Schwellen stehen als Konstanten am
   Kopf von `core.py`.

**Gleichlautende Access-Sätze werden zusammengefasst.** Stehen zwei Datensätze
in Name, Anschrift und E-Mail vollständig überein, ist das keine Auswahl,
sondern eine Dublette in Access — vorgelegt würde sie nur Zeit kosten. Das
Tool behält den gehaltvolleren Satz (die meisten gefüllten Felder, dann der
zuletzt geprüfte) und vermerkt die anderen. In der Oberfläche steht dann
`ID 33817 +1 gleiche`, im Protokoll die verworfenen IDs unter
*Access-Dubletten* — damit sie sich in Access aufräumen lassen.

Zwei Sicherungen, die aus der Erprobung stammen:

* Ein **abweichender Vorname** bei gleicher Adresse wird nie automatisch
  zusammengeführt — das ist fast immer derselbe Haushalt und eine andere
  Person. Eine gemeinsame Familien-E-Mail überstimmt das nicht.
* Sätze mit Eintrag in `Datensatz gelöscht` oder `prüfen` (verzogen,
  verstorben) landen immer zur Handprüfung. Einen Verstorbenen zurück ins
  Mailing zu holen ist der Fehler, den niemand sehen will.

Als Kauf zählt nur eine Rechnung (`RG`/`SR`) mit Umsatz > 0. Rezensions- und
Autorenexemplare gehen auf Lieferschein raus; die Betroffenen werden angelegt,
bekommen aber kein Bestelldatum — ihr Brief kommt über das Merkmal.

### Für wen das Bestelldatum überhaupt zählt

Nur für Kategorie **`K`**. Alle anderen sind über ihre Kategorie dauerhaft im
Mailing — abgelesen am Export selbst, denn der ist die Mailing-Abfrage:

| `VIP/W/K/X` | im Filter ohne aktuellen Kauf | mit |
|---|---|---|
| `K` | 5 278 | 2 766 |
| `VIP` | 4 005 | 143 |
| `BUHA` | 2 411 | 83 |
| `P` | 1 783 | 79 |
| `Wiss` | 832 | 32 |
| `V`, `VHS`, `MS`, `SH` | 869 | 26 |

Bei `K` halten sich beide Gruppen die Waage, überall sonst stehen 91–100 %
ohne aktuellen Kauf im Filter. Buchhandel, Presse, Wissenschaft, Vereine und
VIP bekommen also Post, gleichgültig wann sie zuletzt gekauft haben. Genau
deshalb weist die Spalte *Folge* im Reiter „Unklar" solche Fälle als
folgenlos aus und sortiert sie nach unten (`KATEGORIE_KAUFABHAENGIG` in
`core.py`).

### Zwei Kundennummern auf einem Access-Satz

Lexware führt manche Leute doppelt — beim zweiten Kauf neu angelegt statt
wiedergefunden. Beide Nummern zeigen dann auf denselben Access-Satz; das Tool
legt solche Fälle in **Unklar** vor.

Entscheidet man für beide *aktualisieren*, stünde dieselbe `ID` zweimal in
`aktualisieren.xlsx` — und Access' UPDATE nähme davon irgendeine. Deshalb
werden solche Zeilen beim Schreiben **zusammengeführt**:

* das höhere Bestelljahr gewinnt,
* **beide Kundennummern** kommen ins Feld (`119506, 202515`),
* widersprechen sich Adressfelder, setzt sich die jüngere Bestellung durch.

Jede Zusammenführung wird im Verlauf genannt. Ab dem Folgejahr findet die
Kd.-Nr.-Suche den Satz über **jede** der eingetragenen Nummern, der Fall ist
also einmalig.

## Entwicklung

Kopfloser Probelauf, ohne GUI, gegen die echten Dateien:

```
python -m mailing_list_updater.probelauf \
    mailing_list_updater/Access_Export.xlsx \
    mailing_list_updater/Leware-Kunden_2026_13.08.2026.xlsx \
    mailing_list_updater/Aufträge_2025_komplett.xlsx \
    mailing_list_updater/Aufträge_2026_komplett.xlsx
```

Gibt die Topfgrößen, die Punkteverteilung und Stichproben je Kategorie aus —
der schnellste Weg, die Schwellen zu beurteilen.

Bauen:

```
pyinstaller mailing_list_updater/MailingListUpdater.spec
```

> **Die Merkmalstabellen in `core.py`** — `MERKMAL_GRUPPE` (Kundengruppe →
> `VIP`/`K`/`BUHA`/`P`), `MERKMAL_BRANCHE_GRUPPE` (dasselbe ersatzweise aus der
> Branche) und `MERKMAL_BRANCHE` (Zusatzfelder wie `Autor`, `Presse/ZS`) — sind
> aus den Häufigkeiten des Altbestands abgeleitet, nicht aus einer Vorgabe des
> Verlags. Vor dem ersten scharfen Lauf gegenprüfen.
>
> Die Ableitung aus der **Branche** ist kein Beiwerk: die Auftragsliste führt
> `Kd.Gr.` meist gar nicht, wohl aber `Branche`. Ohne sie landen alle aus
> Aufträgen rekonstruierten Bestandskunden pauschal auf `K` — im Probelauf
> waren das 139 Buchhändler, 16 Autoren und 10 Pressekontakte, die als
> gewöhnliche Endkunden abgelegt worden wären.
>
> Lässt sich das Merkmal weder aus Kundengruppe noch aus Branche bestimmen,
> setzt das Tool `K` und schreibt in die Auffälligkeit, **was** gefehlt hat.
