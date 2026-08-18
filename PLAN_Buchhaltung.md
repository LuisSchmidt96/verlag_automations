# buchungshilfe — PayPal-Sammelzahlungen in Buchungen zerlegen

> **Stand 18.08.2026 — hier weitermachen.**
> Besprochen und entschieden, noch keine Zeile Code. Nichts angelegt, nichts
> geändert; es existiert nur dieser Plan.
>
> **Bevor es losgeht, werden drei Dinge gebraucht:**
> 1. Ein echter **PayPal-CSV-Export** (Aktivitäten → Alle Transaktionen →
>    Herunterladen → CSV). Ohne ihn sind die Spaltennamen unbekannt.
> 2. Ob sich die **OP-Liste aus Lexware nach Excel exportieren** lässt
>    (Berichte → Offene Posten). Luis prüft.
> 3. Der **Steuerschlüssel für die PayPal-Gebühr** — Frage an den
>    Steuerberater, wird nicht geraten.

## Kontext

Die Buchhaltung läuft in **Lexware buchhalter pro** (SKR03). Ein Kollege aus
dem Wareneingang erfasst die Rechnungen. Der laufende Ablauf ist:

```
Extras → Zahlungsverkehr → Online-Kontoauszüge abholen
Buchen → Online-Kontoauszug abgleichen
```

Für Überweisungen funktioniert das gut: steht die Rechnungsnummer im
Verwendungszweck, findet Lexware den offenen Posten selbst.

**Das Werkzeug tritt nicht an die Stelle dieses Ablaufs.** Der Kontoauszug
wird weiterhin dort abgeglichen; die PayPal-Auszahlung wird wie bisher gegen
**1590** gebucht, damit der Auszug fertig ist. Erst danach setzt das Werkzeug
an und teilt den Betrag auf die einzelnen offenen Posten auf.

**PayPal ist der Bruch.** PayPal überweist gesammelt: eine Bankzeile über
einen Betrag, in dem dutzende Kundenzahlungen und die einbehaltenen Gebühren
stecken. Die Zeile wird auf **1590** (durchlaufende Posten) gebucht — und dann
muss jemand von Hand herausfinden, welche offenen Posten damit bezahlt sind.
Lexware kann das nicht, weil es den Inhalt der Auszahlung nicht kennt. Diese
Information steht nur im PayPal-Umsatzbericht.

Zu 99 % zahlen Kunden an den Verlag; gelegentlich wird auch aus dem
PayPal-Guthaben gezahlt.

## Was gemessen und entschieden wurde

**Der Importweg existiert und ist der sichere.** Buchungen kommen über
`Datei → Import → Text/ASCII` hinein und landen im **Buchungsstapel** — also
zur Durchsicht, nicht direkt im Journal. Erst *Buchen → Stapel ausbuchen*
schreibt sie fest. Dieselbe Form wie beim Access-Werkzeug: Datei erzeugen,
Mensch prüft, Mensch bestätigt.

Die ASCII-Schnittstelle kennt genau diese Felder:

```
Belegdatum · Buchungsdatum · Buchungsperiode · Belegnummernkreis · Belegnummer
Buchungstext(79) · Buchungsbetrag · Sollkonto · Habenkonto · Steuerschlüssel
Kostenstelle · Kostenträger · Buchungsbetrag Euro · Währung
```

Daraus folgen drei Entscheidungen:

* **Kein Feld für Beleganhänge.** Der Wunsch „PDF aus dem Mailfach anhängen"
  geht über diesen Weg nicht — das wäre Lexwares Belegverwaltung, ein eigenes
  Thema. Später zu klären, nicht jetzt.
* **Die Automatik nimmt einem bei PayPal nichts ab.** Sie greift, wenn die
  Rechnungsnummer im Verwendungszweck steht — bei Überweisungen also, bei
  PayPal gerade nicht. Genau das ist die Lücke: **welche Rechnung gemeint ist,
  muss das Werkzeug selbst herausfinden.** Erst danach ist die Buchung
  trivial. Die Automatik schließt hinterher nur noch den Posten, den das
  Werkzeug bereits richtig adressiert hat; sie ist Nachbereitung, nicht Hilfe.

  Das Handbuch sagt zum Import ausdrücklich: „Beim Import findet kein
  automatischer Abgleich der Offenen Posten statt" — deshalb der Schritt
  *Buchen → Offene Posten abgleichen → Automatik* danach.
* **Die Bankzeile erzeugt das Werkzeug NICHT.** Die bucht Luis weiterhin selbst
  aus dem Kontoauszug auf 1590. Würde das Werkzeug sie mitliefern, stünde sie
  doppelt.

**Der PayPal-Bericht kommt künftig als CSV** statt als PDF (PayPal:
Aktivitäten → Alle Transaktionen → Herunterladen → CSV). Dieselben Daten, aber
mit Spalten statt Layout. Ein PDF-Leser wäre bei jeder Layoutänderung von
PayPal wieder kaputt.

**Zugeordnet wird über Name und Betrag**, nicht über eine Nummer — die
PayPal-Zahlung trägt keine Rechnungsnummer. Das ist der Kern der Arbeit und
dasselbe Problem wie beim Adressabgleich: kein gemeinsamer Schlüssel. Also
dieselbe Lösung, die sich dort bewährt hat — blocken, punkten, Zweifelsfälle
vorlegen statt raten.

Der Betrag ist dabei der bessere Schlüssel als der Name: er ist exakt, und
zwei offene Rechnungen über denselben Betrag desselben Kunden sind selten.
Der Name entscheidet, wenn der Betrag mehrdeutig ist — und umgekehrt deckt ein
passender Name bei abweichendem Betrag eine Teilzahlung oder einen Skontoabzug
auf, die beide zur Handprüfung gehören.

## Zuschnitt

Nur die PayPal-Aufteilung. Wiederkehrende Buchungen (Hetzner & Co.) kommen
später — und vorher ist zu prüfen, ob Lexware das mit **Buchungsvorlagen** und
**wiederkehrenden Buchungen** nicht längst selbst kann. Etwas nachzubauen, was
das Programm schon mitbringt, wäre die schlechteste Art, Zeit zu investieren.

## Aufbau

Neues Werkzeug nach dem Hausmuster, **kein geteilter Code** mit den anderen
(siehe `CLAUDE.md`):

```
buchungshilfe/
├── __init__.py
├── core.py                  ← Einlesen, Zuordnen, Buchungssätze bauen
├── app.py                   ← Tkinter-GUI
├── main.py
├── Buchungshilfe.spec
└── README.md
```

`tools/update_and_build.ps1` findet die neue `.spec` automatisch.

### Eingaben

| Datei | woher |
|---|---|
| PayPal-Umsatzbericht (CSV) | PayPal → Aktivitäten → Alle Transaktionen → Herunterladen → CSV |
| Offene Posten (Excel) | Lexware → Berichte → Offene Posten → Export nach Excel |

*Noch zu prüfen:* ob sich die OP-Liste wirklich nach Excel exportieren lässt.
Falls nicht, ist der Debitorenstamm plus die Rechnungsliste der Ersatzweg.

### core.py

* `lade_paypal(pfad)` — CSV lesen, tolerant gegen Spaltenreihenfolge (dieselbe
  Kopfzeilensuche wie in `mailing_list_updater/core.py`: nach Pflichtspalten
  suchen statt feste Positionen annehmen).
* `gruppiere_auszahlung(zeilen)` — die Einzelzahlungen einer Auszahlung
  zusammenfassen. PayPal führt die Auszahlung als eigene Zeile („Allgemeine
  Abbuchung" / „Auszahlung"); alles davor seit der letzten Auszahlung gehört
  dazu.
* `finde_op(zahlung, op_liste)` — Zuordnung über Name und Betrag:
  * **Betrag exakt + Name ähnlich** → sicher
  * **Betrag exakt, Name unklar** oder **Name sicher, Betrag abweichend**
    (Teilzahlung, Skonto) → vorlegen
  * sonst → kein Treffer
* `baue_buchungen(...)` — je Zahlung ein Satz, dazu die Gebühren.

### Die Buchungssätze (SKR03)

| Vorgang | Soll | Haben | Belegnummer |
|---|---|---|---|
| Kunde zahlt per PayPal | **1590** | Debitorenkonto | Rechnungsnummer |
| PayPal-Gebühr | **4970** Nebenkosten des Geldverkehrs | **1590** | Auszahlungsdatum |
| Auszahlung aufs Bankkonto | *(bucht Luis selbst)* | | |

Kein Steuerschlüssel bei der OP-Ausbuchung — die Umsatzsteuer steckt bereits
in der Rechnung. **Der Steuerschlüssel der PayPal-Gebühr ist mit dem
Steuerberater zu klären** und wird nicht geraten.

### Die eingebaute Kontrolle: 1590 muss aufgehen

Nach dem Ausbuchen muss das Konto 1590 für diese Auszahlung **auf null**
stehen: Summe der Kundenzahlungen minus Gebühren = Auszahlungsbetrag. Das
Werkzeug rechnet das vor dem Schreiben nach und schreibt die Datei **nicht**,
wenn die Probe nicht aufgeht — dann fehlen Zeilen im Bericht oder eine Zahlung
wurde übersehen.

Das ist das Gegenstück zu `pruefe_zeitraeume()` im Adresswerkzeug: eine
Rechnung, die das Werkzeug selbst prüfen kann, statt sich auf den Bediener zu
verlassen.

### app.py

Drei Reiter, nach dem Muster des Adresswerkzeugs — Liste oben (`ttk.Treeview`
mit Suchfeld), Einzelheiten unten:

| Reiter | Inhalt |
|---|---|
| **Sicher** | Betrag stimmt, Name passt — durchlaufen lassen |
| **Prüfen** | Teilzahlungen, Skonto, mehrdeutige Namen; je Fall die Kandidaten mit Punktzahl |
| **Ohne Zuordnung** | keine offene Rechnung gefunden — Rückzahlung, Spende, private Zahlung, oder die Rechnung fehlt noch |

Oben die Probe: `Auszahlung 1 250,00 € = 47 Zahlungen 1 289,30 € − Gebühren 39,30 €` ✓

### Ausgabe

Nach `buchungen/<Jahr>/<Auszahlungsdatum>/`:

| Datei | Inhalt |
|---|---|
| **`buchungsstapel.txt`** | die ASCII-Importdatei für Lexware |
| `anleitung.txt` | der Importweg Schritt für Schritt |
| `protokoll.xlsx` | jede Zuordnung mit Punktzahl und Begründung |

### Der Weg nach Lexware

Die Schritte 1 und 2 sind der gewohnte Ablauf, unverändert:

1. `Extras → Zahlungsverkehr → Online-Kontoauszüge abholen`
2. `Buchen → Online-Kontoauszug abgleichen` — die PayPal-Auszahlung wie bisher
   gegen **1590** buchen. Der Auszug ist damit erledigt.

Erst jetzt kommt das Werkzeug:

3. Datenbestand sichern.
4. `Datei → Import → Text/ASCII`, Buchungsdaten, Profil einmalig anlegen.
5. `Ansicht → Buchungsstapel` — Buchungen prüfen, bei Bedarf ändern.
6. `Buchen → Stapel ausbuchen`.
7. `Buchen → Offene Posten abgleichen → Automatik` — schließt die Posten, die
   das Werkzeug bereits richtig adressiert hat.
8. Kontrolle: Konto **1590** steht für diese Auszahlung auf null.

Das Handbuch warnt ausdrücklich vor **Mehrfachimport** — Lexware prüft nicht,
ob ein Stapel schon eingelesen wurde. Das Werkzeug legt deshalb je Auszahlung
einen eigenen Ordner an und vermerkt im Protokoll, wann geschrieben wurde.

## Prüfen

1. **Eine echte Auszahlung durchrechnen**, aber nichts importieren: PayPal-CSV
   und OP-Liste laden, Probe muss aufgehen. Weicht sie ab, taugt der Rest
   nicht.
2. **Stichprobe von Hand**: 10 Zuordnungen gegen den PayPal-Bericht und die
   OP-Liste halten.
3. **Import in eine Testfirma** oder in eine Sicherungskopie des Datenbestands,
   nicht im Echtbestand. Danach 1590 kontrollieren.
4. Erst wenn zwei Auszahlungen sauber durchgelaufen sind, im Echtbestand.

## Offene Punkte

* Lässt sich die OP-Liste nach Excel exportieren? (Luis prüft)
* Steuerschlüssel für die PayPal-Gebühr — Frage an den Steuerberater.
* Wie heißen die PayPal-CSV-Spalten in der deutschen Fassung genau? Dafür
  brauche ich einen echten Export, bevor der Leser geschrieben wird.
* Die gelegentlichen Zahlungen **aus** dem PayPal-Guthaben: erst einmal in
  „Ohne Zuordnung" sammeln und von Hand buchen. Automatisieren lohnt sich erst,
  wenn klar ist, wie oft das vorkommt.
* Später: wiederkehrende Buchungen — vorher prüfen, was Lexware selbst kann.
* Später: Belege aus dem Mailfach — geht nicht über die ASCII-Schnittstelle,
  Belegverwaltung von Lexware pro prüfen.

## Nach der Freigabe

Die Entscheidungen dieses Gesprächs (Lexware pro, SKR03, 1590 als
Verrechnungskonto, Importweg über den Buchungsstapel, Zuordnung über Name und
Betrag) gehören in den Projektspeicher — im Planmodus darf ich außer dieser
Datei nichts schreiben, das hole ich als Erstes nach.
