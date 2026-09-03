# Buchdurchgang

Bringt ein Buch **in einem Zug** durch die drei Werkzeuge — Cover, pi&bi,
Shopware — und legt alles in **einen Ordner je Buch**. Nach jedem Schritt steht
eine Checkliste; erst wenn sie abgehakt ist, geht es weiter.

```
0. Buch wählen   Umschlag-PDF + ONIX-XML, beide gegeneinander geprüft
1. Cover         cover_previews      2D / 3D / TIF
2. pi & bi       pi_bi_generator     Presse- und Buchinformation
3. Shop          shopware_publisher  Produkt als Entwurf
4. Ablegen       Buchordner auf den Netzordner, mit Nachprüfung
```

## Kategorien: geraten, nicht gefragt

In Schritt 3 gibt es **keinen Kategorie-Auswähler**. Das Werkzeug sucht je
beteiligter Person im Shop (`Nachname, Vorname`) und setzt, was es findet —
ohne Rückfrage. Wen es *nicht* findet, schreibt es ins Protokoll.

Der Grund: nachsehen und ergänzen tut man ohnehin im Shopware-Backend, wo auch
die Sachkategorien hingehören. Genau das steht in der Checkliste — und das
Backend geht nach dem Anlegen **von selbst auf**.

## Örtlich arbeiten, am Ende ablegen

Die Schritte 1–3 schreiben in einen **örtlichen Arbeitsordner**, erst Schritt 4
legt den fertigen Buchordner auf den Netzordner. Das ist kein Geschmack,
sondern gemessen: derselbe Cover-Schritt braucht örtlich **1,7 s**, über eine
Netzeinbindung lief er wiederholt in Zeitüberschreitungen — er bewegt rund
11 MB. Ein kompletter Durchgang dauert so **rund 4 Sekunden**.

| Einstellung | Bedeutung |
|---|---|
| `arbeitsordner` | die schnelle Werkbank (leer = `~/Buch_Arbeit`) |
| `ablageort` | wohin der fertige Ordner in Schritt 4 wandert |

Beim Ablegen wird **kopiert und nachgeprüft** (Dateigröße am Ziel), nicht bloß
kopiert: über das Netz bricht ein Kopiervorgang gern in der Mitte ab, und eine
halbe Datei sieht aus wie eine ganze. Liegen am Ziel schon gleichnamige
Dateien, wandern sie nach `_alt/<Zeitstempel>/` statt überschrieben zu werden.

Die Photoshop-Zwischendateien (`_slot_*.png`, `_mockup_*.jsx`) bleiben zurück —
sie haben im Artikelordner nichts zu suchen.

**Die örtliche Kopie bleibt stehen.** Sie ist das Sicherheitsnetz, falls beim
Übertragen etwas hakt; wann sie weg kann, sagt die Checkliste. Der Preis:
dieselben Dateien liegen zweimal, und wer später örtlich etwas ändert, ändert
nicht den Netzordner.

## Warum dieses Werkzeug die anderen einbindet

Die Hausregel sagt: **kein geteilter Code zwischen den Werkzeugen**. Sie
verbietet, eine gemeinsame Bibliothek *herauszulösen*. Hier wird nichts
herausgelöst — der Durchgang **benutzt** die drei, sie bleiben einzeln
lauffähig. Die Alternative wäre gewesen, den Photoshop-Aufruf, die
docx-Vorlagen und den Shopware-Client zu duplizieren; das wären rund
3000 Zeilen doppelt.

Der Preis ist Kopplung: ändert sich `cover_previews/core.py`, ändert sich auch
der Durchgang. Deshalb rufen `cover_previews` und `pi_bi_generator` inzwischen
**dieselben** kopflosen Funktionen auf wie er (`core.lauf()` bzw.
`core.erzeuge_alle()`) — es gibt nur einen Weg, nicht zwei, die auseinanderlaufen.

## Die Eingangsprobe

Bevor irgendetwas erzeugt wird, werden Umschlag-PDF und ONIX gegeneinander
gehalten:

| Probe | woher | Folge bei Abweichung |
|---|---|---|
| **ISBN** | Text im PDF vs. `<ProductIdentifier>` | **Abbruch** — etwas ist vertauscht |
| **Format** | gemessene Vorderseite vs. `<Measure>` | **Warnung** — Beschnittzugabe schwankt |

Die Formatprobe fängt den Fall, den die ISBN nicht fängt: richtige Nummer,
falsche Datei — etwa eine ältere Auflage mit anderem Format. Beispiel:

```
gemessen  21,6 x 21,6 cm   (mit 3 mm Beschnitt)
ONIX      21,0 x 21,0 cm   ✓
```

## Der Buchordner

Alles zu einem Buch liegt beieinander, benannt wie in `cover_previews`
(`<Kurzcode>_<Titel>`):

```
<Ablageort>/05-609-4_Inszeniertes Glück/
    2D_300_05-609-4.jpg   2D_72_05-609-4.jpg
    3D_300_05-609-4.jpg   3D_72_05-609-4.jpg
    05-609-4.png          3D_300_05-609-4.tif
    PI_05-609-4.docx      PI_05-609-4.html
    BI_05-609-4.docx      BI_05-609-4.html
    cover_9783955056094.jpg
    durchgang.json        <- Stand und Häkchen
```

Nach Schritt 4 liegt derselbe Ordner auf dem Netz — ohne die
Photoshop-Zwischendateien, aber **mit** `durchgang.json`: dort steht, welcher
Schritt wann gelaufen ist und was abgehakt wurde.

`durchgang.json` liegt **beim Buch**, nicht beim Werkzeug: so sieht auch Wochen
später und an einem anderen Rechner jeder, welcher Schritt wann gelaufen ist,
mit welcher Datei, und was geprüft wurde. Ein unterbrochener Durchgang lässt
sich damit fortsetzen — Fenster zu, Fenster auf, Stand ist wieder da.

## Die Tore

**Weiter** wird erst frei, wenn der Schritt *gelaufen* **und** *abgehakt* ist.
Die Häkchen selbst sind bis zum Lauf gesperrt — eine Checkliste, die man vorher
abhaken kann, prüft nichts.

In der Schrittliste links: `·` offen · `◐` gelaufen, aber nicht abgehakt ·
`✓` fertig.

## Konfiguration

`config.json` liegt neben der .exe und hat je Werkzeug einen Abschnitt:

```json
{
  "ablageort": "\\\\C019\\d\\Online\\Webseite\\Artikeldaten",
  "cover_previews":     { "vorlagen_dir": "..\\CoverPreviews\\_NEU_Vorlage" },
  "pi_bi_generator":    { },
  "shopware_publisher": { "umgebungen": { "dev": { }, "prod": { } } }
}
```

Das ist kein Zufall, sondern nötig: die drei Werkzeuge legen ihre `config.json`
je **neben ihre eigene .exe**. In einer gemeinsamen .exe würden sich alle drei
auf dieselbe Datei stürzen. Deshalb hält der Durchgang seine eigene und reicht
die fertigen Einstellungen hinein; `lade_config()` der anderen wird nie gerufen.

**Die Shopware-Zugangsdaten sind eigene** — sie stehen im Abschnitt
`shopware_publisher` dieser Datei, nicht in der des Publishers. Abtippen muss
man sie trotzdem nicht: der Knopf **„Zugang aus dem ShopwarePublisher
übernehmen"** in Schritt 3 holt Shop-URL, Zugriffsschlüssel, das
**verschlüsselte** Secret samt Salt und die fünf Zuordnungen (Steuer, Währung,
Hersteller, Verkaufskanal, Seiten-Layout) herüber.

Das Master-Passwort wandert dabei **nicht** mit und wird nirgends gespeichert —
es entsperrt das übernommene Secret hinterher genauso wie im Publisher, weil
Chiffretext und Salt selbsttragend sind.

Gesucht wird in `..\ShopwarePublisher\config.json` (so liegen die Werkzeuge
unter `VR-Tools\` nebeneinander) und im Quellbaum daneben.

**`_NEU_Vorlage` wird nicht mitgeliefert** (rund 480 MB). Der Eintrag
`cover_previews.vorlagen_dir` zeigt auf den Ordner neben `CoverPreviews.exe`;
beide liegen unter `VR-Tools\` nebeneinander.

## Grenzen

* Der **3D-Zweig steuert Photoshop per COM** und läuft nur unter Windows.
  Anderswo macht er einen Trockenlauf: es entstehen JSX und Slot-PNGs, aber
  **kein Mockup**. Das Fenster sagt das ausdrücklich — ein halber Lauf soll
  nicht wie ein ganzer aussehen.
* Ist der **Ablageort nicht erreichbar, bricht der Cover-Schritt ab**, statt
  wie das Einzelwerkzeug still nach `cover_output/` auszuweichen. In einer
  Kette wäre ein Buchordner an unerwarteter Stelle fatal: die beiden
  Folgeschritte suchten ins Leere.
* Der **Werbetext wird nicht gekürzt.** Das Werkzeug erzeugt ihn vollständig;
  gekürzt wird wie bisher in Word, und die Checkliste fragt danach.
