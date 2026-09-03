# Buchdurchgang

Bringt ein Buch **in einem Zug** durch die drei Werkzeuge — Cover, pi&bi,
Shopware — und legt alles in **einen Ordner je Buch**. Nach jedem Schritt steht
eine Checkliste; erst wenn sie abgehakt ist, geht es weiter.

```
0. Buch wählen   Umschlag-PDF + ONIX-XML, beide gegeneinander geprüft
1. Cover         cover_previews      2D / 3D / TIF
2. pi & bi       pi_bi_generator     Presse- und Buchinformation
3. Shop          shopware_publisher  Produkt als Entwurf
```

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

**Die Shopware-Zugangsdaten sind eigene.** Schlüssel und verschlüsseltes Secret
stehen im Abschnitt `shopware_publisher` dieser Datei, nicht in der des
Publishers — sie müssen einmal separat gesetzt werden.

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
