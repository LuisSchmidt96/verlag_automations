# Cover-Previews-Generator

Erzeugt aus einem **druckfertigen Umschlag-PDF** (Rückseite | Rücken | Vorderseite,
plus Einschlag/Beschnitt) automatisch die Artikeldaten-Bilder – nach dem
manuellen Photoshop-Ablauf des Verlags:

- **2D-Vorderseite** — flach, auf die Buchkante beschnitten, als JPEG:
  `2D_300_{sc}.jpg` und `2D_72_{sc}.jpg` (**pixelgleich**, nur der DPI-Eintrag
  unterscheidet sich, wie „neu berechnen aus“).
- **3D-Mockup (Vorderseite + Rücken)** — über die zum **Buchformat** passende
  Mockup-Vorlage (`_NEU_Vorlage/…psd`): transparentes `{sc}.png` plus
  `3D_300_{sc}.jpg` / `3D_72_{sc}.jpg`. Gesteuert per Photoshop-COM.
- **Freigestelltes CMYK-TIF** (optional) — `3D_300_{sc}.tif`: nur das Buch
  (ohne Spiegelung/Hintergrund), auf Weiß, mit Vektor-**Beschneidungspfad**
  „Pfad 1“, in **CMYK** mit eingebettetem Profil. Zum Platzieren in Newslettern
  / InDesign auf farbigem Grund. Siehe unten.

Kurzcode `sc` aus der ISBN (z. B. `05-572-1`).

## Vorlagen-Automatik (Buchformat → PSD)

Für jedes Buchformat gibt es eine eigene Mockup-Vorlage (`17x24.psd`, `21x14.psd`
usw.). Das Tool misst aus den Marken den **Vorderseiten-Trim** und wählt die
formatnächste Vorlage **automatisch** (Dropdown-Override möglich, z. B. für
`EBOOK` oder `MaGeBl`). Bei **weißem Umschlag** wird automatisch eine leichte
Grau-Korrektur (Selektive Farbkorrektur, Weiß +10 % Schwarz) angewandt.

Der **Vorlagen-Ordner** (`_NEU_Vorlage`) wird in der GUI gewählt bzw. per
`config.json → vorlagen_dir` gesetzt (z. B. auf den Netzpfad
`…/Artikeldaten/_NEU_VORLAGE`). Die PSDs (~480 MB) sind **nicht** eingecheckt.

## Wie die Smart-Objekte befüllt werden (der springende Punkt)

`vorlagen_map.json` beschreibt jede Vorlage; erzeugt wird sie aus den PSDs mit

```
python -m cover_previews.psd_analyse          # nach jeder PSD-Revision neu laufen lassen
```

Zwei Eigenheiten der Vorlagen muss man kennen — sonst zerfällt das Mockup:

**1. Jede Vorlage hat ihren eigenen Inhaltsmaßstab `k`.** Photoshop bildet den Inhalt
eines Smart-Objekts über einen festen Transform auf die 3D-Fläche ab; die **Maße des
Inhalts** sind dabei der Maßstab. Der Inhaltsraum liegt je nach Vorlage bei 162–310 dpi
(nur `17x24` und `28x21` sind 300 dpi). Wer immer 300 dpi liefert, setzt den Inhalt in
den meisten Vorlagen um bis zu 1,8× zu groß ein. `content_px_per_cm` hält `k` fest;
das Cover wird exakt auf die Slot-Größe gebracht.

**2. Die Rücken-Slots sind absichtlich überbreit.** Der Rücken-Slot von `17x24` ist
z. B. 10 cm breit, trägt sein Motiv aber nur in einem schmalen, an der Cover-Kante
ausgerichteten Streifen — der Rest ist transparent. Die Leinwand nimmt damit **jede**
Rückendicke auf. Genau so füllt das Tool sie: Motiv `Dicke × k` px breit, an der
`anchor`-Kante, Rest transparent. Der Rücken ist dadurch **maßgetreu** (ein 4-cm-Buch
sieht dicker aus als ein 2-cm-Buch) und bleibt trotzdem bündig am Cover.

Der Inhalt behält also **immer** die Größe seines Slots. Liefert man nur den schmalen
Rückenstreifen, skaliert Photoshop den Transform auf dessen Maße — der Rücken schrumpft
zur Außenkante und löst sich sichtbar vom Buch.

Buch und Spiegelung sind eigene Slots (mit teils unterschiedlichen Maßen, `29x22` hat
zwei Cover-Größen) und werden beide befüllt.

## Freigestelltes CMYK-TIF (`3D_300_{sc}.tif`)

Zusätzlich zum 3D-Mockup kann das Tool das Buch **freigestellt** als CMYK-TIF
ausgeben — zum Platzieren in Newslettern / InDesign auf farbigem Grund. Es folgt der
vom Setzer gelieferten Referenz:

- **nur das Buch** (Spiegelung + Hintergrund raus), auf **Weiß**;
- freigestellt über einen **Vektor-Beschneidungspfad** „Pfad 1“ — nicht über einen
  Alphakanal. Für CMYK-Druck ist der Pfad der Standard, den InDesign/RIPs verstehen;
  CMYK+Alpha ist ein 5-Kanal-Sonderfall, den manche Workflows verweigern;
- **CMYK** mit eingebettetem Profil (`changeMode` nutzt den CMYK-Arbeitsfarbraum des
  Photoshop — auf dem Verlagsrechner also dessen Profil), 300 dpi, LZW.

Die weiche, transparente Fassung **mit** Spiegelung gibt es weiter als `{sc}.png`
(RGB, echter Alphakanal) — das ist die Bildschirm-/HTML-Variante. TIF = Druck, PNG =
Bildschirm; sie teilen sich nach Medium auf.

Damit nur das Buch übrig bleibt, wird die **Spiegelungs-Gruppe** der Vorlage
ausgeblendet. Welche Gruppe das ist, leitet `psd_analyse.py` her (Schlüssel
`spiegelung` in `vorlagen_map.json`): Buch- und Spiegelungs-Ebenen trennen sich an der
größten Lücke ihrer `top`-Werte (die Spiegelung sitzt tiefer); zur Spiegelung die
äußerste Vorfahr-Gruppe ohne Buch-Ebene. Der weiße **Falz-Schein** am Rand (Alpha 1)
fällt weg, weil die Auswahl vor dem Pfad um 2 px verkleinert wird — kein weißer Strich.

Steuerung: Ankreuzfeld „3D zusätzlich als CMYK-TIF“ (an die 3D-Ausgabe gekoppelt),
Stand in `config.json → tif_erzeugen`.

## So funktioniert die Marken-Erkennung

Die Schnitt-/Falzmarken stehen im PDF als Vektor-Striche. Die Setzer machen das
allerdings **unterschiedlich**: mal sind die Marken bis an die Blattkante gezogen
(`US_Kleindenkmale`), mal mit Abstand davor abgesetzt (`US_Kellerkinder` — dort
beginnt die Marke erst 9 pt unter der Kante). Auf „berührt den Blattrand“ ist
also kein Verlass.

Erkannt wird deshalb so:

1. Nur **kurze** Striche (unter 15 % der Seitenkante) — lange sind Rahmen, kein
   Marken.
2. Sie müssen im **äußeren Rand** liegen (`marken_band_pt`, Standard 60 pt), also
   außerhalb des Anschnitts.
3. Entscheidend: sie müssen **paarweise** auftreten — dieselbe x-Position oben
   *und* unten (bzw. y links *und* rechts). Motiv im Umschlag tut das praktisch
   nie. Ohne diese Bedingung würde z. B. der **Barcode** von `US_Kellerkinder`
   mitgezählt: er sitzt rund 100 pt über der Unterkante und ist ebenfalls kurz.

Findet das keine vier senkrechten Schnitte, greift die alte, strengere Regel
(Marke berührt die Kante) als Rückfall.

Aus den senkrechten Schnitten sind die **zwei breitesten Felder** die beiden
Buchdeckel, das Feld dazwischen der Rücken.

In der GUI werden die erkannten Linien blau über die Vorschau gelegt und lassen
sich vor dem Rendern **mit der Maus nachjustieren**.

## Ablage (Artikeldaten-Share)

Die fertigen Bilder landen je Buch in einem eigenen Ordner unter

```
\\C019\d\Online\Webseite\Artikeldaten\<Kurzcode>_<Titel>\
```

(einstellbar über `config.json → artikeldaten_dir`; ist der Share nicht
erreichbar, fällt das Tool auf `cover_output/` neben der .exe zurück).

Nach dem Einlesen des PDFs sucht das Tool anhand des Kurzcodes einen **vorhandenen**
Ordner (`05-597-4_Oberkirch`) und wählt ihn aus. Gibt es keinen, schlägt es einen
neuen vor — der Kurzcode steht fest, den Titelteil tippst du daneben ein; vor dem
Anlegen wird gefragt.

Liegen im Zielordner schon gleichnamige Dateien, werden sie **nicht** überschrieben,
sondern nach `_alt/<Zeitstempel>/` verschoben; die neuen behalten die regulären
Namen. Die Konvention auf dem Share ist `.jpg` (nicht `.jpeg`) — sonst lägen die
neuen Bilder neben den alten, statt sie abzulösen.

## Hardcover / Softcover

Die Vorlagen haben eine **Falz**-Ebene (die Rille am Buchdeckel neben dem Rücken).
Der Schalter im GUI blendet sie ein (Hardcover) bzw. aus (Softcover). Die Sichtbarkeit
wird immer **explizit** gesetzt: die Vorlagen sind sich uneinig, wie sie ausgeliefert
werden (bei `16x16` und `29x22` ist die Falz aus, sonst an). `EBOOK` hat keine.

## Bedienung

1. **Umschlag-PDF** wählen → Vorschau mit erkannten Linien; die passende
   **Vorlage** wird automatisch anhand des Formats gesetzt.
2. Linien bei Bedarf justieren (blau ziehen); Vorlage ggf. im Dropdown ändern.
3. Ausgaben ankreuzen (2D / 3D). Einmalig den **Vorlagen-Ordner** setzen.
4. **Erstellen** → 2D sofort (JPEG); für 3D öffnet das Tool die Vorlage in
   Photoshop, setzt Cover + Rücken in die Smart-Objekte, exportiert
   transparentes PNG + flache 3D-JPEGs und schließt sie **ohne zu speichern**
   (`finally`, also auch bei einem Fehler) — die Vorlagen-PSD auf der Platte
   bleibt unberührt.

Die 3D-Bilder bekommen oben/links/rechts `rand_cm` Rand (Standard 1,5 cm), unten
`rand_unten_cm` (Standard **0**): das Bild endet dort, wo die Spiegelung ausläuft,
sonst schwebt das Buch über einer leeren weißen Fläche.

## 3D / Photoshop-Voraussetzungen

- Läuft der 3D-Schritt, muss **Photoshop auf demselben Windows-PC** installiert
  sein (Steuerung per COM, `pywin32`).
- Ohne Windows/Photoshop macht das Tool für 3D einen **Dry-Run**: es schreibt das
  ausgeschnittene Cover-PNG und das erzeugte `.jsx`-Skript in den Ausgabeordner —
  so lässt sich der Photoshop-Schritt prüfen bzw. manuell ausführen.

## Bauen (Windows, aus dem Repo-Wurzelordner)

```
pyinstaller cover_previews/CoverPreviews.spec
```

oder alle Tools zusammen: `.\tools\update_and_build.ps1`.

## Dateien/Ordner

```
cover_previews/
├── core.py        Logik: PDF rastern, Marken erkennen, ausschneiden, 2D/3D
├── app.py         Tkinter-GUI (Vorschau + Linien justieren)
├── main.py        Einstiegspunkt für den .exe-Build
├── CoverPreviews.spec
├── beispiele/     Test-PDF(s) (lokal, nicht eingecheckt)
└── mockups/       optionale Mockup-PSDs (lokal, nicht eingecheckt)
```

Laufzeitdaten (`config.json`, `cover_output/`) legt das Tool direkt neben der
.exe ab – kein `data/`-Unterordner, wie bei den anderen Tools.
