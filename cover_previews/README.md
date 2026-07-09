# Cover-Previews-Generator

Erzeugt aus einem **druckfertigen Umschlag-PDF** (Rückseite | Rücken | Vorderseite,
plus Einschlag/Beschnitt) automatisch die Artikeldaten-Bilder – nach dem
manuellen Photoshop-Ablauf des Verlags:

- **2D-Vorderseite** — flach, auf die Buchkante beschnitten, als JPEG:
  `2D_300_{sc}.jpeg` und `2D_72_{sc}.jpeg` (**pixelgleich**, nur der DPI-Eintrag
  unterscheidet sich, wie „neu berechnen aus“).
- **3D-Mockup (Vorderseite + Rücken)** — über die zum **Buchformat** passende
  Mockup-Vorlage (`_NEU_Vorlage/…psd`): transparentes `{sc}.png` plus
  `3D_300_{sc}.jpg` / `3D_72_{sc}.jpg`. Gesteuert per Photoshop-COM.

Kurzcode `sc` aus der ISBN (z. B. `05-572-1`).

## Vorlagen-Automatik (Buchformat → PSD)

Für jedes Buchformat gibt es eine eigene Mockup-Vorlage (`17x24.psd`, `21x14.psd`
usw.). Das Tool misst aus den Marken den **Vorderseiten-Trim** und wählt die
formatnächste Vorlage **automatisch** (Dropdown-Override möglich, z. B. für
`EBOOK` oder `MaGeBl`). Die Zuordnung Cover/Rücken → Smart-Objekt steckt in der
mitgelieferten `vorlagen_map.json` (Layer-IDs aus einer einmaligen PSD-Analyse;
die gespiegelte Reflexion teilt sich die Smart-Object-Quelle und aktualisiert
sich mit). Bei **weißem Umschlag** wird automatisch eine leichte Grau-Korrektur
(Selektive Farbkorrektur, Weiß +10 % Schwarz) angewandt.

Der **Vorlagen-Ordner** (`_NEU_Vorlage`) wird in der GUI gewählt bzw. per
`config.json → vorlagen_dir` gesetzt (z. B. auf den Netzpfad
`…/Artikeldaten/_NEU_VORLAGE`). Die PSDs (~480 MB) sind **nicht** eingecheckt.

## So funktioniert die Marken-Erkennung

Die Schnitt-/Falzmarken stehen im PDF als Vektor-Striche. Regel: Eine
**vertikale** Marke, die den **oberen oder unteren Blattrand** berührt, markiert
eine senkrechte Schnittlinie (Rückseite | Rücken | Vorderseite); eine
**horizontale** Marke am **linken/rechten Rand** die Ober-/Unterkante. Marken,
die keinen Blattrand berühren (z. B. Logo-Rahmen im Innenbereich), werden
ignoriert. Aus den senkrechten Schnitten sind die **zwei breitesten Felder** die
beiden Buchdeckel, das Feld dazwischen der Rücken.

In der GUI werden die erkannten Linien blau über die Vorschau gelegt und lassen
sich vor dem Rendern **mit der Maus nachjustieren**.

## Bedienung

1. **Umschlag-PDF** wählen → Vorschau mit erkannten Linien; die passende
   **Vorlage** wird automatisch anhand des Formats gesetzt.
2. Linien bei Bedarf justieren (blau ziehen); Vorlage ggf. im Dropdown ändern.
3. Ausgaben ankreuzen (2D / 3D). Einmalig den **Vorlagen-Ordner** setzen.
4. **Erstellen** → 2D sofort (JPEG); für 3D öffnet das Tool im Hintergrund eine
   Kopie der Vorlage in Photoshop, setzt Cover + Rücken in die Smart-Objekte,
   exportiert transparentes PNG + flache 3D-JPEGs und schließt **ohne zu
   speichern** (Original-PSD bleibt unberührt).

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
├── mockups/       optionale Mockup-PSDs (lokal, nicht eingecheckt)
└── data/          config.json + cover_output/ (Laufzeit, nicht eingecheckt)
```
