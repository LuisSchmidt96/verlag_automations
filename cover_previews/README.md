# Cover-Previews-Generator

Erzeugt aus einem **druckfertigen Umschlag-PDF** (Rückseite | Rücken | Vorderseite,
plus Einschlag/Beschnitt) automatisch Vorschau-Bilder:

- **2D-Vorderseite** — flach, auf die Buchkante beschnitten, als PNG.
- **3D-Mockup (Vorderseite + Rücken)** — durch Wiederverwendung eines
  bestehenden Photoshop-Mockups (`.psd` mit Smart-Object), gesteuert per COM.

Jeweils in **300 dpi (Druck)** und **72 dpi / ~800 px (Web)**. Die Dateinamen
folgen der Konvention der anderen Tools: `2D_300_{kurzcode}.png`,
`3D_300_{kurzcode}.png` usw. (Kurzcode aus der ISBN, z. B. `05-572-1`).

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

1. **Umschlag-PDF** wählen → Vorschau mit erkannten Linien erscheint.
2. Linien bei Bedarf justieren (blau ziehen).
3. Ausgaben ankreuzen. Für **3D**: **Mockup-PSD** wählen und den Namen der
   **Smart-Object-Ebene** eintragen (den findet man, indem man das PSD einmal in
   Photoshop öffnet und in der Ebenen-Palette nachsieht).
4. **Erstellen** → 2D wird sofort erzeugt; für 3D öffnet das Tool im Hintergrund
   eine Kopie des Mockups in Photoshop, tauscht das Cover ins Smart-Object,
   exportiert das PNG und schließt **ohne zu speichern** (Original bleibt unberührt).

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
