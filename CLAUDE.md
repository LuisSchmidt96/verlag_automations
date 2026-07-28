# Hinweise für Claude Code

`README.md` beschreibt Aufbau und Build. Hier steht nur, was man sonst erst
mühsam herausfindet.

## Konventionen

* **Alles auf Deutsch** — Code-Kommentare, Docstrings, Commit-Nachrichten, GUI,
  README. Auch Variablen- und Funktionsnamen (`ziel_ordner`, `finde_schnittlinien`).
* **Kein geteilter Code zwischen den Tools.** Jedes Tool ist eigenständig und
  wird zu einer eigenen .exe gebaut; doppelter Code ist hier Absicht, keine
  Schlamperei. Nicht „aufräumen“ und eine gemeinsame Bibliothek daraus machen.
* Muster je Tool: `core.py` (reine Logik, keine UI) / `app.py` (Tkinter) /
  `main.py` / `<Name>.spec`. Laufzeitdaten (`config.json`, Ausgabeordner) liegen
  direkt neben der .exe, nicht in einem `data/`-Unterordner.
* Kommentare erklären das **Warum**, besonders bei den Eigenheiten von Photoshop,
  Shopware und den Druckvorlagen. Diese Begründungen nicht wegkürzen — sie sind
  der Grund, warum der Code so aussieht.

## Umgebung

* Entwickelt wird unter **Linux**, eingesetzt wird unter **Windows**.
* Der 3D-Zweig von `cover_previews` steuert **Photoshop per COM** (`pywin32`) und
  läuft nur unter Windows. Ohne Windows macht er einen Dry-Run und schreibt das
  erzeugte `.jsx` in den Ausgabeordner — damit lässt sich der Photoshop-Schritt
  prüfen, aber **nicht ersetzen**: alles, was den 3D-Zweig anfasst, muss am Ende
  auf einem Windows-Rechner mit Photoshop gegengeprüft werden.
* Die Mockup-PSDs (`cover_previews/_NEU_Vorlage/`, ~480 MB) sind **nicht**
  eingecheckt. Lokal lassen sie sich mit `psd-tools` auslesen (Ebenenbaum,
  Ebenen-IDs, Composite) — das ist der schnellste Weg, Fragen zu den Vorlagen zu
  klären, ohne Photoshop.

## Laufende Arbeiten

* [`cover_previews/WIP_TIF-Freistellung.md`](cover_previews/WIP_TIF-Freistellung.md)
  — freigestelltes CMYK-TIF zum 3D-Mockup **und** der nicht gefundene
  Artikeldaten-Share. Enthält die gemessenen Befunde (woher der weiße Strich an
  der Buchkante kommt, welche Ebenen die Spiegelung sind) und die offenen Punkte.
