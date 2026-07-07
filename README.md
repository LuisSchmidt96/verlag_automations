# verlag_automations

Sammlung eigenständiger Automations-Tools für den Verlag Regionalkultur.
Jedes Tool lebt in einem eigenen Unterordner, wird zu einer **eigenen .exe**
gebaut und ist unabhängig von den anderen (kein geteilter Code – bewusst).

## Struktur

```
verlag_automations/
├── booxpress_etiketten/        ← ein Tool (BOOXpress-Versandetiketten)
│   ├── core.py                 ← reine Logik, ohne UI
│   ├── app.py                  ← Tkinter-GUI
│   ├── main.py                 ← Einstiegspunkt für den .exe-Build
│   ├── BooxpressEtiketten.spec ← PyInstaller-Build-Rezept
│   ├── README.md               ← Details zu diesem Tool
│   ├── samples/                ← Beispiel-Eingaben (nur zur Referenz)
│   └── data/                   ← Laufzeitdaten neben der .exe
│       ├── config.json         (wird bei Bedarf automatisch erzeugt)
│       ├── paketnr.txt         (Zähler)
│       └── kommliste.xlsx      (Stammdaten)
├── tools/
│   └── update_and_build.ps1    ← zieht Git-Änderungen & baut alle Tools neu
├── requirements.txt            ← gemeinsame Python-Abhängigkeiten
└── README.md
```

## Ein neues Tool hinzufügen

1. Neuen Ordner anlegen, z. B. `mein_tool/` mit `__init__.py`, `core.py`,
   `app.py`, `main.py` (nach dem Muster von `booxpress_etiketten/`).
2. `main.py` als Einstiegspunkt: `from mein_tool.app import main`.
3. Laufzeitdaten immer in `mein_tool/data/` ablegen (siehe `core.py`:
   `APP_DIR = _base_dir() / "data"`).
4. Eine eigene `MeinTool.spec` anlegen (Kopie der vorhandenen anpassen:
   Entry-Skript, `name=` und ggf. `collect_all(...)`).
5. Fertig – `tools/update_and_build.ps1` findet die neue `.spec` automatisch.

## Voraussetzungen

- Python 3.12 **inkl. tkinter** (unter Debian/Ubuntu: `sudo apt install python3-tk`)
- Abhängigkeiten: `pip install -r requirements.txt`

## Bauen

Einzelnes Tool (aus dem Repo-Wurzelordner):

```
pyinstaller booxpress_etiketten/BooxpressEtiketten.spec
```

Alle Tools auf einmal (Windows, holt vorher Git-Änderungen):

```
.\tools\update_and_build.ps1
```

Die fertigen Programme liegen anschließend unter `dist/<Name>/`.
