# verlag_automations

Sammlung eigenständiger Automations-Tools für den Verlag Regionalkultur.
Jedes Tool lebt in einem eigenen Unterordner, wird zu einer **eigenen .exe**
gebaut und ist unabhängig von den anderen (kein geteilter Code – bewusst).

## Struktur

```
verlag_automations/
├── booxpress_etiketten/        ← BOOXpress-Versandetiketten (Lexware → Etiketten-docx)
│   ├── core.py                 ← reine Logik, ohne UI
│   ├── app.py                  ← Tkinter-GUI
│   ├── main.py                 ← Einstiegspunkt für den .exe-Build
│   ├── BooxpressEtiketten.spec ← PyInstaller-Build-Rezept
│   ├── README.md               ← Details zu diesem Tool
│   └── samples/                ← Beispiel-Eingaben (nur zur Referenz)
│                                 Laufzeitdaten (config.json, paketnr.txt,
│                                 kommliste.xlsx) liegen neben der .exe
├── pi_bi_generator/            ← ein Tool (PI/BI aus VLB-ONIX-XML)
│   ├── core.py, app.py, main.py, PiBiGenerator.spec, README.md
│   ├── vorlagen/               ← docx-/html-Vorlagen (in die .exe gebündelt)
│   └── beispiele/              ← Beispiel-XML + Referenzdokumente
├── cover_previews/             ← Cover-Previews (Umschlag-PDF → 2D-/3D-Vorschau-PNGs)
│   ├── core.py, app.py, main.py, CoverPreviews.spec, README.md
│   └── mockups/                ← Photoshop-Mockup-PSDs (lokal, nicht eingecheckt)
├── tools/
│   └── update_and_build.ps1    ← zieht Git-Änderungen & baut alle Tools neu
├── requirements.txt            ← gemeinsame Python-Abhängigkeiten
└── README.md
```

Alle Tools folgen demselben Muster (`core.py` / `app.py` / `main.py` /
`<Name>.spec`); Laufzeitdaten liegen direkt neben der .exe.

## Ein neues Tool hinzufügen

1. Neuen Ordner anlegen, z. B. `mein_tool/` mit `__init__.py`, `core.py`,
   `app.py`, `main.py` (nach dem Muster von `booxpress_etiketten/`).
2. `main.py` als Einstiegspunkt: `from mein_tool.app import main`.
3. Laufzeitdaten (config.json, Ausgabeordner, …) immer direkt neben der .exe
   ablegen (siehe `core.py`: `APP_DIR = _base_dir()`) – kein `data/`-Unterordner,
   damit der Anwender sie sofort findet.
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
