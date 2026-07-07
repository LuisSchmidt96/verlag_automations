#!/usr/bin/env python3
"""
PI/BI-Generator (Tkinter-GUI)
=============================

Erzeugt aus einem VLB-ONIX-XML-Datensatz die Presse- (PI) und
Buchinformation (BI) je als docx und html in einem Ausgabeordner.

Dateien/Ordner (neben der .exe):
- vorlagen/   : mitgelieferte docx-/html-Vorlagen
- data/       : config.json (wird beim ersten Start angelegt)
- pi_bi_output/ : Standard-Ausgabeordner

Bauen als .exe (Windows, aus dem Repo-Wurzelordner):
    pyinstaller pi_bi_generator/PiBiGenerator.spec
"""

from __future__ import annotations

import os
import re
import sys
import traceback
from pathlib import Path
from tkinter import (Tk, filedialog, messagebox, StringVar)
from tkinter import ttk

from pi_bi_generator import core


# Anzeige-Reihenfolge der erkannten Felder in der Vorschau
_VORSCHAU_FELDER = [
    ("isbn13_formatiert", "ISBN"),
    ("shortcode", "Kurzcode"),
    ("titel", "Titel"),
    ("band_text", "Band"),
    ("editoren_slash", "Herausgeber"),
    ("mitwirkende", "Mit Beiträgen von"),
    ("umfang_zeile", "Umfang"),
    ("preis", "Preis"),
    ("datum", "Erscheinungsdatum"),
    ("cover_url", "Cover-URL (VLB)"),
]


def _sichere_ordnername(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*]', "_", name).strip().strip(".")
    return name or "PI_BI"


class App(Tk):
    def __init__(self):
        super().__init__()
        self.title("PI/BI-Generator")
        self.geometry("880x720")
        self.minsize(760, 600)

        try:
            self.cfg = core.lade_config()
        except Exception as e:
            messagebox.showerror("Config-Fehler",
                                 f"Konnte config.json nicht laden:\n{e}")
            self.cfg = dict(core.DEFAULT_CONFIG)

        self.buch = None

        self.xml_pfad = StringVar(value="")
        self.cover_pfad = StringVar(value="")
        self.ordner_var = StringVar(value="")
        self.detail_var = StringVar(value=self.cfg.get("detail_fallback_url", ""))
        self.output_var = StringVar(value=str(self._output_basis()))
        self.vorschau_vars = {feld: StringVar(value="")
                              for feld, _ in _VORSCHAU_FELDER}
        self._build_ui()

    # -- Pfade ---------------------------------------------------------

    def _output_basis(self) -> Path:
        raw = self.cfg.get("output_dir", "pi_bi_output")
        p = Path(os.path.expandvars(str(raw))).expanduser()
        if not p.is_absolute():
            p = core.APP_DIR / p
        return p

    # -- UI-Aufbau -----------------------------------------------------

    def _build_ui(self):
        pad = {"padx": 12, "pady": 6}

        header = ttk.Frame(self)
        header.pack(fill="x", **pad)
        ttk.Label(header, text="PI/BI-Generator",
                  font=("Segoe UI", 14, "bold")).pack(side="left")

        # XML-Datei
        frm_xml = ttk.LabelFrame(self, text="VLB-ONIX-XML")
        frm_xml.pack(fill="x", **pad)
        ttk.Entry(frm_xml, textvariable=self.xml_pfad, state="readonly"
                  ).pack(side="left", fill="x", expand=True, padx=8, pady=8)
        ttk.Button(frm_xml, text="Auswählen…", command=self._waehle_xml
                   ).pack(side="left", padx=8, pady=8)

        # Erkannte Daten (Vorschau)
        frm_prev = ttk.LabelFrame(self, text="Erkannte Daten")
        frm_prev.pack(fill="x", **pad)
        inner = ttk.Frame(frm_prev)
        inner.pack(fill="x", padx=8, pady=8)
        inner.columnconfigure(1, weight=1)
        for r, (feld, label) in enumerate(_VORSCHAU_FELDER):
            ttk.Label(inner, text=label + ":", foreground="gray"
                      ).grid(row=r, column=0, sticky="ne", padx=(0, 8), pady=1)
            ttk.Label(inner, textvariable=self.vorschau_vars[feld],
                      wraplength=620, justify="left"
                      ).grid(row=r, column=1, sticky="w", pady=1)

        # Coverbild
        frm_cover = ttk.LabelFrame(self, text="Coverbild")
        frm_cover.pack(fill="x", **pad)
        ttk.Entry(frm_cover, textvariable=self.cover_pfad
                  ).pack(side="left", fill="x", expand=True, padx=8, pady=8)
        ttk.Button(frm_cover, text="Auswählen…", command=self._waehle_cover
                   ).pack(side="left", padx=(0, 4), pady=8)
        ttk.Button(frm_cover, text="Vom Webserver", command=self._cover_web
                   ).pack(side="left", padx=(0, 8), pady=8)

        # Ausgabe-Einstellungen
        frm_out = ttk.LabelFrame(self, text="Ausgabe")
        frm_out.pack(fill="x", **pad)
        grid = ttk.Frame(frm_out)
        grid.pack(fill="x", padx=8, pady=8)
        grid.columnconfigure(1, weight=1)

        ttk.Label(grid, text="Ordnername:").grid(row=0, column=0, sticky="w",
                                                 padx=(0, 8), pady=3)
        ttk.Entry(grid, textvariable=self.ordner_var).grid(
            row=0, column=1, columnspan=2, sticky="ew", pady=3)

        ttk.Label(grid, text="Produkt-Seite (Webshop):").grid(
            row=1, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(grid, textvariable=self.detail_var).grid(
            row=1, column=1, columnspan=2, sticky="ew", pady=3)

        ttk.Label(grid, text="Zielordner:").grid(row=2, column=0, sticky="w",
                                                 padx=(0, 8), pady=3)
        ttk.Label(grid, textvariable=self.output_var, foreground="gray"
                  ).grid(row=2, column=1, sticky="w", pady=3)
        ttk.Button(grid, text="Ändern…", command=self._waehle_output
                   ).grid(row=2, column=2, sticky="e", pady=3)

        # Aktion
        self.btn_run = ttk.Button(self, text="Dokumente erstellen",
                                  command=self._run)
        self.btn_run.pack(pady=14)

        ttk.Label(
            self,
            text="Hinweis: Cover-Thumbnail (…/newsletter_/{Kurzcode}.png) und "
                 "Blick-ins-Buch-PDF für die HTML-Fassung bitte separat auf "
                 "den Webserver hochladen.",
            foreground="gray", wraplength=820, justify="left"
        ).pack(fill="x", padx=12, pady=(0, 8))

    # -- Datei-Auswahl -------------------------------------------------

    def _waehle_xml(self):
        last_dir = self.cfg.get("last_input_dir", "")
        initialdir = last_dir if last_dir and Path(last_dir).is_dir() else None
        pfad = filedialog.askopenfilename(
            title="VLB-ONIX-XML auswählen",
            filetypes=[("ONIX-XML", "*.xml"), ("Alle Dateien", "*.*")],
            initialdir=initialdir,
        )
        if not pfad:
            return
        self.xml_pfad.set(pfad)
        neu_dir = str(Path(pfad).parent)
        if neu_dir != self.cfg.get("last_input_dir"):
            self.cfg["last_input_dir"] = neu_dir
            try:
                core.speichere_config(self.cfg)
            except Exception as e:
                self._log(f"⚠ Konnte last_input_dir nicht speichern: {e}")
        self._lade_vorschau(Path(pfad))

    def _lade_vorschau(self, pfad: Path):
        try:
            self.buch = core.lade_buchdaten(pfad, self.cfg)
        except Exception as e:
            self.buch = None
            for var in self.vorschau_vars.values():
                var.set("")
            self._log(f"\n❌ XML-Fehler:\n{traceback.format_exc()}")
            messagebox.showerror(
                "XML-Fehler", f"Konnte die XML nicht auswerten:\n{e}")
            return
        for feld, var in self.vorschau_vars.items():
            var.set(str(getattr(self.buch, feld, "") or ""))
        # Ordnername vorschlagen (Kurzcode), Cover auto-vorschlagen (lokal)
        self.ordner_var.set(self.buch.shortcode)
        self._suche_lokales_cover(pfad)

    def _suche_lokales_cover(self, xml_pfad: Path):
        """Neben der XML nach einem passenden Coverbild suchen."""
        if not self.buch:
            return
        sc, isbn = self.buch.shortcode, self.buch.isbn13
        kandidaten = [f"2D_300_{sc}.jpg", f"3D_300_{sc}.jpg",
                      f"{isbn}.jpg", f"{sc}.jpg", f"{sc}.png"]
        for name in kandidaten:
            p = xml_pfad.parent / name
            if p.exists():
                self.cover_pfad.set(str(p))
                return

    def _waehle_cover(self):
        last_dir = self.cfg.get("last_input_dir", "")
        initialdir = last_dir if last_dir and Path(last_dir).is_dir() else None
        pfad = filedialog.askopenfilename(
            title="Coverbild auswählen",
            filetypes=[("Bilder", "*.jpg *.jpeg *.png"), ("Alle Dateien", "*.*")],
            initialdir=initialdir,
        )
        if pfad:
            self.cover_pfad.set(pfad)

    def _cover_web(self):
        if not self.buch:
            messagebox.showwarning("Hinweis", "Bitte zuerst eine XML laden.")
            return
        url = core.web_cover_url(self.cfg, self.buch)
        try:
            daten = core.lade_cover_web(url, float(self.cfg.get("cover_timeout", 15)))
        except Exception as e:
            messagebox.showerror(
                "Cover-Download fehlgeschlagen",
                f"Konnte das Cover nicht laden von:\n{url}\n\n{e}")
            return
        cache = core.APP_DIR / "cover_cache"
        cache.mkdir(parents=True, exist_ok=True)
        ziel = cache / f"{self.buch.shortcode}.png"
        ziel.write_bytes(daten)
        self.cover_pfad.set(str(ziel))
        self._log(f"Cover vom Webserver geladen: {url}")

    def _waehle_output(self):
        pfad = filedialog.askdirectory(
            title="Zielordner (Basis) auswählen",
            initialdir=str(self._output_basis()) if self._output_basis().is_dir() else None,
        )
        if pfad:
            self.cfg["output_dir"] = pfad
            self.output_var.set(pfad)
            try:
                core.speichere_config(self.cfg)
            except Exception as e:
                self._log(f"⚠ Konnte output_dir nicht speichern: {e}")

    # -- Ablauf --------------------------------------------------------

    def _run(self):
        self.btn_run.configure(state="disabled")
        self.update_idletasks()
        try:
            self._do_run()
        except Exception as e:
            self._log(f"\n❌ FEHLER:\n{traceback.format_exc()}")
            messagebox.showerror("Fehler", str(e))
        finally:
            self.btn_run.configure(state="normal")

    def _do_run(self):
        if self.buch is None:
            messagebox.showwarning("Hinweis", "Bitte zuerst eine XML laden.")
            return

        ordnername = _sichere_ordnername(self.ordner_var.get())
        ziel = self._output_basis() / ordnername
        if ziel.exists() and any(ziel.iterdir()):
            if not messagebox.askyesno(
                    "Ordner existiert",
                    f"Der Ordner existiert bereits und ist nicht leer:\n{ziel}\n\n"
                    f"Vorhandene Dateien ggf. überschreiben?"):
                return

        # Cover laden (optional)
        cover_bytes = None
        cover_suffix = ".jpg"
        cover_str = self.cover_pfad.get().strip()
        if cover_str:
            cp = Path(cover_str)
            if cp.exists():
                cover_bytes = core.lade_cover_datei(cp)
                cover_suffix = cp.suffix or ".jpg"
            else:
                messagebox.showwarning(
                    "Cover nicht gefunden",
                    f"Coverdatei nicht gefunden:\n{cp}\n\n"
                    f"Es wird ohne aktuelles Cover erzeugt.")

        sc = self.buch.shortcode
        detail_url = self.detail_var.get().strip() or \
            self.cfg.get("detail_fallback_url", "")
        ziel.mkdir(parents=True, exist_ok=True)

        v = core.VORLAGEN_DIR
        core.generiere_docx(v / "pi_vorlage.docx", self.buch, cover_bytes,
                            ziel / f"PI_{sc}.docx")
        core.generiere_docx(v / "bi_vorlage.docx", self.buch, cover_bytes,
                            ziel / f"BI_{sc}.docx")
        core.generiere_html(v / "pi_vorlage.html", self.buch, detail_url,
                            self.cfg, ziel / f"PI {sc}.html")
        core.generiere_html(v / "bi_vorlage.html", self.buch, detail_url,
                            self.cfg, ziel / f"BI {sc}.html")

        erzeugt = 4
        if cover_bytes:
            (ziel / f"cover_{self.buch.isbn13}{cover_suffix}").write_bytes(cover_bytes)
            erzeugt += 1
        else:
            self._log("⚠ Kein Cover gewählt — docx behalten das Platzhalter-Cover.")

        self._log(f"✓ {erzeugt} Dateien in {ziel}")
        cover_hinweis = ("" if cover_bytes else
                         "\n\n⚠ Ohne aktuelles Cover erzeugt "
                         "(kein Coverbild gewählt).")
        if messagebox.askyesno(
                "Fertig",
                f"PI und BI wurden erstellt (docx + html){cover_hinweis}\n\n"
                f"Ordner jetzt öffnen?"):
            self._datei_oeffnen(ziel)

    # -- Helpers ------------------------------------------------------

    def _log(self, msg: str):
        print(msg)

    @staticmethod
    def _datei_oeffnen(pfad: Path):
        if sys.platform == "win32":
            os.startfile(str(pfad))
        elif sys.platform == "darwin":
            os.system(f'open "{pfad}"')
        else:
            os.system(f'xdg-open "{pfad}"')


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
