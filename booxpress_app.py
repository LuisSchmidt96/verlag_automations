#!/usr/bin/env python3
"""
BOOXpress-Etiketten-Generator (Tkinter-GUI)
============================================

Erzeugt BOOXpress-Versandetiketten aus einem Lexware-Aufträge-Export.

Dateien im App-Ordner (neben der .exe):
- config.json    : Konstanten (Verlag-K-Nr, Adresse, Pfad zur Komm-Liste)
- paketnr.txt    : Fortlaufender Paketnummer-Counter (auto-hochgezählt)
- kommliste.xlsx : Aktuelle BOOXpress-Komm-Liste (manuell aktualisieren)
- etiketten_output/ : Ausgabeordner für die erzeugten Etiketten-docx

Bauen als .exe (Windows):
    pip install pyinstaller pandas openpyxl python-docx python-barcode pillow
    pyinstaller --windowed --noconfirm --name BooxpressEtiketten ^
        --collect-all barcode booxpress_app.py
"""

from __future__ import annotations

import os
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path
from tkinter import Tk, filedialog, messagebox, scrolledtext, StringVar
from tkinter import ttk
from tkinter.simpledialog import askinteger

import booxpress_core as core


class App(Tk):
    def __init__(self):
        super().__init__()
        self.title("BOOXpress-Etiketten-Generator")
        self.geometry("760x580")
        self.minsize(680, 480)

        try:
            self.cfg = core.lade_config()
        except Exception as e:
            messagebox.showerror("Config-Fehler",
                                 f"Konnte config.json nicht laden:\n{e}")
            self.cfg = dict(core.DEFAULT_CONFIG)

        self.paketnr_var = StringVar(value=str(core.lade_paketnr()))
        self.auftraege_pfad = StringVar(value="")
        self.startpos_var = StringVar(value="1")
        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 12, "pady": 6}

        # Header
        header = ttk.Frame(self)
        header.pack(fill="x", **pad)
        ttk.Label(header, text="BOOXpress-Etiketten-Generator",
                  font=("Segoe UI", 14, "bold")).pack(side="left")

        # Aufträge-Datei
        frm_file = ttk.LabelFrame(self, text="Lexware-Aufträge-Datei (xlsx)")
        frm_file.pack(fill="x", **pad)
        ttk.Entry(frm_file, textvariable=self.auftraege_pfad
                  ).pack(side="left", fill="x", expand=True, padx=8, pady=8)
        ttk.Button(frm_file, text="Auswählen…", command=self._waehle_datei
                   ).pack(side="left", padx=8, pady=8)

        # Paketnummer
        frm_pkt = ttk.LabelFrame(self, text="Nächste Paketnummer")
        frm_pkt.pack(fill="x", **pad)
        inner = ttk.Frame(frm_pkt)
        inner.pack(fill="x", padx=8, pady=8)
        ttk.Label(inner, textvariable=self.paketnr_var,
                  font=("Segoe UI", 18, "bold")).pack(side="left")
        ttk.Label(inner,
                  text="  (wird nach jedem Lauf automatisch hochgezählt)",
                  foreground="gray").pack(side="left")
        ttk.Button(inner, text="Zurücksetzen…",
                   command=self._reset_paketnr).pack(side="right")

        # Startposition auf dem Bogen
        frm_start = ttk.LabelFrame(self,
                                   text="Startposition auf dem Bogen (1–8)")
        frm_start.pack(fill="x", **pad)
        inner_s = ttk.Frame(frm_start)
        inner_s.pack(fill="x", padx=8, pady=8)
        ttk.Spinbox(inner_s, from_=1, to=8, width=5,
                    textvariable=self.startpos_var,
                    state="readonly").pack(side="left")
        ttk.Label(inner_s,
                  text="  (1 = oben links … 8 = unten rechts; "
                       "für teilbedruckte Bögen)",
                  foreground="gray").pack(side="left")

        # Run
        self.btn_run = ttk.Button(self, text="Etiketten erstellen",
                                  command=self._run_async)
        self.btn_run.pack(pady=12)

        # Log
        frm_log = ttk.LabelFrame(self, text="Verlauf")
        frm_log.pack(fill="both", expand=True, **pad)
        self.log = scrolledtext.ScrolledText(frm_log, height=12,
                                             font=("Consolas", 10),
                                             state="disabled", wrap="word")
        self.log.pack(fill="both", expand=True, padx=8, pady=8)

        self._log(f"App-Ordner: {core.APP_DIR}")
        self._log(f"Komm-Liste: {self.cfg['kommliste_pfad']}")
        if not (core.APP_DIR / self.cfg["kommliste_pfad"]).exists():
            self._log("⚠ Komm-Liste nicht gefunden — bitte aktuelle "
                      "BOOXpress-Komm-Liste in den App-Ordner kopieren.")

    # -- Aktionen -----------------------------------------------------

    def _waehle_datei(self):
        last_dir = self.cfg.get("last_input_dir", "")
        initialdir = last_dir if last_dir and Path(last_dir).is_dir() else None
        pfad = filedialog.askopenfilename(
            title="Lexware-Aufträge-Export auswählen",
            filetypes=[("Excel-Dateien", "*.xlsx *.xls"),
                       ("Alle Dateien", "*.*")],
            initialdir=initialdir,
        )
        if pfad:
            self.auftraege_pfad.set(pfad)
            neu_dir = str(Path(pfad).parent)
            if neu_dir != self.cfg.get("last_input_dir"):
                self.cfg["last_input_dir"] = neu_dir
                try:
                    core.speichere_config(self.cfg)
                except Exception as e:
                    self._log(f"⚠ Konnte last_input_dir nicht speichern: {e}")

    def _reset_paketnr(self):
        neu = askinteger("Paketnummer zurücksetzen",
                         "Neuer Startwert für die Paketnummer:",
                         initialvalue=core.lade_paketnr(), minvalue=1)
        if neu is not None:
            core.speichere_paketnr(neu)
            self.paketnr_var.set(str(neu))
            self._log(f"Paketnummer manuell auf {neu} gesetzt.")

    # -- Worker (in Thread, damit UI nicht einfriert) ------------------

    def _run_async(self):
        self.btn_run.configure(state="disabled")
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        try:
            self._do_run()
        except Exception as e:
            self._log(f"\n❌ FEHLER:\n{traceback.format_exc()}")
            messagebox.showerror("Fehler", str(e))
        finally:
            self.btn_run.configure(state="normal")

    def _do_run(self):
        self._log("\n" + "─" * 60)
        self._log(f"Start: {datetime.now():%Y-%m-%d %H:%M:%S}")

        auftraege_str = self.auftraege_pfad.get().strip()
        if not auftraege_str:
            messagebox.showwarning("Hinweis", "Bitte erst eine Aufträge-Datei auswählen.")
            return
        auftraege_pfad = Path(auftraege_str)
        if not auftraege_pfad.exists():
            raise FileNotFoundError(f"Aufträge-Datei nicht gefunden: {auftraege_pfad}")

        kommliste_pfad = core.APP_DIR / self.cfg["kommliste_pfad"]
        if not kommliste_pfad.exists():
            raise FileNotFoundError(
                f"Komm-Liste nicht gefunden: {kommliste_pfad}\n\n"
                f"Bitte aktuelle BOOXpress-Komm-Liste als "
                f"'{self.cfg['kommliste_pfad']}' in den App-Ordner kopieren."
            )

        self._log(f"Lade Aufträge: {auftraege_pfad.name}")
        auftraege = core.lade_auftraege(auftraege_pfad)
        self._log(f"  {len(auftraege)} Aufträge")

        self._log(f"Lade Komm-Liste: {kommliste_pfad.name}")
        kommliste = core.lade_kommliste(kommliste_pfad)
        self._log(f"  {len(kommliste)} BOOXpress-Empfänger")

        start_pkt = core.lade_paketnr()
        self._log(f"Start-Paketnummer: {start_pkt}")
        etiketten, warnungen = core.generiere_etiketten(
            auftraege, kommliste, self.cfg, start_paketnr=start_pkt
        )
        self._log(f"  → {len(etiketten)} Etiketten erzeugt")
        self._log(f"  → {len(warnungen)} übersprungen")

        if warnungen:
            self._log("\nÜbersprungene Aufträge (vmtl. GLS oder Endkunden):")
            for w in warnungen:
                self._log(f"  Beleg {w['belegnr']:>8}  Kd {w['kd_nr']:>7}  "
                          f"{w['matchcode']:<35}  {w['grund']}")

        if not etiketten:
            self._log("\nKeine BOOXpress-Etiketten zu erzeugen.")
            return

        try:
            start_pos = int(self.startpos_var.get())
        except ValueError:
            start_pos = 1
        start_pos = max(1, min(8, start_pos))

        output_dir = core.APP_DIR / self.cfg["output_dir"]
        output_dir.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_pfad = output_dir / f"Etiketten_{ts}.docx"
        self._log(f"Startposition auf Bogen: {start_pos}")
        self._log(f"Schreibe docx: {output_pfad.name}")
        core.baue_docx(etiketten, output_pfad, self.cfg,
                       start_position=start_pos)

        neu = start_pkt + len(etiketten)
        core.speichere_paketnr(neu)
        self.paketnr_var.set(str(neu))
        self._log(f"\n✓ Fertig. Nächste Paketnummer: {neu}")
        self._log(f"  Datei: {output_pfad}")

        if messagebox.askyesno("Fertig",
                               f"{len(etiketten)} Etiketten erstellt.\n\n"
                               f"Datei jetzt öffnen?"):
            self._datei_oeffnen(output_pfad)

    # -- Helpers ------------------------------------------------------

    def _log(self, msg: str):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")
        self.update_idletasks()

    @staticmethod
    def _datei_oeffnen(pfad: Path):
        if sys.platform == "win32":
            os.startfile(str(pfad))
        elif sys.platform == "darwin":
            os.system(f'open "{pfad}"')
        else:
            os.system(f'xdg-open "{pfad}"')


if __name__ == "__main__":
    App().mainloop()
