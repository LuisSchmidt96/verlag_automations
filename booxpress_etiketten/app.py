#!/usr/bin/env python3
"""
BOOXpress-Etiketten-Generator (Tkinter-GUI)
============================================

Erzeugt BOOXpress-Versandetiketten aus einem Lexware-Aufträge-Export.

Dateien im App-Ordner (neben der .exe):
- config.json    : Konstanten (Verlag-K-Nr, Adresse, Pfad zur Komm-Liste)
- paketnr.txt    : Fortlaufender Paketnummer-Counter (auto-hochgezählt)
- kommliste.xlsx : Aktuelle BOOXpress-Komm-Liste (nur zur Vorbelegung)
- etiketten_output/ : Ausgabeordner für die erzeugten Etiketten-docx

Bauen als .exe (Windows):
    pip install pyinstaller pandas openpyxl python-docx python-barcode pillow
    pyinstaller --windowed --noconfirm --name BooxpressEtiketten ^
        --collect-all barcode booxpress_app.py
"""

from __future__ import annotations

import os
import sys
import traceback
from datetime import datetime
from pathlib import Path
from tkinter import (Tk, Canvas, filedialog, messagebox,
                     StringVar, BooleanVar)
from tkinter import ttk
from tkinter.simpledialog import askinteger

from booxpress_etiketten import core


class App(Tk):
    def __init__(self):
        super().__init__()
        self.title("BOOXpress-Etiketten-Generator")
        self.geometry("1000x760")
        self.minsize(820, 560)

        try:
            self.cfg = core.lade_config()
        except Exception as e:
            messagebox.showerror("Config-Fehler",
                                 f"Konnte config.json nicht laden:\n{e}")
            self.cfg = dict(core.DEFAULT_CONFIG)

        self.paketnr_var = StringVar(value=str(core.lade_paketnr(self.cfg)))
        self.auftraege_pfad = StringVar(value="")
        # 8 Etikettenplätze auf dem Bogen (1=o.links … 8=u.rechts).
        # True = Platz ist frei und soll bedruckt werden.
        self.pos_vars = [BooleanVar(value=True) for _ in range(8)]
        # Eine Zeile je Auftrag (wird beim Laden gefüllt)
        self.auftrag_rows = []
        self.zaehler_var = StringVar(value="")
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
        ttk.Button(frm_file, text="Neu laden", command=self._lade_liste
                   ).pack(side="left", padx=(0, 8), pady=8)

        # Kopfbereich: Paketnummer + Bogen-Plätze nebeneinander
        top = ttk.Frame(self)
        top.pack(fill="x", **pad)

        # Paketnummer
        frm_pkt = ttk.LabelFrame(top, text="Nächste Paketnummer")
        frm_pkt.pack(side="left", fill="both", expand=True)
        inner = ttk.Frame(frm_pkt)
        inner.pack(fill="x", padx=8, pady=8)
        ttk.Label(inner, textvariable=self.paketnr_var,
                  font=("Segoe UI", 18, "bold")).pack(side="left")
        ttk.Button(inner, text="Zurücksetzen…",
                   command=self._reset_paketnr).pack(side="right")

        # Freie Etikettenplätze auf dem Bogen
        frm_grid = ttk.LabelFrame(
            top, text="Freie Plätze auf dem Bogen (angehakt = bedrucken)")
        frm_grid.pack(side="left", fill="both", expand=True, padx=(12, 0))
        grid_inner = ttk.Frame(frm_grid)
        grid_inner.pack(padx=8, pady=8)
        for i, var in enumerate(self.pos_vars):
            r, c = divmod(i, 2)
            ttk.Checkbutton(grid_inner, text=f"Platz {i + 1}", variable=var,
                            command=self._aktualisiere_zaehler
                            ).grid(row=r, column=c, sticky="w", padx=16, pady=2)
        ttk.Label(grid_inner,
                  text="(1 = oben links … 8 = unten rechts)",
                  foreground="gray").grid(row=4, column=0, columnspan=2,
                                          sticky="w", pady=(4, 0))
        grid_btns = ttk.Frame(frm_grid)
        grid_btns.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Button(grid_btns, text="Alle",
                   command=lambda: self._set_alle_plaetze(True)).pack(side="left")
        ttk.Button(grid_btns, text="Keine",
                   command=lambda: self._set_alle_plaetze(False)
                   ).pack(side="left", padx=6)

        # Aufträge-Liste (scrollbar, alle Aufträge, standardmäßig abgewählt)
        frm_auf = ttk.LabelFrame(
            self, text="Aufträge (anhaken zum Bedrucken)")
        frm_auf.pack(fill="both", expand=True, **pad)

        auf_top = ttk.Frame(frm_auf)
        auf_top.pack(fill="x", padx=8, pady=(6, 2))
        ttk.Label(auf_top, textvariable=self.zaehler_var,
                  font=("Segoe UI", 10, "bold")).pack(side="left")
        ttk.Button(auf_top, text="Keine",
                   command=lambda: self._set_alle_auftraege(False)
                   ).pack(side="right")
        ttk.Button(auf_top, text="Alle",
                   command=lambda: self._set_alle_auftraege(True)
                   ).pack(side="right", padx=6)

        auf_body = ttk.Frame(frm_auf)
        auf_body.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self.auf_canvas = Canvas(auf_body, highlightthickness=0)
        vs = ttk.Scrollbar(auf_body, orient="vertical",
                           command=self.auf_canvas.yview)
        self.auf_liste = ttk.Frame(self.auf_canvas)
        self.auf_liste.bind(
            "<Configure>",
            lambda e: self.auf_canvas.configure(
                scrollregion=self.auf_canvas.bbox("all")))
        self.auf_canvas.create_window((0, 0), window=self.auf_liste,
                                      anchor="nw")
        self.auf_canvas.configure(yscrollcommand=vs.set)
        self.auf_canvas.pack(side="left", fill="both", expand=True)
        vs.pack(side="right", fill="y")
        self._binde_mausrad(self.auf_canvas)

        self._platzhalter_zeile("Noch keine Aufträge geladen — "
                                "bitte eine Aufträge-Datei auswählen.")

        # Run
        self.btn_run = ttk.Button(self, text="Etiketten erstellen",
                                  command=self._run)
        self.btn_run.pack(pady=10)

        self._aktualisiere_zaehler()

    # -- Aufträge-Liste ------------------------------------------------

    def _binde_mausrad(self, canvas):
        def _scroll(event):
            if event.num == 4 or event.delta > 0:
                canvas.yview_scroll(-1, "units")
            elif event.num == 5 or event.delta < 0:
                canvas.yview_scroll(1, "units")
        canvas.bind("<Enter>", lambda e: (
            canvas.bind_all("<MouseWheel>", _scroll),
            canvas.bind_all("<Button-4>", _scroll),
            canvas.bind_all("<Button-5>", _scroll)))
        canvas.bind("<Leave>", lambda e: (
            canvas.unbind_all("<MouseWheel>"),
            canvas.unbind_all("<Button-4>"),
            canvas.unbind_all("<Button-5>")))

    def _leere_liste(self):
        for child in self.auf_liste.winfo_children():
            child.destroy()
        self.auftrag_rows = []

    def _platzhalter_zeile(self, text):
        self._leere_liste()
        ttk.Label(self.auf_liste, text=text,
                  foreground="gray").grid(row=0, column=0, padx=8, pady=12,
                                          sticky="w")

    def _baue_auftrag_zeilen(self, kandidaten):
        self._leere_liste()

        kopf = ["", "Match / Beleg", "Empfänger", "Straße", "PLZ / Ort",
                "Kundennr. (VD)"]
        for c, text in enumerate(kopf):
            ttk.Label(self.auf_liste, text=text,
                      font=("Segoe UI", 9, "bold")).grid(
                row=0, column=c, sticky="w", padx=4, pady=(0, 4))

        for i, et in enumerate(kandidaten, start=1):
            sel = BooleanVar(value=False)   # standardmäßig abgewählt
            name = StringVar(value=et.name.replace("\n", " | "))
            strasse = StringVar(value=et.strasse)
            plz_ort = StringVar(value=et.plz_ort)
            vd = StringVar(value=et.vd)

            ttk.Checkbutton(self.auf_liste, variable=sel,
                            command=self._aktualisiere_zaehler
                            ).grid(row=i, column=0, padx=4, pady=1)

            marker = "" if et.gematcht else "⚠ "
            ref = f"{marker}{et.matchcode}  ·  {et.belegnr}"
            ttk.Label(self.auf_liste, text=ref).grid(
                row=i, column=1, sticky="w", padx=4, pady=1)

            ttk.Entry(self.auf_liste, textvariable=name, width=26).grid(
                row=i, column=2, padx=4, pady=1)
            ttk.Entry(self.auf_liste, textvariable=strasse, width=22).grid(
                row=i, column=3, padx=4, pady=1)
            ttk.Entry(self.auf_liste, textvariable=plz_ort, width=18).grid(
                row=i, column=4, padx=4, pady=1)
            ttk.Entry(self.auf_liste, textvariable=vd, width=12).grid(
                row=i, column=5, padx=4, pady=1)

            self.auftrag_rows.append({
                "sel": sel, "name": name, "strasse": strasse,
                "plz_ort": plz_ort, "vd": vd,
                "matchcode": et.matchcode, "belegnr": et.belegnr})

        self.auf_canvas.yview_moveto(0)
        self._aktualisiere_zaehler()

    def _lade_liste(self):
        pfad_str = self.auftraege_pfad.get().strip()
        if not pfad_str:
            return
        auftraege_pfad = Path(pfad_str)
        if not auftraege_pfad.exists():
            self._platzhalter_zeile("Aufträge-Datei nicht gefunden.")
            messagebox.showwarning(
                "Hinweis", f"Aufträge-Datei nicht gefunden:\n{auftraege_pfad}")
            return

        try:
            self._log(f"Lade Aufträge: {auftraege_pfad.name}")
            auftraege = core.lade_auftraege(auftraege_pfad)

            kommliste_pfad = core.APP_DIR / self.cfg["kommliste_pfad"]
            if kommliste_pfad.exists():
                kommliste = core.lade_kommliste(kommliste_pfad)
                self._log(f"  Komm-Liste zur Vorbelegung: "
                          f"{len(kommliste)} Einträge")
            else:
                kommliste = None
                self._log("  ⚠ Keine Komm-Liste — Adressen manuell ergänzen.")

            kandidaten = core.generiere_etiketten(auftraege, kommliste, self.cfg)
        except Exception as e:
            self._log(f"\n❌ FEHLER beim Laden:\n{traceback.format_exc()}")
            messagebox.showerror("Fehler", str(e))
            self._platzhalter_zeile("Laden fehlgeschlagen — siehe Verlauf.")
            return

        if not kandidaten:
            self._platzhalter_zeile("Keine Aufträge in der Datei gefunden.")
            self._log("  Keine Aufträge gefunden.")
            return

        self._baue_auftrag_zeilen(kandidaten)
        ungematcht = sum(1 for k in kandidaten if not k.gematcht)
        self._log(f"  {len(kandidaten)} Aufträge geladen "
                  f"({ungematcht} ohne Komm-Listen-Treffer, ⚠ markiert).")

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
            self._lade_liste()

    def _reset_paketnr(self):
        neu = askinteger("Paketnummer zurücksetzen",
                         "Neuer Startwert für die Paketnummer:",
                         initialvalue=core.lade_paketnr(self.cfg), minvalue=1)
        if neu is not None:
            core.speichere_paketnr(neu, self.cfg)
            self.paketnr_var.set(str(neu))
            self._log(f"Paketnummer manuell auf {neu} gesetzt.")

    def _set_alle_plaetze(self, wert: bool):
        for var in self.pos_vars:
            var.set(wert)
        self._aktualisiere_zaehler()

    def _set_alle_auftraege(self, wert: bool):
        for r in self.auftrag_rows:
            r["sel"].set(wert)
        self._aktualisiere_zaehler()

    def _freie_plaetze(self):
        return [i + 1 for i, v in enumerate(self.pos_vars) if v.get()]

    def _aktualisiere_zaehler(self):
        n = sum(1 for r in self.auftrag_rows if r["sel"].get())
        k = len(self._freie_plaetze())
        txt = f"Ausgewählt: {n} von {k} freien Plätzen"
        if n > k:
            txt += "   ⚠ zu viele!"
        self.zaehler_var.set(txt)

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
        self._log("\n" + "─" * 60)
        self._log(f"Start: {datetime.now():%Y-%m-%d %H:%M:%S}")

        if not self.auftrag_rows:
            messagebox.showwarning(
                "Hinweis", "Bitte zuerst eine Aufträge-Datei laden.")
            return

        positionen = self._freie_plaetze()
        if not positionen:
            messagebox.showwarning(
                "Hinweis", "Bitte mindestens einen Etikettenplatz auswählen.")
            return

        ausgewaehlt = [r for r in self.auftrag_rows if r["sel"].get()]
        if not ausgewaehlt:
            messagebox.showwarning(
                "Hinweis", "Bitte mindestens einen Auftrag anhaken.")
            return
        if len(ausgewaehlt) > len(positionen):
            messagebox.showerror(
                "Zu viele Etiketten",
                f"{len(ausgewaehlt)} Aufträge ausgewählt, aber nur "
                f"{len(positionen)} Plätze auf dem Bogen frei.\n\n"
                f"Bitte weniger Aufträge anhaken oder mehr Plätze freigeben.")
            return

        # Editierte Felder einsammeln + Kundennr. prüfen
        etiketten = []
        for r in ausgewaehlt:
            vd = r["vd"].get().strip()
            if not vd.isdigit() or not 1 <= len(vd) <= 6:
                messagebox.showerror(
                    "Ungültige Kundennummer",
                    f"Kundennr. (VD) muss 1–6 Ziffern sein.\n"
                    f"Ungültig: '{vd}' "
                    f"(Match {r['matchcode']}, Beleg {r['belegnr']}).")
                return
            name = r["name"].get().replace(" | ", "\n").strip()
            etiketten.append(core.Etikett(
                name, r["strasse"].get().strip(),
                r["plz_ort"].get().strip(), vd, r["belegnr"]))

        start_pkt = core.lade_paketnr(self.cfg)
        self._log(f"Start-Paketnummer: {start_pkt}")
        core.setze_barcodes(etiketten, self.cfg, start_pkt)

        output_dir = core.APP_DIR / self.cfg["output_dir"]
        output_dir.mkdir(parents=True, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_pfad = output_dir / f"Etiketten_{ts}.docx"
        self._log(f"Bedruckte Plätze: {positionen[:len(etiketten)]}")
        self._log(f"Schreibe docx: {output_pfad.name}")
        core.baue_docx(etiketten, output_pfad, self.cfg, positionen)

        neu = start_pkt + len(etiketten)
        core.speichere_paketnr(neu, self.cfg)
        self.paketnr_var.set(str(neu))
        self._log(f"\n✓ Fertig. {len(etiketten)} Etiketten. "
                  f"Nächste Paketnummer: {neu}")
        self._log(f"  Datei: {output_pfad}")

        if messagebox.askyesno("Fertig",
                               f"{len(etiketten)} Etiketten erstellt.\n\n"
                               f"Datei jetzt öffnen?"):
            self._datei_oeffnen(output_pfad)

    # -- Helpers ------------------------------------------------------

    def _log(self, msg: str):
        # Verlauf-Panel entfernt — Ausgabe nur noch auf der Konsole (Debug).
        # Wichtige Meldungen laufen ohnehin über messagebox.
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
