#!/usr/bin/env python3
"""
Cover-Previews-Generator (Tkinter-GUI)
======================================

Erzeugt aus einem druckfertigen Umschlag-PDF 2D- und 3D-Vorschau-PNGs.

Ablauf in der GUI:
1. Umschlag-PDF wählen  -> Vorschau mit erkannten Schnittlinien
2. Linien bei Bedarf mit der Maus nachjustieren  (blau = Schnitt)
3. Ausgaben & (für 3D) Mockup-PSD + Smart-Object-Ebene wählen
4. "Erstellen"  -> 2D sofort (Python), 3D über Photoshop (COM)

Dateien/Ordner (neben der .exe):
- data/config.json   : Einstellungen (wird beim ersten Start angelegt)
- data/cover_output/ : Standard-Ausgabeordner
"""

from __future__ import annotations

import os
import sys
import threading
import traceback
from pathlib import Path
from tkinter import (Tk, Canvas, filedialog, messagebox,
                     StringVar, BooleanVar)
from tkinter import ttk

from PIL import Image, ImageTk

from cover_previews import core

PREVIEW_MAX_W = 820          # Breite der Vorschau in Pixeln
PREVIEW_RENDER_DPI = 96      # Rasterung fürs Vorschaubild


class App(Tk):
    def __init__(self):
        super().__init__()
        self.title("Cover-Previews-Generator")
        self.geometry("1040x820")
        self.minsize(900, 640)

        try:
            self.cfg = core.lade_config()
        except Exception as e:
            messagebox.showerror("Config-Fehler",
                                 f"Konnte config.json nicht laden:\n{e}")
            self.cfg = dict(core.DEFAULT_CONFIG)

        self.doc = None                 # fitz.Document
        self.reg = None                 # core.Regionen
        self.sc = "unbekannt"           # Kurzcode aus ISBN
        self._preview_img = None        # ImageTk-Referenz (sonst GC)
        self._scale = 1.0               # Faktor pt -> Canvas-Pixel
        self._drag = None               # (orientierung, index) beim Ziehen

        self.pdf_pfad = StringVar()
        self.psd_pfad = StringVar(value=self.cfg.get("mockup_psd_pfad", ""))
        self.layer_name = StringVar(
            value=(self.cfg.get("mockup_slots") or [{}])[0].get("layer", "COVER"))
        self.out_2d = BooleanVar(value=True)
        self.out_3d = BooleanVar(value=True)
        self.info_var = StringVar(value="Bitte ein Umschlag-PDF wählen.")

        self._build_ui()

    # ------------------------------------------------------------------
    # UI-Aufbau
    # ------------------------------------------------------------------
    def _build_ui(self):
        pad = {"padx": 10, "pady": 5}

        head = ttk.Frame(self)
        head.pack(fill="x", **pad)
        ttk.Label(head, text="Cover-Previews-Generator",
                  font=("Segoe UI", 14, "bold")).pack(side="left")

        # PDF-Auswahl
        frm_pdf = ttk.LabelFrame(self, text="Umschlag-PDF (druckfertig)")
        frm_pdf.pack(fill="x", **pad)
        ttk.Entry(frm_pdf, textvariable=self.pdf_pfad).pack(
            side="left", fill="x", expand=True, padx=8, pady=8)
        ttk.Button(frm_pdf, text="Auswählen…", command=self._waehle_pdf).pack(
            side="left", padx=(0, 4), pady=8)
        ttk.Button(frm_pdf, text="Neu erkennen", command=self._erkenne).pack(
            side="left", padx=(0, 8), pady=8)

        # Vorschau
        frm_prev = ttk.LabelFrame(
            self, text="Vorschau — blaue Linien mit der Maus justieren "
                       "(grün=Rückseite, orange=Rücken, rot=Vorderseite)")
        frm_prev.pack(fill="both", expand=True, **pad)
        self.canvas = Canvas(frm_prev, background="#2b2b2b",
                             highlightthickness=0, height=430)
        self.canvas.pack(fill="both", expand=True, padx=6, pady=6)
        self.canvas.bind("<Button-1>", self._greife_linie)
        self.canvas.bind("<B1-Motion>", self._ziehe_linie)
        self.canvas.bind("<ButtonRelease-1>", self._lasse_linie)

        # Ausgaben + Mockup
        bottom = ttk.Frame(self)
        bottom.pack(fill="x", **pad)

        frm_out = ttk.LabelFrame(bottom, text="Ausgaben")
        frm_out.pack(side="left", fill="both", expand=True)
        ttk.Checkbutton(frm_out, text="2D-Vorderseite",
                        variable=self.out_2d).pack(anchor="w", padx=8, pady=(6, 0))
        ttk.Checkbutton(frm_out, text="3D-Mockup (Photoshop)",
                        variable=self.out_3d).pack(anchor="w", padx=8, pady=(0, 6))
        ttk.Label(frm_out,
                  text=f"Auflösungen: {self.cfg['dpi_print']} dpi (Druck) + "
                       f"{self.cfg['dpi_web']} dpi (Web, ~{self.cfg['web_max_px']} px)",
                  foreground="gray").pack(anchor="w", padx=8, pady=(0, 6))

        frm_ps = ttk.LabelFrame(bottom, text="3D-Mockup (Photoshop-PSD)")
        frm_ps.pack(side="left", fill="both", expand=True, padx=(10, 0))
        row = ttk.Frame(frm_ps)
        row.pack(fill="x", padx=8, pady=(8, 2))
        ttk.Entry(row, textvariable=self.psd_pfad).pack(
            side="left", fill="x", expand=True)
        ttk.Button(row, text="…", width=3, command=self._waehle_psd).pack(
            side="left", padx=(4, 0))
        row2 = ttk.Frame(frm_ps)
        row2.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Label(row2, text="Smart-Object-Ebene:").pack(side="left")
        ttk.Entry(row2, textvariable=self.layer_name, width=18).pack(
            side="left", padx=6)

        # Aktionen
        act = ttk.Frame(self)
        act.pack(fill="x", **pad)
        ttk.Label(act, textvariable=self.info_var).pack(side="left")
        self.btn_run = ttk.Button(act, text="Erstellen", command=self._run)
        self.btn_run.pack(side="right")
        ttk.Button(act, text="Ausgabeordner öffnen",
                   command=self._oeffne_ausgabe).pack(side="right", padx=6)

    # ------------------------------------------------------------------
    # PDF laden + Erkennung
    # ------------------------------------------------------------------
    def _waehle_pdf(self):
        start = self.cfg.get("last_input_dir") or None
        pfad = filedialog.askopenfilename(
            title="Umschlag-PDF wählen", initialdir=start,
            filetypes=[("PDF-Dateien", "*.pdf"), ("Alle Dateien", "*.*")])
        if pfad:
            self.pdf_pfad.set(pfad)
            self.cfg["last_input_dir"] = str(Path(pfad).parent)
            core.speichere_config(self.cfg)
            self._erkenne()

    def _erkenne(self):
        pfad = self.pdf_pfad.get().strip()
        if not pfad:
            return
        try:
            self.doc = core.oeffne_pdf(pfad)
            self.reg = core.finde_schnittlinien(self.doc, self.cfg)
            isbn = core.extrahiere_isbn(self.doc)
            self.sc = core.shortcode_aus_isbn(isbn) if isbn else \
                Path(pfad).stem
            self._zeichne_vorschau()
            n = len(self.reg.x_cuts)
            if self.reg.front is None:
                self.info_var.set(
                    f"Nur {n} vertikale Marken erkannt — Linien bitte prüfen.")
            else:
                self.info_var.set(f"Erkannt. Kurzcode: {self.sc}. "
                                  "Linien bei Bedarf justieren, dann Erstellen.")
        except Exception as e:
            messagebox.showerror("Fehler beim Einlesen",
                                 f"{e}\n\n{traceback.format_exc()}")

    def _zeichne_vorschau(self):
        img = core.rendere_seite(self.doc, PREVIEW_RENDER_DPI)
        W_px = img.width
        # Vorschau auf PREVIEW_MAX_W begrenzen; _pt2px = Canvas-Pixel je Punkt.
        disp_w = min(W_px, PREVIEW_MAX_W)
        disp = img.resize((disp_w, int(img.height * disp_w / img.width)),
                          Image.LANCZOS)
        self._pt2px = disp_w / self.reg.seite_pt[0]     # Punkt -> Canvas-Pixel
        self._preview_img = ImageTk.PhotoImage(disp)
        self.canvas.configure(width=disp.width, height=disp.height)
        self._redraw()

    def _redraw(self):
        c = self.canvas
        c.delete("all")
        if self._preview_img is None:
            return
        c.create_image(0, 0, anchor="nw", image=self._preview_img)
        s = self._pt2px
        # Regionen-Rechtecke
        farben = {"back": "#00c800", "spine": "#ff8c00", "front": "#dc0000"}
        for name, col in farben.items():
            box = getattr(self.reg, name)
            if box:
                c.create_rectangle(box[0]*s, box[1]*s, box[2]*s, box[3]*s,
                                   outline=col, width=3)
        # Schnittlinien (blau, ziehbar)
        h = self._preview_img.height()
        w = self._preview_img.width()
        for i, x in enumerate(self.reg.x_cuts):
            c.create_line(x*s, 0, x*s, h, fill="#1e9bff", width=2,
                          tags=("vcut", f"v{i}"))
        for i, y in enumerate(self.reg.y_cuts):
            c.create_line(0, y*s, w, y*s, fill="#1e9bff", width=2,
                          tags=("hcut", f"h{i}"))

    # ------------------------------------------------------------------
    # Linien ziehen
    # ------------------------------------------------------------------
    def _greife_linie(self, ev):
        if self.reg is None:
            return
        s = self._pt2px
        # nächste vertikale / horizontale Linie in Reichweite (8 px)
        best = None
        for i, x in enumerate(self.reg.x_cuts):
            d = abs(ev.x - x*s)
            if d < 8 and (best is None or d < best[2]):
                best = ("v", i, d)
        for i, y in enumerate(self.reg.y_cuts):
            d = abs(ev.y - y*s)
            if d < 8 and (best is None or d < best[2]):
                best = ("h", i, d)
        self._drag = best[:2] if best else None

    def _ziehe_linie(self, ev):
        if not self._drag:
            return
        orient, i = self._drag
        s = self._pt2px
        if orient == "v":
            self.reg.x_cuts[i] = max(0, min(self.reg.seite_pt[0], ev.x / s))
        else:
            self.reg.y_cuts[i] = max(0, min(self.reg.seite_pt[1], ev.y / s))
        self.reg = core.regionen_aus_cuts(
            self.reg.x_cuts, self.reg.y_cuts, self.reg.seite_pt, self.cfg)
        self._redraw()

    def _lasse_linie(self, _ev):
        self._drag = None

    # ------------------------------------------------------------------
    # Mockup-Auswahl
    # ------------------------------------------------------------------
    def _waehle_psd(self):
        pfad = filedialog.askopenfilename(
            title="Mockup-PSD wählen",
            filetypes=[("Photoshop-Dateien", "*.psd *.psb"),
                       ("Alle Dateien", "*.*")])
        if pfad:
            self.psd_pfad.set(pfad)

    # ------------------------------------------------------------------
    # Erstellen
    # ------------------------------------------------------------------
    def _run(self):
        if self.doc is None or self.reg is None:
            messagebox.showwarning("Kein PDF", "Bitte zuerst ein PDF wählen.")
            return
        if self.reg.front is None:
            messagebox.showwarning(
                "Regionen unklar",
                "Vorder-/Rückseite/Rücken sind nicht bestimmt. Bitte die "
                "blauen Linien justieren (mind. 4 senkrechte Schnitte).")
            return
        # Config aus GUI übernehmen
        self.cfg["mockup_psd_pfad"] = self.psd_pfad.get().strip()
        slot_region = (self.cfg.get("mockup_slots") or
                       [{"region": "front_spine"}])[0].get("region", "front_spine")
        self.cfg["mockup_slots"] = [{"layer": self.layer_name.get().strip() or "COVER",
                                     "region": slot_region}]
        core.speichere_config(self.cfg)

        self.btn_run.configure(state="disabled")
        self.info_var.set("Erzeuge …")
        threading.Thread(target=self._run_worker, daemon=True).start()

    def _run_worker(self):
        erzeugt = []
        fehler = []
        try:
            out_dir = core.APP_DIR / self.cfg["output_dir"]
            img_hi = core.rendere_seite(self.doc, int(self.cfg["dpi_print"]))

            if self.out_2d.get():
                front = core.extrahiere(
                    img_hi, self.reg.box_px("front", int(self.cfg["dpi_print"])))
                erzeugt += core.speichere_2d(front, out_dir, self.sc, self.cfg)

            if self.out_3d.get():
                if not self.cfg["mockup_psd_pfad"]:
                    fehler.append("3D übersprungen: kein Mockup-PSD gewählt.")
                elif sys.platform != "win32":
                    # kein Photoshop/COM -> Dry-Run (Slot-PNG + JSX zum Prüfen)
                    erzeugt += core.erzeuge_3d_photoshop(
                        self.reg, img_hi, self.cfg, out_dir, self.sc,
                        dry_run=True, log=lambda m: None)
                    fehler.append("3D: kein Windows/Photoshop — Dry-Run "
                                  "(JSX + Slot-PNG) geschrieben.")
                else:
                    erzeugt += core.erzeuge_3d_photoshop(
                        self.reg, img_hi, self.cfg, out_dir, self.sc,
                        dry_run=False, log=lambda m: None)
        except Exception as e:
            fehler.append(f"{e}")
            traceback.print_exc()
        self.after(0, self._run_fertig, erzeugt, fehler)

    def _run_fertig(self, erzeugt, fehler):
        self.btn_run.configure(state="normal")
        if erzeugt:
            namen = "\n".join(f"• {Path(p).name}" for p in erzeugt)
            self.info_var.set(f"{len(erzeugt)} Datei(en) erstellt.")
            msg = f"Erstellt:\n{namen}"
            if fehler:
                msg += "\n\nHinweise:\n" + "\n".join(fehler)
            messagebox.showinfo("Fertig", msg)
        else:
            self.info_var.set("Nichts erstellt.")
            messagebox.showerror("Fehler", "\n".join(fehler) or "Unbekannt.")

    def _oeffne_ausgabe(self):
        out_dir = core.APP_DIR / self.cfg["output_dir"]
        out_dir.mkdir(parents=True, exist_ok=True)
        if sys.platform == "win32":
            os.startfile(str(out_dir))
        elif sys.platform == "darwin":
            os.system(f'open "{out_dir}"')
        else:
            os.system(f'xdg-open "{out_dir}"')


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
