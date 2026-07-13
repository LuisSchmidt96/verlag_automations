#!/usr/bin/env python3
"""
Cover-Previews-Generator (Tkinter-GUI)
======================================

Erzeugt aus einem druckfertigen Umschlag-PDF 2D- und 3D-Vorschau-PNGs.

Ablauf in der GUI:
1. Umschlag-PDF wählen  -> Vorschau mit erkannten Schnittlinien; aus der ISBN
   wird der Kurzcode und daraus der Zielordner auf dem Artikeldaten-Share
   gesucht (bzw. ein neuer vorgeschlagen)
2. Linien bei Bedarf mit der Maus nachjustieren  (blau = Schnitt)
3. Ausgaben, Einband (Hardcover = mit Falz) und Mockup-Vorlage wählen
4. "Erstellen"  -> 2D sofort (Python), 3D über Photoshop (COM)

Die Bilder landen in <Artikeldaten>/<Kurzcode>_<Titel>/. Liegen dort schon
gleichnamige Dateien, werden sie nach _alt/<Zeitstempel>/ weggesichert, statt
sie zu überschreiben.

config.json (Einstellungen) liegt direkt neben der .exe.
"""

from __future__ import annotations

import os
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path
from tkinter import (Tk, Canvas, filedialog, messagebox,
                     StringVar, BooleanVar)
from tkinter import ttk

from PIL import Image, ImageTk

from cover_previews import core

PREVIEW_MAX_W = 820          # max. Breite der Vorschau in Pixeln
PREVIEW_MAX_H = 460          # max. Höhe der Vorschau (damit die Bedienelemente
                             # unten sichtbar bleiben)
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
        self._out_dir = None            # Zielordner des laufenden Auftrags
        self._alt = []                  # Dateien, die vorher weggesichert werden
        self._preview_img = None        # ImageTk-Referenz (sonst GC)
        self._scale = 1.0               # Faktor pt -> Canvas-Pixel
        self._drag = None               # (orientierung, index) beim Ziehen

        self.pdf_pfad = StringVar()
        self.vorlagen_dir_var = StringVar(
            value=str(core.vorlagen_dir(self.cfg)))
        self.vorlage_var = StringVar()          # gewählte Vorlagen-Datei
        self.vorlage_info = StringVar(value="—")
        self.out_2d = BooleanVar(value=True)
        self.out_3d = BooleanVar(value=True)
        self.einband_var = StringVar(
            value=self.cfg.get("einband", "hardcover"))
        self.titel_var = StringVar()            # Titelteil des Ordnernamens
        self.ziel_var = StringVar(value="—")    # aufgelöster Zielordner
        self.ziel_status = StringVar(value="")  # vorhanden / wird angelegt
        self.info_var = StringVar(value="Bitte ein Umschlag-PDF wählen.")

        self._build_ui()
        self.titel_var.trace_add("write", lambda *_: self._aktualisiere_ziel())

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

        # Zielordner auf dem Artikeldaten-Share
        frm_ziel = ttk.LabelFrame(self, text="Zielordner (Artikeldaten)")
        frm_ziel.pack(fill="x", **pad)
        z1 = ttk.Frame(frm_ziel)
        z1.pack(fill="x", padx=8, pady=(8, 2))
        ttk.Label(z1, text="Ordnername:  Kurzcode +").pack(side="left")
        ttk.Entry(z1, textvariable=self.titel_var, width=28).pack(
            side="left", padx=6)
        ttk.Label(z1, text="(Titel, optional)", foreground="gray").pack(side="left")
        ttk.Label(z1, textvariable=self.ziel_status,
                  foreground="#0a7").pack(side="left", padx=12)
        z2 = ttk.Frame(frm_ziel)
        z2.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Label(z2, textvariable=self.ziel_var, foreground="gray").pack(
            side="left", fill="x", expand=True)

        # Bedienelemente unten zuerst verankern, damit sie nie verdeckt werden.
        unten = ttk.Frame(self)
        unten.pack(side="bottom", fill="x")

        # Vorschau füllt den verbleibenden Platz darüber.
        frm_prev = ttk.LabelFrame(
            self, text="Vorschau — blaue Linien mit der Maus justieren "
                       "(grün=Rückseite, orange=Rücken, rot=Vorderseite)")
        frm_prev.pack(side="top", fill="both", expand=True, **pad)
        self.canvas = Canvas(frm_prev, background="#2b2b2b",
                             highlightthickness=0, height=300)
        self.canvas.pack(fill="both", expand=True, padx=6, pady=6)
        self.canvas.bind("<Button-1>", self._greife_linie)
        self.canvas.bind("<B1-Motion>", self._ziehe_linie)
        self.canvas.bind("<ButtonRelease-1>", self._lasse_linie)

        # Ausgaben + Mockup
        bottom = ttk.Frame(unten)
        bottom.pack(fill="x", **pad)

        frm_out = ttk.LabelFrame(bottom, text="Ausgaben")
        frm_out.pack(side="left", fill="both", expand=True)
        ttk.Checkbutton(frm_out, text="2D-Vorderseite",
                        variable=self.out_2d).pack(anchor="w", padx=8, pady=(6, 0))
        ttk.Checkbutton(frm_out, text="3D-Mockup (Photoshop)",
                        variable=self.out_3d).pack(anchor="w", padx=8, pady=(0, 6))
        ttk.Label(frm_out,
                  text=f"2D als JPEG, 3D als PNG+JPEG — je {self.cfg['dpi_print']} "
                       f"& {self.cfg['dpi_web']} dpi",
                  foreground="gray").pack(anchor="w", padx=8, pady=(0, 6))

        frm_eb = ttk.LabelFrame(bottom, text="Einband (3D)")
        frm_eb.pack(side="left", fill="both", padx=(10, 0))
        ttk.Radiobutton(frm_eb, text="Hardcover  (mit Falz)",
                        variable=self.einband_var, value="hardcover",
                        command=self._einband_gewechselt).pack(
                            anchor="w", padx=8, pady=(6, 0))
        ttk.Radiobutton(frm_eb, text="Softcover  (ohne Falz)",
                        variable=self.einband_var, value="softcover",
                        command=self._einband_gewechselt).pack(
                            anchor="w", padx=8, pady=(0, 6))
        ttk.Label(frm_eb, text="Falz = Rille am Buchdeckel",
                  foreground="gray").pack(anchor="w", padx=8, pady=(0, 6))

        frm_ps = ttk.LabelFrame(bottom, text="3D-Mockup-Vorlage (nach Buchformat)")
        frm_ps.pack(side="left", fill="both", expand=True, padx=(10, 0))
        row = ttk.Frame(frm_ps)
        row.pack(fill="x", padx=8, pady=(8, 2))
        ttk.Label(row, text="Vorlagen-Ordner:").pack(side="left")
        ttk.Entry(row, textvariable=self.vorlagen_dir_var).pack(
            side="left", fill="x", expand=True, padx=4)
        ttk.Button(row, text="…", width=3, command=self._waehle_vorlagen_dir).pack(
            side="left")
        row2 = ttk.Frame(frm_ps)
        row2.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Label(row2, text="Vorlage:").pack(side="left")
        self.vorlage_box = ttk.Combobox(row2, textvariable=self.vorlage_var,
                                        state="readonly", width=16)
        self.vorlage_box.pack(side="left", padx=6)
        ttk.Label(row2, textvariable=self.vorlage_info,
                  foreground="gray").pack(side="left")
        self._fuelle_vorlagen()

        # Aktionen
        act = ttk.Frame(unten)
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
                self._auto_vorlage()
                self.info_var.set(f"Erkannt. Kurzcode: {self.sc}. "
                                  "Linien bei Bedarf justieren, dann Erstellen.")
            self._ziel_aus_kurzcode()
        except Exception as e:
            messagebox.showerror("Fehler beim Einlesen",
                                 f"{e}\n\n{traceback.format_exc()}")

    # ------------------------------------------------------------------
    # Zielordner
    # ------------------------------------------------------------------
    def _ziel_aus_kurzcode(self):
        """Nach dem Einlesen: vorhandenen Artikelordner suchen bzw. einen neuen
        vorschlagen. Der Titelteil wird aus einem gefundenen Ordner übernommen."""
        vorhanden = core.finde_artikel_ordner(self.sc, self.cfg)
        if vorhanden is not None:
            rest = vorhanden.name[len(self.sc):].lstrip("_ -")
            self.titel_var.set(rest)         # löst _aktualisiere_ziel aus
        self._aktualisiere_ziel()

    def _aktualisiere_ziel(self):
        if not self.sc or self.sc == "unbekannt":
            self.ziel_var.set("—")
            self.ziel_status.set("")
            return
        ordner, existiert = core.ziel_ordner(self.sc, self.titel_var.get(), self.cfg)
        self.ziel_var.set(str(ordner))
        if core.artikeldaten_dir(self.cfg) is None:
            self.ziel_status.set("⚠ Share nicht erreichbar — lokaler Ordner")
        elif existiert:
            self.ziel_status.set("vorhandener Ordner")
        else:
            self.ziel_status.set("wird neu angelegt")

    def _einband_gewechselt(self):
        self.cfg["einband"] = self.einband_var.get()
        core.speichere_config(self.cfg)

    def _zeichne_vorschau(self):
        img = core.rendere_seite(self.doc, PREVIEW_RENDER_DPI)
        # Vorschau in Breite UND Höhe begrenzen (Seitenverhältnis erhalten), damit
        # die Bedienelemente unten immer sichtbar bleiben.
        f = min(PREVIEW_MAX_W / img.width, PREVIEW_MAX_H / img.height, 1.0)
        disp_w = max(1, int(img.width * f))
        disp_h = max(1, int(img.height * f))
        disp = img.resize((disp_w, disp_h), Image.LANCZOS)
        self._pt2px = disp_w / self.reg.seite_pt[0]     # Punkt -> Canvas-Pixel
        self._preview_img = ImageTk.PhotoImage(disp)
        self.canvas.configure(width=disp_w, height=disp_h)
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
    # Vorlagen-Auswahl
    # ------------------------------------------------------------------
    def _fuelle_vorlagen(self):
        """Befüllt die Vorlagen-Combobox aus dem Vorlagen-Ordner."""
        self.cfg["vorlagen_dir"] = self.vorlagen_dir_var.get().strip()
        namen = [e["name"] for e in core.vorlagen_liste(self.cfg)]
        self.vorlage_box.configure(values=namen)

    def _waehle_vorlagen_dir(self):
        pfad = filedialog.askdirectory(title="Vorlagen-Ordner (_NEU_Vorlage) wählen")
        if pfad:
            self.vorlagen_dir_var.set(pfad)
            self.cfg["vorlagen_dir"] = pfad
            core.speichere_config(self.cfg)
            self._fuelle_vorlagen()
            if self.reg is not None:
                self._auto_vorlage()

    def _auto_vorlage(self):
        """Wählt anhand des erkannten Formats automatisch die passende Vorlage."""
        treffer = core.waehle_vorlage(self.reg, self.cfg)
        masse = core.front_masse_cm(self.reg)
        if masse:
            self.vorlage_info.set(f"erkannt: {masse[0]:.1f}×{masse[1]:.1f} cm")
        if treffer:
            self.vorlage_var.set(treffer["name"])
            hinweis = "" if treffer["im_toleranz"] else "  ⚠ prüfen"
            self.vorlage_info.set(
                f"{masse[0]:.1f}×{masse[1]:.1f} cm → {treffer['name']}"
                f" (Δ {treffer['dist_cm']} cm){hinweis}")

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
        self.cfg["vorlagen_dir"] = self.vorlagen_dir_var.get().strip()
        self.cfg["einband"] = self.einband_var.get()
        core.speichere_config(self.cfg)

        out_dir, existiert = core.ziel_ordner(self.sc, self.titel_var.get(), self.cfg)

        if not existiert and not messagebox.askokcancel(
                "Ordner anlegen", f"Der Ordner wird neu angelegt:\n\n{out_dir}"):
            return

        # Vorhandene Dateien nicht stillschweigend überschreiben.
        namen = core.ausgabe_namen(self.sc, self.cfg, self.out_2d.get(),
                                   self.out_3d.get())
        self._alt = core.kollisionen(out_dir, namen)
        if self._alt:
            liste = "\n".join(f"• {p.name}" for p in self._alt)
            if not messagebox.askokcancel(
                    "Dateien vorhanden",
                    f"Im Zielordner liegen schon {len(self._alt)} dieser "
                    f"Datei(en):\n\n{liste}\n\n"
                    "Sie werden nach _alt/<Zeitstempel>/ verschoben, die neuen "
                    "bekommen die regulären Namen.\n\nFortfahren?"):
                return

        self._out_dir = out_dir
        self.btn_run.configure(state="disabled")
        self.info_var.set("Erzeuge …")
        threading.Thread(target=self._run_worker, daemon=True).start()

    def _run_worker(self):
        erzeugt = []
        fehler = []
        try:
            out_dir = self._out_dir
            out_dir.mkdir(parents=True, exist_ok=True)
            if self._alt:
                stempel = datetime.now().strftime("%Y-%m-%d_%H%M%S")
                ziel = core.sichere_weg(self._alt, stempel)
                fehler.append(f"{len(self._alt)} alte Datei(en) nach "
                              f"_alt/{ziel.name}/ verschoben.")
            img_hi = core.rendere_seite(self.doc, int(self.cfg["dpi_print"]))

            if self.out_2d.get():
                front = core.extrahiere(
                    img_hi, self.reg.box_px("front", int(self.cfg["dpi_print"])))
                erzeugt += core.speichere_2d(front, out_dir, self.sc, self.cfg)

            if self.out_3d.get():
                vorlage = self.vorlage_var.get().strip()
                if not vorlage:
                    fehler.append("3D übersprungen: keine Vorlage gewählt.")
                else:
                    dry = sys.platform != "win32"
                    erzeugt += core.erzeuge_3d_photoshop(
                        self.reg, img_hi, self.cfg, out_dir, self.sc, vorlage,
                        dry_run=dry, log=lambda m: None)
                    if dry:
                        fehler.append("3D: kein Windows/Photoshop — Dry-Run "
                                      "(JSX + Slot-PNGs) geschrieben.")
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
        if self.sc and self.sc != "unbekannt":
            out_dir, _ = core.ziel_ordner(self.sc, self.titel_var.get(), self.cfg)
        else:
            out_dir = core.artikeldaten_dir(self.cfg) or \
                core.APP_DIR / self.cfg["output_dir"]
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
