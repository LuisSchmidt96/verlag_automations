#!/usr/bin/env python3
"""
Shopware-Publisher (Tkinter-GUI)
================================

Legt aus einem VLB-ONIX-XML ein Shopware-6-Produkt als **Entwurf** an
(inaktiv, also im Shop nicht sichtbar, bis es freigegeben wird).

Ablauf in der GUI:
1. Einmalig: Shop-URL + Zugangsdaten eintragen -> "Verbinden".
   Danach Steuersatz / Währung / Kategorie / Hersteller auswählen.
2. ONIX-XML wählen -> die erkannten Felder + gefundenen Bilder werden gezeigt.
3. "Als Entwurf anlegen" -> Bilder hochladen + Produkt anlegen/aktualisieren.

Die Bilder kommen vom cover_previews-Tool: sie werden anhand des Kurzcodes im
Artikeldaten-Ordner des Buchs gesucht (2D_72_<sc>.jpg als Cover, 3D als Galerie).

Ein zweiter Lauf zum selben Buch **aktualisiert** das Produkt (gleiche ISBN =
gleiche Produkt-ID), es entsteht kein Duplikat.

config.json (inkl. Zugangsdaten) liegt neben der .exe.
"""

from __future__ import annotations

import os
import sys
import threading
import traceback
import webbrowser
from pathlib import Path
from tkinter import (Tk, filedialog, messagebox, StringVar, BooleanVar, Text,
                     END, DISABLED, NORMAL)
from tkinter import ttk

from shopware_publisher import core


class App(Tk):
    def __init__(self):
        super().__init__()
        self.title("Shopware-Publisher")
        self.geometry("980x780")
        self.minsize(860, 640)

        try:
            self.cfg = core.lade_config()
        except Exception as e:
            messagebox.showerror("Config-Fehler",
                                 f"Konnte config.json nicht laden:\n{e}")
            self.cfg = dict(core.DEFAULT_CONFIG)

        self.felder = None          # geladene Buchdaten
        self.bilder = None          # gefundene Bilddateien
        self._lookups = {}          # Shop-Listen für die Auswahlfelder

        self.shop_url = StringVar(value=self.cfg.get("shop_url", ""))
        self.key_var = StringVar(value=self.cfg.get("access_key_id", ""))
        self.secret_var = StringVar(value=self.cfg.get("secret_access_key", ""))
        self.tax_var = StringVar()
        self.cur_var = StringVar()
        self.cat_var = StringVar()
        self.man_var = StringVar()
        self.xml_pfad = StringVar()
        self.dry_run = BooleanVar(value=False)
        self.status = StringVar(value="Nicht verbunden.")
        self.info_var = StringVar(value="Bitte eine ONIX-XML wählen.")

        self._build_ui()

    # ------------------------------------------------------------------
    def _build_ui(self):
        pad = {"padx": 10, "pady": 5}

        head = ttk.Frame(self)
        head.pack(fill="x", **pad)
        ttk.Label(head, text="Shopware-Publisher",
                  font=("Segoe UI", 14, "bold")).pack(side="left")
        ttk.Label(head, textvariable=self.status,
                  foreground="gray").pack(side="right")

        # Unten verankern, damit die Buttons immer sichtbar bleiben
        unten = ttk.Frame(self)
        unten.pack(side="bottom", fill="x")

        # --- Verbindung ------------------------------------------------
        frm_v = ttk.LabelFrame(self, text="Shop-Verbindung (Admin → Einstellungen "
                                          "→ System → Integrationen)")
        frm_v.pack(fill="x", **pad)
        r1 = ttk.Frame(frm_v); r1.pack(fill="x", padx=8, pady=(8, 2))
        ttk.Label(r1, text="Shop-URL:", width=12).pack(side="left")
        ttk.Entry(r1, textvariable=self.shop_url).pack(
            side="left", fill="x", expand=True)
        r2 = ttk.Frame(frm_v); r2.pack(fill="x", padx=8, pady=2)
        ttk.Label(r2, text="Access-Key:", width=12).pack(side="left")
        ttk.Entry(r2, textvariable=self.key_var).pack(
            side="left", fill="x", expand=True)
        ttk.Label(r2, text="Secret:").pack(side="left", padx=(8, 0))
        ttk.Entry(r2, textvariable=self.secret_var, show="•", width=26).pack(
            side="left", padx=4)
        ttk.Button(r2, text="Verbinden", command=self._verbinde).pack(
            side="left", padx=6)

        # --- Zuordnungen ----------------------------------------------
        r3 = ttk.Frame(frm_v); r3.pack(fill="x", padx=8, pady=(2, 8))
        self.tax_box = self._combo(r3, "Steuer:", self.tax_var, 16)
        self.cur_box = self._combo(r3, "Währung:", self.cur_var, 10)
        self.cat_box = self._combo(r3, "Kategorie:", self.cat_var, 22)
        self.man_box = self._combo(r3, "Hersteller:", self.man_var, 18)

        # --- Buch (ONIX) ----------------------------------------------
        frm_b = ttk.LabelFrame(self, text="Buch (VLB-ONIX-XML)")
        frm_b.pack(fill="x", **pad)
        b1 = ttk.Frame(frm_b); b1.pack(fill="x", padx=8, pady=8)
        ttk.Entry(b1, textvariable=self.xml_pfad).pack(
            side="left", fill="x", expand=True)
        ttk.Button(b1, text="Auswählen…", command=self._waehle_xml).pack(
            side="left", padx=(6, 0))

        # --- Vorschau --------------------------------------------------
        frm_p = ttk.LabelFrame(self, text="Vorschau (was im Shop landet)")
        frm_p.pack(fill="both", expand=True, **pad)
        self.txt = Text(frm_p, height=14, wrap="word", state=DISABLED,
                        background="#fbfbfb")
        self.txt.pack(fill="both", expand=True, padx=6, pady=6)

        # --- Aktionen --------------------------------------------------
        act = ttk.Frame(unten)
        act.pack(fill="x", **pad)
        ttk.Checkbutton(act, text="Dry-Run (nur anzeigen, nichts senden)",
                        variable=self.dry_run).pack(side="left")
        ttk.Label(act, textvariable=self.info_var,
                  foreground="gray").pack(side="left", padx=12)
        self.btn_run = ttk.Button(act, text="Als Entwurf anlegen",
                                  command=self._run)
        self.btn_run.pack(side="right")

    def _combo(self, parent, label, var, width):
        ttk.Label(parent, text=label).pack(side="left")
        box = ttk.Combobox(parent, textvariable=var, state="readonly",
                           width=width)
        box.pack(side="left", padx=(2, 10))
        return box

    # ------------------------------------------------------------------
    # Verbindung + Zuordnungen
    # ------------------------------------------------------------------
    def _verbinde(self):
        self.cfg["shop_url"] = self.shop_url.get().strip()
        self.cfg["access_key_id"] = self.key_var.get().strip()
        self.cfg["secret_access_key"] = self.secret_var.get().strip()
        core.speichere_config(self.cfg)
        self.status.set("Verbinde …")
        threading.Thread(target=self._verbinde_worker, daemon=True).start()

    def _verbinde_worker(self):
        try:
            c = core.ShopClient(self.cfg["shop_url"], self.cfg["access_key_id"],
                                self.cfg["secret_access_key"])
            version = c.verbinde().get("version", "?")
            self._lookups = {
                "tax": c.steuersaetze(), "cur": c.waehrungen(),
                "cat": c.kategorien(), "man": c.hersteller(),
            }
            self.after(0, self._fuelle_zuordnungen, version)
        except Exception as e:
            self.after(0, lambda: (self.status.set("Nicht verbunden."),
                                   messagebox.showerror("Verbindung", str(e))))

    def _fuelle_zuordnungen(self, version):
        self.status.set(f"Verbunden (Shopware {version})")

        def fill(box, var, eintraege, beschriftung, cfg_key, vorauswahl=None):
            self._lookups[cfg_key] = eintraege
            box.configure(values=[beschriftung(e) for e in eintraege])
            gesetzt = self.cfg.get(cfg_key)
            for e in eintraege:
                if e["id"] == gesetzt:
                    var.set(beschriftung(e))
                    return
            if vorauswahl:
                for e in eintraege:
                    if vorauswahl(e):
                        var.set(beschriftung(e))
                        return

        fill(self.tax_box, self.tax_var, self._lookups["tax"],
             lambda e: f"{e['name']} ({e['taxRate']} %)", "tax_id",
             vorauswahl=lambda e: abs(float(e["taxRate"]) - 7.0) < 0.01)
        fill(self.cur_box, self.cur_var, self._lookups["cur"],
             lambda e: e["isoCode"], "currency_id",
             vorauswahl=lambda e: e["isoCode"] == "EUR")
        fill(self.cat_box, self.cat_var, self._lookups["cat"],
             lambda e: e.get("name") or e["id"], "category_id")
        fill(self.man_box, self.man_var, self._lookups["man"],
             lambda e: e.get("name") or e["id"], "manufacturer_id")
        self._merke_zuordnungen()

    def _merke_zuordnungen(self):
        """Auswahl der Comboboxen zurück in die Config schreiben."""
        def pick(var, key, beschriftung):
            eintraege = self._lookups.get(key) or []
            for e in eintraege:
                if beschriftung(e) == var.get():
                    self.cfg[key] = e["id"]
                    if key == "tax_id":
                        self.cfg["tax_rate"] = float(e["taxRate"])
                    return
        pick(self.tax_var, "tax_id", lambda e: f"{e['name']} ({e['taxRate']} %)")
        pick(self.cur_var, "currency_id", lambda e: e["isoCode"])
        pick(self.cat_var, "category_id", lambda e: e.get("name") or e["id"])
        pick(self.man_var, "manufacturer_id", lambda e: e.get("name") or e["id"])
        core.speichere_config(self.cfg)

    # ------------------------------------------------------------------
    # ONIX laden + Vorschau
    # ------------------------------------------------------------------
    def _waehle_xml(self):
        start = self.cfg.get("last_input_dir") or None
        pfad = filedialog.askopenfilename(
            title="ONIX-XML wählen", initialdir=start,
            filetypes=[("ONIX-XML", "*.xml"), ("Alle Dateien", "*.*")])
        if not pfad:
            return
        self.xml_pfad.set(pfad)
        self.cfg["last_input_dir"] = str(Path(pfad).parent)
        core.speichere_config(self.cfg)
        self._lade_buch()

    def _lade_buch(self):
        try:
            self.felder = core.lade_buchfelder(self.xml_pfad.get(), self.cfg)
            self.bilder = core.finde_bilder(self.felder["shortcode"], self.cfg)
            self._zeige_vorschau()
            self.info_var.set(f"Kurzcode {self.felder['shortcode']}")
        except Exception as e:
            messagebox.showerror("Fehler beim Einlesen",
                                 f"{e}\n\n{traceback.format_exc()}")

    def _zeige_vorschau(self):
        f, b = self.felder, self.bilder
        satz = float(self.cfg.get("tax_rate", 7.0))
        brutto = float(f.get("preis_brutto") or 0.0)
        zeilen = [
            f"Name           {core.produkt_name(f)}",
            f"Artikelnummer  {f['isbn13']}   (= EAN)",
            f"Preis          {brutto:.2f} {f['waehrung']} brutto"
            f"   /   {core.netto(brutto, satz):.4f} netto ({satz} %)",
            f"Bestand        {self.cfg.get('default_stock', 0)}",
            f"Status         {'AKTIV' if self.cfg.get('aktiv') else 'ENTWURF (inaktiv)'}",
            "",
        ]
        ordner = b.get("ordner")
        zeilen.append(f"Bilder-Ordner  {ordner if ordner else '— nicht gefunden —'}")
        zeilen.append(f"  Cover        {Path(b['cover']).name if b.get('cover') else '— fehlt —'}")
        for g in b.get("galerie", []):
            zeilen.append(f"  Galerie      {Path(g).name}")
        if not b.get("cover"):
            zeilen.append("  (ohne Bild wird das Produkt trotzdem angelegt)")
        zeilen += ["", "Beschreibung (HTML):", core.baue_beschreibung(f)]

        self.txt.configure(state=NORMAL)
        self.txt.delete("1.0", END)
        self.txt.insert(END, "\n".join(zeilen))
        self.txt.configure(state=DISABLED)

    # ------------------------------------------------------------------
    # Anlegen
    # ------------------------------------------------------------------
    def _run(self):
        if not self.felder:
            messagebox.showwarning("Kein Buch", "Bitte zuerst eine ONIX-XML wählen.")
            return
        if self._lookups:
            self._merke_zuordnungen()
        self.btn_run.configure(state="disabled")
        self.info_var.set("Sende …")
        threading.Thread(target=self._run_worker, daemon=True).start()

    def _run_worker(self):
        try:
            ergebnis = core.veroeffentliche(
                self.felder, self.cfg, self.bilder,
                dry_run=self.dry_run.get(), log=lambda m: None)
            self.after(0, self._fertig, ergebnis, None)
        except Exception as e:
            traceback.print_exc()
            self.after(0, self._fertig, None, e)

    def _fertig(self, ergebnis, fehler):
        self.btn_run.configure(state="normal")
        if fehler:
            self.info_var.set("Fehler.")
            messagebox.showerror("Fehler", str(fehler))
            return
        if self.dry_run.get():
            import json
            self.txt.configure(state=NORMAL)
            self.txt.delete("1.0", END)
            self.txt.insert(END, "DRY-RUN — dieser Payload würde gesendet:\n\n" +
                            json.dumps(ergebnis["payload"], indent=2,
                                       ensure_ascii=False))
            self.txt.configure(state=DISABLED)
            self.info_var.set("Dry-Run fertig — nichts gesendet.")
            return
        self.info_var.set("Als Entwurf angelegt.")
        url = ergebnis.get("admin_url") or ""
        if url and messagebox.askyesno(
                "Fertig",
                "Produkt wurde als Entwurf angelegt/aktualisiert.\n\n"
                "Im Shop-Admin öffnen?"):
            webbrowser.open(url)


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
