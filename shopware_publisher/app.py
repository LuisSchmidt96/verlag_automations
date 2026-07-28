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
from tkinter import (Tk, filedialog, messagebox, simpledialog, StringVar,
                     BooleanVar, Text, END, DISABLED, NORMAL)
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
        self._secrets = {}          # je Umgebung entsperrtes Secret — nur im RAM!
        self._client = None         # verbundener ShopClient (nach „Verbinden“)
        self._existiert_id = None   # ID, falls das Buch im Shop schon existiert

        umg = core.umgebung(self.cfg)
        self.umg_var = StringVar(value=core.aktive_umgebung(self.cfg))
        self.umg_warnung = StringVar(value="")
        self.shop_url = StringVar(value=umg.get("shop_url", ""))
        self.key_var = StringVar(value=umg.get("access_key_id", ""))
        self.secret_status = StringVar(value="gesperrt")
        self.tax_var = StringVar()
        self.cur_var = StringVar()
        self.cat_var = StringVar()
        self.sc_var = StringVar()
        self.xml_pfad = StringVar()
        self.dry_run = BooleanVar(value=False)
        self.status = StringVar(value="Nicht verbunden.")
        self.info_var = StringVar(value="Bitte eine ONIX-XML wählen.")

        self._build_ui()
        # Beim Start Secret entsperren (bzw. Klartext-Altlast verschlüsseln).
        self.after(100, self._start_entsperren)

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

        # Umgebung (dev / prod) — jede mit eigenem Zugang UND eigenen Shop-IDs
        r0 = ttk.Frame(frm_v); r0.pack(fill="x", padx=8, pady=(8, 2))
        ttk.Label(r0, text="Umgebung:", width=12).pack(side="left")
        self.umg_box = ttk.Combobox(r0, textvariable=self.umg_var,
                                    state="readonly", width=10,
                                    values=core.umgebungs_namen(self.cfg))
        self.umg_box.pack(side="left")
        self.umg_box.bind("<<ComboboxSelected>>", self._wechsle_umgebung)
        self.umg_label = ttk.Label(r0, textvariable=self.umg_warnung,
                                   font=("Segoe UI", 10, "bold"))
        self.umg_label.pack(side="left", padx=10)

        r1 = ttk.Frame(frm_v); r1.pack(fill="x", padx=8, pady=2)
        ttk.Label(r1, text="Shop-URL:", width=12).pack(side="left")
        ttk.Entry(r1, textvariable=self.shop_url).pack(
            side="left", fill="x", expand=True)
        # Shopware nennt die beiden Werte im Admin "Zugriffsschlüssel-ID" und
        # "Geheimer Zugriffsschlüssel" — hier genauso, sonst verwechselt man sie.
        r2 = ttk.Frame(frm_v); r2.pack(fill="x", padx=8, pady=2)
        ttk.Label(r2, text="Zugriffsschlüssel-ID:", width=18).pack(side="left")
        ttk.Entry(r2, textvariable=self.key_var).pack(
            side="left", fill="x", expand=True)
        ttk.Label(r2, text="(von Shopware, nicht geheim)",
                  foreground="gray").pack(side="left", padx=6)

        r2b = ttk.Frame(frm_v); r2b.pack(fill="x", padx=8, pady=2)
        ttk.Label(r2b, text="Geheimer Schlüssel:", width=18).pack(side="left")
        ttk.Label(r2b, textvariable=self.secret_status,
                  foreground="#a60").pack(side="left")
        ttk.Button(r2b, text="Geheimen Schlüssel setzen…",
                   command=self._secret_setzen).pack(side="left", padx=6)
        ttk.Button(r2b, text="Entsperren…", command=self._entsperren).pack(
            side="left", padx=(0, 6))
        ttk.Button(r2b, text="Verbinden", command=self._verbinde).pack(
            side="left", padx=6)

        # --- Zuordnungen ----------------------------------------------
        r3 = ttk.Frame(frm_v); r3.pack(fill="x", padx=8, pady=(2, 2))
        self.tax_box = self._combo(r3, "Steuer:", self.tax_var, 16)
        self.cur_box = self._combo(r3, "Währung:", self.cur_var, 10)
        self.cat_box = self._combo(r3, "Kategorie:", self.cat_var, 26)

        # Verkaufskanal, Seiten-Layout und Hersteller sind im Verlagsshop für
        # jedes Buch gleich -> beim Verbinden automatisch aus dem Bestand.
        r4 = ttk.Frame(frm_v); r4.pack(fill="x", padx=8, pady=(2, 8))
        ttk.Label(r4, text="Aus dem Bestand:", width=18).pack(side="left")
        ttk.Label(r4, textvariable=self.sc_var,
                  foreground="gray").pack(side="left")

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
    # Umgebung (dev / prod)
    # ------------------------------------------------------------------
    @property
    def _secret(self):
        """Entsperrtes Secret der aktiven Umgebung (nur im RAM)."""
        return self._secrets.get(core.aktive_umgebung(self.cfg))

    def _merke_zugang(self):
        """Shop-URL/Key aus den Feldern in die aktive Umgebung schreiben.
        Die URL wird dabei normalisiert (fehlendes https:// ergänzt)."""
        umg = core.umgebung(self.cfg)
        url = core.normalisiere_url(self.shop_url.get())
        self.shop_url.set(url)                  # normalisierte Form auch anzeigen
        umg["shop_url"] = url
        umg["access_key_id"] = self.key_var.get().strip()

    def _wechsle_umgebung(self, _ev=None):
        """Umgebung umschalten: Felder, Zuordnungen und Secret gehören zum Shop."""
        self._merke_zugang()                 # alte Umgebung sichern
        if self._lookups:
            self._merke_zuordnungen()
        self.cfg["aktive_umgebung"] = self.umg_var.get()
        core.speichere_config(self.cfg)

        umg = core.umgebung(self.cfg)
        self.shop_url.set(umg.get("shop_url", ""))
        self.key_var.set(umg.get("access_key_id", ""))
        # Zuordnungen sind shopspezifisch (UUIDs!) -> beim Wechsel verwerfen,
        # bis erneut verbunden wurde.
        self._lookups = {}
        self._vorlage = {}
        self._client = None                  # Verbindung gehört zum alten Shop
        self._existiert_id = None
        for box, var in ((self.tax_box, self.tax_var), (self.cur_box, self.cur_var),
                         (self.cat_box, self.cat_var)):
            box.configure(values=[])
            var.set("")
        self.sc_var.set("")
        self.status.set("Nicht verbunden.")
        self._aktualisiere_umgebung()
        if core.hat_secret(umg) and not self._secret:
            self._entsperren()

    def _aktualisiere_umgebung(self):
        """Statusfarbe/Warnung — Produktivshop deutlich sichtbar machen."""
        name = core.aktive_umgebung(self.cfg)
        if core.ist_produktiv(name):
            self.umg_warnung.set("⚠  PRODUKTIVSHOP — Änderungen sind echt")
            self.umg_label.configure(foreground="#c00")
        else:
            self.umg_warnung.set("Testumgebung")
            self.umg_label.configure(foreground="#0a7")
        self._setze_secret_status()

    # ------------------------------------------------------------------
    # Master-Passwort / Secret (je Umgebung)
    # ------------------------------------------------------------------
    def _setze_secret_status(self):
        self.secret_status.set("entsperrt ✓" if self._secret else "gesperrt")

    def _start_entsperren(self):
        """Beim Start: Secret entsperren bzw. Klartext-Altlast verschlüsseln."""
        self._aktualisiere_umgebung()
        umg = core.umgebung(self.cfg)
        klartext = core.klartext_secret_vorhanden(umg)
        if klartext and not core.hat_secret(umg):
            messagebox.showwarning(
                "Secret im Klartext",
                "In der config.json liegt noch ein Secret im Klartext.\n\n"
                "Es wird jetzt mit einem Master-Passwort verschlüsselt.")
            if self._frage_neues_passwort(klartext):
                messagebox.showinfo("Fertig", "Secret ist jetzt verschlüsselt.")
            return
        if core.hat_secret(umg):
            self._entsperren()

    def _entsperren(self):
        """Master-Passwort abfragen und das Secret der Umgebung entschlüsseln."""
        umg = core.umgebung(self.cfg)
        name = core.aktive_umgebung(self.cfg)
        if not core.hat_secret(umg):
            messagebox.showinfo(
                "Kein Secret",
                f"Für „{name}“ ist noch kein Secret hinterlegt — bitte zuerst "
                f"„Secret setzen…“.")
            return
        for _ in range(3):
            pw = simpledialog.askstring(
                "Master-Passwort", f"Master-Passwort für „{name}“:",
                show="•", parent=self)
            if not pw:
                break                       # Abbrechen -> bleibt gesperrt
            try:
                self._secrets[name] = core.hole_secret(umg, pw)
                self._setze_secret_status()
                self.status.set(f"„{name}“ entsperrt — bereit zum Verbinden.")
                return
            except core.PasswortFehler as e:
                messagebox.showerror("Master-Passwort", str(e))
        self._secrets.pop(name, None)
        self._setze_secret_status()

    def _frage_neues_passwort(self, secret: str) -> bool:
        """Neues Master-Passwort setzen und das Secret damit verschlüsseln."""
        name = core.aktive_umgebung(self.cfg)
        pw = simpledialog.askstring(
            "Master-Passwort",
            f"Master-Passwort für „{name}“ — frei wählbar, NICHT von Shopware.\n"
            f"Es verschlüsselt den geheimen Schlüssel auf der Platte:",
            show="•", parent=self)
        if not pw:
            return False
        pw2 = simpledialog.askstring("Master-Passwort",
                                     "Passwort wiederholen:", show="•", parent=self)
        if pw != pw2:
            messagebox.showerror("Master-Passwort",
                                 "Die Passwörter stimmen nicht überein.")
            return False
        try:
            core.setze_secret(core.umgebung(self.cfg), secret, pw)
        except core.PasswortFehler as e:
            messagebox.showerror("Master-Passwort", str(e))
            return False
        # Shop-URL/Access-Key mitspeichern — sonst gingen sie beim Setzen des
        # Secrets verloren (sie stehen nur in den Eingabefeldern).
        self._merke_zugang()
        core.speichere_config(self.cfg)
        self._secrets[name] = secret
        self._setze_secret_status()
        return True

    def _secret_setzen(self):
        """Geheimen Zugriffsschlüssel der Umgebung eintragen (verschlüsselt)."""
        name = core.aktive_umgebung(self.cfg)
        secret = simpledialog.askstring(
            "Geheimer Zugriffsschlüssel",
            f"Geheimer Zugriffsschlüssel für „{name}“\n"
            f"(der lange Wert von Shopware, der nur einmal angezeigt wurde):",
            show="•", parent=self)
        if not secret:
            return
        if self._frage_neues_passwort(secret.strip()):
            messagebox.showinfo(
                "Gespeichert",
                f"Der geheime Schlüssel für „{name}“ ist verschlüsselt in der "
                f"config.json abgelegt.\n"
                f"Ohne das Master-Passwort ist er nicht lesbar.")

    # ------------------------------------------------------------------
    # Verbindung + Zuordnungen
    # ------------------------------------------------------------------
    def _verbinde(self):
        if not self._secret:
            messagebox.showwarning(
                "Gesperrt", "Bitte zuerst das Master-Passwort eingeben "
                            "(„Entsperren…“) bzw. ein Secret setzen.")
            return
        self._merke_zugang()
        core.speichere_config(self.cfg)
        self.status.set("Verbinde …")
        threading.Thread(target=self._verbinde_worker, daemon=True).start()

    def _verbinde_worker(self):
        try:
            umg = core.umgebung(self.cfg)
            c = core.ShopClient(umg.get("shop_url", ""),
                                umg.get("access_key_id", ""),
                                self._secret or "",
                                tls_pruefen=bool(umg.get("tls_pruefen", True)))
            version = c.verbinde().get("version", "?")
            self._client = c            # verbunden halten (für Existenz-Prüfung)
            self._lookups = {
                "tax": c.steuersaetze(), "cur": c.waehrungen(),
                "cat": c.kategorien(),
            }
            # Verkaufskanal, Seiten-Layout und Hersteller aus dem Bestand
            # übernehmen (für jedes Buch gleich). Ohne Verkaufskanal wäre das
            # neue Produkt im Shop unsichtbar.
            vorlage = c.vorlage_vom_bestand()
            for k in ("cms_page_id", "sales_channel_id", "visibility",
                      "manufacturer_id"):
                if vorlage.get(k):
                    umg[k] = vorlage[k]
            self._vorlage = vorlage
            self.after(0, self._fuelle_zuordnungen, version)
        except Exception as e:
            self.after(0, lambda: (self.status.set("Nicht verbunden."),
                                   messagebox.showerror("Verbindung", str(e))))

    def _fuelle_zuordnungen(self, version):
        name = core.aktive_umgebung(self.cfg)
        self.status.set(f"Verbunden mit „{name}“ (Shopware {version})")
        umg = core.umgebung(self.cfg)

        def fill(box, var, eintraege, beschriftung, cfg_key, vorauswahl=None):
            self._lookups[cfg_key] = eintraege
            box.configure(values=[beschriftung(e) for e in eintraege])
            gesetzt = umg.get(cfg_key)
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
        v = getattr(self, "_vorlage", {}) or {}
        kanal = v.get("sales_channel_name") or umg.get("sales_channel_id", "")
        hersteller = v.get("manufacturer_name") or "—"
        if kanal:
            self.sc_var.set(f"Verkaufskanal: {kanal}   ·   Hersteller: "
                            f"{hersteller}   ·   Seiten-Layout übernommen")
        else:
            self.sc_var.set("⚠ kein Verkaufskanal gefunden — Produkt bliebe "
                            "im Shop unsichtbar!")
        self._merke_zuordnungen()

    def _merke_zuordnungen(self):
        """Auswahl der Comboboxen in die aktive Umgebung schreiben (die IDs
        sind shopspezifisch — dev und prod haben verschiedene UUIDs)."""
        umg = core.umgebung(self.cfg)

        def pick(var, key, beschriftung):
            eintraege = self._lookups.get(key) or []
            for e in eintraege:
                if beschriftung(e) == var.get():
                    umg[key] = e["id"]
                    if key == "tax_id":
                        umg["tax_rate"] = float(e["taxRate"])
                    return
        pick(self.tax_var, "tax_id", lambda e: f"{e['name']} ({e['taxRate']} %)")
        pick(self.cur_var, "currency_id", lambda e: e["isoCode"])
        pick(self.cat_var, "category_id", lambda e: e.get("name") or e["id"])
        # Hersteller/Verkaufskanal/CMS werden aus dem Bestand übernommen
        # (kein Dropdown) — hier nichts zu tun.
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
            eff = core.effektiv(self.cfg)
            self.felder = core.lade_buchfelder(self.xml_pfad.get(), eff)
            self.bilder = core.finde_bilder(self.felder["shortcode"], eff)
            self._pruefe_existenz()
            self._zeige_vorschau()
            self.info_var.set(f"Kurzcode {self.felder['shortcode']}")
        except Exception as e:
            messagebox.showerror("Fehler beim Einlesen",
                                 f"{e}\n\n{traceback.format_exc()}")

    def _pruefe_existenz(self):
        """Falls verbunden: schauen, ob das Buch im Shop schon existiert."""
        self._existiert_id = None
        if not (self._client and self.felder):
            return
        nummer = self.felder.get("isbn13_formatiert") or self.felder["isbn13"]
        try:
            self._existiert_id = self._client.produkt_id_zu_nummer(nummer)
        except core.ShopFehler:
            self._existiert_id = None        # nicht blockieren, nur Hinweis

    def _zeige_vorschau(self):
        f, b = self.felder, self.bilder
        eff = core.effektiv(self.cfg)
        satz = float(eff.get("tax_rate", 7.0))
        brutto = float(f.get("preis_brutto") or 0.0)
        zeilen = [
            f"Umgebung       {core.aktive_umgebung(self.cfg)}"
            f"   ({eff.get('shop_url') or '— keine Shop-URL —'})",
            f"Name           {core.produkt_name(f)}",
            f"Artikelnummer  {f['isbn13']}   (= EAN)",
            f"Preis          {brutto:.2f} {f['waehrung']} brutto"
            f"   /   {core.netto(brutto, satz):.4f} netto ({satz} %)",
            f"Bestand        {eff.get('default_stock', 0)}",
            f"Status         {'AKTIV' if eff.get('aktiv') else 'ENTWURF (inaktiv)'}",
        ]
        if not self._client:
            zeilen.append("Im Shop        ? (noch nicht verbunden)")
        elif self._existiert_id:
            zeilen.append("Im Shop        ⚠ EXISTIERT BEREITS — würde "
                          "überschrieben (nur nach Bestätigung)")
        else:
            zeilen.append("Im Shop        neu — wird angelegt")
        zeilen.append("")
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

        # Dev- und Produktivshop liegen auf demselben Server — vor dem Schreiben
        # zeigen, in WELCHEN Shop es geht. Ein Vertipper oder eine vergessene
        # Umgebung soll nicht still im Livesystem landen.
        self._ueberschreiben = False
        if not self.dry_run.get():
            name = core.aktive_umgebung(self.cfg)
            eff = core.effektiv(self.cfg)
            ziel = eff.get("shop_url") or "— keine Shop-URL —"
            warnung = ("\n\n⚠  ACHTUNG: PRODUKTIVSHOP — die Änderung ist echt!"
                       if core.ist_produktiv(name) else "")
            if not messagebox.askokcancel(
                    "Wirklich anlegen?",
                    f"Umgebung: {name}\n{ziel}\n\n"
                    f"Artikelnummer: {self.felder['isbn13']}\n"
                    f"Status: {'AKTIV' if eff.get('aktiv') else 'Entwurf (inaktiv)'}"
                    f"{warnung}"):
                return

            # Bestehendes Buch? Ausdrücklich warnen — die gepflegten Shop-Daten
            # (Name, Beschreibung, Preis, Cover …) würden mit den ONIX-Werten
            # überschrieben. Nur bei klarem „Ja“ freigeben.
            if self._existiert_id:
                if not messagebox.askyesno(
                        "Bestehendes Buch überschreiben?",
                        "Dieses Buch existiert im Shop BEREITS.\n\n"
                        "Beim Fortfahren werden Name, Beschreibung, Preis, "
                        "Bilder und weitere Felder mit den Werten aus der "
                        "ONIX-Datei ÜBERSCHRIEBEN. Bereits im Shop gepflegte "
                        "Angaben gehen dabei verloren.\n\n"
                        "Wirklich überschreiben?", icon="warning", default="no"):
                    self.info_var.set("Abgebrochen — nichts geändert.")
                    return
                self._ueberschreiben = True

        self.btn_run.configure(state="disabled")
        self.info_var.set("Sende …")
        threading.Thread(target=self._run_worker, daemon=True).start()

    def _run_worker(self):
        try:
            ergebnis = core.veroeffentliche(
                self.felder, self.cfg, self.bilder, secret=self._secret or "",
                dry_run=self.dry_run.get(),
                ueberschreiben=getattr(self, "_ueberschreiben", False),
                log=lambda m: None)
            self.after(0, self._fertig, ergebnis, None)
        except Exception as e:
            traceback.print_exc()
            self.after(0, self._fertig, None, e)

    def _fertig(self, ergebnis, fehler):
        self.btn_run.configure(state="normal")
        if fehler:
            # Sicherheitsnetz: Der Kern hat das Überschreiben verweigert (die
            # Vorab-Prüfung lief z. B. nicht, weil beim Laden noch nicht
            # verbunden war). Nichts wurde gesendet — hier ausdrücklich fragen.
            if isinstance(fehler, core.ProduktExistiert):
                self._existiert_id = fehler.produkt_id
                if messagebox.askyesno(
                        "Bestehendes Buch überschreiben?",
                        f"Das Buch {fehler.nummer} existiert im Shop BEREITS "
                        f"und wurde NICHT verändert.\n\n"
                        "Beim Überschreiben gehen im Shop gepflegte Angaben "
                        "(Name, Beschreibung, Preis, Bilder …) verloren.\n\n"
                        "Wirklich überschreiben?", icon="warning", default="no"):
                    self._ueberschreiben = True
                    self.btn_run.configure(state="disabled")
                    self.info_var.set("Überschreibe …")
                    threading.Thread(target=self._run_worker, daemon=True).start()
                else:
                    self.info_var.set("Übersprungen — bestehendes Buch "
                                      "unverändert.")
                return
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
        self._existiert_id = ergebnis["payload"]["id"]   # existiert jetzt
        was = "angelegt" if ergebnis.get("neu") else "aktualisiert"
        self.info_var.set(f"Produkt {was}.")
        url = ergebnis.get("admin_url") or ""
        if url and messagebox.askyesno(
                "Fertig", f"Produkt wurde als Entwurf {was}.\n\n"
                          "Im Shop-Admin öffnen?"):
            webbrowser.open(url)


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
