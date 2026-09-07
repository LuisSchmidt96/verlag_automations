"""Buchdurchgang — Oberfläche.

Ein Fenster, drei Schritte, dazwischen Tore:

    0. Buch wählen   Umschlag-PDF + ONIX-XML, beide Proben
    1. Cover         cover_previews
    2. pi & bi       pi_bi_generator
    3. Shop          shopware_publisher
    4. Presse        Dateien per SFTP auf den Webserver
    5. Ablegen       Buchordner auf den Netzordner

**Weiter** wird erst frei, wenn alle Häkchen des Schritts gesetzt sind. Die
Häkchen liegen in ``durchgang.json`` im Buchordner — nicht beim Werkzeug —,
damit auch Wochen später und an einem anderen Rechner erkennbar bleibt, was
geprüft wurde.
"""

from __future__ import annotations

import os
import queue
import webbrowser
import sys
import threading
import traceback
from pathlib import Path
from tkinter import (Tk, filedialog, messagebox, simpledialog, StringVar,
                     BooleanVar, Text, END, DISABLED, NORMAL)
from tkinter import ttk

from buchdurchgang import core
from shopware_publisher import core as sw

PAD = {"padx": 8, "pady": 6}


class App(Tk):
    def __init__(self):
        super().__init__()
        self.title("Buchdurchgang")
        self.geometry("1040x820")
        self.minsize(920, 700)

        try:
            self.cfg = core.lade_config()
        except Exception as e:
            messagebox.showerror("Config-Fehler", f"config.json:\n{e}")
            self.cfg = dict(core.DEFAULT_CONFIG)

        self.paar = None            # Ergebnis der Eingangsprobe
        self.ordner = None          # Buchordner (Path)
        self.stand = {"schritte": {}}
        self.aktiv = "buch"         # "buch" | "cover" | "pibi" | "shop"
        self._secret = None
        self._client = None
        self._kat_gewaehlt: dict[str, str] = {}   # nur aus dem Vorschlag
        self._kat_fehlend: list[str] = []         # Beteiligte ohne Kategorie
        self._haken: dict[str, list[BooleanVar]] = {}
        # Arbeitsthreads fassen Tk NICHT an — sie legen Nachrichten hier ab,
        # und nur der Hauptthread nimmt sie heraus (_pumpe). Ein `after()` aus
        # einem fremden Thread verklemmt den Tcl-Interpreter, sobald beide
        # gleichzeitig hineingreifen; genau daran hing der Cover-Schritt.
        self._nachrichten: queue.Queue = queue.Queue()
        self._logs: dict[str, Text] = {}

        self.arbeit_var = StringVar(value=str(core.arbeitsordner(self.cfg)))
        self.ablage_var = StringVar(value=self.cfg.get("ablageort", ""))
        self.pdf_var = StringVar()
        self.xml_var = StringVar()
        self.titel_var = StringVar()
        self.ordner_var = StringVar(value="—")
        self.bib_var = StringVar()          # Blick-ins-Buch-PDF (von Hand)
        self._sftp_pw = None                # entsperrt, nur im RAM
        self.status = StringVar(value="Umschlag-PDF und ONIX-XML wählen.")

        self._build()
        self._zeige_schritt("buch")
        self.after(150, self._pumpe)

    # ------------------------------------------------------------------
    def _build(self):
        # --- Ordner ------------------------------------------------------
        # Gearbeitet wird örtlich (schnell), abgelegt am Ende auf dem Netz.
        oben = ttk.LabelFrame(self, text="Ordner")
        oben.pack(fill="x", **PAD)
        for text, var, cmd, hinweis in (
                ("Arbeitsordner:", self.arbeit_var, self._waehle_arbeit,
                 "hier wird gearbeitet (örtlich, schnell)"),
                ("Ablageort:", self.ablage_var, self._waehle_ablage,
                 "dorthin wandert der fertige Ordner in Schritt 5")):
            r = ttk.Frame(oben); r.pack(fill="x", padx=8, pady=3)
            ttk.Label(r, text=text, width=14).pack(side="left")
            ttk.Entry(r, textvariable=var).pack(side="left", fill="x", expand=True)
            ttk.Button(r, text="…", width=3, command=cmd).pack(side="left", padx=(6, 0))
            ttk.Label(r, text=hinweis, foreground="gray").pack(side="left", padx=(8, 0))

        # --- Schrittliste links, Inhalt rechts --------------------------
        mitte = ttk.Frame(self); mitte.pack(fill="both", expand=True, **PAD)

        links = ttk.LabelFrame(mitte, text="Schritte")
        links.pack(side="left", fill="y", padx=(0, 8))
        self.schritt_labels = {}
        for sid, titel in [("buch", "0 — Buch")] + \
                [(s["id"], s["titel"]) for s in core.SCHRITTE]:
            lbl = ttk.Label(links, text=f"  {titel}", width=24, anchor="w")
            lbl.pack(fill="x", padx=6, pady=3)
            lbl.bind("<Button-1>", lambda _e, s=sid: self._springe(s))
            self.schritt_labels[sid] = lbl

        self.rechts = ttk.LabelFrame(mitte, text="Schritt")
        self.rechts.pack(side="left", fill="both", expand=True)

        # --- Fuß ---------------------------------------------------------
        fuss = ttk.Frame(self); fuss.pack(fill="x", **PAD)
        # Die Knöpfe ZUERST packen. Tk verteilt in Reihenfolge; stand die
        # Pfadzeile vorn, beanspruchte sie ihre volle Breite und schob die
        # Knöpfe aus dem Fenster — bei langen Pfaden waren sie schlicht weg.
        self.btn_weiter = ttk.Button(fuss, text="Weiter ▸",
                                     command=self._weiter, state="disabled")
        self.btn_weiter.pack(side="right")
        # Zurück ist immer erlaubt: nachsehen, was ein früherer Schritt
        # gemeldet hat, darf nichts kosten. Gearbeitet wird dort nur, wenn man
        # den jeweiligen Knopf noch einmal drückt.
        self.btn_zurueck = ttk.Button(fuss, text="◂ Zurück",
                                      command=self._zurueck, state="disabled")
        self.btn_zurueck.pack(side="right", padx=(0, 6))
        ttk.Button(fuss, text="Ordner öffnen",
                   command=self._ordner_oeffnen).pack(side="right", padx=6)
        ttk.Label(fuss, textvariable=self.ordner_var, foreground="gray",
                  anchor="w").pack(side="left", fill="x", expand=True)
        ttk.Label(self, textvariable=self.status,
                  foreground="#036").pack(fill="x", padx=16, pady=(0, 8))

    # ------------------------------------------------------------------
    # Gerüst je Schritt
    # ------------------------------------------------------------------
    def _leere_rechts(self):
        for kind in self.rechts.winfo_children():
            kind.destroy()

    def _zeige_schritt(self, sid: str):
        self.aktiv = sid
        self._leere_rechts()
        self.rechts.configure(text={"buch": "0 — Buch wählen"}.get(
            sid, core.schritt(sid)["titel"] if sid != "buch" else ""))
        {"buch": self._baue_buch, "cover": self._baue_cover,
         "pibi": self._baue_pibi, "shop": self._baue_shop,
         "ablegen": self._baue_ablegen, "presse": self._baue_presse}[sid]()
        self._male_schrittliste()
        self._pruefe_tor()

    def _male_schrittliste(self):
        for sid, lbl in self.schritt_labels.items():
            if sid == "buch":
                zeichen = "✓" if self.paar else "·"
            elif core.ist_fertig(self.stand, sid):
                zeichen = "✓"
            elif core.ist_gelaufen(self.stand, sid):
                zeichen = "◐"          # gelaufen, aber nicht abgehakt
            else:
                zeichen = "·"
            titel = lbl.cget("text").strip()[2:] if lbl.cget("text")[:1] in "✓◐·" \
                else lbl.cget("text").strip()
            lbl.configure(text=f"{zeichen} {titel}",
                          foreground="#060" if zeichen == "✓" else
                          ("#960" if zeichen == "◐" else "black"))

    def _log_feld(self, eltern, sid: str) -> Text:
        t = Text(eltern, height=9, wrap="word", state=DISABLED,
                 background="#fbfbfb")
        t.pack(fill="both", expand=True, padx=8, pady=(0, 6))
        self._logs[sid] = t
        return t

    def _knopf(self, name: str):
        """Knopf, falls es ihn noch gibt.

        Läuft ein Schritt und man blättert weiter, wird sein Panel zerstört;
        der Abschluss-Handler fasst dann ins Leere ("invalid command name").
        """
        b = getattr(self, name, None)
        try:
            return b if b is not None and b.winfo_exists() else None
        except Exception:
            return None

    def _leere_log(self, sid: str) -> None:
        """Protokollfeld leeren, falls es gerade angezeigt wird."""
        feld = self._logs.get(sid)
        try:
            if feld is None or not feld.winfo_exists():
                return
            feld.configure(state=NORMAL)
            feld.delete("1.0", END)
            feld.configure(state=DISABLED)
        except Exception:
            pass

    def _melde(self, sid: str):
        """Protokoll-Rückruf für einen Arbeitsthread — schreibt nur in die
        Warteschlange, nie ins Fenster."""
        return lambda m: self._nachrichten.put(("log", sid, str(m)))

    # ------------------------------------------------------------------
    # Schritt 5 — Ablegen
    # ------------------------------------------------------------------
    def _baue_ablegen(self):
        f = ttk.Frame(self.rechts); f.pack(fill="both", expand=True)
        ttk.Label(f, wraplength=620, justify="left", padding=(8, 6),
                  text="Der fertige Buchordner wandert vom Arbeitsordner auf "
                       "den Ablageort. Kopiert wird, danach wird nachgeprüft "
                       "(Dateigröße am Ziel). Die örtliche Kopie bleibt stehen "
                       "— sie ist das Sicherheitsnetz, falls beim Übertragen "
                       "etwas hakt.").pack(anchor="w")

        r = ttk.Frame(f); r.pack(fill="x", padx=8, pady=4)
        ziel = core.share_ordner(self.ordner, self.cfg) if self.ordner else None
        ttk.Label(r, foreground="gray" if ziel else "#a00",
                  text=(f"Ziel: {ziel}" if ziel else
                        "⚠ Ablageort nicht erreichbar — bitte oben prüfen.")
                  ).pack(side="left")
        self.btn_ablegen = ttk.Button(r, text="Auf den Ablageort legen",
                                      command=self._lauf_ablegen,
                                      state="normal" if ziel else "disabled")
        self.btn_ablegen.pack(side="right")

        self.ablegen_log = self._log_feld(f, "ablegen")
        self._baue_checkliste(f, "ablegen")

    def _lauf_ablegen(self):
        self.btn_ablegen.configure(state="disabled")
        self.status.set("Lege ab …")
        quelle = self.ordner

        def arbeite():
            try:
                erg = core.schritt_ablegen(quelle, self.cfg,
                                           log=self._melde("ablegen"))
                self._nachrichten.put(("fertig", "ablegen", erg, None))
            except Exception as e:
                traceback.print_exc()
                self._nachrichten.put(("fertig", "ablegen", None, e))

        threading.Thread(target=arbeite, daemon=True).start()

    def _ablegen_fertig(self, erg, fehler):
        (b := self._knopf("btn_ablegen")) and b.configure(state="normal")
        if fehler:
            self._schreibe(self._logs.get("ablegen"), f"✗ {fehler}")
            self.status.set("Ablegen fehlgeschlagen.")
            messagebox.showerror("Ablegen", str(fehler))
            return
        if erg["gesichert"] is not None:
            self._schreibe(self._logs.get("ablegen"),
                           f"   Vorhandenes nach _alt/{erg['gesichert'].name}/ "
                           f"gesichert.")
        for p in erg["kopiert"]:
            self._schreibe(self._logs.get("ablegen"), f"   {Path(p).name}")
        if erg["uebersprungen"]:
            self._schreibe(self._logs.get("ablegen"),
                           "   übersprungen (Photoshop-Zwischendateien): "
                           + ", ".join(erg["uebersprungen"]))
        if erg["fehler"]:
            for f in erg["fehler"]:
                self._schreibe(self._logs.get("ablegen"), f"✗ {f}")
            self.status.set("Nicht alles angekommen — siehe Protokoll.")
            messagebox.showerror(
                "Unvollständig",
                f"{len(erg['fehler'])} Datei(en) kamen nicht an. Der "
                f"Arbeitsordner bleibt unangetastet.\n\n"
                + "\n".join(erg["fehler"][:5]))
            return
        self._schreibe(self._logs.get("ablegen"),
                       f"✓ {len(erg['kopiert'])} Datei(en) angekommen und "
                       f"nachgeprüft.")
        self._schreibe(self._logs.get("ablegen"),
                       f"   Der Arbeitsordner bleibt stehen: {erg['quelle']}")
        core.vermerke_lauf(self.stand, "ablegen", quelle=str(erg["quelle"]),
                           dateien=erg["kopiert"])
        core.speichere_stand(self.ordner, self.stand)
        core.spiegle_stand(self.ordner, erg["ziel"])
        self._gib_checkliste_frei("ablegen")
        self.status.set("Abgelegt — bitte auf dem Ablageort nachsehen und abhaken.")
        self._male_schrittliste()
        self._pruefe_tor()

    # ------------------------------------------------------------------
    # Schritt 4 — Presse-Dateien auf den Webserver
    # ------------------------------------------------------------------
    def _baue_presse(self):
        f = ttk.Frame(self.rechts); f.pack(fill="both", expand=True)
        ttk.Label(f, wraplength=640, justify="left", padding=(8, 6),
                  text="Die Knöpfe auf der Produktseite (Presseinfo, 2D, 3D, "
                       "Blick ins Buch) sind kein Produktfeld — das Template "
                       "zeigt sie, wenn die Datei am erwarteten Pfad liegt. "
                       "Hochladen und Anzeigen ist hier dasselbe."
                  ).pack(anchor="w")

        z = core.sftp_zugang(self.cfg)
        self.sftp_host = StringVar(value=z.get("host", ""))
        self.sftp_user = StringVar(value=z.get("benutzer", ""))
        r = ttk.Frame(f); r.pack(fill="x", padx=8, pady=3)
        ttk.Label(r, text="SFTP-Server:", width=13).pack(side="left")
        ttk.Entry(r, textvariable=self.sftp_host).pack(
            side="left", fill="x", expand=True)
        r2 = ttk.Frame(f); r2.pack(fill="x", padx=8, pady=3)
        ttk.Label(r2, text="Benutzer:", width=13).pack(side="left")
        ttk.Entry(r2, textvariable=self.sftp_user).pack(
            side="left", fill="x", expand=True)
        ttk.Button(r2, text="Passwort setzen…",
                   command=self._sftp_passwort_setzen).pack(side="left", padx=(6, 0))

        # DATEISYSTEM-Pfad, nicht der URL-Pfad — und je Umgebung ein anderer:
        # dev und prod liegen auf demselben Server nebeneinander.
        umg = sw.aktive_umgebung(core.cfg_shop(self.cfg))
        self.sftp_wurzel = StringVar(value=core.web_wurzel(self.cfg))
        rr = ttk.Frame(f); rr.pack(fill="x", padx=8, pady=2)
        ttk.Label(rr, text=f"Web-Wurzel ({umg}):", width=17).pack(side="left")
        ttk.Entry(rr, textvariable=self.sftp_wurzel).pack(
            side="left", fill="x", expand=True)
        ttk.Label(f, foreground="gray", padding=(8, 0),
                  text=f"→ {core.presse_basis(self.cfg)}/…  und  "
                       f"{core.newsletter_basis(self.cfg)}/…").pack(anchor="w")
        rp = ttk.Frame(f); rp.pack(fill="x", padx=8, pady=(2, 4))
        ttk.Label(rp, text="", width=13).pack(side="left")
        ttk.Button(rp, text="Verbindung prüfen (lädt nichts hoch)",
                   command=self._sftp_pruefen).pack(side="left")

        self.mit_pi = BooleanVar(
            value=bool(self.cfg.get("presseinfo_hochladen", True)))
        r3 = ttk.Frame(f); r3.pack(fill="x", padx=8, pady=(6, 2))
        ttk.Checkbutton(r3, text="Presseinfo mit hochladen (= sie wird auf der "
                                 "Produktseite angeboten)",
                        variable=self.mit_pi,
                        command=self._zeige_presse_liste).pack(side="left")

        self.presse_liste = Text(f, height=6, wrap="none", state=DISABLED,
                                 background="#fbfbfb")
        self.presse_liste.pack(fill="x", padx=8, pady=(2, 6))

        r4 = ttk.Frame(f); r4.pack(fill="x", padx=8, pady=(0, 4))
        self.btn_presse = ttk.Button(r4, text="Auf den Webserver laden",
                                     command=self._lauf_presse)
        self.btn_presse.pack(side="right")

        self.presse_log = self._log_feld(f, "presse")
        self._baue_checkliste(f, "presse")
        self._zeige_presse_liste()

    def _zeige_presse_liste(self):
        """Was hochginge — und was fehlt. Vor dem Klick, nicht danach."""
        if not (self.ordner and self.paar):
            return
        eintraege = core.presse_dateien(
            self.ordner, self.paar["sc"], self.cfg,
            mit_pi=self.mit_pi.get(), bib_pdf=self.bib_var.get() or None)
        self.presse_liste.configure(state=NORMAL)
        self.presse_liste.delete("1.0", END)
        for e in eintraege:
            marke = "✓" if e["da"] else ("⚠" if e["pflicht"] else "·")
            self.presse_liste.insert(
                END, f" {marke}  {e['art']:<18} {e['fern']}\n")
        fehlt = [e["art"] for e in eintraege if not e["da"]]
        if fehlt:
            self.presse_liste.insert(
                END, f"\n Fehlt im Buchordner: {', '.join(fehlt)}")
        self.presse_liste.configure(state=DISABLED)

    def _sftp_passwort_setzen(self):
        pw = simpledialog.askstring("SFTP-Passwort",
                                    f"Passwort für {self.sftp_user.get()}:",
                                    show="*", parent=self)
        if not pw:
            return
        master = simpledialog.askstring("Master-Passwort",
                                        "Passwort zum Verschlüsseln:",
                                        show="*", parent=self)
        if not master:
            return
        self._merke_sftp()
        sw.setze_secret(core.sftp_zugang(self.cfg), pw, master)
        core.speichere_config(self.cfg)
        self._sftp_pw = pw
        messagebox.showinfo("Gesetzt", "SFTP-Passwort verschlüsselt gespeichert.")

    def _merke_sftp(self):
        z = core.sftp_zugang(self.cfg)
        z["host"] = self.sftp_host.get().strip()
        z["benutzer"] = self.sftp_user.get().strip()
        if hasattr(self, "sftp_wurzel"):
            umg = sw.aktive_umgebung(core.cfg_shop(self.cfg))
            z.setdefault("web_wurzel", {})[umg] = \
                self.sftp_wurzel.get().strip().rstrip("/")

    def _sftp_pruefen(self):
        """Nachsehen, wo man landet — bevor irgendetwas hochgeht."""
        self._merke_sftp()
        core.speichere_config(self.cfg)
        if not self._sftp_entsperren():
            return
        self.status.set("Prüfe Verbindung …")

        def arbeite():
            try:
                bericht = core.sftp_pruefen(self.cfg, self._sftp_pw or "")
                self._nachrichten.put(("sftp_bericht", bericht, None))
            except Exception as e:
                self._nachrichten.put(("sftp_bericht", None, e))

        threading.Thread(target=arbeite, daemon=True).start()

    def _sftp_bericht(self, bericht, fehler):
        if fehler:
            self._schreibe(self._logs.get("presse"), f"✗ {fehler}")
            self.status.set("Verbindung fehlgeschlagen.")
            messagebox.showerror("SFTP", str(fehler))
            return
        core.speichere_config(self.cfg)      # ggf. gemerkter Serverschlüssel
        self._schreibe(self._logs.get("presse"),
                       f"✓ Angemeldet. Startverzeichnis: {bericht['start']}")
        if bericht["hostkey_neu"]:
            self._schreibe(self._logs.get("presse"),
                           "   Serverschlüssel gemerkt — künftig wird er "
                           "verglichen.")
        for name, titel in (("presse", "Presse-Ordner"),
                            ("newsletter", "Newsletter-Ordner")):
            b = bericht[name]
            if b["da"]:
                self._schreibe(self._logs.get("presse"),
                               f"✓ {titel}: {b['pfad']}  ({b['anzahl']} "
                               f"Einträge: {', '.join(b['inhalt'][:6])})")
            else:
                self._schreibe(self._logs.get("presse"),
                               f"✗ {titel}: {b['pfad']} — nicht da "
                               f"({b.get('grund', '')}). Pfad korrigieren.")
        self.status.set("Verbindung geprüft.")

    def _sftp_entsperren(self) -> bool:
        if self._sftp_pw:
            return True
        z = core.sftp_zugang(self.cfg)
        if not sw.hat_secret(z):
            messagebox.showwarning("Kein Passwort",
                                   "Bitte zuerst „Passwort setzen…“ drücken.")
            return False
        master = simpledialog.askstring("Master-Passwort", "Master-Passwort:",
                                        show="*", parent=self)
        if not master:
            return False
        try:
            self._sftp_pw = sw.hole_secret(z, master)
            return True
        except sw.PasswortFehler as e:
            messagebox.showerror("Passwort", str(e))
            return False

    def _lauf_presse(self):
        self._merke_sftp()
        core.speichere_config(self.cfg)
        if not self._sftp_entsperren():
            return
        eintraege = core.presse_dateien(
            self.ordner, self.paar["sc"], self.cfg,
            mit_pi=self.mit_pi.get(), bib_pdf=self.bib_var.get() or None)
        da = [e for e in eintraege if e["da"]]
        fehlt = [e["art"] for e in eintraege if not e["da"]]
        if not da:
            messagebox.showwarning(
                "Nichts da", "Keine der erwarteten Dateien liegt im "
                             "Buchordner.")
            return
        hinweis = (f"\n\nNicht dabei (fehlt im Buchordner):\n"
                   + "\n".join(f"• {a}" for a in fehlt)) if fehlt else ""
        if not messagebox.askokcancel(
                "Hochladen?",
                f"{len(da)} Datei(en) gehen auf {core.sftp_zugang(self.cfg)['host']}:"
                f"\n\n" + "\n".join(f"• {e['art']} → {e['fern']}" for e in da)
                + hinweis):
            return

        self.btn_presse.configure(state="disabled")
        self.status.set("Lade auf den Webserver …")
        mit_pi, bib = self.mit_pi.get(), self.bib_var.get() or None
        self.cfg["presseinfo_hochladen"] = mit_pi
        core.speichere_config(self.cfg)

        def arbeite():
            try:
                erg = core.schritt_presse(
                    self.ordner, self.paar["sc"], self.cfg,
                    passwort=self._sftp_pw or "", mit_pi=mit_pi, bib_pdf=bib,
                    log=self._melde("presse"))
                self._nachrichten.put(("fertig", "presse", erg, None))
            except Exception as e:
                traceback.print_exc()
                self._nachrichten.put(("fertig", "presse", None, e))

        threading.Thread(target=arbeite, daemon=True).start()

    def _presse_fertig(self, erg, fehler):
        (b := self._knopf("btn_presse")) and b.configure(state="normal")
        if fehler:
            self._schreibe(self._logs.get("presse"), f"✗ {fehler}")
            self.status.set("Hochladen fehlgeschlagen.")
            messagebox.showerror("Presse", str(fehler))
            return
        core.speichere_config(self.cfg)      # ggf. gemerkter Serverschlüssel
        for e in erg["geladen"]:
            self._schreibe(self._logs.get("presse"), f"   {e['art']} → {e['fern']}")
        for e in erg["fehlend"]:
            self._schreibe(self._logs.get("presse"),
                           f"⚠ nicht hochgeladen (fehlt): {e['art']}")
        if erg["fehler"]:
            for f in erg["fehler"]:
                self._schreibe(self._logs.get("presse"), f"✗ {f}")
            self.status.set("Nicht alles angekommen — siehe Protokoll.")
            messagebox.showerror("Unvollständig",
                                 "\n".join(erg["fehler"][:5]))
            return
        self._schreibe(self._logs.get("presse"),
                       f"✓ {len(erg['geladen'])} Datei(en) angekommen und "
                       f"nachgeprüft.")
        core.vermerke_lauf(self.stand, "presse",
                           quelle=core.sftp_zugang(self.cfg).get("host", ""),
                           dateien=[e["fern"] for e in erg["geladen"]],
                           hinweise=[f"nicht dabei: {e['art']}"
                                     for e in erg["fehlend"]])
        core.speichere_stand(self.ordner, self.stand)
        self._gib_checkliste_frei("presse")
        self._zeige_presse_liste()
        self.status.set("Hochgeladen — bitte die Produktseite ansehen und abhaken.")
        self._male_schrittliste()
        self._pruefe_tor()

    def _pumpe(self):
        """Läuft im Hauptthread und leert die Warteschlange.

        Alles, was ein Arbeitsthread der Oberfläche sagen will, geht hier
        durch — Protokollzeilen wie Ergebnisse.
        """
        fertig = {"cover": self._cover_fertig, "pibi": self._pibi_fertig,
                  "shop": self._shop_fertig, "ablegen": self._ablegen_fertig,
                  "presse": self._presse_fertig}
        try:
            while True:
                art, *rest = self._nachrichten.get_nowait()
                if art == "log":
                    sid, text = rest
                    feld = self._logs.get(sid)
                    if feld is not None and feld.winfo_exists():
                        self._schreibe(feld, text)
                elif art == "fertig":
                    sid, erg, fehler = rest
                    fertig[sid](erg, fehler)
                elif art == "verbunden":
                    self._verbunden(*rest)
                elif art == "vorschlag":
                    self._vorschlag(*rest)
                elif art == "shop_bereit":
                    self._shop_bereit(*rest)
                elif art == "sftp_bericht":
                    self._sftp_bericht(*rest)
        except queue.Empty:
            pass
        except Exception:
            traceback.print_exc()
        self.after(150, self._pumpe)

    def _schreibe(self, feld: Text, text: str):
        # Das Feld kann zerstört sein, wenn der Schritt noch lief und der
        # Bediener inzwischen weitergeblättert hat. Die Meldung ist dann
        # gegenstandslos, kein Fehler.
        if feld is None or not feld.winfo_exists():
            return
        feld.configure(state=NORMAL)
        feld.insert(END, text + "\n")
        feld.see(END)
        feld.configure(state=DISABLED)

    def _baue_checkliste(self, eltern, sid: str):
        """Die Häkchen des Schritts. Erst wenn alle sitzen, geht es weiter.

        Solange der Schritt nicht gelaufen ist, sind sie GESPERRT: eine
        Checkliste, die man vorher abhaken kann, prüft nichts — sie wäre nur
        ein Klick auf dem Weg nach vorn.
        """
        rahmen = ttk.LabelFrame(eltern, text="Bevor es weitergeht — bitte prüfen")
        rahmen.pack(fill="x", padx=8, pady=(6, 8))
        gelaufen = core.ist_gelaufen(self.stand, sid)
        gesetzt = set(self.stand.get("schritte", {}).get(sid, {}).get("haken", []))
        self._haken[sid] = []
        self._haken_felder = getattr(self, "_haken_felder", {})
        self._haken_felder[sid] = []
        for punkt in core.schritt(sid)["checkliste"]:
            v = BooleanVar(value=punkt in gesetzt)
            v.trace_add("write", lambda *_a, s=sid: self._haken_geaendert(s))
            cb = ttk.Checkbutton(rahmen, text=punkt, variable=v,
                                 state="normal" if gelaufen else "disabled")
            cb.pack(anchor="w", padx=8, pady=1)
            self._haken[sid].append(v)
            self._haken_felder[sid].append(cb)
        self._haken_hinweis = getattr(self, "_haken_hinweis", {})
        if not gelaufen:
            hinweis = ttk.Label(rahmen, foreground="gray",
                                text="(erst nach dem Lauf)")
            hinweis.pack(anchor="w", padx=8, pady=(2, 4))
            self._haken_hinweis[sid] = hinweis

    def _gib_checkliste_frei(self, sid: str):
        """Nach einem gelaufenen Schritt die Häkchen freischalten.

        Alles hier kann zerstört sein: läuft ein Schritt und man blättert
        weiter, ist das Panel weg, bevor der Abschluss ankommt. Beim nächsten
        Betreten wird es ohnehin neu gebaut — dann schon freigeschaltet, weil
        der Schritt inzwischen als gelaufen vermerkt ist.
        """
        for cb in getattr(self, "_haken_felder", {}).get(sid, []):
            try:
                if cb.winfo_exists():
                    cb.configure(state="normal")
            except Exception:
                pass
        # Der Hinweis „(erst nach dem Lauf)" hat sich damit erledigt.
        hinweis = getattr(self, "_haken_hinweis", {}).pop(sid, None)
        try:
            if hinweis is not None and hinweis.winfo_exists():
                hinweis.destroy()
        except Exception:
            pass

    def _haken_geaendert(self, sid: str):
        punkte = core.schritt(sid)["checkliste"]
        gesetzt = [p for p, v in zip(punkte, self._haken[sid]) if v.get()]
        core.setze_haken(self.stand, sid, gesetzt)
        if self.ordner:
            core.speichere_stand(self.ordner, self.stand)
            # Ist schon abgelegt worden, den Stand am Ziel nachziehen — sonst
            # stünde dort für immer "abgelegt, aber nichts geprüft".
            if core.ist_gelaufen(self.stand, "ablegen"):
                ziel = core.share_ordner(self.ordner, self.cfg)
                if ziel and ziel.is_dir():
                    core.spiegle_stand(self.ordner, ziel)
        self._male_schrittliste()
        self._pruefe_tor()

    def _pruefe_tor(self):
        """„Weiter“ nur, wenn der Schritt gelaufen UND abgehakt ist."""
        if self.aktiv == "buch":
            frei = self.paar is not None
        else:
            frei = core.ist_fertig(self.stand, self.aktiv)
        letzter = self.aktiv == core.SCHRITTE[-1]["id"]
        self.btn_weiter.configure(state="normal" if frei and not letzter
                                  else "disabled")
        self.btn_zurueck.configure(
            state="disabled" if self.aktiv == "buch" else "normal")

    def _reihenfolge(self) -> list[str]:
        return ["buch"] + [s["id"] for s in core.SCHRITTE]

    def _weiter(self):
        folge = self._reihenfolge()
        i = folge.index(self.aktiv)
        if i + 1 < len(folge):
            self._zeige_schritt(folge[i + 1])

    def _zurueck(self):
        folge = self._reihenfolge()
        i = folge.index(self.aktiv)
        if i > 0:
            self._zeige_schritt(folge[i - 1])

    def _springe(self, sid: str):
        """Zurückspringen ist immer erlaubt, vorspringen nur ins Erreichte."""
        folge = self._reihenfolge()
        if folge.index(sid) <= folge.index(self.aktiv):
            self._zeige_schritt(sid)
            return
        if sid != "buch" and not self.paar:
            messagebox.showinfo("Erst das Buch", "Bitte zuerst Schritt 0.")
            return
        vorher = folge[folge.index(sid) - 1]
        if vorher == "buch" or core.ist_fertig(self.stand, vorher):
            self._zeige_schritt(sid)
        else:
            wohin = ("Schritt 0" if vorher == "buch"
                     else f"„{core.schritt(vorher)['titel']}“")
            messagebox.showinfo(
                "Noch nicht dran",
                f"Bitte zuerst {wohin} abschließen — der Schritt muss gelaufen "
                f"und abgehakt sein.")

    # ------------------------------------------------------------------
    # Schritt 0 — Buch
    # ------------------------------------------------------------------
    def _baue_buch(self):
        f = ttk.Frame(self.rechts); f.pack(fill="both", expand=True)
        for label, var, cmd in (
                ("Umschlag-PDF:", self.pdf_var, self._waehle_pdf),
                ("ONIX-XML:", self.xml_var, self._waehle_xml)):
            r = ttk.Frame(f); r.pack(fill="x", padx=8, pady=4)
            ttk.Label(r, text=label, width=14).pack(side="left")
            ttk.Entry(r, textvariable=var).pack(side="left", fill="x", expand=True)
            ttk.Button(r, text="Auswählen…", command=cmd).pack(side="left", padx=(6, 0))

        # Das Blick-ins-Buch-PDF entsteht nicht im Durchgang — es wird von
        # Hand gebaut und kann alles Mögliche heißen. Deshalb hier auswählen,
        # damit es in Schritt 4 unter dem richtigen Namen hochgeht.
        r_bib = ttk.Frame(f); r_bib.pack(fill="x", padx=8, pady=4)
        ttk.Label(r_bib, text="Blick ins Buch:", width=14).pack(side="left")
        ttk.Entry(r_bib, textvariable=self.bib_var).pack(
            side="left", fill="x", expand=True)
        ttk.Button(r_bib, text="Auswählen…", command=self._waehle_bib).pack(
            side="left", padx=(6, 0))
        ttk.Label(r_bib, text="(optional)", foreground="gray").pack(
            side="left", padx=(8, 0))

        r = ttk.Frame(f); r.pack(fill="x", padx=8, pady=4)
        ttk.Label(r, text="Titel (Ordner):", width=14).pack(side="left")
        e = ttk.Entry(r, textvariable=self.titel_var)
        e.pack(side="left", fill="x", expand=True)
        e.bind("<KeyRelease>", lambda _ev: self._zeige_ordner())
        ttk.Button(r, text="Prüfen", command=self._pruefe).pack(
            side="left", padx=(6, 0))

        self.buch_log = self._log_feld(f, "buch")

    def _waehle_arbeit(self):
        p = filedialog.askdirectory(title="Arbeitsordner wählen",
                                    initialdir=self.arbeit_var.get() or None)
        if p:
            self.arbeit_var.set(p)
            self.cfg["arbeitsordner"] = p
            core.speichere_config(self.cfg)
            self._zeige_ordner()

    def _waehle_ablage(self):
        p = filedialog.askdirectory(title="Ablageort wählen",
                                    initialdir=self.ablage_var.get() or None)
        if p:
            self.ablage_var.set(p)
            self.cfg["ablageort"] = p
            core.speichere_config(self.cfg)
            self._zeige_ordner()

    def _waehle_pdf(self):
        p = filedialog.askopenfilename(
            title="Umschlag-PDF", filetypes=[("PDF", "*.pdf"), ("Alle", "*.*")],
            initialdir=self.cfg.get("last_pdf_dir") or None)
        if p:
            self.pdf_var.set(p)
            self.cfg["last_pdf_dir"] = str(Path(p).parent)

    def _waehle_bib(self):
        p = filedialog.askopenfilename(
            title="Blick ins Buch (PDF)",
            filetypes=[("PDF", "*.pdf"), ("Alle Dateien", "*.*")],
            initialdir=self.cfg.get("last_bib_dir") or None)
        if p:
            self.bib_var.set(p)
            self.cfg["last_bib_dir"] = str(Path(p).parent)
            if self.ordner:
                self.stand["bib_pdf"] = p
                core.speichere_stand(self.ordner, self.stand)

    def _waehle_xml(self):
        p = filedialog.askopenfilename(
            title="VLB-ONIX-XML", filetypes=[("XML", "*.xml"), ("Alle", "*.*")],
            initialdir=self.cfg.get("last_xml_dir") or None)
        if p:
            self.xml_var.set(p)
            self.cfg["last_xml_dir"] = str(Path(p).parent)

    def _pruefe(self):
        if not (self.pdf_var.get() and self.xml_var.get()):
            messagebox.showwarning("Unvollständig",
                                   "Bitte Umschlag-PDF und ONIX-XML wählen.")
            return
        self.cfg["arbeitsordner"] = self.arbeit_var.get().strip()
        self.cfg["ablageort"] = self.ablage_var.get().strip()
        core.speichere_config(self.cfg)
        if self.paar and self.paar.get("doc"):
            try:
                self.paar["doc"].close()
            except Exception:
                pass
        self.paar = None
        self._leere_log("buch")
        try:
            paar = core.pruefe_paar(self.pdf_var.get(), self.xml_var.get(), self.cfg)
        except core.PaarFehler as e:
            self._schreibe(self._logs.get("buch"), f"✗ {e}")
            self.status.set("PDF und ONIX gehören nicht zusammen.")
            messagebox.showerror("Passt nicht zusammen", str(e))
            self._male_schrittliste(); self._pruefe_tor()
            return
        except Exception as e:
            self._schreibe(self._logs.get("buch"), f"✗ {e}")
            traceback.print_exc()
            messagebox.showerror("Fehler beim Einlesen", str(e))
            return

        self.paar = paar
        fel = paar["felder"]
        self._schreibe(self._logs.get("buch"), f"✓ ISBN stimmt überein: {paar['isbn']}")
        self._schreibe(self._logs.get("buch"), f"   Titel        {fel['titel']}")
        if fel.get("untertitel"):
            self._schreibe(self._logs.get("buch"), f"   Untertitel   {fel['untertitel']}")
        self._schreibe(self._logs.get("buch"),
                       f"   Preis        {fel['preis_brutto']:.2f} {fel['waehrung']}")
        if paar["gemessen_cm"] and all(paar["onix_cm"]):
            g, o = paar["gemessen_cm"], paar["onix_cm"]
            self._schreibe(self._logs.get("buch"),
                           f"✓ Format      gemessen {g[0]:.1f} x {g[1]:.1f} cm "
                           f"(mit Beschnitt), ONIX {o[0]:.1f} x {o[1]:.1f} cm")
        for w in paar["warnungen"]:
            self._schreibe(self._logs.get("buch"), f"⚠ {w}")

        if not self.titel_var.get().strip():
            self.titel_var.set(fel["titel"])
        self._zeige_ordner()
        self.stand = core.lade_stand(self.ordner) if self.ordner else {"schritte": {}}
        if self.stand.get("bib_pdf") and not self.bib_var.get():
            self.bib_var.set(self.stand["bib_pdf"])
        self.status.set(f"Kurzcode {paar['sc']} — weiter zum Cover.")
        self._male_schrittliste()
        self._pruefe_tor()

    def _zeige_ordner(self):
        if not self.paar:
            return
        self.cfg["arbeitsordner"] = self.arbeit_var.get().strip()
        self.cfg["ablageort"] = self.ablage_var.get().strip()
        ordner, existiert = core.buchordner(self.paar["sc"],
                                            self.titel_var.get(), self.cfg)
        self.ordner = ordner
        # Nur der Ordnername: der volle Pfad steht bereits oben im Feld
        # „Arbeitsordner", und eine zweite lange Zeile drängt hier alles weg.
        self.ordner_var.set(f"Buchordner: {ordner.name}  "
                            f"({'vorhanden' if existiert else 'wird angelegt'})")

    @staticmethod
    def _oeffnen(pfad) -> None:
        """Datei oder Ordner mit dem Programm des Systems öffnen."""
        p = str(pfad)
        if sys.platform == "win32":
            os.startfile(p)                      # noqa: S606 — Windows-API
        elif sys.platform == "darwin":
            os.system(f'open "{p}"')
        else:
            os.system(f'xdg-open "{p}" >/dev/null 2>&1 &')

    def _ordner_oeffnen(self):
        if self.ordner:
            self._oeffnen(self.ordner)

    # ------------------------------------------------------------------
    # Schritt 1 — Cover
    # ------------------------------------------------------------------
    def _baue_cover(self):
        f = ttk.Frame(self.rechts); f.pack(fill="both", expand=True)
        self.cov_2d = BooleanVar(value=True)
        self.cov_3d = BooleanVar(value=True)
        r = ttk.Frame(f); r.pack(fill="x", padx=8, pady=6)
        ttk.Checkbutton(r, text="2D-Vorderseite", variable=self.cov_2d).pack(side="left")
        ttk.Checkbutton(r, text="3D-Mockup", variable=self.cov_3d).pack(side="left", padx=12)
        self.btn_cover = ttk.Button(r, text="Cover erzeugen", command=self._lauf_cover)
        self.btn_cover.pack(side="right")

        if sys.platform != "win32":
            ttk.Label(f, foreground="#960", wraplength=620, justify="left",
                      text="⚠ Kein Windows: der 3D-Zweig steuert Photoshop und "
                           "läuft hier nur als Trockenlauf. Es entstehen JSX und "
                           "Slot-PNGs, aber KEIN Mockup.").pack(
                anchor="w", padx=8, pady=(0, 4))

        self.cover_log = self._log_feld(f, "cover")
        self._baue_checkliste(f, "cover")

    def _lauf_cover(self):
        self.btn_cover.configure(state="disabled")
        self.status.set("Erzeuge Cover …")
        # Tk-Variablen NUR im Hauptthread lesen — aus einem Arbeitsthread
        # heraus wirft das „main thread is not in main loop".
        titel, mit_2d, mit_3d = (self.titel_var.get(), self.cov_2d.get(),
                                 self.cov_3d.get())

        def arbeite():
            try:
                erg = core.schritt_cover(
                    self.paar, titel, self.cfg,
                    mit_2d=mit_2d, mit_3d=mit_3d,
                    log=self._melde("cover"))
                self._nachrichten.put(("fertig", "cover", erg, None))
            except Exception as e:
                traceback.print_exc()
                self._nachrichten.put(("fertig", "cover", None, e))

        threading.Thread(target=arbeite, daemon=True).start()

    def _cover_fertig(self, erg, fehler):
        (b := self._knopf("btn_cover")) and b.configure(state="normal")
        if fehler:
            self._schreibe(self._logs.get("cover"), f"✗ {fehler}")
            self.status.set("Cover-Schritt fehlgeschlagen.")
            messagebox.showerror("Cover", str(fehler))
            return
        self.ordner = erg["out_dir"]
        for p in erg["erzeugt"]:
            self._schreibe(self._logs.get("cover"), f"   {Path(p).name}")
        for h in erg["hinweise"]:
            self._schreibe(self._logs.get("cover"), f"⚠ {h}")
        self._schreibe(self._logs.get("cover"), f"✓ {len(erg['erzeugt'])} Datei(en).")
        self.stand = core.lade_stand(self.ordner)
        core.vermerke_lauf(self.stand, "cover", quelle=self.paar["pdf_pfad"],
                           dateien=erg["erzeugt"], hinweise=erg["hinweise"])
        core.speichere_stand(self.ordner, self.stand)
        self._gib_checkliste_frei("cover")
        self.status.set("Cover erzeugt — bitte prüfen und abhaken.")

        # Die Checkliste fragt nach weißem Rand und heller Kante — beides
        # sieht man nur im Bild. Also gleich anbieten.
        #
        # Die Web-Fassung ist pixelgleich mit der Druckfassung, nur der
        # eingebettete DPI-Wert unterscheidet sich (siehe speichere_2d in
        # cover_previews). Zum Ansehen genügt eine je Paar.
        ccfg = core.cfg_cover(self.cfg)
        web, druck = (str(int(ccfg.get("dpi_web", 72))),
                      str(int(ccfg.get("dpi_print", 300))))
        bilder = []
        for pfad in erg["erzeugt"]:
            name = Path(pfad).name
            if name.startswith("_") or Path(name).suffix.lower() not in (
                    ".jpg", ".jpeg", ".png"):
                continue
            zwilling = Path(str(pfad).replace(f"_{web}_", f"_{druck}_"))
            if f"_{web}_" in name and zwilling.exists():
                continue
            bilder.append(pfad)

        if bilder and messagebox.askyesno(
                "Bilder ansehen?",
                "Jetzt öffnen, um Rand und Kanten zu prüfen?\n\n"
                + "\n".join(f"• {Path(b).name}" for b in bilder)
                + ("\n\n(Die 72-dpi-Fassungen sind pixelgleich und deshalb "
                   "nicht dabei.)" if len(bilder) < len(erg["erzeugt"]) else "")):
            for b in bilder:
                self._oeffnen(b)
        self._male_schrittliste()
        self._pruefe_tor()

    # ------------------------------------------------------------------
    # Schritt 2 — pi & bi
    # ------------------------------------------------------------------
    def _baue_pibi(self):
        f = ttk.Frame(self.rechts); f.pack(fill="both", expand=True)
        r = ttk.Frame(f); r.pack(fill="x", padx=8, pady=6)
        ttk.Label(r, text="Presse- und Buchinformation, je als .docx und .html "
                          "— das Cover kommt aus Schritt 1.",
                  wraplength=520, justify="left").pack(side="left")
        self.btn_pibi = ttk.Button(r, text="pi & bi erzeugen", command=self._lauf_pibi)
        self.btn_pibi.pack(side="right")
        self.pibi_log = self._log_feld(f, "pibi")
        self._baue_checkliste(f, "pibi")

    def _lauf_pibi(self):
        self.btn_pibi.configure(state="disabled")
        self.status.set("Erzeuge pi & bi …")

        def arbeite():
            try:
                dateien = core.schritt_pibi(
                    self.paar, self.ordner, self.cfg,
                    log=self._melde("pibi"))
                self._nachrichten.put(("fertig", "pibi", dateien, None))
            except Exception as e:
                traceback.print_exc()
                self._nachrichten.put(("fertig", "pibi", None, e))

        threading.Thread(target=arbeite, daemon=True).start()

    def _pibi_fertig(self, dateien, fehler):
        (b := self._knopf("btn_pibi")) and b.configure(state="normal")
        if fehler:
            self._schreibe(self._logs.get("pibi"), f"✗ {fehler}")
            messagebox.showerror("pi & bi", str(fehler))
            return
        for p in dateien:
            self._schreibe(self._logs.get("pibi"), f"   {Path(p).name}")
        self._schreibe(self._logs.get("pibi"), f"✓ {len(dateien)} Datei(en).")
        core.vermerke_lauf(self.stand, "pibi", quelle=self.paar["xml_pfad"],
                           dateien=dateien)
        core.speichere_stand(self.ordner, self.stand)
        self._gib_checkliste_frei("pibi")
        self.status.set("pi & bi erzeugt — in Word kürzen, dann abhaken.")

        # Der nächste Handgriff ist ohnehin Word: Werbetext kürzen und als PDF
        # exportieren. Also gleich anbieten, statt den Ordner suchen zu lassen.
        docs = [d for d in dateien if str(d).lower().endswith(".docx")]
        if docs and messagebox.askyesno(
                "In Word öffnen?",
                "Jetzt in Word öffnen, um den Werbetext auf eine Seite zu "
                "kürzen?\n\n"
                + "\n".join(f"• {Path(d).name}" for d in docs)
                + "\n\nDenk daran, anschließend als PDF zu exportieren — "
                  f"unter PI_{self.paar['sc']}.pdf im selben Ordner, sonst "
                  "fehlt sie in Schritt 4."):
            for d in docs:
                self._oeffnen(d)
        self._male_schrittliste()
        self._pruefe_tor()

    # ------------------------------------------------------------------
    # Schritt 3 — Shop
    # ------------------------------------------------------------------
    def _baue_shop(self):
        f = ttk.Frame(self.rechts); f.pack(fill="both", expand=True)
        scfg = core.cfg_shop(self.cfg)
        umg = sw.umgebung(scfg)

        r = ttk.Frame(f); r.pack(fill="x", padx=8, pady=4)
        ttk.Label(r, text="Umgebung:", width=11).pack(side="left")
        self.umg_var = StringVar(value=sw.aktive_umgebung(scfg))
        box = ttk.Combobox(r, textvariable=self.umg_var, state="readonly",
                           width=10, values=sw.umgebungs_namen(scfg))
        box.pack(side="left")
        box.bind("<<ComboboxSelected>>", self._wechsle_umgebung)
        self.shop_url = StringVar(value=umg.get("shop_url", ""))
        ttk.Entry(r, textvariable=self.shop_url).pack(
            side="left", fill="x", expand=True, padx=(8, 0))

        r2 = ttk.Frame(f); r2.pack(fill="x", padx=8, pady=4)
        ttk.Label(r2, text="Schlüssel:", width=11).pack(side="left")
        self.key_var = StringVar(value=umg.get("access_key_id", ""))
        ttk.Entry(r2, textvariable=self.key_var).pack(
            side="left", fill="x", expand=True)
        ttk.Button(r2, text="Secret setzen…", command=self._secret_setzen).pack(
            side="left", padx=(6, 0))
        ttk.Button(r2, text="Verbinden", command=self._verbinde).pack(
            side="left", padx=(6, 0))

        # Kategorien werden beim Verbinden gesucht und ohne Rückfrage gesetzt.
        # Nachsehen und ergänzen tut man ohnehin im Shopware-Backend — das
        # steht so in der Checkliste, und dafür geht es nachher von selbst auf.
        self.shop_log = self._log_feld(f, "shop")

        # Dry-Run und Anlegen stehen UNTER dem Protokoll: dort schaut man hin,
        # bevor man sendet, und nicht darüber.
        r3 = ttk.Frame(f); r3.pack(fill="x", padx=8, pady=4)
        self.shop_dry = BooleanVar(value=False)
        ttk.Checkbutton(r3, text="Dry-Run (nichts senden)",
                        variable=self.shop_dry).pack(side="left")
        self.btn_shop = ttk.Button(r3, text="Als Entwurf anlegen",
                                   command=self._lauf_shop)
        self.btn_shop.pack(side="right")

        self._baue_checkliste(f, "shop")

        # Wer schon verbunden ist und erst jetzt hier ankommt, bekam bisher
        # keinen Vorschlag — der lief nur beim Verbinden. Das war der Grund,
        # warum das Kategoriefeld im Backend leer blieb.
        if self._client and not self._kat_gewaehlt:
            self._schlage_kategorien_vor()

    def _wechsle_umgebung(self, _ev=None):
        scfg = core.cfg_shop(self.cfg)
        self._merke_zugang()
        scfg["aktive_umgebung"] = self.umg_var.get()
        core.speichere_config(self.cfg)
        umg = sw.umgebung(scfg)
        self.shop_url.set(umg.get("shop_url", ""))
        self.key_var.set(umg.get("access_key_id", ""))
        self._client = None
        self._secret = None
        self._kat_gewaehlt = {}      # gehören zur alten Umgebung

    def _merke_zugang(self):
        umg = sw.umgebung(core.cfg_shop(self.cfg))
        umg["shop_url"] = sw.normalisiere_url(self.shop_url.get())
        umg["access_key_id"] = self.key_var.get().strip()
        self.shop_url.set(umg["shop_url"])

    def _secret_setzen(self):
        s = simpledialog.askstring("Geheimer Schlüssel",
                                   "Shopware-Secret:", show="*", parent=self)
        if not s:
            return
        pw = simpledialog.askstring("Master-Passwort",
                                    "Passwort zum Verschlüsseln:", show="*",
                                    parent=self)
        if not pw:
            return
        sw.setze_secret(sw.umgebung(core.cfg_shop(self.cfg)), s, pw)
        core.speichere_config(self.cfg)
        self._secret = s
        messagebox.showinfo("Gesetzt", "Secret verschlüsselt gespeichert.")

    def _entsperren(self) -> bool:
        if self._secret:
            return True
        umg = sw.umgebung(core.cfg_shop(self.cfg))
        if not sw.hat_secret(umg):
            messagebox.showwarning("Kein Secret",
                                   "Bitte zuerst „Secret setzen…“ drücken.")
            return False
        pw = simpledialog.askstring("Master-Passwort", "Master-Passwort:",
                                    show="*", parent=self)
        if not pw:
            return False
        try:
            self._secret = sw.hole_secret(umg, pw)
            return True
        except sw.PasswortFehler as e:
            messagebox.showerror("Passwort", str(e))
            return False

    def _verbinde(self):
        self._merke_zugang()
        core.speichere_config(self.cfg)
        if not self._entsperren():
            return
        scfg = core.cfg_shop(self.cfg)
        umg = sw.umgebung(scfg)
        self.status.set("Verbinde …")

        def arbeite():
            try:
                version, vorlage = self._verbinden_jetzt()
                self._nachrichten.put(("verbunden", version, vorlage, None))
            except Exception as e:
                self._nachrichten.put(("verbunden", None, None, e))

        threading.Thread(target=arbeite, daemon=True).start()

    def _verbinden_jetzt(self) -> tuple[str, dict]:
        """Verbindung aufbauen und Zuordnungen holen — im Arbeitsthread.

        Steckt in einer eigenen Funktion, weil es zwei Wege hierher gibt: den
        Knopf „Verbinden“ und das Anlegen, das sich bei Bedarf selbst
        verbindet. Das Secret muss vorher entsperrt sein.
        """
        scfg = core.cfg_shop(self.cfg)
        umg = sw.umgebung(scfg)
        c = sw.ShopClient(umg.get("shop_url", ""),
                          umg.get("access_key_id", ""), self._secret,
                          tls_pruefen=bool(umg.get("tls_pruefen", True)))
        version = c.verbinde().get("version", "?")
        vorlage = c.vorlage_vom_bestand()
        for k in ("cms_page_id", "sales_channel_id", "visibility",
                  "manufacturer_id"):
            if vorlage.get(k):
                umg[k] = vorlage[k]
        for schluessel, eintraege, passt in (
                ("tax_id", c.steuersaetze(),
                 lambda e: abs(float(e["taxRate"]) - 7.0) < 0.01),
                ("currency_id", c.waehrungen(),
                 lambda e: e["isoCode"] == "EUR")):
            if not umg.get(schluessel):
                for e in eintraege:
                    if passt(e):
                        umg[schluessel] = e["id"]
                        if schluessel == "tax_id":
                            umg["tax_rate"] = float(e["taxRate"])
                        break
        try:
            kats = c.alle_kategorien()
            ziel = sw.schreibe_kategorien_cache(
                sw.aktive_umgebung(scfg), umg.get("shop_url", ""), kats)
            self._nachrichten.put(
                ("log", "shop", f"   {len(kats)} Kategorien gemerkt ({ziel.name})"))
        except sw.ShopFehler:
            pass
        self._client = c
        return version, vorlage

    def _verbunden(self, version, vorlage, fehler):
        if fehler:
            self.status.set("Nicht verbunden.")
            messagebox.showerror("Verbindung", str(fehler))
            return
        core.speichere_config(self.cfg)
        kanal = (vorlage or {}).get("sales_channel_name") or "⚠ keiner"
        self.status.set(f"Verbunden (Shopware {version}) — Verkaufskanal: {kanal}")
        self._schreibe(self._logs.get("shop"), f"✓ Verbunden mit Shopware {version}")
        self._schlage_kategorien_vor()

    # -- Kategorien ----------------------------------------------------
    def _schlage_kategorien_vor(self):
        if not (self._client and self.paar):
            return

        scfg = core.cfg_shop(self.cfg)
        gemerkt = sw.lade_kategorien_cache(sw.aktive_umgebung(scfg))

        def arbeite():
            # Aus dem Zwischenspeicher raten, wenn er da ist: das sind je Buch
            # ein Dutzend Abfragen weniger, und es funktioniert identisch.
            if gemerkt:
                suche = sw.cache_sucher(gemerkt)
            else:
                def suche(t):
                    try:
                        return self._client.kategorien_suchen(t)
                    except sw.ShopFehler:
                        return []
            treffer, fehlend = sw.kategorie_vorschlaege(
                self.paar["felder"], suche, scfg, log=self._melde("shop"))
            self._nachrichten.put(("vorschlag", treffer, fehlend))

        threading.Thread(target=arbeite, daemon=True).start()

    def _fehlende_personen(self) -> list[dict]:
        """Die Beteiligten hinter den nicht gefundenen Namen."""
        fel = (self.paar or {}).get("felder") or {}
        alle = (fel.get("herausgeber_teile") or []) + (fel.get("autoren_teile") or [])
        offen = set(self._kat_fehlend or [])
        return [pers for pers in alle if (pers.get("name") or "") in offen]

    def _vorschlag(self, treffer, fehlend):
        """Was gefunden wurde, wird gesetzt — ohne Rückfrage.

        Geraten wird nur über die Beteiligten; Sachkategorien und alles, was
        der Shop sonst noch braucht, trägt ein Mensch im Backend nach. Was
        NICHT gefunden wurde, steht trotzdem im Protokoll — sonst fiele es erst
        auf, wenn der Breadcrumb leer bleibt.
        """
        for k in treffer:
            self._kat_gewaehlt[k["id"]] = k.get("name") or k["id"]
            self._schreibe(self._logs.get("shop"), f"   Kategorie: {k.get('name')}")
        if not treffer:
            self._schreibe(self._logs.get("shop"), "   Keine Kategorie gefunden.")
        self._kat_fehlend = list(fehlend)
        for name in fehlend:
            self._schreibe(self._logs.get("shop"), f"⚠ ohne eigene Kategorie: {name}")
        if fehlend:
            self._schreibe(self._logs.get("shop"),
                           "   → beim Anlegen wird gefragt, ob sie entstehen "
                           "sollen.")

    # -- Anlegen -------------------------------------------------------
    def _lauf_shop(self):
        """Der Weg zum Entwurf — erst alle Rückfragen, dann ein Arbeitsgang.

        1. **Kategoriebaum auffrischen** und neu raten. Er kann seit dem
           Verbinden gewachsen sein, etwa weil beim letzten Buch eine
           Autorenkategorie entstand.
        2. Umgebung bestätigen.
        3. Überschreiben bestätigen, falls es das Buch schon gibt.
        4. Fehlende Autorenkategorien: anlegen?

        Alle Fenster im Hauptthread, alles Netz danach in EINEM Arbeitsgang.
        Sonst wechseln sich Rückfragen und Wartezeiten ab, und man weiß nie,
        ob noch etwas kommt.
        """
        if self.shop_dry.get():
            self._shop_senden(ueberschreiben=False, anlegen=[])
            return
        # Nicht verbunden? Dann eben jetzt — das Passwort braucht es ohnehin,
        # ein zusätzlicher Knopfdruck davor bringt niemandem etwas.
        if not self._entsperren():
            return
        self._merke_zugang()
        core.speichere_config(self.cfg)

        self.btn_shop.configure(state="disabled")
        self.status.set("Verbinde …" if not self._client
                        else "Frische Kategorien auf …")
        scfg = core.cfg_shop(self.cfg)
        name_umg = sw.aktive_umgebung(scfg)
        url = sw.umgebung(scfg).get("shop_url", "")
        nummer = self.paar["felder"]["isbn13_formatiert"]

        def arbeite():
            try:
                if self._client is None:
                    version, _ = self._verbinden_jetzt()
                    self._nachrichten.put(
                        ("log", "shop", f"✓ Verbunden mit Shopware {version}"))
                kats = self._client.alle_kategorien()
                sw.schreibe_kategorien_cache(name_umg, url, kats)
                treffer, fehlend = sw.kategorie_vorschlaege(
                    self.paar["felder"], sw.cache_sucher(kats), scfg,
                    log=self._melde("shop"))
                vorhanden = self._client.produkt_id_zu_nummer(nummer)
                self._nachrichten.put(
                    ("shop_bereit", kats, treffer, fehlend, vorhanden, None))
            except Exception as e:
                traceback.print_exc()
                self._nachrichten.put(
                    ("shop_bereit", None, None, None, None, e))

        threading.Thread(target=arbeite, daemon=True).start()

    def _shop_bereit(self, kats, treffer, fehlend, vorhanden, fehler):
        """Alle Rückfragen, hintereinander weg — im Hauptthread."""
        (b := self._knopf("btn_shop")) and b.configure(state="normal")
        if fehler:
            self._schreibe(self._logs.get("shop"), f"✗ {fehler}")
            self.status.set("Fehlgeschlagen.")
            messagebox.showerror("Shop", str(fehler))
            return

        self._vorschlag(treffer, fehlend)      # setzt _kat_gewaehlt/_kat_fehlend

        scfg = core.cfg_shop(self.cfg)
        name_umg = sw.aktive_umgebung(scfg)
        nummer = self.paar["felder"]["isbn13_formatiert"]
        warnung = ("\n\n⚠  PRODUKTIVSHOP — die Änderung ist echt!"
                   if sw.ist_produktiv(name_umg) else "")
        if not messagebox.askokcancel(
                "Wirklich anlegen?",
                f"Umgebung: {name_umg}\n{sw.umgebung(scfg).get('shop_url')}\n\n"
                f"Artikelnummer: {nummer}{warnung}"):
            return

        ueberschreiben = False
        if vorhanden:
            if not messagebox.askyesno(
                    "Buch existiert bereits",
                    f"{nummer} gibt es im Shop schon.\n\nBeim Überschreiben "
                    f"gehen dort gepflegte Angaben verloren (Name, "
                    f"Beschreibung, Preis, Bilder …).\n\nWirklich "
                    f"überschreiben?", icon="warning", default="no"):
                self._schreibe(self._logs.get("shop"), "Abgebrochen — nichts geändert.")
                return
            ueberschreiben = True

        anlegen = self._frage_kategorien(kats)
        self._shop_senden(ueberschreiben=ueberschreiben, anlegen=anlegen)

    def _frage_kategorien(self, kats) -> list[dict]:
        """Fehlende Autorenkategorien vorlegen. Rückgabe: was anzulegen ist.

        Was es schon gibt — auch unter einem ANDEREN Buchstaben — wird
        übernommen statt neu angelegt. Im Bestand hängt „Beiträge … Speyer"
        unter S statt unter B; ohne diese Prüfung entstünde der Eintrag ein
        zweites Mal, und die Bücher verteilten sich auf zwei Kategorien.
        """
        if not self._kat_fehlend:
            return []
        scfg = core.cfg_shop(self.cfg)
        plaene = [sw.plane_autorenkategorie(pers, kats, scfg)
                  for pers in self._fehlende_personen()]

        for pl in plaene:
            if pl["vorhanden"]:
                self._kat_gewaehlt[pl["vorhanden"]["id"]] = pl["vorhanden"]["name"]
                self._schreibe(self._logs.get("shop"),
                               f"✓ vorhanden, wird verwendet: {pl['pfad']}")
            elif not pl["moeglich"]:
                self._schreibe(self._logs.get("shop"), f"✗ {pl['name']}: {pl['grund']}")

        machbar = [pl for pl in plaene if pl["moeglich"]]
        if not machbar:
            return []
        liste = "\n".join(f"• {pl['pfad']}" for pl in machbar)
        warnung = ("\n\n⚠  Das steht danach in der Navigation des "
                   "PRODUKTIVSHOPS!"
                   if sw.ist_produktiv(sw.aktive_umgebung(scfg)) else "")
        if messagebox.askyesno(
                "Autorenkategorie fehlt",
                f"Für diese Beteiligten gibt es noch keine Kategorie:\n\n"
                f"{liste}\n\nAnlegen?{warnung}"):
            return machbar
        self._schreibe(self._logs.get("shop"),
                       "Kategorien nicht angelegt — das Buch bekommt sie nicht.")
        return []

    def _shop_senden(self, *, ueberschreiben: bool, anlegen: list[dict]):
        """Ein Arbeitsgang: erst die Kategorien anlegen, dann veröffentlichen."""
        self.btn_shop.configure(state="disabled")
        self.status.set("Sende …")
        dry = self.shop_dry.get()          # Tk-Variable, nur im Hauptthread
        kategorien = dict(self._kat_gewaehlt)

        def arbeite():
            try:
                for pl in anlegen:
                    vorlage = (self._client.kategorie_holen(pl["vorlage_id"])
                               if pl["vorlage_id"] else None)
                    neu_id = self._client.kategorie_anlegen(
                        pl["name"], pl["eltern_id"], vorlage)
                    kategorien[neu_id] = pl["name"]
                    self._nachrichten.put(
                        ("log", "shop", f"✓ angelegt: {pl['pfad']}"))
                if anlegen:
                    scfg = core.cfg_shop(self.cfg)
                    sw.schreibe_kategorien_cache(
                        sw.aktive_umgebung(scfg),
                        sw.umgebung(scfg).get("shop_url", ""),
                        self._client.alle_kategorien())

                erg = core.schritt_shop(
                    self.paar, self.ordner, self.cfg,
                    secret=self._secret or "",
                    kategorien=sorted(kategorien),
                    dry_run=dry, ueberschreiben=ueberschreiben,
                    log=self._melde("shop"))
                self._nachrichten.put(("fertig", "shop", erg, None))
            except Exception as e:
                traceback.print_exc()
                self._nachrichten.put(("fertig", "shop", None, e))

        threading.Thread(target=arbeite, daemon=True).start()

    def _shop_fertig(self, erg, fehler):
        (b := self._knopf("btn_shop")) and b.configure(state="normal")
        if fehler:
            # Überschrieben wird nach vorheriger Frage; kommt der Fehler
            # trotzdem, hat sich zwischen Prüfen und Senden etwas geändert.
            if isinstance(fehler, sw.ProduktExistiert):
                self._schreibe(self._logs.get("shop"),
                               f"✗ {fehler.nummer} wurde zwischenzeitlich "
                               f"angelegt — nichts geändert.")
                messagebox.showwarning("Inzwischen vorhanden", str(fehler))
                return
            self._schreibe(self._logs.get("shop"), f"✗ {fehler}")
            messagebox.showerror("Shop", str(fehler))
            return
        pl = erg["payload"]
        self._schreibe(self._logs.get("shop"),
                       f"✓ {pl['productNumber']} — {pl['name']}")
        # Nachgesehen und ergänzt wird ohnehin im Backend — also gleich hin.
        # Beim Dry-Run gibt es nichts zu öffnen, da wurde nichts angelegt.
        if erg.get("admin_url"):
            self._schreibe(self._logs.get("shop"), f"   {erg['admin_url']}")
            try:
                webbrowser.open(erg["admin_url"])
            except Exception:
                pass                  # kein Browser ist kein Grund zu scheitern
        core.vermerke_lauf(self.stand, "shop", quelle=self.paar["xml_pfad"],
                           dateien=[m.get("datei", "") for m in erg.get("medien", [])])
        core.speichere_stand(self.ordner, self.stand)
        self._gib_checkliste_frei("shop")
        self.status.set("Im Shop angelegt — bitte im Admin prüfen und abhaken.")
        self._male_schrittliste()
        self._pruefe_tor()



if __name__ == "__main__":
    main()
