"""Buchdurchgang — Oberfläche.

Ein Fenster, drei Schritte, dazwischen Tore:

    0. Buch wählen   Umschlag-PDF + ONIX-XML, beide Proben
    1. Cover         cover_previews
    2. pi & bi       pi_bi_generator
    3. Shop          shopware_publisher

**Weiter** wird erst frei, wenn alle Häkchen des Schritts gesetzt sind. Die
Häkchen liegen in ``durchgang.json`` im Buchordner — nicht beim Werkzeug —,
damit auch Wochen später und an einem anderen Rechner erkennbar bleibt, was
geprüft wurde.
"""

from __future__ import annotations

import os
import queue
import sys
import threading
import traceback
from pathlib import Path
from tkinter import (Tk, filedialog, messagebox, simpledialog, StringVar,
                     BooleanVar, Text, Listbox, END, DISABLED, NORMAL,
                     EXTENDED)
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
        self._kat_gewaehlt: dict[str, str] = {}
        self._kat_sicht: list[dict] = []
        self._kat_liste_daten: list[dict] = []
        self._kat_job = None
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
        self.status = StringVar(value="Umschlag-PDF und ONIX-XML wählen.")
        self.kat_such = StringVar()

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
                 "dorthin wandert der fertige Ordner in Schritt 4")):
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
        ttk.Label(fuss, textvariable=self.ordner_var,
                  foreground="gray").pack(side="left")
        self.btn_weiter = ttk.Button(fuss, text="Weiter ▸",
                                     command=self._weiter, state="disabled")
        self.btn_weiter.pack(side="right")
        ttk.Button(fuss, text="Ordner öffnen",
                   command=self._ordner_oeffnen).pack(side="right", padx=6)
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
         "ablegen": self._baue_ablegen}[sid]()
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

    def _melde(self, sid: str):
        """Protokoll-Rückruf für einen Arbeitsthread — schreibt nur in die
        Warteschlange, nie ins Fenster."""
        return lambda m: self._nachrichten.put(("log", sid, str(m)))

    # ------------------------------------------------------------------
    # Schritt 4 — Ablegen
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
        self.btn_ablegen.configure(state="normal")
        if fehler:
            self._schreibe(self.ablegen_log, f"✗ {fehler}")
            self.status.set("Ablegen fehlgeschlagen.")
            messagebox.showerror("Ablegen", str(fehler))
            return
        if erg["gesichert"] is not None:
            self._schreibe(self.ablegen_log,
                           f"   Vorhandenes nach _alt/{erg['gesichert'].name}/ "
                           f"gesichert.")
        for p in erg["kopiert"]:
            self._schreibe(self.ablegen_log, f"   {Path(p).name}")
        if erg["uebersprungen"]:
            self._schreibe(self.ablegen_log,
                           "   übersprungen (Photoshop-Zwischendateien): "
                           + ", ".join(erg["uebersprungen"]))
        if erg["fehler"]:
            for f in erg["fehler"]:
                self._schreibe(self.ablegen_log, f"✗ {f}")
            self.status.set("Nicht alles angekommen — siehe Protokoll.")
            messagebox.showerror(
                "Unvollständig",
                f"{len(erg['fehler'])} Datei(en) kamen nicht an. Der "
                f"Arbeitsordner bleibt unangetastet.\n\n"
                + "\n".join(erg["fehler"][:5]))
            return
        self._schreibe(self.ablegen_log,
                       f"✓ {len(erg['kopiert'])} Datei(en) angekommen und "
                       f"nachgeprüft.")
        self._schreibe(self.ablegen_log,
                       f"   Der Arbeitsordner bleibt stehen: {erg['quelle']}")
        core.vermerke_lauf(self.stand, "ablegen", quelle=str(erg["quelle"]),
                           dateien=erg["kopiert"])
        core.speichere_stand(self.ordner, self.stand)
        core.spiegle_stand(self.ordner, erg["ziel"])
        self._gib_checkliste_frei("ablegen")
        self.status.set("Abgelegt — bitte auf dem Ablageort nachsehen und abhaken.")
        self._male_schrittliste()
        self._pruefe_tor()

    def _pumpe(self):
        """Läuft im Hauptthread und leert die Warteschlange.

        Alles, was ein Arbeitsthread der Oberfläche sagen will, geht hier
        durch — Protokollzeilen wie Ergebnisse.
        """
        fertig = {"cover": self._cover_fertig, "pibi": self._pibi_fertig,
                  "shop": self._shop_fertig, "ablegen": self._ablegen_fertig}
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
                elif art == "kategorien":
                    self._zeige_kategorien(*rest)
        except queue.Empty:
            pass
        except Exception:
            traceback.print_exc()
        self.after(150, self._pumpe)

    def _schreibe(self, feld: Text, text: str):
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
        if not gelaufen:
            ttk.Label(rahmen, foreground="gray",
                      text="(erst nach dem Lauf)").pack(anchor="w", padx=8,
                                                        pady=(2, 4))

    def _gib_checkliste_frei(self, sid: str):
        """Nach einem gelaufenen Schritt die Häkchen freischalten."""
        for cb in getattr(self, "_haken_felder", {}).get(sid, []):
            cb.configure(state="normal")

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

    def _reihenfolge(self) -> list[str]:
        return ["buch"] + [s["id"] for s in core.SCHRITTE]

    def _weiter(self):
        folge = self._reihenfolge()
        i = folge.index(self.aktiv)
        if i + 1 < len(folge):
            self._zeige_schritt(folge[i + 1])

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
        self.buch_log.configure(state=NORMAL)
        self.buch_log.delete("1.0", END)
        self.buch_log.configure(state=DISABLED)
        try:
            paar = core.pruefe_paar(self.pdf_var.get(), self.xml_var.get(), self.cfg)
        except core.PaarFehler as e:
            self._schreibe(self.buch_log, f"✗ {e}")
            self.status.set("PDF und ONIX gehören nicht zusammen.")
            messagebox.showerror("Passt nicht zusammen", str(e))
            self._male_schrittliste(); self._pruefe_tor()
            return
        except Exception as e:
            self._schreibe(self.buch_log, f"✗ {e}")
            traceback.print_exc()
            messagebox.showerror("Fehler beim Einlesen", str(e))
            return

        self.paar = paar
        fel = paar["felder"]
        self._schreibe(self.buch_log, f"✓ ISBN stimmt überein: {paar['isbn']}")
        self._schreibe(self.buch_log, f"   Titel        {fel['titel']}")
        if fel.get("untertitel"):
            self._schreibe(self.buch_log, f"   Untertitel   {fel['untertitel']}")
        self._schreibe(self.buch_log,
                       f"   Preis        {fel['preis_brutto']:.2f} {fel['waehrung']}")
        if paar["gemessen_cm"] and all(paar["onix_cm"]):
            g, o = paar["gemessen_cm"], paar["onix_cm"]
            self._schreibe(self.buch_log,
                           f"✓ Format      gemessen {g[0]:.1f} x {g[1]:.1f} cm "
                           f"(mit Beschnitt), ONIX {o[0]:.1f} x {o[1]:.1f} cm")
        for w in paar["warnungen"]:
            self._schreibe(self.buch_log, f"⚠ {w}")

        if not self.titel_var.get().strip():
            self.titel_var.set(fel["titel"])
        self._zeige_ordner()
        self.stand = core.lade_stand(self.ordner) if self.ordner else {"schritte": {}}
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
        self.ordner_var.set(f"Arbeitsordner: {ordner}  "
                            f"({'vorhanden' if existiert else 'wird angelegt'})")

    def _ordner_oeffnen(self):
        if not self.ordner:
            return
        p = str(self.ordner)
        if sys.platform == "win32":
            os.startfile(p)
        elif sys.platform == "darwin":
            os.system(f'open "{p}"')
        else:
            os.system(f'xdg-open "{p}" >/dev/null 2>&1 &')

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
        self.btn_cover.configure(state="normal")
        if fehler:
            self._schreibe(self.cover_log, f"✗ {fehler}")
            self.status.set("Cover-Schritt fehlgeschlagen.")
            messagebox.showerror("Cover", str(fehler))
            return
        self.ordner = erg["out_dir"]
        for p in erg["erzeugt"]:
            self._schreibe(self.cover_log, f"   {Path(p).name}")
        for h in erg["hinweise"]:
            self._schreibe(self.cover_log, f"⚠ {h}")
        self._schreibe(self.cover_log, f"✓ {len(erg['erzeugt'])} Datei(en).")
        self.stand = core.lade_stand(self.ordner)
        core.vermerke_lauf(self.stand, "cover", quelle=self.paar["pdf_pfad"],
                           dateien=erg["erzeugt"], hinweise=erg["hinweise"])
        core.speichere_stand(self.ordner, self.stand)
        self._gib_checkliste_frei("cover")
        self.status.set("Cover erzeugt — bitte prüfen und abhaken.")
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
        self.btn_pibi.configure(state="normal")
        if fehler:
            self._schreibe(self.pibi_log, f"✗ {fehler}")
            messagebox.showerror("pi & bi", str(fehler))
            return
        for p in dateien:
            self._schreibe(self.pibi_log, f"   {Path(p).name}")
        self._schreibe(self.pibi_log, f"✓ {len(dateien)} Datei(en).")
        core.vermerke_lauf(self.stand, "pibi", quelle=self.paar["xml_pfad"],
                           dateien=dateien)
        core.speichere_stand(self.ordner, self.stand)
        self._gib_checkliste_frei("pibi")
        self.status.set("pi & bi erzeugt — in Word kürzen, dann abhaken.")
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

        # Zweites Abtippen ist unnötig: Schlüssel und verschlüsseltes Secret
        # lassen sich aus dem ShopwarePublisher übernehmen. Das Master-Passwort
        # bleibt dabei aussen vor — es entsperrt hinterher genauso.
        r2b = ttk.Frame(f); r2b.pack(fill="x", padx=8, pady=(0, 4))
        ttk.Label(r2b, text="", width=11).pack(side="left")
        ttk.Button(r2b, text="Zugang aus dem ShopwarePublisher übernehmen",
                   command=self._zugang_uebernehmen).pack(side="left")

        # Kategorien
        k = ttk.LabelFrame(f, text="Kategorien (je Person eine — wird gesucht "
                                   "und vorgeschlagen)")
        k.pack(fill="x", padx=8, pady=(6, 4))
        such = ttk.Entry(k, textvariable=self.kat_such)
        such.pack(fill="x", padx=6, pady=(6, 2))
        such.bind("<KeyRelease>", self._kat_angestossen)
        self.kat_liste = Listbox(k, selectmode=EXTENDED, height=4,
                                 exportselection=False)
        self.kat_liste.pack(fill="x", padx=6, pady=(0, 6))
        self.kat_liste.bind("<<ListboxSelect>>", self._kat_merken)

        r3 = ttk.Frame(f); r3.pack(fill="x", padx=8, pady=4)
        self.shop_dry = BooleanVar(value=False)
        ttk.Checkbutton(r3, text="Dry-Run (nichts senden)",
                        variable=self.shop_dry).pack(side="left")
        self.btn_shop = ttk.Button(r3, text="Als Entwurf anlegen",
                                   command=self._lauf_shop)
        self.btn_shop.pack(side="right")

        self.shop_log = self._log_feld(f, "shop")
        self._baue_checkliste(f, "shop")
        self._zeige_kategorien()

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
        self._kat_gewaehlt = {}
        self._zeige_kategorien()

    def _merke_zugang(self):
        umg = sw.umgebung(core.cfg_shop(self.cfg))
        umg["shop_url"] = sw.normalisiere_url(self.shop_url.get())
        umg["access_key_id"] = self.key_var.get().strip()
        self.shop_url.set(umg["shop_url"])

    def _zugang_uebernehmen(self):
        try:
            meldung = core.uebernimm_shop_zugang(self.cfg)
        except RuntimeError as e:
            messagebox.showerror("Nicht gefunden", str(e))
            return
        core.speichere_config(self.cfg)
        self._secret = None            # gehört zur alten Umgebung
        self._client = None
        scfg = core.cfg_shop(self.cfg)
        self.umg_var.set(sw.aktive_umgebung(scfg))
        umg = sw.umgebung(scfg)
        self.shop_url.set(umg.get("shop_url", ""))
        self.key_var.set(umg.get("access_key_id", ""))
        self._schreibe(self.shop_log, f"✓ {meldung}")
        self._schreibe(self.shop_log,
                       "   Das Master-Passwort wurde NICHT übernommen — es "
                       "entsperrt beim Verbinden genauso wie im Publisher.")
        self.status.set("Zugang übernommen — jetzt „Verbinden“.")

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
                self._client = c
                self._nachrichten.put(("verbunden", version, vorlage, None))
            except Exception as e:
                self._nachrichten.put(("verbunden", None, None, e))

        threading.Thread(target=arbeite, daemon=True).start()

    def _verbunden(self, version, vorlage, fehler):
        if fehler:
            self.status.set("Nicht verbunden.")
            messagebox.showerror("Verbindung", str(fehler))
            return
        core.speichere_config(self.cfg)
        kanal = (vorlage or {}).get("sales_channel_name") or "⚠ keiner"
        self.status.set(f"Verbunden (Shopware {version}) — Verkaufskanal: {kanal}")
        self._schreibe(self.shop_log, f"✓ Verbunden mit Shopware {version}")
        self._schlage_kategorien_vor()

    # -- Kategorien ----------------------------------------------------
    def _kat_angestossen(self, _ev=None):
        if self._kat_job:
            self.after_cancel(self._kat_job)
        self._kat_job = self.after(400, self._kat_suchen)

    def _kat_suchen(self):
        self._kat_job = None
        if not self._client:
            return
        text = self.kat_such.get().strip()

        def arbeite():
            try:
                treffer = self._client.kategorien_suchen(text)
            except sw.ShopFehler:
                treffer = []
            self._nachrichten.put(("kategorien", treffer))

        threading.Thread(target=arbeite, daemon=True).start()

    def _schlage_kategorien_vor(self):
        if not (self._client and self.paar):
            return

        def arbeite():
            def suche(t):
                try:
                    return self._client.kategorien_suchen(t)
                except sw.ShopFehler:
                    return []
            treffer, fehlend = sw.kategorie_vorschlaege(self.paar["felder"], suche)
            self._nachrichten.put(("vorschlag", treffer, fehlend))

        threading.Thread(target=arbeite, daemon=True).start()

    def _vorschlag(self, treffer, fehlend):
        for k in treffer:
            self._kat_gewaehlt[k["id"]] = k.get("name") or k["id"]
            self._schreibe(self.shop_log, f"   Kategorie: {k.get('name')}")
        for name in fehlend:
            self._schreibe(self.shop_log,
                           f"⚠ ohne eigene Kategorie: {name} — bitte von Hand "
                           f"wählen oder im Admin anlegen.")
        self._zeige_kategorien(treffer)

    def _zeige_kategorien(self, treffer=None):
        if treffer is not None:
            self._kat_sicht = treffer
        gewaehlt = [{"id": i, "name": n} for i, n in self._kat_gewaehlt.items()]
        ids = {k["id"] for k in gewaehlt}
        liste = gewaehlt + [k for k in self._kat_sicht if k["id"] not in ids]
        self.kat_liste.delete(0, END)
        for i, k in enumerate(liste):
            marke = "✓ " if k["id"] in self._kat_gewaehlt else "   "
            self.kat_liste.insert(END, marke + (k.get("name") or k["id"]))
            if k["id"] in self._kat_gewaehlt:
                self.kat_liste.selection_set(i)
        self._kat_liste_daten = liste

    def _kat_merken(self, _ev=None):
        liste = self._kat_liste_daten
        sichtbar = {k["id"] for k in liste}
        gewaehlt = {liste[i]["id"]: liste[i].get("name") or liste[i]["id"]
                    for i in self.kat_liste.curselection() if i < len(liste)}
        behalten = {i: n for i, n in self._kat_gewaehlt.items() if i not in sichtbar}
        self._kat_gewaehlt = {**behalten, **gewaehlt}

    # -- Anlegen -------------------------------------------------------
    def _lauf_shop(self):
        if not self.shop_dry.get():
            if not self._client and not self._entsperren():
                return
            scfg = core.cfg_shop(self.cfg)
            name = sw.aktive_umgebung(scfg)
            warnung = ("\n\n⚠  PRODUKTIVSHOP — die Änderung ist echt!"
                       if sw.ist_produktiv(name) else "")
            if not messagebox.askokcancel(
                    "Wirklich anlegen?",
                    f"Umgebung: {name}\n{sw.umgebung(scfg).get('shop_url')}\n\n"
                    f"Artikelnummer: {self.paar['felder']['isbn13_formatiert']}"
                    f"{warnung}"):
                return
            if not self._kat_gewaehlt and not messagebox.askyesno(
                    "Ohne Kategorie?",
                    "Es ist keine Kategorie gewählt. Das Buch bekäme im Shop "
                    "keinen Breadcrumb.\n\nTrotzdem anlegen?",
                    icon="warning", default="no"):
                return

        self.btn_shop.configure(state="disabled")
        self.status.set("Sende …")
        dry = self.shop_dry.get()          # Tk-Variable, nur im Hauptthread

        def arbeite():
            try:
                erg = core.schritt_shop(
                    self.paar, self.ordner, self.cfg,
                    secret=self._secret or "",
                    kategorien=sorted(self._kat_gewaehlt),
                    dry_run=dry,
                    log=self._melde("shop"))
                self._nachrichten.put(("fertig", "shop", erg, None))
            except Exception as e:
                traceback.print_exc()
                self._nachrichten.put(("fertig", "shop", None, e))

        threading.Thread(target=arbeite, daemon=True).start()

    def _shop_fertig(self, erg, fehler):
        self.btn_shop.configure(state="normal")
        if fehler:
            if isinstance(fehler, sw.ProduktExistiert):
                if messagebox.askyesno(
                        "Buch existiert bereits",
                        f"{fehler.nummer} gibt es im Shop schon und wurde NICHT "
                        f"verändert.\n\nBeim Überschreiben gehen dort gepflegte "
                        f"Angaben verloren.\n\nWirklich überschreiben?",
                        icon="warning", default="no"):
                    self._ueberschreibe()
                    return
                self._schreibe(self.shop_log, "Abgebrochen — nichts geändert.")
                return
            self._schreibe(self.shop_log, f"✗ {fehler}")
            messagebox.showerror("Shop", str(fehler))
            return
        pl = erg["payload"]
        self._schreibe(self.shop_log,
                       f"✓ {pl['productNumber']} — {pl['name']}")
        if erg.get("admin_url"):
            self._schreibe(self.shop_log, f"   {erg['admin_url']}")
        core.vermerke_lauf(self.stand, "shop", quelle=self.paar["xml_pfad"],
                           dateien=[m.get("datei", "") for m in erg.get("medien", [])])
        core.speichere_stand(self.ordner, self.stand)
        self._gib_checkliste_frei("shop")
        self.status.set("Im Shop angelegt — bitte im Admin prüfen und abhaken.")
        self._male_schrittliste()
        self._pruefe_tor()

    def _ueberschreibe(self):
        self.btn_shop.configure(state="disabled")
        dry = self.shop_dry.get()          # Tk-Variable, nur im Hauptthread

        def arbeite():
            try:
                erg = core.schritt_shop(
                    self.paar, self.ordner, self.cfg, secret=self._secret or "",
                    kategorien=sorted(self._kat_gewaehlt),
                    dry_run=dry, ueberschreiben=True,
                    log=self._melde("shop"))
                self._nachrichten.put(("fertig", "shop", erg, None))
            except Exception as e:
                self._nachrichten.put(("fertig", "shop", None, e))

        threading.Thread(target=arbeite, daemon=True).start()

def main():
    App().mainloop()


if __name__ == "__main__":
    main()
