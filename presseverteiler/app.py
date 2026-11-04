#!/usr/bin/env python3
"""
Presseverteiler (Tkinter-GUI)
=============================

Je Buch ein paar Schlagworte ankreuzen, und die Liste der Pressekontakte
steht da — statt einer halben Stunde in der Access-Liste.

Zwei Reiter:
- Auswahl je Buch : Regionen und Themen des Buchs ankreuzen, Treffer prüfen,
                    E-Mails kopieren oder die Liste für den Serienbrief speichern.
- Einordnen       : festlegen, wofür jeder Pressekontakt zuständig ist.

Dateien im App-Ordner (neben der .exe):
- config.json            : zuletzt benutzter Access-Export
- Presseverteiler.xlsx   : wer wofür zuständig ist (in Excel lesbar)
- Presseversand/         : je Buch eine Liste samt Protokoll

Bauen als .exe (Windows):
    pyinstaller presseverteiler/Presseverteiler.spec
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import traceback
from collections import defaultdict
from pathlib import Path
from tkinter import (Tk, Toplevel, StringVar, BooleanVar, Text, END,
                     filedialog, messagebox, ttk)

from presseverteiler import core


HAKEN_AN, HAKEN_AUS = "☑", "☐"

ANSICHTEN = ["alle", "ungeprüft", "ungeprüft und unsicher", "geprüft",
             "bekommt nie etwas", "immer"]

# Breite der Spalte „Zuständigkeit“ im Reiter „Einordnen“: Platz für die
# Bäume der Regionen und Themen nebeneinander.
BREITE_ZUSTAENDIGKEIT = 780


def _oeffne(pfad: Path):
    """Datei oder Ordner mit dem Standardprogramm öffnen.

    Unter Windows tut es der Explorer, unter Linux xdg-open — entwickelt
    wird hier, eingesetzt dort.
    """
    if sys.platform == "win32":
        os.startfile(pfad)                                  # noqa: S606
    else:
        subprocess.Popen(["xdg-open", str(pfad)])


class SchlagwortBaum(ttk.Frame):
    """Die Schlagworte einer Achse als Baum zum Ankreuzen.

    Ein Schlagwort mit mehreren Obergebieten steht unter jedem davon (der
    Kraichgau unter drei Landkreisen) — angekreuzt ist es dann überall.

    In beiden Bäumen geht es von oben nach unten, nur auf zwei Arten:

    * **Buch** (`fuer_buch`): Bei den Regionen kreuzt ein Klick alle
      Teilgebiete mit an (Landkreis Karlsruhe → Kraichgau, Pfinzgau), ein
      zweiter wählt den ganzen Ast wieder ab. Einzelne Teilgebiete lassen sich
      danach abwählen. Bei den Themen nicht — warum, steht bei
      core.MIT_TEILGEBIETEN. Die Obergebiete, die das Buch ebenfalls erreicht,
      werden bewusst nicht markiert: Ein hervorgehobener Knoten „Geschichte“
      las sich, als gehöre der ganze Ast dazu, samt Kirchengeschichte. Das
      sagt stattdessen der Satz über der Trefferliste.
    * **Pressekontakt**: Die Unterbegriffe werden nur als „inbegriffen“
      markiert, nicht angekreuzt. Wer für Geschichte zuständig ist, bekommt
      ohnehin alles darunter. Ausgeschrieben in die Mappe stünden sonst bei
      jeder Redaktion zwanzig Schlagworte statt einem.
    """

    def __init__(self, master, titel: str, fuer_buch: bool,
                 bei_aenderung=None, hoehe: int = 10):
        super().__init__(master)
        self.fuer_buch = fuer_buch
        self.achse = ""
        self.bei_aenderung = bei_aenderung
        self.schlagworte: dict[str, core.Schlagwort] = {}
        self.kinder: dict[str, list[str]] = defaultdict(list)
        self.name_von: dict[str, str] = {}
        self.iids_von: dict[str, list[str]] = defaultdict(list)
        self.gewaehlt_set: set[str] = set()

        kopf = ttk.Frame(self)
        kopf.pack(fill="x")
        ttk.Label(kopf, text=titel,
                  font=("Segoe UI", 10, "bold")).pack(side="left")
        self.zahl = StringVar(value="")
        ttk.Label(kopf, textvariable=self.zahl,
                  foreground="#555555").pack(side="left", padx=6)
        ttk.Button(kopf, text="leeren", width=7,
                   command=self._leeren_klick).pack(side="right")
        self.suche = StringVar()
        ttk.Entry(kopf, textvariable=self.suche, width=12).pack(side="right",
                                                                padx=4)
        ttk.Label(kopf, text="suchen:").pack(side="right")
        self.suche.trace_add("write", lambda *_: self._zeichne(suchen=True))

        koerper = ttk.Frame(self)
        koerper.pack(fill="both", expand=True)
        # Etwas engere Zeilen als üblich: Die Themen sollen ohne Rollbalken
        # ganz zu sehen sein.
        ttk.Style(self).configure("Schlagwort.Treeview", rowheight=18)
        self.baum = ttk.Treeview(koerper, show="tree", height=hoehe,
                                 selectmode="none", style="Schlagwort.Treeview")
        vs = ttk.Scrollbar(koerper, orient="vertical", command=self.baum.yview)
        self.baum.configure(yscrollcommand=vs.set)
        self.baum.pack(side="left", fill="both", expand=True)
        vs.pack(side="left", fill="y")
        self.baum.tag_configure("an", foreground="#000000",
                                font=("Segoe UI", 9, "bold"))
        self.baum.tag_configure("mit", foreground="#2e6da4")
        self.baum.tag_configure("aus", foreground="#444444")
        self.baum.tag_configure("unbekannt", foreground="#a00000")
        self.baum.tag_configure("fund", background="#fff2a8")
        self.baum.bind("<Button-1>", self._klick)

    # -- Inhalt -------------------------------------------------------

    def befuelle(self, schlagworte: dict[str, core.Schlagwort], achse: str):
        self.baum.delete(*self.baum.get_children())
        self.name_von.clear()
        self.iids_von.clear()
        self.kinder = defaultdict(list)
        self.achse = achse
        self.schlagworte = {n: s for n, s in schlagworte.items()
                            if s.achse == achse}
        wurzeln = []
        for n, s in self.schlagworte.items():
            eltern = [p for p in s.gehoert_zu if p in self.schlagworte]
            if not eltern:
                wurzeln.append(n)
            for p in eltern:
                self.kinder[p].append(n)

        aufgeklappt: set[str] = set()

        def einfuegen(eltern_iid: str, name: str, pfad: frozenset):
            # Aufgeklappt nur beim ersten Vorkommen. Sonst stünde der
            # Rhein-Neckar-Kreis samt Kraichgau, Bergstraße und Odenwald unter
            # der Metropolregion ein zweites Mal offen da, und der Baum wäre um
            # ein Drittel länger, ohne etwas Neues zu zeigen.
            iid = self.baum.insert(eltern_iid, "end", text=name,
                                   open=name not in aufgeklappt)
            aufgeklappt.add(name)
            self.name_von[iid] = name
            self.iids_von[name].append(iid)
            if name in pfad:     # Ring in der Mappe — nicht endlos absteigen
                return
            for k in self.kinder[name]:
                einfuegen(iid, k, pfad | {name})

        for w in wurzeln:
            einfuegen("", w, frozenset())
        self._zeichne()

    def gewaehlt(self) -> set[str]:
        return set(self.gewaehlt_set)

    def setze(self, namen):
        self.gewaehlt_set = set(namen)
        # Was nicht im Baum steht (Tippfehler aus Excel), bekommt eine
        # eigene rote Zeile — sonst ginge es beim nächsten Übernehmen
        # unbemerkt verloren, und abwählen könnte man es auch nicht.
        for iid in [i for i, n in self.name_von.items()
                    if n not in self.schlagworte]:
            self.baum.delete(iid)
            name = self.name_von.pop(iid)
            self.iids_von.pop(name, None)
        for name in sorted(self.gewaehlt_set - set(self.schlagworte)):
            iid = self.baum.insert("", 0, text=name)
            self.name_von[iid] = name
            self.iids_von[name].append(iid)
        self._zeichne()

    def leeren(self):
        self.setze(set())

    # -- Bedienung ----------------------------------------------------

    def _leeren_klick(self):
        self.leeren()
        if self.bei_aenderung:
            self.bei_aenderung()

    def _klick(self, event):
        iid = self.baum.identify_row(event.y)
        if not iid or "indicator" in self.baum.identify_element(event.x,
                                                                event.y):
            return None
        name = self.name_von.get(iid)
        if name is None:
            return "break"
        if self.fuer_buch and self.achse in core.MIT_TEILGEBIETEN:
            betroffen = core.mit_unterbegriffen({name}, self.schlagworte)
        else:
            betroffen = {name}
        if name in self.gewaehlt_set:
            self.gewaehlt_set -= betroffen
        else:
            self.gewaehlt_set |= betroffen
        self._zeichne()
        if self.bei_aenderung:
            self.bei_aenderung()
        return "break"

    def _zeichne(self, suchen: bool = False):
        mit = (set() if self.fuer_buch else
               core.mit_unterbegriffen(self.gewaehlt_set, self.schlagworte)
               - self.gewaehlt_set)
        such = self.suche.get().strip().lower()
        erster_fund = None
        for iid, name in self.name_von.items():
            if name not in self.schlagworte:
                text, tags = f"{HAKEN_AN if name in self.gewaehlt_set else HAKEN_AUS} " \
                             f"{name}  (unbekanntes Schlagwort)", ["unbekannt"]
            elif name in self.gewaehlt_set:
                text, tags = f"{HAKEN_AN} {name}", ["an"]
            elif name in mit:
                text, tags = f"{HAKEN_AUS} {name}  (inbegriffen)", ["mit"]
            else:
                text, tags = f"{HAKEN_AUS} {name}", ["aus"]
            if such and such in name.lower():
                tags.append("fund")
                erster_fund = erster_fund or iid
            self.baum.item(iid, text=text, tags=tags)
        if suchen and erster_fund:
            self.baum.see(erster_fund)
        self.zahl.set(f"{len(self.gewaehlt_set)} gewählt"
                      if self.gewaehlt_set else "")


class App(Tk):
    def __init__(self):
        super().__init__()
        self.title("Presseverteiler")
        self.geometry("1440x900")
        self.minsize(1150, 720)
        # Zwei Listen und zwei Schlagwortbäume brauchen jede Zeile Höhe.
        if sys.platform == "win32":
            self.state("zoomed")

        self.cfg = self._lade_config()
        self.access_pfad = StringVar(value=self.cfg.get("access", ""))

        self.kontakte: list[core.Kontakt] = []
        self.nach_id: dict[int, core.Kontakt] = {}
        self.verteiler: core.Verteiler | None = None
        # Änderungszeit der Mappe beim Laden. Weicht sie beim Speichern ab,
        # hat jemand die Mappe in Excel bearbeitet — dann nicht drüberschreiben.
        self.mappe_stand: float | None = None
        self.ungespeichert = False

        # Auswahl je Buch
        self.treffer: list[tuple[core.Kontakt, str]] = []
        self.abgewaehlt: set[int] = set()
        self.hinzu: set[int] = set()

        self._baue_ui()
        self.protocol("WM_DELETE_WINDOW", self._schliessen)
        self.report_callback_exception = self._fehler
        if self.access_pfad.get() and Path(self.access_pfad.get()).is_file():
            self.after(100, self._laden)

    # -- Konfiguration ------------------------------------------------

    def _lade_config(self) -> dict:
        try:
            return json.loads(core.CONFIG_PFAD.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return {}

    def _sichere_config(self):
        try:
            core.CONFIG_PFAD.write_text(json.dumps(
                {"access": self.access_pfad.get()}, indent=2,
                ensure_ascii=False), encoding="utf-8")
        except OSError as e:
            self._log(f"config.json nicht schreibbar: {e}")

    # -- Aufbau -------------------------------------------------------

    def _baue_ui(self):
        pad = {"padx": 12, "pady": 5}

        kopf = ttk.Frame(self)
        kopf.pack(fill="x", **pad)
        ttk.Label(kopf, text="Presseverteiler",
                  font=("Segoe UI", 14, "bold")).pack(side="left")
        self.stand = StringVar(value="noch nichts geladen")
        ttk.Label(kopf, textvariable=self.stand).pack(side="left", padx=16)

        zeile = ttk.Frame(self)
        zeile.pack(fill="x", **pad)
        ttk.Label(zeile, text="Access-Export (xlsx, ganze Kundentabelle)",
                  ).pack(side="left")
        ttk.Entry(zeile, textvariable=self.access_pfad).pack(
            side="left", fill="x", expand=True, padx=8)
        ttk.Button(zeile, text="Durchsuchen…",
                   command=self._waehle_access).pack(side="left")
        ttk.Button(zeile, text="Laden", command=self._laden).pack(side="left",
                                                                  padx=(8, 0))
        ttk.Separator(zeile, orient="vertical").pack(side="left", fill="y",
                                                     padx=10)
        ttk.Button(zeile, text="Mappe in Excel öffnen",
                   command=self._mappe_oeffnen).pack(side="left")

        # Warnbanner — normalerweise unsichtbar
        self.warnung = StringVar(value="")
        self.warn_label = ttk.Label(self, textvariable=self.warnung,
                                    foreground="#a00000", wraplength=1380,
                                    justify="left")

        self.reiter = ttk.Notebook(self)
        self.reiter.pack(fill="both", expand=True, **pad)
        self._baue_auswahl()
        self._baue_einordnen()

        self.verlauf = Text(self, height=4, wrap="word")
        self.verlauf.pack(fill="x", **pad)
        self._log("Bereit. Access-Export wählen und auf Laden klicken.")

    def _baue_auswahl(self):
        rahmen = ttk.Frame(self.reiter, padding=6)
        self.reiter.add(rahmen, text="Auswahl je Buch")

        links = ttk.Frame(rahmen, width=380)
        links.pack(side="left", fill="y")
        links.pack_propagate(False)

        buch = ttk.LabelFrame(links, text="Buch", padding=6)
        buch.pack(fill="x")
        self.buch_titel = StringVar()
        self.buch_isbn = StringVar()
        for text, var in (("Titel", self.buch_titel), ("ISBN", self.buch_isbn)):
            z = ttk.Frame(buch)
            z.pack(fill="x", pady=1)
            ttk.Label(z, text=text, width=6).pack(side="left")
            ttk.Entry(z, textvariable=var).pack(side="left", fill="x",
                                                expand=True)
        ttk.Button(buch, text="Neues Buch (alles leeren)",
                   command=self._neues_buch).pack(anchor="e", pady=(4, 0))

        self.buch_regionen = SchlagwortBaum(
            links, "Regionen des Buchs", True, self._auswahl_neu, hoehe=16)
        self.buch_regionen.pack(fill="both", expand=True, pady=(8, 0))
        self.buch_themen = SchlagwortBaum(
            links, "Themen des Buchs", True, self._auswahl_neu, hoehe=11)
        self.buch_themen.pack(fill="both", expand=True, pady=(8, 0))

        rechts = ttk.Frame(rahmen)
        rechts.pack(side="left", fill="both", expand=True, padx=(10, 0))

        self.erreicht = StringVar(value="Links Regionen und Themen des Buchs "
                                        "ankreuzen.")
        ttk.Label(rechts, textvariable=self.erreicht, wraplength=980,
                  justify="left").pack(fill="x", pady=(0, 4))

        # „warum“ gleich neben die Institution: Das ist die Spalte, mit der
        # man eine Liste prüft. Abteilung und Medium dürfen notfalls aus dem
        # Fenster rutschen.
        spalten = [("sel", "", 34), ("inst", "Institution", 200),
                   ("grund", "warum", 230), ("person", "Person", 140),
                   ("ort", "PLZ / Ort", 130), ("mail", "E-Mail", 190),
                   ("abt", "Abteilung", 130), ("medium", "Medium", 80)]
        koerper = ttk.Frame(rechts)
        koerper.pack(fill="both", expand=True)
        baum = ttk.Treeview(koerper, columns=[s[0] for s in spalten],
                            show="headings", selectmode="extended")
        for schluessel, titel, breite in spalten:
            baum.heading(schluessel, text=titel)
            baum.column(schluessel, width=breite, stretch=schluessel != "sel",
                        anchor="center" if schluessel == "sel" else "w")
        vs = ttk.Scrollbar(koerper, orient="vertical", command=baum.yview)
        baum.configure(yscrollcommand=vs.set)
        baum.pack(side="left", fill="both", expand=True)
        vs.pack(side="left", fill="y")
        baum.tag_configure("aus", foreground="#999999")
        baum.tag_configure("ungeprueft", foreground="#8a5a00")
        baum.tag_configure("hand", foreground="#1f5fa8")
        baum.bind("<Button-1>", self._treffer_klick)
        baum.bind("<space>", lambda e: self._treffer_umschalten(
            [int(i) for i in baum.selection()]))
        baum.bind("<Double-1>", lambda e: self._zum_einordnen())
        self.treffer_baum = baum

        self.treffer_zahl = StringVar(value="")
        ttk.Label(rechts, textvariable=self.treffer_zahl).pack(anchor="w",
                                                               pady=(4, 0))
        ttk.Label(rechts, foreground="#555555", text=(
            "Haken anklicken (oder Leertaste) = für dieses Buch abwählen. "
            "Braun = Einordnung noch ungeprüft (Vorschlag aus Access). "
            "Blau = von Hand hinzugefügt. Doppelklick = Zuständigkeit "
            "bearbeiten.")).pack(anchor="w")

        knoepfe = ttk.Frame(rechts)
        knoepfe.pack(fill="x", pady=(6, 0))
        ttk.Button(knoepfe, text="Kontakt von Hand hinzufügen…",
                   command=self._hinzufuegen_dialog).pack(side="left")
        ttk.Button(knoepfe, text="Zuständigkeit bearbeiten",
                   command=self._zum_einordnen).pack(side="left", padx=6)
        ttk.Button(knoepfe, text="Ordner öffnen",
                   command=self._ordner_oeffnen).pack(side="right")
        ttk.Button(knoepfe, text="Liste speichern (Serienbrief + Protokoll)",
                   command=self._liste_speichern).pack(side="right", padx=6)
        ttk.Button(knoepfe, text="E-Mail-Adressen kopieren (BCC)",
                   command=self._mails_kopieren).pack(side="right")

    def _baue_einordnen(self):
        rahmen = ttk.Frame(self.reiter, padding=6)
        self.reiter.add(rahmen, text="Einordnen")

        # Liste und Zuständigkeit teilen sich den Platz an einem
        # verschiebbaren Steg. Die Zuständigkeit ist breit genug für beide
        # Bäume nebeneinander — übereinander waren sie je neun Zeilen hoch,
        # und man musste bei jedem Kontakt scrollen. Wird das Fenster
        # schmaler, gibt die Liste nach (weight=1), nicht die Bäume.
        teiler = ttk.Panedwindow(rahmen, orient="horizontal")
        teiler.pack(fill="both", expand=True)
        links = ttk.Frame(teiler)
        rechts = ttk.LabelFrame(teiler, text="Zuständigkeit", padding=6,
                                width=BREITE_ZUSTAENDIGKEIT)
        rechts.pack_propagate(False)
        teiler.add(links, weight=1)
        teiler.add(rechts, weight=0)

        filter_ = ttk.Frame(links)
        filter_.pack(fill="x", pady=(0, 4))
        ttk.Label(filter_, text="zeigen:").pack(side="left")
        self.ansicht = StringVar(value=ANSICHTEN[0])
        box = ttk.Combobox(filter_, textvariable=self.ansicht, values=ANSICHTEN,
                           state="readonly", width=24)
        box.pack(side="left", padx=(4, 12))
        box.bind("<<ComboboxSelected>>", lambda e: self._fuelle_einordnen())
        ttk.Label(filter_, text="suchen:").pack(side="left")
        self.ein_suche = StringVar()
        ttk.Entry(filter_, textvariable=self.ein_suche, width=24).pack(
            side="left", padx=4)
        self.ein_suche.trace_add("write", lambda *_: self._fuelle_einordnen())
        self.ein_zahl = StringVar(value="")
        ttk.Label(filter_, textvariable=self.ein_zahl).pack(side="left",
                                                            padx=12)

        spalten = [("inst", "Institution", 190), ("abt", "Abteilung", 120),
                   ("ort", "PLZ / Ort", 110), ("art", "Medienart", 90),
                   ("regionen", "Regionen", 160), ("themen", "Themen", 140),
                   ("immer", "immer", 44), ("geprueft", "geprüft", 52)]
        koerper = ttk.Frame(links)
        koerper.pack(fill="both", expand=True)
        baum = ttk.Treeview(koerper, columns=[s[0] for s in spalten],
                            show="headings", selectmode="extended")
        for schluessel, titel, breite in spalten:
            baum.heading(schluessel, text=titel)
            baum.column(schluessel, width=breite, anchor="w")
        vs = ttk.Scrollbar(koerper, orient="vertical", command=baum.yview)
        baum.configure(yscrollcommand=vs.set)
        baum.pack(side="left", fill="both", expand=True)
        vs.pack(side="left", fill="y")
        baum.tag_configure("ungeprueft", foreground="#8a5a00")
        baum.tag_configure("unsicher", foreground="#c0392b")
        baum.tag_configure("leer", foreground="#888888")
        baum.bind("<<TreeviewSelect>>", lambda e: self._zeige_eintrag())
        self.ein_baum = baum

        umbruch = BREITE_ZUSTAENDIGKEIT - 30
        self.ein_wer = StringVar(value="Links einen Kontakt wählen.")
        ttk.Label(rechts, textvariable=self.ein_wer, wraplength=umbruch,
                  font=("Segoe UI", 10, "bold"), justify="left").pack(fill="x")
        # Was in Access über den Kontakt steht. Die alten Felder und vor allem
        # die Bemerkungen tragen Regeln, die nirgends sonst stehen („möchte
        # alle Kraichgau-Titel an ihn nach Bretten“).
        self.ein_access = StringVar(value="")
        ttk.Label(rechts, textvariable=self.ein_access, wraplength=umbruch,
                  foreground="#555555", justify="left").pack(fill="x",
                                                             pady=(2, 0))
        # Woraus der ungeprüfte Vorschlag folgt — damit man ihn mit einem
        # Blick bestätigen kann, statt den Kontakt neu zu recherchieren.
        self.ein_vorschlag = StringVar(value="")
        self.ein_vorschlag_label = ttk.Label(
            rechts, textvariable=self.ein_vorschlag, wraplength=umbruch,
            foreground="#8a5a00", justify="left")
        self.ein_vorschlag_label.pack(fill="x", pady=(2, 6))

        # Knopf und Felder von unten her zuerst packen — sie müssen immer
        # sichtbar bleiben; den Rest teilen sich die beiden Bäume.
        self.knopf_uebernehmen = ttk.Button(
            rechts, text="Übernehmen (gilt als geprüft) und weiter",
            command=self._uebernehmen, state="disabled")
        self.knopf_uebernehmen.pack(side="bottom", fill="x", pady=(6, 0))
        ttk.Label(rechts, foreground="#555555", wraplength=umbruch,
                  justify="left",
                  text="Leere Regionen = egal woher, leere Themen = egal "
                       "worüber. Ganz leer = bekommt nie etwas.").pack(
            side="bottom", fill="x", pady=(6, 0))
        z = ttk.Frame(rechts)
        z.pack(side="bottom", fill="x", pady=(4, 0))
        ttk.Label(z, text="Bemerkung").pack(side="left")
        self.ein_bemerkung = StringVar()
        ttk.Entry(z, textvariable=self.ein_bemerkung).pack(
            side="left", fill="x", expand=True, padx=(6, 0))
        z = ttk.Frame(rechts)
        z.pack(side="bottom", fill="x", pady=(6, 0))
        self.ein_immer = BooleanVar(value=False)
        ttk.Checkbutton(z, text="bekommt immer jedes Buch",
                        variable=self.ein_immer).pack(side="left")
        # Medienart und Ressort gehen noch nicht in die Auswahl ein (siehe
        # core.MEDIENARTEN) — hier nur, damit man Falsches gleich berichtigt.
        self.ein_ressort = StringVar()
        ttk.Combobox(z, textvariable=self.ein_ressort, state="readonly",
                     values=["", *core.RESSORTS], width=18).pack(side="right")
        ttk.Label(z, text="Ressort").pack(side="right", padx=(12, 4))
        self.ein_medienart = StringVar()
        ttk.Combobox(z, textvariable=self.ein_medienart, state="readonly",
                     values=["", *core.MEDIENARTEN], width=15).pack(
            side="right")
        ttk.Label(z, text="Medienart").pack(side="right", padx=(12, 4))

        baeume = ttk.Frame(rechts)
        baeume.pack(fill="both", expand=True)
        baeume.columnconfigure((0, 1), weight=1, uniform="baum")
        baeume.rowconfigure(0, weight=1)
        self.ein_regionen = SchlagwortBaum(baeume, "Regionen", False, hoehe=20)
        self.ein_regionen.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        self.ein_themen = SchlagwortBaum(baeume, "Themen", False, hoehe=20)
        self.ein_themen.grid(row=0, column=1, sticky="nsew")

    # -- Laden --------------------------------------------------------

    def _waehle_access(self):
        pfad = filedialog.askopenfilename(
            title="Access-Export wählen",
            filetypes=[("Excel", "*.xlsx"), ("alle Dateien", "*.*")])
        if pfad:
            self.access_pfad.set(pfad)
            self._laden()

    def _laden(self):
        pfad = Path(self.access_pfad.get())
        if not pfad.is_file():
            messagebox.showerror("Access-Export",
                                 "Bitte zuerst den Access-Export wählen.")
            return
        if self.ungespeichert and not messagebox.askyesno(
                "Neu laden",
                "Die letzte Änderung konnte nicht gespeichert werden und geht "
                "beim Neuladen verloren. Trotzdem laden?"):
            return
        self.config(cursor="watch")
        self._log(f"Lese {pfad.name} …")
        self.update()
        try:
            kontakte, uebergangen = core.lade_presse(pfad)
            warnungen = []
            if core.VERTEILER_PFAD.is_file():
                v = core.lade_verteiler()
                neu, weg = core.abgleich(v, kontakte)
                if neu:
                    self._log(f"{len(neu)} neue Pressekontakte aus Access "
                              f"aufgenommen (ungeprüft).")
                if weg:
                    warnungen.append(
                        f"{len(weg)} Kontakte aus der Mappe gibt es in Access "
                        f"nicht mehr (oder sie sind als verzogen/verstorben "
                        f"markiert). Sie bleiben in der Mappe, werden aber nie "
                        f"ausgewählt.")
                speichern = bool(neu)
            else:
                v = core.neuer_verteiler(kontakte)
                self._log(f"Mappe {core.VERTEILER_PFAD.name} neu angelegt. "
                          f"{sum(e.eingeordnet for e in v.eintraege.values())} "
                          f"Kontakte haben Vorschläge aus den alten "
                          f"Access-Feldern Region/Fachgebiet bekommen — "
                          f"ungeprüft.")
                speichern = True
        except Exception as e:                              # noqa: BLE001
            self.config(cursor="")
            messagebox.showerror("Laden fehlgeschlagen", str(e))
            self._log(f"Laden fehlgeschlagen: {e}")
            return
        finally:
            self.config(cursor="")

        warnungen += core.pruefe_verteiler(v)
        if uebergangen:
            self._log(f"{uebergangen} Pressesätze übergangen (verzogen, "
                      f"verstorben o. Ä.).")

        self.kontakte = kontakte
        self.nach_id = {k.access_id: k for k in kontakte}
        self.verteiler = v
        self.mappe_stand = (core.VERTEILER_PFAD.stat().st_mtime
                            if core.VERTEILER_PFAD.is_file() else None)
        self.ungespeichert = False
        self._sichere_config()
        if speichern:
            self._speichern()

        for baum, achse in ((self.buch_regionen, core.REGION),
                            (self.buch_themen, core.THEMA),
                            (self.ein_regionen, core.REGION),
                            (self.ein_themen, core.THEMA)):
            vorher = baum.gewaehlt()
            baum.befuelle(v.schlagworte, achse)
            baum.setze(vorher)
        self._zeige_warnungen(warnungen)
        self._fuelle_einordnen()
        self._auswahl_neu()
        self._stand_neu()
        self._log(f"{len(kontakte)} Pressekontakte geladen.")

    def _zeige_warnungen(self, warnungen: list[str]):
        if warnungen:
            text = "\n".join("• " + w for w in warnungen[:8])
            if len(warnungen) > 8:
                text += f"\n… und {len(warnungen) - 8} weitere."
            self.warnung.set(text)
            self.warn_label.pack(fill="x", padx=12, before=self.reiter)
        else:
            self.warnung.set("")
            self.warn_label.pack_forget()

    def _stand_neu(self):
        if not self.verteiler:
            return
        aktiv = [self.verteiler.eintraege[k.access_id] for k in self.kontakte
                 if k.access_id in self.verteiler.eintraege]
        eingeordnet = sum(1 for e in aktiv if e.eingeordnet)
        geprueft = sum(1 for e in aktiv if e.geprueft)
        unsicher = sum(1 for e in aktiv if e.unsicher and not e.geprueft)
        teile = [f"{len(aktiv)} Pressekontakte", f"{geprueft} geprüft"]
        if unsicher:
            teile.append(f"{unsicher} ungeprüft und unsicher")
        teile.append(f"{len(aktiv) - eingeordnet} bekommen nie etwas")
        self.stand.set(" · ".join(teile))

    # -- Speichern ----------------------------------------------------

    def _speichern(self) -> bool:
        """Die ganze Mappe schreiben — nach jeder Änderung, damit nichts
        verloren geht, wenn das Fenster einfach zugemacht wird."""
        pfad = core.VERTEILER_PFAD
        if (self.mappe_stand is not None and pfad.is_file()
                and pfad.stat().st_mtime != self.mappe_stand):
            self.ungespeichert = True
            messagebox.showwarning(
                "Mappe von außen geändert",
                f"„{pfad.name}“ wurde geändert, seit das Programm sie geladen "
                f"hat — vermutlich in Excel. Damit diese Änderungen nicht "
                f"überschrieben werden, wird jetzt nicht gespeichert.\n\n"
                f"Bitte auf „Laden“ klicken und die letzte Einordnung "
                f"wiederholen.")
            return False
        try:
            core.speichere_verteiler(self.verteiler, self.kontakte, pfad)
        except (PermissionError, OSError) as e:
            self.ungespeichert = True
            messagebox.showwarning("Nicht gespeichert", str(e))
            self._log(f"Nicht gespeichert: {e}")
            return False
        self.mappe_stand = pfad.stat().st_mtime
        self.ungespeichert = False
        return True

    def _mappe_oeffnen(self):
        if not core.VERTEILER_PFAD.is_file():
            messagebox.showinfo("Noch keine Mappe",
                                "Die Mappe entsteht beim ersten Laden.")
            return
        messagebox.showinfo(
            "In Excel bearbeiten",
            "Solange die Mappe in Excel offen ist, kann das Programm nicht "
            "speichern.\n\nWer in Excel etwas ändert: dort speichern, Excel "
            "schließen und hier auf „Laden“ klicken.")
        _oeffne(core.VERTEILER_PFAD)

    # -- Auswahl je Buch ----------------------------------------------

    def _neues_buch(self):
        self.buch_titel.set("")
        self.buch_isbn.set("")
        self.buch_regionen.leeren()
        self.buch_themen.leeren()
        self.abgewaehlt.clear()
        self.hinzu.clear()
        self._auswahl_neu()

    def _auswahl_neu(self):
        if not self.verteiler:
            return
        regionen = self.buch_regionen.gewaehlt()
        themen = self.buch_themen.gewaehlt()
        self.treffer = core.auswahl(self.verteiler, self.kontakte, regionen,
                                    themen)
        gefunden = {k.access_id for k, _ in self.treffer}
        for i in sorted(self.hinzu - gefunden):
            if i in self.nach_id:
                self.treffer.append((self.nach_id[i], "von Hand"))
        self.treffer.sort(key=lambda t: t[0].sortierung)

        sw = self.verteiler.schlagworte
        mit = (core.erreicht_vom_buch(regionen, sw, core.REGION)
               | core.erreicht_vom_buch(themen, sw, core.THEMA))
        mit -= regionen | themen
        if not (regionen or themen):
            self.erreicht.set("Links Regionen und Themen des Buchs ankreuzen. "
                              "Ohne Schlagworte bekommen es nur die Kontakte "
                              "mit „immer“.")
        elif mit:
            # Nach oben, nicht nach unten: Ein Buch zur Revolution 1848/49
            # geht an die Geschichtsredaktion, nicht an die Kirchengeschichte.
            self.erreicht.set(
                "Außerdem bekommen es alle, die allgemeiner zuständig sind: "
                + ", ".join(sorted(mit)) + ". Wer nur für ein anderes "
                "Unterthema oder Teilgebiet eingetragen ist, bekommt es nicht.")
        else:
            self.erreicht.set("")
        self._fuelle_treffer()

    def _fuelle_treffer(self):
        baum = self.treffer_baum
        baum.delete(*baum.get_children())
        for k, g in self.treffer:
            e = self.verteiler.eintraege.get(k.access_id)
            aus = k.access_id in self.abgewaehlt
            if aus:
                tags = ("aus",)
            elif g == "von Hand":
                tags = ("hand",)
            elif e and not e.geprueft:
                tags = ("ungeprueft",)
            else:
                tags = ()
            if not e or e.geprueft or g == "von Hand":
                vermerk = ""
            else:
                vermerk = "  (ungeprüft, unsicher)" if e.unsicher \
                    else "  (ungeprüft)"
            baum.insert("", "end", iid=str(k.access_id), tags=tags, values=(
                HAKEN_AUS if aus else HAKEN_AN, k.institution, g + vermerk,
                k.person, k.ort, k.mail if k.mails else "— nur Post —", k.abteilung,
                k.medium))
        gewaehlt = self._empfaenger()
        _, ohne, doppelt = core.mailadressen([k for k, _ in gewaehlt])
        teile = [f"{len(gewaehlt)} Empfänger"]
        if len(gewaehlt) != len(self.treffer):
            teile.append(f"{len(self.treffer) - len(gewaehlt)} abgewählt")
        if ohne:
            teile.append(f"{ohne} ohne E-Mail (nur Post)")
        if doppelt:
            teile.append(f"{doppelt} E-Mail-Adressen doppelt "
                         f"(werden einmal kopiert)")
        self.treffer_zahl.set(" · ".join(teile))

    def _empfaenger(self) -> list[tuple[core.Kontakt, str]]:
        return [(k, g) for k, g in self.treffer
                if k.access_id not in self.abgewaehlt]

    def _treffer_klick(self, event):
        if self.treffer_baum.identify_column(event.x) != "#1":
            return None
        iid = self.treffer_baum.identify_row(event.y)
        if iid:
            self._treffer_umschalten([int(iid)])
        return "break"

    def _treffer_umschalten(self, ids: list[int]):
        for i in ids:
            self.abgewaehlt ^= {i}
        auswahl = self.treffer_baum.selection()
        oben = self.treffer_baum.yview()[0]
        self._fuelle_treffer()
        self.treffer_baum.yview_moveto(oben)
        vorhanden = [i for i in auswahl if self.treffer_baum.exists(i)]
        if vorhanden:
            self.treffer_baum.selection_set(vorhanden)
        return "break"

    def _hinzufuegen_dialog(self):
        if not self.verteiler:
            return
        fenster = Toplevel(self)
        fenster.title("Kontakt von Hand hinzufügen")
        fenster.geometry("900x520")
        fenster.transient(self)

        z = ttk.Frame(fenster, padding=8)
        z.pack(fill="x")
        ttk.Label(z, text="suchen (Institution, Person, Ort, E-Mail):").pack(
            side="left")
        such = StringVar()
        eingabe = ttk.Entry(z, textvariable=such, width=40)
        eingabe.pack(side="left", padx=6)
        eingabe.focus_set()

        spalten = [("inst", "Institution", 260), ("person", "Person", 170),
                   ("ort", "PLZ / Ort", 150), ("mail", "E-Mail", 240)]
        koerper = ttk.Frame(fenster, padding=(8, 0))
        koerper.pack(fill="both", expand=True)
        baum = ttk.Treeview(koerper, columns=[s[0] for s in spalten],
                            show="headings", selectmode="extended")
        for schluessel, titel, breite in spalten:
            baum.heading(schluessel, text=titel)
            baum.column(schluessel, width=breite, anchor="w")
        vs = ttk.Scrollbar(koerper, orient="vertical", command=baum.yview)
        baum.configure(yscrollcommand=vs.set)
        baum.pack(side="left", fill="both", expand=True)
        vs.pack(side="left", fill="y")

        def fuellen(*_):
            baum.delete(*baum.get_children())
            text = such.get().strip().lower()
            schon = {k.access_id for k, _ in self.treffer}
            for k in self.kontakte:
                if k.access_id in schon:
                    continue
                if text and text not in " ".join(
                        (k.institution, k.abteilung, k.person, k.ort,
                         k.mail)).lower():
                    continue
                baum.insert("", "end", iid=str(k.access_id), values=(
                    k.institution, k.person, k.ort, k.mail))

        def hinzu(*_):
            ids = [int(i) for i in baum.selection()]
            if not ids:
                return
            self.hinzu.update(ids)
            self.abgewaehlt.difference_update(ids)
            self._auswahl_neu()
            fuellen()
            self._log(f"{len(ids)} Kontakt(e) von Hand hinzugefügt.")

        such.trace_add("write", fuellen)
        baum.bind("<Double-1>", hinzu)
        k = ttk.Frame(fenster, padding=8)
        k.pack(fill="x")
        ttk.Button(k, text="Hinzufügen", command=hinzu).pack(side="left")
        ttk.Button(k, text="Schließen", command=fenster.destroy).pack(
            side="right")
        fuellen()

    def _mails_kopieren(self):
        empfaenger = self._empfaenger()
        if not empfaenger:
            messagebox.showinfo("Nichts zu kopieren", "Die Liste ist leer.")
            return
        text, ohne, doppelt = core.mailadressen([k for k, _ in empfaenger])
        self.clipboard_clear()
        self.clipboard_append(text)
        anzahl = len(text.split("; ")) if text else 0
        meldung = f"{anzahl} E-Mail-Adressen in die Zwischenablage kopiert."
        if ohne:
            meldung += (f" {ohne} Empfänger haben keine E-Mail — die gehen "
                        f"nur per Post (stehen in der gespeicherten Liste).")
        self._log(meldung)
        messagebox.showinfo("Kopiert", meldung + "\n\nIn Outlook in das Feld "
                                                 "„Bcc“ einfügen.")

    def _liste_speichern(self):
        empfaenger = self._empfaenger()
        if not empfaenger:
            messagebox.showinfo("Nichts zu speichern", "Die Liste ist leer.")
            return
        if not self.buch_titel.get().strip() and not messagebox.askyesno(
                "Ohne Titel",
                "Es ist kein Buchtitel eingetragen. Die Datei bekommt dann nur "
                "das Datum als Namen. Trotzdem speichern?"):
            return
        abgewaehlt = [(k, g) for k, g in self.treffer
                      if k.access_id in self.abgewaehlt]
        try:
            ziel = core.speichere_versand(
                self.buch_titel.get().strip(), self.buch_isbn.get().strip(),
                self.buch_regionen.gewaehlt(), self.buch_themen.gewaehlt(),
                empfaenger, abgewaehlt)
        except OSError as e:
            messagebox.showerror("Nicht gespeichert", str(e))
            return
        self._log(f"Gespeichert: {ziel}")
        messagebox.showinfo("Gespeichert",
                            f"{len(empfaenger)} Empfänger gespeichert in\n"
                            f"{ziel}\n\nDas Blatt „Empfänger“ taugt direkt "
                            f"als Datenquelle für den Serienbrief.")

    def _ordner_oeffnen(self):
        if not core.AUSGABE_ORDNER.exists():
            messagebox.showinfo("Noch nichts da",
                                "Der Ordner entsteht beim ersten Speichern.")
            return
        try:
            _oeffne(core.AUSGABE_ORDNER)
        except Exception as e:                              # noqa: BLE001
            self._log(f"Ordner ließ sich nicht öffnen: {e}")
            messagebox.showinfo("Ordner", str(core.AUSGABE_ORDNER))

    def _zum_einordnen(self):
        auswahl = self.treffer_baum.selection()
        if not auswahl:
            messagebox.showinfo("Zuständigkeit bearbeiten",
                                "Bitte zuerst einen Kontakt in der Liste "
                                "markieren.")
            return
        self.ansicht.set(ANSICHTEN[0])
        self.ein_suche.set("")
        self._fuelle_einordnen()
        self.reiter.select(1)
        vorhanden = [i for i in auswahl if self.ein_baum.exists(i)]
        if vorhanden:
            self.ein_baum.selection_set(vorhanden)
            self.ein_baum.see(vorhanden[0])

    # -- Einordnen ----------------------------------------------------

    def _passt_ansicht(self, e: core.Eintrag) -> bool:
        # „ungeprüft“ zeigt auch die leeren Vorschläge: „bekommt nie etwas“
        # ist ebenfalls eine Einschätzung, die jemand bestätigen soll.
        ansicht = self.ansicht.get()
        if ansicht == "ungeprüft":
            return not e.geprueft
        if ansicht == "ungeprüft und unsicher":
            return not e.geprueft and e.unsicher
        if ansicht == "geprüft":
            return e.geprueft
        if ansicht == "bekommt nie etwas":
            return not e.eingeordnet
        if ansicht == "immer":
            return e.immer
        return True

    def _fuelle_einordnen(self):
        if not self.verteiler:
            return
        baum = self.ein_baum
        auswahl = baum.selection()
        baum.delete(*baum.get_children())
        text = self.ein_suche.get().strip().lower()
        v = self.verteiler
        n = 0
        for k in self.kontakte:
            e = v.eintraege.get(k.access_id)
            if e is None or not self._passt_ansicht(e):
                continue
            if text and text not in " ".join(
                    (k.institution, k.abteilung, k.person, k.ort, k.medium,
                     k.mail, e.medienart, e.begruendung, *e.regionen,
                     *e.themen)).lower():
                continue
            if not e.geprueft and e.unsicher:
                tags = ("unsicher",)
            elif not e.eingeordnet:
                tags = ("leer",)
            elif not e.geprueft:
                tags = ("ungeprueft",)
            else:
                tags = ()
            baum.insert("", "end", iid=str(k.access_id), tags=tags, values=(
                k.institution, k.abteilung, k.ort, e.medienart or k.medium,
                core.als_text(e.regionen, v), core.als_text(e.themen, v),
                "x" if e.immer else "", "x" if e.geprueft else ""))
            n += 1
        self.ein_zahl.set(f"{n} Kontakte")
        vorhanden = [i for i in auswahl if baum.exists(i)]
        if vorhanden:
            baum.selection_set(vorhanden)

    def _zeige_eintrag(self):
        ids = [int(i) for i in self.ein_baum.selection()]
        if not ids or not self.verteiler:
            self.ein_wer.set("Links einen Kontakt wählen.")
            self.ein_access.set("")
            self.ein_vorschlag.set("")
            self.knopf_uebernehmen.configure(state="disabled")
            return
        e = self.verteiler.eintraege[ids[0]]
        k = self.nach_id.get(ids[0])
        if len(ids) > 1:
            self.ein_wer.set(f"{len(ids)} Kontakte markiert — „Übernehmen“ "
                             f"gilt für alle. Angezeigt ist der erste.")
        elif k:
            self.ein_wer.set(" · ".join(t for t in (k.institution, k.abteilung,
                                                    k.person, k.ort) if t))
        if k:
            alt = [f"{f} = {k.feld(f)}" for f in (*core.REGION_FELDER,
                                                  *core.FACH_FELDER)
                   if k.feld(f)]
            zeilen = []
            if k.medium:
                zeilen.append(f"Medium: {k.medium}")
            if alt:
                zeilen.append("Access: " + "; ".join(alt))
            if k.feld("Bemerkungen"):
                zeilen.append("Bemerkungen: " + k.feld("Bemerkungen"))
            self.ein_access.set("\n".join(zeilen))
        if e.begruendung:
            stand = ("geprüft" if e.geprueft
                     else "ungeprüft, UNSICHER" if e.unsicher else "ungeprüft")
            self.ein_vorschlag.set(f"Vorschlag ({stand}): {e.begruendung}")
        else:
            self.ein_vorschlag.set("")
        self.ein_vorschlag_label.configure(
            foreground="#c0392b" if e.unsicher and not e.geprueft
            else "#8a5a00")
        self.ein_regionen.setze(e.regionen)
        self.ein_themen.setze(e.themen)
        self.ein_immer.set(e.immer)
        self.ein_bemerkung.set(e.bemerkung)
        self.ein_medienart.set(e.medienart)
        self.ein_ressort.set(e.ressort)
        self.knopf_uebernehmen.configure(
            state="normal",
            text=("Übernehmen (gilt als geprüft) und weiter" if len(ids) == 1
                  else f"Für {len(ids)} Kontakte übernehmen"))

    def _uebernehmen(self):
        ids = [int(i) for i in self.ein_baum.selection()]
        if not ids:
            return
        if len(ids) > 1 and not messagebox.askyesno(
                "Mehrere Kontakte",
                f"Regionen, Themen und „immer“ für alle {len(ids)} markierten "
                f"Kontakte so setzen?"):
            return
        regionen = self.ein_regionen.gewaehlt()
        themen = self.ein_themen.gewaehlt()
        for i in ids:
            e = self.verteiler.eintraege[i]
            e.regionen, e.themen = set(regionen), set(themen)
            e.immer = self.ein_immer.get()
            e.geprueft = True
            e.unsicher = False
            # Bemerkung, Medienart und Ressort nur bei einem Kontakt: Bei
            # mehreren stünden sonst die Werte des ersten bei allen.
            if len(ids) == 1:
                e.bemerkung = self.ein_bemerkung.get().strip()
                e.medienart = self.ein_medienart.get()
                e.ressort = self.ein_ressort.get()
        gespeichert = self._speichern()

        # Weiter zum nächsten Kontakt der Liste — so lässt sich eine Ansicht
        # wie „ungeprüft“ in einem Zug durchklicken.
        alle = list(self.ein_baum.get_children())
        letzter = str(ids[-1])
        naechster = None
        if letzter in alle:
            pos = alle.index(letzter)
            naechster = next((i for i in alle[pos + 1:]
                              if int(i) not in ids), None)
        self._fuelle_einordnen()
        if naechster and self.ein_baum.exists(naechster):
            self.ein_baum.selection_set(naechster)
            self.ein_baum.see(naechster)
        self._auswahl_neu()
        self._stand_neu()
        if gespeichert:
            self._log(f"Übernommen und gespeichert: {len(ids)} Kontakt(e).")

    # -- Verlauf, Fehler, Schluss -------------------------------------

    def _log(self, text):
        self.verlauf.insert(END, text + "\n")
        self.verlauf.see(END)
        self.update_idletasks()

    def _fehler(self, typ, wert, tb):
        """Unerwartete Fehler sichtbar machen — als .exe ohne Konsole ginge
        ein Traceback sonst ins Leere."""
        text = "".join(traceback.format_exception(typ, wert, tb))
        self._log(text)
        messagebox.showerror("Unerwarteter Fehler", str(wert))

    def _schliessen(self):
        if self.ungespeichert and not messagebox.askyesno(
                "Nicht gespeichert",
                "Die letzte Änderung an der Mappe konnte nicht gespeichert "
                "werden. Trotzdem schließen?"):
            return
        self.destroy()


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
