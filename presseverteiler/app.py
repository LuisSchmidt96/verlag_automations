#!/usr/bin/env python3
"""
Presseverteiler (Tkinter-GUI)
=============================

Je Buch ein paar Schlagworte ankreuzen, und ein Vorschlag aus den Listen
ähnlicher früherer Bücher steht da — statt einer halben Stunde in der
Access-Liste.

Zwei Reiter:
- Auswahl je Buch : Regionen und Themen des Buchs ankreuzen, Vorschlag prüfen,
                    mit einem Klick alles ausgeben.
- Kontakte        : Regionen und Themen je Kontakt pflegen. Das ist Suchhilfe
                    und hilft neuen Kontakten, für die Auswahl aber keine Pflicht.

Neben der .exe:
- config.json            : zuletzt benutzter Access-Export, ggf. andere Ablage

In der Ablage auf dem NAS (core.AUSGABE_ORDNER):
- <Titel>/               : je Buch die Listen für E-Mail und Post samt Protokoll
- _Presseverteiler/Pressehistorie.xlsx  : wer welches Buch bekommen hat
- _Presseverteiler/Presseverteiler.xlsx : Regionen und Themen je Kontakt

Bauen als .exe (Windows):
    pyinstaller presseverteiler/Presseverteiler.spec
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import traceback
from collections import Counter, defaultdict
from pathlib import Path
from tkinter import (Tk, Toplevel, StringVar, BooleanVar, Text, END,
                     filedialog, messagebox, ttk)

from presseverteiler import core


HAKEN_AN, HAKEN_AUS = "☑", "☐"

ANSICHTEN = ["alle", "ungeprüft", "ungeprüft und unsicher", "geprüft",
             "nie gewählt", "immer"]

# Breite der rechten Spalte im Reiter „Kontakte“: Platz für die Bäume der
# Regionen und Themen nebeneinander.
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
      core.MIT_TEILGEBIETEN.
    * **Pressekontakt**: Die Unterbegriffe werden nur als „inbegriffen“
      markiert, nicht angekreuzt. Ausgeschrieben in die Mappe stünden sonst bei
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
        # Ohne eigene Wahl der gemeinsame Export auf dem NAS: Ein neuer Rechner
        # muss dann nichts einstellen.
        self.access_pfad = StringVar(value=self.cfg.get("access") or str(
            self._ablage() / core.DATEN_ORDNER / core.ACCESS_PFAD.name))

        self.kontakte: list[core.Kontakt] = []
        self.nach_id: dict[int, core.Kontakt] = {}
        self.verteiler: core.Verteiler | None = None
        # Änderungszeit der Mappe beim Laden. Weicht sie beim Speichern ab,
        # hat jemand die Mappe in Excel bearbeitet — dann nicht drüberschreiben.
        self.mappe_stand: float | None = None
        self.ungespeichert = False

        self.historie: list[core.Buch] = []
        self.gewaehlt_zahl: Counter = Counter()   # je Kontakt: in wie vielen Listen

        # Auswahl je Buch. Die Häkchen ergeben sich aus dem Vorschlag und den
        # Klicks darauf. Die Klicks stehen getrennt, damit sie erhalten bleiben,
        # wenn man danach noch eine Region dazunimmt.
        self.treffer: list[core.Treffer] = []
        # (Spalte, absteigend); None = Reihenfolge des Vorschlags
        self.treffer_ordnung: tuple[str, bool] | None = None
        self.nachbarn: list[tuple[float, core.Buch]] = []
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
            self.cfg["access"] = self.access_pfad.get()
            core.CONFIG_PFAD.write_text(json.dumps(
                self.cfg, indent=2, ensure_ascii=False), encoding="utf-8")
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

        # Die früheren Bücher zeigen, woher der Vorschlag kommt. Zwei Bücher
        # derselben Reihe oder desselben Orts bekamen oft dieselbe Liste, darum
        # lässt sich eine davon direkt übernehmen.
        aehnlich = ttk.LabelFrame(rechts, text="Ähnliche frühere Bücher",
                                  padding=4)
        aehnlich.pack(fill="x")
        spalten = [("titel", "Titel", 460), ("datum", "Datum", 80),
                   ("zahl", "Empfänger", 70), ("regionen", "Regionen", 200),
                   ("themen", "Themen", 170)]
        koerper = ttk.Frame(aehnlich)
        koerper.pack(fill="x")
        baum = ttk.Treeview(koerper, columns=[s_[0] for s_ in spalten],
                            show="headings", selectmode="browse", height=5)
        for schluessel, titel, breite in spalten:
            baum.heading(schluessel, text=titel)
            baum.column(schluessel, width=breite, anchor="w")
        vs = ttk.Scrollbar(koerper, orient="vertical", command=baum.yview)
        baum.configure(yscrollcommand=vs.set)
        baum.pack(side="left", fill="x", expand=True)
        vs.pack(side="left", fill="y")
        baum.bind("<Double-1>", lambda e: self._liste_uebernehmen())
        self.aehnlich_baum = baum
        ttk.Button(aehnlich, text="Liste dieses Buchs übernehmen",
                   command=self._liste_uebernehmen).pack(anchor="e",
                                                         pady=(4, 0))

        self.erreicht = StringVar(value="Links Regionen und Themen des Buchs "
                                        "ankreuzen.")
        ttk.Label(rechts, textvariable=self.erreicht, wraplength=980,
                  justify="left").pack(fill="x", pady=(6, 4))

        suche = ttk.Frame(rechts)
        suche.pack(fill="x", pady=(0, 4))
        ttk.Label(suche, text="suchen:").pack(side="left")
        self.treffer_suche = StringVar()
        self.treffer_suche.trace_add("write",
                                     lambda *_: self._fuelle_treffer())
        ttk.Entry(suche, textvariable=self.treffer_suche, width=28).pack(
            side="left", padx=(4, 12))
        ttk.Label(suche, foreground="#555555", text=(
            "Klick auf eine Spaltenüberschrift sortiert, der zweite dreht um, "
            "der dritte stellt den Vorschlag wieder her.")).pack(side="left")

        # „Quote“ und „warum“ gleich neben die Institution: Das sind die
        # Spalten, mit denen man eine Liste prüft. Person und Medienart dürfen
        # notfalls aus dem Fenster rutschen.
        spalten = [("sel", "", 34), ("inst", "Institution", 200),
                   ("abt", "Abteilung", 150), ("quote", "Quote", 56),
                   ("grund", "warum", 170), ("ort", "PLZ / Ort", 130),
                   ("mail", "E-Mail", 180), ("person", "Person", 130),
                   ("medium", "Medienart", 90)]
        self.treffer_spalten = {s_[0]: s_[1] for s_ in spalten}
        koerper = ttk.Frame(rechts)
        koerper.pack(fill="both", expand=True)
        baum = ttk.Treeview(koerper, columns=[s_[0] for s_ in spalten],
                            show="headings", selectmode="extended")
        for schluessel, titel, breite in spalten:
            schmal = schluessel in ("sel", "quote")
            baum.heading(schluessel, text=titel,
                         command=lambda s_=schluessel:
                         self._treffer_sortieren(s_))
            baum.column(schluessel, width=breite, stretch=not schmal,
                        anchor="center" if schmal else "w")
        vs = ttk.Scrollbar(koerper, orient="vertical", command=baum.yview)
        baum.configure(yscrollcommand=vs.set)
        baum.pack(side="left", fill="both", expand=True)
        vs.pack(side="left", fill="y")
        baum.tag_configure("aus", foreground="#999999")
        baum.tag_configure("hand", foreground="#1f5fa8")
        baum.bind("<Button-1>", self._treffer_klick)
        baum.bind("<space>", lambda e: self._treffer_umschalten(
            [int(i) for i in baum.selection()]))
        baum.bind("<Double-1>", self._treffer_doppelklick)
        self.treffer_baum = baum

        self.treffer_zahl = StringVar(value="")
        ttk.Label(rechts, textvariable=self.treffer_zahl).pack(anchor="w",
                                                               pady=(4, 0))
        ttk.Label(rechts, foreground="#555555", text=(
            "Haken anklicken (oder Leertaste) = an- oder abwählen. Grau = "
            "nicht angekreuzt, aber naheliegend: Hier lohnt der Blick nach "
            "unten. Blau = von Hand dazugenommen. Doppelklick = Kontakt "
            "ansehen.")).pack(anchor="w")

        knoepfe = ttk.Frame(rechts)
        knoepfe.pack(fill="x", pady=(6, 0))
        ttk.Button(knoepfe, text="Kontakte suchen und hinzufügen…",
                   command=self._hinzufuegen_dialog).pack(side="left")
        ttk.Button(knoepfe, text="Kontakt ansehen",
                   command=self._zum_einordnen).pack(side="left", padx=6)
        # Ein Knopf für alles: Einzelne Schritte vergisst man, und ohne den
        # Eintrag in die Historie lernt das Tool die Liste nie.
        ttk.Button(knoepfe, text="Fertig – alles ausgeben",
                   command=self._alles_ausgeben).pack(side="right")
        ttk.Label(knoepfe, foreground="#555555", text=(
            "Ordner mit allen Listen, E-Mail-Adressen in die Zwischenablage"
            )).pack(side="right", padx=8)

    def _baue_einordnen(self):
        rahmen = ttk.Frame(self.reiter, padding=6)
        self.reiter.add(rahmen, text="Kontakte")

        # Liste und rechte Spalte teilen sich den Platz an einem
        # verschiebbaren Steg. Die rechte Spalte ist breit genug für beide
        # Bäume nebeneinander — übereinander waren sie je neun Zeilen hoch,
        # und man musste bei jedem Kontakt scrollen. Wird das Fenster
        # schmaler, gibt die Liste nach (weight=1), nicht die Bäume.
        teiler = ttk.Panedwindow(rahmen, orient="horizontal")
        teiler.pack(fill="both", expand=True)
        links = ttk.Frame(teiler)
        rechts = ttk.LabelFrame(teiler, text="Regionen und Themen", padding=6,
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
                   ("listen", "Listen", 46), ("immer", "immer", 44),
                   ("geprueft", "geprüft", 52)]
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
                  text="Wer ein Buch bekommt, ergibt sich aus früheren Listen. "
                       "Regionen und Themen helfen beim Suchen und bei neuen "
                       "Kontakten ohne Vorgeschichte.").pack(
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
        ttk.Checkbutton(z, text="bei jedem Buch vorankreuzen",
                        variable=self.ein_immer).pack(side="left")
        # Medienart und Ressort gehen nicht in den Vorschlag ein (siehe
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
        if self._ablage_pruefen() is None:
            return
        mappe, historie_pfad = self._mappe_pfad(), self._historie_pfad()
        self.config(cursor="watch")
        self._log(f"Lese {pfad.name} …")
        self.update()
        try:
            kontakte, uebergangen = core.lade_presse(pfad)
            warnungen = []
            if mappe.is_file():
                v = core.lade_verteiler(mappe)
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
                self._log(f"Mappe {mappe.name} neu angelegt. "
                          f"{sum(e.eingeordnet for e in v.eintraege.values())} "
                          f"Kontakte haben Vorschläge aus den alten "
                          f"Access-Feldern Region/Fachgebiet bekommen — "
                          f"ungeprüft.")
                speichern = True
            historie = core.lade_historie(historie_pfad)
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

        if not historie:
            warnungen.append(
                f"Keine {historie_pfad.name} in {historie_pfad.parent} "
                f"gefunden. Ohne frühere Listen gibt es nur Vorschläge nach "
                f"Gegend.")
        self.kontakte = kontakte
        self.nach_id = {k.access_id: k for k in kontakte}
        self.verteiler = v
        self._historie_setzen(historie)
        self.mappe_stand = (mappe.stat().st_mtime
                            if mappe.is_file() else None)
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
        self._log(f"{len(kontakte)} Pressekontakte und {len(historie)} frühere "
                  f"Bücher geladen.")

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
        nie = sum(1 for k in self.kontakte
                  if not self.gewaehlt_zahl[k.access_id])
        self.stand.set(f"{len(aktiv)} Pressekontakte · {len(self.historie)} "
                       f"frühere Bücher · {nie} Kontakte noch nie gewählt")

    # -- Speichern ----------------------------------------------------

    def _speichern(self) -> bool:
        """Die ganze Mappe schreiben — nach jeder Änderung, damit nichts
        verloren geht, wenn das Fenster einfach zugemacht wird."""
        pfad = self._mappe_pfad()
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
        if not self._mappe_pfad().is_file():
            messagebox.showinfo("Noch keine Mappe",
                                "Die Mappe entsteht beim ersten Laden.")
            return
        messagebox.showinfo(
            "In Excel bearbeiten",
            "Solange die Mappe in Excel offen ist, kann das Programm nicht "
            "speichern.\n\nWer in Excel etwas ändert: dort speichern, Excel "
            "schließen und hier auf „Laden“ klicken.")
        _oeffne(self._mappe_pfad())

    # -- Auswahl je Buch ----------------------------------------------

    def _historie_setzen(self, historie: list[core.Buch]):
        self.historie = historie
        self.gewaehlt_zahl = Counter(i for b in historie for i in b.empfaenger)

    def _neues_buch(self):
        self.buch_titel.set("")
        self.buch_isbn.set("")
        self.treffer_suche.set("")
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
        vs = core.vorschlag(regionen, themen, self.historie, self.verteiler,
                            self.kontakte)
        self.nachbarn = vs.nachbarn
        self.treffer = vs.treffer
        schon = {t.kontakt.access_id for t in self.treffer}
        von_hand = [core.Treffer(self.nach_id[i], 0.0, 0, False, "von Hand")
                    for i in sorted(self.hinzu - schon) if i in self.nach_id]
        von_hand.sort(key=lambda t: t.kontakt.sortierung)
        self.treffer = von_hand + self.treffer

        if not (regionen or themen):
            self.erreicht.set("Links Regionen und Themen des Buchs ankreuzen.")
        elif not self.nachbarn:
            self.erreicht.set("Kein ähnliches früheres Buch gefunden. Unten "
                              "stehen nur Kontakte aus der Gegend, nichts ist "
                              "vorangekreuzt.")
        else:
            prozent = round(core.SCHWELLE * 100)
            self.erreicht.set(
                f"Vorschlag aus den {len(self.nachbarn)} ähnlichsten Büchern: "
                f"Angekreuzt ist, wer bei mindestens {prozent} % davon dabei "
                f"war. Das trifft erfahrungsgemäß rund die Hälfte der "
                f"richtigen Liste, die andere Hälfte steht oft weiter unten "
                f"oder muss gesucht werden.")
        self._fuelle_aehnlich()
        self._fuelle_treffer()

    def _fuelle_aehnlich(self):
        baum = self.aehnlich_baum
        baum.delete(*baum.get_children())
        for _, b in self.nachbarn:
            baum.insert("", "end", iid=str(b.nummer), values=(
                b.titel, b.datum.strftime("%m/%Y"), len(b.empfaenger),
                "; ".join(sorted(b.regionen)), "; ".join(sorted(b.themen))))

    def _angekreuzt(self, t: core.Treffer) -> bool:
        i = t.kontakt.access_id
        return i in self.hinzu or (t.vorgeschlagen and i not in self.abgewaehlt)

    def _such_text(self, k: core.Kontakt) -> str:
        e = self.verteiler.eintraege.get(k.access_id)
        return core.alphabetisch(" ".join(
            (k.institution, k.abteilung, k.person, k.ort, k.mail, k.medium,
             e.medienart if e else "", *(e.regionen if e else ()),
             *(e.themen if e else ()))))

    def _treffer_werte(self, t: core.Treffer) -> dict[str, str]:
        k = t.kontakt
        e = self.verteiler.eintraege.get(k.access_id)
        return {"inst": k.institution, "abt": k.abteilung,
                "quote": f"{round(100 * t.stimmen)} %" if t.stimmen else "",
                "grund": t.warum, "ort": k.ort,
                "mail": k.mail if k.mails else "", "person": k.person,
                "medium": (e.medienart if e else "") or k.medium}

    def _treffer_anzeige(self) -> list[core.Treffer]:
        """Was die Liste zeigt: nach der Suche gefiltert, dann sortiert.

        Beides ändert nur die Anzeige. Wer gerade ausgeblendet ist, bleibt
        angekreuzt und geht mit, sonst könnte eine vergessene Suche still
        Empfänger kosten.
        """
        worte = core.alphabetisch(self.treffer_suche.get()).split()

        def passt(t: core.Treffer) -> bool:
            text = self._such_text(t.kontakt)
            return all(w in text for w in worte)

        zeilen = [t for t in self.treffer if passt(t)] if worte else self.treffer
        if not self.treffer_ordnung:
            return zeilen
        spalte, absteigend = self.treffer_ordnung
        # Gleichstände alphabetisch: Die Sortierung danach ist stabil, auch
        # absteigend.
        zeilen = sorted(zeilen, key=lambda t: t.kontakt.sortierung)
        if spalte == "sel":
            return sorted(zeilen, key=self._angekreuzt, reverse=absteigend)
        if spalte in ("quote", "grund"):
            # Nach Stimmen, nicht nach dem Text: Neuere Bücher wiegen mehr,
            # darum zählt „in 5 von 15“ nicht immer weniger als „in 6 von 15“.
            return sorted(zeilen, key=lambda t: t.stimmen, reverse=absteigend)

        def text(t):
            return core.alphabetisch(self._treffer_werte(t)[spalte])

        # Leere Felder immer nach hinten, auch bei Z–A: Sie sagen nichts aus.
        return (sorted((t for t in zeilen if text(t)), key=text,
                       reverse=absteigend)
                + [t for t in zeilen if not text(t)])

    def _treffer_sortieren(self, spalte: str):
        """Erster Klick sortiert, der zweite dreht um, der dritte stellt die
        Reihenfolge des Vorschlags wieder her. Haken und Quote beginnen mit
        den Stärksten, alles andere bei A."""
        start = spalte in ("sel", "quote", "grund")
        ordnung = self.treffer_ordnung
        if not ordnung or ordnung[0] != spalte:
            self.treffer_ordnung = (spalte, start)
        elif ordnung[1] == start:
            self.treffer_ordnung = (spalte, not start)
        else:
            self.treffer_ordnung = None
        for s_, titel in self.treffer_spalten.items():
            pfeil = ""
            if self.treffer_ordnung and self.treffer_ordnung[0] == s_:
                pfeil = " ▼" if self.treffer_ordnung[1] else " ▲"
            self.treffer_baum.heading(s_, text=titel + pfeil)
        self._fuelle_treffer()

    def _fuelle_treffer(self):
        baum = self.treffer_baum
        baum.delete(*baum.get_children())
        anzeige = self._treffer_anzeige()
        for t in anzeige:
            k = t.kontakt
            w = self._treffer_werte(t)
            an = self._angekreuzt(t)
            tags = (("aus",) if not an
                    else ("hand",) if k.access_id in self.hinzu else ())
            baum.insert("", "end", iid=str(k.access_id), tags=tags, values=(
                HAKEN_AN if an else HAKEN_AUS, w["inst"], w["abt"], w["quote"],
                w["grund"], w["ort"], w["mail"] or "— nur Post —", w["person"],
                w["medium"]))
        gewaehlt = self._empfaenger()
        _, ohne, doppelt = core.mailadressen([k for k, _ in gewaehlt])
        teile = [f"{len(gewaehlt)} Empfänger angekreuzt"]
        rest = len(self.treffer) - len(gewaehlt)
        if rest:
            teile.append(f"{rest} weitere zur Auswahl")
        if ohne:
            teile.append(f"{ohne} ohne E-Mail (nur Post)")
        if doppelt:
            teile.append(f"{doppelt} E-Mail-Adressen doppelt "
                         f"(werden einmal kopiert)")
        if len(anzeige) < len(self.treffer):
            teile.insert(0, f"Suche zeigt {len(anzeige)} von "
                            f"{len(self.treffer)}")
        self.treffer_zahl.set(" · ".join(teile))

    def _empfaenger(self) -> list[tuple[core.Kontakt, str]]:
        return [(t.kontakt, t.warum) for t in self.treffer
                if self._angekreuzt(t)]

    def _treffer_klick(self, event):
        # Überschriften müssen durch, auch die der Hakenspalte: Sie sortieren.
        if (self.treffer_baum.identify_region(event.x, event.y) != "cell"
                or self.treffer_baum.identify_column(event.x) != "#1"):
            return None
        iid = self.treffer_baum.identify_row(event.y)
        if iid:
            self._treffer_umschalten([int(iid)])
        return "break"

    def _treffer_doppelklick(self, event):
        # Zweimal schnell auf eine Überschrift ist Umsortieren, kein Öffnen.
        if self.treffer_baum.identify_region(event.x, event.y) == "cell":
            self._zum_einordnen()

    def _treffer_umschalten(self, ids: list[int]):
        nach_id = {t.kontakt.access_id: t for t in self.treffer}
        for i in ids:
            if i not in nach_id:
                continue
            if self._angekreuzt(nach_id[i]):
                self.hinzu.discard(i)
                if nach_id[i].vorgeschlagen:
                    self.abgewaehlt.add(i)
            else:
                self.abgewaehlt.discard(i)
                if not nach_id[i].vorgeschlagen:
                    self.hinzu.add(i)
        auswahl = self.treffer_baum.selection()
        oben = self.treffer_baum.yview()[0]
        self._fuelle_treffer()
        self.treffer_baum.yview_moveto(oben)
        vorhanden = [i for i in auswahl if self.treffer_baum.exists(i)]
        if vorhanden:
            self.treffer_baum.selection_set(vorhanden)
        return "break"

    def _liste_uebernehmen(self):
        """Die Liste eines früheren Buchs als Ausgangspunkt nehmen."""
        auswahl = self.aehnlich_baum.selection()
        if not auswahl:
            messagebox.showinfo("Liste übernehmen",
                                "Bitte zuerst oben ein Buch markieren.")
            return
        buch = next(b for _, b in self.nachbarn if str(b.nummer) == auswahl[0])
        ids = {i for i in buch.empfaenger if i in self.nach_id}
        self.abgewaehlt = {t.kontakt.access_id for t in self.treffer
                           if t.vorgeschlagen} - ids
        self.hinzu = ids - {t.kontakt.access_id for t in self.treffer
                            if t.vorgeschlagen}
        self._auswahl_neu()
        weg = len(buch.empfaenger) - len(ids)
        self._log(f"Liste von „{buch.titel}“ übernommen: {len(ids)} Kontakte"
                  + (f", {weg} gibt es in Access nicht mehr." if weg else "."))

    def _hinzufuegen_dialog(self):
        if not self.verteiler:
            return
        fenster = Toplevel(self)
        fenster.title("Kontakte suchen und hinzufügen")
        fenster.geometry("900x520")
        fenster.transient(self)

        z = ttk.Frame(fenster, padding=8)
        z.pack(fill="x")
        ttk.Label(z, text="suchen (Medium, Person, Ort, Region, Thema):").pack(
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
            text = core.alphabetisch(such.get().strip())
            schon = {k.access_id for k, _ in self._empfaenger()}
            for k in self.kontakte:
                if k.access_id in schon:
                    continue
                if text and text not in self._such_text(k):
                    continue
                baum.insert("", "end", iid=str(k.access_id), values=(
                    k.institution, k.person, k.ort, k.mail))

        def hinzu(*_, ids=None):
            ids = ids or [int(i) for i in baum.selection()]
            if not ids:
                return
            self.hinzu.update(ids)
            self.abgewaehlt.difference_update(ids)
            self._auswahl_neu()
            fuellen()
            self._log(f"{len(ids)} Kontakt(e) von Hand hinzugefügt.")

        def alle_hinzu():
            # Südkurier, Schwarzwälder Bote und Schwäbische Zeitung gehen breit
            # raus, meist an 40–46 % der Familie: „Südkurier“ suchen, alle
            # nehmen, die übrigen wieder abwählen.
            hinzu(ids=[int(i) for i in baum.get_children()])

        such.trace_add("write", fuellen)
        baum.bind("<Double-1>", hinzu)
        k = ttk.Frame(fenster, padding=8)
        k.pack(fill="x")
        ttk.Button(k, text="Markierte hinzufügen", command=hinzu).pack(
            side="left")
        ttk.Button(k, text="Alle Treffer hinzufügen",
                   command=alle_hinzu).pack(side="left", padx=6)
        ttk.Button(k, text="Schließen", command=fenster.destroy).pack(
            side="right")
        fuellen()

    def _ablage(self) -> Path:
        gewaehlt = self.cfg.get("ablage")
        return Path(gewaehlt) if gewaehlt else core.AUSGABE_ORDNER

    def _mappe_pfad(self) -> Path:
        return self._ablage() / core.DATEN_ORDNER / core.VERTEILER_PFAD.name

    def _historie_pfad(self) -> Path:
        return self._ablage() / core.DATEN_ORDNER / core.HISTORIE_PFAD.name

    def _ablage_pruefen(self) -> Path | None:
        """Die Ablage, oder None, wenn jetzt nicht gespeichert werden soll.

        Nie still woandershin ausweichen: Eine Liste, die nur auf diesem
        Rechner liegt, sucht im Verlag später niemand.
        """
        ordner = self._ablage()
        if ordner.is_dir():
            return ordner
        if not messagebox.askyesno(
                "Ablage nicht erreichbar",
                f"Der Ordner auf dem NAS ist nicht erreichbar:\n{ordner}\n\n"
                f"Dort liegen die Listen und die Pressehistorie. Besteht die "
                f"Verbindung zum Netz? Dann einfach noch einmal versuchen, "
                f"die Auswahl bleibt erhalten.\n\nStattdessen dauerhaft "
                f"einen anderen Ordner wählen (etwa denselben über einen "
                f"Laufwerksbuchstaben)?"):
            return None
        neu = filedialog.askdirectory(title="Ordner „Presselisten für "
                                            "Rezensionsversand“ wählen")
        if not neu:
            return None
        self.cfg["ablage"] = neu
        self._sichere_config()
        # Liegt dort noch keine Mappe, kommt die geladene mit: Sonst legte das
        # nächste Laden eine leere an, und die Einordnungen wären weg.
        if self.verteiler and not self._mappe_pfad().is_file():
            self._speichern()
        return Path(neu)

    def _alles_ausgeben(self):
        """Ein Klick für alles: Ordner mit allen Listen anlegen, die
        E-Mail-Adressen in die Zwischenablage legen, die Liste in die Historie
        aufnehmen und den Ordner öffnen."""
        empfaenger = self._empfaenger()
        if not empfaenger:
            messagebox.showinfo("Nichts auszugeben", "Die Liste ist leer.")
            return
        if not self.buch_titel.get().strip() and not messagebox.askyesno(
                "Ohne Titel",
                "Es ist kein Buchtitel eingetragen. Der Ordner bekommt dann "
                "nur das Datum als Namen. Trotzdem ausgeben?"):
            return
        ablage = self._ablage_pruefen()
        if ablage is None:
            return
        titel = self.buch_titel.get().strip()
        isbn = self.buch_isbn.get().strip()
        regionen = self.buch_regionen.gewaehlt()
        themen = self.buch_themen.gewaehlt()
        abgewaehlt = [(t.kontakt, t.warum) for t in self.treffer
                      if t.kontakt.access_id in self.abgewaehlt]
        try:
            ziel = core.speichere_versand(titel, isbn, regionen, themen,
                                          empfaenger, abgewaehlt, ablage)
        except OSError as e:
            messagebox.showerror("Nicht ausgegeben", str(e))
            return
        self._log(f"Ausgegeben: {ziel}")

        # Dieselbe Reihenfolge wie in den Dateien, sonst stünde in der
        # Zwischenablage womöglich eine andere Schreibweise einer doppelten
        # Sammeladresse als in E-Mail-Adressen.txt.
        kontakte = sorted((k for k, _ in empfaenger), key=lambda k: k.sortierung)
        text, _, _ = core.mailadressen(kontakte)
        self.clipboard_clear()
        self.clipboard_append(text)

        # Frisch einlesen statt die geladene Fassung schreiben: Jemand könnte
        # die Historie inzwischen in Excel bearbeitet haben, etwa eine Liste
        # gelöscht, die als Vorbild nichts taugte. Fehlt sie am Ort (gerade
        # erst einen anderen Ordner gewählt), zieht die geladene mit um.
        vermerk = "Die Liste zählt ab morgen für neue Vorschläge."
        try:
            pfad = self._historie_pfad()
            historie = (core.lade_historie(pfad) if pfad.is_file()
                        else list(self.historie))
            core.buch_merken(historie, titel or "(ohne Titel)", isbn, regionen,
                             themen, [k.access_id for k in kontakte])
            core.speichere_historie(historie, self.kontakte, pfad)
            self._historie_setzen(historie)
            self._stand_neu()
        except Exception as e:                              # noqa: BLE001
            vermerk = (f"Die Liste ist ausgegeben, aber nicht in die Historie "
                       f"gekommen und zählt nicht für neue Vorschläge:\n{e}")
            self._log(f"Historie nicht gespeichert: {e}")

        adressen = len(core.einzelne_mails(kontakte))
        anschriften = sum(1 for k in kontakte if k.hat_anschrift)
        keiner = sum(1 for k in kontakte if not k.mails and not k.hat_anschrift)
        mails = (f"Die {adressen} E-Mail-Adressen liegen schon in der "
                 f"Zwischenablage: in Outlook in das Feld „Bcc“ einfügen."
                 if adressen else "Niemand auf der Liste hat eine E-Mail-Adresse.")
        warnung = (f"\n\n{keiner} Empfänger haben weder E-Mail noch "
                   f"Anschrift und stehen nur in {core.DATEI_EMPFAENGER}."
                   if keiner else "")
        messagebox.showinfo(
            "Fertig",
            f"{len(empfaenger)} Empfänger, ausgegeben nach\n{ziel}\n\n"
            f"{mails}\n\n"
            f"Im Ordner, der sich gleich öffnet:\n"
            f"• {core.DATEI_EMPFAENGER}: alle, mit Protokoll\n"
            f"• {core.DATEI_MAIL}: {adressen} Adressen\n"
            f"• {core.DATEI_POST}: {anschriften} Anschriften für Serienbrief "
            f"und Etiketten\n"
            f"• {core.DATEI_BCC}: die Adressen noch einmal für Bcc"
            f"{warnung}\n\n{vermerk}")
        # Erst nach der Meldung öffnen, sonst schiebt sich der Explorer davor.
        try:
            _oeffne(ziel)
        except Exception as e:                              # noqa: BLE001
            self._log(f"Ordner ließ sich nicht öffnen: {e}")

    def _zum_einordnen(self):
        auswahl = self.treffer_baum.selection()
        if not auswahl:
            messagebox.showinfo("Kontakt ansehen",
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

    # -- Kontakte -----------------------------------------------------

    def _passt_ansicht(self, e: core.Eintrag) -> bool:
        ansicht = self.ansicht.get()
        if ansicht == "ungeprüft":
            return not e.geprueft
        if ansicht == "ungeprüft und unsicher":
            return not e.geprueft and e.unsicher
        if ansicht == "geprüft":
            return e.geprueft
        if ansicht == "nie gewählt":
            return not self.gewaehlt_zahl[e.access_id]
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
                self.gewaehlt_zahl[k.access_id] or "",
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
