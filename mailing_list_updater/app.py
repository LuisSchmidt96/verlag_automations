#!/usr/bin/env python3
"""
Mailinglisten-Abgleich (Tkinter-GUI)
=====================================

Gleicht den Access-Adressstamm gegen die Lexware-Exporte ab und erzeugt zwei
Excel-Dateien für den Rückweg nach Access: eine zum Anfügen, eine zum
Aktualisieren.

Dateien im App-Ordner (neben der .exe):
- config.json               : zuletzt benutzte Pfade
- mailing_output/<Jahr>/    : Ausgabe und entscheidungen.json

Bauen als .exe (Windows):
    pyinstaller mailing_list_updater/MailingListUpdater.spec
"""

from __future__ import annotations

import json
import traceback
from datetime import date
from pathlib import Path
from tkinter import (Tk, Canvas, StringVar, Text, END, TclError, filedialog,
                     messagebox, ttk, Listbox, SINGLE)

from mailing_list_updater import core


# Spalten der vier Listen. Die erste ist immer das Häkchen.
SPALTEN = {
    core.FALL_NEU: [("sel", "", 34), ("kdnr", "Kd.-Nr.", 70),
                    ("name", "Name / Firma", 240), ("ort", "PLZ / Ort", 150),
                    ("strasse", "Straße", 180), ("jahr", "Bestelljahr", 80),
                    ("hinweis", "Auffälligkeit", 260)],
    core.FALL_AKTUALISIEREN: [("sel", "", 34), ("kdnr", "Kd.-Nr.", 70),
                              ("name", "Name / Firma", 220),
                              ("id", "Access-ID", 70), ("punkte", "Punkte", 60),
                              ("jahr", "Bestelljahr", 80),
                              ("abw", "wird in Access geändert", 260)],
    core.FALL_UNKLAR: [("sel", "", 34), ("kdnr", "Kd.-Nr.", 70),
                       ("name", "Name / Firma", 200), ("kand", "Kand.", 46),
                       ("punkte", "bester", 52), ("abstand", "Abstand", 58),
                       ("jahr", "Bestelljahr", 76),
                       ("folge", "Folge", 190),
                       ("hinweis", "warum unklar", 260)],
    core.FALL_OHNE_AUFTRAG: [("sel", "", 34), ("kdnr", "Kd.-Nr.", 70),
                             ("name", "Name / Firma", 240),
                             ("ort", "PLZ / Ort", 150),
                             ("gruppe", "Kundengruppe", 160),
                             ("hinweis", "Hinweis", 260)],
}

# Spalte, welche die anfängliche Sortierung am ehesten beschreibt. Sie
# bekommt die Markierung, damit erkennbar ist, wonach die Liste geordnet ist
# und wohin man klicken muss, um es rückgängig zu machen.
VORGABE_SPALTE = {
    core.FALL_NEU: "hinweis",
    core.FALL_AKTUALISIEREN: "abw",
    core.FALL_UNKLAR: "hinweis",
    core.FALL_OHNE_AUFTRAG: "hinweis",
}

REITER_TITEL = {
    core.FALL_NEU: "Neu anlegen",
    core.FALL_AKTUALISIEREN: "Aktualisieren",
    core.FALL_UNKLAR: "Unklar",
    core.FALL_OHNE_AUFTRAG: "Ohne Auftrag",
}

HAKEN_AN, HAKEN_AUS = "☑", "☐"


def _sortwert(wert):
    """Zahlen numerisch, alles andere alphabetisch — und Leeres immer
    zuletzt, damit ein Klick auf „Auffälligkeit“ die gefüllten Zeilen
    zeigt und nicht dreihundert leere."""
    text = str(wert if wert is not None else "").strip()
    try:
        return (0, float(text), "")
    except ValueError:
        return (1 if text else 2, 0.0, text.lower())


def _name(satz: dict) -> str:
    person = " ".join(x for x in (str(satz.get("Vorname") or ""),
                                  str(satz.get("Name") or "")) if x).strip()
    firma = str(satz.get("Firma") or satz.get("Institution") or "").strip()
    if person and firma:
        return f"{person} · {firma}"
    return person or firma


class App(Tk):
    def __init__(self):
        super().__init__()
        self.title("Mailinglisten-Abgleich")
        self.geometry("1360x860")
        self.minsize(1100, 700)

        self.cfg = self._lade_config()
        self.access_pfad = StringVar(value=self.cfg.get("access", ""))
        self.kunden_pfad = StringVar(value=self.cfg.get("kunden", ""))
        self.auftrag_pfade: list[str] = list(self.cfg.get("auftraege", []))

        self.zuordnungen: list[core.Zuordnung] = []
        self.access: list[dict] = []
        self.access_spalten: list[str] = []
        self.befund = None
        self.access_warnung = None
        # Baumeintrag-ID -> Zuordnung, je Reiter
        self.zeilen: dict[str, dict[str, core.Zuordnung]] = {}
        self.baeume: dict[str, ttk.Treeview] = {}
        # Reiter -> (Spalte, absteigend) oder None für die inhaltliche
        # Vorgabe „Fragwürdiges zuerst“
        self.sortierung: dict[str, tuple | None] = {}
        self.detail_rahmen: dict[str, ttk.Frame] = {}

        self._baue_ui()

    # -- Konfiguration ------------------------------------------------

    def _lade_config(self) -> dict:
        try:
            return json.loads(core.CONFIG_PFAD.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return {}

    def _sichere_config(self):
        try:
            core.CONFIG_PFAD.write_text(json.dumps({
                "access": self.access_pfad.get(),
                "kunden": self.kunden_pfad.get(),
                "auftraege": self.auftrag_pfade,
            }, indent=2, ensure_ascii=False), encoding="utf-8")
        except OSError as e:
            self._log(f"config.json nicht schreibbar: {e}")

    # -- Aufbau -------------------------------------------------------

    def _baue_ui(self):
        pad = {"padx": 12, "pady": 6}

        kopf = ttk.Frame(self)
        kopf.pack(fill="x", **pad)
        ttk.Label(kopf, text="Mailinglisten-Abgleich",
                  font=("Segoe UI", 14, "bold")).pack(side="left")

        # Dateien
        frm = ttk.LabelFrame(self, text="Eingabedateien")
        frm.pack(fill="x", **pad)

        self._dateizeile(frm, "Access-Volltabelle (xlsx)", self.access_pfad)
        self._dateizeile(frm, "Lexware-Kunden (xlsx)", self.kunden_pfad)

        auf = ttk.Frame(frm)
        auf.pack(fill="x", padx=8, pady=(2, 8))
        ttk.Label(auf, text="Aufträge (xlsx, mehrere)", width=26,
                  anchor="w").pack(side="left")
        self.auftrag_liste = Listbox(auf, height=3, selectmode=SINGLE)
        self.auftrag_liste.pack(side="left", fill="x", expand=True, padx=(0, 8))
        for p in self.auftrag_pfade:
            self.auftrag_liste.insert(END, p)
        knoepfe = ttk.Frame(auf)
        knoepfe.pack(side="left")
        ttk.Button(knoepfe, text="Hinzufügen…",
                   command=self._auftrag_hinzu).pack(fill="x")
        ttk.Button(knoepfe, text="Entfernen",
                   command=self._auftrag_weg).pack(fill="x", pady=(4, 0))

        leiste = ttk.Frame(self)
        leiste.pack(fill="x", **pad)
        ttk.Button(leiste, text="Abgleichen",
                   command=self._abgleichen).pack(side="left")
        ttk.Button(leiste, text="Dateien schreiben",
                   command=self._schreiben).pack(side="left", padx=8)
        self.zaehler = StringVar(value="noch nicht abgeglichen")
        ttk.Label(leiste, textvariable=self.zaehler).pack(side="left", padx=16)

        # Warnbanner — normalerweise unsichtbar
        self.warnung = StringVar(value="")
        self.warn_label = ttk.Label(self, textvariable=self.warnung,
                                    foreground="#a00000", wraplength=1300,
                                    justify="left")

        self.reiter = ttk.Notebook(self)
        self.reiter.pack(fill="both", expand=True, **pad)
        for fall in (core.FALL_NEU, core.FALL_AKTUALISIEREN,
                     core.FALL_UNKLAR, core.FALL_OHNE_AUFTRAG):
            self._baue_reiter(fall)

        self.verlauf = Text(self, height=6, wrap="word")
        self.verlauf.pack(fill="x", **pad)
        self._log("Bereit. Dateien wählen und auf Abgleichen klicken.")

    def _dateizeile(self, eltern, beschriftung, var):
        z = ttk.Frame(eltern)
        z.pack(fill="x", padx=8, pady=(8, 2))
        ttk.Label(z, text=beschriftung, width=26, anchor="w").pack(side="left")
        ttk.Entry(z, textvariable=var).pack(side="left", fill="x",
                                            expand=True, padx=(0, 8))
        ttk.Button(z, text="Auswählen…",
                   command=lambda: self._waehle(var)).pack(side="left")

    def _baue_reiter(self, fall):
        rahmen = ttk.Frame(self.reiter)
        self.reiter.add(rahmen, text=REITER_TITEL[fall])

        # Liste oben. Bewusst ttk.Treeview statt der Canvas-und-Checkbutton-
        # Bauweise der anderen Tools: hier stehen bis zu ~1000 Zeilen drin,
        # und ein Widget je Zeile wäre spürbar zäh.
        oben = ttk.Frame(rahmen)
        oben.pack(fill="both", expand=True)

        spalten = SPALTEN[fall]
        baum = ttk.Treeview(oben, columns=[s[0] for s in spalten],
                            show="headings", selectmode="browse")
        for schluessel, titel, breite in spalten:
            baum.heading(schluessel, text=titel,
                         command=lambda f=fall, s=schluessel:
                         self._kopf_klick(f, s))
            baum.column(schluessel, width=breite,
                        anchor="center" if schluessel == "sel" else "w")
        vs = ttk.Scrollbar(oben, orient="vertical", command=baum.yview)
        baum.configure(yscrollcommand=vs.set)
        baum.pack(side="left", fill="both", expand=True)
        vs.pack(side="left", fill="y")

        baum.tag_configure("aus", foreground="#999999")
        baum.tag_configure("auffaellig", background="#fff4e0")
        baum.bind("<Button-1>", lambda e, f=fall: self._klick(e, f))
        baum.bind("<<TreeviewSelect>>", lambda e, f=fall: self._zeige_detail(f))
        self.baeume[fall] = baum
        self.zeilen[fall] = {}

        werkzeug = ttk.Frame(rahmen)
        werkzeug.pack(fill="x", pady=(4, 0))
        ttk.Button(werkzeug, text="Alle anhaken",
                   command=lambda f=fall: self._alle(f, True)).pack(side="left")
        ttk.Button(werkzeug, text="Alle abwählen",
                   command=lambda f=fall: self._alle(f, False)
                   ).pack(side="left", padx=6)
        ttk.Separator(werkzeug, orient="vertical").pack(side="left", fill="y",
                                                        padx=10)
        ttk.Button(werkzeug, text="Zeile zurücksetzen",
                   command=lambda f=fall: self._reset_zeile(f)).pack(side="left")
        ttk.Button(werkzeug, text="Alle Entscheidungen verwerfen",
                   command=self._reset_alles).pack(side="left", padx=6)

        detail = ttk.LabelFrame(rahmen, text="Einzelheiten")
        detail.pack(fill="both", expand=True, pady=(6, 0))
        self.detail_rahmen[fall] = detail

    # -- Dateiauswahl -------------------------------------------------

    def _waehle(self, var):
        pfad = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx"), ("Alle Dateien", "*.*")])
        if pfad:
            var.set(pfad)
            self._sichere_config()

    def _auftrag_hinzu(self):
        for pfad in filedialog.askopenfilenames(
                filetypes=[("Excel", "*.xlsx"), ("Alle Dateien", "*.*")]):
            if pfad not in self.auftrag_pfade:
                self.auftrag_pfade.append(pfad)
                self.auftrag_liste.insert(END, pfad)
        self._sichere_config()

    def _auftrag_weg(self):
        for i in reversed(self.auftrag_liste.curselection()):
            self.auftrag_liste.delete(i)
            del self.auftrag_pfade[i]
        self._sichere_config()

    # -- Abgleich -----------------------------------------------------

    def _abgleichen(self):
        if not self.access_pfad.get() or not self.kunden_pfad.get():
            messagebox.showwarning(
                "Dateien fehlen",
                "Access-Volltabelle und Lexware-Kunden werden gebraucht.")
            return
        if not self.auftrag_pfade:
            messagebox.showwarning(
                "Aufträge fehlen",
                "Ohne Aufträge gibt es kein Bestelljahr.\n\n"
                "Die Aufträge müssen mindestens so weit zurückreichen wie "
                "der Kundenexport — im Zweifel das laufende und das vorige "
                "Jahr laden.")
            return
        try:
            self._log("Einlesen …")
            self.access_spalten, access = core.lade_access(self.access_pfad.get())
            self.access = access
            kunden_spalten, kunden = core.lade_lexware_kunden(
                self.kunden_pfad.get())
            auftraege = core.lade_auftraege(self.auftrag_pfade)
            self._log(f"  Access {len(access)}, Lexware {len(kunden)}, "
                      f"Aufträge {len(auftraege)} Zeilen")

            # Fehlende Spalten sofort benennen — sonst rät man später am
            # schlechten Ergebnis herum, woran es lag.
            for zeile in core.pruefe_spalten(
                    self.access_spalten, kunden_spalten,
                    core.auftrags_spalten(auftraege)):
                self._log("  Hinweis: " + zeile)

            self.access_warnung = core.pruefe_access_vollstaendigkeit(
                access, date.today().year)
            if self.access_warnung:
                self._log("  WARNUNG: " + self.access_warnung)

            vorher = len(kunden)
            kunden = core.ergaenze_kunden_aus_auftraegen(kunden, auftraege)
            if len(kunden) > vorher:
                self._log(f"  {len(kunden) - vorher} Bestandskunden aus den "
                          f"Adressspalten der Aufträge nachgebildet")

            lage = core.bestelljahre(auftraege)
            self.befund = core.pruefe_zeitraeume(kunden, lage)

            warnungen = []
            if self.access_warnung:
                warnungen.append(self.access_warnung)
            if self.befund.verdaechtig:
                warnungen.append(self.befund.text)
                self._log("WARNUNG: " + self.befund.text)
            if warnungen:
                self.warnung.set("\n\n".join("⚠  " + w for w in warnungen))
                self.warn_label.pack(fill="x", padx=12, pady=(0, 6),
                                     before=self.reiter)
            else:
                self.warn_label.pack_forget()

            self._log("Abgleichen …")
            self.zuordnungen = core.gleiche_alle_ab(kunden, access, lage)

            # Gleichlautende Access-Sätze wurden zusammengefasst. Die Zahl
            # gehört genannt: sie ist kein Zwischenergebnis, sondern ein
            # Befund über die Datenbank, den der Verlag aufräumen kann.
            dubletten = {i for z in self.zuordnungen for k in z.kandidaten
                         for i in k.dubletten}
            if dubletten:
                self._log(f"  {len(dubletten)} gleichlautende Access-Sätze zu "
                          f"ihren Zwillingen zusammengefasst — sie standen "
                          f"sonst als Auswahl da, bei der es nichts zu wählen "
                          f"gibt. Welche, steht in protokoll.xlsx unter "
                          f"„Access-Dubletten“.")
            self._lade_entscheidungen()
            self._fuelle_alle()
            self._sichere_config()
            self._log("Fertig. " + self.zaehler.get())
        except Exception as e:                      # noqa: BLE001
            self._log("FEHLER: " + traceback.format_exc())
            messagebox.showerror("Abgleich fehlgeschlagen", str(e))

    # -- Listen füllen ------------------------------------------------

    def _fuelle_alle(self):
        for fall in self.baeume:
            self._fuelle(fall)
        n = {f: len(self.zeilen[f]) for f in self.baeume}
        # Wer auf "nichts tun" steht, erzeugt keine einzige Ausgabezeile —
        # sein Bestelldatum bleibt also alt, und er kann aus dem Mailing
        # fallen. Diese Zahl gehört sichtbar neben die Töpfe.
        uebergangen = sum(1 for z in self.zuordnungen
                          if z.aktion == core.AKTION_NICHTS)
        self.zaehler.set(
            f"neu {n[core.FALL_NEU]} · aktualisieren "
            f"{n[core.FALL_AKTUALISIEREN]} · unklar {n[core.FALL_UNKLAR]} · "
            f"ohne Auftrag {n[core.FALL_OHNE_AUFTRAG]}"
            f"   —   {uebergangen} übergangen (ohne Bestelldatum-Update)")

    def _fuelle(self, fall):
        baum = self.baeume[fall]
        baum.delete(*baum.get_children())
        self.zeilen[fall] = {}

        # Erst alles bewerten, dann die fragwürdigen nach oben. Wer tausend
        # Zeilen durchsieht, soll die zweifelhaften zuerst sehen und nicht
        # zufällig verteilt zwischen den unstrittigen. Über die Spaltenköpfe
        # lässt sich jederzeit anders sortieren.
        zeilen = []
        for z in self.zuordnungen:
            if z.fall != fall:
                continue
            werte, auffaellig = self._zeilenwerte(fall, z)
            zeilen.append((self._sortschluessel(fall, z, auffaellig),
                           werte, auffaellig, z))

        gewaehlt = self.sortierung.get(fall)
        if gewaehlt is None:
            zeilen.sort(key=lambda t: t[0])
        else:
            schluessel, absteigend = gewaehlt
            i = [s[0] for s in SPALTEN[fall]].index(schluessel)
            zeilen.sort(key=lambda t: _sortwert(t[1][i]), reverse=absteigend)
            # Leere Zellen bleiben in BEIDE Richtungen unten. Sonst brächte
            # ein absteigender Klick auf „Auffälligkeit“ dreihundert leere
            # Zeilen nach oben — das Gegenteil dessen, wonach man sortiert.
            # Der zweite Durchgang wirkt, weil Pythons sort stabil ist.
            zeilen.sort(key=lambda t: 1 if not str(t[1][i] or "").strip() else 0)
        self._kopf_beschriften(fall)

        for _, werte, auffaellig, z in zeilen:
            tags = []
            if auffaellig:
                tags.append("auffaellig")
            if z.aktion == core.AKTION_NICHTS:
                tags.append("aus")
            eintrag = baum.insert("", END, values=werte, tags=tags)
            self.zeilen[fall][eintrag] = z

    def _sortschluessel(self, fall, z, auffaellig):
        """Kleiner Wert = weiter oben = eher anzusehen."""
        punkte = z.bester.punkte if z.bester else 0.0
        kdnr = str(z.lexware.get("Kd.-Nr") or "")

        if fall == core.FALL_AKTUALISIEREN:
            # Viele Abweichungen und ein schwacher Treffer sind das, worüber
            # man zweimal schauen will; die glatten 100er ohne Abweichung
            # brauchen niemanden.
            return (0 if auffaellig else 1, punkte, kdnr)
        if fall == core.FALL_UNKLAR:
            # Nicht nach Punkten: die schwachen Treffer (60–65) sind die
            # LEICHTEN Fälle — offensichtlich kein Treffer, also neu anlegen.
            # Teuer sind die anderen: ein gesperrter Satz (verstorben,
            # verzogen), den man versehentlich reaktiviert, und zwei gleich
            # gute Kandidaten, bei denen die falsche Wahl zwei Menschen zu
            # einem macht. Die zuerst.
            zweiter = z.kandidaten[1].punkte if len(z.kandidaten) > 1 else 0.0
            abstand = punkte - zweiter if len(z.kandidaten) > 1 else 999
            # Wer ohnehin Post bekommt, ganz nach unten: dort ist die
            # Entscheidung folgenlos, und Zeit hat man nur einmal.
            folgenlos = bool(z.kandidaten) and all(
                core.dauerhaft_im_mailing(k.access) for k in z.kandidaten)
            if folgenlos:
                rang = 3
            elif z.bester and z.bester.gesperrt:
                rang = 0
            elif zweiter >= core.SCHWELLE_UNKLAR:
                rang = 1
            else:
                rang = 2
            # Innerhalb der Gruppe der knappste Vorsprung zuerst: dort ist die
            # Verwechslungsgefahr am groessten.
            return (rang, abstand, -punkte, kdnr)
        # Neu anlegen / Ohne Auftrag: Auffälliges zuerst.
        return (0 if auffaellig else 1, kdnr)

    def _zeilenwerte(self, fall, z):
        lex = z.lexware
        kdnr = str(lex.get("Kd.-Nr") or "")
        haken = HAKEN_AUS if z.aktion == core.AKTION_NICHTS else HAKEN_AN
        auffaellig = False

        # Die Zeilen zeigen den Satz so, wie er nach Häkchen und
        # Handkorrekturen aussieht — nicht den rohen Lexware-Satz. Sonst
        # bliebe in der Liste ein Name stehen, den man unten gerade geändert
        # hat, und die Spalte widerspräche der Maske darunter.
        if fall == core.FALL_NEU:
            satz = core.baue_neuen_satz(z, self.access_spalten,
                                        date.today().year, date.today())
            probleme = core.auffaelligkeiten(lex, satz) + z.hinweise
            auffaellig = bool(probleme)
            return ((haken, kdnr, _name(satz),
                     f'{satz.get("PLZ") or ""} {satz.get("Ort") or ""}'.strip(),
                     f'{satz.get("Straße") or ""} '
                     f'{satz.get("Hausnummer") or ""}'.strip(),
                     satz.get("Bestelldatum") or "—",
                     "; ".join(probleme)), auffaellig)

        if fall == core.FALL_AKTUALISIEREN:
            zeile = core.baue_aktualisierung(z, date.today().year, date.today())
            geaendert = core.geaenderte_felder(z, zeile)
            return ((haken, kdnr, _name(zeile),
                     (z.ziel or {}).get("ID", ""),
                     f"{z.bester.punkte:.0f}" if z.bester else "",
                     zeile.get("Bestelldatum") or "—",
                     ", ".join(geaendert) or "nur Bestelldatum"),
                    bool(geaendert))

        if fall == core.FALL_UNKLAR:
            # Der Abstand zum Zweitplatzierten sagt mehr als die Punktzahl
            # allein: ein Vorsprung von 20 heißt "der Erste ist es", ein
            # Vorsprung von 0 heißt "hier muss jemand hinsehen".
            zweiter = z.kandidaten[1].punkte if len(z.kandidaten) > 1 else None
            abstand = (f"+{z.bester.punkte - zweiter:.0f}"
                       if z.bester and zweiter is not None else "—")
            # Was hängt an der Entscheidung? Steht der Empfänger über seine
            # Kategorie ohnehin dauerhaft im Mailing, ändert das Bestelldatum
            # nichts daran, ob ein Brief kommt. Solche Fälle darf man mit
            # gutem Gewissen liegen lassen.
            dauerhaft = {core.dauerhaft_im_mailing(k.access)
                         for k in z.kandidaten}
            dauerhaft.discard("")
            if z.kandidaten and all(core.dauerhaft_im_mailing(k.access)
                                    for k in z.kandidaten):
                folge = f"bekommt Post ohnehin ({'/'.join(sorted(dauerhaft))})"
            else:
                folge = "Brief hängt am Bestelldatum"
            return ((haken, kdnr, _name(lex), len(z.kandidaten),
                     f"{z.bester.punkte:.0f}" if z.bester else "",
                     abstand, z.bestelljahr or "—", folge,
                     " | ".join(z.unklar_grund)), True)

        return ((haken, kdnr, _name(lex),
                 f'{lex.get("Plz") or ""} {lex.get("Ort") or ""}'.strip(),
                 f'{lex.get("Kundengruppe") or ""}/{lex.get("Branche") or ""}',
                 " | ".join(z.hinweise)), False)

    def _kopf_klick(self, fall, schluessel):
        """Erster Klick sortiert aufsteigend, der nächste absteigend, der
        dritte stellt die anfängliche Ordnung wieder her — sonst käme man an
        die Sortierung „Auffälliges zuerst" nicht mehr heran, ohne den
        Abgleich zu wiederholen."""
        aktuell = self.sortierung.get(fall)
        if aktuell is None or aktuell[0] != schluessel:
            self.sortierung[fall] = (schluessel, False)
        elif not aktuell[1]:
            self.sortierung[fall] = (schluessel, True)
        else:
            self.sortierung[fall] = None
        self._fuelle(fall)

    def _kopf_beschriften(self, fall):
        """Pfeil an die Spalte setzen, nach der geordnet ist."""
        gewaehlt = self.sortierung.get(fall)
        for schluessel, titel, _ in SPALTEN[fall]:
            if gewaehlt and gewaehlt[0] == schluessel:
                zeichen = " ▼" if gewaehlt[1] else " ▲"
            elif gewaehlt is None and VORGABE_SPALTE.get(fall) == schluessel:
                zeichen = " ◆"          # die anfängliche, inhaltliche Ordnung
            else:
                zeichen = ""
            self.baeume[fall].heading(schluessel, text=titel + zeichen)

    # -- Häkchen ------------------------------------------------------

    def _klick(self, ereignis, fall):
        baum = self.baeume[fall]
        if baum.identify_region(ereignis.x, ereignis.y) != "cell":
            return
        if baum.identify_column(ereignis.x) != "#1":
            return
        eintrag = baum.identify_row(ereignis.y)
        if eintrag:
            self._umschalten(fall, eintrag)
            return "break"

    def _umschalten(self, fall, eintrag, an=None):
        z = self.zeilen[fall][eintrag]
        aus = z.aktion == core.AKTION_NICHTS
        an = aus if an is None else an
        if an:
            z.aktion = (core.AKTION_AKTUALISIEREN
                        if fall == core.FALL_AKTUALISIEREN and z.ziel
                        else core.AKTION_NEU)
        else:
            z.aktion = core.AKTION_NICHTS
        z.beruehrt = True
        baum = self.baeume[fall]
        baum.set(eintrag, "sel", HAKEN_AN if an else HAKEN_AUS)
        tags = [t for t in baum.item(eintrag, "tags") if t != "aus"]
        baum.item(eintrag, tags=tags if an else tags + ["aus"])

    def _alle(self, fall, an):
        for eintrag in self.baeume[fall].get_children():
            self._umschalten(fall, eintrag, an)

    # -- Zurücksetzen -------------------------------------------------

    def _reset_zeile(self, fall):
        auswahl = self.baeume[fall].selection()
        if not auswahl:
            messagebox.showinfo("Nichts gewählt",
                                "Erst eine Zeile in der Liste anklicken.")
            return
        z = self.zeilen[fall][auswahl[0]]
        core.zuruecksetzen(z)
        z.beruehrt = False
        self._fuelle_alle()
        self._springe_zu(z, z.fall)
        self._log(f'Kd.-Nr. {z.lexware.get("Kd.-Nr")} zurückgesetzt '
                  f'auf „{REITER_TITEL[z.fall]}“.')

    def _reset_alles(self):
        if not self.zuordnungen:
            return
        if not messagebox.askyesno(
                "Alle Entscheidungen verwerfen",
                "Sämtliche Häkchen, Zuweisungen und Handkorrekturen dieses "
                "Laufs gehen verloren und alles steht wieder so da, wie der "
                "Abgleich es vorgeschlagen hat.\n\nDie eingelesenen Dateien "
                "bleiben unberührt. Fortfahren?"):
            return
        for z in self.zuordnungen:
            core.zuruecksetzen(z)
            z.beruehrt = False
        # Auch die Sicherung leeren, sonst kehrten die verworfenen
        # Entscheidungen beim nächsten Start zurück.
        self._entscheidungspfad().unlink(missing_ok=True)
        self._fuelle_alle()
        self._log(f"{len(self.zuordnungen)} Fälle auf den Vorschlag des "
                  f"Abgleichs zurückgesetzt.")

    # -- Einzelheiten -------------------------------------------------

    def _zeige_detail(self, fall):
        rahmen = self.detail_rahmen[fall]
        for kind in rahmen.winfo_children():
            kind.destroy()
        auswahl = self.baeume[fall].selection()
        if not auswahl:
            return
        z = self.zeilen[fall][auswahl[0]]

        if fall == core.FALL_AKTUALISIEREN:
            self._detail_abweichungen(rahmen, z, fall)
        elif fall == core.FALL_UNKLAR:
            self._detail_kandidaten(rahmen, z, fall)
        else:
            self._detail_felder(rahmen, z, fall)

    def _detail_felder(self, rahmen, z, fall):
        """Alle Felder des künftigen Access-Satzes, frei änderbar.

        Bewusst ALLE, auch die leeren: gerade an die muss man ran. Steht in
        einem Lexware-Satz die Straße im PLZ-Feld, sind `Straße` und
        `Hausnummer` leer — und wären sie ausgeblendet, ließe sich der Satz
        hier gar nicht geradeziehen.
        """
        satz = core.baue_neuen_satz(z, self.access_spalten, date.today().year,
                                    date.today())

        # Auch hier muss man die Entscheidung zurücknehmen können: wer im
        # Reiter "Aktualisieren" auf "neu anlegen" geklickt hat, landet hier
        # und soll es sich anders überlegen dürfen, ohne den Abgleich neu zu
        # fahren.
        if z.kandidaten and fall == core.FALL_NEU:
            leiste = ttk.Frame(rahmen)
            leiste.pack(fill="x", padx=8, pady=(6, 0))
            ttk.Label(leiste,
                      text=f"Es gäbe {len(z.kandidaten)} möglichen Treffer in "
                           f"Access — stattdessen aktualisieren:").pack(
                side="left")
            for k in z.kandidaten[:3]:
                ttk.Button(
                    leiste, text=f'ID {k.access.get("ID")} ({k.punkte:.0f})',
                    command=lambda kk=k, zz=z: self._waehle_ziel(
                        zz, kk, core.FALL_NEU)).pack(side="left", padx=4)

        ttk.Label(rahmen, text="So wird der Access-Satz angelegt "
                               "(Änderungen wirken sofort; leere Felder sind "
                               "ausfüllbar):").pack(anchor="w", padx=8,
                                                    pady=(6, 4))
        self._feldmaske(rahmen, z, fall, satz)

    def _scrollbereich(self, eltern) -> ttk.Frame:
        """Ein scrollbarer Bereich; gibt den Rahmen zurück, in den man legt.

        Der Einzelheiten-Bereich hat feste Höhe, sein Inhalt nicht: 50 Felder
        beim Anlegen, bis zu acht Kandidaten beim Unklaren. Ohne Rollbalken
        rutschen die Knöpfe darunter aus dem Bild — und dann kommt man an sie
        gar nicht mehr heran.
        """
        koerper = ttk.Frame(eltern)
        koerper.pack(fill="both", expand=True)
        leinwand = Canvas(koerper, highlightthickness=0)
        vs = ttk.Scrollbar(koerper, orient="vertical", command=leinwand.yview)
        inhalt = ttk.Frame(leinwand)
        inhalt.bind("<Configure>", lambda e: leinwand.configure(
            scrollregion=leinwand.bbox("all")))
        leinwand.create_window((0, 0), window=inhalt, anchor="nw")
        leinwand.configure(yscrollcommand=vs.set)
        leinwand.pack(side="left", fill="both", expand=True)
        vs.pack(side="left", fill="y")

        def rollen(schritt):
            # bind_all bleibt hängen, wenn der Bereich längst neu gebaut wurde
            # — dann zeigt der Verweis auf ein zerstörtes Widget.
            try:
                leinwand.yview_scroll(schritt, "units")
            except TclError:
                pass

        leinwand.bind_all("<MouseWheel>",
                          lambda e: rollen(int(-e.delta / 120)))
        leinwand.bind_all("<Button-4>", lambda e: rollen(-1))
        leinwand.bind_all("<Button-5>", lambda e: rollen(1))
        return inhalt

    def _feldmaske(self, rahmen, z, fall, satz: dict):
        """Ein scrollbares Gitter aus Beschriftung und Eingabefeld je Feld."""
        gitter = self._scrollbereich(rahmen)
        gaengig, rest = core.ordne_felder(satz)
        zeile = self._feldgruppe(gitter, z, fall, gaengig, 0)
        if rest:
            ttk.Separator(gitter, orient="horizontal").grid(
                row=zeile, column=0, columnspan=6, sticky="ew", pady=(8, 2))
            ttk.Label(gitter, text="Weitere Felder",
                      foreground="#666666").grid(
                row=zeile + 1, column=0, columnspan=6, sticky="w", padx=8)
            self._feldgruppe(gitter, z, fall, rest, zeile + 2)

    def _feldgruppe(self, gitter, z, fall, felder, ab_zeile: int) -> int:
        """Legt die Felder dreispaltig ab. Gibt die nächste freie Zeile zurück."""
        for i, (feld, wert) in enumerate(felder):
            r, s = divmod(i, 3)
            r += ab_zeile
            leer = wert in (None, "")
            ttk.Label(gitter, text=feld + ":", width=20, anchor="e",
                      foreground="#999999" if leer else "black").grid(
                row=r, column=s * 2, sticky="e", padx=(8, 2), pady=1)
            if feld == "ID":
                # Der Schlüssel, über den die UPDATE-Abfrage verbindet — den
                # zu ändern hieße, einen fremden Datensatz zu überschreiben.
                ttk.Label(gitter, text=str(wert), foreground="#666666").grid(
                    row=r, column=s * 2 + 1, sticky="w", padx=(0, 8))
                continue
            var = StringVar(value="" if leer else str(wert))
            ttk.Entry(gitter, textvariable=var, width=26).grid(
                row=r, column=s * 2 + 1, sticky="w", padx=(0, 8))
            var.trace_add("write",
                          lambda *_, f=feld, v=var, zz=z, fa=fall:
                          self._feld_geaendert(zz, f, v.get(), fa))
        return ab_zeile + (len(felder) + 2) // 3

    def _feld_geaendert(self, z, feld, wert, fall):
        z.aenderungen[feld] = wert
        z.beruehrt = True
        self._frische_zeile(z, fall)

    def _frische_zeile(self, z, fall):
        """Die Listenzeile neu berechnen — nach jeder Änderung an der
        Entscheidung, gleich ob Häkchen oder getipptes Feld."""
        for eintrag, zz in self.zeilen[fall].items():
            if zz is not z:
                continue
            werte, auffaellig = self._zeilenwerte(fall, z)
            baum = self.baeume[fall]
            # Auch die Einfärbung neu setzen — sonst bleibt eine Zeile orange,
            # deren Auffälligkeit man gerade behoben hat.
            tags = []
            if auffaellig:
                tags.append("auffaellig")
            if z.aktion == core.AKTION_NICHTS:
                tags.append("aus")
            baum.item(eintrag, values=werte, tags=tags)
            return

    def _detail_abweichungen(self, rahmen, z, fall):
        abw = core.abweichungen(z)
        kopf = (f'Access-ID {(z.ziel or {}).get("ID")} · '
                f'{z.bester.punkte:.0f} Punkte · {z.bester.begruendung}'
                if z.bester else "")
        ttk.Label(rahmen, text=kopf).pack(anchor="w", padx=8, pady=(6, 4))

        if abw:
            tabelle = ttk.Frame(rahmen)
            tabelle.pack(fill="x", padx=8)
            for spalte, titel in enumerate(("Feld", "in Access", "aus Lexware",
                                            "übernehmen")):
                ttk.Label(tabelle, text=titel,
                          font=("Segoe UI", 9, "bold")).grid(
                    row=0, column=spalte, sticky="w", padx=8, pady=(0, 2))
            for i, (feld, (alt, neu)) in enumerate(sorted(abw.items())):
                # Ist der Access-Wert deutlich länger, steht dort meist mehr
                # als in Lexware — ein ausgeschriebener Name oder ein Vermerk
                # wie "zzz_keine Werbeanrufe!". Übernehmen löscht das. Deshalb
                # hervorheben, damit es beim Durchsehen nicht untergeht.
                aermer = bool(alt) and len(alt) > len(neu) + 3
                ttk.Label(tabelle, text=("⚠ " if aermer else "") + feld,
                          foreground="#a06000" if aermer else "black").grid(
                    row=i + 1, column=0, sticky="w", padx=8)
                ttk.Label(tabelle, text=(alt or "—").replace("\n", " ⏎ "),
                          foreground="#a06000" if aermer else "#666666").grid(
                    row=i + 1, column=1, sticky="w", padx=8)
                ttk.Label(tabelle, text=neu).grid(row=i + 1, column=2,
                                                  sticky="w", padx=8)
                var = StringVar(value="1" if feld in z.uebernehmen else "0")
                ttk.Checkbutton(
                    tabelle, variable=var, onvalue="1", offvalue="0",
                    command=lambda f=feld, v=var, zz=z, fa=fall:
                    self._uebernahme(zz, f, v.get() == "1", fa)).grid(
                    row=i + 1, column=3, sticky="w", padx=8)
        else:
            ttk.Label(rahmen, text="Keine Abweichungen — es wird nur das "
                                   "Bestelldatum gesetzt.").pack(anchor="w",
                                                                 padx=8)

        # Weicht die Anschrift stark ab, ist es oft gar kein Umzug, sondern
        # eine zweite Adresse desselben Menschen — privat und dienstlich. Beide
        # sollen dann einen Brief bekommen, also braucht es hier dieselbe Wahl
        # wie im Reiter "Unklar" und nicht nur die Übernahme-Häkchen.
        leiste = ttk.Frame(rahmen)
        leiste.pack(fill="x", padx=8, pady=(8, 0))
        ttk.Label(leiste, text="Stattdessen:").pack(side="left")
        ttk.Button(leiste, text="neuen Datensatz anlegen (Zweitadresse)",
                   command=lambda zz=z, fa=fall: self._als_neu(zz, fa)).pack(
            side="left", padx=6)
        ttk.Button(leiste, text="nichts tun",
                   command=lambda zz=z, fa=fall: self._als_nichts(zz, fa)).pack(
            side="left")

        # Darunter derselbe Editor wie beim Anlegen, damit sich auch hier
        # etwas von Hand geradeziehen lässt. Gezeigt werden genau die Felder,
        # welche die UPDATE-Abfrage schreibt — was hier steht, steht hinterher
        # in Access, nicht mehr und nicht weniger.
        ttk.Label(rahmen, text="So sieht der Access-Satz hinterher aus "
                               "(Änderungen wirken sofort und haben Vorrang "
                               "vor den Häkchen oben):").pack(
            anchor="w", padx=8, pady=(10, 4))
        zeile = core.baue_aktualisierung(z, date.today().year, date.today())
        self._feldmaske(rahmen, z, fall, zeile)

    def _uebernahme(self, z, feld, an, fall):
        if an:
            z.uebernehmen.add(feld)
        else:
            z.uebernehmen.discard(feld)
        z.beruehrt = True
        # Beides hängt am Häkchen: die Zeile oben (was geändert wird) und der
        # Editor unten (der Endwert).
        self._frische_zeile(z, fall)
        self._zeige_detail(fall)

    def _detail_kandidaten(self, rahmen, z, fall):
        """Oben, was gilt und was zu tun ist; darunter scrollbar die Auswahl.

        Die Lexware-Anschrift steht NICHT mehr als Kopfzeile da: sie erscheint
        im Gitter darunter ohnehin, Feld für Feld und damit vergleichbar.
        Zweimal dasselbe zu zeigen macht die Maske nur voll.
        """
        lex = z.lexware
        kopf = ttk.Frame(rahmen)
        kopf.pack(fill="x", padx=8, pady=(6, 0))
        jahr = (f"Bestelljahr {z.bestelljahr}" if z.bestelljahr
                else "kein Bestelljahr (Freiexemplar)")
        ttk.Label(kopf, text=f'Kd.-Nr. {lex.get("Kd.-Nr")}  ·  {jahr}',
                  foreground="#666666").pack(anchor="w")
        for g in z.unklar_grund:
            ttk.Label(rahmen, text="· " + g, foreground="#a06000").pack(
                anchor="w", padx=8)
        for h in z.hinweise:
            ttk.Label(rahmen, text="· " + h, foreground="#888888").pack(
                anchor="w", padx=8)

        anzahl = len(z.kandidaten[:8])
        frage = ttk.Frame(rahmen)
        frage.pack(fill="x", padx=8, pady=(8, 4))
        ttk.Label(frage,
                  text=(f"{anzahl} passender Satz in Access." if anzahl == 1
                        else f"{anzahl} passende Sätze in Access."),
                  font=("Segoe UI", 9, "bold")).pack(side="left")
        # Diese beiden Knöpfe stehen ÜBER der Liste, nicht darunter: bei acht
        # Kandidaten waren sie sonst aus dem sichtbaren Bereich geschoben und
        # damit unerreichbar.
        ttk.Button(frage, text="stattdessen neu anlegen (Zweitadresse)",
                   command=lambda zz=z, fa=fall: self._als_neu(zz, fa)).pack(
            side="right")
        ttk.Button(frage, text="nichts tun",
                   command=lambda zz=z, fa=fall:
                   self._als_nichts(zz, fa)).pack(side="right", padx=6)

        gitter = self._scrollbereich(rahmen)
        self._kandidatengitter(gitter, z, fall)

    # Farben des Feldvergleichs. Übereinstimmung ist die Nachricht, deshalb
    # ist sie kräftig; Nichtwissen bleibt blass, damit eine fehlende E-Mail
    # nicht wie eine falsche aussieht.
    FARBE = {core.GLEICH: "#0a7a3a", core.ANDERS: "#b03030",
             core.UNBEKANNT: "#aaaaaa"}
    ZEICHEN = {core.GLEICH: "✓", core.ANDERS: "✗", core.UNBEKANNT: "·"}

    def _kandidatengitter(self, gitter, z, fall):
        """Kandidaten Feld für Feld gegen den Lexware-Satz stellen.

        Bei einem einzigen Kandidaten untereinander — da ist Platz und man
        will jedes Feld lesen. Bei mehreren nebeneinander, sonst müsste man
        scrollen, um zwei Kandidaten zu vergleichen, und genau das ist ja die
        Aufgabe.
        """
        kandidaten = z.kandidaten[:8]
        if not kandidaten:
            return

        if len(kandidaten) == 1:
            k = kandidaten[0]
            ttk.Label(gitter, text="Feld", font=("Segoe UI", 9, "bold")).grid(
                row=0, column=0, sticky="w", padx=(0, 12))
            ttk.Label(gitter, text="Lexware",
                      font=("Segoe UI", 9, "bold")).grid(row=0, column=1,
                                                         sticky="w", padx=12)
            dublett = (f' (+ {len(k.dubletten)} gleichlautende: '
                       f'{", ".join(str(i) for i in k.dubletten)})'
                       if k.dubletten else "")
            ttk.Label(gitter, text=f'Access ID {k.access.get("ID")}{dublett}',
                      font=("Segoe UI", 9, "bold")).grid(row=0, column=2,
                                                         sticky="w", padx=12)
            for i, (feld, l, a, befund) in enumerate(
                    core.vergleiche(z.lexware, k.access), start=1):
                farbe = self.FARBE[befund]
                ttk.Label(gitter, text=f"{self.ZEICHEN[befund]} {feld}",
                          foreground=farbe).grid(row=i, column=0, sticky="w",
                                                 padx=(0, 12))
                ttk.Label(gitter, text=l or "—", foreground=farbe).grid(
                    row=i, column=1, sticky="w", padx=12)
                ttk.Label(gitter, text=a or "—", foreground=farbe).grid(
                    row=i, column=2, sticky="w", padx=12)
            self._kandidatenknopf(gitter, z, k, fall,
                                  len(core.VERGLEICHSFELDER) + 1, 0,
                                  "diesen Access-Satz aktualisieren")
            return

        # Mehrere: Felder als Spalten, Kandidaten als Zeilen.
        werte = core.lexware_werte(z.lexware)
        anzahl = len(core.VERGLEICHSFELDER)

        # Kopfzeile mit den FELDNAMEN. Ohne sie weiß niemand, was in Spalte 4
        # steht — die Werte allein sagen es nicht.
        ttk.Label(gitter, text="", width=16).grid(row=0, column=0)
        for s, feld in enumerate(core.VERGLEICHSFELDER, start=1):
            ttk.Label(gitter, text=feld, font=("Segoe UI", 8, "bold"),
                      foreground="#555555").grid(row=0, column=s, sticky="w",
                                                 padx=6)

        # Die Bezugszeile: das, wogegen alles darunter verglichen wird.
        ttk.Label(gitter, text="Lexware", font=("Segoe UI", 9, "bold")).grid(
            row=1, column=0, sticky="w", padx=(0, 8))
        for s, feld in enumerate(core.VERGLEICHSFELDER, start=1):
            ttk.Label(gitter, text=str(werte.get(feld) or "—")[:22],
                      font=("Segoe UI", 9, "bold")).grid(row=1, column=s,
                                                         sticky="w", padx=6)
        ttk.Separator(gitter, orient="horizontal").grid(
            row=2, column=0, columnspan=anzahl + 3, sticky="ew", pady=3)

        for i, k in enumerate(kandidaten):
            r = i + 3
            vorsprung = ""
            if i == 0 and len(kandidaten) > 1:
                vorsprung = f" (+{k.punkte - kandidaten[1].punkte:.0f})"
            dublett = (f' +{len(k.dubletten)} gleiche' if k.dubletten else "")
            ttk.Label(gitter,
                      text=f'{k.punkte:.0f}{vorsprung} · '
                           f'ID {k.access.get("ID")}{dublett}',
                      font=("Segoe UI", 9, "bold") if i == 0 else None,
                      foreground="#a00000" if k.gesperrt else "black").grid(
                row=r, column=0, sticky="w", padx=(0, 8))
            for s, (feld, l, a, befund) in enumerate(
                    core.vergleiche(z.lexware, k.access), start=1):
                ttk.Label(gitter, text=(a or "—")[:22],
                          foreground=self.FARBE[befund]).grid(
                    row=r, column=s, sticky="w", padx=6)
            # Beschriftung ausgeschrieben: "übernehmen" liest sich wie
            # "diesen Satz neu anlegen", gemeint ist aber das Gegenteil —
            # in diesen vorhandenen Access-Satz hineinschreiben.
            self._kandidatenknopf(gitter, z, k, fall, r, anzahl + 1,
                                  "diesen aktualisieren")

        ttk.Label(gitter,
                  text="grün = stimmt mit Lexware überein   ·   "
                       "rot = weicht ab   ·   grau = auf einer Seite leer",
                  foreground="#777777", font=("Segoe UI", 8)).grid(
            row=len(kandidaten) + 3, column=0, columnspan=anzahl + 3,
            sticky="w", pady=(6, 0))

    def _kandidatenknopf(self, gitter, z, k, fall, zeile, spalte, text):
        if k.gesperrt:
            ttk.Label(gitter, text=f"⛔ {k.sperrgrund}",
                      foreground="#a00000").grid(row=zeile, column=spalte,
                                                 sticky="w", padx=8)
        ttk.Button(gitter, text=text,
                   command=lambda: self._waehle_ziel(z, k, fall)).grid(
            row=zeile, column=spalte + 1, sticky="w", padx=8, pady=1)

    def _waehle_ziel(self, z, kandidat, fall):
        z.ziel = kandidat.access
        z.beruehrt = True
        z.aktion = core.AKTION_AKTUALISIEREN
        z.fall = core.FALL_AKTUALISIEREN
        if core.UEBERNAHME_VORGABE:
            # Gleiche Vorbelegung wie bei den eindeutigen Treffern, sonst
            # verhielten sich von Hand zugewiesene Fälle anders als der Rest.
            z.uebernehmen = set(core.abweichungen(z))
        self._fuelle_alle()
        self._log(f'Kd.-Nr. {z.lexware.get("Kd.-Nr")} → Access-ID '
                  f'{kandidat.access.get("ID")} (aktualisieren)')
        self._springe_zu(z, core.FALL_AKTUALISIEREN)

    def _als_neu(self, z, fall):
        z.ziel = None
        z.beruehrt = True
        z.aktion = core.AKTION_NEU
        z.fall = core.FALL_NEU
        self._fuelle_alle()
        self._log(f'Kd.-Nr. {z.lexware.get("Kd.-Nr")} → neuer Datensatz')
        self._springe_zu(z, core.FALL_NEU)

    def _springe_zu(self, z, fall):
        """Nach einer Entscheidung im Reiter „Unklar" den Fall dort öffnen, wo
        er gelandet ist — sonst müsste man ihn unter tausend Zeilen suchen,
        nur um noch eine Kleinigkeit nachzubessern."""
        for eintrag, zz in self.zeilen[fall].items():
            if zz is z:
                self.reiter.select(list(self.baeume).index(fall))
                baum = self.baeume[fall]
                baum.selection_set(eintrag)
                baum.see(eintrag)
                self._zeige_detail(fall)
                return

    def _als_nichts(self, z, fall):
        z.beruehrt = True
        z.aktion = core.AKTION_NICHTS
        self._fuelle_alle()

    # -- Entscheidungen sichern ---------------------------------------

    def _ausgabeordner(self) -> Path:
        return core.APP_DIR / "mailing_output" / str(date.today().year)

    def _entscheidungspfad(self) -> Path:
        return self._ausgabeordner() / "entscheidungen.json"

    def _sichere_entscheidungen(self):
        """Die Durchsicht von über tausend Fällen soll eine Mittagspause oder
        einen Absturz überleben."""
        daten = {}
        for z in self.zuordnungen:
            kdnr = str(z.lexware.get("Kd.-Nr") or "")
            if not kdnr or not z.beruehrt:
                continue
            daten[kdnr] = {
                "aktion": z.aktion,
                "fall": z.fall,
                "ziel_id": (z.ziel or {}).get("ID"),
                "uebernehmen": sorted(z.uebernehmen),
                "aenderungen": {k: str(v) for k, v in z.aenderungen.items()},
            }
        pfad = self._entscheidungspfad()
        pfad.parent.mkdir(parents=True, exist_ok=True)
        pfad.write_text(json.dumps(daten, indent=1, ensure_ascii=False),
                        encoding="utf-8")

    def _lade_entscheidungen(self):
        pfad = self._entscheidungspfad()
        try:
            daten = json.loads(pfad.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return
        nach_id = {}
        for z in self.zuordnungen:
            for k in z.kandidaten:
                nach_id[k.access.get("ID")] = k.access
        wieder = 0
        for z in self.zuordnungen:
            eintrag = daten.get(str(z.lexware.get("Kd.-Nr") or ""))
            if not eintrag:
                continue
            z.aktion = eintrag.get("aktion", z.aktion)
            z.fall = eintrag.get("fall", z.fall)
            if eintrag.get("ziel_id") is not None:
                z.ziel = nach_id.get(eintrag["ziel_id"], z.ziel)
            z.uebernehmen = set(eintrag.get("uebernehmen", []))
            z.aenderungen = dict(eintrag.get("aenderungen", {}))
            z.beruehrt = True
            wieder += 1
        if wieder:
            self._log(f"{wieder} frühere Entscheidungen aus "
                      f"entscheidungen.json übernommen.")

    # -- Schreiben ----------------------------------------------------

    def _schreiben(self):
        if not self.zuordnungen:
            messagebox.showwarning("Nichts da", "Erst abgleichen.")
            return
        try:
            ordner = self._ausgabeordner()
            self._sichere_entscheidungen()
            ergebnis = core.schreibe_alles(
                ordner, self.zuordnungen, self.access, self.access_spalten,
                date.today().year, date.today(), self.befund,
                self.access_warnung)
            self._log(f"Geschrieben nach {ordner}:")
            for name, wert in ergebnis.items():
                self._log(f"   {name}: {wert}")
            _, zusammengefuehrt = core.verschmolzene_aktualisierungen(
                self.zuordnungen, date.today().year, date.today())
            for anmerkung in zusammengefuehrt:
                self._log("   " + anmerkung)

            uebergangen = [z for z in self.zuordnungen
                           if z.aktion == core.AKTION_NICHTS]
            mit_jahr = sum(1 for z in uebergangen if z.bestelljahr)
            if uebergangen:
                self._log(
                    f"   {len(uebergangen)} Fälle übergangen — davon "
                    f"{mit_jahr} mit einem Bestelljahr, das damit NICHT nach "
                    f"Access kommt. Wer dort ein altes Datum stehen hat und "
                    f"kein Merkmal trägt, fällt aus dem Mailing.")
            messagebox.showinfo(
                "Fertig",
                f"Dateien liegen in\n{ordner}\n\n"
                f"neu anlegen: {ergebnis['neu_anlegen.xlsx']}\n"
                f"aktualisieren: {ergebnis['aktualisieren.xlsx']}\n\n"
                "Wie sie nach Access kommen, steht in access_import.txt.")
        except Exception as e:                      # noqa: BLE001
            self._log("FEHLER: " + traceback.format_exc())
            messagebox.showerror("Schreiben fehlgeschlagen", str(e))

    # -- Verlauf ------------------------------------------------------

    def _log(self, text):
        self.verlauf.insert(END, text + "\n")
        self.verlauf.see(END)
        self.update_idletasks()


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
