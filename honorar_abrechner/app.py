"""
Honorar-Abrechner — Oberfläche
===============================
Vier Schritte in der Reihenfolge, in der im Januar gearbeitet wird:
Stammdaten pflegen, Stückzahlen erfassen, durchrechnen, Ausgaben erzeugen.

Alles, was länger dauert als ein Lidschlag, läuft in einem Arbeitsthread.
Arbeitsthreads fassen Tk NICHT an — sie legen Nachrichten in eine
Warteschlange, und nur der Hauptthread nimmt sie heraus (``_pumpe``).
"""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import traceback
from datetime import date
from pathlib import Path
from tkinter import (Tk, StringVar, BooleanVar, IntVar, END, NORMAL, DISABLED,
                     filedialog, messagebox, ttk, Text, Toplevel)

from honorar_abrechner import core


def _spur() -> None:
    """Fehlerspur ausgeben, wenn es ein stderr gibt.

    Der Fensterbau (``console=False``) hat keines: dort ist ``sys.stderr``
    None, und ``traceback.print_exc()`` schriebe in ein Nichts — was einen
    ZWEITEN Fehler auslöst, ausgerechnet im Fehlerzweig.
    """
    if sys.stderr is not None:
        traceback.print_exc()


PAD = {"padx": 12, "pady": 6}

SPALTEN_STAMM = [("name", "Empfänger", 300), ("ort", "PLZ / Ort", 170),
                 ("iban", "Bankverbindung", 230), ("buecher", "Bücher", 70),
                 ("hinweis", "Hinweis", 240)]

SPALTEN_ERFASSUNG = [("isbn", "ISBN", 90), ("titel", "Buchtitel", 320),
                     ("empf", "Empfänger", 220), ("art", "Vergütungsart", 130),
                     ("verkauft", "verk. Ex.", 90), ("eigenkauf", "Eigenkauf", 90),
                     ("stand", "Stand", 110)]

SPALTEN_DURCHLAUF = [("name", "Empfänger", 280), ("art", "Vergütungsart", 120),
                     ("posten", "Bücher", 70), ("netto", "Netto", 95),
                     ("mwst", "MwSt", 85), ("brutto", "Auszahlung", 105),
                     ("brief", "Brief", 60), ("hinweis", "Grund / Hinweis", 340)]


def _sortwert(wert):
    """Zahlen numerisch, alles andere alphabetisch — Leeres immer zuletzt."""
    text = str(wert if wert is not None else "").strip()
    try:
        return (0, float(text.replace(".", "").replace(",", ".")), "")
    except ValueError:
        return (1 if text else 2, 0.0, text.lower())


class App(Tk):
    def __init__(self):
        super().__init__()
        self.title("Honorar-Abrechner")
        self.geometry("1420x880")
        self.minsize(1100, 700)

        try:
            self.cfg = core.lade_config()
        except Exception as e:
            messagebox.showerror("Config-Fehler", str(e))
            self.cfg = dict(core.DEFAULT_CONFIG)

        self.bestand = core.Bestand()
        self.abrechnungen: list[core.Abrechnung] = []
        self.jahr = IntVar(value=date.today().year - 1)
        self.status = StringVar(value="Bereit.")
        self.nur_offene = BooleanVar(value=False)
        self.suche_stamm = StringVar()
        self.suche_erfassung = StringVar()

        # Arbeitsthreads legen hier ab, der Hauptthread holt heraus. Ein
        # after() aus einem fremden Thread verklemmt den Tcl-Interpreter.
        self._nachrichten: queue.Queue = queue.Queue()
        self._sortierung: dict[str, tuple] = {}

        self._baue_ui()
        self.after(150, self._pumpe)
        self.after(200, self._starte_laden)
        self.protocol("WM_DELETE_WINDOW", self._beenden)

    # -----------------------------------------------------------------
    # Aufbau
    # -----------------------------------------------------------------

    def _baue_ui(self):
        kopf = ttk.Frame(self)
        kopf.pack(fill="x", **PAD)
        ttk.Label(kopf, text="Honorar-Abrechner",
                  font=("Segoe UI", 14, "bold")).pack(side="left")
        ttk.Label(kopf, text="Abrechnungsjahr:").pack(side="left", padx=(24, 4))
        ttk.Spinbox(kopf, from_=2000, to=2100, width=6,
                    textvariable=self.jahr,
                    command=self._jahr_gewechselt).pack(side="left")
        ttk.Button(kopf, text="Bestand neu laden",
                   command=self._starte_laden).pack(side="right")

        self.reiter = ttk.Notebook(self)
        self.reiter.pack(fill="both", expand=True, padx=12, pady=(0, 6))
        self._baue_stammdaten()
        self._baue_erfassung()
        self._baue_durchlauf()
        self._baue_ausgaben()

        leiste = ttk.Frame(self)
        leiste.pack(fill="x", padx=12, pady=(0, 10))
        ttk.Label(leiste, textvariable=self.status).pack(side="left")

    def _liste(self, eltern, spalten, doppelklick=None):
        rahmen = ttk.Frame(eltern)
        rahmen.pack(fill="both", expand=True, **PAD)
        baum = ttk.Treeview(rahmen, columns=[s[0] for s in spalten],
                            show="headings", selectmode="browse")
        for schluessel, titel, breite in spalten:
            baum.heading(schluessel, text=titel,
                         command=lambda b=baum, s=schluessel: self._sortiere(b, s))
            baum.column(schluessel, width=breite, stretch=(breite > 200))
        leiste = ttk.Scrollbar(rahmen, orient="vertical", command=baum.yview)
        baum.configure(yscrollcommand=leiste.set)
        baum.pack(side="left", fill="both", expand=True)
        leiste.pack(side="right", fill="y")
        if doppelklick:
            baum.bind("<Double-1>", doppelklick)
        return baum

    def _sortiere(self, baum, spalte):
        richtung = self._sortierung.get(id(baum))
        umgekehrt = bool(richtung and richtung[0] == spalte and not richtung[1])
        self._sortierung[id(baum)] = (spalte, umgekehrt)
        zeilen = [(baum.set(k, spalte), k) for k in baum.get_children("")]
        zeilen.sort(key=lambda z: _sortwert(z[0]), reverse=umgekehrt)
        for i, (_, k) in enumerate(zeilen):
            baum.move(k, "", i)

    # --- Reiter 1: Stammdaten -----------------------------------------

    def _baue_stammdaten(self):
        seite = ttk.Frame(self.reiter)
        self.reiter.add(seite, text="Stammdaten")

        oben = ttk.Frame(seite)
        oben.pack(fill="x", **PAD)
        ttk.Label(oben, text="Suchen:").pack(side="left")
        feld = ttk.Entry(oben, textvariable=self.suche_stamm, width=36)
        feld.pack(side="left", padx=6)
        feld.bind("<KeyRelease>", lambda e: self._zeige_stammdaten())
        ttk.Button(oben, text="Bestand speichern",
                   command=self._speichern).pack(side="right")
        ttk.Button(oben, text="Altmappe importieren …",
                   command=self._importieren).pack(side="right", padx=6)

        self.baum_stamm = self._liste(seite, SPALTEN_STAMM,
                                      doppelklick=self._zeige_empfaenger)
        ttk.Label(seite, text="Doppelklick auf eine Zeile zeigt Adresse, "
                              "Bücher und Konditionen.").pack(anchor="w",
                                                              padx=12)

    # --- Reiter 2: Jahreserfassung ------------------------------------

    def _baue_erfassung(self):
        seite = ttk.Frame(self.reiter)
        self.reiter.add(seite, text="Jahreserfassung")

        oben = ttk.Frame(seite)
        oben.pack(fill="x", **PAD)
        ttk.Label(oben, text="Suchen:").pack(side="left")
        feld = ttk.Entry(oben, textvariable=self.suche_erfassung, width=36)
        feld.pack(side="left", padx=6)
        feld.bind("<KeyRelease>", lambda e: self._zeige_erfassung())
        ttk.Checkbutton(oben, text="nur noch nicht erfasste",
                        variable=self.nur_offene,
                        command=self._zeige_erfassung).pack(side="left", padx=12)
        self.fortschritt = StringVar(value="")
        ttk.Label(oben, textvariable=self.fortschritt).pack(side="right")

        self.baum_erfassung = self._liste(seite, SPALTEN_ERFASSUNG)
        self.baum_erfassung.bind("<Double-1>", self._bearbeite_zelle)
        self.baum_erfassung.bind("<Return>", self._bearbeite_zelle)
        ttk.Label(seite, text="Doppelklick oder Eingabetaste auf „verk. Ex.“ "
                              "bzw. „Eigenkauf“ öffnet das Eingabefeld; "
                              "Eingabetaste springt in die nächste Zeile."
                  ).pack(anchor="w", padx=12)

    # --- Reiter 3: Durchlauf ------------------------------------------

    def _baue_durchlauf(self):
        seite = ttk.Frame(self.reiter)
        self.reiter.add(seite, text="Durchlauf")

        oben = ttk.Frame(seite)
        oben.pack(fill="x", **PAD)
        ttk.Button(oben, text="Jetzt durchrechnen",
                   command=self._rechnen).pack(side="left")
        ttk.Button(oben, text="Gegen die Altmappe prüfen …",
                   command=self._gegenprobe).pack(side="left", padx=8)
        self.summe = StringVar(value="")
        ttk.Label(oben, textvariable=self.summe,
                  font=("Segoe UI", 10, "bold")).pack(side="right")

        self.baum_durchlauf = self._liste(seite, SPALTEN_DURCHLAUF)
        self.baum_durchlauf.tag_configure("kein_brief", foreground="#777777")
        self.baum_durchlauf.tag_configure("problem", foreground="#a4262c")

    # --- Reiter 4: Ausgaben -------------------------------------------

    def _baue_ausgaben(self):
        seite = ttk.Frame(self.reiter)
        self.reiter.add(seite, text="Ausgaben")

        knoepfe = ttk.LabelFrame(seite, text="Erzeugen")
        knoepfe.pack(fill="x", **PAD)
        self.btn_briefe = ttk.Button(knoepfe, text="Briefe (.docx)",
                                     command=self._briefe)
        self.btn_briefe.grid(row=0, column=0, **PAD)
        self.btn_pdf = ttk.Button(knoepfe, text="Briefe nach PDF (nur Windows)",
                                  command=self._pdf)
        self.btn_pdf.grid(row=0, column=1, **PAD)
        self.btn_listen = ttk.Button(knoepfe, text="Zahlungsliste + KSK-Liste",
                                     command=self._listen)
        self.btn_listen.grid(row=0, column=2, **PAD)
        ttk.Button(knoepfe, text="Ausgabeordner öffnen",
                   command=self._ordner_oeffnen).grid(row=0, column=3, **PAD)

        self.protokollfeld = Text(seite, height=20, state=DISABLED, wrap="word")
        self.protokollfeld.pack(fill="both", expand=True, **PAD)

    # -----------------------------------------------------------------
    # Warteschlange
    # -----------------------------------------------------------------

    def _melde(self):
        """Protokoll-Rückruf für einen Arbeitsthread — schreibt nur in die
        Warteschlange, nie ins Fenster."""
        return lambda text: self._nachrichten.put(("log", str(text)))

    def _pumpe(self):
        """Läuft im Hauptthread und leert die Warteschlange."""
        fertig = {"laden": self._laden_fertig, "import": self._import_fertig,
                  "rechnen": self._rechnen_fertig, "briefe": self._briefe_fertig,
                  "pdf": self._pdf_fertig, "listen": self._listen_fertig,
                  "speichern": self._speichern_fertig,
                  "gegenprobe": self._gegenprobe_fertig}
        try:
            while True:
                art, *rest = self._nachrichten.get_nowait()
                if art == "log":
                    self._schreibe(rest[0])
                elif art == "status":
                    self.status.set(rest[0])
                elif art == "fertig":
                    sid, erg, fehler = rest
                    fertig[sid](erg, fehler)
        except queue.Empty:
            pass
        except Exception:
            _spur()
        self.after(150, self._pumpe)

    def _schreibe(self, text: str):
        feld = self.protokollfeld
        if feld is None or not feld.winfo_exists():
            return
        feld.configure(state=NORMAL)
        feld.insert(END, text + "\n")
        feld.see(END)
        feld.configure(state=DISABLED)

    def _starte(self, sid: str, arbeit):
        """Einen Arbeitsschritt im Hintergrund anstoßen."""
        def laeuft():
            try:
                self._nachrichten.put(("fertig", sid, arbeit(), None))
            except Exception as e:
                _spur()
                self._nachrichten.put(("fertig", sid, None, e))
        threading.Thread(target=laeuft, daemon=True).start()

    def _fehler(self, fehler, titel="Fehler"):
        self.status.set("Fehlgeschlagen.")
        self._schreibe(f"FEHLER: {fehler}")
        messagebox.showerror(titel, str(fehler))

    # -----------------------------------------------------------------
    # Laden / Speichern / Import
    # -----------------------------------------------------------------

    def _starte_laden(self):
        hinweis = core.fremde_arbeitssperre()
        if hinweis:
            messagebox.showwarning("Bestand wird bearbeitet", hinweis)
        self.status.set("Lade Bestand …")
        self._starte("laden", core.lade_bestand)

    def _laden_fertig(self, bestand, fehler):
        if fehler:
            return self._fehler(fehler, "Bestand konnte nicht geladen werden")
        self.bestand = bestand
        core.setze_arbeitssperre()
        for w in bestand.warnungen:
            self._schreibe("Hinweis beim Laden: " + w)
        if bestand.warnungen:
            self.reiter.select(3)
        self.status.set(
            f"{len(bestand.empfaenger)} Empfänger, "
            f"{sum(1 for _ in bestand.buecher())} Bücher geladen.")
        self._zeige_stammdaten()
        self._zeige_erfassung()

    def _speichern(self):
        self.status.set("Speichere …")
        self._starte("speichern",
                     lambda: core.speichere_bestand(self.bestand))

    def _speichern_fertig(self, pfad, fehler):
        if fehler:
            return self._fehler(fehler, "Bestand konnte nicht gespeichert werden")
        self.status.set(f"Gespeichert: {pfad}")
        self._schreibe(f"Bestand gespeichert: {pfad}")

    def _importieren(self):
        if self.bestand.empfaenger and not messagebox.askokcancel(
                "Altmappe importieren",
                "Der Import ersetzt den gesamten aktuellen Bestand.\n\n"
                "Fortfahren?"):
            return
        pfad = filedialog.askopenfilename(
            title="Alte Honorar-Mappe wählen",
            initialdir=self.cfg.get("last_input_dir") or str(core.APP_DIR),
            filetypes=[("Excel-Mappe", "*.xlsx"), ("Alle Dateien", "*.*")])
        if not pfad:
            return
        self.cfg["last_input_dir"] = str(Path(pfad).parent)
        core.speichere_config(self.cfg)
        jahr = self.jahr.get()
        self.reiter.select(3)
        self.status.set("Importiere …")
        self._starte("import",
                     lambda: core.importiere_alt(pfad, jahr, log=self._melde()))

    def _import_fertig(self, erg, fehler):
        if fehler:
            return self._fehler(fehler, "Import fehlgeschlagen")
        bestand, prot = erg
        self.bestand = bestand
        ziel = core.ausgabeordner(self.cfg, self.jahr.get()) / "import_protokoll.xlsx"
        try:
            core.schreibe_importprotokoll(prot, ziel)
            self._schreibe(f"Importprotokoll: {ziel}")
        except Exception as e:
            self._schreibe(f"Protokoll konnte nicht geschrieben werden: {e}")
        for w in prot.warnungen:
            self._schreibe("WARNUNG: " + " ".join(w.split()))
        self.status.set(
            f"Import fertig: {len(bestand.empfaenger)} Empfänger, "
            f"{sum(1 for _ in bestand.buecher())} Bücher, "
            f"{len(prot.warnungen)} Warnungen. Noch nicht gespeichert.")
        self._zeige_stammdaten()
        self._zeige_erfassung()
        messagebox.showinfo(
            "Import fertig",
            f"{len(bestand.empfaenger)} Empfänger und "
            f"{sum(1 for _ in bestand.buecher())} Bücher gelesen.\n\n"
            f"{len(prot.warnungen)} Stellen brauchen einen Blick — sie stehen "
            f"im Protokoll.\n\nDer Bestand ist noch NICHT gespeichert.")

    def _jahr_gewechselt(self):
        self._zeige_erfassung()
        self.abrechnungen = []
        for k in self.baum_durchlauf.get_children(""):
            self.baum_durchlauf.delete(k)
        self.summe.set("")

    # -----------------------------------------------------------------
    # Reiter 1: Stammdaten anzeigen
    # -----------------------------------------------------------------

    def _zeige_stammdaten(self):
        baum = self.baum_stamm
        for k in baum.get_children(""):
            baum.delete(k)
        suche = self.suche_stamm.get().strip().lower()
        for e in self.bestand.empfaenger:
            if suche and suche not in (
                    e.anzeigename + " " + e.ort + " " + e.iban).lower():
                continue
            offen = [b for b in e.buecher if b.nachpflege]
            hinweis = ""
            if not e.iban:
                hinweis = "keine Bankverbindung"
            elif offen:
                hinweis = f"{len(offen)} Buch/Bücher zur Nachpflege"
            baum.insert("", "end", iid=e.kennung, values=(
                e.anzeigename, f"{e.plz} {e.ort}".strip(), e.iban,
                len(e.buecher), hinweis))

    def _zeige_empfaenger(self, ereignis=None):
        auswahl = self.baum_stamm.selection()
        if not auswahl:
            return
        e = next((x for x in self.bestand.empfaenger
                  if x.kennung == auswahl[0]), None)
        if e is None:
            return
        fenster = Toplevel(self)
        fenster.title(e.anzeigename)
        fenster.geometry("900x620")
        kopf = ttk.LabelFrame(fenster, text="Anschrift")
        kopf.pack(fill="x", **PAD)
        text = "\n".join(core.anschrift_zeilen(e)) or "(keine Anschrift)"
        ttk.Label(kopf, text=text, justify="left").pack(anchor="w", **PAD)
        ttk.Label(kopf, text=f"IBAN: {e.iban or '—'}\n"
                             f"Aktenzeichen: {e.aktenzeichen or '—'}\n"
                             f"E-Mail: {e.email or '—'}",
                  justify="left").pack(anchor="w", **PAD)

        unten = ttk.LabelFrame(fenster, text="Bücher und Konditionen")
        unten.pack(fill="both", expand=True, **PAD)
        feld = Text(unten, wrap="word")
        feld.pack(fill="both", expand=True, **PAD)
        for b in e.buecher:
            feld.insert(END, f"{b.isbn or '—'}  {b.titel}\n")
            feld.insert(END, f"     {b.verguetungsart} · "
                             f"{core._kondition_text(b.kondition)}\n")
            if b.mwst_pflichtig:
                feld.insert(END, f"     mehrwertsteuerpflichtig "
                                 f"({b.mwst_satz:g} %)\n")
            if b.stillgelegt:
                feld.insert(END, f"     STILLGELEGT: {b.stillgelegt_grund}\n")
            for hinweis in b.nachpflege:
                feld.insert(END, f"     ! {hinweis}\n")
            if b.quelle:
                feld.insert(END, f"     Herkunft: {b.quelle}\n")
            feld.insert(END, "\n")
        feld.configure(state=DISABLED)

    # -----------------------------------------------------------------
    # Reiter 2: Jahreserfassung
    # -----------------------------------------------------------------

    def _erfassungszeilen(self):
        jahr = self.jahr.get()
        suche = self.suche_erfassung.get().strip().lower()
        for e in self.bestand.empfaenger:
            for b in e.buecher:
                if b.stillgelegt:
                    continue
                jw = b.jahre.get(jahr)
                if self.nur_offene.get() and jw is not None and jw.erfasst:
                    continue
                if suche and suche not in (
                        b.titel + " " + b.isbn + " " + e.anzeigename).lower():
                    continue
                yield e, b, jw

    def _zeige_erfassung(self):
        baum = self.baum_erfassung
        for k in baum.get_children(""):
            baum.delete(k)
        jahr = self.jahr.get()
        for e, b, jw in self._erfassungszeilen():
            stand = core.kumulierte_menge(b, jahr)
            baum.insert("", "end", iid=b.kennung, values=(
                b.isbn, b.titel, e.anzeigename, b.verguetungsart,
                "" if jw is None or jw.verkauft is None else jw.verkauft,
                "" if jw is None else jw.eigenkauf,
                "" if stand is None else stand))
        gesamt = sum(1 for e in self.bestand.empfaenger
                     for b in e.buecher if not b.stillgelegt)
        erfasst = sum(1 for e in self.bestand.empfaenger for b in e.buecher
                      if not b.stillgelegt and jahr in b.jahre
                      and b.jahre[jahr].erfasst)
        self.fortschritt.set(f"{erfasst} von {gesamt} Büchern erfasst")

    def _bearbeite_zelle(self, ereignis):
        """Stückzahl direkt im Baum ändern.

        Bei 280 Zeilen entscheidet der Tastenfluss über die Brauchbarkeit:
        Eingabetaste übernimmt und springt eine Zeile weiter, Esc verwirft.
        """
        baum = self.baum_erfassung
        if ereignis.type == "4":                       # Mausklick
            spalte = baum.identify_column(ereignis.x)
            zeile = baum.identify_row(ereignis.y)
        else:                                          # Eingabetaste
            auswahl = baum.selection()
            zeile = auswahl[0] if auswahl else ""
            spalte = "#5"
        if not zeile or spalte not in ("#5", "#6"):
            return
        feldname = "verkauft" if spalte == "#5" else "eigenkauf"
        x, y, breite, hoehe = baum.bbox(zeile, spalte)
        wert = StringVar(value=baum.set(zeile, feldname))
        eingabe = ttk.Entry(baum, textvariable=wert, justify="right")
        eingabe.place(x=x, y=y, width=breite, height=hoehe)
        eingabe.focus_set()
        eingabe.selection_range(0, END)

        def uebernehmen(weiter: bool):
            text = wert.get().strip()
            eingabe.destroy()
            buch = self._buch(zeile)
            if buch is None:
                return
            jw = buch.jahr(self.jahr.get())
            zahl = core._ganzzahl(text) if text else None
            if feldname == "verkauft":
                jw.verkauft = zahl
            else:
                jw.eigenkauf = zahl or 0
            baum.set(zeile, feldname, "" if zahl is None else zahl)
            self._zeige_fortschritt()
            if weiter:
                nachbar = baum.next(zeile)
                if nachbar:
                    baum.selection_set(nachbar)
                    baum.focus(nachbar)
                    baum.see(nachbar)

        eingabe.bind("<Return>", lambda e: uebernehmen(True))
        eingabe.bind("<FocusOut>", lambda e: uebernehmen(False))
        eingabe.bind("<Escape>", lambda e: eingabe.destroy())

    def _zeige_fortschritt(self):
        jahr = self.jahr.get()
        gesamt = sum(1 for e in self.bestand.empfaenger
                     for b in e.buecher if not b.stillgelegt)
        erfasst = sum(1 for e in self.bestand.empfaenger for b in e.buecher
                      if not b.stillgelegt and jahr in b.jahre
                      and b.jahre[jahr].erfasst)
        self.fortschritt.set(f"{erfasst} von {gesamt} Büchern erfasst")

    def _buch(self, kennung: str):
        for e in self.bestand.empfaenger:
            for b in e.buecher:
                if b.kennung == kennung:
                    return b
        return None

    # -----------------------------------------------------------------
    # Reiter 3: Durchlauf
    # -----------------------------------------------------------------

    def _rechnen(self):
        jahr = self.jahr.get()
        self.status.set("Rechne …")
        self._starte("rechnen",
                     lambda: core.rechne_alle(self.bestand, jahr, self.cfg))

    def _rechnen_fertig(self, abrechnungen, fehler):
        if fehler:
            return self._fehler(fehler, "Berechnung fehlgeschlagen")
        self.abrechnungen = abrechnungen
        baum = self.baum_durchlauf
        for k in baum.get_children(""):
            baum.delete(k)
        for ab in sorted(abrechnungen, key=lambda a: -a.brutto):
            hinweise = list(ab.gruende) + list(ab.probleme)
            marken = []
            if not ab.brief:
                marken.append("kein_brief")
            elif ab.probleme:
                marken.append("problem")
            baum.insert("", "end", iid=ab.empfaenger.kennung, values=(
                ab.empfaenger.anzeigename, ab.verguetungsart, len(ab.posten),
                core.euro(ab.netto), core.euro(ab.mwst), core.euro(ab.brutto),
                "Ja" if ab.brief else "—", " | ".join(hinweise)),
                tags=marken)
        mit = [a for a in abrechnungen if a.brief]
        self.summe.set(
            f"{len(mit)} Briefe · Auszahlung {core.euro(sum(a.brutto for a in mit))} "
            f"· KSK {core.euro(sum(a.ksk_netto for a in abrechnungen))}")
        self.status.set("Durchgerechnet.")

    def _gegenprobe(self):
        pfad = filedialog.askopenfilename(
            title="Alte Honorar-Mappe zum Vergleich wählen",
            initialdir=self.cfg.get("last_input_dir") or str(core.APP_DIR),
            filetypes=[("Excel-Mappe", "*.xlsx"), ("Alle Dateien", "*.*")])
        if not pfad:
            return
        jahr = self.jahr.get()
        self.reiter.select(3)
        self.status.set("Vergleiche mit der Altmappe …")
        self._starte("gegenprobe", lambda: core.pruefe_gegen_excel(
            self.bestand, pfad, jahr, self.cfg))

    def _gegenprobe_fertig(self, abweichungen, fehler):
        if fehler:
            return self._fehler(fehler, "Gegenprobe fehlgeschlagen")
        unerklaert = [a for a in abweichungen if not a.grund]
        self._schreibe(
            f"\nGegenprobe: {len(abweichungen)} Abweichungen, "
            f"davon {len(unerklaert)} ohne Erklärung.")
        for a in abweichungen:
            self._schreibe(
                f"  {a.was}: {a.wer or a.titel} — Altmappe {core.euro(a.soll)}, "
                f"Werkzeug {core.euro(a.ist)}"
                + (f" ({a.grund})" if a.grund else "  ← ungeklärt"))
        self.status.set(
            f"Gegenprobe: {len(unerklaert)} ungeklärte Abweichungen."
            if unerklaert else "Gegenprobe: alles stimmt überein.")

    # -----------------------------------------------------------------
    # Reiter 4: Ausgaben
    # -----------------------------------------------------------------

    def _bereit(self) -> bool:
        if not self.abrechnungen:
            messagebox.showinfo(
                "Erst durchrechnen",
                "Bitte zuerst im Reiter „Durchlauf“ auf „Jetzt durchrechnen“.")
            return False
        return True

    def _briefe(self):
        if not self._bereit():
            return
        ordner = core.ausgabeordner(self.cfg, self.jahr.get()) / "briefe"
        self.status.set("Erzeuge Briefe …")
        self._starte("briefe", lambda: core.erzeuge_briefe(
            self.abrechnungen, self.cfg, ordner, log=self._melde()))

    def _briefe_fertig(self, erg, fehler):
        if fehler:
            return self._fehler(fehler, "Briefe konnten nicht erzeugt werden")
        erzeugt, uebergangen = erg
        self._briefe_pfade = erzeugt
        self._schreibe(f"\n{len(erzeugt)} Briefe erzeugt.")
        if uebergangen:
            self._schreibe(f"Ohne Brief ({len(uebergangen)}):")
            for u in uebergangen:
                self._schreibe("   " + u)
        self.status.set(f"{len(erzeugt)} Briefe erzeugt.")

    def _pdf(self):
        pfade = getattr(self, "_briefe_pfade", [])
        if not pfade:
            messagebox.showinfo(
                "Erst Briefe erzeugen",
                "Die PDF-Umwandlung arbeitet mit den zuvor erzeugten "
                ".docx-Dateien. Bitte zuerst „Briefe (.docx)“.")
            return
        if sys.platform != "win32" and not messagebox.askokcancel(
                "Kein Windows",
                "Die Umwandlung braucht Word unter Windows.\n\n"
                "Hier wird der Schritt nur vermerkt und übersprungen. "
                "Trotzdem fortfahren?"):
            return
        self.status.set("Wandle nach PDF …")
        self._starte("pdf",
                     lambda: core.wandle_nach_pdf(pfade, log=self._melde()))

    def _pdf_fertig(self, erzeugt, fehler):
        if fehler:
            return self._fehler(fehler, "PDF-Umwandlung fehlgeschlagen")
        self.status.set(f"{len(erzeugt)} PDF-Dateien erzeugt.")

    def _listen(self):
        if not self._bereit():
            return
        jahr = self.jahr.get()
        ordner = core.ausgabeordner(self.cfg, jahr)

        def arbeit():
            a = core.schreibe_zahlungsliste(
                self.abrechnungen, jahr, ordner / f"Zahlungsliste_{jahr}.xlsx")
            b = core.schreibe_ksk_liste(
                self.abrechnungen, jahr,
                ordner / f"Kuenstlersozialkasse_{jahr}.xlsx", self.cfg)
            c = core.schreibe_laufprotokoll(
                self.abrechnungen, jahr, ordner / "protokoll.xlsx")
            return [a, b, c]

        self.status.set("Schreibe Listen …")
        self._starte("listen", arbeit)

    def _listen_fertig(self, pfade, fehler):
        if fehler:
            return self._fehler(fehler, "Listen konnten nicht geschrieben werden")
        for p in pfade:
            self._schreibe(f"geschrieben: {p}")
        self.status.set("Listen geschrieben.")

    def _ordner_oeffnen(self):
        self._datei_oeffnen(core.ausgabeordner(self.cfg, self.jahr.get()))

    @staticmethod
    def _datei_oeffnen(pfad):
        pfad = str(pfad)
        try:
            if sys.platform == "win32":
                os.startfile(pfad)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", pfad])
            else:
                subprocess.Popen(["xdg-open", pfad])
        except Exception as e:
            messagebox.showwarning("Öffnen fehlgeschlagen", str(e))

    def _beenden(self):
        core.loese_arbeitssperre()
        self.destroy()


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
