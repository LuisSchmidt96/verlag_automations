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

SPALTEN_ERFASSUNG = [("isbn", "ISBN", 90), ("titel", "Buchtitel", 300),
                     ("empf", "Empfänger", 200), ("art", "Vergütungsart", 120),
                     ("verkauft", "verkaufte Ex. ✎", 110),
                     ("eigenkauf", "Eigenkauf ✎", 100),
                     ("stand", "Stand bis Vorjahr", 120),
                     # Ohne diese Spalte tippt man Zahlen für ein Buch ein,
                     # das gar nicht ausgezahlt wird, und erfährt es nie.
                     ("bes", "Besonderheit", 230)]

SPALTEN_DURCHLAUF = [("name", "Empfänger", 280), ("art", "Vergütungsart", 120),
                     ("posten", "Bücher", 70), ("netto", "Netto", 95),
                     ("mwst", "MwSt", 85), ("brutto", "Auszahlung", 105),
                     ("brief", "Brief", 60), ("hinweis", "Grund / Hinweis", 340)]

# Im Sonderfall-Modus zeigt dieselbe Liste etwas anderes: nicht Empfänger mit
# ihrer Auszahlung, sondern die einzelnen Bücher mit dem Betrag, der wegen
# der offenen Entscheidung NICHT fließt. Sonst stünden dort neunzehn Zeilen
# mit lauter Nullen — die Empfänger haben ja gerade keine normale Auszahlung.
SPALTEN_SONDERFAELLE = [("name", "Empfänger", 260), ("titel", "Buchtitel", 300),
                        ("isbn", "ISBN", 80), ("ex", "Vergütungs-Ex.", 110),
                        ("betrag", "wird NICHT ausgezahlt", 160),
                        ("iban", "Bankverbindung", 210),
                        ("quelle", "in der Altmappe", 170)]


# Die vier Schritte in der Reihenfolge, in der gearbeitet wird. Die
# Beschriftung sagt, was zu TUN ist, nicht wie der Programmteil heißt.
REITER_MELDUNGEN = "Meldungen"

SCHRITTE = [
    ("stamm", "1. Autoren und Bücher"),
    ("zahlen", "2. Zahlen eintragen"),
    ("rechnen", "3. Nachrechnen"),
    ("ausgeben", "4. Briefe und Listen"),
]


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
        # Ob seit dem letzten Speichern etwas geändert wurde. Silke arbeitet
        # einmal im Jahr damit und tippt Hunderte Zahlen ein — ohne diesen
        # Merker wäre ein versehentliches Schließen der ganze Vormittag.
        self.geaendert = False
        # Welche Arbeitsschritte gerade laufen — gegen Doppelklicks.
        self._laufend: set[str] = set()
        self._rechnung_veraltet = False
        # Nach so vielen eingetippten Zahlen wird still zwischengespeichert.
        self._seit_sicherung = 0

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
        ttk.Label(kopf, text="Abrechnung für das Jahr:").pack(
            side="left", padx=(24, 4))
        ttk.Spinbox(kopf, from_=2000, to=2100, width=6, state="readonly",
                    textvariable=self.jahr,
                    command=self._jahr_gewechselt).pack(side="left")
        ttk.Button(kopf, text="Gespeicherten Stand neu laden",
                   command=self._starte_laden).pack(side="right")

        # Eine eigene Schrittleiste hätte die Reiterbeschriftungen nur Wort
        # für Wort wiederholt. Der Stand gehört deshalb IN die Reiter
        # (✓ / ◐ / ·), und diese Zeile sagt stattdessen etwas, das dort
        # nicht steht: was als Nächstes zu tun ist.
        self.naechstes = StringVar(value="")
        ttk.Label(self, textvariable=self.naechstes,
                  font=("Segoe UI", 10, "bold"), foreground="#1a5c1a"
                  ).pack(anchor="w", padx=12, pady=(0, 4))

        self.reiter = ttk.Notebook(self)
        self.reiter.pack(fill="both", expand=True, padx=12, pady=(0, 6))
        self._baue_stammdaten()
        self._baue_erfassung()
        self._baue_durchlauf()
        self._baue_ausgaben()
        self._baue_meldungen()

        leiste = ttk.Frame(self)
        leiste.pack(fill="x", padx=12, pady=(0, 10))
        ttk.Label(leiste, textvariable=self.status).pack(side="left")
        self._male_schrittleiste()

    def _male_schrittleiste(self):
        """Stand in die Reiterbeschriftungen schreiben und sagen, was folgt.

        ✓ heißt fertig, ◐ heißt angefangen, · heißt noch gar nichts.
        """
        jahr = self.jahr.get()
        gesamt = sum(1 for e in self.bestand.empfaenger
                     for b in e.buecher if not b.stillgelegt)
        offen = sum(1 for e in self.bestand.empfaenger for b in e.buecher
                    if not b.stillgelegt
                    and not (jahr in b.jahre and b.jahre[jahr].erfasst))
        veraltet = getattr(self, "_rechnung_veraltet", False)
        zeichen = [
            "✓" if self.bestand.empfaenger else "·",
            ("·" if not gesamt else "✓" if offen == 0
             else "◐" if offen < gesamt else "·"),
            "✓" if self.abrechnungen and not veraltet else
            ("◐" if self.abrechnungen else "·"),
            "✓" if getattr(self, "_etwas_erzeugt", False) else "·",
        ]
        for i, (_, titel) in enumerate(SCHRITTE):
            self.reiter.tab(i, text=f"{zeichen[i]} {titel}")

        # Genau eine Handlung, nie eine Liste — sonst ist es wieder geraten.
        if not self.bestand.empfaenger:
            satz = ("Als Nächstes: auf Reiter 1 die Daten aus der alten "
                    "Excel-Tabelle holen.")
        elif offen:
            satz = (f"Als Nächstes: auf Reiter 2 noch {offen} Bücher "
                    f"eintragen.")
        elif not self.abrechnungen or veraltet:
            satz = ("Als Nächstes: auf Reiter 3 die Beträge berechnen.")
        elif not getattr(self, "_etwas_erzeugt", False):
            satz = ("Als Nächstes: das Ergebnis auf Reiter 3 durchsehen, "
                    "dann auf Reiter 4 die Briefe erzeugen.")
        elif self.geaendert:
            satz = ("Als Nächstes: die Eingaben speichern — Reiter 1, "
                    "Knopf oben rechts.")
        else:
            satz = "Fertig. Alles erzeugt und gespeichert."
        self.naechstes.set(satz)
        self._zaehle_meldungen()

    def _formular(self, titel: str, felder: list, werte: dict) -> dict | None:
        """Ein einfaches Eingabefenster: je Zeile Beschriftung und Feld.

        ``felder`` ist eine Liste von (Schlüssel, Beschriftung, Art), wobei
        die Art „text“, „zahl“ oder eine Liste zur Auswahl ist. Gibt die
        geänderten Werte zurück oder None, wenn abgebrochen wurde.
        """
        fenster = Toplevel(self)
        fenster.title(titel)
        fenster.transient(self)
        fenster.grab_set()
        rahmen = ttk.Frame(fenster)
        rahmen.pack(fill="both", expand=True, **PAD)

        eingaben = {}
        for i, (schluessel, beschriftung, art) in enumerate(felder):
            ttk.Label(rahmen, text=beschriftung).grid(
                row=i, column=0, sticky="e", padx=(0, 8), pady=3)
            var = StringVar(value=str(werte.get(schluessel, "") or ""))
            if isinstance(art, (list, tuple)):
                feld = ttk.Combobox(rahmen, textvariable=var, width=46,
                                    values=list(art), state="readonly")
            else:
                feld = ttk.Entry(rahmen, textvariable=var, width=48)
            feld.grid(row=i, column=1, sticky="w", pady=3)
            eingaben[schluessel] = (var, art)
            if i == 0:
                feld.focus_set()

        ergebnis = {}
        fertig = {"ok": False}

        def uebernehmen():
            for schluessel, (var, art) in eingaben.items():
                text = var.get().strip()
                if art == "zahl" and text:
                    zahl = core._komma(text)
                    if zahl is None:
                        messagebox.showwarning(
                            "Keine Zahl",
                            f"Bei „{dict((s, b) for s, b, _ in felder)[schluessel]}“ "
                            f"steht „{text}“ — dort gehört eine Zahl hin.",
                            parent=fenster)
                        return
                    ergebnis[schluessel] = zahl
                else:
                    ergebnis[schluessel] = text
            fertig["ok"] = True
            fenster.destroy()

        leiste = ttk.Frame(fenster)
        leiste.pack(fill="x", padx=12, pady=(0, 12))
        ttk.Button(leiste, text="Übernehmen", command=uebernehmen).pack(side="right")
        ttk.Button(leiste, text="Abbrechen",
                   command=fenster.destroy).pack(side="right", padx=6)
        fenster.bind("<Return>", lambda e: uebernehmen())
        fenster.bind("<Escape>", lambda e: fenster.destroy())
        self.wait_window(fenster)
        return ergebnis if fertig["ok"] else None

    FELDER_EMPFAENGER = [
        ("anrede", "Anrede", ["", "Herr", "Frau", "Damen und Herren",
                              "Frau und Herr"]),
        ("titel_akad", "Titel (Dr., Prof.)", "text"),
        ("vorname", "Vorname", "text"),
        ("name", "Nachname", "text"),
        ("institution", "Einrichtung / Firma", "text"),
        ("strasse", "Straße und Hausnummer", "text"),
        ("plz", "PLZ", "text"),
        ("ort", "Ort", "text"),
        ("land", "Land (nur wenn im Ausland)", "text"),
        ("iban", "Bankverbindung (IBAN)", "text"),
        ("aktenzeichen", "Aktenzeichen / Kontoinhaber", "text"),
        ("email", "E-Mail", "text"),
    ]

    def _neuer_empfaenger(self):
        werte = self._formular("Neuen Autor anlegen",
                               self.FELDER_EMPFAENGER, {})
        if werte is None:
            return
        if not (werte["name"] or werte["institution"]):
            messagebox.showwarning(
                "Name fehlt",
                "Bitte mindestens einen Nachnamen oder eine Einrichtung "
                "eintragen — sonst lässt sich der Brief nicht adressieren.")
            return
        e = core.Empfaenger(kennung=self.bestand.naechste_kennung("E"),
                            beruehrt=True, **werte)
        self.bestand.empfaenger.append(e)
        self.geaendert = True
        self._zeige_stammdaten()
        self._male_schrittleiste()
        self.baum_stamm.selection_set(e.kennung)
        self.baum_stamm.see(e.kennung)
        self.status.set(f"„{e.anzeigename}“ angelegt — noch nicht gespeichert.")

    def _aendere_empfaenger(self, e):
        werte = self._formular(f"Angaben ändern — {e.anzeigename}",
                               self.FELDER_EMPFAENGER,
                               {k: getattr(e, k) for k, _, _ in
                                self.FELDER_EMPFAENGER})
        if werte is None:
            return
        for schluessel, wert in werte.items():
            setattr(e, schluessel, wert)
        e.beruehrt = True
        self.geaendert = True
        self._zeige_stammdaten()
        self._zeige_erfassung()
        self.status.set("Geändert — noch nicht gespeichert.")

    # --- Buchdialog ---------------------------------------------------
    # Eigenes Fenster statt eines Feldrasters: was der Autor je Exemplar
    # bekommt, ist keine Liste von vierzehn Feldern, sondern eine
    # Entscheidung mit zwei Wegen — fester Betrag oder Anteil am
    # Verlagsabgabepreis. Und die Staffel ist kein eigenes Feld, sondern der
    # Honorarsatz selbst, nur mengenabhängig. Frei eingetippter Vertragstext,
    # den das Werkzeug deuten muss, hat hier nichts zu suchen: beim Import
    # war Deuten alternativlos, beim Eingeben ist es die schlechteste
    # Loesung.

    def _neues_buch(self, e):
        buch = core.Buch(kennung=self.bestand.naechste_kennung("B"))
        if not self._buchfenster(f"Neues Buch für {e.anzeigename}", buch):
            return
        e.buecher.append(buch)
        self.geaendert = True
        self._rechnung_veraltet = bool(self.abrechnungen)
        self._zeige_stammdaten()
        self._zeige_erfassung()
        self._male_schrittleiste()
        self.status.set(f"„{buch.titel}“ angelegt — noch nicht gespeichert.")

    def _aendere_buch(self, e, buch):
        """Ein bestehendes Buch ändern — sonst waere ein Vertippen nur durch
        Neuanlegen zu beheben, und Loeschen gibt es nicht."""
        if not self._buchfenster(f"Buch ändern — {buch.titel}", buch):
            return
        self.geaendert = True
        self._rechnung_veraltet = bool(self.abrechnungen)
        self._zeige_stammdaten()
        self._zeige_erfassung()
        self.status.set(f"„{buch.titel}“ geändert — noch nicht gespeichert.")

    def _buchfenster(self, titel: str, buch) -> bool:
        """Buchangaben erfassen. True, wenn übernommen wurde."""
        k = buch.kondition
        fenster = Toplevel(self)
        fenster.title(titel)
        fenster.transient(self)
        fenster.grab_set()

        def feld(eltern, zeile, beschriftung, var, breite=26, werte=None):
            ttk.Label(eltern, text=beschriftung).grid(
                row=zeile, column=0, sticky="e", padx=(0, 8), pady=3)
            if werte:
                w = ttk.Combobox(eltern, textvariable=var, width=breite - 2,
                                 values=werte, state="readonly")
            else:
                w = ttk.Entry(eltern, textvariable=var, width=breite)
            w.grid(row=zeile, column=1, sticky="w", pady=3)
            return w

        # --- Das Buch ---------------------------------------------------
        oben = ttk.LabelFrame(fenster, text="Das Buch")
        oben.pack(fill="x", **PAD)
        v_titel = StringVar(value=buch.titel)
        v_isbn = StringVar(value=buch.isbn)
        v_art = StringVar(value=buch.verguetungsart or "Honorar")
        v_notiz = StringVar(value=buch.notizen)
        feld(oben, 0, "Buchtitel", v_titel, 46).focus_set()
        feld(oben, 1, "ISBN (Kurzform, z. B. 05-300)", v_isbn, 20)
        feld(oben, 2, "Art der Zahlung", v_art, 24,
             werte=list(core.VERGUETUNGSARTEN))
        feld(oben, 3, "Notizen", v_notiz, 46)

        # --- Was je Exemplar gezahlt wird -------------------------------
        geld = ttk.LabelFrame(fenster, text="Was der Autor je Exemplar bekommt")
        geld.pack(fill="x", **PAD)
        v_weg = StringVar(value="fest" if k.betrag_je_ex is not None
                          else "anteil")

        fest = ttk.Frame(geld)
        ttk.Radiobutton(geld, text="Ein fester Betrag je Exemplar",
                        variable=v_weg, value="fest",
                        command=lambda: umschalten()).grid(
                            row=0, column=0, sticky="w", padx=8, pady=(6, 0))
        fest.grid(row=1, column=0, sticky="w", padx=32)
        v_fest = StringVar(value="" if k.betrag_je_ex is None
                           else f"{k.betrag_je_ex:g}")
        feld(fest, 0, "Betrag in €", v_fest, 12)

        ttk.Radiobutton(geld, text="Ein Anteil am Verlagsabgabepreis",
                        variable=v_weg, value="anteil",
                        command=lambda: umschalten()).grid(
                            row=2, column=0, sticky="w", padx=8, pady=(10, 0))
        anteil = ttk.Frame(geld)
        anteil.grid(row=3, column=0, sticky="w", padx=32)
        v_preis = StringVar(value="" if k.ladenpreis is None
                            else f"{k.ladenpreis:g}")
        v_mwst = StringVar(value="keine" if not k.mwst_im_preis
                           else f"{k.mwst_im_preis:g} %")
        v_rabatt = StringVar(value=f"{k.verlagsrabatt:g}"
                             if k.rabatt_anwenden else "")
        v_teiler = StringVar(value=str(k.teiler or 1))
        feld(anteil, 0, "Ladenpreis in €", v_preis, 12)
        feld(anteil, 1, "darin enthaltene MwSt", v_mwst, 12,
             werte=["keine", "7 %", "19 %"])
        feld(anteil, 2, "Verlagsrabatt in % (meist 40)", v_rabatt, 12)
        feld(anteil, 3, "geteilt durch (Mitautoren)", v_teiler, 12)

        # --- Der Honorarsatz, gleich ob fest oder gestaffelt ------------
        satzrahmen = ttk.LabelFrame(
            fenster, text="Honorarsatz — gleichbleibend oder nach Menge gestaffelt")
        satzrahmen.pack(fill="x", **PAD)
        v_satzart = StringVar(value="staffel" if k.staffel else "fest")
        ttk.Radiobutton(satzrahmen, text="Immer derselbe Satz",
                        variable=v_satzart, value="fest",
                        command=lambda: umschalten()).grid(
                            row=0, column=0, sticky="w", padx=8, pady=(6, 0))
        einfach = ttk.Frame(satzrahmen)
        einfach.grid(row=1, column=0, sticky="w", padx=32)
        v_satz = StringVar(value="" if k.satz is None else f"{k.satz:g}")
        feld(einfach, 0, "Satz in %", v_satz, 10)

        ttk.Radiobutton(satzrahmen,
                        text="Gestaffelt — der Satz steigt mit der Menge",
                        variable=v_satzart, value="staffel",
                        command=lambda: umschalten()).grid(
                            row=2, column=0, sticky="w", padx=8, pady=(10, 0))
        stufen = ttk.Frame(satzrahmen)
        stufen.grid(row=3, column=0, sticky="w", padx=32)
        ttk.Label(stufen, text="bis … Exemplare").grid(row=0, column=1)
        ttk.Label(stufen, text="Satz in %").grid(row=0, column=2)

        stufenzeilen: list[tuple] = []

        def stufe_anlegen(grenze="", satz=""):
            i = len(stufenzeilen)
            if i >= 6:
                return
            ttk.Label(stufen, text=f"{i + 1}.").grid(row=i + 1, column=0,
                                                     padx=(0, 6))
            vg, vs = StringVar(value=grenze), StringVar(value=satz)
            eg = ttk.Entry(stufen, textvariable=vg, width=12, justify="right")
            es = ttk.Entry(stufen, textvariable=vs, width=8, justify="right")
            eg.grid(row=i + 1, column=1, pady=2)
            es.grid(row=i + 1, column=2, padx=(8, 0), pady=2)
            for e_ in (eg, es):
                e_.bind("<KeyRelease>", lambda _e: vorschau())
            stufenzeilen.append((vg, vs))

        vorhandene = list(k.staffel) or [(None, None)]
        for grenze, satz in vorhandene:
            stufe_anlegen("" if grenze is None else str(grenze),
                          "" if satz is None else f"{satz:g}")
        while len(stufenzeilen) < 2:
            stufe_anlegen()
        ttk.Button(stufen, text="+ Stufe",
                   command=lambda: (stufe_anlegen(), vorschau())).grid(
                       row=99, column=1, sticky="w", pady=(6, 0))
        ttk.Label(satzrahmen, justify="left", foreground="#555555", text=(
            "Die letzte Stufe ohne Mengenangabe gilt nach oben offen. Welche "
            "Stufe in einem Jahr greift, entscheidet der Stand zu\n"
            "Jahresbeginn (Spalte „Stand bis Vorjahr“).")
        ).grid(row=4, column=0, sticky="w", padx=8, pady=(4, 8))

        # --- Sonderregeln ------------------------------------------------
        sonder = ttk.LabelFrame(fenster, text="Sonderregeln")
        sonder.pack(fill="x", **PAD)
        v_mwstpflicht = BooleanVar(value=buch.mwst_pflichtig)
        v_schwelle = BooleanVar(value=k.schwelle_zehn)
        v_frei = StringVar(value=str(k.freimenge) if k.freimenge else "")
        ttk.Checkbutton(sonder, text="Autor ist mehrwertsteuerpflichtig",
                        variable=v_mwstpflicht).grid(row=0, column=0,
                                                     sticky="w", padx=8, pady=2)
        ttk.Checkbutton(sonder,
                        text="Kein Honorar unter zehn Exemplaren im Jahr",
                        variable=v_schwelle).grid(row=1, column=0, sticky="w",
                                                  padx=8, pady=2)
        frei = ttk.Frame(sonder)
        frei.grid(row=2, column=0, sticky="w", padx=8, pady=2)
        feld(frei, 0, "Freimenge: die ersten … Exemplare ohne Honorar",
             v_frei, 10)

        # --- Vorschau ----------------------------------------------------
        ergebnis = StringVar()
        ttk.Label(fenster, textvariable=ergebnis, font=("Segoe UI", 10, "bold"),
                  foreground="#1a5c1a", wraplength=620, justify="left"
                  ).pack(anchor="w", padx=12, pady=(0, 4))

        def lies() -> tuple:
            """Die Eingaben in eine Kondition übersetzen. (Kondition, Fehler)"""
            kond = core.Kondition()
            fehler = []
            if v_weg.get() == "fest":
                betrag = core._komma(v_fest.get())
                if betrag is None:
                    fehler.append("Es fehlt der Betrag je Exemplar.")
                kond.betrag_je_ex = betrag
            else:
                preis = core._komma(v_preis.get())
                if preis is None:
                    fehler.append("Es fehlt der Ladenpreis.")
                kond.ladenpreis = preis
                kond.mwst_im_preis = ({"7 %": 7.0, "19 %": 19.0}
                                      .get(v_mwst.get()))
                rabatt = core._komma(v_rabatt.get())
                kond.rabatt_anwenden = rabatt is not None
                kond.verlagsrabatt = rabatt if rabatt is not None else 40.0
                kond.teiler = int(core._komma(v_teiler.get(), 1) or 1)
                if v_satzart.get() == "fest":
                    satz = core._komma(v_satz.get())
                    if satz is None:
                        fehler.append("Es fehlt der Honorarsatz.")
                    kond.satz = satz
                else:
                    staffel, letzte = [], None
                    for vg, vs in stufenzeilen:
                        satz = core._komma(vs.get())
                        if satz is None:
                            continue
                        grenze = core._ganzzahl(vg.get())
                        staffel.append((grenze, satz))
                    offen = [s for s in staffel if s[0] is None]
                    begrenzt = sorted((s for s in staffel if s[0] is not None),
                                      key=lambda s: s[0])
                    if len(offen) > 1:
                        fehler.append("Nur EINE Stufe darf ohne Mengenangabe "
                                      "bleiben — sie gilt nach oben offen.")
                    if not staffel:
                        fehler.append("Es ist keine Stufe ausgefüllt.")
                    elif not offen:
                        fehler.append("Die oberste Stufe braucht keine "
                                      "Mengenangabe — sie gilt für alles "
                                      "darüber. Bitte dort die Menge leeren.")
                    grenzen = [s[0] for s in begrenzt]
                    if len(set(grenzen)) != len(grenzen):
                        fehler.append("Zwei Stufen haben dieselbe Menge.")
                    kond.staffel = begrenzt + offen[:1]
                    # Welche Stufe gilt, entscheidet der kumulierte Stand.
                    # Ist der unbekannt — bei den meisten Büchern der Fall —,
                    # greift `satz` als Rückfall. Den darf das Formular NICHT
                    # überschreiben: bei einem importierten Buch steht dort
                    # der Satz, den der Verlag tatsächlich angewandt hat.
                    # Auf die erste Stufe zu setzen machte aus 14 % 12 % und
                    # damit aus 1,33 € 1,14 €.
                    kond.satz = (k.satz if k.satz is not None
                                 else (begrenzt or offen)[0][1]
                                 if (begrenzt or offen) else None)
            kond.schwelle_zehn = v_schwelle.get()
            kond.freimenge = int(core._komma(v_frei.get(), 0) or 0)
            return kond, fehler

        def vorschau(*_):
            kond, fehler = lies()
            if fehler:
                ergebnis.set("… " + fehler[0])
                return
            if kond.staffel:
                teile = []
                untere = 1
                for grenze, satz in kond.staffel:
                    kopie = core.Kondition(**{**kond.__dict__, "satz": satz,
                                              "staffel": []})
                    betrag = core.satz_aus_kondition(kopie)
                    if grenze is None:
                        teile.append(f"ab {untere}: {core.euro(betrag)}")
                    else:
                        teile.append(f"{untere}–{grenze}: {core.euro(betrag)}")
                        untere = grenze + 1
                ergebnis.set("Ergibt je Exemplar — " + ", ".join(teile))
            else:
                ergebnis.set("Ergibt "
                             + core.euro(core.satz_aus_kondition(kond))
                             + " je Exemplar.")

        def umschalten():
            anteilig = v_weg.get() == "anteil"
            for w in fest.winfo_children():
                w.configure(state="normal" if not anteilig else "disabled")
            for rahmen in (anteil, einfach, stufen):
                for w in rahmen.winfo_children():
                    if w.winfo_class() in ("TEntry", "TCombobox", "TButton"):
                        w.configure(state="disabled" if not anteilig else
                                    ("readonly" if w.winfo_class() == "TCombobox"
                                     else "normal"))
            if anteilig:
                gestaffelt = v_satzart.get() == "staffel"
                for w in einfach.winfo_children():
                    if w.winfo_class() in ("TEntry",):
                        w.configure(state="disabled" if gestaffelt else "normal")
                for w in stufen.winfo_children():
                    if w.winfo_class() in ("TEntry", "TButton"):
                        w.configure(state="normal" if gestaffelt else "disabled")
            vorschau()

        for var in (v_fest, v_preis, v_rabatt, v_satz, v_teiler, v_frei):
            var.trace_add("write", lambda *_: vorschau())
        v_mwst.trace_add("write", lambda *_: vorschau())
        umschalten()

        # --- Knöpfe ------------------------------------------------------
        fertig = {"ok": False}

        def uebernehmen():
            if not v_titel.get().strip():
                messagebox.showwarning("Titel fehlt",
                                       "Ohne Buchtitel geht es nicht.",
                                       parent=fenster)
                return
            kond, fehler = lies()
            if fehler:
                messagebox.showwarning("Bitte noch ergänzen",
                                       "\n".join(fehler), parent=fenster)
                return
            buch.titel = v_titel.get().strip()
            buch.isbn = v_isbn.get().strip()
            buch.verguetungsart = v_art.get() or "Honorar"
            buch.notizen = v_notiz.get().strip()
            buch.mwst_pflichtig = v_mwstpflicht.get()
            buch.kondition = kond
            fertig["ok"] = True
            fenster.destroy()

        leiste = ttk.Frame(fenster)
        leiste.pack(fill="x", padx=12, pady=(0, 12))
        ttk.Button(leiste, text="Übernehmen",
                   command=uebernehmen).pack(side="right")
        ttk.Button(leiste, text="Abbrechen",
                   command=fenster.destroy).pack(side="right", padx=6)
        fenster.bind("<Escape>", lambda _e: fenster.destroy())
        self.wait_window(fenster)
        return fertig["ok"]

    def _anleitung(self, eltern, text: str):
        """Ein Satz oben auf jedem Reiter: was ist hier zu tun.

        Silke benutzt das Werkzeug einmal im Jahr und liest keine Anleitung.
        Was sie wissen muss, muss dort stehen, wo sie gerade hinsieht.
        """
        ttk.Label(eltern, text=text, justify="left", wraplength=1300,
                  font=("Segoe UI", 10)).pack(anchor="w", padx=12, pady=(10, 0))

    def _weiter(self, eltern, nummer: int, text: str):
        """Der Knopf, der zum nächsten Schritt führt."""
        rahmen = ttk.Frame(eltern)
        rahmen.pack(fill="x", padx=12, pady=(0, 10))
        ttk.Button(rahmen, text=text,
                   command=lambda: self.reiter.select(nummer)).pack(side="right")

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
        self.reiter.add(seite, text="1. Autoren und Bücher")
        self._anleitung(seite, (
            "Hier stehen alle Autoren mit ihren Büchern, Anschriften und "
            "Bankverbindungen. Beim allerersten Mal holen Sie die Daten mit "
            "dem Knopf rechts aus der alten Excel-Tabelle. Danach müssen Sie "
            "hier nur noch etwas tun, wenn ein Autor umzieht oder ein Buch "
            "dazukommt."))

        oben = ttk.Frame(seite)
        oben.pack(fill="x", **PAD)
        ttk.Label(oben, text="Suchen:").pack(side="left")
        feld = ttk.Entry(oben, textvariable=self.suche_stamm, width=36)
        feld.pack(side="left", padx=6)
        feld.bind("<KeyRelease>", lambda e: self._zeige_stammdaten())
        ttk.Button(oben, text="Speichern",
                   command=self._speichern).pack(side="right")
        ttk.Button(oben, text="Daten aus der alten Excel-Tabelle holen …",
                   command=self._importieren).pack(side="right", padx=6)
        ttk.Button(oben, text="Neuer Autor …",
                   command=self._neuer_empfaenger).pack(side="left", padx=(16, 4))
        ttk.Button(oben, text="Neues Buch …",
                   command=self._neues_buch_zur_auswahl).pack(side="left")

        self.baum_stamm = self._liste(seite, SPALTEN_STAMM,
                                      doppelklick=self._zeige_empfaenger)
        ttk.Label(seite, text="Doppelklick auf eine Zeile zeigt Anschrift, "
                              "Bücher und die vereinbarte Vergütung."
                  ).pack(anchor="w", padx=12)
        self._weiter(seite, 1, "Weiter zu Schritt 2: Zahlen eintragen  ▸")

    # --- Reiter 2: Jahreserfassung ------------------------------------

    def _baue_erfassung(self):
        seite = ttk.Frame(self.reiter)
        self.reiter.add(seite, text="2. Zahlen eintragen")
        self._anleitung(seite, (
            "Das ist die eigentliche Arbeit: Tragen Sie für jedes Buch ein, "
            "wie viele Exemplare im Abrechnungsjahr verkauft wurden und wie "
            "viele davon der Autor selbst gekauft hat."))

        oben = ttk.Frame(seite)
        oben.pack(fill="x", **PAD)
        ttk.Label(oben, text="Suchen:").pack(side="left")
        feld = ttk.Entry(oben, textvariable=self.suche_erfassung, width=36)
        feld.pack(side="left", padx=6)
        feld.bind("<KeyRelease>", lambda e: self._zeige_erfassung())
        ttk.Checkbutton(oben, text="nur Bücher ohne Zahl",
                        variable=self.nur_offene,
                        command=self._zeige_erfassung).pack(side="left", padx=12)
        ttk.Button(oben, text="Zur ersten Zeile ohne Zahl",
                   command=self._springe_zur_luecke).pack(side="left", padx=6)
        self.fortschritt = StringVar(value="")
        ttk.Label(oben, textvariable=self.fortschritt,
                  font=("Segoe UI", 10, "bold")).pack(side="right")

        self.baum_erfassung = self._liste(seite, SPALTEN_ERFASSUNG)
        self.baum_erfassung.tag_configure("gesondert", foreground="#8a5a00")
        self.baum_erfassung.bind("<Double-1>", self._bearbeite_zelle)
        self.baum_erfassung.bind("<Return>", self._bearbeite_zelle)
        # Der Unterschied zwischen einer leeren Zelle und einer 0 entscheidet
        # über Geld. Er muss dastehen, nicht in einem Fachwort stecken.
        hinweis = ttk.Frame(seite)
        hinweis.pack(fill="x", padx=12, pady=(0, 8))
        ttk.Label(hinweis, justify="left", text=(
            "Beschreibbar sind nur die beiden Spalten mit dem Stift (✎). "
            "Doppelklick öffnet das Feld, die Eingabetaste übernimmt und "
            "springt eine Zeile tiefer, Tabulator wechselt zum Eigenkauf, "
            "Esc verwirft.")
        ).pack(anchor="w")
        ttk.Label(hinweis, justify="left", foreground="#8a5a00", text=(
            "Wichtig: Ein leeres Feld heißt „habe ich noch nicht eingetragen“. "
            "Hat sich ein Buch im ganzen Jahr nicht verkauft, tragen Sie dort "
            "bitte eine 0 ein — sonst fehlt das Buch in der Abrechnung.")
        ).pack(anchor="w")
        ttk.Label(hinweis, justify="left", foreground="#8a5a00", text=(
            "Braune Zeilen werden gesondert abgerechnet: ihre Zahlen laufen "
            "für Staffel und Freimenge mit, lösen aber keine Auszahlung aus. "
            "Was an einem Buch sonst besonders ist, steht rechts.")
        ).pack(anchor="w")
        self._weiter(seite, 2, "Weiter zu Schritt 3: Nachrechnen  ▸")

    # --- Reiter 3: Durchlauf ------------------------------------------

    def _baue_durchlauf(self):
        seite = ttk.Frame(self.reiter)
        self.reiter.add(seite, text="3. Nachrechnen")
        self._anleitung(seite, (
            "Hier sehen Sie, was jeder Autor bekommt, bevor ein einziger "
            "Brief geschrieben wird. Grau heißt: bekommt keinen Brief, weil "
            "nichts auszuzahlen ist. Rot heißt: bitte einmal ansehen."))

        oben = ttk.Frame(seite)
        oben.pack(fill="x", **PAD)
        ttk.Button(oben, text="Beträge berechnen",
                   command=self._rechnen).pack(side="left")
        ttk.Button(oben, text="Mit der alten Excel-Tabelle vergleichen …",
                   command=self._gegenprobe).pack(side="left", padx=8)
        self.nur_sonderfaelle = BooleanVar(value=False)
        ttk.Checkbutton(oben, text="nur die offenen Sonderfälle",
                        variable=self.nur_sonderfaelle,
                        command=self._male_durchlauf).pack(side="left", padx=12)
        self.summe = StringVar(value="")
        ttk.Label(oben, textvariable=self.summe,
                  font=("Segoe UI", 10, "bold")).pack(side="right")

        self.baum_durchlauf = self._liste(seite, SPALTEN_DURCHLAUF,
                                          doppelklick=self._zeige_rechenweg)
        self.baum_durchlauf.tag_configure("kein_brief", foreground="#777777")
        self.baum_durchlauf.tag_configure("problem", foreground="#a4262c")
        self.baum_durchlauf.tag_configure("sonderfall", foreground="#8a5a00")
        self.zurueck = StringVar(value="")
        ttk.Label(seite, textvariable=self.zurueck, foreground="#8a5a00",
                  wraplength=1300, justify="left").pack(anchor="w", padx=12)
        # Legende: Farben ohne Erklärung sind keine Auskunft.
        legende = ttk.Frame(seite)
        legende.pack(fill="x", padx=12)
        ttk.Label(legende, text="grau = bekommt keinen Brief",
                  foreground="#777777").pack(side="left", padx=(0, 20))
        ttk.Label(legende, text="rot = bitte ansehen, bevor der Brief hinausgeht",
                  foreground="#a4262c").pack(side="left", padx=(0, 20))
        ttk.Label(legende, text="braun = wartet auf eine Entscheidung",
                  foreground="#8a5a00").pack(side="left")
        ttk.Label(legende, text="  ·  Doppelklick auf eine Zeile zeigt, wie der "
                                "Betrag zustande kommt").pack(side="left")
        self._weiter(seite, 3, "Weiter zu Schritt 4: Briefe und Listen  ▸")

    # --- Reiter 4: Ausgaben -------------------------------------------

    def _baue_ausgaben(self):
        seite = ttk.Frame(self.reiter)
        self.reiter.add(seite, text="4. Briefe und Listen")
        self._anleitung(seite, (
            "Zum Schluss werden die Briefe und die beiden Listen geschrieben. "
            "Alles landet in einem Ordner, den Sie unten öffnen können. "
            "Autoren ohne Auszahlung bekommen keinen Brief — sie stehen aber "
            "mit Begründung im Protokoll."))

        knoepfe = ttk.LabelFrame(seite, text="Was soll erzeugt werden?")
        knoepfe.pack(fill="x", **PAD)
        self.btn_briefe = ttk.Button(knoepfe, text="Briefe schreiben (Word)",
                                     command=self._briefe)
        self.btn_briefe.grid(row=0, column=0, **PAD)
        self.btn_pdf = ttk.Button(knoepfe, text="Briefe als PDF (braucht Word)",
                                  command=self._pdf)
        self.btn_pdf.grid(row=0, column=1, **PAD)
        self.btn_listen = ttk.Button(knoepfe, text="Zahlungsliste und KSK-Meldung",
                                     command=self._listen)
        self.btn_listen.grid(row=0, column=2, **PAD)
        ttk.Button(knoepfe, text="Ordner mit den Ergebnissen öffnen",
                   command=self._ordner_oeffnen).grid(row=0, column=3, **PAD)

        self.ergebnis = StringVar(value="Noch nichts erzeugt.")
        ttk.Label(seite, textvariable=self.ergebnis, justify="left",
                  wraplength=1300).pack(anchor="w", padx=12, pady=(4, 0))
        ttk.Label(seite, justify="left", foreground="#555555", text=(
            "Einzelheiten zu jedem Schritt stehen im Reiter „Meldungen“.")
        ).pack(anchor="w", padx=12, pady=(2, 0))

    # --- Reiter 5: Meldungen ------------------------------------------

    def _baue_meldungen(self):
        """Ein eigener Reiter für alles, was das Werkzeug zu sagen hat.

        Vorher lag das Protokollfeld auf dem Ausgaben-Reiter — und weil das
        Werkzeug beim Import dorthin sprang, stand der Bediener beim ersten
        Kontakt vor den Knöpfen „Briefe schreiben“ und „PDF“, ausgerechnet
        während noch gar nichts gerechnet war.
        """
        seite = ttk.Frame(self.reiter)
        self.reiter.add(seite, text=REITER_MELDUNGEN)
        self._anleitung(seite, (
            "Hier schreibt das Werkzeug mit, was es tut und was ihm "
            "aufgefallen ist. Beim Einlesen der alten Excel-Tabelle stehen "
            "hier die Stellen, die jemand ansehen sollte."))
        self.protokollfeld = Text(seite, height=24, state=DISABLED, wrap="word")
        self.protokollfeld.pack(fill="both", expand=True, **PAD)
        unten = ttk.Frame(seite)
        unten.pack(fill="x", padx=12, pady=(0, 10))
        ttk.Button(unten, text="Meldungen leeren",
                   command=self._leere_meldungen).pack(side="right")

    def _leere_meldungen(self):
        feld = self.protokollfeld
        feld.configure(state=NORMAL)
        feld.delete("1.0", END)
        feld.configure(state=DISABLED)
        self._meldungen = 0
        self._zaehle_meldungen()

    def _zeige_meldungen(self):
        """Zum Meldungsreiter wechseln (Index 4)."""
        self.reiter.select(len(SCHRITTE))

    def _zaehle_meldungen(self):
        """Die Zahl der Meldungen in die Reiterbeschriftung schreiben."""
        anzahl = getattr(self, "_meldungen", 0)
        self.reiter.tab(len(SCHRITTE),
                        text=f"{REITER_MELDUNGEN} ({anzahl})" if anzahl
                        else REITER_MELDUNGEN)

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
                    self._laufend.discard(sid)
                    if not self._laufend:
                        self.configure(cursor="")
                    fertig[sid](erg, fehler)
        except queue.Empty:
            pass
        except Exception:
            _spur()
        self.after(150, self._pumpe)

    def _schreibe(self, text: str):
        feld = getattr(self, "protokollfeld", None)
        if feld is None or not feld.winfo_exists():
            return
        self._meldungen = getattr(self, "_meldungen", 0) + 1
        feld.configure(state=NORMAL)
        feld.insert(END, text + "\n")
        feld.see(END)
        feld.configure(state=DISABLED)

    def _starte(self, sid: str, arbeit):
        """Einen Arbeitsschritt im Hintergrund anstoßen.

        Ein zweiter Klick, weil sich scheinbar nichts rührt, darf denselben
        Schritt nicht ein zweites Mal starten — zwei Threads, die dieselbe
        Datei schreiben, oder zwei Importläufe, die nacheinander den Bestand
        ersetzen, enden unvorhersehbar.
        """
        if sid in self._laufend:
            return
        self._laufend.add(sid)
        self.configure(cursor="watch")

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
            messagebox.showwarning("Jemand anderes arbeitet gerade daran", hinweis)
        # Ein Zwischenstand, der jünger ist als der gespeicherte Bestand,
        # stammt von einem Lauf, der nicht ordentlich beendet wurde.
        angebot = core.offener_zwischenstand()
        if angebot and messagebox.askyesno("Nicht gespeicherte Eingaben", angebot):
            self.status.set("Hole die nicht gespeicherten Eingaben zurück …")
            self._starte("laden",
                         lambda: core.lade_bestand(core.WIEDERHERSTELLUNG_PFAD))
            self.geaendert = True
            return
        self.status.set("Lade die gespeicherten Daten …")
        self._starte("laden", core.lade_bestand)

    def _laden_fertig(self, bestand, fehler):
        if fehler:
            return self._fehler(fehler, "Die gespeicherten Daten konnten nicht geöffnet werden")
        self.bestand = bestand
        core.setze_arbeitssperre()
        for w in bestand.warnungen:
            self._schreibe("Hinweis beim Laden: " + w)
        if bestand.warnungen:
            self._zeige_meldungen()
        self.status.set(
            f"{len(bestand.empfaenger)} Empfänger, "
            f"{sum(1 for _ in bestand.buecher())} Bücher geladen.")
        self._zeige_stammdaten()
        self._zeige_erfassung()
        self._male_schrittleiste()

    def _speichern(self):
        self.status.set("Speichere …")
        jahr = self.jahr.get()
        self._starte("speichern",
                     lambda: core.speichere_bestand(self.bestand, jahr=jahr,
                                                    cfg=self.cfg))

    def _speichern_fertig(self, pfad, fehler):
        if fehler:
            return self._fehler(fehler, "Speichern hat nicht geklappt")
        self.geaendert = False
        self._seit_sicherung = 0
        core.verwerfe_zwischenstand()
        self.status.set(f"Gespeichert: {pfad}")
        self._schreibe(f"Gespeichert: {pfad}")

    def _importieren(self):
        if self.bestand.empfaenger and not messagebox.askokcancel(
                "Altmappe importieren",
                "Der Import ersetzt den gesamten aktuellen Bestand.\n\n"
                "Fortfahren?"):
            return
        pfad = filedialog.askopenfilename(
            title="Die alte Excel-Tabelle mit den Honoraren auswählen",
            initialdir=self.cfg.get("last_input_dir") or str(core.APP_DIR),
            filetypes=[("Excel-Tabelle", "*.xlsx"), ("Alle Dateien", "*.*")])
        if not pfad:
            return
        self.cfg["last_input_dir"] = str(Path(pfad).parent)
        core.speichere_config(self.cfg)
        jahr = self.jahr.get()
        self._zeige_meldungen()
        self.status.set("Importiere …")
        self._starte("import",
                     lambda: core.importiere_alt(pfad, jahr, log=self._melde()))

    def _import_fertig(self, erg, fehler):
        if fehler:
            return self._fehler(fehler, "Die alte Tabelle konnte nicht gelesen werden")
        bestand, prot = erg
        self.bestand = bestand
        self.geaendert = True
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
        self._male_schrittleiste()
        messagebox.showinfo(
            "Import fertig",
            f"{len(bestand.empfaenger)} Empfänger und "
            f"{sum(1 for _ in bestand.buecher())} Bücher gelesen.\n\n"
            f"{len(prot.warnungen)} Stellen brauchen einen Blick — sie stehen "
            f"im Protokoll.\n\nDer Bestand ist noch NICHT gespeichert.")

    def _jahr_gewechselt(self):
        self._zeige_erfassung()
        self._etwas_erzeugt = False
        self.abrechnungen = []
        for k in self.baum_durchlauf.get_children(""):
            self.baum_durchlauf.delete(k)
        self.summe.set("")
        self._male_schrittleiste()

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
            hinweise = []
            # Eine fehlende Bankverbindung nur dort melden, wo sie gebraucht
            # wird: 126 der 250 Empfänger haben keine, aber die meisten sind
            # stillgelegt oder warten auf eine Entscheidung. Ein Hinweis, der
            # bei der Hälfte steht, wird überlesen.
            aktiv = any(not b.stillgelegt and not b.gesondert
                        for b in e.buecher)
            if not e.iban and aktiv:
                hinweise.append("keine Bankverbindung")
            gesondert = sum(1 for b in e.buecher
                            if b.gesondert and not b.stillgelegt)
            if gesondert:
                hinweise.append(f"{gesondert} wartet auf Entscheidung"
                                if gesondert == 1 else
                                f"{gesondert} warten auf Entscheidung")
            offen = sum(1 for b in e.buecher if b.nachpflege)
            if offen:
                hinweise.append("1 Buch bitte prüfen" if offen == 1
                                else f"{offen} Bücher bitte prüfen")
            hinweis = " · ".join(hinweise)
            baum.insert("", "end", iid=e.kennung, values=(
                e.anzeigename, f"{e.plz} {e.ort}".strip(), e.iban,
                len(e.buecher), hinweis))

    def _gewaehlter_empfaenger(self):
        auswahl = self.baum_stamm.selection()
        if not auswahl:
            messagebox.showinfo(
                "Erst einen Autor auswählen",
                "Bitte in der Liste die Zeile anklicken, um die es geht.")
            return None
        return next((x for x in self.bestand.empfaenger
                     if x.kennung == auswahl[0]), None)

    def _neues_buch_zur_auswahl(self):
        e = self._gewaehlter_empfaenger()
        if e is not None:
            self._neues_buch(e)

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

        knoepfe = ttk.Frame(fenster)
        knoepfe.pack(fill="x", padx=12)
        ttk.Button(knoepfe, text="Angaben ändern …",
                   command=lambda: (fenster.destroy(),
                                    self._aendere_empfaenger(e))).pack(side="left")
        ttk.Button(knoepfe, text="Neues Buch …",
                   command=lambda: (fenster.destroy(),
                                    self._neues_buch(e))).pack(side="left", padx=6)

        unten = ttk.LabelFrame(
            fenster, text="Bücher — und wie das Honorar berechnet wird")
        unten.pack(fill="both", expand=True, **PAD)
        spalten = [("isbn", "ISBN", 80), ("titel", "Buchtitel", 300),
                   ("art", "Art der Zahlung", 130), ("satz", "€ je Ex.", 90),
                   ("bes", "Besonderheit", 260)]
        rahmen = ttk.Frame(unten)
        rahmen.pack(fill="both", expand=True, **PAD)
        baum = ttk.Treeview(rahmen, columns=[s[0] for s in spalten],
                            show="headings", selectmode="browse")
        for schluessel, titel, breite in spalten:
            baum.heading(schluessel, text=titel)
            baum.column(schluessel, width=breite, stretch=(breite > 200))
        leiste = ttk.Scrollbar(rahmen, orient="vertical", command=baum.yview)
        baum.configure(yscrollcommand=leiste.set)
        baum.pack(side="left", fill="both", expand=True)
        leiste.pack(side="right", fill="y")
        baum.tag_configure("still", foreground="#888888")

        def fuelle():
            for kind in baum.get_children(""):
                baum.delete(kind)
            for b in e.buecher:
                baum.insert("", "end", iid=b.kennung, values=(
                    b.isbn, b.titel, b.verguetungsart,
                    core.euro(core.satz_aus_kondition(b.kondition)),
                    core.besonderheit(b)),
                    tags=("still",) if b.stillgelegt else ())
        fuelle()

        def gewaehltes_buch():
            auswahl = baum.selection()
            if not auswahl:
                messagebox.showinfo(
                    "Erst ein Buch auswählen",
                    "Bitte in der Liste die Zeile anklicken, um die es geht.",
                    parent=fenster)
                return None
            return next((b for b in e.buecher if b.kennung == auswahl[0]), None)

        def aendern(_ereignis=None):
            b = gewaehltes_buch()
            if b is not None:
                self._aendere_buch(e, b)
                fuelle()

        baum.bind("<Double-1>", aendern)
        ttk.Label(unten, text="Doppelklick auf ein Buch ändert seine Angaben."
                  ).pack(anchor="w", padx=12)

        buchknoepfe = ttk.Frame(fenster)
        buchknoepfe.pack(fill="x", padx=12, pady=(0, 12))
        ttk.Button(buchknoepfe, text="Buch ändern …",
                   command=aendern).pack(side="left")
        ttk.Button(
            buchknoepfe, text="Neues Buch …",
            command=lambda: (self._neues_buch(e), fuelle())).pack(side="left",
                                                                 padx=6)

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
                "" if stand is None else stand,
                core.besonderheit(b)),
                tags=("gesondert",) if b.gesondert else ())
        self._zeige_fortschritt()

    def _bearbeite_zelle(self, ereignis):
        """Doppelklick oder Eingabetaste öffnet das Feld unter dem Zeiger."""
        baum = self.baum_erfassung
        if ereignis.type == "4":                       # Mausklick
            spalte = baum.identify_column(ereignis.x)
            zeile = baum.identify_row(ereignis.y)
        else:                                          # Eingabetaste
            auswahl = baum.selection()
            zeile = auswahl[0] if auswahl else ""
            spalte = "#5"
        self._oeffne_eingabe(zeile, spalte)

    def _oeffne_eingabe(self, zeile: str, spalte: str):
        """Das Eingabefeld über einer Zelle.

        Bei rund 490 Zeilen entscheidet der Tastenfluss darüber, ob die
        Arbeit eine Stunde oder einen Vormittag dauert:

        * Eingabetaste übernimmt und öffnet SOFORT das Feld eine Zeile
          tiefer — nicht bloß die Auswahl, sonst muss man je Zeile zweimal
          drücken.
        * Tabulator springt zwischen „verkaufte Ex.“ und „Eigenkauf“.
        * Esc verwirft wirklich. Das ist kniffliger als es aussieht: das
          Zerstören eines Feldes löst in Tk ein FocusOut aus, und daran hängt
          das Übernehmen — ohne Merker würde Esc den Wert also speichern,
          genau umgekehrt zur Erwartung.
        """
        baum = self.baum_erfassung
        if not zeile or spalte not in ("#5", "#6"):
            return
        # Ohne see() liefert bbox für eine weggescrollte Zeile "" — das gäbe
        # einen Fehler, der im Fensterbau spurlos verschwände.
        baum.see(zeile)
        kasten = baum.bbox(zeile, spalte)
        if not kasten:
            return
        x, y, breite, hoehe = kasten
        feldname = "verkauft" if spalte == "#5" else "eigenkauf"
        wert = StringVar(value=baum.set(zeile, feldname))
        eingabe = ttk.Entry(baum, textvariable=wert, justify="right")
        eingabe.place(x=x, y=y, width=breite, height=hoehe)
        eingabe.focus_set()
        eingabe.selection_range(0, END)
        erledigt = {"ja": False}

        def schliesse():
            erledigt["ja"] = True
            eingabe.destroy()
            baum.focus_set()          # sonst tut die Eingabetaste nichts mehr

        def verwerfen(_e=None):
            schliesse()

        def uebernehmen(weiter: bool, andere_spalte: str = ""):
            if erledigt["ja"]:
                return
            text = wert.get().strip()
            zahl, meldung = core.lies_stueckzahl(text)
            # Eine unlesbare Eingabe darf NICHT stillschweigend zu einer
            # anderen Zahl oder zu einem leeren Feld werden. Der deutsche
            # Tausenderpunkt ist die gefährliche Stelle: „1.234“ wurde früher
            # klaglos zu 1 — aus 1234 verkauften Exemplaren wurde eines.
            if meldung:
                messagebox.showwarning("Bitte noch einmal eintragen", meldung)
                eingabe.focus_set()
                eingabe.selection_range(0, END)
                return
            schliesse()
            buch = self._buch(zeile)
            if buch is None:
                return
            jw = buch.jahr(self.jahr.get())
            if feldname == "verkauft":
                jw.verkauft = zahl
            else:
                jw.eigenkauf = zahl or 0
            if baum.exists(zeile):
                baum.set(zeile, feldname, "" if zahl is None else zahl)
            self.geaendert = True
            # Ab jetzt stimmt das zuletzt gerechnete Ergebnis nicht mehr.
            # Ohne diesen Merker entstünden die Briefe aus den ALTEN Zahlen —
            # also ausgerechnet dann falsch, wenn sorgfältig nachgebessert
            # wurde.
            self._rechnung_veraltet = bool(self.abrechnungen)
            self._seit_sicherung += 1
            if self._seit_sicherung >= 20:
                # Stiller Zwischenstand. Läuft im Hauptthread, dauert aber
                # nur Sekundenbruchteile und passiert nur alle 20 Eingaben.
                self._seit_sicherung = 0
                if core.sichere_zwischenstand(self.bestand, self.jahr.get(),
                                              self.cfg):
                    self.status.set("Zwischenstand gesichert.")
            self._zeige_fortschritt()
            self._male_schrittleiste()
            if andere_spalte:
                self.after(1, lambda: self._oeffne_eingabe(zeile, andere_spalte))
            elif weiter:
                nachbar = baum.next(zeile)
                if nachbar:
                    baum.selection_set(nachbar)
                    baum.focus(nachbar)
                    self.after(1, lambda: self._oeffne_eingabe(nachbar, spalte))

        def weiter_tab(_e=None):
            uebernehmen(False, "#6" if spalte == "#5" else "#5")
            return "break"

        eingabe.bind("<Return>", lambda e: uebernehmen(True))
        eingabe.bind("<Tab>", weiter_tab)
        eingabe.bind("<FocusOut>", lambda e: uebernehmen(False))
        eingabe.bind("<Escape>", verwerfen)

    def _springe_zur_luecke(self):
        """Zur ersten Zeile, bei der noch keine Zahl steht.

        Wer nach der Mittagspause zurückkommt, soll nicht 490 Zeilen
        durchscrollen, um die Stelle wiederzufinden.
        """
        baum = self.baum_erfassung
        for kennung in baum.get_children(""):
            if not str(baum.set(kennung, "verkauft")).strip():
                baum.selection_set(kennung)
                baum.focus(kennung)
                baum.see(kennung)
                baum.focus_set()
                self.status.set("Hier geht es weiter.")
                return
        self.status.set("Bei allen Büchern steht eine Zahl.")

    def _zeige_fortschritt(self):
        """Zeigt, was NOCH FEHLT — nicht, was schon geschafft ist.

        „384 von 389 erledigt“ klingt nach Feierabend; „noch 5 Bücher ohne
        Zahl“ sagt, was zu tun ist. Genau diese fünf verhindern sonst eine
        vollständige Abrechnung.
        """
        jahr = self.jahr.get()
        gesamt = sum(1 for e in self.bestand.empfaenger
                     for b in e.buecher if not b.stillgelegt)
        offen = sum(1 for e in self.bestand.empfaenger for b in e.buecher
                    if not b.stillgelegt
                    and not (jahr in b.jahre and b.jahre[jahr].erfasst))
        if not gesamt:
            self.fortschritt.set("")
        elif offen:
            self.fortschritt.set(f"Noch {offen} von {gesamt} Büchern ohne Zahl")
        else:
            self.fortschritt.set(f"✓ Alle {gesamt} Bücher haben eine Zahl")

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
            return self._fehler(fehler, "Beim Rechnen ist etwas schiefgegangen")
        self.abrechnungen = abrechnungen
        self._rechnung_veraltet = False
        self._male_durchlauf()
        # Was wegen einer offenen Entscheidung NICHT ausgezahlt wird, gehört
        # bei jedem Durchlauf vor Augen — nicht nur einmal ins Importprotokoll.
        anzahl, betrag = core.zurueckgehalten(self.bestand, self.jahr.get(),
                                              self.cfg)
        if anzahl:
            self.zurueck.set(
                f"⚠ {anzahl} Empfänger bekommen zusammen "
                f"{core.euro(betrag)} NICHT — ihre Verträge stehen auf "
                f"„gesondert abrechnen“ und warten auf eine Entscheidung "
                f"des Verlags.")
        else:
            self.zurueck.set("")
        self.status.set("Fertig gerechnet.")
        self._male_schrittleiste()

    def _gegenprobe(self):
        pfad = filedialog.askopenfilename(
            title="Die alte Excel-Tabelle zum Vergleichen auswählen",
            initialdir=self.cfg.get("last_input_dir") or str(core.APP_DIR),
            filetypes=[("Excel-Tabelle", "*.xlsx"), ("Alle Dateien", "*.*")])
        if not pfad:
            return
        jahr = self.jahr.get()
        self._zeige_meldungen()
        self.status.set("Vergleiche mit der alten Excel-Tabelle …")
        self._starte("gegenprobe", lambda: core.pruefe_gegen_excel(
            self.bestand, pfad, jahr, self.cfg))

    def _gegenprobe_fertig(self, abweichungen, fehler):
        if fehler:
            return self._fehler(fehler, "Der Vergleich hat nicht geklappt")
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

    def _male_durchlauf(self):
        """Die Liste füllen — in einem von zwei Modi.

        „Alle“ zeigt Empfänger mit ihrer Auszahlung. „Nur die offenen
        Sonderfälle“ zeigt stattdessen die BÜCHER mit dem Betrag, der wegen
        der offenen Entscheidung nicht fließt: neunzehn der zwanzig
        Empfänger haben gar keine normale Auszahlung, ihre Zeilen wären in
        der gewöhnlichen Ansicht durchweg null.
        """
        baum = self.baum_durchlauf
        for k in baum.get_children(""):
            baum.delete(k)
        if not self.abrechnungen:
            return

        if self.nur_sonderfaelle.get():
            self._setze_spalten(SPALTEN_SONDERFAELLE)
            zeilen = core.gesonderte_posten(self.bestand, self.jahr.get(),
                                            self.cfg)
            self._sonderzeilen = zeilen
            for i, z in enumerate(zeilen):
                baum.insert("", "end", iid=f"s{i}", values=(
                    z["Empfänger"], z["Buchtitel"], z["ISBN"],
                    f"{z['Vergütungs-Ex.']} Ex.", core.euro(z["Betrag"]),
                    z["Bankverbindung"] or "— fehlt —",
                    (z["in der Altmappe"] or "").replace("!", ", Zeile ")),
                    tags=("sonderfall",))
            summe = core.runde(sum(z["Betrag"] for z in zeilen))
            empfaenger = len({z["Empfänger"] for z in zeilen})
            self.summe.set(
                f"{len(zeilen)} Bücher von {empfaenger} Autoren · "
                f"zusammen {core.euro(summe)}, die NICHT ausgezahlt werden")
            self.status.set(
                "Häkchen wegnehmen zeigt wieder alle Empfänger.")
            return

        self._setze_spalten(SPALTEN_DURCHLAUF)
        offen = {z["Empfänger"] for z in core.gesonderte_posten(
            self.bestand, self.jahr.get(), self.cfg)}
        for ab in sorted(self.abrechnungen, key=lambda a: -a.brutto):
            hinweise = list(ab.gruende) + list(ab.probleme)
            marken = []
            if ab.empfaenger.anzeigename in offen:
                marken.append("sonderfall")
            elif not ab.brief:
                marken.append("kein_brief")
            elif ab.probleme:
                marken.append("problem")
            baum.insert("", "end", iid=ab.empfaenger.kennung, values=(
                ab.empfaenger.anzeigename, ab.verguetungsart, len(ab.posten),
                core.euro(ab.netto), core.euro(ab.mwst), core.euro(ab.brutto),
                "Ja" if ab.brief else "—", " | ".join(hinweise)),
                tags=marken)
        mit = [a for a in self.abrechnungen if a.brief]
        meldbar = [a for a in self.abrechnungen if a.ksk_netto > 0]
        self.summe.set(
            f"{len(mit)} Briefe · Auszahlung "
            f"{core.euro(sum(a.brutto for a in mit))} · an die "
            f"Künstlersozialkasse zu melden: "
            f"{core.euro(sum(a.ksk_netto for a in meldbar))} "
            f"({len(meldbar)} Autoren)")

    def _zeige_sonderfall(self, nummer: int):
        """Wie der einbehaltene Betrag zustande kommt — und warum er bleibt."""
        zeilen = getattr(self, "_sonderzeilen", [])
        if nummer >= len(zeilen):
            return
        z = zeilen[nummer]
        buch = next((b for _, b in self.bestand.buecher()
                     if b.kennung == z.get("_kennung")), None)
        fenster = Toplevel(self)
        fenster.title(f"Offener Sonderfall — {z['Buchtitel']}")
        fenster.geometry("860x520")
        feld = Text(fenster, wrap="word", font=("Segoe UI", 10))
        feld.pack(fill="both", expand=True, **PAD)
        feld.insert(END, f"{z['Empfänger']}\n{z['Buchtitel']}\n\n")
        if buch is not None:
            posten = core.betrag_zeile(buch, self.jahr.get(), self.cfg)
            if posten is not None:
                for zeile in core.rechenweg(posten):
                    feld.insert(END, zeile + "\n")
                feld.insert(END, "\n")
        feld.insert(END, "─" * 68 + "\n")
        feld.insert(END,
                    f"Dieser Betrag von {core.euro(z['Betrag'])} wird "
                    f"NICHT ausgezahlt.\n\n"
                    f"Das Buch steht auf „gesondert abrechnen“. In der alten "
                    f"Excel-Tabelle stand es im Blatt „Zahlung ab XX Ex.“ "
                    f"({z['in der Altmappe']}), das dort nur einen Stand "
                    f"führte und keine Auszahlungen auslöste — die Summe "
                    f"jener Spalte war negativ, und fast niemand daraus kam "
                    f"in der Zahlungsliste vor.\n\n"
                    f"Soll der Betrag doch fließen, nimmt man im Bestand im "
                    f"Blatt „Bücher“ den Haken in der Spalte „Gesondert "
                    f"abrechnen“ weg. Das ist eine Entscheidung des "
                    f"Verlags.\n")
        feld.configure(state=DISABLED)

    def _setze_spalten(self, spalten):
        """Die Spalten der Durchlauf-Liste umschalten."""
        baum = self.baum_durchlauf
        if getattr(self, "_spalten_jetzt", None) == spalten:
            return
        self._spalten_jetzt = spalten
        baum.configure(columns=[s[0] for s in spalten])
        for schluessel, titel, breite in spalten:
            baum.heading(schluessel, text=titel,
                         command=lambda s=schluessel: self._sortiere(baum, s))
            baum.column(schluessel, width=breite, stretch=(breite > 200))

    def _zeige_rechenweg(self, ereignis=None):
        """„Wie dieser Betrag zustande kommt“ — in ganzen Sätzen.

        In Excel konnte man in die Zelle klicken und die Formel lesen. Ohne
        eine Entsprechung dazu bleibt nur, dem Werkzeug zu glauben — und
        genau das tut niemand bei einer Abrechnung über 20.000 €.
        """
        auswahl = self.baum_durchlauf.selection()
        if not auswahl:
            return
        if auswahl[0].startswith("s"):
            # Im Sonderfall-Modus steht in der Zeile ein Buch, kein Empfänger.
            self._zeige_sonderfall(int(auswahl[0][1:]))
            return
        ab = next((a for a in self.abrechnungen
                   if a.empfaenger.kennung == auswahl[0]), None)
        if ab is None:
            return
        fenster = Toplevel(self)
        fenster.title(f"Wie der Betrag zustande kommt — {ab.empfaenger.anzeigename}")
        fenster.geometry("900x640")
        feld = Text(fenster, wrap="word", font=("Segoe UI", 10))
        feld.pack(fill="both", expand=True, **PAD)
        for posten in ab.posten:
            for zeile in core.rechenweg(posten):
                feld.insert(END, zeile + "\n")
            feld.insert(END, "\n")
        feld.insert(END, "─" * 70 + "\n")
        feld.insert(END, f"Summe netto:      {core.euro(ab.netto)}\n")
        if ab.mwst:
            feld.insert(END, f"Mehrwertsteuer:   {core.euro(ab.mwst)}\n")
        if ab.verrechnet:
            feld.insert(END, f"einbehalten für den offenen Vorschuss: "
                             f"−{core.euro(ab.verrechnet)}\n")
        feld.insert(END, f"Auszahlung:       {core.euro(ab.brutto)}\n")
        if ab.gruende:
            feld.insert(END, "\nEs geht kein Brief hinaus: "
                             + "; ".join(ab.gruende) + "\n")
        feld.configure(state=DISABLED)

    # -----------------------------------------------------------------
    # Reiter 4: Ausgaben
    # -----------------------------------------------------------------

    def _bereit(self) -> bool:
        if not self.abrechnungen:
            messagebox.showinfo(
                "Erst die Beträge berechnen",
                "Bitte zuerst auf Reiter „3. Nachrechnen“ den Knopf "
                "„Beträge berechnen“ drücken.")
            return False
        if getattr(self, "_rechnung_veraltet", False):
            messagebox.showwarning(
                "Die Zahlen haben sich geändert",
                "Seit dem letzten Rechnen wurden Stückzahlen geändert.\n\n"
                "Bitte auf Reiter „3. Nachrechnen“ noch einmal „Beträge "
                "berechnen“ drücken — sonst stünden in den Briefen die alten "
                "Beträge.")
            self.reiter.select(2)
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
            return self._fehler(fehler, "Die Briefe konnten nicht geschrieben werden")
        erzeugt, uebergangen = erg
        self._briefe_pfade = erzeugt
        self._schreibe(f"\n{len(erzeugt)} Briefe erzeugt.")
        if uebergangen:
            self._schreibe(f"Ohne Brief ({len(uebergangen)}):")
            for u in uebergangen:
                self._schreibe("   " + u)
        self._etwas_erzeugt = True
        self.ergebnis.set(
            f"{len(erzeugt)} Briefe geschrieben, {len(uebergangen)} Autoren "
            f"ohne Brief (Begründung im Protokoll).")
        self.status.set(f"{len(erzeugt)} Briefe erzeugt.")
        self._male_schrittleiste()

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
        self.ergebnis.set(f"{len(erzeugt)} PDF-Dateien erzeugt.")
        self.status.set(f"{len(erzeugt)} PDF-Dateien erzeugt.")

    def _listen(self):
        if not self._bereit():
            return
        jahr = self.jahr.get()
        ordner = core.ausgabeordner(self.cfg, jahr)

        def arbeit():
            return [core.schreibe_listen(
                self.abrechnungen, jahr,
                ordner / f"Zahlungsliste_{jahr}.xlsx", self.cfg,
                bestand=self.bestand)]

        self.status.set("Schreibe Listen …")
        self._starte("listen", arbeit)

    def _listen_fertig(self, pfade, fehler):
        if fehler:
            return self._fehler(fehler, "Die Listen konnten nicht geschrieben werden")
        for p in pfade:
            self._schreibe(f"geschrieben: {p}")
        self._etwas_erzeugt = True
        self.ergebnis.set(
            "Zahlungsliste, Meldung an die Künstlersozialkasse und Protokoll "
            "stehen als drei Blätter in einer Datei.")
        self.status.set("Listen geschrieben.")
        self._male_schrittleiste()

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
        if self.geaendert:
            antwort = messagebox.askyesnocancel(
                "Änderungen speichern?",
                "Sie haben Zahlen geändert, die noch nicht gespeichert sind.\n\n"
                "Ja  — speichern und schließen\n"
                "Nein — schließen und die Änderungen verwerfen")
            if antwort is None:
                return                      # Abbrechen: Fenster bleibt offen
            if antwort:
                try:
                    core.speichere_bestand(self.bestand)
                except Exception as e:
                    messagebox.showerror("Speichern fehlgeschlagen", str(e))
                    return                  # nicht schließen, sonst ist es weg
        core.loese_arbeitssperre()
        self.destroy()


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
