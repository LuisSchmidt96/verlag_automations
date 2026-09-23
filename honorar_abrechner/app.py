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

    FELDER_BUCH = [
        ("titel", "Buchtitel", "text"),
        ("isbn", "ISBN (Kurzform, z. B. 05-300)", "text"),
        ("verguetungsart", "Art der Zahlung", list(core.VERGUETUNGSARTEN)),
        ("betrag_je_ex", "Fester Betrag je Exemplar (€)", "zahl"),
        ("ladenpreis", "oder: Ladenpreis (€)", "zahl"),
        ("mwst_im_preis", "davon MwSt herausrechnen (7 / 19, sonst leer)", "zahl"),
        ("verlagsrabatt", "Verlagsrabatt in % (meist 40)", "zahl"),
        ("satz", "Honorarsatz in %", "zahl"),
        ("teiler", "geteilt durch (Mitautoren)", "zahl"),
        ("mwst_pflichtig", "Autor ist mehrwertsteuerpflichtig",
         ["Nein", "Ja"]),
        ("schwelle_zehn", "Kein Honorar unter zehn Exemplaren",
         ["Nein", "Ja"]),
        ("freimenge", "Freimenge (erste N Exemplare ohne Honorar)", "zahl"),
        ("notizen", "Notizen", "text"),
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

    def _neues_buch(self, e):
        werte = self._formular(f"Neues Buch für {e.anzeigename}",
                               self.FELDER_BUCH,
                               {"verlagsrabatt": 40, "teiler": 1,
                                "verguetungsart": "Honorar",
                                "mwst_pflichtig": "Nein",
                                "schwelle_zehn": "Nein"})
        if werte is None:
            return
        if not werte["titel"]:
            messagebox.showwarning("Titel fehlt",
                                   "Ohne Buchtitel geht es nicht.")
            return
        buch = core.Buch(kennung=self.bestand.naechste_kennung("B"))
        self._uebertrage_buch(buch, werte)
        e.buecher.append(buch)
        self.geaendert = True
        self._zeige_stammdaten()
        self._zeige_erfassung()
        self._male_schrittleiste()
        self.status.set(f"„{buch.titel}“ angelegt — noch nicht gespeichert.")

    @staticmethod
    def _uebertrage_buch(buch, werte: dict):
        """Die Formularwerte in das Buch schreiben."""
        buch.titel = werte["titel"]
        buch.isbn = werte["isbn"]
        buch.verguetungsart = werte["verguetungsart"] or "Honorar"
        buch.mwst_pflichtig = werte["mwst_pflichtig"] == "Ja"
        buch.notizen = werte["notizen"]
        k = buch.kondition
        k.betrag_je_ex = werte["betrag_je_ex"] or None
        k.ladenpreis = werte["ladenpreis"] or None
        k.mwst_im_preis = werte["mwst_im_preis"] or None
        k.verlagsrabatt = werte["verlagsrabatt"] or 40.0
        k.rabatt_anwenden = bool(werte["verlagsrabatt"])
        k.satz = werte["satz"] or None
        k.teiler = int(werte["teiler"] or 1)
        k.freimenge = int(werte["freimenge"] or 0)
        k.schwelle_zehn = werte["schwelle_zehn"] == "Ja"

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
        self.baum_erfassung.bind("<Double-1>", self._bearbeite_zelle)
        self.baum_erfassung.bind("<Return>", self._bearbeite_zelle)
        # Der Unterschied zwischen einer leeren Zelle und einer 0 entscheidet
        # über Geld. Er muss dastehen, nicht in einem Fachwort stecken.
        hinweis = ttk.Frame(seite)
        hinweis.pack(fill="x", padx=12, pady=(0, 8))
        ttk.Label(hinweis, justify="left", text=(
            "Doppelklick auf eine Zahl öffnet das Feld. Die Eingabetaste "
            "übernimmt und springt eine Zeile tiefer, Esc verwirft.")
        ).pack(anchor="w")
        ttk.Label(hinweis, justify="left", foreground="#8a5a00", text=(
            "Wichtig: Ein leeres Feld heißt „habe ich noch nicht eingetragen“. "
            "Hat sich ein Buch im ganzen Jahr nicht verkauft, tragen Sie dort "
            "bitte eine 0 ein — sonst fehlt das Buch in der Abrechnung.")
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
        self.summe = StringVar(value="")
        ttk.Label(oben, textvariable=self.summe,
                  font=("Segoe UI", 10, "bold")).pack(side="right")

        self.baum_durchlauf = self._liste(seite, SPALTEN_DURCHLAUF,
                                          doppelklick=self._zeige_rechenweg)
        self.baum_durchlauf.tag_configure("kein_brief", foreground="#777777")
        self.baum_durchlauf.tag_configure("problem", foreground="#a4262c")
        # Legende: Farben ohne Erklärung sind keine Auskunft.
        legende = ttk.Frame(seite)
        legende.pack(fill="x", padx=12)
        ttk.Label(legende, text="grau = bekommt keinen Brief",
                  foreground="#777777").pack(side="left", padx=(0, 20))
        ttk.Label(legende, text="rot = bitte ansehen, bevor der Brief hinausgeht",
                  foreground="#a4262c").pack(side="left")
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
            offen = [b for b in e.buecher if b.nachpflege]
            hinweis = ""
            if not e.iban:
                hinweis = "keine Bankverbindung"
            elif offen:
                hinweis = f"{len(offen)} Buch/Bücher zur Nachpflege"
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

        unten = ttk.LabelFrame(fenster, text="Bücher — und wie das Honorar berechnet wird")
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
        # Nur positive Honorare werden gemeldet. Würde man alle aufsummieren,
        # kürzten sich die negativen (Autoren mit mehr Rückgaben als
        # Verkäufen) heraus und der angezeigte Meldebetrag wäre zu niedrig —
        # er stünde in keiner Beziehung zu dem, was in der Liste landet.
        meldbar = [a for a in abrechnungen if a.ksk_netto > 0]
        self.summe.set(
            f"{len(mit)} Briefe · Auszahlung "
            f"{core.euro(sum(a.brutto for a in mit))} · an die "
            f"Künstlersozialkasse zu melden: "
            f"{core.euro(sum(a.ksk_netto for a in meldbar))} "
            f"({len(meldbar)} Autoren)")
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

    def _zeige_rechenweg(self, ereignis=None):
        """„Wie dieser Betrag zustande kommt“ — in ganzen Sätzen.

        In Excel konnte man in die Zelle klicken und die Formel lesen. Ohne
        eine Entsprechung dazu bleibt nur, dem Werkzeug zu glauben — und
        genau das tut niemand bei einer Abrechnung über 20.000 €.
        """
        auswahl = self.baum_durchlauf.selection()
        if not auswahl:
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
                ordner / f"Zahlungsliste_{jahr}.xlsx", self.cfg)]

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
