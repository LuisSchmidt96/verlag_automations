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
import re
import subprocess
import sys
import threading
import traceback
from datetime import date
from pathlib import Path
from tkinter import (Tk, StringVar, BooleanVar, IntVar, END, NORMAL, DISABLED,
                     TclError, filedialog, messagebox, ttk, Text, Toplevel,
                     Listbox)

from honorar_abrechner import core


def dez(zahl: float) -> str:
    """Eine Zahl so, wie man sie hier schreibt: 0,84 statt 0.84."""
    return f"{zahl:g}".replace(".", ",")


def _sperre_eingabe(fenster, versuche: int = 20) -> None:
    """Das Fenster modal machen — erst, wenn es auf dem Bildschirm ist.

    Unter X11 scheitert grab_set mit „grab failed: window not viewable“,
    solange der Fenstermanager das neue Fenster noch nicht gezeigt hat. Die
    Ausnahme brach den Aufbau des Buch-Dialogs ab, noch bevor ein einziges
    Feld darin stand: übrig blieb ein leeres Fenster mit Titel. Also nicht
    warten (das zeigte das Fenster leer an), sondern es kurz darauf noch
    einmal versuchen, während der Inhalt schon entsteht.
    """
    try:
        fenster.grab_set()
    except TclError:
        if versuche and fenster.winfo_exists():
            fenster.after(50, lambda: fenster.winfo_exists()
                          and _sperre_eingabe(fenster, versuche - 1))


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

SPALTEN_BUECHER = [("isbn", "ISBN", 90), ("titel", "Buchtitel", 330),
                   ("autor", "Autor / Empfänger", 260),
                   ("art", "Art der Zahlung", 130), ("satz", "€ je Ex.", 90),
                   ("bes", "Besonderheit", 320)]

SPALTEN_ERFASSUNG = [("isbn", "ISBN", 90), ("titel", "Buchtitel", 300),
                     ("empf", "Empfänger", 200), ("art", "Vergütungsart", 120),
                     ("verkauft", "verkaufte Ex. ✎", 110),
                     ("eigenkauf", "Eigenkauf ✎", 100),
                     # Hier landet, was das Vorjahr schuldig geblieben ist:
                     # ein Fehlbetrag aus Rückgaben, oder eine Menge, die
                     # schon abgerechnet wurde. Ohne diese Spalte ist der
                     # Hinweis „im nächsten Jahr unter Korrektur eintragen“
                     # eine Anweisung, die man nicht befolgen kann.
                     ("korrektur", "Korrektur ✎", 95),
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


# Die Schritte in der Reihenfolge, in der gearbeitet wird. Die Beschriftung
# sagt, was zu TUN ist, nicht wie der Programmteil heißt. Autoren und Bücher
# waren ein Reiter; Bücher fand man dann nur über den Doppelklick auf ihren
# Autor — wer ein Buch sucht, weiß aber oft nicht, wem es gehört.
REITER_MELDUNGEN = "Meldungen"

SCHRITTE = [
    ("stamm", "1. Autoren"),
    ("buecher", "2. Bücher"),
    ("zahlen", "3. Zahlen eintragen"),
    ("rechnen", "4. Nachrechnen"),
    ("ausgeben", "5. Briefe und Listen"),
]


# Beträge und Mengen stehen in den Listen als fertiger Text: „903,00 €“,
# „1.483,14 €“, „265 Ex.“. Die Zahl steht vorn, dahinter kommt die Einheit.
_RX_ZAHL_VORN = re.compile(r"^\s*([+-]?\d{1,3}(?:\.\d{3})+|[+-]?\d+)(,\d+)?")


def _sortwert(wert):
    """Zahlen numerisch, alles andere alphabetisch — Leeres immer zuletzt.

    Die Einheit muss dabei abgeschnitten werden. Vorher scheiterte die
    Umwandlung am Eurozeichen, und dann wurde alphabetisch sortiert: „903“
    stand vor „92“, weil die Null vor der Zwei kommt. Bei einer Spalte, in
    der man nach dem größten Betrag sucht, ist das genau verkehrt.
    """
    text = str(wert if wert is not None else "").strip()
    if not text:
        return (2, 0.0, "")
    treffer = _RX_ZAHL_VORN.match(text)
    if treffer:
        ganz = treffer.group(1).replace(".", "")
        nachkomma = (treffer.group(2) or ",0").replace(",", ".")
        try:
            return (0, float(ganz + nachkomma), "")
        except ValueError:
            pass
    return (1, 0.0, text.lower())


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
        self.suche_buecher = StringVar()
        self.zeige_stillgelegte = BooleanVar(value=False)
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
        self._baue_buecher()
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
            "✓" if gesamt else "·",
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
            satz = (f"Als Nächstes: auf Reiter 3 noch {offen} Bücher "
                    f"eintragen.")
        elif not self.abrechnungen or veraltet:
            satz = ("Als Nächstes: auf Reiter 4 die Beträge berechnen.")
        elif not getattr(self, "_etwas_erzeugt", False):
            satz = ("Als Nächstes: das Ergebnis auf Reiter 4 durchsehen, "
                    "dann auf Reiter 5 die Briefe erzeugen.")
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
        _sperre_eingabe(fenster)
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

    def _neues_buch(self, e, eltern=None):
        buch = core.Buch(kennung=self.bestand.naechste_kennung("B"))
        # Angehängt wird erst beim „Übernehmen“ — das erledigt der Dialog,
        # weil dort auch Mitautoren dazukommen können.
        if not self._buchfenster(f"Neues Buch für {e.anzeigename}",
                                 [(e, buch)], eltern):
            return
        self.geaendert = True
        self._rechnung_veraltet = bool(self.abrechnungen)
        self._zeige_stammdaten()
        self._zeige_erfassung()
        self._male_schrittleiste()
        self.status.set(f"„{buch.titel}“ angelegt — noch nicht gespeichert.")

    def _aendere_buch(self, e, buch, eltern=None):
        """Ein bestehendes Buch ändern — mit allen seinen Autoren."""
        gruppe = self._buchgruppe(e, buch)
        if not self._buchfenster(f"Buch ändern — {buch.titel}", gruppe,
                                 eltern):
            return
        self.geaendert = True
        self._rechnung_veraltet = bool(self.abrechnungen)
        self._zeige_stammdaten()
        self._zeige_erfassung()
        self.status.set(f"„{buch.titel}“ geändert — noch nicht gespeichert.")

    @staticmethod
    def _buchschluessel(b) -> str:
        """Was zwei Einträge zum selben Buch macht: die ISBN, sonst der Titel."""
        return b.isbn.strip() or b.titel.strip().lower()

    def _buchgruppe(self, e, buch) -> list:
        """Das Buch samt den Einträgen seiner Mitautoren.

        Ein Buch mit zwei Autoren steht im Bestand zweimal, einmal je Autor
        — so rechnet und schreibt jeder Brief für sich. Bearbeitet wird es
        trotzdem als EIN Buch: der Ladenpreis ist für alle derselbe, nur der
        Anteil unterscheidet sich.
        """
        gruppe = [(e, buch)]
        schluessel = self._buchschluessel(buch)
        if not schluessel:
            return gruppe
        for x, b in self.bestand.buecher():
            if b is not buch and self._buchschluessel(b) == schluessel:
                gruppe.append((x, b))
        return gruppe

    # --- Entfernen ----------------------------------------------------
    # Loeschen ist hier fast nie das Richtige: ein Titel, der nicht mehr
    # abgerechnet wird, gehoert stillgelegt — dann bleibt die Historie
    # erhalten und man sieht in fuenf Jahren noch, was 2025 gezahlt wurde.
    # Entfernt wird nur, was versehentlich angelegt wurde. Deshalb nennt
    # die Rueckfrage beides und zaehlt auf, was verloren geht.

    def _entferne_buch(self, e, buch, danach=None, eltern=None) -> bool:
        jahre = sorted(j for j, w in buch.jahre.items() if w.erfasst)
        if jahre:
            verlust = (f"\n\nDamit gehen die erfassten Zahlen aus "
                       f"{len(jahre)} Jahren ({jahre[0]}–{jahre[-1]}) "
                       f"unwiderruflich verloren.\n\nSoll der Titel nur "
                       f"nicht mehr abgerechnet werden, ist „Buch ändern“ → "
                       f"„Wird nicht mehr abgerechnet“ der richtige Weg — "
                       f"dann bleibt die Historie erhalten.")
        else:
            verlust = "\n\nFür dieses Buch sind keine Zahlen erfasst."
        if not messagebox.askokcancel(
                "Buch entfernen",
                f"„{buch.titel}“ von {e.anzeigename} wirklich entfernen?"
                + verlust, icon="warning", default="cancel",
                parent=eltern or self):
            return False
        e.buecher.remove(buch)
        self.geaendert = True
        self._rechnung_veraltet = bool(self.abrechnungen)
        self._zeige_stammdaten()
        self._zeige_erfassung()
        self._male_schrittleiste()
        if danach is not None:
            danach()
        self.status.set(f"„{buch.titel}“ entfernt — noch nicht gespeichert.")
        return True

    def _entferne_empfaenger(self, e) -> bool:
        n = len(e.buecher)
        jahre = sorted({j for b in e.buecher for j, w in b.jahre.items()
                        if w.erfasst})
        teile = []
        if n:
            teile.append(f"{n} " + ("Buch" if n == 1 else "Bücher"))
        if jahre:
            teile.append(f"die Zahlen aus {len(jahre)} Jahren "
                         f"({jahre[0]}–{jahre[-1]})")
        if teile:
            verlust = ("\n\nDamit gehen " + " und ".join(teile)
                       + " unwiderruflich verloren.\n\nSoll nur nicht mehr "
                         "abgerechnet werden, lassen sich die Bücher einzeln "
                         "stilllegen — dann bleibt die Historie erhalten.")
        else:
            verlust = "\n\nZu diesem Autor sind keine Bücher erfasst."
        if not messagebox.askokcancel(
                "Autor entfernen",
                f"„{e.anzeigename}“ wirklich aus dem Bestand entfernen?"
                + verlust, icon="warning", default="cancel"):
            return False
        self.bestand.empfaenger.remove(e)
        self.geaendert = True
        self._rechnung_veraltet = bool(self.abrechnungen)
        self._zeige_stammdaten()
        self._zeige_erfassung()
        self._male_schrittleiste()
        self.status.set(f"„{e.anzeigename}“ entfernt — noch nicht "
                        f"gespeichert.")
        return True

    def _entferne_gewaehlten_empfaenger(self):
        e = self._gewaehlter_empfaenger()
        if e is not None:
            self._entferne_empfaenger(e)

    def _buchfenster(self, titel: str, gruppe: list, eltern=None) -> bool:
        """Buchangaben erfassen. True, wenn übernommen wurde.

        ``gruppe`` ist das Buch mit allen Autoren: [(Empfänger, Buch), …],
        das zuerst genannte ist das, von dem aus geöffnet wurde. Ein Buch,
        das noch in keinem Empfänger steht, wird beim Übernehmen angehängt.

        ``eltern`` ist das Fenster, aus dem heraus geöffnet wird — meist das
        Autorenfenster. Hinge der Dialog am Hauptfenster, holte der
        Fenstermanager beim Schließen das Hauptfenster nach vorn, und das
        Autorenfenster verschwände dahinter, als wäre es mit zugegangen.
        """
        fenster = Toplevel(eltern or self)
        fenster.title(titel)
        fenster.transient(eltern or self)
        _sperre_eingabe(fenster)
        try:
            return self._baue_buchfenster(fenster, gruppe)
        except Exception:
            # Sonst bliebe ein leeres Fenster mit Titel stehen und hielte
            # die Eingabesperre — der Fehler selbst landete nur im Terminal.
            if fenster.winfo_exists():
                fenster.destroy()
            raise

    def _baue_buchfenster(self, fenster, gruppe: list) -> bool:
        erster_e, buch = gruppe[0]
        jahr = self.jahr.get()
        # Was das BUCH je Exemplar abwirft, und wer wie viel Prozent davon
        # bekommt. Bei einem Autor ist der Anteil 100 %.
        k, anteile_anfang, basisfehler = core.buch_basis(
            [b for _, b in gruppe])

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

        def name(e):
            return " · ".join(e.anzeigename.split("\n"))

        def stueck(var):
            """Ganze Stückzahl; „1.000“ ist tausend, nicht eins."""
            return core._ganzzahl(var.get().strip().replace(".", ""))

        def zeige(person):
            w = self._zeige_empfaenger(e=person, eltern=fenster)
            if w is None:
                return
            _sperre_eingabe(w)
            fenster.wait_window(w)
            if fenster.winfo_exists():
                _sperre_eingabe(fenster)

        # --- Hinweise aus dem Import ------------------------------------
        # „bitte prüfen“ stand in der Liste, aber WAS zu prüfen ist, stand
        # nur in einer Spalte der Excel-Datei. Hier steht es, und mit dem
        # Haken verschwindet es — Infos ohne Handlungsbedarf bleiben stehen.
        v_erledigt = BooleanVar(value=False)
        hinweise = [(e, b, h) for e, b in gruppe for h in b.nachpflege]
        if hinweise:
            hinw = ttk.LabelFrame(fenster, text="Hinweise aus dem Import")
            hinw.pack(fill="x", **PAD)
            irgendwas_pruefen = False
            for e, b, h in hinweise:
                pruefen = h in core.pruefhinweise(b)
                irgendwas_pruefen |= pruefen
                wer = f"{name(e)}: " if len(gruppe) > 1 else ""
                ttk.Label(hinw, text=("⚠ " if pruefen else "ℹ ") + wer + h,
                          wraplength=900, justify="left").pack(
                              anchor="w", padx=8, pady=1)
            if irgendwas_pruefen:
                ttk.Checkbutton(hinw, text="geprüft — die ⚠-Hinweise "
                                           "entfernen", variable=v_erledigt
                                ).pack(anchor="w", padx=8, pady=(4, 6))

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
        geld = ttk.LabelFrame(fenster, text="Was das Buch je Exemplar an "
                                            "Honorar abwirft")
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
                           else dez(k.betrag_je_ex))
        feld(fest, 0, "Betrag in €", v_fest, 12)

        ttk.Radiobutton(geld, text="Ein Anteil am Verlagsabgabepreis",
                        variable=v_weg, value="anteil",
                        command=lambda: umschalten()).grid(
                            row=2, column=0, sticky="w", padx=8, pady=(10, 0))
        anteil = ttk.Frame(geld)
        anteil.grid(row=3, column=0, sticky="w", padx=32)
        v_preis = StringVar(value="" if k.ladenpreis is None
                            else dez(k.ladenpreis))
        v_mwst = StringVar(value="keine" if not k.mwst_im_preis
                           else dez(k.mwst_im_preis) + " %")
        v_rabatt = StringVar(value=dez(k.verlagsrabatt)
                             if k.rabatt_anwenden else "")
        feld(anteil, 0, "Ladenpreis in €", v_preis, 12)
        feld(anteil, 1, "darin enthaltene MwSt", v_mwst, 12,
             werte=["keine", "7 %", "19 %"])
        feld(anteil, 2, "Verlagsrabatt in % (meist 40)", v_rabatt, 12)

        # --- Der Honorarsatz, gleich ob fest oder gestaffelt ------------
        # Er gehört zum Anteil am Verlagsabgabepreis und steht deshalb im
        # selben Kasten darunter. Als eigener Kasten sah es aus, als gälte
        # bei einem festen Betrag zusätzlich noch ein Prozentsatz — zumal
        # „Immer derselbe Satz“ dort angewählt erschien.
        satzrahmen = ttk.Frame(geld)
        satzrahmen.grid(row=4, column=0, sticky="w", padx=32, pady=(6, 0))
        ttk.Label(satzrahmen, text="Honorarsatz — gleichbleibend oder nach "
                                   "Menge gestaffelt").grid(
                                       row=0, column=0, sticky="w", pady=(4, 0))
        satz_gemerkt = "staffel" if k.staffel else "fest"
        v_satzart = StringVar(value=satz_gemerkt)
        r_einfach = ttk.Radiobutton(satzrahmen, text="Immer derselbe Satz",
                                    variable=v_satzart, value="fest",
                                    command=lambda: umschalten())
        r_einfach.grid(row=1, column=0, sticky="w", padx=8, pady=(4, 0))
        einfach = ttk.Frame(satzrahmen)
        einfach.grid(row=2, column=0, sticky="w", padx=32)
        v_satz = StringVar(value="" if k.satz is None else dez(k.satz))
        feld(einfach, 0, "Satz in %", v_satz, 10)

        r_staffel = ttk.Radiobutton(
            satzrahmen, text="Gestaffelt — der Satz steigt mit der Menge",
            variable=v_satzart, value="staffel", command=lambda: umschalten())
        r_staffel.grid(row=3, column=0, sticky="w", padx=8, pady=(10, 0))
        stufen = ttk.Frame(satzrahmen)
        stufen.grid(row=4, column=0, sticky="w", padx=32)
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
                          "" if satz is None else dez(satz))
        while len(stufenzeilen) < 2:
            stufe_anlegen()
        ttk.Button(stufen, text="+ Stufe",
                   command=lambda: (stufe_anlegen(), vorschau())).grid(
                       row=99, column=1, sticky="w", pady=(6, 0))
        ttk.Label(satzrahmen, justify="left", foreground="#555555", text=(
            "Die letzte Stufe ohne Mengenangabe gilt nach oben offen. Wird im "
            "Lauf des Jahres eine\nGrenze überschritten, gilt der höhere Satz "
            "für die Exemplare darüber.")
        ).grid(row=5, column=0, sticky="w", padx=8, pady=(4, 8))

        # --- Autor bzw. Mitautoren, rechts daneben ----------------------
        # Ein Autor: nur sein Name und der Knopf zu ihm. Mehrere: wer wie
        # viel Prozent vom Honorar des Buches bekommt. So steht das Geld
        # einmal fürs Buch da, und die Aufteilung daneben — statt dass
        # jeder Eintrag seinen fertigen Teil trägt und niemand die Summe
        # sieht (bei „Worte für Orte“ zahlte der Verlag so unbemerkt das
        # Doppelte).
        autoren = [{"e": e, "buch": b, "neu": b not in e.buecher,
                    "anteil": StringVar(value=dez(a) if anteile_anfang
                                        else "")}
                   for (e, b), a in zip(gruppe, anteile_anfang
                                        or [0.0] * len(gruppe))]
        anteile_text_anfang = [z["anteil"].get() for z in autoren]
        rechts = ttk.Frame(geld)
        rechts.grid(row=0, column=1, rowspan=5, sticky="ne", padx=(16, 8),
                    pady=(6, 8))
        geld.columnconfigure(1, weight=1)
        anzeige: dict = {}

        def mitautor_hinzu():
            e = self._waehle_empfaenger(eltern=fenster)
            if e is None:
                return
            if any(z["e"] is e for z in autoren):
                messagebox.showinfo("Schon dabei",
                                    f"{name(e)} steht schon bei diesem Buch.",
                                    parent=fenster)
                return
            # Die Kennung erst beim Übernehmen: zwei neue Einträge bekämen
            # sonst dieselbe, weil keiner von beiden schon im Bestand steht.
            autoren.append({"e": e, "buch": core.Buch(kennung=""),
                            "neu": True, "anteil": StringVar()})
            # Gleichmäßig verteilen — angezeigt wird es sofort, und wer es
            # anders vereinbart hat, trägt es um.
            for z in autoren:
                z["anteil"].set(dez(core.runde(100 / len(autoren), 2)))
            zeichne_autoren()
            vorschau()

        def zeichne_autoren():
            for w in rechts.winfo_children():
                w.destroy()
            anzeige.clear()
            mehrere = (len(autoren) > 1 or basisfehler
                       or core._komma(autoren[0]["anteil"].get()) != 100)
            if not mehrere:
                kasten = ttk.LabelFrame(rechts, text="Autor")
                kasten.pack(anchor="ne")
                ttk.Label(kasten, text=name(autoren[0]["e"]),
                          font=("Segoe UI", 9, "bold")).grid(
                              row=0, column=0, sticky="w", padx=8, pady=6)
                ttk.Button(kasten, text="öffnen", width=7,
                           command=lambda: zeige(autoren[0]["e"])).grid(
                               row=0, column=1, padx=(6, 8))
                ttk.Button(kasten, text="+ Mitautor …",
                           command=mitautor_hinzu).grid(
                               row=1, column=0, columnspan=2, sticky="w",
                               padx=8, pady=(0, 8))
                return
            kasten = ttk.LabelFrame(rechts, text="Mitautoren — Anteil am "
                                                 "Honorar des Buches")
            kasten.pack(anchor="ne")
            for i, z in enumerate(autoren):
                schrift = {"font": ("Segoe UI", 9, "bold")} if i == 0 else {}
                ttk.Label(kasten, text=name(z["e"]), **schrift).grid(
                    row=i, column=0, sticky="w", padx=(8, 6), pady=2)
                ein = ttk.Entry(kasten, textvariable=z["anteil"], width=7,
                                justify="right")
                ein.grid(row=i, column=1, sticky="e")
                ein.bind("<KeyRelease>", lambda _e: vorschau())
                ttk.Label(kasten, text="%").grid(row=i, column=2, sticky="w",
                                                 padx=(2, 8))
                anzeige[id(z)] = StringVar()
                ttk.Label(kasten, textvariable=anzeige[id(z)]).grid(
                    row=i, column=3, sticky="e", padx=(0, 8))
                ttk.Button(kasten, text="öffnen", width=7,
                           command=lambda x=z["e"]: zeige(x)).grid(
                               row=i, column=4, padx=(0, 8))
            n = len(autoren)
            ttk.Separator(kasten).grid(row=n, column=0, columnspan=5,
                                       sticky="ew", padx=8, pady=(4, 2))
            anzeige["summe"] = StringVar()
            anzeige["summe_label"] = ttk.Label(
                kasten, textvariable=anzeige["summe"],
                font=("Segoe UI", 9, "bold"))
            anzeige["summe_label"].grid(row=n + 1, column=0, columnspan=4,
                                        sticky="w", padx=8)
            ttk.Button(kasten, text="+ Mitautor …",
                       command=mitautor_hinzu).grid(
                           row=n + 1, column=4, padx=(0, 8), pady=(0, 4))
            anzeige["fehler"] = StringVar(value=basisfehler)
            ttk.Label(kasten, textvariable=anzeige["fehler"],
                      foreground="#8a1c1c", wraplength=360,
                      justify="left").grid(row=n + 2, column=0, columnspan=5,
                                           sticky="w", padx=8, pady=(2, 6))
        zeichne_autoren()

        # --- Sonderregeln ------------------------------------------------
        sonder = ttk.LabelFrame(fenster, text="Sonderregeln")
        sonder.pack(fill="x", **PAD)
        fuer = f" ({name(erster_e)})" if len(autoren) > 1 else ""
        v_mwstpflicht = BooleanVar(value=buch.mwst_pflichtig)
        v_schwelle = BooleanVar(value=k.schwelle_zehn)
        # Die Freimenge wird so eingegeben, wie sie im Vertrag steht: „ab dem
        # 201. Exemplar“. Dazu gehört der Stand — ohne ihn weiß das Werkzeug
        # nicht, wie weit es bis zur Schwelle noch ist, und zahlt im ersten
        # Jahr ab dem ersten Exemplar.
        v_ab = StringVar(value=str(k.freimenge + 1) if k.freimenge else "")
        bisher_alt = core.freimenge_stand(buch, jahr)
        v_bisher = StringVar(
            value=str(bisher_alt) if bisher_alt is not None
            else ("0" if not buch.jahre else ""))
        frei_info = StringVar()
        # Ein Haken davor, damit sichtbar ist, OB die Regel gilt — ein leeres
        # Feld allein sieht aus wie „vergessen auszufüllen“.
        v_frei_an = BooleanVar(value=bool(k.freimenge))
        v_voraus_an = BooleanVar(value=bool(buch.vorauszahlung))
        ttk.Checkbutton(sonder, text="Autor ist mehrwertsteuerpflichtig" + fuer,
                        variable=v_mwstpflicht).grid(row=0, column=0,
                                                     sticky="w", padx=8, pady=2)
        ttk.Checkbutton(sonder,
                        text="Kein Honorar unter zehn Exemplaren im Jahr",
                        variable=v_schwelle).grid(row=1, column=0, sticky="w",
                                                  padx=8, pady=2)
        frei = ttk.Frame(sonder)
        frei.grid(row=2, column=0, sticky="w", padx=8, pady=2)
        ttk.Checkbutton(frei, text="Honorar erst ab dem … verkauften Exemplar",
                        variable=v_frei_an,
                        command=lambda: vorschau()).grid(
                            row=0, column=0, sticky="w", pady=3)
        w_ab = ttk.Entry(frei, textvariable=v_ab, width=10)
        w_ab.grid(row=0, column=1, sticky="w", padx=(8, 0), pady=3)
        ttk.Label(frei, text=f"davon bis Ende {jahr - 1} schon verkauft").grid(
            row=1, column=0, sticky="e", pady=3)
        w_bisher = ttk.Entry(frei, textvariable=v_bisher, width=10)
        w_bisher.grid(row=1, column=1, sticky="w", padx=(8, 0), pady=3)
        ttk.Label(frei, textvariable=frei_info, foreground="#555555").grid(
            row=1, column=2, sticky="w", padx=(8, 0))

        # Ein Buch aus der Abrechnung nehmen, ohne es zu verlieren. Das ist
        # der haeufigste Grund, warum ein Titel verschwindet — vergriffen,
        # Autor verstorben, unbekannt verzogen —, und bisher ging es nur von
        # Hand in der Excel.
        v_still = BooleanVar(value=buch.stillgelegt)
        v_grund = StringVar(value=buch.stillgelegt_grund)
        v_gesondert = BooleanVar(value=buch.gesondert)
        # Eine Vorauszahlung schmilzt nicht von selbst ab: das Werkzeug
        # verrechnet sie beim Rechnen und schreibt in den Hinweis, was
        # danach noch offen ist — eingetragen wird der neue Restbetrag hier
        # von Hand. Automatisch zurueckschreiben waere bequemer und falsch:
        # zweimal rechnen wuerde die Vorauszahlung zweimal abziehen.
        v_voraus = StringVar(value=dez(buch.vorauszahlung)
                             if buch.vorauszahlung else "")
        vor = ttk.Frame(sonder)
        vor.grid(row=3, column=0, sticky="w", padx=8, pady=2)
        ttk.Checkbutton(vor, text="noch offene Vorauszahlung (€)" + fuer
                        + " — wird verrechnet, bevor ausgezahlt wird",
                        variable=v_voraus_an,
                        command=lambda: vorschau()).grid(
                            row=0, column=0, sticky="w", pady=3)
        w_voraus = ttk.Entry(vor, textvariable=v_voraus, width=10)
        w_voraus.grid(row=0, column=1, sticky="w", padx=(8, 0), pady=3)

        ttk.Separator(sonder, orient="horizontal").grid(
            row=4, column=0, sticky="ew", padx=8, pady=(8, 4))
        ttk.Checkbutton(
            sonder, text="Wird nicht mehr abgerechnet (vergriffen, "
                         "verstorben, kein Honorar mehr)",
            variable=v_still, command=lambda: grund_umschalten()).grid(
                row=5, column=0, sticky="w", padx=8, pady=2)
        grundrahmen = ttk.Frame(sonder)
        grundrahmen.grid(row=6, column=0, sticky="w", padx=32, pady=2)
        w_grund = feld(grundrahmen, 0, "Grund", v_grund, 40)
        ttk.Checkbutton(
            sonder, text="Sonderfall: Honorar entscheidet der Verlag von "
                         "Hand — die Zahlen laufen mit, ausgezahlt wird "
                         "nichts automatisch",
            variable=v_gesondert).grid(row=7, column=0, sticky="w",
                                       padx=8, pady=2)

        def grund_umschalten():
            w_grund.configure(state="normal" if v_still.get() else "disabled")
        grund_umschalten()

        # --- Vorschau ----------------------------------------------------
        ergebnis = StringVar()
        ttk.Label(fenster, textvariable=ergebnis, font=("Segoe UI", 10, "bold"),
                  foreground="#1a5c1a", wraplength=900, justify="left"
                  ).pack(anchor="w", padx=12, pady=(0, 4))

        def lies() -> tuple:
            """Die Eingaben übersetzen. (Kondition des Buches,
            Vorauszahlung, Fehler)"""
            kond = core.Kondition(satz_runden=k.satz_runden,
                                  mwst_aufschlagen=k.mwst_aufschlagen,
                                  freimenge_ab_jahr=k.freimenge_ab_jahr)
            fehler = []

            def zahl(var, feldname, vorgabe=None):
                """Leer heißt Vorgabe — unlesbar heißt Fehler.

                Beides in einen Topf zu werfen ist hier teuer: ein vertipptes
                „4o“ im Verlagsrabatt hiesse „gar kein Rabatt“, und der Autor
                bekaeme auf einen Schlag zwei Drittel mehr. Ohne Meldung.
                """
                text = var.get().strip()
                if not text:
                    return vorgabe
                wert = core._komma(text)
                if wert is None:
                    fehler.append(f"„{text}“ kann ich bei {feldname} nicht "
                                  f"als Zahl lesen. Bitte nur Ziffern, "
                                  f"Komma erlaubt.")
                    return vorgabe
                return wert

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
                rabatt = zahl(v_rabatt, "Verlagsrabatt")
                kond.rabatt_anwenden = rabatt is not None
                kond.verlagsrabatt = rabatt if rabatt is not None else 40.0
                if v_satzart.get() == "fest":
                    satz = core._komma(v_satz.get())
                    if satz is None:
                        fehler.append("Es fehlt der Honorarsatz.")
                    kond.satz = satz
                else:
                    staffel = []
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
            # Ohne Haken gilt die Regel nicht — was im ausgegrauten Feld
            # noch steht, wird ignoriert.
            kond.freimenge = 0
            if v_frei_an.get():
                ab = stueck(v_ab)
                if ab is None or ab < 2:
                    fehler.append("Bitte eintragen, ab dem wievielten "
                                  "Exemplar Honorar gezahlt wird (mindestens "
                                  "2) — oder den Haken davor entfernen.")
                else:
                    kond.freimenge = ab - 1
                    if stueck(v_bisher) is None:
                        fehler.append(
                            f"Bitte eintragen, wie viele Exemplare bis Ende "
                            f"{jahr - 1} schon verkauft waren — bei einem "
                            f"neuen Buch 0.")
            voraus = 0.0
            if v_voraus_an.get():
                voraus = zahl(v_voraus, "Vorauszahlung", None)
                if voraus is None:
                    fehler.append("Bitte den offenen Betrag der Vorauszahlung "
                                  "eintragen — oder den Haken davor "
                                  "entfernen.")
                    voraus = 0.0
                elif voraus < 0:
                    fehler.append("Eine Vorauszahlung kann nicht negativ "
                                  "sein.")
            return kond, voraus, fehler

        def lies_anteile() -> tuple:
            """(Anteile in %, Fehler) — die Summe muss 100 % ergeben."""
            werte = [core._komma(z["anteil"].get()) for z in autoren]
            if any(w is None or w < 0 for w in werte):
                return None, ("Bitte bei jedem Autor seinen Anteil in % "
                              "eintragen.")
            summe = sum(werte)
            if abs(summe - 100) > 0.01:
                return werte, (f"Die Anteile ergeben zusammen "
                               f"{dez(core.runde(summe, 2))} % statt 100 %.")
            return werte, ""

        def freimenge_zeigen():
            ab, bisher = stueck(v_ab), stueck(v_bisher)
            an = v_frei_an.get()
            w_ab.configure(state="normal" if an else "disabled")
            w_bisher.configure(
                state="normal" if an and ab and ab > 1 else "disabled")
            w_voraus.configure(
                state="normal" if v_voraus_an.get() else "disabled")
            if not an or not ab or ab < 2:
                frei_info.set("")
            elif bisher is None:
                frei_info.set("")
            elif bisher >= ab - 1:
                frei_info.set("→ Schwelle erreicht, jedes weitere Exemplar "
                              "wird vergütet")
            else:
                frei_info.set(f"→ noch {ab - 1 - bisher} Exemplare bis zum "
                              f"Honorar")

        def je_exemplar(kond) -> str:
            if kond.staffel:
                teile, untere = [], 1
                for grenze, satz in kond.staffel:
                    kopie = core.Kondition(**{**kond.__dict__, "satz": satz,
                                              "staffel": []})
                    betrag = core.satz_aus_kondition(kopie)
                    if grenze is None:
                        teile.append(f"ab {untere}: {core.euro(betrag)}")
                    else:
                        teile.append(f"{untere}–{grenze}: {core.euro(betrag)}")
                        untere = grenze + 1
                return ", ".join(teile)
            return core.euro(core.satz_aus_kondition(kond))

        def vorschau(*_):
            freimenge_zeigen()
            kond, voraus, fehler = lies()
            anteile, anteilfehler = lies_anteile()
            if "summe" in anzeige:
                summe = sum(a for a in (anteile or []) if a is not None)
                anzeige["summe"].set(f"zusammen {dez(core.runde(summe, 2))} %")
                anzeige["summe_label"].configure(
                    foreground="#8a1c1c" if anteilfehler else "#1a5c1a")
                for z, a in zip(autoren, anteile or [None] * len(autoren)):
                    if fehler or a is None:
                        anzeige[id(z)].set("")
                        continue
                    anteilig = core.Kondition(**{**kond.__dict__,
                                                 "anteil": a, "teiler": 1})
                    anzeige[id(z)].set("= " + (
                        core.euro(core.satz_aus_kondition(anteilig))
                        if not anteilig.staffel else
                        core.euro(core.satz_aus_kondition(
                            core.Kondition(**{**anteilig.__dict__,
                                              "satz": anteilig.staffel[0][1],
                                              "staffel": []})))
                        + " …"))
            if fehler:
                ergebnis.set("… " + fehler[0])
                return
            if kond.staffel:
                ergebnis.set("Das Buch ergibt je Exemplar — "
                             + je_exemplar(kond))
            else:
                ergebnis.set("Das Buch ergibt " + je_exemplar(kond)
                             + " je Exemplar.")

        def umschalten():
            nonlocal satz_gemerkt
            anteilig = v_weg.get() == "anteil"
            # Bei festem Betrag gibt es keinen Satz: dann ist auch keiner
            # der beiden Knöpfe angewählt. Die Wahl wird gemerkt und kommt
            # zurück, sobald wieder „Anteil“ gewählt ist.
            if anteilig:
                if not v_satzart.get():
                    v_satzart.set(satz_gemerkt)
            elif v_satzart.get():
                satz_gemerkt = v_satzart.get()
                v_satzart.set("")
            for w in (r_einfach, r_staffel):
                w.configure(state="normal" if anteilig else "disabled")
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

        for var in (v_fest, v_preis, v_rabatt, v_satz, v_ab, v_bisher):
            var.trace_add("write", lambda *_: vorschau())
        v_mwst.trace_add("write", lambda *_: vorschau())
        umschalten()

        # --- Knöpfe ------------------------------------------------------
        # Was für alle Autoren gilt, wird nur geschrieben, wenn es hier
        # geändert wurde. Sonst überschriebe das Öffnen eines Buchs still die
        # Angaben eines Mitautors, die zufällig anders stehen — etwa eine
        # eigene Notiz zu dessen Vertrag.
        def textstand():
            return {"titel": v_titel.get(), "isbn": v_isbn.get(),
                    "art": v_art.get(), "notiz": v_notiz.get(),
                    "still": (v_still.get(), v_grund.get()),
                    "gesondert": v_gesondert.get()}

        def geldstand():
            return (v_weg.get(), v_fest.get(), v_preis.get(), v_mwst.get(),
                    v_rabatt.get(), v_satzart.get(), v_satz.get(),
                    tuple((a.get(), b.get()) for a, b in stufenzeilen),
                    v_schwelle.get(), v_frei_an.get(), v_ab.get(),
                    v_bisher.get())

        text_anfang, geld_anfang = textstand(), geldstand()
        fertig = {"ok": False}

        def uebernehmen():
            if not v_titel.get().strip():
                messagebox.showwarning("Titel fehlt",
                                       "Ohne Buchtitel geht es nicht.",
                                       parent=fenster)
                return
            kond, voraus, fehler = lies()
            # Die Kondition wird nur neu geschrieben, wenn am Geld oder an
            # den Anteilen etwas geändert wurde. Dann aber für alle Autoren
            # — und dann müssen die Anteile stimmen.
            anteile, anteilfehler = lies_anteile()
            neu_schreiben = (geldstand() != geld_anfang
                             or [z["anteil"].get() for z in autoren]
                             != anteile_text_anfang
                             or any(z["neu"] for z in autoren))
            if neu_schreiben and anteilfehler:
                fehler.append(anteilfehler)
            if fehler:
                messagebox.showwarning("Bitte noch ergänzen",
                                       "\n".join(fehler), parent=fenster)
                return
            if neu_schreiben and basisfehler and not messagebox.askokcancel(
                    "Angaben vereinheitlichen",
                    "Die Einträge der Autoren rechneten bisher auf "
                    "unterschiedlichen Grundlagen. Mit „OK“ gilt für alle "
                    "das Honorar des Buches von hier, aufgeteilt nach den "
                    "eingetragenen Anteilen.", parent=fenster):
                return
            jetzt = textstand()

            def geaendert(schluessel):
                return jetzt[schluessel] != text_anfang[schluessel]

            for i, z in enumerate(autoren):
                b, neu = z["buch"], z["neu"]
                if neu or geaendert("titel"):
                    b.titel = v_titel.get().strip()
                if neu or geaendert("isbn"):
                    b.isbn = v_isbn.get().strip()
                if neu or geaendert("art"):
                    b.verguetungsart = v_art.get() or "Honorar"
                if neu or geaendert("notiz"):
                    b.notizen = v_notiz.get().strip()
                if neu or geaendert("still"):
                    b.stillgelegt = v_still.get()
                    b.stillgelegt_grund = (v_grund.get().strip()
                                           if v_still.get() else "")
                if neu or geaendert("gesondert"):
                    b.gesondert = v_gesondert.get()
                if i == 0:
                    # MwSt-Pflicht und Vorauszahlung gehören zum Autor, nicht
                    # zum Buch: sie gelten für den, von dem aus geöffnet wurde.
                    b.mwst_pflichtig = v_mwstpflicht.get()
                    b.vorauszahlung = voraus
                if neu_schreiben:
                    alte_freimenge = b.kondition.freimenge
                    b.kondition = core.Kondition(**{
                        **kond.__dict__, "staffel": list(kond.staffel),
                        "teiler": 1, "anteil": anteile[i]})
                    if kond.freimenge:
                        bisher = stueck(v_bisher)
                        if (neu or bisher != bisher_alt
                                or kond.freimenge != alte_freimenge):
                            core.setze_freimenge_stand(b, jahr, bisher)
                if v_erledigt.get():
                    weg = set(core.pruefhinweise(b))
                    b.nachpflege = [h for h in b.nachpflege if h not in weg]
                if neu:
                    if not b.kennung:
                        b.kennung = self.bestand.naechste_kennung("B")
                    z["e"].buecher.append(b)
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
        # Zweimal sortieren, weil Pythons Sortierung stabil ist: erst nach
        # Wert (auf Wunsch rückwärts), dann nach Gruppe. So bleiben leere
        # Zellen in BEIDE Richtungen unten — sonst stünden sie beim
        # Rückwärtssortieren ganz oben und verdeckten genau die größten
        # Beträge, die man sucht.
        zeilen.sort(key=lambda z: _sortwert(z[0])[1:], reverse=umgekehrt)
        zeilen.sort(key=lambda z: _sortwert(z[0])[0])
        for i, (_, k) in enumerate(zeilen):
            baum.move(k, "", i)

    # --- Reiter 1: Stammdaten -----------------------------------------

    def _baue_stammdaten(self):
        seite = ttk.Frame(self.reiter)
        self.reiter.add(seite, text="1. Autoren")
        self._anleitung(seite, (
            "Hier stehen alle Autoren mit Anschriften und Bankverbindungen. "
            "Beim allerersten Mal holen Sie die Daten mit dem Knopf rechts "
            "aus der alten Excel-Tabelle. Danach müssen Sie hier nur noch "
            "etwas tun, wenn ein Autor umzieht oder dazukommt — Bücher "
            "stehen auf Reiter 2."))

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
        ttk.Button(oben, text="Autor entfernen …",
                   command=self._entferne_gewaehlten_empfaenger).pack(
                       side="left", padx=(12, 0))

        self.baum_stamm = self._liste(seite, SPALTEN_STAMM,
                                      doppelklick=self._zeige_empfaenger)
        ttk.Label(seite, text="Doppelklick auf eine Zeile zeigt Anschrift, "
                              "Bücher und die vereinbarte Vergütung."
                  ).pack(anchor="w", padx=12)
        self._weiter(seite, 1, "Weiter zu Schritt 2: Bücher  ▸")

    # --- Reiter 2: Bücher ---------------------------------------------

    def _baue_buecher(self):
        seite = ttk.Frame(self.reiter)
        self.reiter.add(seite, text="2. Bücher")
        self._anleitung(seite, (
            "Hier stehen alle Bücher mit der vereinbarten Vergütung. "
            "Doppelklick auf ein Buch ändert seine Angaben — Preis, Satz, ab "
            "dem wievielten Exemplar Honorar gezahlt wird. Ein neues Buch "
            "legen Sie mit „Neues Buch …“ an; gefragt wird zuerst, zu "
            "welchem Autor es gehört."))

        oben = ttk.Frame(seite)
        oben.pack(fill="x", **PAD)
        ttk.Label(oben, text="Suchen:").pack(side="left")
        feld = ttk.Entry(oben, textvariable=self.suche_buecher, width=36)
        feld.pack(side="left", padx=6)
        feld.bind("<KeyRelease>", lambda e: self._zeige_buecher())
        ttk.Button(oben, text="Neues Buch …",
                   command=self._neues_buch_waehlen).pack(side="left",
                                                          padx=(16, 4))
        ttk.Button(oben, text="Buch ändern …",
                   command=self._aendere_gewaehltes_buch).pack(side="left")
        ttk.Button(oben, text="Buch entfernen …",
                   command=self._entferne_gewaehltes_buch).pack(
                       side="left", padx=(12, 0))
        # Gut hundert stillgelegte Titel: sichtbar nur auf Wunsch, sonst
        # gehen die laufenden darin unter.
        ttk.Checkbutton(oben, text="auch Bücher zeigen, die nicht mehr "
                                   "abgerechnet werden",
                        variable=self.zeige_stillgelegte,
                        command=self._zeige_buecher).pack(side="right")

        self.baum_buecher = self._liste(
            seite, SPALTEN_BUECHER,
            doppelklick=lambda _e: self._aendere_gewaehltes_buch())
        self.baum_buecher.tag_configure("still", foreground="#888888")
        ttk.Label(seite, text="Doppelklick auf ein Buch ändert seine "
                              "Angaben.").pack(anchor="w", padx=12)
        self._weiter(seite, 2, "Weiter zu Schritt 3: Zahlen eintragen  ▸")

    # --- Reiter 3: Jahreserfassung ------------------------------------

    def _baue_erfassung(self):
        seite = ttk.Frame(self.reiter)
        self.reiter.add(seite, text="3. Zahlen eintragen")
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
        self._weiter(seite, 3, "Weiter zu Schritt 4: Nachrechnen  ▸")

    # --- Reiter 4: Durchlauf ------------------------------------------

    def _baue_durchlauf(self):
        seite = ttk.Frame(self.reiter)
        self.reiter.add(seite, text="4. Nachrechnen")
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
        self._weiter(seite, 4, "Weiter zu Schritt 5: Briefe und Listen  ▸")

    # --- Reiter 5: Ausgaben -------------------------------------------

    def _baue_ausgaben(self):
        seite = ttk.Frame(self.reiter)
        self.reiter.add(seite, text="5. Briefe und Listen")
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

    # --- Reiter 6: Meldungen ------------------------------------------

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
        """Zum Meldungsreiter wechseln (der Reiter nach den Schritten)."""
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
            offen = sum(1 for b in e.buecher if core.pruefhinweise(b))
            if offen:
                hinweise.append("1 Buch bitte prüfen" if offen == 1
                                else f"{offen} Bücher bitte prüfen")
            hinweis = " · ".join(hinweise)
            baum.insert("", "end", iid=e.kennung, values=(
                e.anzeigename, f"{e.plz} {e.ort}".strip(), e.iban,
                len(e.buecher), hinweis))
        # Jede Änderung an Autoren oder Büchern zeichnet diese Liste neu —
        # die Bücherliste hängt sich daran, statt an zehn Stellen gerufen
        # zu werden.
        if hasattr(self, "baum_buecher"):
            self._zeige_buecher()

    # -----------------------------------------------------------------
    # Reiter 2: Bücher anzeigen
    # -----------------------------------------------------------------

    def _zeige_buecher(self):
        baum = self.baum_buecher
        auswahl = baum.selection()
        for k in baum.get_children(""):
            baum.delete(k)
        suche = self.suche_buecher.get().strip().lower()
        alle = self.zeige_stillgelegte.get()
        # Ein Buch mit mehreren Autoren steht im Bestand einmal je Autor —
        # in der Liste aber nur einmal, mit allen Namen. Die Zeile trägt die
        # Kennung des ersten Eintrags; der Dialog holt die übrigen dazu.
        gruppen: dict[str, list] = {}
        for e, b in self.bestand.buecher():
            gruppen.setdefault(self._buchschluessel(b) or b.kennung,
                               []).append((e, b))
        for gruppe in gruppen.values():
            still = all(b.stillgelegt for _, b in gruppe)
            if still and not alle:
                continue
            e, b = gruppe[0]
            namen = " / ".join(" · ".join(x.anzeigename.split("\n"))
                               for x, _ in gruppe)
            if suche and suche not in (b.titel + " " + b.isbn + " "
                                       + namen).lower():
                continue
            betraege = " / ".join(core.euro(core.satz_aus_kondition(
                y.kondition)) for _, y in gruppe)
            baum.insert("", "end", iid=b.kennung, values=(
                b.isbn, b.titel, namen, b.verguetungsart, betraege,
                core.besonderheit(b)),
                tags=("still",) if still else ())
        # Nach dem Ändern steht die Liste neu da; ohne das hier wäre das
        # gerade bearbeitete Buch aus dem Blick.
        if auswahl and baum.exists(auswahl[0]):
            baum.selection_set(auswahl[0])
            baum.see(auswahl[0])

    def _gewaehltes_buch(self):
        auswahl = self.baum_buecher.selection()
        if not auswahl:
            messagebox.showinfo(
                "Erst ein Buch auswählen",
                "Bitte in der Liste die Zeile anklicken, um die es geht.")
            return None
        return next(((e, b) for e, b in self.bestand.buecher()
                     if b.kennung == auswahl[0]), None)

    def _aendere_gewaehltes_buch(self):
        paar = self._gewaehltes_buch()
        if paar is not None:
            self._aendere_buch(*paar)

    def _entferne_gewaehltes_buch(self):
        paar = self._gewaehltes_buch()
        if paar is None:
            return
        gruppe = self._buchgruppe(*paar)
        if len(gruppe) == 1:
            self._entferne_buch(*paar)
            return
        # Für ein Buch mit mehreren Autoren EINE Rückfrage, nicht drei.
        namen = "\n".join(f"  · {e.anzeigename}" for e, _ in gruppe)
        jahre = sorted({j for _, b in gruppe for j, w in b.jahre.items()
                        if w.erfasst})
        verlust = (f"\n\nDamit gehen die erfassten Zahlen aus {len(jahre)} "
                   f"Jahren unwiderruflich verloren." if jahre else "")
        if not messagebox.askokcancel(
                "Buch entfernen",
                f"„{paar[1].titel}“ gehört {len(gruppe)} Autoren:\n{namen}"
                f"\n\nBei allen entfernen?{verlust}\n\nSoll der Titel nur "
                f"nicht mehr abgerechnet werden, ist „Buch ändern“ → „Wird "
                f"nicht mehr abgerechnet“ der richtige Weg.",
                icon="warning", default="cancel"):
            return
        for e, b in gruppe:
            e.buecher.remove(b)
        self.geaendert = True
        self._rechnung_veraltet = bool(self.abrechnungen)
        self._zeige_stammdaten()
        self._zeige_erfassung()
        self._male_schrittleiste()
        self.status.set(f"„{paar[1].titel}“ entfernt — noch nicht "
                        f"gespeichert.")

    def _neues_buch_waehlen(self):
        """Neues Buch: erst den Autor wählen, dann die Angaben."""
        auswahl = self.baum_buecher.selection()
        vorschlag = next((e for e, b in self.bestand.buecher()
                          if auswahl and b.kennung == auswahl[0]), None)
        e = self._waehle_empfaenger(vorschlag)
        if e is None:
            return
        vorher = len(e.buecher)
        self._neues_buch(e)
        if len(e.buecher) > vorher:
            neu = e.buecher[-1].kennung
            if self.baum_buecher.exists(neu):
                self.baum_buecher.selection_set(neu)
                self.baum_buecher.see(neu)

    def _waehle_empfaenger(self, vorschlag=None, eltern=None):
        """Ein kleines Fenster: zu welchem Autor gehört das neue Buch?

        Eine Liste mit Suchfeld statt einer Klappliste — bei 250 Namen
        findet man in einer Klappliste niemanden.
        """
        fenster = Toplevel(eltern or self)
        fenster.title("Zu welchem Autor gehört das Buch?")
        fenster.transient(eltern or self)
        _sperre_eingabe(fenster)
        rahmen = ttk.Frame(fenster)
        rahmen.pack(fill="both", expand=True, **PAD)
        such = StringVar()
        ttk.Label(rahmen, text="Suchen:").grid(row=0, column=0, sticky="w")
        eingabe = ttk.Entry(rahmen, textvariable=such, width=40)
        eingabe.grid(row=0, column=1, sticky="ew", padx=(6, 0))
        liste = Listbox(rahmen, width=60, height=18, exportselection=False)
        liste.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(8, 0))
        rahmen.rowconfigure(1, weight=1)
        rahmen.columnconfigure(1, weight=1)
        ttk.Label(rahmen, foreground="#555555", text=(
            "Steht der Autor noch nicht in der Liste, legen Sie ihn zuerst "
            "auf Reiter 1 mit „Neuer Autor …“ an.")).grid(
                row=2, column=0, columnspan=2, sticky="w", pady=(6, 0))

        alle = sorted(self.bestand.empfaenger,
                      key=lambda x: x.anzeigename.lower())
        sichtbar: list = []

        def fuelle(*_):
            text = such.get().strip().lower()
            sichtbar[:] = [x for x in alle
                           if not text or text in x.anzeigename.lower()]
            liste.delete(0, END)
            for x in sichtbar:
                # Mehrzeilige Institutionen zeigt die Listbox sonst als
                # Kästchen mitten im Namen.
                liste.insert(END, " · ".join(x.anzeigename.split("\n")))
            if sichtbar:
                i = sichtbar.index(vorschlag) if vorschlag in sichtbar else 0
                liste.selection_set(i)
                liste.see(i)
        such.trace_add("write", fuelle)
        fuelle()

        gewaehlt = {"e": None}

        def nehmen(_e=None):
            i = liste.curselection()
            if not i:
                return
            gewaehlt["e"] = sichtbar[i[0]]
            fenster.destroy()

        liste.bind("<Double-1>", nehmen)
        leiste = ttk.Frame(fenster)
        leiste.pack(fill="x", padx=12, pady=(0, 12))
        ttk.Button(leiste, text="Weiter", command=nehmen).pack(side="right")
        ttk.Button(leiste, text="Abbrechen",
                   command=fenster.destroy).pack(side="right", padx=6)
        fenster.bind("<Return>", nehmen)
        fenster.bind("<Escape>", lambda _e: fenster.destroy())
        eingabe.focus_set()
        self.wait_window(fenster)
        if eltern is not None and eltern.winfo_exists():
            _sperre_eingabe(eltern)     # die Sperre ging an dieses Fenster
        return gewaehlt["e"]

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

    def _zeige_empfaenger(self, ereignis=None, e=None, eltern=None):
        """Das Fenster eines Autors. Ohne ``e`` der in der Liste gewählte."""
        if e is None:
            auswahl = self.baum_stamm.selection()
            if not auswahl:
                return None
            e = next((x for x in self.bestand.empfaenger
                      if x.kennung == auswahl[0]), None)
        if e is None:
            return None
        fenster = Toplevel(eltern or self)
        if eltern is not None:
            fenster.transient(eltern)
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
        ttk.Button(knoepfe, text="Autor entfernen …",
                   command=lambda: (self._entferne_empfaenger(e)
                                    and fenster.destroy())).pack(side="left")

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
                self._aendere_buch(e, b, eltern=fenster)
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
            command=lambda: (self._neues_buch(e, eltern=fenster), fuelle())).pack(side="left",
                                                                 padx=6)

        def entfernen():
            b = gewaehltes_buch()
            if b is not None:
                self._entferne_buch(e, b, danach=fuelle, eltern=fenster)

        ttk.Button(buchknoepfe, text="Buch entfernen …",
                   command=entfernen).pack(side="left", padx=(12, 0))
        return fenster

    # -----------------------------------------------------------------
    # Reiter 3: Jahreserfassung
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
                "" if jw is None or not jw.korrektur else jw.korrektur,
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
        if not zeile or spalte not in ("#5", "#6", "#7"):
            return
        # Ohne see() liefert bbox für eine weggescrollte Zeile "" — das gäbe
        # einen Fehler, der im Fensterbau spurlos verschwände.
        baum.see(zeile)
        kasten = baum.bbox(zeile, spalte)
        if not kasten:
            return
        x, y, breite, hoehe = kasten
        feldname = {"#5": "verkauft", "#6": "eigenkauf",
                    "#7": "korrektur"}[spalte]
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
            elif feldname == "eigenkauf":
                jw.eigenkauf = zahl or 0
            else:
                jw.korrektur = zahl or 0
            if baum.exists(zeile):
                # Eine 0 in Eigenkauf oder Korrektur ist der Normalfall und
                # soll die Spalte nicht zupflastern.
                baum.set(zeile, feldname,
                         "" if zahl is None
                         or (not zahl and feldname != "verkauft") else zahl)
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
            # Tabulator geht die drei Zahlenspalten der Reihe nach durch.
            uebernehmen(False, {"#5": "#6", "#6": "#7", "#7": "#5"}[spalte])
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
    # Reiter 4: Durchlauf
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
        feld = self._rechenfeld(fenster)
        feld.pack(fill="both", expand=True, **PAD)
        feld.insert(END, f"{z['Empfänger']}\n")
        if buch is not None:
            posten = core.betrag_zeile(buch, self.jahr.get(), self.cfg)
            if posten is not None:
                self._schreibe_rechenweg(feld, posten)
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
        feld = self._rechenfeld(fenster)
        feld.pack(fill="both", expand=True, **PAD)
        for posten in ab.posten:
            self._schreibe_rechenweg(feld, posten)
            feld.insert(END, "\n")
        feld.insert(END, "─" * 62 + "\n")
        if len(ab.posten) > 1 or ab.mwst or ab.verrechnet:
            feld.insert(END, f"\tSumme netto\t{core.euro(ab.netto)}\n")
        if ab.mwst:
            feld.insert(END, f"\tMehrwertsteuer\t{core.euro(ab.mwst)}\n")
        if ab.verrechnet:
            feld.insert(END, f"\teinbehalten für den offenen Vorschuss\t"
                             f"− {core.euro(ab.verrechnet)}\n")
        feld.insert(END, f"\tAuszahlung\t{core.euro(ab.brutto)}\n", "summe")
        if ab.gruende:
            feld.insert(END, "\nEs geht kein Brief hinaus: "
                             + "; ".join(ab.gruende) + "\n")
        feld.configure(state=DISABLED)

    @staticmethod
    def _rechenfeld(fenster) -> Text:
        """Ein Textfeld für den Rechenweg: Beschriftung links eingerückt,
        Beträge an einem rechtsbündigen Tabulator — wie auf dem Papier."""
        feld = Text(fenster, wrap="word", font=("Segoe UI", 10),
                    tabs=("0.8c", "13c", "right"), spacing1=1)
        feld.tag_configure("titel", font=("Segoe UI", 10, "bold"),
                           spacing1=8, spacing3=4)
        feld.tag_configure("summe", font=("Segoe UI", 10, "bold"))
        # Ein langer Hinweis bricht sonst am linken Rand um statt unter
        # seinem eigenen Anfang.
        feld.tag_configure("notiz", lmargin2="0.8c")
        return feld

    @staticmethod
    def _schreibe_rechenweg(feld, posten) -> None:
        titel, *rest = core.rechenweg(posten)
        feld.insert(END, titel + "\n", "titel")
        for i, zeile in enumerate(rest):
            # Die Ergebniszeile eines Buches (die letzte mit Betrag) fett.
            if i == len(rest) - 1:
                marke = "summe"
            elif zeile.count("\t") == 1:
                marke = "notiz"
            else:
                marke = ()
            feld.insert(END, zeile + "\n", marke)

    # -----------------------------------------------------------------
    # Reiter 5: Ausgaben
    # -----------------------------------------------------------------

    def _bereit(self) -> bool:
        if not self.abrechnungen:
            messagebox.showinfo(
                "Erst die Beträge berechnen",
                "Bitte zuerst auf Reiter „4. Nachrechnen“ den Knopf "
                "„Beträge berechnen“ drücken.")
            return False
        if getattr(self, "_rechnung_veraltet", False):
            messagebox.showwarning(
                "Die Zahlen haben sich geändert",
                "Seit dem letzten Rechnen wurden Stückzahlen geändert.\n\n"
                "Bitte auf Reiter „4. Nachrechnen“ noch einmal „Beträge "
                "berechnen“ drücken — sonst stünden in den Briefen die alten "
                "Beträge.")
            self.reiter.select(3)
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
