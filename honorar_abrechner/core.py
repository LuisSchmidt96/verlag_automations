"""
Honorar-Abrechner Core-Logik
=============================
Reine Datenlogik ohne UI-Abhängigkeiten.

Der Verlag rechnet einmal im Januar die Autorenhonorare des Vorjahres ab.
Bisher steckte die gesamte Geschäftslogik in von Hand getippten Zellformeln
einer Excel-Mappe (``Honorare_<Jahr>.xlsx``, 10 Blätter, 142 Spalten). Dieses
Modul hält die Konditionen stattdessen als Daten und rechnet sie aus.

Der Bestand selbst bleibt bewusst eine Excel-Mappe (``Honorarbestand.xlsx``):
sie muss im Januar auch ohne das Werkzeug zu öffnen und notfalls von Hand
weiterzuführen sein. Die beiden Nachteile dieser Wahl — Excel sperrt die Datei,
und Typen verrutschen beim Speichern von Hand — werden hier ausdrücklich
abgefangen (siehe ``pruefe_sperre`` und die ``_text``/``_ganzzahl``-Helfer).
"""

from __future__ import annotations

import ast
import os
import json
import re
import shutil
import socket
import sys
from copy import deepcopy
from decimal import Decimal, ROUND_HALF_UP
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import openpyxl
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Emu, Pt
from docx.table import _Row
from docx.text.paragraph import Paragraph
from openpyxl.styles import Alignment, Font, PatternFill


# ---------------------------------------------------------------------
# Daten-Ordner (Config, Bestand, Ausgabe)
# ---------------------------------------------------------------------

def _base_dir() -> Path:
    """Ordner neben der .exe (PyInstaller-Build) bzw. neben dem Tool-Code."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


# Sämtliche Laufzeitdateien (config.json, Honorarbestand.xlsx sowie der
# Ausgabeordner honorar_output/) liegen direkt neben der .exe — bewusst KEIN
# data-Unterordner, damit die Stammdaten leicht zu finden sind.
APP_DIR = _base_dir()
APP_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_PFAD = APP_DIR / "config.json"
BESTAND_PFAD = APP_DIR / "Honorarbestand.xlsx"
SICHERUNG_PFAD = APP_DIR / "Honorarbestand.bak.xlsx"
SPERR_PFAD = APP_DIR / "honorarbestand.sperre"
# Zwischenstand während der Erfassung. Wer einen Vormittag lang Zahlen
# eintippt, soll einen Absturz oder einen Stromausfall überleben.
WIEDERHERSTELLUNG_PFAD = APP_DIR / "Honorarbestand.wiederherstellung.xlsx"


def _vorlagen_dir() -> Path:
    """Nur-Lese-Vorlagen (Briefvorlage).

    Im PyInstaller-Build liegen sie unter ``sys._MEIPASS`` (per ``datas``
    gebündelt), sonst neben diesem Modul.
    """
    if getattr(sys, "frozen", False):
        p = Path(getattr(sys, "_MEIPASS", "")) / "honorar_abrechner" / "vorlagen"
        if p.exists():
            return p
    return Path(__file__).parent / "vorlagen"


# ---------------------------------------------------------------------
# Konfiguration
# ---------------------------------------------------------------------

DEFAULT_CONFIG = {
    "output_dir": "honorar_output",
    "absender_ort": "Ubstadt-Weiher",
    "unterzeichner": "Silke Freitag",
    "mwst_satz": 7.0,
    "verlagsrabatt": 40.0,
    # Die Schwelle aus der Spalte „Keine Berechnung bei 10 o. weniger.“
    # Gezahlt wird AB dieser Menge; darunter entfällt das Honorar.
    "schwelle_menge": 10,
    # Konten für die Künstlersozialkasse-Liste (SKR03, s. PLAN_Buchhaltung.md).
    # Entscheidend ist NICHT die Höhe des Steuersatzes, sondern ob der Autor
    # überhaupt mehrwertsteuerpflichtig ist: wer es nicht ist, wird auf ein
    # eigenes Konto „Honorare" ohne Umsatzsteuer gebucht. Im Altbestand sind
    # das 54 der 74 Buchungen — der größere Teil.
    "ksk_konto_19": "4780",
    "ksk_konto_7": "4781",
    "ksk_konto_0": "4782",
    "ksk_bezeichnung_19": "Fremdarbeiten (19%)",
    "ksk_bezeichnung_7": "Fremdarbeiten (7%)",
    "ksk_bezeichnung_0": "Honorare",
    "ksk_gegenkonto": "001610",
    "ksk_ust_konto_7": "001571",
    "ksk_ust_konto_19": "001576",
    "last_input_dir": "",
    "config_version": 1,
}


def lade_config() -> dict:
    if CONFIG_PFAD.exists():
        with open(CONFIG_PFAD, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        for k, v in DEFAULT_CONFIG.items():
            cfg.setdefault(k, v)
        return cfg
    with open(CONFIG_PFAD, "w", encoding="utf-8") as f:
        json.dump(DEFAULT_CONFIG, f, indent=2, ensure_ascii=False)
    return dict(DEFAULT_CONFIG)


def speichere_config(cfg: dict) -> None:
    with open(CONFIG_PFAD, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)


# ---------------------------------------------------------------------
# Wertebereiche
# ---------------------------------------------------------------------

# Im Bestand vorgefunden. „Mittelrückfluss“ steht nur im Archivblatt, gehört
# aber dazu, sonst fällt der Import darüber.
VERGUETUNGSARTEN = ("Honorar", "Rückfluss", "Erlösanteil",
                    "Darlehensrückzahlung", "Mittelrückfluss", "Spende")

# Nur diese Art zählt für die Künstlersozialkasse — Rückflüsse und
# Erlösanteile sind keine Honorare im Sinne der KSK.
KSK_ART = "Honorar"


# ---------------------------------------------------------------------
# Datenmodell
# ---------------------------------------------------------------------

@dataclass
class Kondition:
    """Wie sich das Entgelt je Exemplar ergibt.

    Der Altbestand rechnet uneinheitlich: mal vom Netto-, mal vom
    Brutto-Ladenpreis, mal ohne Verlagsrabatt, mal geteilt auf mehrere
    Autoren. Damit jede Altzeile 1:1 abbildbar bleibt, wird der Betrag je
    Exemplar ENTWEDER direkt gesetzt (``betrag_je_ex``) ODER aus seinen
    Bestandteilen gerechnet — nie beides.
    """

    betrag_je_ex: float | None = None    # fester Betrag, schlägt alles andere
    ladenpreis: float | None = None
    mwst_im_preis: float | None = None   # 7 / 19 — wird herausgerechnet
    mwst_aufschlagen: bool = False       # eine Altzeile rechnet sie DRAUF
    verlagsrabatt: float = 40.0          # Ladenpreis minus 40 % = Abgabepreis
    rabatt_anwenden: bool = True
    satz: float | None = None            # Honorarsatz in Prozent
    teiler: int = 1                      # Aufteilung auf Mitautoren
    # Staffel als [(Obergrenze, Satz), ...]; die letzte Stufe hat Obergrenze
    # None und gilt nach oben offen.
    staffel: list[tuple[int | None, float]] = field(default_factory=list)
    freimenge: int = 0                   # „ab dem 201. verkauften Exemplar“
    freimenge_ab_jahr: int | None = None
    schwelle_zehn: bool = False
    # Ob der Betrag je Exemplar auf Cent gerundet wird, BEVOR er mit der
    # Menge malgenommen wird. Die Altmappe macht beides: meist steht ein
    # ROUND(...) um die Satzformel, bei einigen Zeilen aber nicht — dort
    # rechnet Excel mit dem vollen Bruch weiter. Der Unterschied macht bei
    # 20 Exemplaren schon acht Cent aus, deshalb wird er mitgeführt.
    satz_runden: bool = True


@dataclass
class Jahreswert:
    """Die Stückzahlen eines Buches in einem Jahr.

    ``verkauft is None`` heißt „noch nicht erfasst“ — ausdrücklich etwas
    anderes als 0 („nichts verkauft“). Der Altbestand verwechselt das an
    mehreren Stellen; hier wird es auseinandergehalten, damit die Abrechnung
    eine fehlende Eingabe meldet, statt stillschweigend null zu rechnen.

    ``eigenkauf`` und ``korrektur`` sind dagegen echte Wahlfelder: leer heißt
    dort 0, so wie es im Altbestand auch gemeint war.
    """

    verkauft: int | None = None
    eigenkauf: int = 0
    korrektur: int = 0
    # Kumulierter Saldo nach diesem Jahr — nur bei Verträgen mit Freimenge
    # oder Staffel gefüllt, sonst None.
    vortrag: int | None = None

    @property
    def erfasst(self) -> bool:
        return self.verkauft is not None


@dataclass
class Buch:
    kennung: str = ""
    titel: str = ""
    isbn: str = ""                       # Verlagskurzform, z. B. „05-300“
    verguetungsart: str = "Honorar"
    kondition: Kondition = field(default_factory=Kondition)
    mwst_pflichtig: bool = False
    mwst_satz: float = 7.0
    jahre: dict[int, Jahreswert] = field(default_factory=dict)
    # Sonderfall Vorauszahlung: offener Restbetrag, der jedes Jahr um das
    # errechnete Honorar schrumpft. Erst wenn er null ist, wird ausgezahlt.
    vorauszahlung: float = 0.0
    # Wird auf einem eigenen Weg abgerechnet und darf NICHT automatisch
    # mitausgezahlt werden. Im Altbestand betrifft das die Verträge aus dem
    # Blatt „Zahlung ab XX Ex.“: dort läuft ein Zähler auf die vereinbarte
    # Freimenge zu, und die Betragsspalte zeigt den Stand, nicht eine
    # fällige Zahlung — die Summe dieser Spalte ist −9.954 €. Nur vier der
    # 54 Personen tauchen überhaupt in der Zahlungsliste auf, mit anderen
    # Beträgen. Wer sie automatisch auszahlt, überweist Geld, das nie
    # geflossen ist.
    gesondert: bool = False
    stillgelegt: bool = False
    stillgelegt_grund: str = ""
    notizen: str = ""
    nachpflege: list[str] = field(default_factory=list)
    # Woher der Satz stammt, z. B. „2025!Z108“. Nur Dokumentation — aber die
    # einzige Möglichkeit, eine Zahl später noch bis in die Altmappe
    # zurückzuverfolgen.
    quelle: str = ""

    def jahr(self, jahr: int) -> Jahreswert:
        """Den Jahrgang holen und dabei anlegen, wenn es ihn noch nicht gibt."""
        if jahr not in self.jahre:
            self.jahre[jahr] = Jahreswert()
        return self.jahre[jahr]


@dataclass
class Empfaenger:
    kennung: str = ""
    lf_nr: str = ""                      # die alte „lf. Nr.“, zum Wiederfinden
    autorenart: str = ""                 # VR | Badenia | Guderjahn | …
    anrede: str = ""                     # Herr | Frau | Damen und Herren | …
    titel_akad: str = ""                 # Dr. / Prof. Dr.
    vorname: str = ""
    name: str = ""
    institution: str = ""
    strasse: str = ""
    plz: str = ""
    ort: str = ""
    land: str = ""
    iban: str = ""
    aktenzeichen: str = ""               # „Buchungszeichen …“, „Kontoinhaber: …“
    email: str = ""
    # Kein Mensch, sondern ein Platzhalter für eine Gruppe — im Altbestand
    # steht dort „verschiedene Autoren“. So ein Posten darf weder einen
    # Brief noch eine Überweisung auslösen: der Betrag ist ein Topf, der
    # erst auf die Beteiligten zu verteilen ist.
    sammelposten: bool = False
    buecher: list[Buch] = field(default_factory=list)
    beruehrt: bool = False               # vom Bediener angefasst

    @property
    def anzeigename(self) -> str:
        person = " ".join(x for x in (self.vorname, self.name) if x).strip()
        if person and self.institution:
            return f"{person} · {self.institution}"
        return person or self.institution or "(ohne Namen)"

    @property
    def dateiname_basis(self) -> str:
        """Der Namensteil der Briefdatei, nach dem Muster der Altbriefe:
        „Buck, Dieter“ bzw. bei reinen Einrichtungen „Stadtarchiv Stuttgart“."""
        if self.name and self.vorname:
            return f"{self.name}, {self.vorname}"
        return self.name or self.institution or self.kennung


@dataclass
class Bestand:
    """Alles, was in ``Honorarbestand.xlsx`` steht."""

    empfaenger: list[Empfaenger] = field(default_factory=list)
    # Beim Laden aufgelaufene Beanstandungen (verwaiste Zeilen, unlesbare
    # Werte). Die Oberfläche zeigt sie an, statt sie zu verschlucken.
    warnungen: list[str] = field(default_factory=list)

    def buecher(self):
        for e in self.empfaenger:
            for b in e.buecher:
                yield e, b

    def naechste_kennung(self, praefix: str) -> str:
        """Fortlaufende Kennung, die keine bestehende überschreibt."""
        vorhanden = set()
        if praefix == "E":
            vorhanden = {e.kennung for e in self.empfaenger}
        else:
            vorhanden = {b.kennung for _, b in self.buecher()}
        n = len(vorhanden) + 1
        while f"{praefix}{n:04d}" in vorhanden:
            n += 1
        return f"{praefix}{n:04d}"


# ---------------------------------------------------------------------
# Typ-Helfer für das Lesen aus Excel
# ---------------------------------------------------------------------
# Excel verwischt beim Speichern von Hand die Unterschiede, auf die es hier
# ankommt: aus der PLZ „69469“ wird die Zahl 69469, aus „Ja“ wird „ja“, und
# eine leere Zelle ist von einer Null kaum noch zu trennen. Jedes Feld geht
# deshalb durch einen dieser Helfer, nie direkt aus der Zelle in das Modell.

def _text(wert) -> str:
    """Zellwert als Text. Aus einer Ganzzahl wird „69469“, nicht „69469.0“."""
    if wert is None:
        return ""
    if isinstance(wert, float) and wert.is_integer():
        return str(int(wert))
    if isinstance(wert, datetime):
        return wert.strftime("%d.%m.%Y")
    return " ".join(str(wert).split()) if isinstance(wert, str) else str(wert)


def _mehrzeilig(wert) -> str:
    """Wie _text, aber Zeilenumbrüche bleiben erhalten (Notizen, Institution)."""
    if wert is None:
        return ""
    if isinstance(wert, float) and wert.is_integer():
        return str(int(wert))
    return str(wert).strip()


def _ganzzahl(wert, vorgabe: int | None = None) -> int | None:
    """Ganzzahl oder ``vorgabe``, wenn die Zelle leer oder unlesbar ist."""
    if wert is None or (isinstance(wert, str) and not wert.strip()):
        return vorgabe
    try:
        return int(round(float(str(wert).replace(",", ".").strip())))
    except (TypeError, ValueError):
        return vorgabe


def _komma(wert, vorgabe: float | None = None) -> float | None:
    """Dezimalzahl; akzeptiert auch „1,49“ aus einer als Text geführten Zelle."""
    if wert is None or (isinstance(wert, str) and not wert.strip()):
        return vorgabe
    try:
        return float(str(wert).replace(",", ".").strip())
    except (TypeError, ValueError):
        return vorgabe


def _ja_nein(wert, vorgabe: bool = False) -> bool:
    """„Ja“/„ja“/„JA“/„x“/True → True. Der Bestand schreibt es uneinheitlich."""
    if wert is None:
        return vorgabe
    if isinstance(wert, bool):
        return wert
    t = str(wert).strip().lower()
    if not t:
        return vorgabe
    return t in ("ja", "j", "x", "wahr", "true", "1")


def _js(wert: bool) -> str:
    return "Ja" if wert else "Nein"


def norm_kopf(wert) -> str:
    """Kopfzellen vergleichbar machen.

    Die Altmappe führt Spaltennamen mit eingebetteten Zeilenumbrüchen
    (``'Vergütungs-\\nexemplare\\n 2025'``); ohne diese Normalisierung findet
    die Kopfzeilensuche sie nicht wieder.
    """
    return " ".join(str(wert or "").split()).strip().lower()


# ---------------------------------------------------------------------
# Kopfzeilensuche
# ---------------------------------------------------------------------
# Nach dem Muster aus mailing_list_updater/core.py: die Kopfzeile wird
# GESUCHT, nicht an einer festen Stelle angenommen. Das überlebt es, wenn
# jemand oben eine Zeile einfügt oder Spalten umsortiert.

def finde_kopfzeile(zeilen: list[tuple], pflichtspalten: set[str],
                    ab: int = 0) -> int:
    """Index der ersten Zeile, die alle Pflichtspalten enthält (0-basiert).

    ``ab`` erlaubt es, in einem Blatt mit mehreren untereinanderliegenden
    Blöcken die nächste Kopfzeile zu suchen — das Blatt „Sonderfälle“ ist so
    gebaut.
    """
    noetig = {norm_kopf(s) for s in pflichtspalten}
    for i in range(ab, len(zeilen)):
        vorhanden = {norm_kopf(z) for z in zeilen[i] if z is not None}
        if noetig <= vorhanden:
            return i
    return -1


def spalten_index(kopfzeile: tuple) -> dict[str, int]:
    """Abbildung normalisierter Spaltenname → Spaltenindex."""
    idx = {}
    for i, z in enumerate(kopfzeile):
        name = norm_kopf(z)
        if name and name not in idx:
            idx[name] = i
    return idx


def hole(zeile: tuple, idx: dict[str, int], spalte: str):
    """Zellwert über den Spaltennamen, tolerant gegen fehlende Spalten."""
    i = idx.get(norm_kopf(spalte))
    if i is None or i >= len(zeile):
        return None
    return zeile[i]


# ---------------------------------------------------------------------
# Aufbau der Bestandsmappe
# ---------------------------------------------------------------------
# Eine Definition je Blatt, von Lesen UND Schreiben benutzt — sonst laufen
# die beiden Seiten mit der Zeit auseinander.

BLATT_EMPFAENGER = "Empfänger"
BLATT_BUECHER = "Bücher"
BLATT_JAHRE = "Jahreswerte"        # alte Fassung, wird noch gelesen
BLATT_STAFFELN = "Staffeln"
BLATT_HINWEISE = "Hinweise"
# Das Arbeitsblatt des laufenden Jahres. Es heißt „Abrechnung <Jahr>“, damit
# auf den ersten Blick klar ist, worum es geht — und damit beim Jahreswechsel
# nicht versehentlich in die Zahlen des Vorjahres getippt wird.
BLATT_ABRECHNUNG = "Abrechnung"
BLATT_HISTORIE = "Historie"
BLATT_REGELN = "Regeln"

SPALTEN_EMPFAENGER = [
    "Kennung", "lf. Nr.", "Autorenart", "Anrede", "Titel", "Vorname", "Name",
    "Institution", "Straße", "PLZ", "Ort", "Land", "IBAN",
    "Aktenzeichen / Kontoinhaber", "E-Mail", "Sammelposten",
]

SPALTEN_BUECHER = [
    "Buch-Kennung", "Empfänger-Kennung", "Buchtitel", "ISBN", "Vergütungsart",
    "MwSt-pflichtig", "MwSt-Satz", "Betrag je Ex.", "Ladenpreis",
    "MwSt im Preis", "MwSt aufschlagen", "Verlagsrabatt", "Rabatt anwenden",
    "Satz", "Teiler", "Satz runden", "Freimenge", "Freimenge ab Jahr",
    "Schwelle 10",
    "Vorauszahlung", "Gesondert abrechnen", "Stillgelegt", "Grund",
    "Notizen", "Nachpflege", "Quelle",
]

SPALTEN_JAHRE = [
    "Buch-Kennung", "Jahr", "verkauft", "Eigenkauf", "Korrektur", "Vortrag",
]

# Das Arbeitsblatt. Links steht, worum es geht, in der Mitte wird getippt,
# rechts steht das Ergebnis der letzten Berechnung. Die Kennung muss mit —
# über sie findet das Werkzeug die Zeile wieder —, steht aber ganz vorn und
# schmal, damit sie nicht stört.
SPALTEN_ABRECHNUNG = [
    "Buch-Kennung", "Autor / Einrichtung", "Buchtitel", "ISBN",
    "Vergütungsart", "€ je Ex.",
    "verkaufte Ex.", "Eigenkauf", "Korrektur",
    "Stand bis Vorjahr", "Vergütungs-Ex.", "Betrag netto", "MwSt",
    "Besonderheit",
]
# Nur diese drei werden aus dem Blatt zurückgelesen. Alles andere ist
# Anzeige und wird beim Speichern neu geschrieben.
EINGABESPALTEN = ("verkaufte Ex.", "Eigenkauf", "Korrektur")

SPALTEN_HISTORIE = [
    "Buch-Kennung", "Buchtitel", "Jahr", "verkauft", "Eigenkauf",
    "Korrektur", "Vortrag",
]

SPALTEN_REGELN = [
    "Art", "Autor / Einrichtung", "Buchtitel", "Was gilt hier",
    "Im Klartext", "Buch-Kennung",
]

SPALTEN_STAFFELN = ["Buch-Kennung", "Stufe", "bis Menge", "Satz"]

# Ohne diese Spalten ist das Blatt nicht verwertbar — fehlt eine, gibt es eine
# verständliche Meldung statt eines Absturzes tief in der Zuordnung.
PFLICHT_EMPFAENGER = {"Kennung", "Name", "Vorname", "Institution"}
PFLICHT_BUECHER = {"Buch-Kennung", "Empfänger-Kennung", "Buchtitel", "ISBN"}
PFLICHT_JAHRE = {"Buch-Kennung", "Jahr", "verkauft"}
PFLICHT_ABRECHNUNG = {"Buch-Kennung", "verkaufte Ex.", "Eigenkauf"}
PFLICHT_HISTORIE = {"Buch-Kennung", "Jahr", "verkauft"}
PFLICHT_STAFFELN = {"Buch-Kennung", "bis Menge", "Satz"}

# Spaltenbreiten, damit die Mappe ohne Nachjustieren lesbar ist.
BREITEN = {
    "Kennung": 10, "Buch-Kennung": 13, "Empfänger-Kennung": 18, "lf. Nr.": 9,
    "Autorenart": 12, "Anrede": 18, "Titel": 10, "Vorname": 16, "Name": 22,
    "Institution": 34, "Straße": 26, "PLZ": 8, "Ort": 20, "Land": 14,
    "IBAN": 26, "Aktenzeichen / Kontoinhaber": 30, "E-Mail": 30,
    "Sammelposten": 13,
    "Buchtitel": 42, "ISBN": 10, "Vergütungsart": 18, "MwSt-pflichtig": 14,
    "MwSt-Satz": 10, "Betrag je Ex.": 13, "Ladenpreis": 11,
    "MwSt im Preis": 13, "MwSt aufschlagen": 16, "Verlagsrabatt": 13,
    "Rabatt anwenden": 15, "Satz": 8, "Teiler": 8, "Satz runden": 12,
    "Freimenge": 10,
    "Freimenge ab Jahr": 16, "Schwelle 10": 12, "Vorauszahlung": 13,
    "Gesondert abrechnen": 18,
    "Stillgelegt": 11, "Grund": 26, "Notizen": 60, "Nachpflege": 40,
    "Quelle": 14,
    "Jahr": 8, "verkauft": 10, "Eigenkauf": 11, "Korrektur": 11,
    "Vortrag": 10, "Stufe": 8, "bis Menge": 11,
    "Art": 15, "Autor / Einrichtung": 34, "€ je Ex.": 11, "verkaufte Ex.": 13,
    "Stand bis Vorjahr": 16, "Vergütungs-Ex.": 14, "Betrag netto": 13,
    "MwSt": 10, "Besonderheit": 30, "Was gilt hier": 22, "Im Klartext": 90,
}

# Zahlenformate. Ohne das „@“ bei PLZ und ISBN macht Excel beim nächsten
# Öffnen wieder Zahlen daraus und frisst führende Nullen.
# Bei den Geldbeträgen steht „[$-407]“ davor: das ist die Kennung für Deutsch
# (Deutschland). Ohne sie entscheidet die Ländereinstellung des Rechners, ob
# aus 3428.79 ein „3.428,79“ oder ein „3,428.79“ wird — und genau das Zweite
# kam bisher heraus.
GELDFORMAT = "[$-407]#,##0.00"
PROZENTFORMAT = "[$-407]0.##"

FORMATE = {
    "PLZ": "@", "ISBN": "@", "IBAN": "@", "lf. Nr.": "@", "Kennung": "@",
    "Buch-Kennung": "@", "Empfänger-Kennung": "@",
    "€ je Ex.": GELDFORMAT, "Betrag netto": GELDFORMAT, "MwSt": GELDFORMAT,
    "Betrag je Ex.": GELDFORMAT, "Ladenpreis": GELDFORMAT,
    "Vorauszahlung": GELDFORMAT, "Satz": "[$-407]0.####",
    "MwSt-Satz": PROZENTFORMAT, "Verlagsrabatt": PROZENTFORMAT,
    "MwSt im Preis": PROZENTFORMAT,
}

HINWEIS_TEXT = [
    ("Honorarbestand — bitte vor dem Ändern lesen", True),
    ("", False),
    ("Diese Mappe ist der Datenbestand des Honorar-Abrechners. Sie darf von "
     "Hand geöffnet und geändert werden — dafür ist sie gemacht. Damit das "
     "gutgeht, gelten fünf Regeln:", False),
    ("", False),
    ("1. Spaltennamen nicht umbenennen, löschen oder übersetzen. Das Werkzeug "
     "sucht die Spalten über ihren Namen. Die Reihenfolge darf sich ändern, "
     "der Name nicht.", False),
    ("2. Kennungen (E0001, B0001 …) nicht ändern. Sie verbinden die Blätter "
     "miteinander. Wird eine Kennung geändert, verliert das Buch seinen Autor "
     "und seine Jahreswerte.", False),
    ("3. Getippt wird im Blatt „Abrechnung <Jahr>“, und zwar nur in den drei "
     "gelb hinterlegten Spalten: verkaufte Ex., Eigenkauf und Korrektur. "
     "Alles andere auf diesem Blatt ist Anzeige und wird beim nächsten "
     "Speichern neu berechnet.", False),
    ("4. Eine LEERE Zelle bei „verkaufte Ex.“ heißt „noch nicht "
     "eingetragen“ — nicht „nichts verkauft“. Wenn ein Buch sich in einem "
     "Jahr wirklich nicht verkauft hat, gehört dort eine 0 hinein. Das "
     "Werkzeug meldet fehlende Eingaben; eine irrtümliche 0 kann es nicht "
     "erkennen.", False),
    ("5. Was an einem Buch besonders ist — Staffel, Freimenge, "
     "Vorauszahlung, Schwelle —, steht in der Spalte „Besonderheit“ und "
     "ausführlich im Blatt „Regeln“. Die Jahre vor dem laufenden stehen im "
     "Blatt „Historie“; der daraus errechnete Stand erscheint im "
     "Arbeitsblatt in der Spalte „Stand bis Vorjahr“.", False),
    ("6. Das Werkzeug schreibt diese Mappe beim Speichern vollständig neu. "
     "Eigene Formeln, Farben, Kommentare und zusätzliche Spalten gehen dabei "
     "verloren. Wer etwas festhalten will, schreibt es in „Notizen“.", False),
    ("7. Die Datei schließen, bevor das Werkzeug speichert. Solange sie in "
     "Excel offen ist, kann das Werkzeug nicht schreiben — es sagt das dann "
     "auch.", False),
    ("", False),
    ("Vor jedem Speichern legt das Werkzeug eine Sicherungskopie unter "
     "Honorarbestand.bak.xlsx an. Wer sich vertan hat, holt sich von dort den "
     "Stand vor der letzten Änderung zurück.", False),
]


# ---------------------------------------------------------------------
# Sperre
# ---------------------------------------------------------------------

def pruefe_sperre(pfad: Path = None) -> str:
    """Prüft, ob die Bestandsmappe gerade in Excel offen ist.

    Excel legt neben einer geöffneten Datei eine versteckte Sperrdatei
    ``~$<Name>.xlsx`` an und verweigert fremde Schreibzugriffe. Ohne diese
    Prüfung stürbe das Speichern mit einem PermissionError — und zwar erst,
    nachdem eine Stunde Erfassung im Fenster steht. Gibt eine Meldung im
    Klartext zurück oder "" wenn alles frei ist.
    """
    pfad = Path(pfad or BESTAND_PFAD)
    sperrdatei = pfad.parent / f"~${pfad.name}"
    if sperrdatei.exists():
        return (f"Der Bestand „{pfad.name}“ ist gerade in Excel geöffnet.\n"
                f"Bitte dort schließen und erneut versuchen.")
    return ""


def setze_arbeitssperre() -> None:
    """Vermerkt, wer den Bestand gerade bearbeitet.

    Rein informativ — sie hindert niemanden, macht aber sichtbar, dass eine
    Zweite schon dran ist. Vorbild ist das Check-out aus
    PLAN_Auto-Update_und_Buchdaten.md.
    """
    try:
        SPERR_PFAD.write_text(
            f"{socket.gethostname()}\n{datetime.now():%d.%m.%Y %H:%M}\n",
            encoding="utf-8")
    except OSError:
        pass  # Eine fehlgeschlagene Notiz darf den Start nicht verhindern.


def loese_arbeitssperre() -> None:
    try:
        SPERR_PFAD.unlink(missing_ok=True)
    except OSError:
        pass


def fremde_arbeitssperre() -> str:
    """Meldung, wenn ein anderer Rechner den Bestand offen hat."""
    if not SPERR_PFAD.exists():
        return ""
    try:
        zeilen = SPERR_PFAD.read_text(encoding="utf-8").splitlines()
    except OSError:
        return ""
    rechner = zeilen[0] if zeilen else "?"
    wann = zeilen[1] if len(zeilen) > 1 else "?"
    if rechner == socket.gethostname():
        return ""
    return (f"Der Bestand wird seit {wann} auf „{rechner}“ bearbeitet.\n"
            f"Wenn beide gleichzeitig speichern, gewinnt der Letzte.")


# ---------------------------------------------------------------------
# Bestand lesen
# ---------------------------------------------------------------------

def _blatt_zeilen(wb, name: str) -> list[tuple]:
    if name not in wb.sheetnames:
        raise ValueError(
            f"Im Bestand fehlt das Blatt „{name}“.\n"
            f"Vorhanden sind: {', '.join(wb.sheetnames)}.")
    return [tuple(z) for z in wb[name].iter_rows(values_only=True)]


def _lies_blatt(wb, name: str, pflicht: set[str]) -> tuple[dict, list[tuple]]:
    """Kopfzeile suchen und die Datenzeilen darunter zurückgeben."""
    zeilen = _blatt_zeilen(wb, name)
    kopf_i = finde_kopfzeile(zeilen, pflicht)
    if kopf_i < 0:
        fehlend = ", ".join(sorted(pflicht))
        raise ValueError(
            f"Im Blatt „{name}“ wurde keine Kopfzeile gefunden.\n"
            f"Erwartet werden mindestens die Spalten: {fehlend}.")
    idx = spalten_index(zeilen[kopf_i])
    daten = [z for z in zeilen[kopf_i + 1:] if any(w is not None for w in z)]
    return idx, daten


def lade_bestand(pfad: Path = None) -> Bestand:
    """Den Bestand aus ``Honorarbestand.xlsx`` lesen.

    Gibt es die Datei noch nicht, kommt ein leerer Bestand zurück — das ist
    der erste Start, kein Fehler.
    """
    pfad = Path(pfad or BESTAND_PFAD)
    if not pfad.exists():
        return Bestand()

    meldung = pruefe_sperre(pfad)
    if meldung:
        raise ValueError(meldung)

    wb = openpyxl.load_workbook(pfad, read_only=True, data_only=True)
    try:
        bestand = Bestand()

        idx, daten = _lies_blatt(wb, BLATT_EMPFAENGER, PFLICHT_EMPFAENGER)
        nach_kennung: dict[str, Empfaenger] = {}
        for z in daten:
            kennung = _text(hole(z, idx, "Kennung"))
            if not kennung:
                bestand.warnungen.append(
                    f"{BLATT_EMPFAENGER}: eine Zeile ohne Kennung wurde "
                    f"übergangen ({_text(hole(z, idx, 'Name'))}).")
                continue
            e = Empfaenger(
                kennung=kennung,
                lf_nr=_text(hole(z, idx, "lf. Nr.")),
                autorenart=_text(hole(z, idx, "Autorenart")),
                anrede=_text(hole(z, idx, "Anrede")),
                titel_akad=_text(hole(z, idx, "Titel")),
                vorname=_text(hole(z, idx, "Vorname")),
                name=_text(hole(z, idx, "Name")),
                institution=_mehrzeilig(hole(z, idx, "Institution")),
                strasse=_text(hole(z, idx, "Straße")),
                plz=_text(hole(z, idx, "PLZ")),
                ort=_text(hole(z, idx, "Ort")),
                land=_text(hole(z, idx, "Land")),
                iban=_text(hole(z, idx, "IBAN")),
                aktenzeichen=_mehrzeilig(
                    hole(z, idx, "Aktenzeichen / Kontoinhaber")),
                email=_text(hole(z, idx, "E-Mail")),
                sammelposten=_ja_nein(hole(z, idx, "Sammelposten")),
            )
            if kennung in nach_kennung:
                bestand.warnungen.append(
                    f"{BLATT_EMPFAENGER}: Kennung „{kennung}“ kommt mehrfach "
                    f"vor; nur der erste Satz wird verwendet.")
                continue
            nach_kennung[kennung] = e
            bestand.empfaenger.append(e)

        idx, daten = _lies_blatt(wb, BLATT_BUECHER, PFLICHT_BUECHER)
        buecher: dict[str, Buch] = {}
        for z in daten:
            bk = _text(hole(z, idx, "Buch-Kennung"))
            ek = _text(hole(z, idx, "Empfänger-Kennung"))
            if not bk:
                continue
            if ek not in nach_kennung:
                bestand.warnungen.append(
                    f"{BLATT_BUECHER}: „{_text(hole(z, idx, 'Buchtitel'))}“ "
                    f"({bk}) verweist auf den unbekannten Empfänger „{ek}“ "
                    f"und wurde übergangen.")
                continue
            nachpflege = _mehrzeilig(hole(z, idx, "Nachpflege"))
            b = Buch(
                kennung=bk,
                titel=_mehrzeilig(hole(z, idx, "Buchtitel")),
                isbn=_text(hole(z, idx, "ISBN")),
                verguetungsart=_text(hole(z, idx, "Vergütungsart")) or "Honorar",
                mwst_pflichtig=_ja_nein(hole(z, idx, "MwSt-pflichtig")),
                mwst_satz=_komma(hole(z, idx, "MwSt-Satz"), 7.0),
                vorauszahlung=_komma(hole(z, idx, "Vorauszahlung"), 0.0),
                gesondert=_ja_nein(hole(z, idx, "Gesondert abrechnen")),
                stillgelegt=_ja_nein(hole(z, idx, "Stillgelegt")),
                stillgelegt_grund=_mehrzeilig(hole(z, idx, "Grund")),
                notizen=_mehrzeilig(hole(z, idx, "Notizen")),
                nachpflege=[t for t in nachpflege.split(" | ") if t],
                quelle=_text(hole(z, idx, "Quelle")),
                kondition=Kondition(
                    betrag_je_ex=_komma(hole(z, idx, "Betrag je Ex.")),
                    ladenpreis=_komma(hole(z, idx, "Ladenpreis")),
                    mwst_im_preis=_komma(hole(z, idx, "MwSt im Preis")),
                    mwst_aufschlagen=_ja_nein(
                        hole(z, idx, "MwSt aufschlagen")),
                    verlagsrabatt=_komma(hole(z, idx, "Verlagsrabatt"), 40.0),
                    rabatt_anwenden=_ja_nein(
                        hole(z, idx, "Rabatt anwenden"), True),
                    satz=_komma(hole(z, idx, "Satz")),
                    teiler=_ganzzahl(hole(z, idx, "Teiler"), 1) or 1,
                    satz_runden=_ja_nein(hole(z, idx, "Satz runden"), True),
                    freimenge=_ganzzahl(hole(z, idx, "Freimenge"), 0) or 0,
                    freimenge_ab_jahr=_ganzzahl(
                        hole(z, idx, "Freimenge ab Jahr")),
                    schwelle_zehn=_ja_nein(hole(z, idx, "Schwelle 10")),
                ),
            )
            if bk in buecher:
                bestand.warnungen.append(
                    f"{BLATT_BUECHER}: Buch-Kennung „{bk}“ kommt mehrfach vor.")
                continue
            buecher[bk] = b
            nach_kennung[ek].buecher.append(b)

        def merke_jahr(blatt, bk, jahr, verkauft, eigenkauf, korrektur,
                       vortrag=None):
            if not bk or jahr is None:
                return
            if bk not in buecher:
                bestand.warnungen.append(
                    f"{blatt}: Jahrgang {jahr} verweist auf das unbekannte "
                    f"Buch „{bk}“ und wurde übergangen.")
                return
            buecher[bk].jahre[jahr] = Jahreswert(
                # Leer heißt „nicht erfasst“ — deshalb hier KEIN Vorgabewert 0.
                verkauft=verkauft, eigenkauf=eigenkauf or 0,
                korrektur=korrektur or 0, vortrag=vortrag)

        # Das Arbeitsblatt des laufenden Jahres. Es heißt „Abrechnung <Jahr>“;
        # das Jahr steht im Blattnamen, damit es beim Wechsel nicht
        # verwechselt werden kann.
        for name in wb.sheetnames:
            if not name.startswith(BLATT_ABRECHNUNG + " "):
                continue
            jahr = _ganzzahl(name.split()[-1])
            if jahr is None:
                continue
            idx, daten = _lies_blatt(wb, name, PFLICHT_ABRECHNUNG)
            for z in daten:
                merke_jahr(name, _text(hole(z, idx, "Buch-Kennung")), jahr,
                           _ganzzahl(hole(z, idx, "verkaufte Ex.")),
                           _ganzzahl(hole(z, idx, "Eigenkauf"), 0),
                           _ganzzahl(hole(z, idx, "Korrektur"), 0))

        if BLATT_HISTORIE in wb.sheetnames:
            idx, daten = _lies_blatt(wb, BLATT_HISTORIE, PFLICHT_HISTORIE)
            for z in daten:
                merke_jahr(BLATT_HISTORIE, _text(hole(z, idx, "Buch-Kennung")),
                           _ganzzahl(hole(z, idx, "Jahr")),
                           _ganzzahl(hole(z, idx, "verkauft")),
                           _ganzzahl(hole(z, idx, "Eigenkauf"), 0),
                           _ganzzahl(hole(z, idx, "Korrektur"), 0),
                           _ganzzahl(hole(z, idx, "Vortrag")))

        # Ältere Bestände führten alle Jahre in einem Blatt „Jahreswerte“.
        # Die werden weiter gelesen, damit eine gespeicherte Mappe nach dem
        # Umbau nicht wertlos wird.
        if BLATT_JAHRE in wb.sheetnames:
            idx, daten = _lies_blatt(wb, BLATT_JAHRE, PFLICHT_JAHRE)
            for z in daten:
                merke_jahr(BLATT_JAHRE, _text(hole(z, idx, "Buch-Kennung")),
                           _ganzzahl(hole(z, idx, "Jahr")),
                           _ganzzahl(hole(z, idx, "verkauft")),
                           _ganzzahl(hole(z, idx, "Eigenkauf"), 0),
                           _ganzzahl(hole(z, idx, "Korrektur"), 0),
                           _ganzzahl(hole(z, idx, "Vortrag")))

        if BLATT_STAFFELN in wb.sheetnames:
            idx, daten = _lies_blatt(wb, BLATT_STAFFELN, PFLICHT_STAFFELN)
            roh: dict[str, list] = {}
            for z in daten:
                bk = _text(hole(z, idx, "Buch-Kennung"))
                satz = _komma(hole(z, idx, "Satz"))
                if not bk or satz is None:
                    continue
                if bk not in buecher:
                    bestand.warnungen.append(
                        f"{BLATT_STAFFELN}: Stufe verweist auf das unbekannte "
                        f"Buch „{bk}“ und wurde übergangen.")
                    continue
                stufe = _ganzzahl(hole(z, idx, "Stufe"), 0) or 0
                roh.setdefault(bk, []).append(
                    (stufe, _ganzzahl(hole(z, idx, "bis Menge")), satz))
            for bk, stufen in roh.items():
                # Nach Stufennummer ordnen; die offene Stufe (bis Menge leer)
                # gehört immer ans Ende.
                stufen.sort(key=lambda s: (s[1] is None, s[1] or 0, s[0]))
                buecher[bk].kondition.staffel = [(g, s) for _, g, s in stufen]

        return bestand
    finally:
        wb.close()


# ---------------------------------------------------------------------
# Bestand schreiben
# ---------------------------------------------------------------------

def _schreibe_blatt(wb, name: str, spalten: list[str],
                    zeilen: list[dict]) -> None:
    """Ein Datenblatt anlegen, nach dem Hausmuster mit fixierter Kopfzeile."""
    ws = wb.create_sheet(name)
    ws.append(spalten)
    for z in ws[1]:
        z.font = Font(bold=True)
    for zeile in zeilen:
        ws.append([zeile.get(s) for s in spalten])
    ws.freeze_panes = "A2"
    for i, s in enumerate(spalten, start=1):
        buchstabe = ws.cell(row=1, column=i).column_letter
        ws.column_dimensions[buchstabe].width = BREITEN.get(s, 14)
        fmt = FORMATE.get(s)
        if fmt:
            # Format auf die ganze Spalte legen, auch auf noch leere Zeilen —
            # sonst macht Excel aus einer nachgetragenen PLZ wieder eine Zahl.
            for zelle in ws[buchstabe][1:]:
                zelle.number_format = fmt


def _kennzeichne_eingabespalten(ws, spalten: list[str]) -> None:
    """Die Spalten, in die getippt wird, farblich hinterlegen.

    Vierzehn Spalten nebeneinander, und nur drei davon sind zum Ausfüllen —
    ohne Markierung rät man jedes Jahr aufs Neue, welche.
    """
    hell = PatternFill("solid", fgColor="FFF6DA")
    for i, name in enumerate(spalten, start=1):
        if name not in EINGABESPALTEN:
            continue
        buchstabe = ws.cell(row=1, column=i).column_letter
        for zelle in ws[buchstabe]:
            zelle.fill = hell


def _schreibe_hinweise(wb) -> None:
    ws = wb.create_sheet(BLATT_HINWEISE)
    ws.column_dimensions["A"].width = 110
    for text, fett in HINWEIS_TEXT:
        ws.append([text])
        zelle = ws.cell(row=ws.max_row, column=1)
        zelle.alignment = Alignment(wrapText=True, vertical="top")
        if fett:
            zelle.font = Font(bold=True, size=13)


def sichere_zwischenstand(bestand: Bestand, jahr: int | None = None,
                          cfg: dict | None = None) -> Path | None:
    """Stillen Zwischenstand schreiben, ohne den richtigen Bestand anzufassen.

    Bewusst in eine eigene Datei: der Bediener hat noch nicht gespeichert,
    also darf auch nichts Gespeichertes überschrieben werden. Schlägt das
    Schreiben fehl, wird das absichtlich verschluckt — ein Zwischenstand ist
    eine Zugabe und darf die laufende Arbeit nie unterbrechen.
    """
    try:
        return speichere_bestand(bestand, WIEDERHERSTELLUNG_PFAD,
                                 mit_sicherung=False, jahr=jahr, cfg=cfg)
    except Exception:
        return None


def verwerfe_zwischenstand() -> None:
    try:
        WIEDERHERSTELLUNG_PFAD.unlink(missing_ok=True)
    except OSError:
        pass


def offener_zwischenstand() -> str:
    """Meldung, wenn ein Zwischenstand jünger ist als der Bestand.

    Gibt "" zurück, wenn es nichts wiederherzustellen gibt.
    """
    if not WIEDERHERSTELLUNG_PFAD.exists():
        return ""
    try:
        zwischen = WIEDERHERSTELLUNG_PFAD.stat().st_mtime
        bestand = BESTAND_PFAD.stat().st_mtime if BESTAND_PFAD.exists() else 0
    except OSError:
        return ""
    if zwischen <= bestand:
        return ""
    wann = datetime.fromtimestamp(zwischen).strftime("%d.%m.%Y um %H:%M")
    return (f"Beim letzten Mal wurde das Programm beendet, ohne zu speichern.\n"
            f"Es liegen Eingaben von {wann} vor.\n\n"
            f"Sollen sie zurückgeholt werden?")


def besonderheit(buch: "Buch") -> str:
    """Ein kurzes Wort für das, was an diesem Buch nicht gewöhnlich ist.

    Dieselbe Quelle wie das Blatt „Regeln“ — was hier steht, ist dort
    ausführlich erklärt.
    """
    teile = []
    for art in regelarten(buch):
        if art == "Freimenge":
            teile.append(f"Honorar ab dem {buch.kondition.freimenge + 1}. Ex.")
        elif art == "Mitautoren":
            teile.append(f"geteilt durch {buch.kondition.teiler}")
        elif art == "Schwelle":
            teile.append("erst ab 10 Ex.")
        elif art == "Gesondert":
            teile.append("gesondert abrechnen")
        elif art == "Stillgelegt":
            teile.append("stillgelegt")
        elif art == "Vorauszahlung":
            teile.append("Vorauszahlung offen")
        elif art == "Bitte prüfen":
            teile.append("bitte prüfen")
        else:
            teile.append(art)
    return ", ".join(teile)


# Die Sonderregeln, nach Tragweite geordnet. Wer das Blatt öffnet, soll
# oben das finden, was eine Abrechnung wirklich verändert — und nicht erst
# durch hundert stillgelegte Titel scrollen.
REGELARTEN = [
    ("Vorauszahlung", "Vorauszahlung offen"),
    ("Staffel", "Gestaffelter Satz"),
    ("Freimenge", "Freimenge"),
    ("Gesondert", "Gesondert abrechnen"),
    ("Schwelle", "Erst ab zehn Exemplaren"),
    ("Mitautoren", "Auf mehrere Autoren geteilt"),
    ("Bitte prüfen", "Beim Einlesen aufgefallen"),
    ("Stillgelegt", "Wird nicht mehr abgerechnet"),
]


def regelarten(buch: "Buch") -> list[str]:
    """Welche Sonderregeln an diesem Buch hängen — leer heißt: keine."""
    k = buch.kondition
    arten = []
    if buch.vorauszahlung:
        arten.append("Vorauszahlung")
    if k.staffel:
        arten.append("Staffel")
    if k.freimenge:
        arten.append("Freimenge")
    if buch.gesondert:
        arten.append("Gesondert")
    if k.schwelle_zehn:
        arten.append("Schwelle")
    if k.teiler != 1:
        arten.append("Mitautoren")
    if buch.nachpflege:
        arten.append("Bitte prüfen")
    if buch.stillgelegt:
        arten.append("Stillgelegt")
    return arten


def regelzeilen(bestand: Bestand) -> list[dict]:
    """Die besonderen Abmachungen in ganzen Sätzen, je Buch eine Zeile.

    NUR Bücher mit einer Besonderheit. Die gewöhnliche Rechnung — Ladenpreis,
    Rabatt, Satz — gehört nicht hierher: sie steht im Blatt „Bücher“ und im
    Werkzeug hinter dem Doppelklick. Stünde sie auch hier, wären es 412 statt
    193 Zeilen, und die 24 Staffeln, um die es geht, lägen darin begraben.

    Geordnet nach Tragweite: was die Auszahlung am stärksten verändert,
    steht oben; stillgelegte Titel stehen zuletzt.
    """
    rang = {name: i for i, (name, _) in enumerate(REGELARTEN)}
    zeilen = []
    for e, b in bestand.buecher():
        arten = regelarten(b)
        if not arten:
            continue
        k = b.kondition
        saetze = []
        if b.vorauszahlung:
            saetze.append(
                f"Es ist noch eine Vorauszahlung von {euro(b.vorauszahlung)} "
                f"offen. Sie wird vom Honorar abgezogen, bis sie getilgt ist.")
        if k.staffel:
            stufen, untere = [], 1
            for grenze, satz in k.staffel:
                if grenze is None:
                    stufen.append(f"ab {untere} Exemplaren {satz:g} %")
                else:
                    stufen.append(f"{untere} bis {grenze} Exemplare {satz:g} %")
                    untere = grenze + 1
            saetze.append("Gestaffelter Satz: " + ", ".join(stufen)
                          + ". Welche Stufe gilt, entscheidet der Stand zu "
                            "Jahresbeginn (Spalte „Stand bis Vorjahr“).")
        if k.freimenge:
            saetze.append(
                f"Die ersten {k.freimenge} Exemplare werden nicht vergütet. "
                f"Solange der Stand darunter liegt, wird nichts gezahlt — "
                f"und auch nichts zurückgefordert.")
        if b.gesondert:
            saetze.append(
                "Wird auf einem eigenen Weg abgerechnet und löst hier keine "
                "Auszahlung aus. Die Zahlen laufen mit, damit Staffel und "
                "Freimenge stimmen.")
        if k.schwelle_zehn:
            saetze.append(
                "Unter zehn Vergütungsexemplaren im Jahr entfällt das "
                "Honorar. Bei genau zehn wird gezahlt.")
        if k.teiler != 1:
            saetze.append(f"Das Honorar teilen sich {k.teiler} Autoren.")
        if b.stillgelegt:
            saetze.append("Wird nicht mehr abgerechnet"
                          + (f" — {b.stillgelegt_grund}." if b.stillgelegt_grund
                             else "."))
        for hinweis in b.nachpflege:
            saetze.append("Beim Einlesen aufgefallen: " + hinweis)

        zeilen.append({
            "_rang": min(rang[a] for a in arten),
            "Art": arten[0],
            "Buch-Kennung": b.kennung,
            "Autor / Einrichtung": e.anzeigename,
            "Buchtitel": b.titel,
            "Was gilt hier": ", ".join(arten),
            "Im Klartext": " ".join(saetze),
        })
    zeilen.sort(key=lambda z: (z["_rang"], z["Autor / Einrichtung"].lower()))
    for z in zeilen:
        z.pop("_rang")
    return zeilen


def speichere_bestand(bestand: Bestand, pfad: Path = None,
                      mit_sicherung: bool = True, jahr: int | None = None,
                      cfg: dict | None = None) -> Path:
    """Den Bestand schreiben — Sicherungskopie zuerst, dann atomar ersetzen.

    Atomar über Temp-Datei + os.replace, weil die Mappe auf dem Netzlaufwerk
    liegen kann: bricht das Schreiben ab, steht dort noch die alte, heile
    Datei statt einer halben.
    """
    pfad = Path(pfad or BESTAND_PFAD)
    meldung = pruefe_sperre(pfad)
    if meldung:
        raise ValueError(meldung)

    if mit_sicherung and pfad.exists():
        try:
            shutil.copy2(pfad, SICHERUNG_PFAD)
        except OSError as e:
            raise ValueError(
                f"Die Sicherungskopie konnte nicht angelegt werden: {e}\n"
                f"Es wurde nichts geschrieben.") from e

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    _schreibe_hinweise(wb)

    _schreibe_blatt(wb, BLATT_EMPFAENGER, SPALTEN_EMPFAENGER, [
        {
            "Kennung": e.kennung, "lf. Nr.": e.lf_nr,
            "Autorenart": e.autorenart, "Anrede": e.anrede,
            "Titel": e.titel_akad, "Vorname": e.vorname, "Name": e.name,
            "Institution": e.institution, "Straße": e.strasse, "PLZ": e.plz,
            "Ort": e.ort, "Land": e.land, "IBAN": e.iban,
            "Aktenzeichen / Kontoinhaber": e.aktenzeichen, "E-Mail": e.email,
            "Sammelposten": _js(e.sammelposten),
        }
        for e in bestand.empfaenger
    ])

    _schreibe_blatt(wb, BLATT_BUECHER, SPALTEN_BUECHER, [
        {
            "Buch-Kennung": b.kennung, "Empfänger-Kennung": e.kennung,
            "Buchtitel": b.titel, "ISBN": b.isbn,
            "Vergütungsart": b.verguetungsart,
            "MwSt-pflichtig": _js(b.mwst_pflichtig), "MwSt-Satz": b.mwst_satz,
            "Betrag je Ex.": b.kondition.betrag_je_ex,
            "Ladenpreis": b.kondition.ladenpreis,
            "MwSt im Preis": b.kondition.mwst_im_preis,
            "MwSt aufschlagen": _js(b.kondition.mwst_aufschlagen),
            "Verlagsrabatt": b.kondition.verlagsrabatt,
            "Rabatt anwenden": _js(b.kondition.rabatt_anwenden),
            "Satz": b.kondition.satz, "Teiler": b.kondition.teiler,
            "Satz runden": _js(b.kondition.satz_runden),
            "Freimenge": b.kondition.freimenge or None,
            "Freimenge ab Jahr": b.kondition.freimenge_ab_jahr,
            "Schwelle 10": _js(b.kondition.schwelle_zehn),
            "Vorauszahlung": b.vorauszahlung or None,
            "Gesondert abrechnen": _js(b.gesondert),
            "Stillgelegt": _js(b.stillgelegt), "Grund": b.stillgelegt_grund,
            "Notizen": b.notizen, "Nachpflege": " | ".join(b.nachpflege),
            "Quelle": b.quelle,
        }
        for e, b in bestand.buecher()
    ])

    # --- Das Arbeitsblatt des laufenden Jahres -------------------------
    if jahr is None:
        vorhanden = [j for _, b in bestand.buecher() for j in b.jahre]
        jahr = max(vorhanden) if vorhanden else date.today().year - 1

    abrechnungszeilen, historienzeilen = [], []
    # Nach Autor und Titel sortiert — so sucht ein Mensch, nicht nach
    # Kennungen.
    sortiert = sorted(bestand.buecher(),
                      key=lambda p: (p[0].anzeigename.lower(), p[1].titel.lower()))
    for e, b in sortiert:
        if b.stillgelegt:
            # Stillgelegte Bücher gehören nicht ins Arbeitsblatt — sie
            # blähten es um ein Fünftel auf, ohne dass je etwas einzutragen
            # wäre. Ihre Zahlen wandern vollständig in die Historie, damit
            # nichts verlorengeht.
            for j in sorted(b.jahre):
                h = b.jahre[j]
                historienzeilen.append({
                    "Buch-Kennung": b.kennung, "Buchtitel": b.titel, "Jahr": j,
                    "verkauft": h.verkauft, "Eigenkauf": h.eigenkauf,
                    "Korrektur": h.korrektur or None, "Vortrag": h.vortrag,
                })
            continue
        jw = b.jahre.get(jahr)
        posten = betrag_zeile(b, jahr, cfg or {}) if jw else None
        abrechnungszeilen.append({
            "Buch-Kennung": b.kennung,
            "Autor / Einrichtung": e.anzeigename,
            "Buchtitel": b.titel,
            "ISBN": b.isbn,
            "Vergütungsart": b.verguetungsart,
            "€ je Ex.": satz_je_ex(b, jahr),
            # Leer heißt „noch nicht eingetragen“ — deshalb keine 0 erzwingen.
            "verkaufte Ex.": jw.verkauft if jw else None,
            "Eigenkauf": jw.eigenkauf if jw and jw.verkauft is not None else None,
            "Korrektur": (jw.korrektur or None) if jw else None,
            "Stand bis Vorjahr": kumulierte_menge(b, jahr),
            "Vergütungs-Ex.": posten.verguetungs_ex if posten else None,
            "Betrag netto": posten.netto if posten else None,
            "MwSt": posten.mwst if posten else None,
            "Besonderheit": besonderheit(b),
        })
        for j in sorted(b.jahre):
            if j == jahr:
                continue
            h = b.jahre[j]
            historienzeilen.append({
                "Buch-Kennung": b.kennung, "Buchtitel": b.titel, "Jahr": j,
                "verkauft": h.verkauft, "Eigenkauf": h.eigenkauf,
                "Korrektur": h.korrektur or None, "Vortrag": h.vortrag,
            })

    blattname = f"{BLATT_ABRECHNUNG} {jahr}"
    _schreibe_blatt(wb, blattname, SPALTEN_ABRECHNUNG, abrechnungszeilen)
    _kennzeichne_eingabespalten(wb[blattname], SPALTEN_ABRECHNUNG)
    _schreibe_blatt(wb, BLATT_HISTORIE, SPALTEN_HISTORIE, historienzeilen)

    staffelzeilen = []
    for _, b in bestand.buecher():
        for i, (grenze, satz) in enumerate(b.kondition.staffel, start=1):
            staffelzeilen.append({
                "Buch-Kennung": b.kennung, "Stufe": i,
                "bis Menge": grenze, "Satz": satz,
            })
    _schreibe_blatt(wb, BLATT_STAFFELN, SPALTEN_STAFFELN, staffelzeilen)
    _schreibe_blatt(wb, BLATT_REGELN, SPALTEN_REGELN, regelzeilen(bestand))

    tmp = pfad.parent / f".{pfad.name}.{os.getpid()}.tmp"
    try:
        wb.save(tmp)
        os.replace(tmp, pfad)
    except OSError as e:
        tmp.unlink(missing_ok=True)
        raise ValueError(
            f"Der Bestand konnte nicht geschrieben werden: {e}\n"
            f"Ist die Datei in Excel geöffnet?") from e
    finally:
        wb.close()
    return pfad


# ---------------------------------------------------------------------
# Die Satzformel der Altmappe zurückübersetzen
# ---------------------------------------------------------------------
# In der Altmappe ist der Betrag je Exemplar keine Zahl, sondern eine von
# Hand getippte Formel, die den ganzen Vertrag kodiert:
#
#     =ROUND(16.9/1.07*0.6*0.12,2)
#      └ Ladenpreis 16,90 brutto
#              └ /1,07 → netto
#                    └ ×0,6 → Verlagsabgabepreis (Ladenpreis minus 40 %)
#                        └ ×0,12 → 12 % Honorarsatz
#
# Geschrieben wird das uneinheitlich: mal `/1.07`, mal `*100/107`, mal gar
# nicht (dann ist vom Bruttopreis gerechnet), mal `/3` für drei Mitautoren.
# Ein Muster-Regex je Schreibweise erfasst nur zwei Drittel der Zeilen —
# deshalb wird die Formel stattdessen in ihre Faktoren zerlegt und jeder
# Faktor einzeln gedeutet. Damit gehen 191 der 192 Formeln auf den Cent auf.

_RX_PROZENTLITERAL = re.compile(r"([\d.]+)\s*%")


def _tokens(ausdruck: str) -> list[tuple[float, bool]]:
    """Ein reines Produkt/Quotient in seine Zahlen zerlegen, IN REIHENFOLGE.

    Die Reihenfolge ist nicht schmückendes Beiwerk, sondern trägt Bedeutung:
    in ``11.9*40/100*10/100`` ist das erste ``40/100`` der Verlagsabgabepreis
    und das zweite ``10/100`` der Honorarsatz. Ohne Reihenfolge sind die
    beiden nicht auseinanderzuhalten.

    Rückgabe: [(Zahl, steht_im_nenner), ...]. Alles außer Zahl, Mal und
    Geteilt gilt als nicht zerlegbar.
    """
    baum = ast.parse(ausdruck, mode="eval").body
    folge: list[tuple[float, bool]] = []

    def geh(knoten, unten: bool):
        if isinstance(knoten, ast.Constant) and isinstance(
                knoten.value, (int, float)):
            folge.append((float(knoten.value), unten))
        elif isinstance(knoten, ast.BinOp) and isinstance(knoten.op, ast.Mult):
            geh(knoten.left, unten)
            geh(knoten.right, unten)
        elif isinstance(knoten, ast.BinOp) and isinstance(knoten.op, ast.Div):
            geh(knoten.left, unten)
            geh(knoten.right, not unten)
        else:
            raise ValueError("nicht zerlegbar")

    geh(baum, False)
    return folge


def _loese_mwst(folge: list[tuple[float, bool]]) -> tuple[float | None, bool]:
    """Den Mehrwertsteuer-Faktor herausnehmen.

    Vier Schreibweisen kommen im Bestand vor und meinen dreimal dasselbe:
    ``/1.07`` und ``*100/107`` rechnen die Steuer aus dem Ladenpreis heraus,
    ``*107/100`` schlägt sie auf (eine einzige Altzeile).
    """
    for satz, teiler in ((7.0, 1.07), (19.0, 1.19)):
        for i, (wert, unten) in enumerate(folge):
            if unten and abs(wert - teiler) < 1e-9:
                del folge[i]
                return satz, False
    for satz, hundert in ((7.0, 107.0), (19.0, 119.0)):
        for i, (wert, unten) in enumerate(folge):
            if unten and abs(wert - hundert) < 1e-9:
                for j, (w2, u2) in enumerate(folge):
                    if not u2 and abs(w2 - 100.0) < 1e-9:
                        del folge[max(i, j)]
                        del folge[min(i, j)]
                        return satz, False
            if not unten and abs(wert - hundert) < 1e-9:
                for j, (w2, u2) in enumerate(folge):
                    if u2 and abs(w2 - 100.0) < 1e-9:
                        del folge[max(i, j)]
                        del folge[min(i, j)]
                        return satz, True
    return None, False


def _fasse_faktoren(folge: list[tuple[float, bool]]) -> list[float]:
    """Aus der Zahlenfolge multiplikative Faktoren machen.

    ``40`` gefolgt von ``/100`` ist EIN Faktor (0,4), nicht zwei — sonst
    zerfällt der Prozentsatz in zwei sinnlose Hälften.
    """
    faktoren: list[float] = []
    i = 0
    while i < len(folge):
        wert, unten = folge[i]
        if not unten and i + 1 < len(folge):
            w2, u2 = folge[i + 1]
            if u2 and abs(w2 - 100.0) < 1e-9:
                faktoren.append(wert / 100.0)
                i += 2
                continue
        faktoren.append(1.0 / wert if unten else wert)
        i += 1
    return faktoren


def zerlege_satzformel(formel: str) -> tuple[Kondition | None, str]:
    """``=ROUND(16.9/1.07*0.6*0.12,2)`` → Kondition(16,90 · 7 % · 40 % · 12 %).

    Gibt (Kondition, "") zurück, wenn die Zerlegung aufgeht und sich der
    Zellwert damit nachrechnen lässt, sonst (None, Begründung) — dann
    übernimmt der Aufrufer nur den fertigen Zahlenwert und merkt die Zeile
    zur Nachpflege vor. Geraten wird nichts.
    """
    if not isinstance(formel, str) or not formel.strip().startswith("="):
        return None, "keine Formel"

    text = formel.strip()
    gerundet = "ROUND" in text.upper()
    inneres = text[1:].strip()
    # Die Klammer um alles herum kommt vor: =(ROUND(...)/2)
    while inneres.startswith("(") and inneres.endswith(")"):
        inneres = inneres[1:-1].strip()
    # ROUND(...) auflösen. Die Rundung selbst wird nicht nachgebildet: sie
    # sitzt in satz_aus_kondition an derselben Stelle, und das Ergebnis wird
    # unten ohnehin gegen den Zellwert geprüft.
    while True:
        m = re.search(r"ROUND\s*\(", inneres, re.I)
        if not m:
            break
        tiefe, i = 0, m.end() - 1
        for i in range(m.end() - 1, len(inneres)):
            if inneres[i] == "(":
                tiefe += 1
            elif inneres[i] == ")":
                tiefe -= 1
                if tiefe == 0:
                    break
        argumente = inneres[m.end():i]
        # letzte oberste Komma-Trennung ist die Stellenzahl
        tiefe, schnitt = 0, None
        for k, z in enumerate(argumente):
            if z == "(":
                tiefe += 1
            elif z == ")":
                tiefe -= 1
            elif z == "," and tiefe == 0:
                schnitt = k
        if schnitt is not None:
            argumente = argumente[:schnitt]
        inneres = inneres[:m.start()] + "(" + argumente + ")" + inneres[i + 1:]

    # Excel schreibt Prozente als Literal: 5% meint 0,05.
    inneres = _RX_PROZENTLITERAL.sub(r"(\1/100)", inneres)

    try:
        folge = _tokens(inneres)
    except (SyntaxError, ValueError, RecursionError):
        return None, f"Formel nicht zerlegbar: {text}"
    if not folge:
        return None, f"Formel ohne Zahlen: {text}"

    kond = Kondition()
    kond.satz_runden = gerundet
    kond.mwst_im_preis, kond.mwst_aufschlagen = _loese_mwst(folge)

    # Mitautoren-Teiler: eine Division durch 2, 3 oder 4. Eine /10 ist keine
    # Aufteilung, sondern „10 %" — deshalb die enge Liste.
    for i, (wert, unten) in enumerate(folge):
        if unten and wert in (2.0, 3.0, 4.0):
            kond.teiler = int(wert)
            del folge[i]
            break

    faktoren = _fasse_faktoren(folge)
    if not faktoren:
        return None, f"Formel ohne Preis: {text}"

    # Der erste Faktor über 1 ist der Ladenpreis.
    preis_i = next((i for i, f in enumerate(faktoren) if f > 1.0), None)
    if preis_i is None:
        # Kein Preis, nur eine Zahl — das ist ein fester Betrag je Exemplar.
        return None, f"Formel ohne Ladenpreis: {text}"
    kond.ladenpreis = faktoren.pop(preis_i)

    # Was übrig bleibt, sind Anteile. Zwei davon heißen: erst der
    # Verlagsabgabepreis, dann der Honorarsatz. Nur einer heißt: nur der
    # Honorarsatz, gerechnet direkt vom (Netto-)Ladenpreis.
    if len(faktoren) >= 2:
        kond.verlagsrabatt = round((1.0 - faktoren[0]) * 100, 6)
        kond.rabatt_anwenden = True
        kond.satz = round(faktoren[1] * 100, 6)
        rest = faktoren[2:]
    elif len(faktoren) == 1:
        kond.rabatt_anwenden = False
        kond.satz = round(faktoren[0] * 100, 6)
        rest = []
    else:
        kond.rabatt_anwenden = False
        rest = []

    if rest:
        return None, f"Formel nicht vollständig gedeutet (Rest {rest}): {text}"
    if kond.satz is None and kond.teiler == 1:
        # Ohne Satz UND ohne Teiler bliebe nur der nackte Ladenpreis stehen —
        # das ist keine Kondition, sondern ein Missverständnis. ``=24.9/2``
        # (die Hälfte des Verkaufspreises einer CD) ist dagegen in Ordnung.
        return None, f"Formel ohne Satz: {text}"
    return kond, ""


def satz_aus_kondition(kond: Kondition, satz: float | None = None) -> float:
    """Betrag je Exemplar aus den Bestandteilen.

    Gerundet wird der Satz je Exemplar auf volle Cent, NICHT erst das
    spätere Produkt aus Menge und Satz. Und zwar IMMER: in der Altmappe
    fehlte das RUNDEN in acht Formeln, und dann steht im Brief „20 × 1,85 €
    = 36,92 €“ — eine Rechnung, die kein Autor nachrechnen kann. Der Schalter
    ``satz_runden`` wird deshalb nicht mehr beachtet.
    """
    if kond.betrag_je_ex is not None:
        return runde(kond.betrag_je_ex)
    if kond.ladenpreis is None:
        return 0.0
    wert = kond.ladenpreis
    if kond.mwst_im_preis:
        if kond.mwst_aufschlagen:
            wert *= 1 + kond.mwst_im_preis / 100
        else:
            wert /= 1 + kond.mwst_im_preis / 100
    if kond.rabatt_anwenden:
        wert *= 1 - kond.verlagsrabatt / 100
    genutzt = satz if satz is not None else kond.satz
    if genutzt is not None:
        wert *= genutzt / 100
    if kond.teiler and kond.teiler != 1:
        wert /= kond.teiler
    return runde(wert)


# ---------------------------------------------------------------------
# Die Staffel aus dem Notizentext ziehen
# ---------------------------------------------------------------------
# Die Vertragsstaffel steht in der Altmappe nur als Fließtext in „Notizen":
#
#   „lt. Vertrag bis 2500 Ex. 12% v. Nettoverlagsabgabepreis, ab 2501 Ex.
#    13%, Stand 2023: 1460 Ex."
#
# Gerechnet wird sie gar nicht — der Prozentsatz ist von Hand in die
# Satzformel getippt. Genau das soll das Werkzeug abnehmen, also wird der
# Text gelesen. Greift keiner der Ausdrücke, bleibt die Staffel leer und die
# Zeile kommt zur Nachpflege: geraten wird nichts.

_RX_BIS = re.compile(
    r"(?:bis|1\s*[–\-]\s*)\s*([\d.]+)\s*(?:Ex(?:emplare?n?)?\.?)?"
    r"\s*[,:]?\s*([\d,]+)\s*%", re.I)
_RX_AB = re.compile(
    r"ab\s+(?:dem\s+)?([\d.]+)\.?\s*(?:Ex(?:emplare?n?)?\.?)?"
    r"\s*[,:]?\s*([\d,]+)\s*%", re.I)
_RX_STAND = re.compile(r"Stand\s*(\d{4})\s*:\s*([\d.]+)\s*Ex", re.I)
_RX_FREIMENGE = re.compile(
    r"ab\s+dem\s+(\d+)\.\s+von\s+uns\s+verkauft", re.I)


def _menge(text: str) -> int:
    """„2.500" und „2500" sind dieselbe Zahl — der Punkt ist ein Tausenderpunkt."""
    return int(text.replace(".", ""))


def _prozent(text: str) -> float:
    return float(text.replace(",", "."))


def staffel_als_text(staffel: list) -> str:
    """Die Staffel so schreiben, wie sie im Vertrag steht.

    Gegenstück zu ``zerlege_staffel``: was das Werkzeug verstanden hat,
    muss man zurücklesen können — sonst weiß niemand, ob der eingetippte
    Vertragstext richtig angekommen ist.
    """
    if not staffel:
        return ""
    teile, untere = [], 1
    for grenze, satz in staffel:
        if grenze is None:
            teile.append(f"ab {untere} Ex. {satz:g} %")
        else:
            teile.append(f"bis {grenze} Ex. {satz:g} %")
            untere = grenze + 1
    return ", ".join(teile)


def zerlege_staffel(notiz: str) -> tuple[list, int | None, int | None, int]:
    """(Stufen, Stand-Jahr, Stand-Menge, Freimenge) aus dem Notizentext.

    Stufen als [(Obergrenze, Satz), …]; die letzte Stufe hat die Obergrenze
    None und gilt nach oben offen.
    """
    if not notiz:
        return [], None, None, 0

    # Zwei Arten von Angabe, die verschieden zu lesen sind:
    #   „bis 3500 Ex. 12 %"  nennt die OBERGRENZE dieser Stufe.
    #   „ab 3501 Ex. 13 %"   nennt die UNTERGRENZE dieser Stufe — und legt
    #                        damit die Obergrenze der VORIGEN auf 3500 fest.
    # Beides in einen Topf zu werfen ergibt die falsche Staffel: aus
    # „bis 3500 12 %, ab 3501 13 %, ab 5001 15 %" würde sonst eine Stufe
    # bis 3501 statt bis 5000.
    angaben: list[tuple[int, float, str]] = []
    for m in _RX_BIS.finditer(notiz):
        angaben.append((_menge(m.group(1)), _prozent(m.group(2)), "bis"))
    for m in _RX_AB.finditer(notiz):
        angaben.append((_menge(m.group(1)), _prozent(m.group(2)), "ab"))
    angaben.sort(key=lambda a: (a[0], a[2] == "ab"))

    stufen: list[tuple[int | None, float]] = []
    for grenze, satz, art in angaben:
        if art == "ab":
            # Die vorige Stufe endet genau davor.
            if stufen:
                if stufen[-1][1] == satz:
                    continue          # dieselbe Stufe, nur anders formuliert
                stufen[-1] = (grenze - 1, stufen[-1][1])
            stufen.append((None, satz))
        else:
            if stufen and stufen[-1][1] == satz:
                stufen[-1] = (grenze, satz)
            else:
                stufen.append((grenze, satz))
    if stufen:
        # Die oberste Stufe gilt nach oben offen.
        stufen[-1] = (None, stufen[-1][1])
        # Eine einzelne Stufe ist keine Staffel, sondern ein fester Satz.
        if len(stufen) < 2:
            stufen = []

    m = _RX_STAND.search(notiz)
    stand_jahr = int(m.group(1)) if m else None
    stand_menge = _menge(m.group(2)) if m else None

    m = _RX_FREIMENGE.search(notiz)
    # „ab dem 201. verkauften Exemplar" heißt: die ersten 200 sind frei.
    freimenge = int(m.group(1)) - 1 if m else 0

    return stufen, stand_jahr, stand_menge, freimenge


# ---------------------------------------------------------------------
# Import der Altmappe
# ---------------------------------------------------------------------
# Einmalig auszuführen. Liest `Honorare_<Jahr>.xlsx` und baut daraus den
# neuen Bestand. Jede Unklarheit landet im Protokoll und als Nachpflege-
# Vermerk am Buch — repariert wird nichts, denn was richtig ist, entscheidet
# der Verlag und nicht das Werkzeug.

ALT_BLATT_HAUPT = "2025"
ALT_BLATT_HISTORIE = "Zahlung ab XX Ex."
ALT_BLATT_SONDER = "Sonderfälle"
ALT_BLATT_ARCHIV = "keine Zahlung mehr"
ALT_BLATT_LUBW = "LUBW"

# Ohne diese fünf Spalten ist ein Stammblatt nicht als solches zu erkennen.
ALT_PFLICHT = {"Vorname", "Name", "BUCHTITEL", "ISBN", "VERGÜT-ART"}

_RX_JAHRSPALTE = re.compile(
    r"^(?:verk\.?|verkaufte|tats\.?)\s*ex\.?\s*(\d{4})$")
_RX_NUR_JAHR = re.compile(r"^(\d{4})$")


def finde_spalte(idx: dict[str, int], *kandidaten: str) -> int | None:
    """Spaltenindex über mehrere mögliche Schreibweisen.

    Die Altmappe schreibt dieselbe Spalte je Blatt anders — „HONORARper Ex."
    hier, „HONORAR\\nper Ex." dort. Deshalb erst exakt suchen, dann als
    Wortanfang.
    """
    for k in kandidaten:
        i = idx.get(norm_kopf(k))
        if i is not None:
            return i
    for k in kandidaten:
        n = norm_kopf(k).replace(" ", "")
        for name, i in idx.items():
            if name.replace(" ", "").startswith(n):
                return i
    return None


def _zelle(zeile: tuple, i: int | None):
    if i is None or i >= len(zeile):
        return None
    return zeile[i]


def finde_jahresspalten(kopfzeile: tuple) -> dict[int, dict[str, int]]:
    """Die Jahrgänge eines Blattes: Jahr → {verkauft, eigenkauf, vergütung}.

    Das Historienblatt führt 41 Jahresspalten von 2003 bis 2025, und zwar
    unregelmäßig: 2003 und 2004 stehen als nackte Jahreszahl, 2005–2007 als
    Paar, ab 2008 als Dreiergruppe. Deshalb wird von links nach rechts
    gelaufen und jede Spalte dem zuletzt gesehenen Jahr zugeschlagen.
    """
    jahre: dict[int, dict[str, int]] = {}
    aktuell: int | None = None
    for i, roh in enumerate(kopfzeile):
        name = norm_kopf(roh)
        if not name:
            continue
        m = _RX_JAHRSPALTE.match(name) or _RX_NUR_JAHR.match(name)
        if m:
            aktuell = int(m.group(1))
            jahre.setdefault(aktuell, {})["verkauft"] = i
            continue
        if aktuell is None:
            continue
        if name.startswith("eigen"):
            jahre[aktuell].setdefault("eigenkauf", i)
        elif "vergüt" in name and "ex" in name:
            # Steht das Jahr in der Spalte selbst, gilt dieses.
            m2 = re.search(r"(\d{4})", name)
            ziel = int(m2.group(1)) if m2 else aktuell
            jahre.setdefault(ziel, {}).setdefault("verguetung", i)
    return {j: sp for j, sp in jahre.items() if "verkauft" in sp}


def finde_bloecke(ws_formeln, spalte_auszahlung: str) -> list[list[int]]:
    """Welche Zeilen zu einer Auszahlung — also zu einem Brief — gehören.

    Die Altmappe hält das nirgends sauber fest. Die einzige belastbare
    Quelle ist die von Hand gebaute Summenkette in der Auszahlungsspalte:

        =X108+Z108+X109+Z109+…

    Aus ihr werden die Zeilennummern gelesen. Das ist verlässlicher als eine
    Gruppierung nach Namen (dieselbe Person kommt im Bestand mehrfach vor),
    aber nicht fehlerfrei — eine Kette summiert zwei verschiedene Personen.
    Diese Fälle meldet der Import, statt sie stillschweigend zu glätten.
    """
    bloecke = []
    for zeile in ws_formeln.iter_rows(min_col=1, max_col=ws_formeln.max_column):
        for zelle in zeile:
            if zelle.column_letter != spalte_auszahlung:
                continue
            formel = zelle.value
            if not (isinstance(formel, str) and formel.startswith("=")):
                continue
            nummern = sorted({int(n) for n in re.findall(r"[A-Z]+(\d+)", formel)})
            # Summen über ganze Spalten (=SUM(AA2:AA271)) sind Gesamtzeilen,
            # keine Empfänger.
            if nummern and "SUM" not in formel.upper() and "!" not in formel:
                bloecke.append(nummern)
    return bloecke


def _schluessel(vorname: str, name: str, institution: str) -> str:
    return "|".join(" ".join(str(t or "").split()).lower()
                    for t in (vorname, name, institution))


@dataclass
class ImportProtokoll:
    zeilen: list[dict] = field(default_factory=list)
    warnungen: list[str] = field(default_factory=list)

    def merke(self, blatt: str, zeile: int, empfaenger: str, buch: str,
              kondition: str, hinweis: str = "") -> None:
        self.zeilen.append({
            "Blatt": blatt, "Zeile": zeile, "Empfänger": empfaenger,
            "Buch": buch, "Kondition": kondition, "Hinweis": hinweis,
        })

    def warne(self, text: str) -> None:
        self.warnungen.append(text)


def _alt_kopf(ws_v, pflicht: set[str], ab: int = 0):
    """Kopfzeile eines Altblattes suchen und die Spaltenzuordnung liefern."""
    zeilen = [tuple(z) for z in ws_v.iter_rows(values_only=True)]
    kopf_i = finde_kopfzeile(zeilen, pflicht, ab=ab)
    if kopf_i < 0:
        return -1, {}, zeilen
    return kopf_i, spalten_index(zeilen[kopf_i]), zeilen


def _stammfelder(zeile: tuple, idx: dict) -> dict:
    """Die Adressfelder, die in allen Altblättern gleich heißen."""
    akad = ""
    for sp in ("Titel_1", "Titel_2", "Titel_3"):
        t = _text(_zelle(zeile, finde_spalte(idx, sp)))
        if t and t not in akad:
            akad = (akad + " " + t).strip()
    return {
        "lf_nr": _text(_zelle(zeile, finde_spalte(idx, "lf. Nr."))),
        "autorenart": _text(_zelle(zeile, finde_spalte(idx, "Autorenart"))),
        "anrede": _text(_zelle(zeile, finde_spalte(idx, "Anrede"))),
        "titel_akad": akad,
        "vorname": _text(_zelle(zeile, finde_spalte(idx, "Vorname"))),
        "name": _text(_zelle(zeile, finde_spalte(idx, "Name"))),
        "institution": _mehrzeilig(
            _zelle(zeile, finde_spalte(idx, "Institution"))),
        "strasse": _text(_zelle(zeile, finde_spalte(idx, "Straße"))),
        "plz": _text(_zelle(zeile, finde_spalte(idx, "PLZ"))),
        "ort": _text(_zelle(zeile, finde_spalte(idx, "Ort"))),
        "land": _text(_zelle(zeile, finde_spalte(idx, "Land"))),
        "iban": _text(_zelle(zeile, finde_spalte(idx, "Bankverbindung"))),
        "email": _text(_zelle(
            zeile, finde_spalte(idx, "Abrechnung per Mail senden", "E-Mail"))),
    }


def _hole_empfaenger(bestand: Bestand, register: dict, felder: dict
                     ) -> Empfaenger:
    """Empfänger aus dem Register holen oder neu anlegen.

    Dieselbe Person steht in mehreren Altblättern (aktiv, Historie,
    stillgelegt). Zusammengeführt wird über Vorname, Name und Institution —
    genauer geht es nicht, weil es keinen Schlüssel gibt.
    """
    schluessel = _schluessel(felder["vorname"], felder["name"],
                             felder["institution"])
    if schluessel in register:
        e = register[schluessel]
        # Leere Felder aus dem späteren Blatt nachtragen, gefüllte behalten:
        # das zuerst gelesene Blatt ist das jüngere und damit das bessere.
        for k, v in felder.items():
            if v and not getattr(e, k, ""):
                setattr(e, k, v)
        return e
    e = Empfaenger(kennung=bestand.naechste_kennung("E"), **felder)
    # „verschiedene Autoren“ ist kein Mensch, sondern der Sammelposten einer
    # Anthologie. Ein Brief dorthin wäre unzustellbar und eine Überweisung
    # ginge ins Leere.
    if "verschiedene autoren" in _schluessel(
            felder.get("vorname", ""), felder.get("name", ""), ""):
        e.sammelposten = True
    register[schluessel] = e
    bestand.empfaenger.append(e)
    return e


def _richte_zeile_aus(zeile: tuple, idx: dict) -> tuple | None:
    """Eine um eine Spalte verrutschte Zeile geradeziehen.

    Im Blatt „keine Zahlung mehr“ sind 66 der 105 Zeilen ab BUCHTITEL um
    eine Spalte nach rechts verschoben: dort steckt eine namenlose
    Leerspalte, die die Kopfzeile nicht kennt (vermutlich ein „Land“, das
    im Hauptblatt an dieser Stelle steht und hier nie beschriftet wurde).
    Der Adressteil davor sitzt richtig, alles ab dem Buchtitel nicht — und
    damit auch sämtliche Jahresspalten. Ungeprüft übernommen ergäbe das
    Bücher ohne Titel, einen Honorarsatz, der keiner ist, und eine
    Historie, die um ein Jahr danebenliegt.

    Rückgabe: die berichtigte Zeile, oder None, wenn nichts zu tun war.
    """
    i = finde_spalte(idx, "BUCHTITEL")
    if i is None or i + 2 >= len(zeile):
        return None
    if _zelle(zeile, i) is not None:
        return None
    rechts = _zelle(zeile, i + 1)
    if not (isinstance(rechts, str) and rechts.strip()):
        return None
    # Nur wenn die Zelle wirklich LEER ist, darf sie herausfallen —
    # sonst ginge ein Wert verloren statt einer Lücke.
    return zeile[:i] + zeile[i + 1:]


def _lies_buch(bestand: Bestand, zeile: tuple, zeile_f: tuple, idx: dict,
               blatt: str, nr: int, prot: ImportProtokoll) -> Buch | None:
    """Ein Buch samt Kondition aus einer Altzeile."""
    titel = _mehrzeilig(_zelle(zeile, finde_spalte(idx, "BUCHTITEL")))
    isbn = _text(_zelle(zeile, finde_spalte(idx, "ISBN")))
    if not titel and not isbn:
        return None

    i_satz = finde_spalte(idx, "HONORARper Ex.", "HONORAR per Ex.", "HONORAR")
    roh = _zelle(zeile_f, i_satz)
    wert = _komma(_zelle(zeile, i_satz))

    buch = Buch(
        kennung=bestand.naechste_kennung("B"),
        titel=titel,
        isbn=isbn,
        verguetungsart=_text(
            _zelle(zeile, finde_spalte(idx, "VERGÜT-ART"))) or "Honorar",
        mwst_pflichtig=_ja_nein(
            _zelle(zeile, finde_spalte(idx, "MwSt-Pflichtig?", "MwSt-Pflichtig"))),
        notizen=_mehrzeilig(_zelle(zeile, finde_spalte(idx, "Notizen"))),
        quelle=f"{blatt}!Z{nr}",
    )

    # Die eine Altzeile mit 19 % statt 7 % erkennt man nur an ihrer
    # MwSt-Formel, nicht an einer eigenen Spalte.
    i_mwst = finde_spalte(idx, "MwSt.")
    mwst_formel = _zelle(zeile_f, i_mwst)
    if isinstance(mwst_formel, str) and "0.19" in mwst_formel:
        buch.mwst_satz = 19.0

    hinweis = ""
    kond, grund = zerlege_satzformel(roh) if isinstance(roh, str) else (None, "")
    if kond is None:
        # Kein Vertrag herauszulesen — dann wenigstens der Betrag, damit die
        # Abrechnung auf den Cent stimmt. Die Kondition holt der Verlag nach.
        kond = Kondition(betrag_je_ex=wert)
        if isinstance(roh, str) and roh.startswith("="):
            hinweis = grund
            buch.nachpflege.append(f"Satzformel nicht gedeutet: {roh}")
    else:
        # Gegenprobe: erst wenn der zerlegte Vertrag den Zellwert trifft,
        # wird er übernommen. Sonst gilt wieder der nackte Betrag.
        if wert is not None and abs(satz_aus_kondition(kond) - wert) > 0.005:
            hinweis = (f"zerlegt {satz_aus_kondition(kond):.2f} ≠ "
                       f"Mappe {wert:.2f} — Betrag übernommen")
            buch.nachpflege.append(hinweis)
            kond = Kondition(betrag_je_ex=wert)

    kond.schwelle_zehn = (
        _ja_nein(_zelle(
            zeile, finde_spalte(idx, "Keine Berechnung bei 10 o. weniger.")))
        or zehnerregel_in_notiz(buch.notizen))

    # Staffel und Freimenge stehen als Fließtext in den Notizen.
    stufen, stand_jahr, stand_menge, freimenge = zerlege_staffel(buch.notizen)
    if stufen:
        kond.staffel = stufen
        hinweis = (hinweis + "; " if hinweis else "") + \
            f"Staffel aus Notiz: {stufen}"
    if freimenge:
        kond.freimenge = freimenge
    if stand_jahr and stand_menge is not None:
        # Der Stand ist eine Momentaufnahme aus der Notiz, kein gerechneter
        # Wert. Er wird als Vortrag jenes Jahres abgelegt, damit die Staffel
        # überhaupt einen Anknüpfungspunkt hat — und vermerkt, woher er kommt.
        buch.jahr(stand_jahr).vortrag = stand_menge
        buch.nachpflege.append(
            f"kumulierter Stand {stand_menge} Ex. aus Notiz (Stand {stand_jahr})")

    buch.kondition = kond
    prot.merke(blatt, nr, "", titel,
               _kondition_text(kond), hinweis)
    return buch


def _kondition_text(kond: Kondition) -> str:
    if kond.betrag_je_ex is not None:
        return f"fester Betrag {kond.betrag_je_ex:.2f} €/Ex."
    if kond.ladenpreis is None:
        return "ohne Angabe"
    teile = [f"Ladenpreis {kond.ladenpreis:.2f} €"]
    if kond.mwst_im_preis:
        teile.append(("+" if kond.mwst_aufschlagen else "−") +
                     f"{kond.mwst_im_preis:.0f} % MwSt")
    if kond.rabatt_anwenden:
        teile.append(f"−{kond.verlagsrabatt:.0f} % Verlagsrabatt")
    if kond.satz is not None:
        teile.append(f"× {kond.satz:g} %")
    if kond.teiler != 1:
        teile.append(f"÷ {kond.teiler}")
    if kond.staffel:
        teile.append(f"Staffel {kond.staffel}")
    if kond.freimenge:
        teile.append(f"Freimenge {kond.freimenge}")
    return ", ".join(teile)


def _lies_jahre(buch: Buch, zeile: tuple, idx: dict,
                jahresspalten: dict, standard_jahr: int | None = None) -> None:
    """Die Stückzahlen aller Jahrgänge einer Zeile übernehmen."""
    i_korrektur = finde_spalte(
        idx, "Korrektur Vorjahr oder bereits abgerechnet")
    for jahr, sp in jahresspalten.items():
        verkauft = _ganzzahl(_zelle(zeile, sp.get("verkauft")))
        eigenkauf = _ganzzahl(_zelle(zeile, sp.get("eigenkauf")), 0) or 0
        if verkauft is None and not eigenkauf:
            continue
        jw = buch.jahr(jahr)
        jw.verkauft = verkauft
        jw.eigenkauf = eigenkauf
        if jahr == standard_jahr and i_korrektur is not None:
            jw.korrektur = _ganzzahl(_zelle(zeile, i_korrektur), 0) or 0
        i_verg = sp.get("verguetung")
        gezeigt = _ganzzahl(_zelle(zeile, i_verg)) if i_verg is not None else None
        if i_verg is not None and verkauft is not None and gezeigt is None:
            # In einzelnen Altzeilen ist die Vergütungsexemplar-Zelle leer,
            # obwohl Stückzahlen dastehen — dort hat jemand die Formel
            # gelöscht. Excel rechnet deshalb mit 0 weiter, das Werkzeug mit
            # der echten Menge. Der Unterschied ist gewollt, muss aber
            # sichtbar sein.
            buch.nachpflege.append(
                f"Vergütungsexemplare {jahr} waren in der Altmappe leer "
                f"(dort Betrag 0), hier aus {verkauft} Ex. gerechnet")
        elif (gezeigt is not None and verkauft is not None
                and gezeigt != verkauft - eigenkauf):
            # Die Altmappe trägt in dieser Spalte den KUMULIERTEN Stand vor,
            # nicht die Jahresmenge. Ohne ihn wäre die Staffel blind und aus
            # einem Jahr mit Rückgaben („−69 Ex.“) würde ein negativer
            # Saldo, obwohl in Wahrheit 1540 Exemplare aufgelaufen sind.
            jw.vortrag = gezeigt


# Die Verträge im Blatt „Zahlung ab XX Ex.“ formulieren die Schwelle auf
# ein Dutzend Arten: „ab dem 201. von uns verkauften“, „ab dem 300sten“,
# „Ab dem 501. Exemplar“, „ab 501. Ex.“, „ab dem 1.000. Exemplar“,
# „Honorarverzicht bei den ersten 300 Ex.“, „Falls mehr als 1.000 Ex. …“.
# _RX_FREIMENGE kennt nur die erste; damit hatten 24 von 68 Büchern eine
# Freimenge. Bewusst NUR für dieses Blatt: im Hauptblatt stehen dieselben
# Sätze bei Büchern, die die Schwelle längst hinter sich haben und keine
# Vorgeschichte mitbringen — eine Freimenge dort striche ihr Honorar.
_ZAHL = r"(\d{1,3}(?:\.\d{3})+|\d+)"
_RX_SCHWELLE_AB = re.compile(
    r"\bab\s+(?:dem\s+)?" + _ZAHL + r"\s*(?:\.|sten\b|ten\b)?\s*"
    r"(?=(?:\S+\s+){0,2}?(?:Ex\b|Ex\.|Exemplar|von\s+uns|verkauft|"
    r"für\s+jedes|,|das\s+der))", re.I)
_RX_SCHWELLE_ERSTE = re.compile(
    r"(?:bei|für)\s+den\s+ersten\s+" + _ZAHL + r"\s*(?:Ex|Exemplar|verkauft)",
    re.I)
_RX_SCHWELLE_MEHR = re.compile(r"mehr\s+als\s+" + _ZAHL + r"\s*Ex", re.I)
_SCHWELLE_WORTE = {"zweihundertersten": 201, "dreihundertersten": 301,
                   "fünfhundertersten": 501, "tausendersten": 1001}
# „Honorar für E-Book erst, wenn für das Buch Honorar fällig ist“: die
# Schwelle hängt an einem ANDEREN Buch. Das kann die Freimenge nicht
# abbilden — solche Titel bleiben ein Sonderfall.
_RX_SCHWELLE_FREMD = re.compile(r"erst,?\s+wenn", re.I)


def freimenge_aus_notiz(notiz: str) -> int | None:
    """Die Freimenge aus einem Vertragstext, wörtlich gelesen.

    „ab dem 300sten“ heißt: 299 sind frei. So hat es auch die Altmappe
    gerechnet — bei allen Laufzetteln, deren Beginn sich nachrechnen lässt,
    stimmt die wörtliche Lesart (bis auf Mörderliebchen, siehe _laufzettel_einrichten).
    """
    text = notiz or ""
    if m := _RX_SCHWELLE_ERSTE.search(text):
        return int(m.group(1).replace(".", ""))
    if m := _RX_SCHWELLE_AB.search(text):
        n = int(m.group(1).replace(".", ""))
        return n - 1 if n > 1 else None      # „ab dem 1.“ ist keine Schwelle
    for wort, n in _SCHWELLE_WORTE.items():
        if wort in text.lower():
            return n - 1
    if m := _RX_SCHWELLE_MEHR.search(text):
        return int(m.group(1).replace(".", ""))
    return None


# „Falls sich der Verkauf auf weniger als 10 Ex. pro Kalenderjahr
# reduziert, entfällt das Honorar“ — in 81 Notizen, aber nur bei 30 war in
# der Altmappe auch die Spalte angekreuzt. Ob die übrigen bezahlt wurden,
# war Ermessen. Der Verlag hat entschieden (Oktober 2026): es gilt, was im
# Vertrag steht — der Haken folgt der Notiz.
_RX_ZEHNERREGEL = re.compile(
    r"(weniger|unter)\s+(als\s+)?(10|zehn)\b"
    r"|\b(10|zehn)\s*(Ex\.?|Exemplare?)?\s*(o\.|oder)\s*weniger", re.I)


def zehnerregel_in_notiz(notiz: str) -> bool:
    return bool(_RX_ZEHNERREGEL.search(notiz or ""))


def setze_zehnerregel(bestand: Bestand) -> list[str]:
    """Für einen schon importierten Bestand: den Haken „kein Honorar unter
    zehn Exemplaren“ überall setzen, wo die Notiz es sagt. Nur setzen, nie
    entfernen. Gibt die betroffenen Titel zurück."""
    geaendert = []
    for e, b in bestand.buecher():
        if not b.kondition.schwelle_zehn and zehnerregel_in_notiz(b.notizen):
            b.kondition.schwelle_zehn = True
            geaendert.append(f"{e.anzeigename} — {b.titel}")
    return geaendert


def _laufzettel_beginn(buch: Buch) -> int | None:
    """Mit welcher Freimenge der Laufzettel der Altmappe begonnen hat.

    Im ersten Jahr mit Saldo ist verkauft − Saldo genau die Freimenge,
    sofern davor nichts verkauft wurde. Sonst lässt es sich nicht sagen.
    """
    jahre = sorted(buch.jahre)
    erst = next((j for j in jahre if buch.jahre[j].vortrag is not None), None)
    if erst is None:
        return None
    if any(buch.jahre[j].verkauft for j in jahre if j < erst):
        return None
    jw = buch.jahre[erst]
    if jw.verkauft is None:
        return None
    beginn = jw.verkauft - jw.eigenkauf - jw.vortrag
    return beginn if beginn > 0 else None


def _laufzettel_einrichten(buch: Buch) -> str:
    """Ein Buch aus „Zahlung ab XX Ex.“ auf eine gewöhnliche Freimenge setzen.

    Das Blatt ist kein Sonderweg, sondern der Laufzettel für Verträge, die
    erst ab einer bestimmten Stückzahl Honorar zahlen — die Freimenge
    leistet genau das. „Gesondert“ bleibt nur, wessen Schwelle sich nicht
    als Zahl fassen lässt (ab der zweiten Auflage, Garantiehonorar, E-Book
    erst mit dem gedruckten Buch). Gibt zurück, was ins Protokoll gehört, oder "".
    """
    notiz = buch.notizen or ""
    laut_vertrag = freimenge_aus_notiz(notiz)
    if laut_vertrag is None or _RX_SCHWELLE_FREMD.search(notiz):
        return ""
    beginn = _laufzettel_beginn(buch)
    # Die Schwelle steht in der Notiz, und die gilt (Entscheidung des
    # Verlags, Oktober 2026). Bei „Mörderliebchen“ sagt die Notiz „ab dem
    # 200sten“, der Laufzettel der Altmappe begann aber 2017 bei 51
    # verkauften mit „noch 199“ — also mit 250 freien. Der Laufzettel wird
    # deshalb um den Unterschied verschoben: Ende 2024 fehlen dann noch 120
    # Exemplare statt 171. Bei sechs weiteren Büchern macht das ein Exemplar.
    buch.kondition.freimenge = laut_vertrag
    if beginn is not None and beginn != laut_vertrag:
        for jw in buch.jahre.values():
            if jw.vortrag is not None:
                jw.vortrag += beginn - laut_vertrag
    # Startpunkt vor dem ersten Jahr: gezählt wird von null an. Ohne ihn
    # fände die Freimenge im ersten Jahr keinen Anknüpfungspunkt, und es
    # würde ab dem ersten Exemplar gezahlt — bei „Wandern in den Rheinauen“
    # 280,90 € für 265 von 500 freien Exemplaren. Die Altmappe trägt den
    # Saldo erst NACH dem Jahr ein, für das erste Jahr hilft er also nicht.
    # Spätere Salden bleiben der Anker; der Startpunkt ändert nur das
    # erste Jahr.
    if buch.jahre:
        erstes = min(buch.jahre)
        buch.jahr(erstes - 1).vortrag = -buch.kondition.freimenge
    buch.gesondert = False
    if beginn is not None and abs(beginn - laut_vertrag) > 1:
        return (f"„{buch.titel}“: Die Notiz nennt Honorar ab dem "
                f"{laut_vertrag + 1}. Exemplar, der Laufzettel der Altmappe "
                f"hat mit {beginn} freien gezählt. Die Notiz gilt; der Stand "
                f"ist um {beginn - laut_vertrag} Exemplare angepasst.")
    return ""


def stelle_laufzettel_um(bestand: Bestand) -> tuple[int, list[str]]:
    """Für einen schon importierten Bestand: die Laufzettel-Bücher umstellen.

    Berührt nur Bücher aus „Zahlung ab XX Ex.“, die noch auf „gesondert“
    stehen. Gibt die Zahl der umgestellten Bücher und die Prüfhinweise zurück.
    """
    n, hinweise = 0, []
    for _, b in bestand.buecher():
        if not (b.gesondert and b.quelle.startswith(ALT_BLATT_HISTORIE)):
            continue
        hinweis = _laufzettel_einrichten(b)
        if not b.gesondert:
            n += 1
        if hinweis:
            hinweise.append(hinweis)
    return n, hinweise


def _blatt_zeilenweise(bestand, register, wbf, wbv, blattname, prot,
                       stillgelegt=False, standard_jahr=None, ab=0,
                       gesondert=False):
    """Ein Stammblatt Zeile für Zeile einlesen (ohne Blockbildung).

    Für die Blätter, in denen jede Zeile für sich steht: Historie, Archiv,
    Sonderfälle. Das Hauptblatt braucht die Blockbildung und geht einen
    eigenen Weg.
    """
    if blattname not in wbv.sheetnames:
        prot.warne(f"Blatt „{blattname}“ fehlt in der Altmappe.")
        return 0
    ws_v, ws_f = wbv[blattname], wbf[blattname]
    kopf_i, idx, zeilen = _alt_kopf(ws_v, ALT_PFLICHT, ab=ab)
    if kopf_i < 0:
        prot.warne(f"Blatt „{blattname}“: keine Kopfzeile gefunden.")
        return 0
    jahresspalten = finde_jahresspalten(zeilen[kopf_i])
    i_grund = finde_spalte(idx, "Begründung")
    i_voraus = finde_spalte(idx, "Restbetrag", "Vorauszahlung")

    anzahl = 0
    verrutscht: list[int] = []
    for r in range(kopf_i + 2, ws_v.max_row + 1):
        zeile = tuple(c.value for c in ws_v[r])
        if not any(w is not None for w in zeile):
            continue
        # Eine zweite Kopfzeile bedeutet: hier beginnt ein neuer Block mit
        # anderer Spalteneinteilung (so ist „Sonderfälle“ gebaut).
        if norm_kopf(_zelle(zeile, finde_spalte(idx, "Vorname"))) == "vorname":
            anzahl += _blatt_zeilenweise(
                bestand, register, wbf, wbv, blattname, prot,
                stillgelegt=stillgelegt, standard_jahr=standard_jahr,
                ab=r - 1, gesondert=gesondert)
            break
        felder = _stammfelder(zeile, idx)
        if not (felder["vorname"] or felder["name"] or felder["institution"]):
            continue
        zeile_f = tuple(c.value for c in ws_f[r])
        gerade = _richte_zeile_aus(zeile, idx)
        if gerade is not None:
            zeile = gerade
            zeile_f = _richte_zeile_aus(zeile_f, idx) or zeile_f
            verrutscht.append(r)
        buch = _lies_buch(bestand, zeile, zeile_f, idx, blattname, r, prot)
        if buch is None:
            continue
        e = _hole_empfaenger(bestand, register, felder)
        buch.gesondert = gesondert
        buch.stillgelegt = stillgelegt
        if stillgelegt:
            buch.stillgelegt_grund = _text(_zelle(zeile, i_grund))
        if i_voraus is not None:
            buch.vorauszahlung = _komma(_zelle(zeile, i_voraus), 0.0) or 0.0
        _lies_jahre(buch, zeile, idx, jahresspalten, standard_jahr)
        if gesondert:
            hinweis = _laufzettel_einrichten(buch)
            if hinweis:
                prot.warne(f"{blattname}, Zeile {r}: {hinweis}")
        e.buecher.append(buch)
        prot.zeilen[-1]["Empfänger"] = e.anzeigename
        anzahl += 1
    if verrutscht:
        prot.warne(
            f"Blatt „{blattname}“: {len(verrutscht)} Zeilen sind ab der "
            f"Spalte BUCHTITEL um eine Spalte verschoben (Zeilen "
            f"{verrutscht[0]}–{verrutscht[-1]}). Sie wurden beim Lesen "
            f"geradegezogen; ohne das hätten diese Bücher keinen Titel und "
            f"eine um ein Jahr verschobene Historie.")
    return anzahl


def _hauptblatt(bestand, register, wbf, wbv, blattname, jahr, prot):
    """Das Abrechnungsblatt — hier entscheidet die Summenkette die Gruppierung."""
    ws_v, ws_f = wbv[blattname], wbf[blattname]
    kopf_i, idx, zeilen = _alt_kopf(ws_v, ALT_PFLICHT)
    if kopf_i < 0:
        prot.warne(f"Blatt „{blattname}“: keine Kopfzeile gefunden.")
        return 0
    jahresspalten = finde_jahresspalten(zeilen[kopf_i])
    if not jahresspalten:
        jahresspalten = {jahr: {
            "verkauft": finde_spalte(idx, f"verk. Ex. {jahr}", "verk. Ex."),
            "eigenkauf": finde_spalte(idx, "Eigenkauf"),
        }}

    i_aus = finde_spalte(idx, "Auszahlungs-Betrag")
    spalte_aus = ws_v.cell(row=1, column=i_aus + 1).column_letter
    bloecke = finde_bloecke(ws_f, spalte_aus)

    zu_block: dict[int, list[int]] = {}
    for b in bloecke:
        for r in b:
            zu_block.setdefault(r, b)

    erledigt: set[int] = set()
    anzahl = 0
    for r in range(kopf_i + 2, ws_v.max_row + 1):
        if r in erledigt:
            continue
        zeilennummern = zu_block.get(r, [r])
        # Zeilen, die in keiner Summenkette vorkommen, stehen für sich. Das
        # ist meistens ein Versehen in der Altmappe, deshalb der Vermerk.
        allein = r not in zu_block

        saetze = []
        for nr in zeilennummern:
            if nr < kopf_i + 2 or nr > ws_v.max_row:
                continue
            zeile = tuple(c.value for c in ws_v[nr])
            if not any(w is not None for w in zeile):
                continue
            saetze.append((nr, zeile))
        if not saetze:
            erledigt.update(zeilennummern)
            continue

        felder_je_zeile = [(nr, _stammfelder(z, idx)) for nr, z in saetze]
        # Eine Summenkette darf nur EINE Person umfassen. Tut sie es nicht,
        # ist das ein Fehler in der Altmappe — gemeldet, nicht geglättet.
        schluessel = {_schluessel(f["vorname"], f["name"], f["institution"])
                      for _, f in felder_je_zeile if f["name"] or f["institution"]}
        gemischt = len(schluessel) > 1
        if gemischt:
            namen = " / ".join(sorted(
                f["name"] or f["institution"] for _, f in felder_je_zeile))
            prot.warne(
                f"{blattname}, Zeilen {zeilennummern}: die Summenkette in "
                f"{spalte_aus} fasst verschiedene Personen zusammen ({namen}). "
                f"Sie wurden trotzdem als EIN Empfänger übernommen, so wie die "
                f"Altmappe es getan hat — bitte prüfen.")

        leit = next((f for _, f in felder_je_zeile
                     if f["name"] or f["institution"]), felder_je_zeile[0][1])
        e = _hole_empfaenger(bestand, register, leit)
        if allein:
            prot.warne(
                f"{blattname}, Zeile {r}: steht in keiner Summenkette "
                f"({e.anzeigename}) — als eigener Empfänger übernommen.")
            allein_hinweis = (
                f"Diese Zeile stand in der Altmappe in keiner Summenkette; "
                f"sie bekommt hier einen eigenen Brief. Bitte prüfen, ob das "
                f"richtig ist.")
        else:
            allein_hinweis = ""

        for nr, zeile in saetze:
            zeile_f = tuple(c.value for c in ws_f[nr])
            buch = _lies_buch(bestand, zeile, zeile_f, idx, blattname, nr, prot)
            if buch is None:
                continue
            if allein_hinweis:
                buch.nachpflege.append(allein_hinweis)
            if gemischt:
                # Sonst sieht der Bediener später nur einen Betrag und
                # erfährt nie, dass die Altmappe ihn einem anderen gutschrieb.
                # Genau dieser Fall braucht einen Blick.
                buch.nachpflege.append(
                    f"In der Altmappe wurde dieses Buch über die Summenkette "
                    f"in {spalte_aus}{max(zeilennummern)} mit einer ANDEREN "
                    f"Person zusammen abgerechnet ({namen}). Wem der Betrag "
                    f"zusteht, muss der Verlag entscheiden.")
            _lies_jahre(buch, zeile, idx, jahresspalten, jahr)
            e.buecher.append(buch)
            prot.zeilen[-1]["Empfänger"] = e.anzeigename
            anzahl += 1
        erledigt.update(zeilennummern)
    return anzahl


def _anthologie(bestand, register, wbv, blattname, jahr, prot):
    """Ein Anthologie-Blatt: ein Topfbetrag, aufgeteilt auf die Beitragenden.

    Modelliert als ein Empfänger je Person mit genau einem „Buch“, dessen
    Betrag je Exemplar der Anteil ist und dessen Menge 1 beträgt. So braucht
    es keinen eigenen Sonderweg durch die ganze Rechnung.
    """
    if blattname not in wbv.sheetnames:
        return 0
    ws = wbv[blattname]
    kopf_i, idx, zeilen = _alt_kopf(ws, {"Name", "Bankverbindung"})
    if kopf_i < 0:
        prot.warne(f"Blatt „{blattname}“: keine Kopfzeile gefunden.")
        return 0
    i_name = finde_spalte(idx, "Name")
    i_ort = finde_spalte(idx, "Ort")
    i_strasse = finde_spalte(idx, "Straße")
    i_bank = finde_spalte(idx, "Bankverbindung")
    i_web = finde_spalte(idx, "Web")
    anzahl = 0
    for r in range(kopf_i + 2, ws.max_row + 1):
        zeile = tuple(c.value for c in ws[r])
        name = _text(_zelle(zeile, i_name))
        if not name or name.lower().startswith("summe"):
            continue
        # Der Anteil steht in der Spalte rechts neben der Bankverbindung.
        betrag = None
        for i in range(len(zeile)):
            if isinstance(zeile[i], (int, float)) and i > (i_bank or 0):
                betrag = float(zeile[i])
                break
        # „Nachname, Vorname“ oder „Vorname Nachname“ — beides kommt vor.
        if "," in name:
            nachname, _, vorname = name.partition(",")
        else:
            teile = name.rsplit(" ", 1)
            vorname, nachname = (teile[0], teile[1]) if len(teile) == 2 else ("", name)
        ort = _text(_zelle(zeile, i_ort))
        plz, _, ortsname = ort.partition(" ")
        felder = {
            "lf_nr": "", "autorenart": "", "anrede": "", "titel_akad": "",
            "vorname": vorname.strip(), "name": nachname.strip(),
            "institution": "", "strasse": _text(_zelle(zeile, i_strasse)),
            "plz": plz if plz.isdigit() else "",
            "ort": ortsname.strip() if plz.isdigit() else ort,
            "land": "", "iban": _text(_zelle(zeile, i_bank)),
            "email": _text(_zelle(zeile, i_web)),
        }
        e = _hole_empfaenger(bestand, register, felder)
        buch = Buch(
            kennung=bestand.naechste_kennung("B"),
            titel=blattname, isbn="", verguetungsart="Honorar",
            kondition=Kondition(betrag_je_ex=betrag),
            quelle=f"{blattname}!Z{r}",
            notizen=(f"Beteiligt an der Anthologie „{blattname}“. Der "
                     f"Topfbetrag wird unter allen Beteiligten geteilt; "
                     f"zuletzt ergab das {betrag} € je Person."),
            nachpflege=[
                "Anthologie: der Topfbetrag wird nicht automatisch verteilt. "
                "Wenn in diesem Jahr ausgeschüttet werden soll, den Anteil "
                "je Person von Hand eintragen."],
        )
        # KEINE Zahlung für das laufende Jahr erfinden. Der Topf der
        # Altmappe war längst ausgeschüttet — bei „Tödliche Häppchen“ am
        # 27.01.2022 —, und ein Import, der daraus 15 neue Auszahlungen und
        # 15 Briefe macht, verschickt Geld ein zweites Mal.
        buch.jahr(jahr).verkauft = 0
        e.buecher.append(buch)
        prot.merke(blattname, r, e.anzeigename, blattname,
                   f"Anteil {betrag}" if betrag else "kein Anteil", "")
        anzahl += 1
    return anzahl


def _zahlungslisten_satz(wbv, betrag: float) -> dict | None:
    """Den Empfänger zu einem Betrag aus der alten Zahlungsliste holen.

    Das LUBW-Blatt nennt nur Titel und Beträge, nicht den Zahlungsempfänger.
    Der steht in der Zahlungsliste — mit Name, Bankverbindung und der
    Vergütungsart. Ohne diesen Umweg hieße der Empfänger „LUBW“, hätte keine
    Bankverbindung, und der Brief ginge an den falschen Namen.
    """
    blatt = next((b for b in wbv.sheetnames
                  if b.lower().startswith("zahlungsliste")), None)
    if blatt is None:
        return None
    ws = wbv[blatt]
    kopf_i, idx, _ = _alt_kopf(ws, {"Vorname", "Name", "Institution"})
    if kopf_i < 0:
        return None
    for r in range(kopf_i + 2, ws.max_row + 1):
        zeile = tuple(c.value for c in ws[r])
        wert = _komma(_zelle(zeile, finde_spalte(idx, "Auszahlungs-Betrag")))
        if wert is None or abs(wert - betrag) > 0.005:
            continue
        return {
            "vorname": _text(_zelle(zeile, finde_spalte(idx, "Vorname"))),
            "name": _text(_zelle(zeile, finde_spalte(idx, "Name"))),
            "institution": _mehrzeilig(
                _zelle(zeile, finde_spalte(idx, "Institution"))),
            "iban": _text(_zelle(zeile, finde_spalte(idx, "Bankverbindung"))),
            "aktenzeichen": _mehrzeilig(_zelle(
                zeile, finde_spalte(idx, "Aktenzeichen / Kontoinhaber"))),
            "art": _text(_zelle(zeile, finde_spalte(idx, "VERGÜT-ART"))),
        }
    return None


def _iban_schluessel(iban) -> str:
    return "".join(str(iban or "").split()).upper()


def _aktenzeichen_nachtragen(bestand: Bestand, wbv, prot) -> int:
    """Aktenzeichen und Kontoinhaber aus der alten Zahlungsliste übernehmen.

    Die Abrechnungsblätter haben dafür keine Spalte — die Angabe steht nur
    in der Zahlungsliste, und ohne sie kommt etwa bei einer Stadtkasse die
    Überweisung ohne Buchungszeichen an und lässt sich dort nicht zuordnen.

    Zugeordnet wird über die IBAN, nicht über den Namen: die Zahlungsliste
    schreibt „Ev. Bildungszentrum Hospitalhof“, wo das Hauptblatt
    „Evangelisches …“ hat, und nennt beim Stadtarchiv Karlsruhe keinen
    Ansprechpartner. Über den Namen fänden sich drei von acht, über die IBAN
    sieben — der achte steht im Bestand gar nicht.
    """
    blatt = next((b for b in wbv.sheetnames
                  if b.lower().startswith("zahlungsliste")), None)
    if blatt is None:
        return 0
    ws = wbv[blatt]
    kopf_i, idx, _ = _alt_kopf(ws, {"Bankverbindung",
                                    "Aktenzeichen / Kontoinhaber"})
    if kopf_i < 0:
        prot.warne(f"Blatt „{blatt}“: keine Kopfzeile gefunden — "
                   f"Aktenzeichen wurden nicht übernommen.")
        return 0
    nach_iban: dict[str, list[Empfaenger]] = {}
    for e in bestand.empfaenger:
        if e.iban:
            nach_iban.setdefault(_iban_schluessel(e.iban), []).append(e)

    anzahl = 0
    for r in range(kopf_i + 2, ws.max_row + 1):
        zeile = tuple(c.value for c in ws[r])
        az = _mehrzeilig(_zelle(
            zeile, finde_spalte(idx, "Aktenzeichen / Kontoinhaber")))
        if not az:
            continue
        iban = _text(_zelle(zeile, finde_spalte(idx, "Bankverbindung")))
        wer = " ".join(t for t in (
            _text(_zelle(zeile, finde_spalte(idx, "Vorname"))),
            _text(_zelle(zeile, finde_spalte(idx, "Name"))),
            _mehrzeilig(_zelle(zeile, finde_spalte(idx, "Institution"))))
            if t) or f"Zeile {r}"
        treffer = nach_iban.get(_iban_schluessel(iban), []) if iban else []
        if not treffer:
            prot.warne(
                f"{blatt}, Zeile {r}: „{wer}“ hat das Aktenzeichen „{az}“, "
                f"steht aber mit dieser IBAN in keinem Abrechnungsblatt — "
                f"nicht übernommen. Falls weiter gezahlt wird, als Autor "
                f"von Hand anlegen.")
            continue
        for e in treffer:
            if not e.aktenzeichen:
                e.aktenzeichen = az
                anzahl += 1
            elif e.aktenzeichen != az:
                prot.warne(
                    f"{blatt}, Zeile {r}: bei „{e.anzeigename}“ steht schon "
                    f"„{e.aktenzeichen}“, die Zahlungsliste sagt „{az}“ — "
                    f"das vorhandene bleibt, bitte prüfen.")
    return anzahl


def trage_aktenzeichen_nach(bestand: Bestand, altmappe) -> ImportProtokoll:
    """Für einen schon importierten Bestand: nur die Aktenzeichen nachholen.

    Ein neuer Import wäre der falsche Weg — er verwürfe alles, was seitdem
    eingetragen wurde. Hier werden ausschließlich LEERE Aktenzeichen gefüllt.
    """
    prot = ImportProtokoll()
    wbv = openpyxl.load_workbook(Path(altmappe), data_only=True)
    try:
        n = _aktenzeichen_nachtragen(bestand, wbv, prot)
    finally:
        wbv.close()
    prot.warne(f"{n} Aktenzeichen aus der Zahlungsliste übernommen.")
    return prot


def _lubw(bestand, register, wbv, jahr, prot):
    """Das LUBW-Blatt: ein Empfänger, viele Titel, fester Betrag je Exemplar."""
    if ALT_BLATT_LUBW not in wbv.sheetnames:
        return 0
    ws = wbv[ALT_BLATT_LUBW]
    kopf_i, idx, zeilen = _alt_kopf(ws, {"ISBN", "Titel", "Einzelbetrag"})
    if kopf_i < 0:
        prot.warne(f"Blatt „{ALT_BLATT_LUBW}“: keine Kopfzeile gefunden.")
        return 0

    felder = {k: "" for k in ("lf_nr", "autorenart", "anrede", "titel_akad",
                              "vorname", "name", "strasse", "plz", "ort",
                              "land", "iban", "email", "aktenzeichen")}
    felder["institution"] = "LUBW"
    felder["anrede"] = "Damen und Herren"
    art = "Rückfluss"

    # Den echten Zahlungsempfänger über den Gesamtbetrag in der alten
    # Zahlungsliste suchen.
    i_gesamt = finde_spalte(idx, "Gesamtbetrag")
    summe = 0.0
    for r in range(kopf_i + 2, ws.max_row + 1):
        wert = _komma(_zelle(tuple(c.value for c in ws[r]), i_gesamt))
        if wert is not None:
            summe += wert
    summe = round(summe / 2, 2)   # die Summenzeile zählt alles doppelt
    satz = _zahlungslisten_satz(wbv, summe)
    if satz:
        for schluessel in ("vorname", "name", "institution", "iban",
                           "aktenzeichen"):
            if satz.get(schluessel):
                felder[schluessel] = satz[schluessel]
        art = satz.get("art") or art
        prot.warne(
            f"{ALT_BLATT_LUBW}: Zahlungsempfänger aus der Zahlungsliste "
            f"übernommen — „{felder['institution'] or felder['name']}“, "
            f"Vergütungsart „{art}“. Das LUBW-Blatt selbst nennt keinen "
            f"Empfänger; bitte prüfen, ob das stimmt.")
    else:
        prot.warne(
            f"{ALT_BLATT_LUBW}: In der alten Zahlungsliste war kein "
            f"Empfänger zum Gesamtbetrag von {summe:.2f} € zu finden. Der "
            f"Empfänger heißt vorläufig „LUBW“ und hat keine "
            f"Bankverbindung — bitte nachtragen.")

    e = _hole_empfaenger(bestand, register, felder)
    i_isbn = finde_spalte(idx, "ISBN")
    i_titel = finde_spalte(idx, "Titel")
    i_einzel = finde_spalte(idx, "Einzelbetrag")
    i_menge = finde_spalte(idx, f"Menge {jahr}", "Menge")
    anzahl = 0
    for r in range(kopf_i + 2, ws.max_row + 1):
        zeile = tuple(c.value for c in ws[r])
        isbn = _text(_zelle(zeile, i_isbn))
        titel = _mehrzeilig(_zelle(zeile, i_titel))
        einzel = _komma(_zelle(zeile, i_einzel))
        if not isbn or einzel is None:
            continue
        buch = Buch(
            kennung=bestand.naechste_kennung("B"),
            titel=titel, isbn=isbn, verguetungsart=art,
            kondition=Kondition(betrag_je_ex=einzel),
            quelle=f"{ALT_BLATT_LUBW}!Z{r}",
            notizen="Pauschale je Exemplar laut LUBW-Vereinbarung.",
        )
        buch.jahr(jahr).verkauft = _ganzzahl(_zelle(zeile, i_menge), 0) or 0
        e.buecher.append(buch)
        prot.merke(ALT_BLATT_LUBW, r, e.anzeigename, titel,
                   f"fester Betrag {einzel:.2f} €/Ex.", "")
        anzahl += 1
    return anzahl


def importiere_alt(pfad, jahr: int = 2025, log=None) -> tuple[Bestand, ImportProtokoll]:
    """Die alte Honorar-Mappe einlesen und daraus den neuen Bestand bauen.

    Reihenfolge mit Absicht: erst das Archiv, dann Historie und Sonderfälle,
    zuletzt das Abrechnungsblatt. So gewinnt immer das jüngste Blatt, wenn
    dieselbe Person mehrfach vorkommt — beim Zusammenführen werden nur noch
    LEERE Felder nachgetragen.
    """
    def melde(text):
        if log:
            log(text)

    pfad = Path(pfad)
    if not pfad.exists():
        raise ValueError(f"Die Altmappe wurde nicht gefunden:\n{pfad}")

    melde(f"Lese {pfad.name} …")
    wbv = openpyxl.load_workbook(pfad, data_only=True)
    wbf = openpyxl.load_workbook(pfad, data_only=False)
    try:
        bestand = Bestand()
        prot = ImportProtokoll()
        register: dict[str, Empfaenger] = {}

        n = _blatt_zeilenweise(bestand, register, wbf, wbv, ALT_BLATT_ARCHIV,
                               prot, stillgelegt=True)
        melde(f"  {ALT_BLATT_ARCHIV}: {n} stillgelegte Bücher")

        n = _blatt_zeilenweise(bestand, register, wbf, wbv, ALT_BLATT_HISTORIE,
                               prot, standard_jahr=jahr, gesondert=True)
        melde(f"  {ALT_BLATT_HISTORIE}: {n} Bücher mit Jahreshistorie")
        # Das Blatt ist der Laufzettel für Verträge mit Freimenge. Wo die
        # Schwelle eine Zahl ist, wird es eine gewöhnliche Freimenge; nur
        # der Rest bleibt „gesondert“ und wartet auf den Verlag.
        rest = [b for _, b in bestand.buecher()
                if b.gesondert and b.quelle.startswith(ALT_BLATT_HISTORIE)]
        prot.warne(
            f"{ALT_BLATT_HISTORIE}: {n - len(rest)} von {n} Büchern zahlen "
            f"ab einer festen Stückzahl und laufen als Freimenge mit. "
            f"{len(rest)} sind als „gesondert abrechnen“ übernommen und "
            f"lösen KEINE Auszahlung aus — ihre Schwelle ist keine Zahl "
            f"(ab der zweiten Auflage, Garantiehonorar, E-Book erst mit dem "
            f"Buch). Wie sie abgerechnet werden, muss der Verlag entscheiden.")

        n = _blatt_zeilenweise(bestand, register, wbf, wbv, ALT_BLATT_SONDER,
                               prot, standard_jahr=jahr)
        melde(f"  {ALT_BLATT_SONDER}: {n} Bücher")

        blatt_haupt = str(jahr) if str(jahr) in wbv.sheetnames else ALT_BLATT_HAUPT
        n = _hauptblatt(bestand, register, wbf, wbv, blatt_haupt, jahr, prot)
        melde(f"  {blatt_haupt}: {n} Bücher")

        for blatt in wbv.sheetnames:
            if blatt.lower().startswith("tödliche"):
                n = _anthologie(bestand, register, wbv, blatt, jahr, prot)
                melde(f"  {blatt}: {n} Beitragende")

        n = _lubw(bestand, register, wbv, jahr, prot)
        melde(f"  {ALT_BLATT_LUBW}: {n} Titel")

        n = _aktenzeichen_nachtragen(bestand, wbv, prot)
        melde(f"  Zahlungsliste: {n} Aktenzeichen übernommen")

        melde(f"Fertig: {len(bestand.empfaenger)} Empfänger, "
              f"{sum(1 for _ in bestand.buecher())} Bücher, "
              f"{len(prot.warnungen)} Warnungen.")
        return bestand, prot
    finally:
        wbv.close()
        wbf.close()


def schreibe_importprotokoll(prot: ImportProtokoll, ziel: Path) -> Path:
    """Die Spur, an der später nachvollziehbar ist, warum etwas so importiert
    wurde. Ohne sie ist eine Zahl im Bestand nicht mehr zu erklären."""
    ziel = Path(ziel)
    ziel.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    spalten = ["Blatt", "Zeile", "Empfänger", "Buch", "Kondition", "Hinweis"]
    ws = wb.create_sheet("Zeilen")
    ws.append(spalten)
    for z in ws[1]:
        z.font = Font(bold=True)
    for zeile in prot.zeilen:
        ws.append([zeile.get(s) for s in spalten])
    ws.freeze_panes = "A2"
    for i, s in enumerate(spalten, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = \
            {"Blatt": 20, "Zeile": 8, "Empfänger": 34, "Buch": 40,
             "Kondition": 52, "Hinweis": 60}.get(s, 16)

    ws2 = wb.create_sheet("Warnungen")
    ws2.append(["Warnung"])
    ws2["A1"].font = Font(bold=True)
    ws2.column_dimensions["A"].width = 130
    for w in prot.warnungen:
        ws2.append([w])
        ws2.cell(row=ws2.max_row, column=1).alignment = Alignment(wrapText=True)
    ws2.freeze_panes = "A2"

    wb.save(ziel)
    wb.close()
    return ziel


# ---------------------------------------------------------------------
# Berechnung
# ---------------------------------------------------------------------

@dataclass
class Posten:
    """Eine Zeile der Titeltabelle im Abrechnungsbrief."""

    buch: Buch
    jahr: int = 0
    satz: float = 0.0            # Betrag je Exemplar in Euro
    satz_prozent: float | None = None   # der angewandte Honorarsatz
    verkauft: int = 0
    eigenkauf: int = 0
    korrektur: int = 0
    verguetungs_ex: int = 0
    netto: float = 0.0
    mwst: float = 0.0

    @property
    def brutto(self) -> float:
        return runde(self.netto + self.mwst)


@dataclass
class Abrechnung:
    """Was ein Empfänger für ein Jahr bekommt — und ob ein Brief hinausgeht."""

    empfaenger: Empfaenger
    jahr: int
    posten: list[Posten] = field(default_factory=list)
    netto: float = 0.0
    mwst: float = 0.0
    brutto: float = 0.0
    ksk_netto: float = 0.0          # nur echte Honorare, für die KSK
    verrechnet: float = 0.0         # gegen offene Vorauszahlungen verrechnet
    brief: bool = False
    gruende: list[str] = field(default_factory=list)   # warum KEIN Brief
    probleme: list[str] = field(default_factory=list)  # was nachzusehen ist

    @property
    def verguetungsart(self) -> str:
        """Die Art für Dateiname und Zahlungsliste.

        Mischt ein Brief mehrere Arten, gewinnt die betragsmäßig größte —
        und die Abrechnung vermerkt es als Problem.
        """
        nach_art: dict[str, float] = {}
        for p in self.posten:
            nach_art[p.buch.verguetungsart] = \
                nach_art.get(p.buch.verguetungsart, 0.0) + p.brutto
        if not nach_art:
            return "Honorar"
        return max(nach_art.items(), key=lambda x: x[1])[0]


def kumulierte_menge(buch: Buch, jahr: int) -> int | None:
    """Kumulierte Vergütungsexemplare VOR dem Abrechnungsjahr.

    Grundlage für die Staffel und für Verträge mit Freimenge. Gibt None
    zurück, wenn es gar keine Vorgeschichte gibt — dann gilt schlicht der
    feste Satz, und es wird nichts vorgetragen.
    """
    vorjahre = sorted(j for j in buch.jahre if j < jahr)
    if not vorjahre:
        return None

    # Jüngster festgehaltener Vortrag ist der Anker; davor muss nichts mehr
    # aufsummiert werden.
    anker = None
    for j in reversed(vorjahre):
        if buch.jahre[j].vortrag is not None:
            anker = j
            break

    if anker is not None:
        summe = buch.jahre[anker].vortrag
        rest = [j for j in vorjahre if j > anker]
        bekannt = True
    else:
        summe = 0
        rest = vorjahre
        bekannt = False

    for j in rest:
        jw = buch.jahre[j]
        if jw.verkauft is None:
            continue
        bekannt = True
        summe += jw.verkauft - jw.eigenkauf + jw.korrektur

    if not bekannt:
        return None
    if anker is None and buch.kondition.freimenge:
        # Die Freimenge wird genau einmal abgezogen, im ersten erfassten Jahr.
        summe -= buch.kondition.freimenge
    return summe


def angewandter_satz(buch: Buch, jahr: int) -> float | None:
    """Welcher Prozentsatz in diesem Jahr gilt — mit Staffel der Stufensatz."""
    kond = buch.kondition
    if kond.staffel:
        stand = kumulierte_menge(buch, jahr)
        if stand is not None:
            for grenze, satz in kond.staffel:
                if grenze is None or stand <= grenze:
                    return satz
    return kond.satz


def satz_je_ex(buch: Buch, jahr: int) -> float:
    """Betrag je Exemplar für dieses Jahr.

    Bei einer Staffel entscheidet der kumulierte Stand ZU JAHRESBEGINN, welche
    Stufe gilt — und dieser eine Satz wird auf die ganze Jahresmenge
    angewandt. So rechnet der Verlag es heute auch von Hand. (Die Alternative
    wäre, die Jahresmenge auf die Stufen aufzuteilen; das ist eine offene
    Frage an den Verlag, siehe README.)
    """
    kond = buch.kondition
    if kond.staffel:
        stand = kumulierte_menge(buch, jahr)
        if stand is not None:
            for grenze, satz in kond.staffel:
                if grenze is None or stand <= grenze:
                    return satz_aus_kondition(kond, satz)
    return satz_aus_kondition(kond)


def verguetungs_ex(buch: Buch, jahr: int) -> int | None:
    """Vergütungsexemplare des Jahres: verkauft − Eigenkauf + Korrektur.

    Bei Verträgen mit Freimenge kommt der kumulierte Saldo hinzu, der
    jahrelang negativ bleiben darf — gezahlt wird erst, wenn er ins Plus
    dreht. Gibt None zurück, wenn das Jahr noch nicht erfasst ist.

    Nur ein NEGATIVER Saldo wird verrechnet. Ist die Freimenge schon
    überschritten, wurden die Exemplare darüber im Jahr des Überschreitens
    bezahlt; sie noch einmal aufzuschlagen hieße, sie jedes Jahr erneut zu
    vergüten.
    """
    jw = buch.jahre.get(jahr)
    if jw is None or jw.verkauft is None:
        return None
    menge = jw.verkauft - jw.eigenkauf + jw.korrektur
    if buch.kondition.freimenge:
        vortrag = kumulierte_menge(buch, jahr)
        if vortrag is not None:
            menge += min(vortrag, 0)
    return menge


def freimenge_stand(buch: Buch, jahr: int) -> int | None:
    """Wie viele Exemplare bis Ende des Vorjahres auf die Freimenge zählen.

    Gespeichert ist der Saldo GEGEN die Freimenge (−171 heißt: noch 171
    bis zum Honorar) — so führt ihn auch der Laufzettel der Altmappe. Für
    Menschen ist die Zahl der bisher verkauften Exemplare greifbarer; sie
    ist Saldo plus Freimenge. None, wenn es keine Vorgeschichte gibt.
    """
    stand = kumulierte_menge(buch, jahr)
    if stand is None or not buch.kondition.freimenge:
        return None
    return stand + buch.kondition.freimenge


def setze_freimenge_stand(buch: Buch, jahr: int, bisher: int) -> None:
    """Den Stand „bis Ende des Vorjahres verkauft“ von Hand festlegen.

    Abgelegt als Vortrag des Vorjahres; der jüngste Vortrag ist der Anker,
    ältere Jahre zählen dann nicht mehr mit.
    """
    buch.jahr(jahr - 1).vortrag = bisher - buch.kondition.freimenge


def betrag_zeile(buch: Buch, jahr: int, cfg: dict | None = None) -> Posten | None:
    """Einen Posten rechnen. None, wenn das Jahr nicht erfasst ist."""
    cfg = cfg or {}
    menge = verguetungs_ex(buch, jahr)
    if menge is None:
        return None
    jw = buch.jahre[jahr]
    satz = satz_je_ex(buch, jahr)

    schwelle = int(cfg.get("schwelle_menge", 10))
    if buch.kondition.freimenge and menge < 0:
        # Die Freimenge ist noch nicht erreicht: von den ersten 500
        # Exemplaren bekommt der Autor nichts. Daraus darf aber KEINE
        # Forderung gegen ihn werden — er schuldet nichts für Exemplare,
        # die nie vergütet wurden. Der Saldo wird vorgetragen, der Betrag
        # dieses Jahres ist null.
        #
        # Ohne diese Regel verschlang ein einzelnes solches Buch die
        # Auszahlung des ganzen Autors: im Bestand 2025 fiel ein Empfänger
        # dadurch von 6,69 € auf −1.038,59 €.
        netto = 0.0
    elif buch.kondition.schwelle_zehn and menge < schwelle:
        # „Keine Berechnung bei 10 o. weniger." — gemeint ist: unter der
        # Schwelle entfällt das Honorar, BEI genau zehn wird gezahlt.
        netto = 0.0
    else:
        netto = runde(menge * satz)

    mwst = runde(netto * buch.mwst_satz / 100) if buch.mwst_pflichtig else 0.0
    return Posten(buch=buch, jahr=jahr, satz=satz,
                  satz_prozent=angewandter_satz(buch, jahr),
                  verkauft=jw.verkauft, eigenkauf=jw.eigenkauf,
                  korrektur=jw.korrektur, verguetungs_ex=menge,
                  netto=netto, mwst=mwst)


def staffel_ohne_stand(buch: Buch, jahr: int) -> str:
    """Meldet, wenn eine Staffel mangels Stand nicht greifen kann.

    Nur dann, wenn der Rueckfallsatz UNTER der untersten Vertragsstufe
    liegt — sonst ist der gespeicherte Satz die von Hand getippte Stufe
    und alles in Ordnung. Ein Cent Unterschied zaehlt nicht, der entsteht
    schon durchs Runden.
    """
    kond = buch.kondition
    if not kond.staffel or kumulierte_menge(buch, jahr) is not None:
        return ""
    unterste = min(satz_aus_kondition(kond, s) for _, s in kond.staffel)
    gilt = satz_aus_kondition(kond)
    if gilt >= unterste - 0.005:
        return ""
    return (f"„{buch.titel}“: Für dieses Buch ist eine Staffel vereinbart, "
            f"aber der bis {jahr} verkaufte Gesamtstand ist nicht bekannt. "
            f"Gerechnet wird mit {euro(gilt)} je Exemplar — die unterste "
            f"Vertragsstufe wäre {euro(unterste)}. Bitte prüfen.")


def rechne_empfaenger(e: Empfaenger, jahr: int,
                      cfg: dict | None = None) -> Abrechnung:
    """Alle Bücher eines Empfängers zu einer Abrechnung zusammenfassen."""
    cfg = cfg or {}
    ab = Abrechnung(empfaenger=e, jahr=jahr)

    for buch in e.buecher:
        if buch.stillgelegt:
            continue
        if buch.gesondert:
            # Zählt für die Historie und die Staffel, aber nicht für die
            # Auszahlung. Der Betrag wird trotzdem ausgewiesen, damit er
            # nicht unbemerkt verschwindet.
            posten = betrag_zeile(buch, jahr, cfg)
            if posten is not None and posten.netto:
                ab.probleme.append(
                    f"„{buch.titel}“ wird gesondert abgerechnet und ist hier "
                    f"NICHT enthalten (rechnerisch {euro(posten.netto)}).")
            continue
        posten = betrag_zeile(buch, jahr, cfg)
        if posten is None:
            if any(j >= jahr - 1 for j in buch.jahre):
                ab.probleme.append(
                    f"„{buch.titel}“: Stückzahl {jahr} noch nicht erfasst.")
            continue
        # Eine vereinbarte Staffel, deren Stand niemand kennt, faellt still
        # auf den gespeicherten Satz zurueck. Meist ist das genau eine der
        # Stufen — der Verlag hat sie von Hand getippt. Manchmal aber nicht:
        # bei „D'accord mit de Welt" steht in der Altmappe
        # =ROUND((17.9/1.07*0.05)/2,2), also 5 % UND nochmal halbiert,
        # waehrend der Vertrag in den Notizen 10 % nennt. Herauskommen 0,42 €
        # statt 0,84 € je Exemplar. Gerechnet wird weiter wie bisher — was
        # richtig ist, entscheidet der Verlag —, aber gesagt wird es.
        hinweis = staffel_ohne_stand(buch, jahr)
        if hinweis:
            ab.probleme.append(hinweis)
        if posten.verguetungs_ex > 0 and not posten.satz:
            # Sechs Laufzettel-Bücher haben in der Altmappe keinen Betrag je
            # Exemplar — solange die Freimenge läuft, fällt das nicht auf.
            # Ist sie erreicht, ergäbe es stillschweigend 0 €.
            ab.probleme.append(
                f"„{buch.titel}“: {posten.verguetungs_ex} Exemplare sind zu "
                f"vergüten, aber für das Buch ist kein Betrag je Exemplar "
                f"hinterlegt — gerechnet wurde 0 €. Bitte im Buch nachtragen.")
        ab.posten.append(posten)
        ab.netto = runde(ab.netto + posten.netto)
        ab.mwst = runde(ab.mwst + posten.mwst)
        if buch.verguetungsart == KSK_ART:
            ab.ksk_netto = runde(ab.ksk_netto + posten.netto)

    ab.brutto = runde(ab.netto + ab.mwst)

    # Offene Vorauszahlungen zuerst verrechnen — ausgezahlt wird nur, was
    # darüber hinausgeht.
    offen = runde(sum(b.vorauszahlung for b in e.buecher
                      if not b.stillgelegt and b.vorauszahlung > 0))
    if offen > 0 and ab.brutto > 0:
        ab.verrechnet = runde(min(offen, ab.brutto))
        ab.brutto = runde(ab.brutto - ab.verrechnet)
        ab.probleme.append(
            f"{ab.verrechnet:.2f} € mit der offenen Vorauszahlung verrechnet "
            f"(danach noch offen: {offen - ab.verrechnet:.2f} €).")

    # Ein Brief geht nur hinaus, wenn es etwas zu berichten gibt.
    if not ab.posten:
        ab.gruende.append("kein erfasstes Buch in diesem Jahr")
    elif ab.brutto < 0:
        # Mehr Rückgaben als Verkäufe. Ausgezahlt wird nichts — aber der
        # Fehlbetrag verschwindet damit auch nicht von selbst. Wer ihn im
        # nächsten Jahr verrechnen will, trägt ihn dort in „Korrektur“ ein;
        # sonst bekommt der Autor im Folgejahr zu viel.
        ab.gruende.append(
            f"Auszahlungsbetrag {euro(ab.brutto)} — mehr Rückgaben als "
            f"Verkäufe, es wird nichts überwiesen")
        ab.probleme.append(
            f"Fehlbetrag von {euro(abs(ab.brutto))}: soll er im nächsten "
            f"Jahr verrechnet werden? Dann dort bei „Korrektur“ eintragen.")
    elif ab.brutto == 0:
        ab.gruende.append("nichts auszuzahlen")
    elif e.sammelposten:
        ab.gruende.append(
            f"Sammelposten, kein einzelner Empfänger — die {ab.brutto:.2f} € "
            f"sind auf die Beteiligten zu verteilen")
    ab.brief = not ab.gruende

    if ab.brief and not e.iban:
        ab.probleme.append("keine Bankverbindung hinterlegt")
    # Nach dem Import kann ein Empfänger Bücher tragen, die in der Altmappe
    # auf getrennten Blättern standen und getrennt abgerechnet wurden. Sie
    # landen jetzt in EINEM Brief — das ist gewollt, aber der Bediener soll
    # es beim ersten Durchlauf sehen und nicht erst beim Nachrechnen.
    blaetter = {p.buch.quelle.split("!")[0] for p in ab.posten if p.buch.quelle}
    if len(blaetter) > 1:
        ab.probleme.append(
            "fasst Bücher aus mehreren Blättern der Altmappe zusammen: "
            + ", ".join(sorted(blaetter)))
    arten = {p.buch.verguetungsart for p in ab.posten}
    if len(arten) > 1:
        ab.probleme.append(
            "mehrere Vergütungsarten in einem Brief: " + ", ".join(sorted(arten)))
    for buch in e.buecher:
        for hinweis in buch.nachpflege:
            ab.probleme.append(f"„{buch.titel}“: {hinweis}")
    return ab


def rechne_alle(bestand: Bestand, jahr: int,
                cfg: dict | None = None) -> list[Abrechnung]:
    return [rechne_empfaenger(e, jahr, cfg) for e in bestand.empfaenger]


def zurueckgehalten(bestand: Bestand, jahr: int,
                    cfg: dict | None = None) -> tuple[int, float]:
    """Wie viel wegen „Gesondert abrechnen“ NICHT ausgezahlt wird.

    Diese Verträge warten auf eine Entscheidung des Verlags. Solange sie
    aussteht, fließt kein Geld — und genau das darf nicht in einem
    Protokoll versanden, das einmal beim Import gelesen wird. Die Zahl
    gehört bei jedem Durchlauf vor Augen.
    """
    cfg = cfg or {}
    empfaenger, summe = set(), 0.0
    for e, b in bestand.buecher():
        if not b.gesondert or b.stillgelegt:
            continue
        posten = betrag_zeile(b, jahr, cfg)
        if posten is None or posten.netto <= 0:
            continue
        empfaenger.add(e.kennung)
        summe = runde(summe + posten.brutto)
    return len(empfaenger), summe


# ---------------------------------------------------------------------
# Gegenprobe an der Altmappe
# ---------------------------------------------------------------------
# Die eigentliche Abnahme. Solange die Altmappe existiert, ist sie die
# Wahrheit: rechnet das Werkzeug für dasselbe Jahr etwas anderes, ist erst zu
# klären, warum — und nicht die Zahl anzupassen, bis sie passt.

# ``satz_runden`` steht nur noch als Spur im Bestand: es merkt sich, dass
# die Altmappe bei diesem Buch NICHT gerundet hat. Gerechnet wird immer
# gerundet (siehe satz_aus_kondition) — die Gegenprobe muss den Unterschied
# erklären können.
_GRUND_GERUNDET = "Satz je Exemplar auf Cent gerundet (Altmappe: ungerundet)"


_GRUND_ZEHN = ("unter zehn Exemplaren kein Honorar — steht in der Notiz, "
               "in der Altmappe nicht angekreuzt (Entscheidung Okt. 2026)")


def _zehnerregel_greift(buch: Buch, jahr: int) -> bool:
    if not (buch.kondition.schwelle_zehn and zehnerregel_in_notiz(buch.notizen)):
        return False
    menge = verguetungs_ex(buch, jahr)
    return menge is not None and 0 < menge < 10


def _empfaengergrund(e: Empfaenger, jahr: int) -> str:
    """Warum die Summe eines Empfängers von der Altmappe abweichen darf."""
    gruende = []
    aktiv = [b for b in e.buecher if not b.stillgelegt]
    if any(not b.kondition.satz_runden for b in aktiv):
        gruende.append(_GRUND_GERUNDET)
    if any(_zehnerregel_greift(b, jahr) for b in aktiv):
        gruende.append(_GRUND_ZEHN)
    return "; ".join(gruende)


def _abweichungsgrund(buch: Buch, jahr: int) -> str:
    """Warum das Werkzeug hier anders rechnet als die Altmappe.

    Die Altmappe wendet weder Staffel noch Freimenge an — beides steht dort
    nur als Fließtext in den Notizen. Wo das Werkzeug den Vertrag jetzt
    wirklich anwendet, MUSS eine Abweichung herauskommen; sie ist der
    gewünschte Fortschritt und kein Fehler. Alles andere ist einer.
    """
    gruende = []
    if buch.kondition.staffel and kumulierte_menge(buch, jahr) is not None:
        gruende.append("Staffel angewandt (Altmappe: fester Satz)")
    if buch.kondition.freimenge and kumulierte_menge(buch, jahr) is not None:
        gruende.append("Freimenge angewandt (Altmappe: ohne Vortrag)")
    if buch.vorauszahlung:
        gruende.append("Vorauszahlung verrechnet")
    if not buch.kondition.satz_runden:
        gruende.append(_GRUND_GERUNDET)
    if _zehnerregel_greift(buch, jahr):
        gruende.append(_GRUND_ZEHN)
    for hinweis in buch.nachpflege:
        if "Altmappe" in hinweis:
            gruende.append(hinweis)
    return "; ".join(gruende)


@dataclass
class Abweichung:
    was: str            # "Zeile" oder "Auszahlung"
    wer: str
    titel: str
    soll: float
    ist: float
    grund: str = ""

    @property
    def differenz(self) -> float:
        return runde(self.ist - self.soll)


def _vergleichsschluessel(vorname: str, name: str, institution: str) -> str:
    """Zwei Schreibweisen derselben Person vergleichbar machen.

    Nur für die Gegenprobe. Die Altmappe schreibt dieselbe Person mal mit,
    mal ohne Institution; verglichen wird deshalb über Vor- und Nachnamen,
    und nur wo beide fehlen über die Institution.
    """
    def sauber(s):
        s = unicodedata.normalize("NFKD", str(s or ""))
        return re.sub(r"[^a-z0-9]", "", s.lower())
    person = sauber(vorname) + "|" + sauber(name)
    return person if person != "|" else "einrichtung:" + sauber(institution)[:24]


def pruefe_zahlungsliste(bestand: Bestand, pfad_alt, jahr: int,
                         cfg: dict | None = None) -> list[Abweichung]:
    """Die Auszahlungen gegen die alte Zahlungsliste stellen.

    Die Rechenprobe allein genügt nicht: sie prüft Beträge, aber nicht, ob
    am Ende die richtigen Leute in der Liste stehen. Genau dort steckten die
    Fehler — ein Empfänger hieß „LUBW“ statt nach der Stiftung, und eine
    Anthologie erzeugte fünfzehn Auszahlungen für Geld, das Jahre zuvor
    geflossen war.
    """
    cfg = cfg or {}
    pfad_alt = Path(pfad_alt)
    wb = openpyxl.load_workbook(pfad_alt, data_only=True)
    abweichungen: list[Abweichung] = []
    try:
        blatt = next((b for b in wb.sheetnames
                      if b.lower().startswith("zahlungsliste")), None)
        if blatt is None:
            return []
        ws = wb[blatt]
        kopf_i, idx, _ = _alt_kopf(ws, {"Vorname", "Name", "Institution"})
        if kopf_i < 0:
            return []
        alt: dict[str, tuple] = {}
        for r in range(kopf_i + 2, ws.max_row + 1):
            zeile = tuple(c.value for c in ws[r])
            vorname = _text(_zelle(zeile, finde_spalte(idx, "Vorname")))
            name = _text(_zelle(zeile, finde_spalte(idx, "Name")))
            inst = _mehrzeilig(_zelle(zeile, finde_spalte(idx, "Institution")))
            if not (vorname or name or inst):
                continue
            betrag = _komma(_zelle(
                zeile, finde_spalte(idx, "Auszahlungs-Betrag")))
            alt[_vergleichsschluessel(vorname, name, inst)] = (
                " ".join(x for x in (vorname, name, inst) if x), betrag)

        meins: dict[str, tuple] = {}
        grund_je: dict[str, str] = {}
        for ab in rechne_alle(bestand, jahr, cfg):
            if not ab.brief:
                continue
            e = ab.empfaenger
            schluessel = _vergleichsschluessel(e.vorname, e.name, e.institution)
            meins[schluessel] = (e.anzeigename, ab.brutto)
            grund_je[schluessel] = _empfaengergrund(e, jahr)

        for schluessel, (wer, soll) in alt.items():
            if soll is None:
                continue               # Zeile ohne Betrag: nichts zu zahlen
            if schluessel not in meins:
                abweichungen.append(Abweichung(
                    "Zahlung", wer, "", soll, 0.0,
                    "steht in der alten Zahlungsliste, bekommt hier nichts"))
            elif abs(meins[schluessel][1] - soll) > 0.005:
                abweichungen.append(Abweichung(
                    "Zahlung", wer, "", soll, meins[schluessel][1],
                    grund_je.get(schluessel, "")))
        for schluessel, (wer, ist) in meins.items():
            if schluessel not in alt:
                abweichungen.append(Abweichung(
                    "Zahlung", wer, "", 0.0, ist,
                    "neu — stand nicht in der alten Zahlungsliste"))
        return abweichungen
    finally:
        wb.close()


def pruefe_ksk(bestand: Bestand, pfad_alt, jahr: int,
               cfg: dict | None = None) -> list[Abweichung]:
    """Die Meldebeträge gegen das alte Künstlersozialkasse-Blatt stellen.

    Verglichen werden nur die Buchungen mit Belegdatum 31.12. — alles
    andere auf jenem Blatt stammt aus anderen Vorgängen (Lektorate,
    Einzelhonorare) und hat mit der Jahresabrechnung nichts zu tun.
    """
    cfg = cfg or {}
    pfad_alt = Path(pfad_alt)
    wb = openpyxl.load_workbook(pfad_alt, data_only=True)
    abweichungen: list[Abweichung] = []
    try:
        blatt = next((b for b in wb.sheetnames
                      if "sozialkasse" in b.lower()), None)
        if blatt is None:
            return []
        ws = wb[blatt]
        alt: dict[str, tuple] = {}
        for r in range(1, ws.max_row + 1):
            zeile = [c.value for c in ws[r]]
            if _text(_zelle(tuple(zeile), 0)) != f"31.12.{jahr}":
                continue
            text = _text(_zelle(tuple(zeile), 2))
            wer = text.split("_")[0].strip().rstrip(",")
            nachname, _, vorname = wer.partition(",")
            betrag = _komma(_zelle(tuple(zeile), 4), 0.0) or 0.0
            alt[_vergleichsschluessel(vorname, nachname, wer)] = (wer, betrag)

        meins: dict[str, tuple] = {}
        auszahlung: dict[str, float] = {}
        grund_je: dict[str, str] = {}
        for ab in rechne_alle(bestand, jahr, cfg):
            e = ab.empfaenger
            schluessel = _vergleichsschluessel(e.vorname, e.name, e.institution)
            auszahlung[schluessel] = ab.brutto
            grund_je[schluessel] = _empfaengergrund(e, jahr)
            if ab.ksk_netto <= 0:
                continue
            meins[schluessel] = (e.anzeigename, ab.ksk_netto)

        for schluessel, (wer, soll) in alt.items():
            if schluessel not in meins:
                abweichungen.append(Abweichung(
                    "KSK", wer, "", soll, 0.0,
                    "stand im alten KSK-Blatt, wird hier nicht gemeldet"))
            elif abs(meins[schluessel][1] - soll) > 0.005:
                unterschied = abs(meins[schluessel][1] - soll)
                # Die Beträge im alten KSK-Blatt wurden von Hand übertragen;
                # dabei ist manches auf volle Cent oder ganze Euro gerundet
                # worden (Buck steht dort mit 2214 statt 2214,01). Ein Cent
                # ist deshalb kein Befund, sondern eine Abschreibspur.
                if unterschied <= 0.015:
                    grund = "Centdifferenz — im alten Blatt von Hand übertragen"
                elif abs(auszahlung.get(schluessel, 0.0) - soll) <= (
                        # Mit gerundetem Satz weicht auch die Auszahlung
                        # um ein paar Cent von der Altmappe ab.
                        0.5 if _GRUND_GERUNDET in grund_je.get(schluessel, "")
                        else 0.015):
                    # Im alten Blatt steht der AUSZAHLUNGSBETRAG statt der
                    # Honorarsumme. Damit wurden Rückflüsse und Erlösanteile
                    # mitgemeldet, die der Künstlersozialkasse nicht
                    # zustehen. Das ist ein Fehler der Altmappe, kein
                    # Rechenfehler hier — die Altmappe selbst kommt in ihrer
                    # eigenen KSK-Spalte auf denselben Wert wie dieses
                    # Werkzeug.
                    grund = ("im alten Blatt wurde der Auszahlungsbetrag "
                             "gemeldet statt der Honorarsumme — dort sind "
                             "Rückflüsse mitgezählt, die nicht zur KSK "
                             "gehören")
                else:
                    grund = grund_je.get(schluessel, "")
                abweichungen.append(Abweichung(
                    "KSK", wer, "", soll, meins[schluessel][1], grund))
        for schluessel, (wer, ist) in meins.items():
            if schluessel not in alt:
                abweichungen.append(Abweichung(
                    "KSK", wer, "", 0.0, ist,
                    "neu — stand nicht im alten KSK-Blatt"))
        return abweichungen
    finally:
        wb.close()


def pruefe_gegen_excel(bestand: Bestand, pfad_alt, jahr: int = 2025,
                       cfg: dict | None = None) -> list[Abweichung]:
    """Jede Zeile, jede Auszahlung UND jede Ausgabeliste nachrechnen."""
    pfad_alt = Path(pfad_alt)
    wbv = openpyxl.load_workbook(pfad_alt, data_only=True)
    wbf = openpyxl.load_workbook(pfad_alt, data_only=False)
    abweichungen: list[Abweichung] = []
    try:
        blatt = str(jahr) if str(jahr) in wbv.sheetnames else ALT_BLATT_HAUPT
        ws_v, ws_f = wbv[blatt], wbf[blatt]
        kopf_i, idx, zeilen = _alt_kopf(ws_v, ALT_PFLICHT)
        i_titel = finde_spalte(idx, "BUCHTITEL")
        i_netto = finde_spalte(idx, "Betrag o. MwSt.")
        i_aus = finde_spalte(idx, "Auszahlungs-Betrag")
        spalte_aus = ws_v.cell(row=1, column=i_aus + 1).column_letter

        # Zeilenweise vergleichen, NICHT nach Titel gruppiert: derselbe
        # Titel steht mehrfach im Blatt, wenn sich zwei Autoren ein Buch
        # teilen. Deshalb trägt jedes Buch seine Herkunftszeile mit sich.
        nach_quelle = {}
        for e in bestand.empfaenger:
            for buch in e.buecher:
                if buch.quelle:
                    nach_quelle[buch.quelle] = (e, buch)

        for r in range(kopf_i + 2, ws_v.max_row + 1):
            zeile = tuple(c.value for c in ws_v[r])
            titel = _mehrzeilig(_zelle(zeile, i_titel))
            soll = _komma(_zelle(zeile, i_netto))
            if not titel or soll is None:
                continue
            treffer = nach_quelle.get(f"{blatt}!Z{r}")
            if treffer is None:
                continue
            e, buch = treffer
            posten = betrag_zeile(buch, jahr, cfg)
            if posten is None:
                # Kein Posten UND Sollbetrag null heißt schlicht: dieses Buch
                # hatte in dem Jahr nichts — das ist keine Abweichung.
                if abs(round(soll, 2)) > 0.005:
                    abweichungen.append(Abweichung(
                        "Zeile", e.anzeigename, titel, round(soll, 2), 0.0,
                        "Jahr nicht erfasst"))
                continue
            if abs(posten.netto - round(soll, 2)) > 0.005:
                abweichungen.append(Abweichung(
                    "Zeile", e.anzeigename, titel, round(soll, 2),
                    posten.netto, _abweichungsgrund(buch, jahr)))

        # Soll je Auszahlung: die Summenketten der Altmappe.
        for nummern in finde_bloecke(ws_f, spalte_aus):
            letzte = max(nummern)
            soll = _komma(ws_v.cell(row=letzte, column=i_aus + 1).value)
            if soll is None:
                continue
            im_block = [nach_quelle[f"{blatt}!Z{nr}"] for nr in nummern
                        if f"{blatt}!Z{nr}" in nach_quelle]
            if not im_block:
                continue
            ist = 0.0
            gruende = set()
            for e, buch in im_block:
                posten = betrag_zeile(buch, jahr, cfg)
                if posten is None:
                    continue
                ist = runde(ist + posten.brutto)
                grund = _abweichungsgrund(buch, jahr)
                if grund:
                    gruende.add(grund)
            if abs(ist - round(soll, 2)) > 0.005:
                abweichungen.append(Abweichung(
                    "Auszahlung", im_block[0][0].anzeigename,
                    ", ".join(sorted({b.titel for _, b in im_block}))[:60],
                    round(soll, 2), ist, "; ".join(sorted(gruende))))

        # Die Rechnung kann stimmen und die Ausgabe trotzdem falsch sein —
        # falscher Empfänger, doppelte Zahlung, fehlende Meldung. Deshalb
        # werden beide Listen mitgeprüft.
        abweichungen += pruefe_zahlungsliste(bestand, pfad_alt, jahr, cfg)
        abweichungen += pruefe_ksk(bestand, pfad_alt, jahr, cfg)
        return abweichungen
    finally:
        wbv.close()
        wbf.close()


# ---------------------------------------------------------------------
# Ausgaben
# ---------------------------------------------------------------------

def ausgabeordner(cfg: dict, jahr: int) -> Path:
    """honorar_output/<Jahr>/ neben der .exe."""
    ordner = APP_DIR / cfg.get("output_dir", "honorar_output") / str(jahr)
    ordner.mkdir(parents=True, exist_ok=True)
    return ordner


def runde(betrag: float, stellen: int = 2) -> float:
    """Kaufmännisch runden — 5,635 wird 5,64, nicht 5,63.

    Pythons eingebautes ``round`` rundet die Hälfte zur geraden Ziffer und
    rechnet dabei auf der Fließkommadarstellung: aus 5,635 wird 5,63, aus
    2,675 wird 2,67. Excel rundet die Hälfte immer auf. Bei einer Abrechnung
    über 20.000 € und 500 Posten ist das kein Schönheitsfehler, sondern ein
    Cent, der irgendwo fehlt — und der Verlag rechnet gegen eine Excel nach.
    """
    if betrag is None:
        return betrag
    muster = Decimal(1).scaleb(-stellen)
    return float(Decimal(repr(float(betrag))).quantize(muster, ROUND_HALF_UP))


def euro(betrag: float) -> str:
    """1234.5 → „1.234,50 €" — deutsche Schreibweise wie in den Altbriefen."""
    s = f"{betrag:,.2f}".replace(",", "#").replace(".", ",").replace("#", ".")
    return f"{s} €"


def _sicherer_dateiname(text: str) -> str:
    """Was Windows im Dateinamen nicht duldet, muss weg — Umlaute dürfen
    bleiben, die stehen auch in den Altbriefen („Rückfluss")."""
    for z in '<>:"/\\|?*':
        text = text.replace(z, "-")
    return " ".join(text.split()).strip(". ")


def gesonderte_posten(bestand: Bestand, jahr: int,
                      cfg: dict | None = None) -> list[dict]:
    """Die Bücher, die auf eine Entscheidung des Verlags warten.

    Je Buch eine Zeile — mit Betrag und Herkunft aus der Altmappe, damit
    jemand die Sache zu Ende bringen kann, ohne sie erst wieder suchen zu
    müssen.
    """
    cfg = cfg or {}
    zeilen = []
    for e, b in bestand.buecher():
        if not b.gesondert or b.stillgelegt:
            continue
        posten = betrag_zeile(b, jahr, cfg)
        if posten is None or posten.netto <= 0:
            continue
        zeilen.append({
            "_kennung": b.kennung,
            "Empfänger": e.anzeigename,
            "Buchtitel": b.titel,
            "ISBN": b.isbn,
            "Vergütungsart": b.verguetungsart,
            "Vergütungs-Ex.": posten.verguetungs_ex,
            "Betrag": posten.brutto,
            "Bankverbindung": e.iban,
            "in der Altmappe": b.quelle,
        })
    zeilen.sort(key=lambda z: -z["Betrag"])
    return zeilen


def _blatt_gesondert(wb, bestand: Bestand, jahr: int, cfg: dict):
    """Ein eigenes Blatt für die offenen Sonderfälle.

    Ohne das stehen sie nur grau in einer Liste von 250 Zeilen — und
    neunzehn der zwanzig Empfänger bekommen gar keinen Brief, tauchen also
    nirgends auf, wo man sie bearbeiten würde.
    """
    zeilen = gesonderte_posten(bestand, jahr, cfg)
    if not zeilen:
        return None
    spalten = ["Empfänger", "Buchtitel", "ISBN", "Vergütungsart",
               "Vergütungs-Ex.", "Betrag", "Bankverbindung", "in der Altmappe"]
    # „_kennung“ ist nur für die Oberfläche da und gehört nicht ins Blatt.
    ws = wb.create_sheet(f"Offene Sonderfälle {jahr}")
    ws.append(["Diese Beträge werden NICHT ausgezahlt, solange die Verträge "
               "auf „gesondert abrechnen“ stehen."])
    ws.cell(row=1, column=1).font = Font(bold=True, size=12)
    ws.append(["Sie stammen alle aus dem Blatt „Zahlung ab XX Ex.“ der "
               "Altmappe, das dort nur einen Stand führte und keine "
               "Auszahlungen auslöste."])
    ws.append([])
    ws.append(spalten)
    for z in ws[ws.max_row]:
        z.font = Font(bold=True)
    erste = ws.max_row + 1
    for zeile in zeilen:
        ws.append([zeile.get(s) for s in spalten])
    ws.append([None, None, None, "Summe", None,
               f"=SUM(F{erste}:F{ws.max_row})"])
    ws.cell(row=ws.max_row, column=4).font = Font(bold=True)
    ws.freeze_panes = f"A{erste}"
    for i, s in enumerate(spalten, start=1):
        b = ws.cell(row=4, column=i).column_letter
        ws.column_dimensions[b].width = {
            "Empfänger": 34, "Buchtitel": 40, "ISBN": 10,
            "Vergütungsart": 16, "Vergütungs-Ex.": 14, "Betrag": 12,
            "Bankverbindung": 28, "in der Altmappe": 24}.get(s, 14)
        if s == "Betrag":
            for zelle in ws[b][3:]:
                zelle.number_format = GELDFORMAT
        if s in ("ISBN", "Bankverbindung", "in der Altmappe"):
            for zelle in ws[b][3:]:
                zelle.number_format = "@"
    return ws


def schreibe_listen(abrechnungen: list[Abrechnung], jahr: int, ziel: Path,
                    cfg: dict, bestand: Bestand | None = None) -> Path:
    """Zahlungsliste, KSK-Meldung und Protokoll in EINER Mappe.

    Drei Dateien für drei Listen waren drei Gelegenheiten, die falsche zu
    öffnen oder eine zu übersehen. Sie gehören zusammen: dieselbe
    Abrechnung, dasselbe Jahr, derselbe Arbeitsgang.

    Die Datei heißt weiterhin `Zahlungsliste_<Jahr>.xlsx` — das ist der Name,
    unter dem im Verlag danach gesucht wird, und die Zahlungsliste steht als
    erstes Blatt darin. „Abrechnung_<Jahr>“ wäre zu leicht mit dem Blatt
    „Abrechnung <Jahr>“ im Bestand zu verwechseln.
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    _blatt_zahlungsliste(wb, abrechnungen, jahr)
    _blatt_ksk(wb, abrechnungen, jahr, cfg)
    _blatt_protokoll(wb, abrechnungen, jahr)
    if bestand is not None:
        _blatt_gesondert(wb, bestand, jahr, cfg)
    ziel = Path(ziel)
    ziel.parent.mkdir(parents=True, exist_ok=True)
    wb.save(ziel)
    wb.close()
    return ziel


def _blatt_zahlungsliste(wb, abrechnungen: list[Abrechnung], jahr: int):
    """Die Liste für die Überweisungen — Aufbau wie das bisherige Blatt."""
    spalten = ["Vorname", "Name", "Institution", "VERGÜT-ART",
               "Auszahlungs-Betrag", "Bankverbindung",
               "Aktenzeichen / Kontoinhaber", "Überweisung"]
    ws = wb.create_sheet(f"Zahlungsliste {jahr}")
    ws.append(spalten)
    for z in ws[1]:
        z.font = Font(bold=True)
    for ab in abrechnungen:
        if not ab.brief:
            continue
        e = ab.empfaenger
        ws.append([e.vorname, e.name, e.institution, ab.verguetungsart,
                   ab.brutto, e.iban, e.aktenzeichen, None])
    ws.freeze_panes = "A2"
    for i, s in enumerate(spalten, start=1):
        b = ws.cell(row=1, column=i).column_letter
        ws.column_dimensions[b].width = {
            "Vorname": 16, "Name": 22, "Institution": 34, "VERGÜT-ART": 18,
            "Auszahlungs-Betrag": 17, "Bankverbindung": 28,
            "Aktenzeichen / Kontoinhaber": 30, "Überweisung": 14}.get(s, 14)
        if s == "Auszahlungs-Betrag":
            for zelle in ws[b][1:]:
                zelle.number_format = GELDFORMAT
        if s in ("Bankverbindung", "Überweisung"):
            for zelle in ws[b][1:]:
                zelle.number_format = "@"
    return ws


def ksk_satz(ab: Abrechnung) -> float:
    """Mit welchem Umsatzsteuersatz dieser Empfänger zu melden ist.

    0 heißt: nicht mehrwertsteuerpflichtig. Das ist keine Randgruppe — im
    Altbestand sind es 54 von 74 Buchungen, und sie gehören auf ein eigenes
    Konto („Honorare“), nicht zu den Fremdarbeiten mit Steuer.
    """
    saetze = {p.buch.mwst_satz for p in ab.posten
              if p.buch.verguetungsart == KSK_ART and p.buch.mwst_pflichtig}
    return max(saetze) if saetze else 0.0


def _blatt_ksk(wb, abrechnungen: list[Abrechnung], jahr: int, cfg: dict):
    """Die Meldung an die Künstlersozialkasse, im Buchungsformat der Mappe.

    Nur echte Honorare zählen — Rückflüsse, Erlösanteile und
    Darlehensrückzahlungen sind keine Honorare im Sinne der KSK.

    Drei Abschnitte, nicht zwei: getrennt wird danach, ob der Autor
    mehrwertsteuerpflichtig ist und mit welchem Satz. Wer es nicht ist,
    kommt auf ein eigenes Konto ohne Umsatzsteuer.

    Negative Beträge bleiben draußen. Ein Autor mit mehr Rückgaben als
    Verkäufen hat kein negatives Honorar zu melden; die Altmappe kennt
    solche Zeilen auch nicht. Sie stehen stattdessen im Protokoll, damit
    niemand sie übersieht.

    Bewusst nur eine Excel: der Weg über die ASCII-Schnittstelle nach Lexware
    (s. PLAN_Buchhaltung.md) ist eine eigene Entscheidung mit eigener
    Prüfung — das Handbuch warnt ausdrücklich vor Mehrfachimport.
    """
    spalten = ["Belegdatum", "Belegnummer", "Buchungstext", "Gegenkonto",
               "Sollbetrag EUR", "Habenbetrag EUR", "USt-Konto", "USt-%"]
    ws = wb.create_sheet(f"Künstlersozialkasse {jahr}")

    nach_satz: dict[float, list] = {}
    for ab in abrechnungen:
        if ab.ksk_netto <= 0:
            continue
        nach_satz.setdefault(ksk_satz(ab), []).append(ab)

    konten = {19.0: ("ksk_konto_19", "ksk_bezeichnung_19", "ksk_ust_konto_19"),
              7.0: ("ksk_konto_7", "ksk_bezeichnung_7", "ksk_ust_konto_7"),
              0.0: ("ksk_konto_0", "ksk_bezeichnung_0", None)}

    belegdatum = f"31.12.{jahr}"
    for satz in sorted(nach_satz, reverse=True):
        schluessel = konten.get(satz, konten[7.0])
        konto = cfg.get(schluessel[0], "")
        bezeichnung = cfg.get(schluessel[1], f"Fremdarbeiten ({satz:g}%)")
        ust = cfg.get(schluessel[2], "") if schluessel[2] else ""
        ws.append(["Konto", konto, bezeichnung])
        ws.cell(row=ws.max_row, column=1).font = Font(bold=True)
        ws.append(spalten)
        for z in ws[ws.max_row]:
            z.font = Font(bold=True)
        erste = ws.max_row + 1
        for ab in sorted(nach_satz[satz],
                         key=lambda a: a.empfaenger.dateiname_basis.lower()):
            e = ab.empfaenger
            wer = f"{e.name}, {e.vorname}".strip(", ") or e.institution
            ws.append([belegdatum, "", f"{wer}_Honorar {jahr}",
                       cfg.get("ksk_gegenkonto"), ab.ksk_netto, "",
                       ust, f"{satz:g},00".replace(".", ",")])
        letzte = ws.max_row
        summenzeile = [None, None, None, "Summe",
                       f"=SUM(E{erste}:E{letzte})"]
        if satz:
            summenzeile += ["netto=", f"=E{letzte + 1}*100/{100 + satz:g}"]
        ws.append(summenzeile)
        ws.cell(row=ws.max_row, column=4).font = Font(bold=True)
        ws.append([])

    # Gesamtsumme, wie sie auch in der Altmappe unten steht.
    gesamt = runde(sum(a.ksk_netto for a in abrechnungen if a.ksk_netto > 0))
    ws.append([None, None, "KSK-Gesamtsumme", None, gesamt])
    ws.cell(row=ws.max_row, column=3).font = Font(bold=True)

    # Wer wegen eines negativen Honorars nicht gemeldet wird.
    negativ = [a for a in abrechnungen if a.ksk_netto < 0]
    if negativ:
        ws.append([])
        ws.append([None, None,
                   "Nicht gemeldet, weil das Honorar negativ ist "
                   "(mehr Rückgaben als Verkäufe):"])
        ws.cell(row=ws.max_row, column=3).font = Font(bold=True)
        for a in sorted(negativ, key=lambda x: x.ksk_netto):
            ws.append([None, None, a.empfaenger.anzeigename, None, a.ksk_netto])

    ws.column_dimensions["A"].width = 13
    ws.column_dimensions["B"].width = 13
    ws.column_dimensions["C"].width = 46
    for sp in "DEFGH":
        ws.column_dimensions[sp].width = 15
    # Soll- und Habenbetrag in deutscher Schreibweise; Belegdatum,
    # Belegnummer und die Kontonummern bleiben Text, damit Excel aus
    # „001610“ keine 1610 macht.
    for zelle in ws["E"][1:] + ws["F"][1:]:
        zelle.number_format = GELDFORMAT
    for sp in ("A", "B", "D", "G", "H"):
        for zelle in ws[sp][1:]:
            zelle.number_format = "@"
    return ws


def _blatt_protokoll(wb, abrechnungen: list[Abrechnung], jahr: int):
    """Wer bekommt einen Brief, wer nicht, und was ist offen.

    Gerade die Empfänger OHNE Brief gehören hier hinein: sonst bleibt
    unklar, ob sie übersehen wurden oder bewusst leer ausgingen.
    """
    spalten = ["Empfänger", "Vergütungsart", "Bücher", "Netto", "MwSt",
               "Auszahlung", "KSK-Netto", "Brief", "Grund / Hinweis"]
    ws = wb.create_sheet(f"Protokoll {jahr}")
    ws.append(spalten)
    for z in ws[1]:
        z.font = Font(bold=True)
    for ab in sorted(abrechnungen, key=lambda a: -a.brutto):
        hinweise = list(ab.gruende) + list(ab.probleme)
        ws.append([ab.empfaenger.anzeigename, ab.verguetungsart,
                   len(ab.posten), ab.netto, ab.mwst, ab.brutto,
                   ab.ksk_netto, "Ja" if ab.brief else "Nein",
                   " | ".join(hinweise)])
        ws.cell(row=ws.max_row, column=9).alignment = Alignment(wrapText=True)
    ws.freeze_panes = "A2"
    for i, s in enumerate(spalten, start=1):
        b = ws.cell(row=1, column=i).column_letter
        ws.column_dimensions[b].width = {
            "Empfänger": 38, "Vergütungsart": 16, "Bücher": 8, "Netto": 11,
            "MwSt": 10, "Auszahlung": 12, "KSK-Netto": 11, "Brief": 7,
            "Grund / Hinweis": 80}.get(s, 14)
        if s in ("Netto", "MwSt", "Auszahlung", "KSK-Netto"):
            for zelle in ws[b][1:]:
                zelle.number_format = GELDFORMAT
    return ws


# ---------------------------------------------------------------------
# Abrechnungsbriefe
# ---------------------------------------------------------------------
# Zwei Seiten, wie die Altbriefe: Seite 1 hoch mit Briefkopf, Anschrift und
# Betrag, Seite 2 quer mit der Titeltabelle. Die Vorlage trägt Platzhalter in
# der Form {{NAME}}, jeder in GENAU EINEM Run — sonst zerfällt der Platzhalter
# beim Ersetzen, weil Word Text gern mitten im Wort auf Runs aufteilt.

TOKEN_TABELLE = "{{POSTEN}}"

# Steht in den Dokumenteigenschaften der selbstgebauten Briefvorlage. Die
# echte Vorlage des Verlags trägt sie nicht — daran erkennt das Werkzeug,
# ob es noch mit dem Nachbau arbeitet.
NACHBAU_KENNUNG = "Nachbau aus den Muster-PDFs — Honorar-Abrechner"


def vorlage_ist_nachbau(vorlage: Path) -> bool:
    """Ob die Briefvorlage noch der Nachbau ist und nicht die echte."""
    try:
        return Document(str(vorlage)).core_properties.comments == NACHBAU_KENNUNG
    except Exception:
        return False

BRIEFKOPF_SPALTEN = [
    ["Verlag Regionalkultur GmbH & Co. KG", "Bahnhofstr. 2 • 76698 Ubstadt-Weiher",
     "Sitz der Gesellschaft Ubstadt-Weiher",
     "Handelsregister Amtsgericht Mannheim A 709462", "",
     "Geschäftsführer:", "Reiner Schmidt, Andrea Sitzler"],
    ["Verlag Regionalkultur Verwaltungs GmbH (Komplementärin)",
     "Bahnhofstr. 2 • 76698 Ubstadt-Weiher",
     "Sitz der Gesellschaft Ubstadt-Weiher",
     "Amtsgericht Mannheim HRB 736977", "", "", ""],
    ["Bankverbindungen:", "Postbank Karlsruhe",
     "IBAN: DE91 6601 0075 0199 4437 54", "BIC: PBNKDEFF", "",
     "Sparkasse Kraichgau", "IBAN: DE87 6635 0036 0000 7090 74"],
    ["USt.-Id-Nr.: DE334279458", "Gerichtsstand Bruchsal", "", "", "", "",
     "BIC: BRUSDE66XXX"],
]

RUECKSENDEZEILE = ("Verlag Regionalkultur GmbH & Co. KG • Bahnhofstraße 2 • "
                   "76698 Ubstadt-Weiher")

BRIEFTEXT = [
    "hiermit übersenden wir Ihnen die Aufstellung über den Verkauf Ihres "
    "Buches / Ihrer Bücher.",
    "Wir werden Ihnen den Gesamtbetrag in Höhe von {{BETRAG}} auf Ihr "
    "Konto Nr.: {{IBAN}} überweisen.",
    "Wir danken für Ihre Mitarbeit und wünschen Ihnen Gesundheit und ein "
    "schönes und erfolgreiches Jahr.",
]

# Die Titeltabelle auf Seite 2, Spalte für Spalte: Überschrift, Breite und
# Ausrichtung. Ohne feste Breiten verteilt Word gleichmäßig — dann bricht der
# Buchtitel dreizeilig um, während „Eigenkauf“ mit vier Zeichen eine
# handbreite Spalte belegt. Die Maße sind am Muster-PDF abgenommen; in der
# Summe 24,9 cm, die Seite bietet quer 26,7 cm zwischen den Rändern.
SPALTEN_TABELLE = [
    ("Titel",              Cm(7.6), "mitte"),
    ("€/je Ex.",           Cm(2.0), "mitte"),
    ("VK {{JAHR}}",        Cm(2.2), "mitte"),
    ("Eigenkauf",          Cm(2.3), "mitte"),
    ("Vergütungs-Ex.",     Cm(2.9), "mitte"),
    ("Vergütungs-Art",     Cm(3.0), "mitte"),
    ("Betrag Netto",       Cm(2.7), "rechts"),
    ("MwSt. {{MWSTSATZ}}", Cm(2.2), "rechts"),
]

AUSRICHTUNG = {
    "links": WD_ALIGN_PARAGRAPH.LEFT,
    "mitte": WD_ALIGN_PARAGRAPH.CENTER,
    "rechts": WD_ALIGN_PARAGRAPH.RIGHT,
}


def _setze_spaltenraster(tab, breiten) -> None:
    """Die Spaltenbreiten ins ``tblGrid`` schreiben.

    python-docx setzt beim Zuweisen von ``zelle.width`` nur das ``tcW`` der
    einzelnen Zelle. Bei festem Tabellenlayout richtet Word sich aber nach
    dem ``tblGrid`` — und das steht weiterhin auf acht gleich breiten
    Spalten. Ergebnis: alle Spalten gleich breit, der Buchtitel bricht
    dreizeilig um. Also beides setzen.
    """
    raster = tab._tbl.find(qn("w:tblGrid"))
    if raster is None:
        return
    vorhanden = raster.findall(qn("w:gridCol"))
    # Das Raster muss GENAU so viele Spalten haben wie die Tabelle. Bleibt
    # eine überzählige stehen (etwa nachdem die MwSt-Spalte entfernt wurde),
    # rechnet Word weiter mit acht Spalten und die Breiten verrutschen.
    for ueberzaehlig in vorhanden[len(breiten):]:
        raster.remove(ueberzaehlig)
    while len(raster.findall(qn("w:gridCol"))) < len(breiten):
        raster.append(OxmlElement("w:gridCol"))
    for spalte, breite in zip(raster.findall(qn("w:gridCol")), breiten):
        spalte.set(qn("w:w"), str(Emu(int(breite)).twips))
    # Gesamtbreite ebenfalls festschreiben, sonst rechnet Word sie sich neu.
    tblpr = tab._tbl.tblPr
    for alt in tblpr.findall(qn("w:tblW")):
        tblpr.remove(alt)
    gesamt = OxmlElement("w:tblW")
    gesamt.set(qn("w:type"), "dxa")
    gesamt.set(qn("w:w"), str(sum(Emu(int(b)).twips for b in breiten)))
    tblpr.append(gesamt)


def _schmale_zellraender(tab, mm: int = 30) -> None:
    """Innenabstand der Zellen verkleinern (Angabe in Twips)."""
    tblpr = tab._tbl.tblPr
    for alt in tblpr.findall(qn("w:tblCellMar")):
        tblpr.remove(alt)
    raender = OxmlElement("w:tblCellMar")
    for seite in ("top", "left", "bottom", "right"):
        rand = OxmlElement(f"w:{seite}")
        rand.set(qn("w:w"), str(mm))
        rand.set(qn("w:type"), "dxa")
        raender.append(rand)
    tblpr.append(raender)


def _kopfzeile_wiederholen(tab) -> None:
    """Die Überschriftenzeile auf jeder Folgeseite wiederholen.

    Ein Autor mit vierzig Titeln bekommt sonst eine zweite Seite ohne
    Spaltenüberschriften — und niemand weiß mehr, welche Zahl was ist.
    """
    trpr = tab.rows[0]._tr.get_or_add_trPr()
    kopf = OxmlElement("w:tblHeader")
    kopf.set(qn("w:val"), "true")
    trpr.append(kopf)


def _absatz(ziel, text="", groesse=11, fett=False, schrift="Times New Roman",
            abstand_nach=0):
    p = ziel.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(groesse)
    run.font.name = schrift
    run.bold = fett
    p.paragraph_format.space_after = Pt(abstand_nach)
    return p


def baue_briefvorlage(ziel: Path) -> Path:
    """Die Briefvorlage erzeugen — nachgebaut aus den Muster-PDFs.

    Der Verlag hat die Originalvorlage bisher nicht geliefert; dies hier ist
    der Nachbau, damit überhaupt etwas Prüfbares vorliegt. Sobald die echte
    Vorlage da ist, wird sie einfach an dieselbe Stelle gelegt — die
    Platzhalter müssen dann nur gleich heißen.
    """
    doc = Document()

    # --- Seite 1: hoch -------------------------------------------------
    abschnitt = doc.sections[0]
    abschnitt.orientation = WD_ORIENT.PORTRAIT
    abschnitt.page_width, abschnitt.page_height = Cm(21), Cm(29.7)
    abschnitt.left_margin = abschnitt.right_margin = Cm(2.5)
    abschnitt.top_margin = Cm(2.5)
    abschnitt.bottom_margin = Cm(2.0)

    # Der Briefkopf steht in der Fußzeile, wie im Muster-PDF: vier schmale
    # Spalten in 7 pt.
    fuss = abschnitt.footer
    fuss.paragraphs[0].text = ""
    tabelle = fuss.add_table(rows=7, cols=4, width=Cm(16))
    tabelle.autofit = False
    # Die vier Spalten tragen unterschiedlich lange Angaben. Gleich breit
    # bricht jede zweite Zeile mitten im Wort um („Ubstadt-/Weiher“).
    breiten = (Cm(4.5), Cm(4.4), Cm(4.4), Cm(3.4))
    _setze_spaltenraster(tabelle, breiten)
    # Word gibt jeder Zelle links und rechts knapp 2 mm Rand mit. Bei 6-pt-
    # Text in einer 4,5-cm-Spalte ist das der Unterschied zwischen einer und
    # zwei Zeilen.
    _schmale_zellraender(tabelle)
    for spalte, zeilen in enumerate(BRIEFKOPF_SPALTEN):
        for zeile, text in enumerate(zeilen):
            zelle = tabelle.cell(zeile, spalte)
            zelle.width = breiten[spalte]
            absatz = zelle.paragraphs[0]
            absatz.text = ""
            absatz.paragraph_format.space_after = Pt(0)
            run = absatz.add_run(text)
            run.font.size = Pt(6)
            run.font.name = "Arial"

    _absatz(doc, RUECKSENDEZEILE, groesse=7, schrift="Arial", abstand_nach=14)
    _absatz(doc, "{{ANSCHRIFT}}", groesse=11, abstand_nach=0)
    _absatz(doc, "", abstand_nach=24)
    _absatz(doc, "{{ORT_DATUM}}", groesse=11, abstand_nach=24)
    _absatz(doc, "{{BETREFF}}", groesse=11, fett=True, abstand_nach=24)
    _absatz(doc, "{{ANREDE}}", groesse=11, abstand_nach=12)
    for text in BRIEFTEXT:
        _absatz(doc, text, groesse=11, abstand_nach=12)
    _absatz(doc, "Mit freundlichen Grüßen", groesse=11, abstand_nach=0)
    _absatz(doc, "Ihr verlag regionalkultur", groesse=11, abstand_nach=48)
    _absatz(doc, "{{UNTERZEICHNER}}", groesse=11)

    # --- Seite 2: quer -------------------------------------------------
    quer = doc.add_section()
    quer.orientation = WD_ORIENT.LANDSCAPE
    quer.page_width, quer.page_height = Cm(29.7), Cm(21)
    quer.left_margin = quer.right_margin = Cm(1.5)
    # Oben Luft lassen, wie im Muster — die Tabelle klebt sonst am Rand.
    quer.top_margin = Cm(3.0)
    quer.bottom_margin = Cm(1.5)
    # Die Fußzeile des Briefkopfs gehört nicht auf die Tabellenseite.
    quer.footer.is_linked_to_previous = False
    quer.footer.paragraphs[0].text = ""

    # Die Tabelle wird mit der breitesten Fassung angelegt; die MwSt-Spalte
    # entfernt der Brieferzeuger, wenn sie nicht gebraucht wird.
    tab = doc.add_table(rows=2, cols=len(SPALTEN_TABELLE))
    tab.style = "Table Grid"
    tab.alignment = WD_TABLE_ALIGNMENT.CENTER
    # Ohne autofit=False ignoriert Word die Breiten und verteilt gleichmäßig.
    tab.autofit = False

    muster = ["{{TITEL}}", "{{SATZ}}", "{{VK}}", "{{EIGENKAUF}}",
              "{{VERGEX}}", "{{ART}}", "{{NETTO}}", "{{MWST}}"]

    for i, (titel, breite, wohin) in enumerate(SPALTEN_TABELLE):
        for zeile_nr, inhalt in ((0, titel), (1, muster[i])):
            zelle = tab.cell(zeile_nr, i)
            # Die Breite muss an JEDER Zelle stehen, nicht nur an der Spalte —
            # Word liest sie aus den Zellen, nicht aus dem Spaltenraster.
            zelle.width = breite
            zelle.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            absatz = zelle.paragraphs[0]
            absatz.text = ""
            absatz.alignment = AUSRICHTUNG[wohin if zeile_nr else "mitte"]
            absatz.paragraph_format.space_before = Pt(2)
            absatz.paragraph_format.space_after = Pt(2)
            run = absatz.add_run(inhalt)
            run.font.size = Pt(9)
            run.font.name = "Arial"
            run.bold = (zeile_nr == 0)

    _setze_spaltenraster(tab, [b for _, b, _ in SPALTEN_TABELLE])
    _kopfzeile_wiederholen(tab)

    # Das Logo steht im Muster unten links, nicht über der Tabelle.
    _absatz(doc, "", abstand_nach=18)
    _absatz(doc, "verlag regionalkultur", groesse=10, schrift="Verdana")

    # Die Vorlage als Nachbau kennzeichnen. Sobald der Verlag die echte
    # Word-Vorlage an dieselbe Stelle legt, fehlt diese Markierung und der
    # Hinweis verschwindet von selbst — niemand muss daran denken.
    doc.core_properties.comments = NACHBAU_KENNUNG

    ziel = Path(ziel)
    ziel.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(ziel))
    return ziel


def _textruns(absatz):
    return [r for r in absatz.runs if r.text]


def _alle_absaetze(doc):
    """Alle Absätze aus Body, Tabellen und Kopf-/Fußzeilen."""
    yield from doc.paragraphs
    for tbl in doc.tables:
        for zeile in tbl.rows:
            for zelle in zeile.cells:
                yield from zelle.paragraphs
    for abschnitt in doc.sections:
        for kf in (abschnitt.header, abschnitt.first_page_header,
                   abschnitt.footer, abschnitt.first_page_footer):
            if kf is None:
                continue
            yield from kf.paragraphs
            for tbl in kf.tables:
                for zeile in tbl.rows:
                    for zelle in zeile.cells:
                        yield from zelle.paragraphs


def _ersetze_in_absaetzen(absaetze, zuordnung: dict) -> None:
    for p in absaetze:
        runs = _textruns(p)
        if not runs:
            continue
        for run in runs:
            for platzhalter, wert in zuordnung.items():
                if platzhalter in run.text:
                    run.text = run.text.replace(platzhalter, str(wert))


def _ersetze_platzhalter(doc, zuordnung: dict) -> None:
    _ersetze_in_absaetzen(list(_alle_absaetze(doc)), zuordnung)


def _fuelle_mehrzeilig(doc, token: str, zeilen: list[str]) -> None:
    """Einen Platzhalter-Absatz durch je einen Absatz pro Zeile ersetzen,
    unter Beibehaltung von Schrift und Absatzformat der Vorlage."""
    ziel = None
    for p in doc.paragraphs:
        if p.text.strip() == token:
            ziel = p
            break
    if ziel is None:
        return
    for text in (zeilen or [""]):
        neu = deepcopy(ziel._p)
        absatz = Paragraph(neu, ziel._parent)
        runs = _textruns(absatz)
        if runs:
            runs[0].text = text
            for r in runs[1:]:
                r.text = ""
        ziel._p.addprevious(neu)
    ziel._p.getparent().remove(ziel._p)


def _fuelle_tabelle(doc, zeilen: list[list[str]]) -> None:
    """Die Musterzeile der Titeltabelle je Posten klonen.

    ``_fuelle_mehrzeilig`` kann nur Absätze vervielfältigen; eine
    Tabellenzeile braucht einen eigenen Weg, weil ihr <w:tr> geklont werden
    muss, damit Zellbreiten und Rahmen erhalten bleiben.
    """
    if not doc.tables:
        return
    tab = doc.tables[-1]
    if len(tab.rows) < 2:
        return
    muster = tab.rows[1]
    for werte in zeilen:
        neu = deepcopy(muster._tr)
        muster._tr.addprevious(neu)
        zeile = _Row(neu, tab)
        for i, zelle in enumerate(zeile.cells):
            text = werte[i] if i < len(werte) else ""
            for j, p in enumerate(zelle.paragraphs):
                runs = _textruns(p)
                if not runs:
                    continue
                runs[0].text = text if j == 0 else ""
                for r in runs[1:]:
                    r.text = ""
    muster._tr.getparent().remove(muster._tr)


def _entferne_spalte(doc, index: int) -> None:
    """Eine Spalte aus der Titeltabelle nehmen.

    Empfänger ohne Mehrwertsteuerpflicht bekommen die MwSt-Spalte gar nicht
    erst zu sehen — so ist es auch in den Altbriefen. Eine zweite Vorlage
    dafür zu pflegen wäre die schlechtere Lösung.

    Die frei werdende Breite bekommt die Titelspalte. Sonst stünde die
    Tabelle danach schmal und linkslastig auf der Seite.
    """
    if not doc.tables:
        return
    tab = doc.tables[-1]
    frei = SPALTEN_TABELLE[index][1] if index < len(SPALTEN_TABELLE) else Cm(0)
    for zeile in tab.rows:
        zellen = zeile.cells
        if index < len(zellen):
            tr = zeile._tr
            tc = zellen[index]._tc
            if tc.getparent() is tr:
                tr.remove(tc)
        if zeile.cells:
            # Länge + Länge ergibt in python-docx ein nacktes int — wieder in
            # eine Länge fassen, sonst fehlt später die twips-Umrechnung.
            zeile.cells[0].width = Emu(int(SPALTEN_TABELLE[0][1]) + int(frei))
    breiten = [b for _, b, _ in SPALTEN_TABELLE]
    breiten[0] = Emu(int(breiten[0]) + int(frei))
    del breiten[index]
    _setze_spaltenraster(tab, breiten)


def anschrift_zeilen(e: Empfaenger) -> list[str]:
    """Das Anschriftenfeld, wie es in den Altbriefen aussieht.

    Bei Personen steht „Herrn"/„Frau" voran, bei reinen Einrichtungen steht
    nur der Name der Einrichtung da — genau so wie im Muster-PDF für das
    Stadtarchiv Stuttgart.
    """
    zeilen = []
    person = " ".join(x for x in (e.titel_akad, e.vorname, e.name) if x).strip()
    if person:
        if e.anrede == "Herr":
            zeilen.append("Herrn")
        elif e.anrede == "Frau":
            zeilen.append("Frau")
        if e.institution:
            zeilen.append(e.institution.replace("\n", ", "))
        zeilen.append(person)
    elif e.institution:
        zeilen.append(e.institution.replace("\n", ", "))
    if e.strasse:
        zeilen.append(e.strasse)
    ort = " ".join(x for x in (e.plz, e.ort) if x).strip()
    if ort:
        zeilen.append(ort)
    if e.land:
        zeilen.append(e.land)
    return zeilen


def briefanrede(e: Empfaenger) -> str:
    """„Sehr geehrter Herr Buck," / „Sehr geehrte Damen und Herren,"."""
    nachname = " ".join(x for x in (e.titel_akad, e.name) if x).strip()
    if e.anrede == "Herr" and nachname:
        return f"Sehr geehrter Herr {nachname},"
    if e.anrede == "Frau" and nachname:
        return f"Sehr geehrte Frau {nachname},"
    if e.anrede == "Frau und Herr" and nachname:
        return f"Sehr geehrte Frau und Herr {nachname},"
    return "Sehr geehrte Damen und Herren,"


def generiere_brief(vorlage: Path, ab: Abrechnung, cfg: dict,
                    ziel: Path) -> Path:
    """Einen Abrechnungsbrief als .docx erzeugen."""
    doc = Document(str(vorlage))
    jahr = ab.jahr
    mit_mwst = ab.mwst > 0
    mwst_satz = max((p.buch.mwst_satz for p in ab.posten
                     if p.buch.mwst_pflichtig), default=7.0)

    zeilen = []
    for p in ab.posten:
        # Bücher, mit denen im Abrechnungsjahr gar nichts passiert ist,
        # stehen auch in den Altbriefen nicht drin. Neun Zeilen „0 Ex. ·
        # 0,00 €“ machen die Aufstellung nur unübersichtlich. Ein Buch mit
        # Verkäufen bleibt drin, auch wenn der Betrag durch die Schwelle auf
        # null fällt — das will der Autor sehen.
        if not p.verkauft and not p.eigenkauf and not p.netto:
            continue
        werte = [p.buch.titel, euro(p.satz), f"{p.verkauft} Ex.",
                 f"{p.eigenkauf} Ex.", f"{p.verguetungs_ex} Ex.",
                 p.buch.verguetungsart, euro(p.netto)]
        if mit_mwst:
            werte.append(euro(p.mwst))
        zeilen.append(werte)
    if not mit_mwst:
        _entferne_spalte(doc, len(SPALTEN_TABELLE) - 1)
    _fuelle_tabelle(doc, zeilen)

    _fuelle_mehrzeilig(doc, "{{ANSCHRIFT}}", anschrift_zeilen(ab.empfaenger))
    monat = "Januar"
    _ersetze_platzhalter(doc, {
        "{{ORT_DATUM}}": f"{cfg.get('absender_ort', 'Ubstadt-Weiher')}, "
                         f"im {monat} {jahr + 1}",
        "{{BETREFF}}": f"Abrechnung {jahr}",
        "{{ANREDE}}": briefanrede(ab.empfaenger),
        "{{BETRAG}}": euro(ab.brutto),
        "{{IBAN}}": ab.empfaenger.iban or "— bitte Bankverbindung nachtragen —",
        "{{UNTERZEICHNER}}": cfg.get("unterzeichner", ""),
        "{{JAHR}}": str(jahr),
        "{{MWSTSATZ}}": f"{mwst_satz:g} %",
    })

    # Die Vorlage als Nachbau kennzeichnen. Sobald der Verlag die echte
    # Word-Vorlage an dieselbe Stelle legt, fehlt diese Markierung und der
    # Hinweis verschwindet von selbst — niemand muss daran denken.
    doc.core_properties.comments = NACHBAU_KENNUNG

    ziel = Path(ziel)
    ziel.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(ziel))
    return ziel


def erzeuge_briefe(abrechnungen: list[Abrechnung], cfg: dict, ordner: Path,
                   log=None) -> tuple[list[Path], list[str]]:
    """Alle Briefe erzeugen — ohne die, für die es keinen Anlass gibt."""
    def melde(text):
        if log:
            log(text)

    vorlage = _vorlagen_dir() / "brief_vorlage.docx"
    if not vorlage.exists():
        raise ValueError(
            f"Die Briefvorlage fehlt:\n{vorlage}\n"
            f"Ohne sie lassen sich keine Briefe erzeugen.")

    if vorlage_ist_nachbau(vorlage):
        melde("Hinweis: Die Briefe entstehen noch aus der nachgebauten "
              "Vorlage, nicht aus der echten des Verlags. Schriften und "
              "Abstände können abweichen — bitte einen Brief ansehen, bevor "
              "alle verschickt werden.")

    ordner = Path(ordner)
    ordner.mkdir(parents=True, exist_ok=True)
    erzeugt: list[Path] = []
    uebergangen: list[str] = []
    for ab in abrechnungen:
        if not ab.brief:
            uebergangen.append(
                f"{ab.empfaenger.anzeigename}: {', '.join(ab.gruende)}")
            continue
        name = _sicherer_dateiname(
            f"{ab.empfaenger.dateiname_basis}_{ab.verguetungsart}-{ab.jahr}")
        ziel = ordner / f"{name}.docx"
        generiere_brief(vorlage, ab, cfg, ziel)
        erzeugt.append(ziel)
        melde(f"  {ziel.name}")
    melde(f"{len(erzeugt)} Briefe, {len(uebergangen)} ohne Brief.")
    return erzeugt, uebergangen


def wandle_nach_pdf(pfade: list[Path], log=None) -> list[Path]:
    """Die Briefe per Word nach PDF wandeln — nur unter Windows.

    Word wird EINMAL gestartet und am Ende wieder geschlossen; je Brief eine
    neue Word-Instanz zu starten dauert bei hundert Briefen ein Vielfaches.

    Unter Linux gibt es kein Word: dort wird der Schritt übersprungen und
    vermerkt. Die .docx sind dann trotzdem fertig — geprüft werden muss die
    Umwandlung aber auf einem Windows-Rechner.
    """
    def melde(text):
        if log:
            log(text)

    if sys.platform != "win32":
        melde("PDF-Umwandlung übersprungen — dafür wird Windows mit Word "
              "gebraucht. Die .docx-Dateien sind erzeugt.")
        return []

    # Bewusst lokal importiert: pywin32 gibt es nur unter Windows.
    import pythoncom
    import win32com.client

    # COM muss in JEDEM Thread einzeln angemeldet werden. Diese Funktion
    # läuft aus einem Arbeitsthread heraus (die Oberfläche darf nicht
    # einfrieren), und ohne diesen Aufruf scheitert schon das Dispatch mit
    # „CoInitialize has not been called“.
    pythoncom.CoInitialize()
    erzeugt: list[Path] = []
    fehlgeschlagen: list[str] = []
    word = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        # Ohne das bleibt Word bei der ersten Rückfrage stehen — etwa wenn
        # eine Datei noch offen ist — und niemand sieht das Fenster, weil es
        # unsichtbar gestartet wurde.
        word.DisplayAlerts = 0

        for pfad in pfade:
            pfad = Path(pfad)
            ziel = pfad.with_suffix(".pdf")
            doc = None
            try:
                doc = word.Documents.Open(str(pfad.resolve()))
                doc.SaveAs(str(ziel.resolve()), FileFormat=17)  # 17 = PDF
                erzeugt.append(ziel)
                melde(f"  {ziel.name}")
            except Exception as e:
                # Ein einzelner kaputter Brief darf nicht die übrigen
                # einundneunzig verhindern.
                fehlgeschlagen.append(f"{pfad.name}: {e}")
                melde(f"  ! {pfad.name} — {e}")
            finally:
                if doc is not None:
                    try:
                        doc.Close(False)
                    except Exception:
                        pass
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    if fehlgeschlagen:
        melde(f"{len(fehlgeschlagen)} Briefe konnten nicht umgewandelt "
              f"werden — die .docx sind aber da.")
    melde(f"{len(erzeugt)} PDF-Dateien erzeugt.")
    return erzeugt


def lies_stueckzahl(text) -> tuple[int | None, str]:
    """Eine eingetippte Stückzahl streng lesen.

    Gibt (Zahl, "") zurück oder (None, Klartextmeldung). Die Meldung ist für
    den Bediener gedacht und wird genau so angezeigt.

    Warum streng: ``float()`` macht aus dem deutschen Tausenderpunkt „1.234"
    klaglos die Zahl 1,234 und daraus gerundet eine **1**. Aus 1234 verkauften
    Exemplaren würde eines — ohne jede Meldung, mitten in einer Abrechnung
    über 20.000 €. Lieber einmal nachfragen als einmal falsch überweisen.

    Negative Zahlen sind erlaubt und gewollt: Rückgaben kommen im Bestand vor
    (im Musterbrief des Stadtarchivs steht „-9 Ex.").
    """
    roh = " ".join(str(text if text is not None else "").split())
    if not roh:
        return None, ""

    # Tausendertrennung wegnehmen, aber nur wo sie wirklich Tausender trennt:
    # „1.234", „12 345", „1.234.567".
    if re.fullmatch(r"-?\d{1,3}([. ]\d{3})+", roh):
        roh = roh.replace(".", "").replace(" ", "")

    if re.fullmatch(r"[+-]?\d+", roh):
        return int(roh), ""

    if re.fullmatch(r"[+-]?\d+[.,]\d+", roh):
        return None, (f"„{text}“ ist eine Kommazahl. Exemplare gibt es nur "
                      f"ganz — bitte eine ganze Zahl eintragen.")

    return None, (f"„{text}“ kann ich nicht als Stückzahl lesen.\n\n"
                  f"Bitte nur eine ganze Zahl eintragen, zum Beispiel 1234. "
                  f"Ein Minus ist erlaubt (etwa -9 für Rückgaben). Hat sich "
                  f"das Buch nicht verkauft, tragen Sie eine 0 ein.")


def rechenweg(posten: Posten) -> list[str]:
    """Wie dieser Betrag zustande kommt — als schriftliche Rechnung.

    In Excel konnte man in die Zelle klicken und die Formel lesen. Genau
    daran hängt das Vertrauen in eine Abrechnung über 20.000 €. Ein Betrag,
    den man nicht nachvollziehen kann, wird entweder blind geglaubt oder das
    Werkzeug wird nicht benutzt — beides schlecht.

    Jede Rechenzeile ist „\tBeschriftung\tWert“: das Fenster setzt den
    Wert an einen rechtsbündigen Tabulator, so stehen die Zahlen
    untereinander wie auf dem Papier. Als ein einziger Satz mit Kommas
    dazwischen war der Weg vom Ladenpreis zum Honorar kaum zu lesen.
    Die erste Zeile ist der Buchtitel.
    """
    b = posten.buch
    k = b.kondition
    zeilen = [b.titel]

    def ex(n) -> str:
        return f"{n} Ex."

    def pz(x) -> str:
        # „1,875 %“ statt „1.875 %“ — die Seite ist sonst durchgehend deutsch.
        return f"{x:g}".replace(".", ",")

    # 1. Die Menge
    zeilen.append(f"\tverkauft\t{ex(posten.verkauft)}")
    abgezogen = False
    if posten.eigenkauf:
        zeilen.append(f"\tabzüglich Eigenkauf\t− {ex(posten.eigenkauf)}")
        abgezogen = True
    if posten.korrektur:
        zeichen = "+" if posten.korrektur > 0 else "−"
        zeilen.append(f"\tKorrektur aus dem Vorjahr\t"
                      f"{zeichen} {ex(abs(posten.korrektur))}")
        abgezogen = True
    if k.freimenge:
        ab = k.freimenge + 1
        stand = kumulierte_menge(b, posten.jahr)
        if stand is None:
            zeilen.append(f"\tHonorar ab dem {ab}. Exemplar — ohne "
                          f"Vorjahreszahlen gilt das als erreicht")
        elif stand < 0:
            zeilen.append(f"\tnoch frei bis zum {ab}. Exemplar\t"
                          f"− {ex(-stand)}")
        else:
            zeilen.append(f"\tdas {ab}. Exemplar war schon vor "
                          f"{posten.jahr} erreicht")
        abgezogen = True
    if abgezogen or posten.verguetungs_ex != posten.verkauft:
        zeilen.append(f"\tvergütet werden\t"
                      f"{ex(max(posten.verguetungs_ex, 0))}")
    zeilen.append("")

    # 2. Der Betrag je Exemplar
    if k.betrag_je_ex is not None:
        zeilen.append(f"\tfest vereinbart je Exemplar\t"
                      f"{euro(k.betrag_je_ex)}")
    elif k.ladenpreis is not None:
        zeilen.append(f"\tLadenpreis\t{euro(k.ladenpreis)}")
        wert = k.ladenpreis
        if k.mwst_im_preis and not k.mwst_aufschlagen:
            wert /= 1 + k.mwst_im_preis / 100
            zeilen.append(f"\tohne {pz(k.mwst_im_preis)} % MwSt\t{euro(wert)}")
        elif k.mwst_im_preis:
            wert *= 1 + k.mwst_im_preis / 100
            zeilen.append(f"\tmit {pz(k.mwst_im_preis)} % MwSt\t{euro(wert)}")
        if k.rabatt_anwenden:
            wert *= 1 - k.verlagsrabatt / 100
            zeilen.append(f"\tabzüglich {pz(k.verlagsrabatt)} % Verlagsrabatt"
                          f"\t{euro(wert)}")
        if k.teiler != 1:
            if posten.satz_prozent is not None:
                zeilen.append(f"\tdavon {pz(posten.satz_prozent)} % Honorar\t"
                              f"{euro(wert * posten.satz_prozent / 100)}")
            zeilen.append(f"\tgeteilt durch {k.teiler} Mitautoren — "
                          f"je Exemplar\t{euro(posten.satz)}")
        elif posten.satz_prozent is not None:
            zeilen.append(f"\tdavon {pz(posten.satz_prozent)} % Honorar — "
                          f"je Exemplar\t{euro(posten.satz)}")
        else:
            zeilen.append(f"\tje Exemplar\t{euro(posten.satz)}")
    if k.staffel:
        stufen = []
        untere = 1
        for grenze, satz in k.staffel:
            if grenze is None:
                stufen.append(f"ab {untere} Ex. {pz(satz)} %")
            else:
                stufen.append(f"{untere}–{grenze} Ex. {pz(satz)} %")
                untere = grenze + 1
        zeilen.append("\tStaffel laut Vertrag: " + " · ".join(stufen))
        stand = kumulierte_menge(b, posten.jahr)
        if stand is None:
            # Ohne Vorjahreszahlen lässt sich die Stufe nicht bestimmen. Das
            # muss dastehen — sonst zeigt die Auskunft eine Staffel und
            # rechnet daneben mit einem anderen Satz.
            zeilen.append(
                f"\tWie viele Exemplare vor {posten.jahr} verkauft wurden, "
                f"ist hier nicht hinterlegt. Deshalb wird mit dem zuletzt "
                f"vereinbarten Satz von {pz(posten.satz_prozent)} % gerechnet.")
        else:
            zeilen.append(
                f"\tBis Ende {posten.jahr - 1} waren es {stand} Exemplare — "
                f"damit gilt die Stufe mit {pz(posten.satz_prozent)} %.")
    zeilen.append("")

    # 3. Das Ergebnis
    if k.freimenge and posten.verguetungs_ex < 0:
        zeilen.append(f"\tes fehlen noch {posten.verguetungs_ex * -1} "
                      f"Exemplare bis zum Honorar\t{euro(0)}")
    elif k.schwelle_zehn and posten.netto == 0 and posten.verguetungs_ex < 10:
        zeilen.append(f"\tunter zehn Exemplaren entfällt das Honorar "
                      f"laut Vertrag\t{euro(0)}")
    else:
        zeilen.append(f"\t{ex(posten.verguetungs_ex)} × {euro(posten.satz)}"
                      f"\t{euro(posten.netto)}")
    if posten.mwst:
        zeilen.append(f"\tzuzüglich {pz(b.mwst_satz)} % Mehrwertsteuer\t"
                      f"{euro(posten.mwst)}")
        zeilen.append(f"\tzusammen\t{euro(posten.brutto)}")
    return zeilen
