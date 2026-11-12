#!/usr/bin/env python3
"""
Presseverteiler Core-Logik
==========================
Reine Datenlogik ohne UI-Abhängigkeiten — wird von der GUI und vom
kopflosen `probelauf.py` gleichermaßen benutzt.

Aufgabe: für ein neues Buch die Pressekontakte auswählen, die es bekommen
sollen. Bisher sucht der Pressereferent sie je Buch von Hand aus der
Access-Liste heraus — rund 30 Minuten und fehleranfällig.

Das Tool schlägt die Empfänger aus den Listen ähnlicher früherer Bücher vor
(`Pressehistorie.xlsx`). Der erste Entwurf fragte stattdessen, wofür jeder
Kontakt zuständig ist. Das haben die alten Presselisten widerlegt: Die Hälfte
der Empfänger hat mit dem Verbreitungsgebiet nichts zu tun (Befunde in
`WIP_Presselisten-Auswertung.md`). Regionen und Themen je Kontakt stehen
weiter in `Presseverteiler.xlsx`, dienen aber nur noch als Suchhilfe und für
Kontakte ohne Vorgeschichte.

Die Adressen bleiben in Access; das Tool liest den Export bei jedem Start
frisch und schreibt nie zurück. Beide Mappen sind Excel und kein JSON, damit
der Verlag sie auch ohne das Tool öffnen und notfalls von Hand weiterführen
kann.
"""

from __future__ import annotations

import io
import math
import os
import re
import shutil
import sys
import tempfile
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------
# Daten-Ordner (Config, Mappe, Ausgabe)
# ---------------------------------------------------------------------

def _base_dir() -> Path:
    """Ordner neben der .exe (PyInstaller-Build) bzw. neben dem Tool-Code."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


# Neben der .exe liegt nur die config.json: Welcher Access-Export, gilt je
# Rechner. Die beiden Mappen liegen auf dem NAS, siehe unten.
APP_DIR = _base_dir()
CONFIG_PFAD = APP_DIR / "config.json"


def _netz_ablage() -> Path:
    """Wo der Verlag die Presselisten ablegt: auf dem NAS, nicht neben der .exe.

    Dort liegen alle Listen seit 2020, und nur dort findet sie der Rest des
    Verlags. Unter Linux hängt gvfs die Freigabe unter diesem Pfad ein.
    """
    ordner = "VR-Akquise/MK/Presse/Presselisten für Rezensionsversand"
    if sys.platform == "win32":
        return Path(r"\\VR-Archiv\vr-intern") / ordner
    return Path(f"/run/user/{os.getuid()}/gvfs/smb-share:"
                f"server=vr-archiv.local,share=vr-intern") / ordner


AUSGABE_ORDNER = _netz_ablage()

# Historie und Mappe liegen anders als bei den übrigen Tools nicht neben der
# .exe, sondern bei den Listen auf dem NAS. Neben der .exe lägen sie in der
# lokalen Kopie, die launch.ps1 unter %LOCALAPPDATA% anlegt: ungesichert, und
# jeder Rechner lernte für sich. Die Historie ist aber das, woraus der
# Vorschlag lernt, die geht nicht verloren und gehört allen.
DATEN_ORDNER = "_Presseverteiler"
VERTEILER_PFAD = AUSGABE_ORDNER / DATEN_ORDNER / "Presseverteiler.xlsx"
HISTORIE_PFAD = AUSGABE_ORDNER / DATEN_ORDNER / "Pressehistorie.xlsx"
# Der Access-Export liegt mit dort, damit nicht jeder Rechner seinen eigenen,
# verschieden alten hat. Neu exportiert wird an genau diese Stelle.
ACCESS_PFAD = AUSGABE_ORDNER / DATEN_ORDNER / "Access_Export.xlsx"


def _mappe_lesen(pfad):
    """Öffnet eine xlsx aus dem Arbeitsspeicher.

    Eine xlsx ist ein Zip, und beim Lesen springt openpyxl darin umher. Über
    die gvfs-Einbindung des NAS schlägt das immer wieder fehl („Input/output
    error“, „File is not a zip file“), am Stück gelesen klappt es zuverlässig.
    Nebenbei hält das Programm so keine Datei auf dem NAS offen, während
    jemand den Access-Export neu ablegt.
    """
    return openpyxl.load_workbook(io.BytesIO(Path(pfad).read_bytes()),
                                  read_only=True, data_only=True)


def _mappe_ablegen(wb, pfad: Path):
    """Schreibt eine Mappe ohne Halbfertiges am Ziel: erst im Speicher, dann
    am Stück als .neu.xlsx, dann umbenennen. Am Stück aus demselben Grund wie
    beim Lesen; das Umbenennen sorgt dafür, dass ein Abbruch mitten im
    Schreiben die alte Mappe nicht beschädigt."""
    puffer = io.BytesIO()
    wb.save(puffer)
    pfad.parent.mkdir(parents=True, exist_ok=True)
    zwischen = pfad.with_name(pfad.stem + ".neu.xlsx")
    zwischen.write_bytes(puffer.getvalue())
    zwischen.replace(pfad)

REGION = "Region"
THEMA = "Thema"
ACHSEN = (REGION, THEMA)


# ---------------------------------------------------------------------
# Typ-Helfer für das Lesen aus Excel
# ---------------------------------------------------------------------
# Excel verwischt beim Speichern von Hand die Unterschiede, auf die es hier
# ankommt: aus der PLZ „69469“ wird die Zahl 69469, aus der ID 17404 ein
# 17404.0. Jedes Feld geht deshalb durch einen dieser Helfer, nie direkt aus
# der Zelle in das Modell.

def _text(wert) -> str:
    """Zellwert als Text. Aus einer Ganzzahl wird „69469“, nicht „69469.0“."""
    if wert is None:
        return ""
    if isinstance(wert, float) and wert.is_integer():
        return str(int(wert))
    if isinstance(wert, datetime):
        return wert.strftime("%d.%m.%Y")
    return " ".join(str(wert).split()) if isinstance(wert, str) else str(wert)


def _ganzzahl(wert) -> int | None:
    try:
        return int(round(float(str(wert).replace(",", ".").strip())))
    except (TypeError, ValueError):
        return None


def _haken(wert) -> bool:
    """„x“, „X“, „ja“, 1 — alles, was man in Excel als Haken tippen würde."""
    return _text(wert).strip().lower() in {"x", "ja", "j", "1", "wahr", "true"}


def _liste(wert) -> list[str]:
    """„Kraichgau; Landkreis Karlsruhe“ → zwei Einträge.

    Nur das Semikolon trennt. Ein Komma darf in einem Schlagwort vorkommen
    (die Shop-Kategorien heißen „Nachname, Vorname“ — wer je eine davon
    übernimmt, soll sie nicht zerschnitten bekommen).
    """
    return [t.strip() for t in _text(wert).split(";") if t.strip()]


def _sicherer_dateiname(text: str) -> str:
    """Was Windows im Dateinamen nicht duldet, muss weg — Umlaute dürfen
    bleiben."""
    for z in '<>:"/\\|?*':
        text = text.replace(z, "-")
    return " ".join(text.split()).strip(". ")


def alphabetisch(text: str) -> str:
    """Vergleichsform für Sortierung und Suche: Ä zählt wie A, wie im
    Telefonbuch. Sonst stünde Überlingen hinter Zürich, und wer „Sudkurier“
    tippt, fände den SÜDKURIER nicht."""
    zerlegt = unicodedata.normalize("NFKD", text.casefold())
    return "".join(z for z in zerlegt if not unicodedata.combining(z))


# ---------------------------------------------------------------------
# Access-Export einlesen
# ---------------------------------------------------------------------
# Derselbe Export wie für den Mailinglisten-Abgleich: die komplette
# Kundentabelle, nicht eine Abfrage. Auf feste Zeilennummern ist kein
# Verlass — die Kopfzeile wird gesucht.

ACCESS_PFLICHT = {"ID", "Name", "Presse/ZS"}


def finde_kopfzeile(zeilen: list[tuple], pflichtspalten: set[str]) -> int:
    """Index der ersten Zeile, die alle Pflichtspalten enthält (0-basiert)."""
    for i, zeile in enumerate(zeilen):
        vorhanden = {str(z).strip() for z in zeile if z is not None}
        if pflichtspalten <= vorhanden:
            return i
    raise ValueError(
        "Kopfzeile nicht gefunden — erwartet wurden die Spalten "
        + ", ".join(sorted(pflichtspalten))
    )


def _blatt_als_dicts(zeilen: list[tuple],
                     pflichtspalten: set[str]) -> list[dict]:
    kopf_i = finde_kopfzeile(zeilen, pflichtspalten)
    spalten = [(str(z).strip() if z is not None else f"_leer{i}")
               for i, z in enumerate(zeilen[kopf_i])]
    daten = []
    for zeile in zeilen[kopf_i + 1:]:
        if not any(z not in (None, "") for z in zeile):
            continue
        daten.append({name: (zeile[i] if i < len(zeile) else None)
                      for i, name in enumerate(spalten)})
    return daten


def lade_tabelle(pfad, pflichtspalten: set[str]) -> list[dict]:
    wb = _mappe_lesen(pfad)
    try:
        zeilen = [tuple(z) for z in wb.worksheets[0].iter_rows(values_only=True)]
    finally:
        wb.close()
    return _blatt_als_dicts(zeilen, pflichtspalten)


# Wer als Presse gilt. In Access gibt es zwei Felder dafür, und sie sind
# nicht deckungsgleich: `Presse/ZS` (Presse, ZS = Zeitschrift, dazu Tippfehler
# wie „Prese“) und die Mailing-Kategorie `VIP/W/K/X` = „P“. Gemessen im
# August 2026 tragen 6 Sätze nur das „P“ und 13 Pressesätze die Kategorie
# „VIP“ — wer nur eines der Felder prüft, verliert sie. „BIB“ (Bibliotheken)
# steht zwar in derselben Spalte, ist aber keine Presse.
def ist_presse(satz: dict) -> bool:
    presse_zs = _text(satz.get("Presse/ZS")).upper()
    kategorie = _text(satz.get("VIP/W/K/X")).upper()
    return (presse_zs.startswith("PRES") or presse_zs in {"ZS", "ZW"}
            or kategorie == "P")


@dataclass
class Kontakt:
    """Ein Pressekontakt, wie er in Access steht. Die ganze Zeile bleibt
    erhalten — der Serienbrief braucht Anrede, Titel und Anschrift."""
    access_id: int
    satz: dict

    def feld(self, name: str) -> str:
        return _text(self.satz.get(name))

    @property
    def institution(self) -> str:
        return self.feld("Institution")

    @property
    def abteilung(self) -> str:
        return " / ".join(t for t in (self.feld("Abteilung"),
                                      self.feld("Unterabteilung")) if t)

    @property
    def person(self) -> str:
        return " ".join(t for t in (self.feld("Vorname"),
                                    self.feld("Name")) if t)

    @property
    def ort(self) -> str:
        return " ".join(t for t in (self.feld("PLZ"), self.feld("Ort")) if t)

    @property
    def medium(self) -> str:
        # „zeitung“, „Zeitung“, „Tageszeitung“ — für die Anzeige reicht es,
        # den ersten Buchstaben zu vereinheitlichen.
        m = self.feld("Medium")
        return m[:1].upper() + m[1:] if m else ""

    @property
    def mail(self) -> str:
        return self.feld("eMail").strip().strip(";,")

    @property
    def mails(self) -> list[str]:
        """Die einzelnen Adressen des Felds.

        In rund 240 Pressesätzen stehen mehrere Adressen in einem Feld, meist
        mit Komma getrennt („redaktion@…, info@…“). Outlook trennt im
        BCC-Feld aber nur am Semikolon — unzerlegt wäre das eine ungültige
        Adresse. Was kein @ hat („palmyra-verlag“), ist keine Adresse.
        """
        return [t for t in re.split(r"[,;\s]+", self.mail) if "@" in t]

    @property
    def sortierung(self) -> tuple:
        return (alphabetisch(self.institution) or "~",
                alphabetisch(self.abteilung), alphabetisch(self.feld("Name")))

    @property
    def hat_anschrift(self) -> bool:
        return bool(self.feld("Straße") and self.feld("Ort"))


def lade_presse(pfad) -> tuple[list[Kontakt], int]:
    """Alle Pressekontakte aus dem Access-Export.

    Gibt (Kontakte, Zahl der übergangenen) zurück. Übergangen werden Sätze,
    die in „Datensatz gelöscht“ einen Vermerk tragen (verzogen, verstorben,
    „Mailing kam zurück“) — an die soll kein Buch mehr gehen.
    """
    kontakte, uebergangen = [], 0
    for satz in lade_tabelle(pfad, ACCESS_PFLICHT):
        if not ist_presse(satz):
            continue
        if _text(satz.get("Datensatz gelöscht")):
            uebergangen += 1
            continue
        access_id = _ganzzahl(satz.get("ID"))
        if access_id is None:
            continue
        kontakte.append(Kontakt(access_id, satz))
    kontakte.sort(key=lambda k: k.sortierung)
    return kontakte, uebergangen


# ---------------------------------------------------------------------
# Schlagworte
# ---------------------------------------------------------------------
# Startliste, abgeschrieben aus dem Kategoriebaum des Shops (prod, Stand
# 07.09.2026, `shopware_publisher/kategorien_prod.json`). Sie steht hier fest
# im Code statt aus dem Shop geholt zu werden, weil der Baum ohnehin von Hand
# nachgearbeitet werden musste — und weil die Liste nach dem ersten Start im
# Blatt „Schlagworte“ der Mappe lebt. Ab dann gilt die Mappe, nicht dieser
# Code.
#
# Weggelassen sind Kategorien, die nur nach Format oder Alphabet ordnen
# (Audio-CDs, DVDs, E-Books, Kalender, Bildbände, Schriftenreihen,
# „Ortsgeschichte/A-B“ …): danach wählt niemand Presse aus.
#
# „gehört zu“ macht Bücher vergleichbar: Ein Buch über den Pfinzgau gilt damit
# auch als ähnlich zu Büchern über den Landkreis Karlsruhe, die Region
# Karlsruhe und ganz Baden. Mehrere Obergebiete sind erlaubt, weil Landschaften
# sich nicht an Kreisgrenzen halten (der Kraichgau liegt in drei Landkreisen).
#
# Herkunft „Ergänzung“: im Shop gibt es diese nicht, für die Presse braucht
# man sie aber. „Geschichte“ fehlt im Shop als Oberbegriff, obwohl „his“ das
# häufigste Fachgebiet der Pressekontakte in Access ist; „Pfalz“ ist nicht
# „Kurpfalz“ (die Rheinpfalz berichtet nicht über Heidelberg).
STANDARD_SCHLAGWORTE: list[tuple[str, str, str, str]] = [
    # (Schlagwort, Achse, gehört zu — mit ; getrennt, Herkunft)

    # -- Regionen: Großräume --
    ("Baden-Württemberg", REGION, "", "Ergänzung"),
    ("Baden", REGION, "Baden-Württemberg", "Ergänzung"),
    ("Württemberg", REGION, "Baden-Württemberg", "Shop"),
    ("Pfalz", REGION, "Rheinland-Pfalz", "Ergänzung"),
    ("Rheinland-Pfalz", REGION, "", "Shop"),
    ("Saarland", REGION, "", "Ergänzung"),
    ("Hessen", REGION, "", "Shop"),
    ("Franken", REGION, "", "Shop"),
    ("Unterfranken", REGION, "Franken", "Ergänzung"),
    ("Elsass-Lothringen", REGION, "", "Shop"),
    ("Schweiz", REGION, "", "Shop"),
    ("Ostsee", REGION, "", "Shop"),
    ("Metropolregion Rhein-Neckar", REGION, "", "Shop"),
    ("Kurpfalz", REGION, "Metropolregion Rhein-Neckar", "Shop"),
    ("Karlsruhe (Region)", REGION, "Baden", "Shop"),

    # -- Regionen: Kreise und Städte --
    ("Landkreis Karlsruhe", REGION, "Karlsruhe (Region)", "Shop"),
    ("Karlsruhe (Stadt)", REGION, "Karlsruhe (Region)", "Ergänzung"),
    ("Landkreis Rastatt", REGION, "Karlsruhe (Region)", "Shop"),
    ("Baden-Baden", REGION, "Karlsruhe (Region)", "Ergänzung"),
    ("Enzkreis", REGION, "Baden", "Shop"),
    ("Pforzheim", REGION, "Baden", "Ergänzung"),
    ("Rhein-Neckar-Kreis", REGION, "Baden; Metropolregion Rhein-Neckar",
     "Shop"),
    ("Heidelberg", REGION, "Baden; Metropolregion Rhein-Neckar", "Ergänzung"),
    ("Mannheim", REGION, "Baden; Metropolregion Rhein-Neckar", "Ergänzung"),
    ("Neckar-Odenwald-Kreis", REGION, "Baden; Metropolregion Rhein-Neckar",
     "Shop"),
    ("Rhein-Pfalz-Kreis", REGION, "Pfalz; Metropolregion Rhein-Neckar",
     "Shop"),
    ("Landkreis Lörrach", REGION, "Baden", "Shop"),
    ("Landkreis Rottweil", REGION, "Württemberg", "Shop"),
    ("Hohenlohekreis", REGION, "Württemberg", "Shop"),
    ("Landkreis Heilbronn", REGION, "Württemberg", "Ergänzung"),
    ("Heilbronn (Stadt)", REGION, "Württemberg", "Ergänzung"),
    ("Landkreis Calw", REGION, "Württemberg", "Ergänzung"),
    ("Region Stuttgart", REGION, "Württemberg", "Ergänzung"),
    ("Stuttgart", REGION, "Region Stuttgart", "Ergänzung"),

    # Ergänzt im November 2026, als die Pressekontakte zum ersten Mal
    # eingeschätzt wurden: Dort sitzt Lokalpresse, für die es sonst nur
    # „Baden“, „Württemberg“ oder „Pfalz“ gäbe. Ein Buch über Freiburg ließe
    # sich dann nicht von einem über den Kraichgau unterscheiden.
    ("Freiburg / Breisgau", REGION, "Baden", "Ergänzung"),
    ("Landkreis Waldshut", REGION, "Baden", "Ergänzung"),
    ("Schwarzwald-Baar-Kreis", REGION, "Baden; Schwarzwald", "Ergänzung"),
    ("Landkreis Konstanz", REGION, "Baden; Bodensee", "Ergänzung"),
    ("Landkreis Tübingen", REGION, "Württemberg", "Ergänzung"),
    ("Landkreis Reutlingen", REGION, "Württemberg; Schwäbische Alb",
     "Ergänzung"),
    ("Zollernalbkreis", REGION, "Württemberg; Schwäbische Alb", "Ergänzung"),
    ("Landkreis Freudenstadt", REGION, "Württemberg; Schwarzwald",
     "Ergänzung"),
    ("Landkreis Schwäbisch Hall", REGION, "Württemberg", "Ergänzung"),
    ("Ostwürttemberg", REGION, "Württemberg", "Ergänzung"),
    ("Ulm / Alb-Donau", REGION, "Württemberg", "Ergänzung"),
    ("Oberschwaben", REGION, "Württemberg", "Ergänzung"),
    ("Bodenseekreis", REGION, "Württemberg; Bodensee", "Ergänzung"),
    ("Ludwigshafen", REGION, "Pfalz; Metropolregion Rhein-Neckar",
     "Ergänzung"),
    ("Speyer", REGION, "Pfalz; Metropolregion Rhein-Neckar", "Ergänzung"),
    ("Südpfalz", REGION, "Pfalz", "Ergänzung"),
    ("Westpfalz", REGION, "Pfalz", "Ergänzung"),
    ("Rheinhessen", REGION, "Rheinland-Pfalz", "Ergänzung"),

    # -- Regionen: Landschaften --
    ("Kraichgau", REGION,
     "Landkreis Karlsruhe; Rhein-Neckar-Kreis; Landkreis Heilbronn", "Shop"),
    ("Pfinzgau", REGION, "Landkreis Karlsruhe", "Shop"),
    ("Bergstraße", REGION, "Rhein-Neckar-Kreis; Hessen", "Shop"),
    ("Odenwald", REGION, "Neckar-Odenwald-Kreis; Rhein-Neckar-Kreis; Hessen",
     "Shop"),
    ("Ortenau", REGION, "Baden", "Shop"),
    # Müllheim und Badenweiler liegen im Kreis Breisgau-Hochschwarzwald.
    ("Markgräflerland", REGION, "Landkreis Lörrach; Freiburg / Breisgau",
     "Shop"),
    ("Oberrhein", REGION, "Baden", "Shop"),
    ("Schwarzwald", REGION, "Baden-Württemberg", "Shop"),
    ("Schwäbische Alb", REGION, "Württemberg", "Shop"),
    ("Bodensee", REGION, "Baden-Württemberg", "Shop"),
    ("Zabergäu", REGION, "Landkreis Heilbronn", "Shop"),
    ("Tauberfranken", REGION, "Franken; Baden-Württemberg", "Shop"),
    ("Taubertal", REGION, "Tauberfranken", "Shop"),
    ("Haardt", REGION, "Pfalz", "Shop"),
    ("Weinstraße", REGION, "Pfalz", "Shop"),

    # -- Themen --
    ("Geschichte", THEMA, "", "Ergänzung"),
    # Ortsgeschichte gehört bewusst NICHT zu Geschichte: Ein Ortsbuch ist Sache
    # der Lokalpresse. Unter „Geschichte“ sähe jede Ortschronik den
    # Fachbüchern ähnlich, deren Listen an Geschichtszeitschriften gehen.
    ("Ortsgeschichte", THEMA, "", "Shop"),
    ("Kirchengeschichte", THEMA, "Geschichte", "Shop"),
    ("Jüdische Geschichte", THEMA, "Geschichte", "Shop"),
    ("Alltags- / Kulturgeschichte", THEMA, "Geschichte", "Shop"),
    ("Archäologie", THEMA, "Geschichte", "Shop"),
    ("Ortssippenbuch (Familienbuch)", THEMA, "Ortsgeschichte", "Shop"),
    ("Historische Landkarten", THEMA, "Geschichte", "Shop"),
    ("Politik und Soziales", THEMA, "Geschichte", "Shop"),
    ("Kriegs- und Nachkriegszeiten", THEMA, "Politik und Soziales", "Shop"),
    ("Migration und Vertreibung", THEMA, "Politik und Soziales", "Shop"),
    ("Revolution 1848/49", THEMA, "Politik und Soziales", "Shop"),
    ("Technik und Wirtschaft", THEMA, "", "Shop"),
    ("Eisenbahn", THEMA, "Technik und Wirtschaft", "Shop"),
    # IfU = Institut für Unternehmensgeschichte
    ("IfU", THEMA, "Technik und Wirtschaft; Geschichte", "Shop"),
    ("Kunst und Musik", THEMA, "", "Shop"),
    ("Bildende Kunst", THEMA, "Kunst und Musik", "Shop"),
    ("Darstellende Kunst / Theater", THEMA, "Kunst und Musik", "Shop"),
    ("Musik", THEMA, "Kunst und Musik", "Shop"),
    ("Belletristik / Literatur", THEMA, "", "Shop"),
    ("Belletristik", THEMA, "Belletristik / Literatur", "Shop"),
    ("Krimis", THEMA, "Belletristik / Literatur", "Shop"),
    ("Literaturgeschichte", THEMA, "Belletristik / Literatur", "Shop"),
    ("Lyrik", THEMA, "Belletristik / Literatur", "Shop"),
    ("Mundart", THEMA, "Belletristik / Literatur", "Shop"),
    ("Kinder- u. Jugendliteratur", THEMA, "", "Shop"),
    ("Natur und Ökologie", THEMA, "", "Shop"),
    ("Reise- / Kultur- / Naturführer", THEMA, "", "Shop"),
    ("Wanderführer", THEMA, "Reise- / Kultur- / Naturführer", "Shop"),
    ("Kulinarisches", THEMA, "", "Shop"),
    ("Schule / Lernen", THEMA, "", "Shop"),
    # Die beiden häufigsten Access-Fachgebiete, die der Shop nicht kennt
    # („sport“ 26-mal, „architektur“ 19-mal). Vereinschroniken und Bücher
    # über Bauten gibt es im Programm; die Sport- und Architekturredaktionen
    # hätten sonst kein Schlagwort.
    ("Sport", THEMA, "", "Ergänzung"),
    ("Architektur / Denkmalpflege", THEMA, "", "Ergänzung"),
]


@dataclass
class Schlagwort:
    name: str
    achse: str
    gehoert_zu: list[str] = field(default_factory=list)
    herkunft: str = ""


# ---------------------------------------------------------------------
# Startwerte aus den alten Access-Feldern
# ---------------------------------------------------------------------
# In Access gibt es Region 1–3 und Fachgebiet 1–3, aber gemessen im August
# 2026 sind 76 % der Pressesätze ohne Region und 64 % ohne Fachgebiet, und
# die Schreibweisen gehen durcheinander. Was sich eindeutig übersetzen lässt,
# wird als *ungeprüfter Vorschlag* übernommen; der Rest landet im Blatt
# „Altwerte“, damit man sieht, was liegen geblieben ist.
#
# Eine leere Liste heißt „bewusst nichts“: „Berlin“ steht in Access bei den
# Hauptstadtredaktionen — das ist ihr Sitz, nicht ihr Berichtsgebiet. Und
# „allg“ bzw. „reg“ sagen gerade, dass es keine Themenbeschränkung gibt.
UMSCHLUESSEL_REGION: dict[str, list[str]] = {
    "südwest": ["Baden-Württemberg"],
    "bawü": ["Baden-Württemberg"],
    "baden-württemberg": ["Baden-Württemberg"],
    "baden": ["Baden"],
    "württ": ["Württemberg"],
    "oberschwaben": ["Oberschwaben"],
    "aalen": ["Ostwürttemberg"],
    "stuttgart": ["Stuttgart"],
    "odenwald": ["Odenwald"],
    "pfalz": ["Pfalz"],
    "pf": ["Pforzheim"],
    "kraichgau": ["Kraichgau"],
    "kurpfalz": ["Kurpfalz"],
    "rhein-neckar": ["Metropolregion Rhein-Neckar"],
    "rhein-main": ["Hessen"],
    "hessen": ["Hessen"],
    "rheinhessen": ["Rheinhessen"],
    "rheinland-pfalz": ["Rheinland-Pfalz"],
    "unterfranken": ["Unterfranken"],
    "saarland": ["Saarland"],
    "franken": ["Franken"],
    "schweinfurt": ["Unterfranken"],
    "kitzingen": ["Unterfranken"],
    "würzburg": ["Unterfranken"],
    "tauberfranken": ["Tauberfranken"],
    "enz": ["Enzkreis"],
    "schwäbische alb": ["Schwäbische Alb"],
    "schwarzwald": ["Schwarzwald"],
    "bodensee": ["Bodensee"],
    "schweiz": ["Schweiz"],
    "graubünden": ["Schweiz"],
    "ka": ["Karlsruhe (Region)"],
    "karlsruhe": ["Karlsruhe (Region)"],
    "albtal": ["Karlsruhe (Region)"],
    "albgau": ["Karlsruhe (Region)"],
    "bruchsal": ["Landkreis Karlsruhe"],
    "rastatt": ["Landkreis Rastatt"],
    "bühl": ["Landkreis Rastatt"],
    "murgtal": ["Landkreis Rastatt"],
    "achern": ["Ortenau"],
    "mannheim": ["Mannheim"],
    "hd": ["Heidelberg"],
    "rnk": ["Rhein-Neckar-Kreis"],
    "breisgau-hochschwarzwald": ["Freiburg / Breisgau"],
    "hohenlohe": ["Hohenlohekreis"],
    "berlin": [],
    "allg": [],
}

UMSCHLUESSEL_FACH: dict[str, list[str]] = {
    "his": ["Geschichte"],
    "neueste g.": ["Geschichte"],
    "nat": ["Natur und Ökologie"],
    "natur": ["Natur und Ökologie"],
    "jagd": ["Natur und Ökologie"],
    "lit": ["Belletristik / Literatur"],
    "theol": ["Kirchengeschichte"],
    "kunst": ["Kunst und Musik"],
    "mus": ["Musik"],
    # „klass“ heißt klassische Altertumswissenschaft, nicht Klassik: Es
    # steht bei Germania, Gnomon, Antike Welt und den Bonner Jahrbüchern,
    # bei keinem einzigen Sender.
    "klass": ["Archäologie"],
    "bühne": ["Darstellende Kunst / Theater"],
    "bahn": ["Eisenbahn"],
    "jud": ["Jüdische Geschichte"],
    "päd": ["Schule / Lernen"],
    "wirtschaft": ["Technik und Wirtschaft"],
    "tech": ["Technik und Wirtschaft"],
    "tec": ["Technik und Wirtschaft"],
    "gen": ["Ortssippenbuch (Familienbuch)"],
    "archä": ["Archäologie"],
    "ns": ["Kriegs- und Nachkriegszeiten"],
    "pol": ["Politik und Soziales"],
    "mil": ["Kriegs- und Nachkriegszeiten"],
    "militär": ["Kriegs- und Nachkriegszeiten"],
    "freizeit": ["Reise- / Kultur- / Naturführer"],
    "sport": ["Sport"],
    "architektur": ["Architektur / Denkmalpflege"],
    "arch": ["Architektur / Denkmalpflege"],
    "allg": [],
    "reg": [],
}

REGION_FELDER = ("Region 1", "Region 2", "Region 3")
FACH_FELDER = ("Fachgebiet 1", "Fachgebiet 2", "Fachgebiet 3")


def startwerte(kontakt: Kontakt) -> tuple[set, set, list[tuple[str, str]]]:
    """(Regionen, Themen, nicht übersetzte (Feld, Wert)) aus den Access-Feldern."""
    regionen, themen, offen = set(), set(), []
    for felder, tabelle, ziel in ((REGION_FELDER, UMSCHLUESSEL_REGION, regionen),
                                  (FACH_FELDER, UMSCHLUESSEL_FACH, themen)):
        for feld in felder:
            wert = kontakt.feld(feld)
            if not wert:
                continue
            treffer = tabelle.get(wert.lower())
            if treffer is None:
                offen.append((feld.rsplit(" ", 1)[0], wert))
            else:
                ziel.update(treffer)
    return regionen, themen, offen


# ---------------------------------------------------------------------
# Der Verteiler: Regionen und Themen je Kontakt
# ---------------------------------------------------------------------
# Medienart und Ressort gehen nicht in den Vorschlag ein. Erfasst sind sie,
# weil sie bei der ersten Einschätzung fast umsonst mit anfielen und beim
# Durchsehen helfen. Feste Werte statt Freitext, damit man später danach filtern kann;
# das Access-Feld „Medium“ zeigt, was sonst daraus wird („Zeitung“,
# „zeitung“, „Tageszeitung“, „TV“, „Fernsehen“).
MEDIENARTEN = ["Tageszeitung", "Wochenblatt", "Amtsblatt", "Magazin",
               "Fachzeitschrift", "Hörfunk", "Fernsehen", "Online", "Agentur",
               "Person", "Sonstige"]
RESSORTS = ["Lokales", "Kultur", "Rezensionen", "Chefredaktion",
            "Redaktion allgemein", "anderes Ressort"]


@dataclass
class Eintrag:
    """Was über einen Pressekontakt bekannt ist, außer der Anschrift.

    Die Anzeigefelder (Institution, Name …) sind eine Abschrift aus Access,
    damit die Mappe in Excel lesbar ist. Maßgeblich für Adressen bleibt
    Access — die Abschrift wird bei jedem Speichern erneuert.
    """
    access_id: int
    regionen: set = field(default_factory=set)
    themen: set = field(default_factory=set)
    immer: bool = False
    geprueft: bool = False
    bemerkung: str = ""
    anzeige: dict = field(default_factory=dict)
    # Woher ein ungeprüfter Vorschlag kommt und wie sicher er ist. Bei der
    # ersten Einschätzung (presseverteiler/einschaetzung.py) steht hier, was
    # der Kontakt ist („Lokalredaktion der BNN in Bretten“). Mit dieser
    # Begründung lässt sich ein Vorschlag in Sekunden bestätigen, ohne sie
    # müsste man jeden Kontakt neu recherchieren. „unsicher“ markiert, was
    # zuerst durchgesehen werden sollte.
    begruendung: str = ""
    unsicher: bool = False
    medienart: str = ""
    ressort: str = ""

    @property
    def eingeordnet(self) -> bool:
        return bool(self.regionen or self.themen or self.immer)


@dataclass
class Verteiler:
    schlagworte: dict[str, Schlagwort] = field(default_factory=dict)
    eintraege: dict[int, Eintrag] = field(default_factory=dict)
    # Was aus Access nicht zu übersetzen war, als {(Feld, Wert): Anzahl}.
    # Ermittelt beim Neuanlegen, danach aus dem Blatt mitgelesen.
    altwerte: dict = field(default_factory=dict)

    def von_achse(self, achse: str) -> list[Schlagwort]:
        return [s for s in self.schlagworte.values() if s.achse == achse]

    def kanonisch(self, name: str) -> str:
        """Schreibweise aus dem Blatt „Schlagworte“ — „kraichgau“ wird
        „Kraichgau“. Unbekanntes bleibt, wie es ist; pruefe_verteiler meldet es."""
        if name in self.schlagworte:
            return name
        klein = name.lower()
        for s in self.schlagworte:
            if s.lower() == klein:
                return s
        return name


def standard_schlagworte() -> dict[str, Schlagwort]:
    return {name: Schlagwort(name, achse, _liste(oben), herkunft)
            for name, achse, oben, herkunft in STANDARD_SCHLAGWORTE}


def neuer_verteiler(kontakte: list[Kontakt]) -> Verteiler:
    """Erster Start: Schlagworte aus der Startliste, Zuständigkeiten aus den
    alten Access-Feldern — als ungeprüfte Vorschläge."""
    v = Verteiler(schlagworte=standard_schlagworte())
    for k in kontakte:
        regionen, themen, offen = startwerte(k)
        v.eintraege[k.access_id] = Eintrag(k.access_id, regionen, themen)
        for schluessel in offen:
            v.altwerte[schluessel] = v.altwerte.get(schluessel, 0) + 1
    return v


def abgleich(v: Verteiler, kontakte: list[Kontakt]) -> tuple[list[int], list[int]]:
    """Neue Pressekontakte aus Access aufnehmen, verschwundene melden.

    Gibt (neu aufgenommene IDs, IDs in der Mappe ohne Access-Satz) zurück.
    Verschwundene werden NICHT gelöscht: Vielleicht ist der Satz nur als
    verzogen markiert worden, oder jemand hat ihn in Access neu angelegt —
    dann soll die Einordnung nicht verloren sein. Sie werden nur nie
    ausgewählt, weil es zu ihnen keine Anschrift gibt.
    """
    in_access = {k.access_id for k in kontakte}
    neu = []
    for k in kontakte:
        if k.access_id not in v.eintraege:
            regionen, themen, _ = startwerte(k)
            v.eintraege[k.access_id] = Eintrag(k.access_id, regionen, themen)
            neu.append(k.access_id)
    verschwunden = sorted(i for i in v.eintraege if i not in in_access)
    return neu, verschwunden


def pruefe_verteiler(v: Verteiler) -> list[str]:
    """Tippfehler und Widersprüche in der Mappe, als Klartext.

    Wer in Excel „Kraichau“ einträgt, bekäme sonst stillschweigend keinen
    Treffer mehr — genau die Art Fehler, die das Tool abstellen soll.
    """
    meldungen = []
    for s in v.schlagworte.values():
        if s.achse not in ACHSEN:
            meldungen.append(f"Schlagwort „{s.name}“: Achse „{s.achse}“ ist "
                             f"weder „{REGION}“ noch „{THEMA}“.")
        for oben in s.gehoert_zu:
            if oben not in v.schlagworte:
                meldungen.append(f"Schlagwort „{s.name}“ gehört zu „{oben}“ — "
                                 f"das gibt es im Blatt „Schlagworte“ nicht.")
            elif v.schlagworte[oben].achse != s.achse:
                meldungen.append(f"Schlagwort „{s.name}“ ({s.achse}) gehört zu "
                                 f"„{oben}“ ({v.schlagworte[oben].achse}).")
    unbekannt: dict[str, int] = {}
    for e in v.eintraege.values():
        for name, achse in ([(r, REGION) for r in e.regionen]
                            + [(t, THEMA) for t in e.themen]):
            s = v.schlagworte.get(name)
            if s is None or s.achse != achse:
                unbekannt[f"{achse} „{name}“"] = unbekannt.get(
                    f"{achse} „{name}“", 0) + 1
    for text, n in sorted(unbekannt.items()):
        meldungen.append(f"{text} steht bei {n} Kontakt(en), ist aber kein "
                         f"bekanntes Schlagwort dieser Achse.")
    for spalte, attr, erlaubt in (("Medienart", "medienart", MEDIENARTEN),
                                  ("Ressort", "ressort", RESSORTS)):
        fremd: dict[str, int] = {}
        for e in v.eintraege.values():
            wert = getattr(e, attr)
            if wert and wert not in erlaubt:
                fremd[wert] = fremd.get(wert, 0) + 1
        for wert, n in sorted(fremd.items()):
            meldungen.append(f"{spalte} „{wert}“ steht bei {n} Kontakt(en), "
                             f"erlaubt sind: {', '.join(erlaubt)}.")
    return meldungen


# ---------------------------------------------------------------------
# Ober- und Teilgebiete
# ---------------------------------------------------------------------

def mit_obergebieten(namen, schlagworte: dict[str, Schlagwort]) -> set[str]:
    """Die Schlagworte samt allem, wozu sie gehören — über beliebig viele
    Stufen. Ein Ring in der Mappe (A gehört zu B, B zu A) darf nicht zum
    Hänger werden, daher die Liste der schon besuchten."""
    ergebnis, offen = set(), list(namen)
    while offen:
        name = offen.pop()
        if name in ergebnis:
            continue
        ergebnis.add(name)
        s = schlagworte.get(name)
        if s:
            offen.extend(s.gehoert_zu)
    return ergebnis


def mit_unterbegriffen(namen, schlagworte: dict[str, Schlagwort]) -> set[str]:
    """Die Schlagworte samt allem, was zu ihnen gehört — die Gegenrichtung
    von mit_obergebieten.

    Damit kreuzt die Oberfläche beim Buch die Teilgebiete gleich mit an: Ein
    Buch über den Landkreis Karlsruhe betrifft auch Bretten. Das geschieht
    sichtbar im Baum, so lässt sich ein Teilgebiet, das mit dem Buch nichts zu
    tun hat, für dieses eine Buch wieder abwählen.
    """
    kinder: dict[str, list[str]] = {}
    for s in schlagworte.values():
        for oben in s.gehoert_zu:
            kinder.setdefault(oben, []).append(s.name)
    ergebnis, offen = set(), list(namen)
    while offen:
        name = offen.pop()
        if name in ergebnis:
            continue
        ergebnis.add(name)
        offen.extend(kinder.get(name, []))
    return ergebnis


# Achsen, bei denen ein Häkchen am Buch die Teilgebiete mitnimmt. Nur die
# Regionen: Ein Buch über den Landkreis Karlsruhe handelt auch vom
# Kraichgau-Teil darin — eine Fläche enthält ihre Teile. Bei den Themen
# stimmt das nicht: Ein Geschichtsbuch ist nicht deshalb Kirchengeschichte,
# und eine Ortsgeschichte kein Ortssippenbuch (mit Mitnahme ging das
# Östringen-Buch im Probelauf an acht Genealogie-Redaktionen).
MIT_TEILGEBIETEN = {REGION}


def spitzen(namen, schlagworte: dict[str, Schlagwort]) -> set[str]:
    """Die obersten Häkchen: angekreuzt, ohne dass ihr Oberbegriff es ist.

    Die Oberfläche kreuzt bei einer Region die Teilgebiete mit an. Gemeint hat
    man aber nur die Spitze: Wer „Landkreis Karlsruhe“ anklickt, meint nicht
    den Kraichgau samt Heilbronner Teil. Deshalb vergleicht der Vorschlag nur
    die Spitzen, sonst sähe das Buch dem Landkreis Heilbronn ähnlich.
    """
    namen = set(namen)
    return {n for n in namen
            if not (n in schlagworte
                    and set(schlagworte[n].gehoert_zu) & namen)}


# ---------------------------------------------------------------------
# Historie: wer welches Buch bekommen hat
# ---------------------------------------------------------------------
# Rund 350 Presselisten aus den Jahren 2020–2026 zeigen, wie die Auswahl
# wirklich läuft. Sie lässt sich nicht aus einer Zuständigkeit je Kontakt
# ableiten: Die Rheinpfalz-Kultur bekommt 81 % der Bücher aus Rhein-Neckar und
# 69 % aus dem Raum Karlsruhe, obwohl sie dort nicht erscheint. Die alten
# Listen selbst sind aber eine brauchbare Vorlage. 98 % der Empfänger eines
# neuen Buchs standen schon in früheren Listen.

BLATT_BUECHER = "Bücher"
BLATT_EMPFAENGER = "Empfänger"
BUCH_SPALTEN = ["Nr", "Datum", "Titel", "ISBN", "Gattung", "Regionen",
                "Themen", "Empfänger", "Herkunft"]
EMPFAENGER_SPALTEN = ["Nr", "Access-ID", "Institution", "Ort"]

HISTORIE_HINWEISE = [
    "Pressehistorie — wer welches Buch bekommen hat",
    "",
    "Aus dieser Mappe schlägt das Programm „Presseverteiler“ die Empfänger "
    "eines neuen Buchs vor. Jede gespeicherte Liste kommt automatisch dazu.",
    "",
    "Blatt „Bücher“: je Buch eine Zeile. Regionen und Themen stehen wie im "
    "Programm, mehrere mit Semikolon getrennt.",
    "Blatt „Empfänger“: je Buch und Kontakt eine Zeile. Nr verweist auf das "
    "Blatt „Bücher“. Maßgeblich ist die Access-ID; Institution und Ort stehen "
    "nur zur Ansicht da.",
    "",
    "Von Hand ändern geht, solange das Programm geschlossen ist. Eine Liste, "
    "die als Vorbild nichts taugt, lässt sich löschen: die Zeile im Blatt "
    "„Bücher“ entfernen, die Empfänger dazu dürfen stehen bleiben.",
]


@dataclass
class Buch:
    """Ein Buch aus der Historie samt den Empfängern seiner Presseliste."""
    nummer: int
    datum: date
    titel: str
    isbn: str = ""
    # Nur zur Ansicht. Beim Vergleich zählt sie nicht: Themen und Gattung sagen
    # dasselbe, ohne eines von beiden war der Vorschlag im Test genauso gut.
    gattung: str = ""
    regionen: set = field(default_factory=set)
    themen: set = field(default_factory=set)
    empfaenger: set = field(default_factory=set)
    herkunft: str = ""
    # {Access-ID: (Institution, Ort)} — Abschrift zur Ansicht. Sie hält auch
    # Kontakte fest, die es in Access nicht mehr gibt.
    anzeige: dict = field(default_factory=dict)


def _als_datum(wert) -> date | None:
    if isinstance(wert, datetime):
        return wert.date()
    if isinstance(wert, date):
        return wert
    try:
        return datetime.strptime(_text(wert), "%Y-%m-%d").date()
    except ValueError:
        return None


def lade_historie(pfad: Path = None) -> list[Buch]:
    """Alle Bücher, älteste zuerst. Ohne Mappe eine leere Liste: Dann gibt es
    keine Vorschläge aus früheren Büchern, das Programm läuft aber."""
    pfad = Path(pfad or HISTORIE_PFAD)
    if not pfad.is_file():
        return []
    wb = _mappe_lesen(pfad)
    try:
        buch_zeilen = _blatt_zeilen(wb, BLATT_BUECHER)
        empf_zeilen = _blatt_zeilen(wb, BLATT_EMPFAENGER)
    finally:
        wb.close()

    buecher: dict[int, Buch] = {}
    for z in _blatt_als_dicts(buch_zeilen, {"Nr", "Datum", "Titel"}):
        nummer, datum = _ganzzahl(z.get("Nr")), _als_datum(z.get("Datum"))
        if nummer is None or datum is None:
            continue
        buecher[nummer] = Buch(
            nummer, datum, _text(z.get("Titel")), _text(z.get("ISBN")),
            _text(z.get("Gattung")), set(_liste(z.get("Regionen"))),
            set(_liste(z.get("Themen"))), herkunft=_text(z.get("Herkunft")))
    for z in _blatt_als_dicts(empf_zeilen, {"Nr", "Access-ID"}):
        buch = buecher.get(_ganzzahl(z.get("Nr")))
        access_id = _ganzzahl(z.get("Access-ID"))
        if buch is None or access_id is None:
            continue
        buch.empfaenger.add(access_id)
        buch.anzeige[access_id] = (_text(z.get("Institution")),
                                   _text(z.get("Ort")))
    return sorted(buecher.values(), key=lambda b: (b.datum, b.nummer))


def speichere_historie(buecher: list[Buch], kontakte: list[Kontakt],
                       pfad: Path = None) -> Path:
    """Schreibt die ganze Mappe neu, wie speichere_verteiler über eine
    Nebendatei, damit ein Abbruch die alte Mappe nicht zerstört."""
    pfad = Path(pfad or HISTORIE_PFAD)
    sperre = pruefe_sperre(pfad)
    if sperre:
        raise PermissionError(sperre)
    nach_id = {k.access_id: k for k in kontakte}
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = BLATT_HINWEISE
    for zeile in HISTORIE_HINWEISE:
        ws.append([zeile])
    ws["A1"].font = Font(bold=True, size=14)
    ws.column_dimensions["A"].width = 120
    for zelle in ws["A"]:
        zelle.alignment = Alignment(wrap_text=True, vertical="top")

    ws = wb.create_sheet(BLATT_BUECHER)
    _kopfzeile(ws, BUCH_SPALTEN, [7, 12, 60, 20, 18, 40, 40, 10, 40])
    for b in buecher:
        ws.append([b.nummer, b.datum.isoformat(), b.titel, b.isbn, b.gattung,
                   "; ".join(sorted(b.regionen)), "; ".join(sorted(b.themen)),
                   len(b.empfaenger), b.herkunft])
    ws.auto_filter.ref = ws.dimensions

    ws = wb.create_sheet(BLATT_EMPFAENGER)
    _kopfzeile(ws, EMPFAENGER_SPALTEN, [7, 10, 40, 24])
    for b in buecher:
        for i in sorted(b.empfaenger):
            k = nach_id.get(i)
            institution, ort = ((k.institution, k.ort) if k
                                else b.anzeige.get(i, ("", "")))
            ws.append([b.nummer, i, institution, ort])
    ws.auto_filter.ref = ws.dimensions

    pfad.parent.mkdir(parents=True, exist_ok=True)
    _mappe_ablegen(wb, pfad)
    return pfad


def buch_merken(buecher: list[Buch], titel: str, isbn: str, regionen, themen,
                empfaenger, heute: date = None) -> Buch:
    """Nimmt eine gespeicherte Liste in die Historie auf.

    Wer dieselbe Liste am selben Tag noch einmal speichert, hat sie korrigiert.
    Sie ersetzt dann den alten Stand, sonst zählte das Buch doppelt.
    """
    heute = heute or date.today()
    for b in buecher:
        if b.datum == heute and (b.titel, b.isbn) == (titel, isbn):
            b.regionen, b.themen = set(regionen), set(themen)
            b.empfaenger = set(empfaenger)
            return b
    nummer = max((b.nummer for b in buecher), default=0) + 1
    b = Buch(nummer, heute, titel, isbn, "", set(regionen), set(themen),
             set(empfaenger), herkunft="Presseverteiler")
    buecher.append(b)
    return b


# ---------------------------------------------------------------------
# Der Vorschlag
# ---------------------------------------------------------------------
# Das Buch wird mit allen früheren Büchern verglichen. Die ähnlichsten stimmen
# ab: Wer in vielen ihrer Listen stand, steht weit oben. Damit kommen die
# Kulturredaktionen von selbst nach oben, die bei fast jedem Buch dabei sind,
# und ebenso die Fachblätter eines Themas. Eine Regel je Kontakt braucht es
# dafür nicht.
#
# Die Werte sind an den Listen bis Ende 2024 eingestellt und an denen ab 2025
# blind geprüft (siehe blindtest.py). Wer sie ändert, muss das wiederholen.
NACHBARN = 15            # so viele ähnliche Bücher stimmen ab
SCHWELLE = 0.30          # ab diesem Stimmenanteil ist ein Kontakt vorangekreuzt
ALTERUNG_JAHRE = 2       # nach so vielen Jahren zählt eine Liste nur noch 1/e
# Kontakte, die laut Einschätzung in der Gegend des Buchs sitzen, rücken in der
# Liste etwas nach oben. Vorangekreuzt werden sie dadurch nicht: Das hat im
# Test nur 1–2 Punkte gebracht, hilft aber neuen Kontakten ohne Vorgeschichte.
ZUSCHLAG_GEGEND = 0.15


@dataclass
class Treffer:
    kontakt: Kontakt
    stimmen: float            # Anteil der gewichteten Stimmen, 0 bis 1
    buecher: int              # in so vielen der ähnlichen Bücher dabei
    vorgeschlagen: bool       # vorangekreuzt
    warum: str


@dataclass
class Vorschlag:
    treffer: list[Treffer] = field(default_factory=list)
    nachbarn: list[tuple[float, Buch]] = field(default_factory=list)


def _merkmale(regionen, themen,
              schlagworte: dict[str, Schlagwort]) -> dict[str, float]:
    """Das Buch als gewichtete Merkmale. Was genau angekreuzt ist, zählt mehr
    als die Obergebiete, die daraus folgen."""
    oben = spitzen(regionen, schlagworte)
    merkmale = {f"R:{n}": 1.0 for n in mit_obergebieten(oben, schlagworte)}
    merkmale.update({f"R:{n}": 2.0 for n in oben})
    merkmale.update({f"T:{n}": 1.0
                     for n in mit_obergebieten(themen, schlagworte)})
    merkmale.update({f"T:{n}": 1.5 for n in themen})
    return merkmale


def _kosinus(a: dict[str, float], b: dict[str, float]) -> float:
    skalar = sum(wert * b[m] for m, wert in a.items() if m in b)
    laenge = math.sqrt(sum(x * x for x in a.values())
                       * sum(x * x for x in b.values()))
    return skalar / laenge if laenge else 0.0


def aehnliche_buecher(regionen, themen, historie: list[Buch],
                      schlagworte: dict[str, Schlagwort],
                      stichtag: date = None, anzahl: int = None,
                      alterung: float = None) -> list[tuple[float, Buch]]:
    """Die ähnlichsten Bücher vor dem Stichtag, als (Gewicht, Buch).

    Seltene Merkmale zählen mehr: Dass zwei Bücher in Baden-Württemberg
    spielen, sagt wenig, dass beide vom Zabergäu handeln, sagt viel. Neuere
    Listen wiegen mehr, weil die Listen über die Jahre kleiner geworden sind
    und Kontakte wechseln.

    Nur Bücher *vor* dem Stichtag: Der Blindtest stellt damit sicher, dass der
    Vorschlag nichts aus der Zukunft kennt. Im Betrieb ist der Stichtag heute,
    und eine am selben Tag gespeicherte Liste bleibt bis morgen außen vor.
    """
    stichtag = stichtag or date.today()
    anzahl = NACHBARN if anzahl is None else anzahl
    alterung = ALTERUNG_JAHRE if alterung is None else alterung
    pool = [b for b in historie if b.datum < stichtag and b.empfaenger]
    if not pool or not (regionen or themen):
        return []
    roh = [(b, _merkmale(b.regionen, b.themen, schlagworte)) for b in pool]
    haeufig: dict[str, int] = {}
    for _, m in roh:
        for name in m:
            haeufig[name] = haeufig.get(name, 0) + 1

    def gewichtet(m: dict[str, float]) -> dict[str, float]:
        return {name: wert * math.log(1 + len(pool) / (1 + haeufig.get(name, 0)))
                for name, wert in m.items()}

    ziel = gewichtet(_merkmale(regionen, themen, schlagworte))
    bewertet = []
    for b, m in roh:
        alter = (stichtag - b.datum).days / 365
        gewicht = _kosinus(ziel, gewichtet(m)) * math.exp(-alter / alterung)
        if gewicht > 0:
            bewertet.append((gewicht, b))
    bewertet.sort(key=lambda x: (-x[0], -x[1].datum.toordinal(), x[1].nummer))
    return bewertet[:anzahl]


def vorschlag(regionen, themen, historie: list[Buch], v: Verteiler,
              kontakte: list[Kontakt], stichtag: date = None,
              anzahl: int = None, schwelle: float = None,
              alterung: float = None) -> Vorschlag:
    """Die Empfänger für ein Buch, die wahrscheinlichsten zuerst.

    `regionen` sind die im Baum angekreuzten, samt mitgenommenen Teilgebieten.
    Die Liste enthält auch Kontakte unter der Schwelle: Die Hälfte seiner
    Empfänger steht erst dort, und genau da soll der Pressereferent nachsehen.
    """
    schwelle = SCHWELLE if schwelle is None else schwelle
    nachbarn = aehnliche_buecher(regionen, themen, historie, v.schlagworte,
                                 stichtag, anzahl, alterung)
    gesamt = sum(g for g, _ in nachbarn) or 1.0
    stimmen: dict[int, float] = {}
    zahl: dict[int, int] = {}
    for g, b in nachbarn:
        for i in b.empfaenger:
            stimmen[i] = stimmen.get(i, 0.0) + g / gesamt
            zahl[i] = zahl.get(i, 0) + 1

    gegend = mit_unterbegriffen(regionen, v.schlagworte)
    wertung = []
    for k in kontakte:
        e = v.eintraege.get(k.access_id)
        s = stimmen.get(k.access_id, 0.0)
        hier = bool(e and e.regionen & gegend)
        immer = bool(e and e.immer)
        if not (s or hier or immer):
            continue
        if immer:
            warum = "immer"
        elif s:
            warum = (f"in {zahl[k.access_id]} von {len(nachbarn)} "
                     f"ähnlichen Büchern")
        else:
            warum = "sitzt in der Gegend, bisher nicht bei ähnlichen Büchern"
        wertung.append((s + (ZUSCHLAG_GEGEND if hier else 0.0), Treffer(
            k, s, zahl.get(k.access_id, 0), immer or s >= schwelle, warum)))
    wertung.sort(key=lambda x: (-x[1].vorgeschlagen, -x[0],
                                x[1].kontakt.sortierung))
    return Vorschlag([t for _, t in wertung], nachbarn)


def einzelne_mails(kontakte: list[Kontakt]) -> list[tuple[Kontakt, str]]:
    """Jede Adresse einmal, mit dem Kontakt, bei dem sie zuerst steht."""
    gesehen, paare = set(), []
    for k in kontakte:
        for adresse in k.mails:
            if adresse.lower() not in gesehen:
                gesehen.add(adresse.lower())
                paare.append((k, adresse))
    return paare


def mailadressen(kontakte: list[Kontakt]) -> tuple[str, int, int]:
    """Die Adressen für das BCC-Feld, mit Semikolon getrennt.

    Gibt (Text, Zahl der Kontakte ohne E-Mail, Zahl doppelter Adressen)
    zurück. Doppelte kommen vor, weil Sammeladressen („redaktion@…“) bei
    mehreren Personen stehen — die Redaktion soll die Meldung trotzdem nur
    einmal bekommen. Stehen mehrere Adressen bei einem Kontakt, gehen alle
    mit: lieber Redakteur und Redaktion als keiner von beiden.
    """
    paare = einzelne_mails(kontakte)
    ohne = sum(1 for k in kontakte if not k.mails)
    doppelt = sum(len(k.mails) for k in kontakte) - len(paare)
    return "; ".join(adresse for _, adresse in paare), ohne, doppelt


# ---------------------------------------------------------------------
# Mappe lesen und schreiben
# ---------------------------------------------------------------------

BLATT_HINWEISE = "Hinweise"
BLATT_PRESSE = "Presse"
BLATT_SCHLAGWORTE = "Schlagworte"
BLATT_ALTWERTE = "Altwerte aus Access"

# Reihenfolge der Spalten im Blatt „Presse“. Gelesen wird nach Namen, nicht
# nach Position — wer in Excel Spalten verschiebt, macht nichts kaputt.
PRESSE_SPALTEN = ["Access-ID", "Institution", "Abteilung", "Person", "Ort",
                  "Medium", "Regionen", "Themen", "immer", "geprüft",
                  "unsicher", "Medienart", "Ressort", "Begründung",
                  "Bemerkung"]
PRESSE_PFLICHT = {"Access-ID", "Regionen", "Themen"}
ANZEIGE_SPALTEN = ("Institution", "Abteilung", "Person", "Ort", "Medium")
SCHLAGWORT_SPALTEN = ["Schlagwort", "Achse", "gehört zu", "Herkunft"]

HINWEISE = [
    "Presseverteiler — wer welches Buch bekommt",
    "",
    "Diese Mappe gehört zum Programm „Presseverteiler“. Sie hält fest, wofür "
    "jeder Pressekontakt zuständig ist. Die Anschriften stehen NICHT hier, "
    "sondern in Access; das Programm liest sie bei jedem Start aus dem "
    "Access-Export.",
    "",
    "Blatt „Presse“: je Pressekontakt eine Zeile.",
    "  Regionen / Themen: Schlagworte aus dem Blatt „Schlagworte“, mehrere mit "
    "Semikolon getrennt („Kraichgau; Landkreis Karlsruhe“).",
    "  immer: x = ist bei jedem Buch vorangekreuzt.",
    "  geprüft: x = die Einordnung ist durchgesehen. Leer = Vorschlag, noch "
    "nicht bestätigt.",
    "  unsicher: x = der Vorschlag ist geraten; diese zuerst durchsehen.",
    "  Medienart / Ressort: Art des Mediums und Zuständigkeit in der "
    "Redaktion, aus festen Listen. Gehen noch nicht in die Auswahl ein.",
    "  Begründung: woraus der Vorschlag folgt („Lokalredaktion in Bretten“).",
    "  Bemerkung: eigene Notizen.",
    "  Institution bis Medium: nur zur Ansicht, wird beim Speichern aus "
    "Access erneuert.",
    "",
    "Wer ein Buch bekommt, schlägt das Programm aus den Listen ähnlicher "
    "früherer Bücher vor (Mappe „Pressehistorie.xlsx“). Regionen und Themen "
    "je Kontakt dienen nur noch als Suchhilfe und helfen neuen Kontakten ohne "
    "Vorgeschichte.",
    "",
    "Blatt „Schlagworte“: „gehört zu“ macht die Obergebiete. Ein Buch über "
    "den Pfinzgau gilt damit auch als ähnlich zu Büchern über den Landkreis "
    "Karlsruhe, die Region Karlsruhe und Baden.",
    "  Herkunft „Shop“: Kategorie aus dem Webshop. „Ergänzung“: fehlt im Shop, "
    "für die Presse aber nötig.",
    "",
    "Von Hand ändern geht — aber nur, solange das Programm geschlossen ist. "
    "Blätter und Spaltenüberschriften nicht umbenennen.",
]

_KOPF_FONT = Font(bold=True, color="FFFFFF")
_KOPF_FILL = PatternFill("solid", fgColor="44546A")


def pruefe_sperre(pfad: Path = None) -> str:
    """Prüft, ob die Mappe gerade in Excel offen ist.

    Excel legt neben einer geöffneten Datei eine versteckte Sperrdatei
    ``~$<Name>.xlsx`` an und verweigert fremde Schreibzugriffe. Ohne diese
    Prüfung stürbe das Speichern mit einem PermissionError. Gibt eine Meldung
    im Klartext zurück oder "" wenn alles frei ist.
    """
    pfad = Path(pfad or VERTEILER_PFAD)
    if (pfad.parent / f"~${pfad.name}").exists():
        return (f"Die Mappe „{pfad.name}“ ist gerade in Excel geöffnet.\n"
                f"Bitte dort schließen und erneut versuchen.")
    return ""


def _blatt_zeilen(wb, name: str) -> list[tuple]:
    if name not in wb.sheetnames:
        raise ValueError(f"In der Mappe fehlt das Blatt „{name}“.\n"
                         f"Vorhanden sind: {', '.join(wb.sheetnames)}.")
    return [tuple(z) for z in wb[name].iter_rows(values_only=True)]


def lade_verteiler(pfad: Path = None) -> Verteiler:
    pfad = Path(pfad or VERTEILER_PFAD)
    wb = _mappe_lesen(pfad)
    try:
        schlagwort_zeilen = _blatt_zeilen(wb, BLATT_SCHLAGWORTE)
        presse_zeilen = _blatt_zeilen(wb, BLATT_PRESSE)
        # Das Blatt wird nur beim Neuanlegen ermittelt. Ohne Mitlesen ginge
        # es beim ersten Speichern verloren.
        alt_zeilen = (_blatt_zeilen(wb, BLATT_ALTWERTE)
                      if BLATT_ALTWERTE in wb.sheetnames else [])
    finally:
        wb.close()

    v = Verteiler()
    if alt_zeilen:
        for z in _blatt_als_dicts(alt_zeilen, {"Feld", "Wert in Access"}):
            feld_, wert = _text(z.get("Feld")), _text(z.get("Wert in Access"))
            if feld_ and wert:
                v.altwerte[(feld_, wert)] = _ganzzahl(z.get("Anzahl")) or 0
    for z in _blatt_als_dicts(schlagwort_zeilen, {"Schlagwort", "Achse"}):
        name = _text(z.get("Schlagwort"))
        if name:
            v.schlagworte[name] = Schlagwort(
                name, _text(z.get("Achse")).capitalize(),
                _liste(z.get("gehört zu")), _text(z.get("Herkunft")))
    # „gehört zu“ in der kanonischen Schreibweise — sonst reißt ein
    # kleingeschriebenes „baden“ die Kette ab.
    for s in v.schlagworte.values():
        s.gehoert_zu = [v.kanonisch(o) for o in s.gehoert_zu]

    for z in _blatt_als_dicts(presse_zeilen, PRESSE_PFLICHT):
        access_id = _ganzzahl(z.get("Access-ID"))
        if access_id is None:
            continue
        v.eintraege[access_id] = Eintrag(
            access_id,
            {v.kanonisch(r) for r in _liste(z.get("Regionen"))},
            {v.kanonisch(t) for t in _liste(z.get("Themen"))},
            _haken(z.get("immer")), _haken(z.get("geprüft")),
            _text(z.get("Bemerkung")),
            {s: _text(z.get(s)) for s in ANZEIGE_SPALTEN},
            _text(z.get("Begründung")), _haken(z.get("unsicher")),
            _text(z.get("Medienart")), _text(z.get("Ressort")))
    return v


def _kopfzeile(ws, spalten: list[str], breiten: list[int]):
    ws.append(spalten)
    for i, breite in enumerate(breiten, start=1):
        zelle = ws.cell(row=1, column=i)
        zelle.font, zelle.fill = _KOPF_FONT, _KOPF_FILL
        ws.column_dimensions[get_column_letter(i)].width = breite
    ws.freeze_panes = "A2"


def als_text(namen: set, v: Verteiler) -> str:
    """Schlagworte in der Reihenfolge des Blatts „Schlagworte“, nicht
    alphabetisch — dort stehen Großräume vor Kreisen vor Landschaften."""
    reihenfolge = {n: i for i, n in enumerate(v.schlagworte)}
    return "; ".join(sorted(namen, key=lambda n: (reihenfolge.get(n, 10**6), n)))


def speichere_verteiler(v: Verteiler, kontakte: list[Kontakt],
                        pfad: Path = None) -> Path:
    """Schreibt die ganze Mappe neu.

    Erst in eine Nebendatei, dann umbenennen: Bricht das Schreiben mittendrin
    ab (Netzlaufwerk, Strom), bleibt die alte Mappe heil.
    """
    pfad = Path(pfad or VERTEILER_PFAD)
    sperre = pruefe_sperre(pfad)
    if sperre:
        raise PermissionError(sperre)

    nach_id = {k.access_id: k for k in kontakte}
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = BLATT_HINWEISE
    for zeile in HINWEISE:
        ws.append([zeile])
    ws["A1"].font = Font(bold=True, size=14)
    ws.column_dimensions["A"].width = 120
    for zelle in ws["A"]:
        zelle.alignment = Alignment(wrap_text=True, vertical="top")

    ws = wb.create_sheet(BLATT_PRESSE)
    _kopfzeile(ws, PRESSE_SPALTEN,
               [10, 34, 24, 24, 22, 14, 44, 40, 7, 8, 8, 16, 18, 50, 30])

    def zeilenfolge(e: Eintrag):
        k = nach_id.get(e.access_id)
        if k:
            return k.sortierung
        return ((e.anzeige.get("Institution") or "~").lower(), "", "")

    for e in sorted(v.eintraege.values(), key=zeilenfolge):
        k = nach_id.get(e.access_id)
        if k:
            anzeige = {"Institution": k.institution, "Abteilung": k.abteilung,
                       "Person": k.person, "Ort": k.ort, "Medium": k.medium}
        else:
            anzeige = e.anzeige
        ws.append([e.access_id, *(anzeige.get(s, "") for s in ANZEIGE_SPALTEN),
                   als_text(e.regionen, v), als_text(e.themen, v),
                   "x" if e.immer else "", "x" if e.geprueft else "",
                   "x" if e.unsicher else "", e.medienart, e.ressort,
                   e.begruendung, e.bemerkung])
    ws.auto_filter.ref = ws.dimensions

    ws = wb.create_sheet(BLATT_SCHLAGWORTE)
    _kopfzeile(ws, SCHLAGWORT_SPALTEN, [32, 10, 60, 12])
    for s in v.schlagworte.values():
        ws.append([s.name, s.achse, "; ".join(s.gehoert_zu), s.herkunft])
    ws.auto_filter.ref = ws.dimensions

    if v.altwerte:
        ws = wb.create_sheet(BLATT_ALTWERTE)
        _kopfzeile(ws, ["Feld", "Wert in Access", "Anzahl"], [14, 60, 8])
        for (feld_, wert), n in sorted(v.altwerte.items(),
                                       key=lambda x: (x[0][0], -x[1])):
            ws.append([feld_, wert, n])

    pfad.parent.mkdir(parents=True, exist_ok=True)
    _mappe_ablegen(wb, pfad)
    return pfad


# ---------------------------------------------------------------------
# Versand je Buch: ein Ordner mit je einer Liste pro Versandweg
# ---------------------------------------------------------------------
# Ein Ordner je Buch, damit der Pressereferent die Datei für seinen Versandweg
# sofort findet und nicht erst das richtige Blatt suchen muss. „Empfänger.xlsx“
# ist zugleich der Nachweis, wer das Buch bekommen hat. Die Ordner liegen neben
# denen, die er seit 2020 von Hand anlegt, und heißen darum wie diese nach dem
# Buch, ohne Datum und ISBN.
#
# Die Post-Liste enthält alle mit Anschrift, nicht nur die ohne E-Mail: Die
# wären im Median 0 je Liste, höchstens 4 (Listen ab 2025). Damit sie sich
# trotzdem herausfiltern lassen, gibt es die Spalte „auch per E-Mail“.

DATEI_EMPFAENGER = "Empfänger.xlsx"
DATEI_MAIL = "E-Mail.xlsx"
DATEI_POST = "Post.xlsx"
DATEI_BCC = "E-Mail-Adressen.txt"

VERSAND_SPALTEN = ["Access-ID", "Anrede", "Titel", "Vorname", "Name",
                   "Institution", "Abteilung", "Unterabteilung", "c/o",
                   "Straße", "Hausnummer", "PLZ", "Ort", "Land", "eMail",
                   "Medium", "Grund"]
MAIL_SPALTEN = ["Institution", "Abteilung", "Person", "eMail"]
POST_SPALTEN = ["Access-ID", "Anrede", "Titel", "Vorname", "Name",
                "Institution", "Abteilung", "Unterabteilung", "c/o",
                "Straße", "Hausnummer", "PLZ", "Ort", "Land",
                "auch per E-Mail"]


def _titel_person(k: Kontakt) -> str:
    return " ".join(t for t in (k.feld("Titel 1"), k.feld("Titel 2"),
                                k.feld("Titel 3")) if t)


def _mappe_empfaenger(titel: str, isbn: str, regionen, themen,
                      empfaenger: list[tuple[Kontakt, str]],
                      abgewaehlt: list[tuple[Kontakt, str]]):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Empfänger"
    _kopfzeile(ws, VERSAND_SPALTEN,
               [10, 10, 10, 16, 20, 34, 24, 18, 14, 24, 8, 8, 18, 8, 32, 14, 40])
    for k, g in empfaenger:
        ws.append([k.access_id, k.feld("Anrede"), _titel_person(k),
                   k.feld("Vorname"), k.feld("Name"), k.institution,
                   k.feld("Abteilung"), k.feld("Unterabteilung"),
                   k.feld("c/o"), k.feld("Straße"), k.feld("Hausnummer"),
                   k.feld("PLZ"), k.feld("Ort"), k.feld("Land"), k.mail,
                   k.medium, g])
    ws.auto_filter.ref = ws.dimensions

    ws = wb.create_sheet("Buch")
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 90
    for schluessel, wert in (
            ("Titel", titel), ("ISBN", isbn),
            ("Datum", date.today().strftime("%d.%m.%Y")),
            ("Regionen", "; ".join(sorted(regionen))),
            ("Themen", "; ".join(sorted(themen))),
            ("Empfänger", len(empfaenger)),
            ("davon ohne E-Mail", sum(1 for k, _ in empfaenger if not k.mails)),
            ("davon ohne Anschrift",
             sum(1 for k, _ in empfaenger if not k.hat_anschrift))):
        ws.append([schluessel, wert])
        ws.cell(row=ws.max_row, column=1).font = Font(bold=True)

    if abgewaehlt:
        ws = wb.create_sheet("Abgewählt")
        _kopfzeile(ws, ["Access-ID", "Institution", "Person", "wäre gewesen"],
                   [10, 34, 24, 50])
        for k, g in abgewaehlt:
            ws.append([k.access_id, k.institution, k.person, g])
    return wb


def _mappe_mail(kontakte: list[Kontakt]):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "E-Mail"
    _kopfzeile(ws, MAIL_SPALTEN, [34, 24, 24, 40])
    for k, adresse in einzelne_mails(kontakte):
        ws.append([k.institution, k.abteilung, k.person, adresse])
    ws.auto_filter.ref = ws.dimensions
    return wb


def _mappe_post(kontakte: list[Kontakt]):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Post"
    _kopfzeile(ws, POST_SPALTEN,
               [10, 10, 10, 16, 20, 34, 24, 18, 14, 24, 8, 8, 18, 8, 14])
    for k in kontakte:
        if k.hat_anschrift:
            ws.append([k.access_id, k.feld("Anrede"), _titel_person(k),
                       k.feld("Vorname"), k.feld("Name"), k.institution,
                       k.feld("Abteilung"), k.feld("Unterabteilung"),
                       k.feld("c/o"), k.feld("Straße"), k.feld("Hausnummer"),
                       k.feld("PLZ"), k.feld("Ort"), k.feld("Land"),
                       "ja" if k.mails else "nein"])
    ws.auto_filter.ref = ws.dimensions
    return wb


def _buch_ordner(wurzel: Path, titel: str, isbn: str) -> Path:
    """Der Ordner des Buchs.

    Am selben Tag gilt erneutes Speichern als Korrektur und ersetzt die Listen.
    Stammen sie von einem anderen Tag, ist es vermutlich das gleichnamige
    Jahrbuch des Vorjahrs. Dessen Liste darf nicht verloren gehen, das Buch
    bekommt dann einen eigenen Ordner mit Datum.
    """
    heute = date.today()
    name = _sicherer_dateiname(titel[:60] or isbn) or heute.isoformat()
    alt = wurzel / name / DATEI_EMPFAENGER
    if alt.exists() and date.fromtimestamp(alt.stat().st_mtime) != heute:
        return wurzel / f"{name} {heute.isoformat()}"
    return wurzel / name


def speichere_versand(titel: str, isbn: str, regionen, themen,
                      empfaenger: list[tuple[Kontakt, str]],
                      abgewaehlt: list[tuple[Kontakt, str]],
                      ordner: Path = None) -> Path:
    """Schreibt den Ordner des Buchs und gibt ihn zurück.

    Alle Dateien entstehen jedes Mal neu. So bleibt beim erneuten Speichern
    keine Liste vom alten Stand liegen, auch wenn sie nun leer wäre.
    """
    ziel = _buch_ordner(Path(ordner or AUSGABE_ORDNER), titel, isbn)

    # Alphabetisch: Die Listen liest ein Mensch, die Reihenfolge des Vorschlags
    # sagt dort nichts mehr.
    empfaenger = sorted(empfaenger, key=lambda e: e[0].sortierung)
    kontakte = [k for k, _ in empfaenger]
    mappen = {
        DATEI_EMPFAENGER: _mappe_empfaenger(titel, isbn, regionen, themen,
                                            empfaenger, abgewaehlt),
        DATEI_MAIL: _mappe_mail(kontakte),
        DATEI_POST: _mappe_post(kontakte)}

    # Erst alle prüfen, dann schreiben: Sonst stünde bei einer offenen Datei
    # ein Teil des Ordners auf dem neuen Stand und ein Teil auf dem alten.
    for datei in mappen:
        sperre = pruefe_sperre(ziel / datei)
        if sperre:
            raise OSError(sperre)

    # Erst örtlich schreiben, dann am Stück kopieren: Eine xlsx ist ein Zip,
    # beim Schreiben springt openpyxl in der Datei herum, und das verträgt die
    # gvfs-Einbindung des NAS schlecht. copyfile statt copy2, weil gvfs das
    # Setzen der Dateirechte ablehnen kann.
    with tempfile.TemporaryDirectory() as tmp:
        for datei, wb in mappen.items():
            wb.save(Path(tmp) / datei)
        (Path(tmp) / DATEI_BCC).write_text(mailadressen(kontakte)[0],
                                           encoding="utf-8")
        ziel.mkdir(parents=True, exist_ok=True)
        for pfad in Path(tmp).iterdir():
            shutil.copyfile(pfad, ziel / pfad.name)
    return ziel
