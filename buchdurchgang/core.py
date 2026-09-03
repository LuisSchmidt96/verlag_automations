"""Buchdurchgang — Cover, pi&bi und Shopware in einem Zug.

Reine Logik, keine Oberfläche.

Dieses Werkzeug ist die Ausnahme von der Hausregel „kein geteilter Code":
es **benutzt** die drei anderen Werkzeuge, statt sie nachzubauen. Die Regel
verbietet, eine gemeinsame Bibliothek herauszulösen — hier wird nichts
herausgelöst, die drei bleiben einzeln lauffähig. Der Preis ist Kopplung:
ändert sich `cover_previews/core.py`, ändert sich auch dieses Werkzeug. Das
ist bewusst so entschieden (02.09.2026); die Alternative wäre gewesen, den
Photoshop-Aufruf, die docx-Vorlagen und den Shopware-Client zu duplizieren.

**Die Konfiguration ist eigenständig.** Die drei Werkzeuge legen ihre
`config.json` je neben ihre eigene .exe — in einer gemeinsamen .exe würden
sich alle drei auf dieselbe Datei stürzen. Deshalb hält der Durchgang seine
eigene Konfiguration mit einem Abschnitt je Werkzeug und reicht die fertigen
Dicts hinein; `lade_config()` der anderen wird nie aufgerufen.
"""

from __future__ import annotations

import json
import os
import sys
from datetime import datetime
from pathlib import Path

from cover_previews import core as cp
from pi_bi_generator import core as pb
from shopware_publisher import core as sw


# ---------------------------------------------------------------------
# Ablage
# ---------------------------------------------------------------------

def _base_dir() -> Path:
    """Ordner neben der .exe (PyInstaller-Build) bzw. neben dem Tool-Code."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


APP_DIR = _base_dir()
APP_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_PFAD = APP_DIR / "config.json"


DEFAULT_CONFIG: dict = {
    # Basis, unter der je Buch ein Ordner "<Kurzcode>_<Titel>" entsteht.
    # Zum Testen der Netzordner, im Echtbetrieb die Artikeldaten.
    "ablageort": r"\\C019\d\Online\Webseite\Artikeldaten",
    # Abweichungen von den Vorgaben der drei Werkzeuge. Leer = deren Vorgabe.
    "cover_previews": {},
    "pi_bi_generator": {},
    # Der Shopware-Abschnitt ist eine vollständige Publisher-Konfiguration
    # (mit "umgebungen"), damit Zugangsdaten und Verschlüsselung unverändert
    # funktionieren.
    "shopware_publisher": {},
    # Wie weit dürfen gemessenes Cover-Format und ONIX-Maße auseinanderliegen,
    # bevor gewarnt wird? Der Beschnitt macht das gemessene Maß größer
    # (3 mm ringsum = 0,6 cm je Kante); 1,0 cm lässt auch 5 mm Beschnitt durch.
    "format_tol_cm": 1.0,
    "config_version": 1,
}


def lade_config() -> dict:
    cfg = dict(DEFAULT_CONFIG)
    if CONFIG_PFAD.exists():
        try:
            cfg.update(json.loads(CONFIG_PFAD.read_text(encoding="utf-8")))
        except (OSError, json.JSONDecodeError):
            pass
    for k, v in DEFAULT_CONFIG.items():
        cfg.setdefault(k, v)

    # Der Shopware-Abschnitt wird beim ersten Start vollständig angelegt: dort
    # liegen Zugangsdaten und das verschlüsselte Secret, und die müssen beim
    # Speichern erhalten bleiben. Deshalb ist er kein Overlay, sondern eine
    # echte Publisher-Konfiguration, die unmittelbar bearbeitet wird.
    shop = cfg.get("shopware_publisher") or {}
    if not shop.get("umgebungen"):
        grund = json.loads(json.dumps(sw.DEFAULT_CONFIG))
        grund.update(shop)
        grund["umgebungen"] = {"dev": dict(sw.DEFAULT_UMGEBUNG),
                               "prod": dict(sw.DEFAULT_UMGEBUNG)}
        grund["aktive_umgebung"] = "dev"
        shop = grund
    cfg["shopware_publisher"] = shop
    return cfg


def speichere_config(cfg: dict) -> None:
    CONFIG_PFAD.write_text(json.dumps(cfg, indent=2, ensure_ascii=False),
                           encoding="utf-8")


def _misch(vorgabe: dict, eigenes: dict) -> dict:
    zusammen = json.loads(json.dumps(vorgabe))      # frische Kopie
    zusammen.update(eigenes or {})
    return zusammen


def cfg_cover(cfg: dict) -> dict:
    """Konfiguration für cover_previews — der Ablageort ist dort
    `artikeldaten_dir`."""
    c = _misch(cp.DEFAULT_CONFIG, cfg.get("cover_previews", {}))
    c["artikeldaten_dir"] = cfg.get("ablageort", "")
    return c


def cfg_pibi(cfg: dict) -> dict:
    return _misch(pb.DEFAULT_CONFIG, cfg.get("pi_bi_generator", {}))


def cfg_shop(cfg: dict) -> dict:
    """Die Shopware-Konfiguration — **lebend**, nicht kopiert.

    Wer hier das Secret setzt oder eine Zuordnung wählt, schreibt in die
    Konfiguration des Durchgangs; ein anschließendes ``speichere_config``
    behält es. Eine Kopie würde solche Änderungen still verschlucken.
    """
    return cfg.setdefault("shopware_publisher", {})


# ---------------------------------------------------------------------
# Die Schritte
# ---------------------------------------------------------------------

# Die Checklisten sind bewusst fest: sie sollen jedes Mal dieselben Fragen
# stellen. Was dabei zu prüfen ist, weiß nur ein Mensch — das Werkzeug greift
# nicht in Texte ein und zählt nichts nach.
SCHRITTE: list[dict] = [
    {
        "id": "cover",
        "titel": "1 — Cover",
        "checkliste": [
            "Schnittmarken saßen richtig",
            "2D-Vorderseite ohne weißen Rand",
            "3D-Mockup sauber (keine helle Kante an der Buchkante)",
            "TIF freigestellt — oder nicht gebraucht",
        ],
    },
    {
        "id": "pibi",
        "titel": "2 — pi & bi",
        "checkliste": [
            "Werbetext in Word auf eine Seite gekürzt",
            "PDF aus Word exportiert",
            "Titel, Untertitel und Preis stimmen",
        ],
    },
    {
        "id": "shop",
        "titel": "3 — Shop",
        "checkliste": [
            "Kategorien gewählt",
            "Beschreibung im Admin angesehen",
            "Cover sitzt",
            "(optional) Buch auf aktiv gestell",
        ],
    },
]


def schritt(schritt_id: str) -> dict:
    for s in SCHRITTE:
        if s["id"] == schritt_id:
            return s
    raise KeyError(schritt_id)


# ---------------------------------------------------------------------
# Die Eingangsprobe
# ---------------------------------------------------------------------

class PaarFehler(RuntimeError):
    """Umschlag-PDF und ONIX gehören nicht zusammen."""


def pruefe_paar(pdf_pfad, xml_pfad, cfg: dict) -> dict:
    """Umschlag-PDF und ONIX gegeneinander halten, bevor irgendetwas läuft.

    Zwei Proben:

    1. **ISBN** — aus dem PDF-Text und aus der ONIX. Weichen sie ab, ist etwas
       vertauscht; das ist ein Abbruch, kein Hinweis.
    2. **Format** — die gemessene Vorderseite gegen die ONIX-Maße. Das fängt
       den Fall, den die ISBN nicht fängt: richtige Nummer, falsche Datei (etwa
       eine ältere Auflage mit anderem Format). Weil die Beschnittzugabe
       schwankt, ist das nur eine **Warnung**; entscheiden muss ein Mensch.

    Rückgabe: {"sc", "isbn", "felder", "buch", "reg", "doc", "pdf_pfad",
               "xml_pfad", "gemessen_cm", "onix_cm", "warnungen"}
    Das offene ``doc`` wird mitgegeben, damit der Cover-Schritt das PDF nicht
    ein zweites Mal einlesen muss.
    """
    ccfg = cfg_cover(cfg)
    doc = cp.oeffne_pdf(pdf_pfad)
    warnungen: list[str] = []

    isbn_pdf = cp.extrahiere_isbn(doc)
    if not isbn_pdf:
        doc.close()
        raise PaarFehler(
            f"Im Umschlag-PDF steht keine ISBN ({Path(pdf_pfad).name}).\n"
            "Ohne sie lässt sich nicht prüfen, ob PDF und ONIX zusammengehören.")

    felder = sw.lade_buchfelder(xml_pfad, cfg_shop(cfg))
    if isbn_pdf != felder["isbn13"]:
        doc.close()
        raise PaarFehler(
            f"PDF und ONIX gehören nicht zusammen:\n\n"
            f"    Umschlag-PDF   {isbn_pdf}\n"
            f"    ONIX-XML       {felder['isbn13']}  ({felder['titel']})")

    reg = cp.finde_schnittlinien(doc, ccfg)
    gemessen = cp.front_masse_cm(reg)
    onix = (felder.get("breite_cm"), felder.get("hoehe_cm"))

    if reg.front is None:
        warnungen.append(
            "Die Schnittmarken geben Vorder-/Rückseite/Rücken nicht her — im "
            "Cover-Schritt müssen die blauen Linien von Hand justiert werden.")
    elif gemessen and all(onix):
        tol = float(cfg.get("format_tol_cm", 1.0))
        db, dh = abs(gemessen[0] - onix[0]), abs(gemessen[1] - onix[1])
        if db > tol or dh > tol:
            warnungen.append(
                f"Format weicht ab: gemessen {gemessen[0]:.1f} x "
                f"{gemessen[1]:.1f} cm (mit Beschnitt), ONIX sagt "
                f"{onix[0]:.1f} x {onix[1]:.1f} cm. Ist das wirklich der "
                f"Umschlag zu diesem Buch?")
    elif not all(onix):
        warnungen.append("Die ONIX nennt keine Maße — Formatprobe entfällt.")

    buch = pb.lade_buchdaten(xml_pfad, cfg_pibi(cfg))

    return {"sc": cp.shortcode_aus_isbn(isbn_pdf), "isbn": isbn_pdf,
            "felder": felder, "buch": buch, "reg": reg, "doc": doc,
            "pdf_pfad": str(pdf_pfad), "xml_pfad": str(xml_pfad),
            "gemessen_cm": gemessen, "onix_cm": onix, "warnungen": warnungen}


# ---------------------------------------------------------------------
# Buchordner und Stand
# ---------------------------------------------------------------------

STAND_DATEI = "durchgang.json"


def buchordner(sc: str, titel: str, cfg: dict) -> tuple[Path, bool]:
    """Der Ordner dieses Buchs — Namensregel wie in cover_previews."""
    return cp.ziel_ordner(sc, titel, cfg_cover(cfg))


def lade_stand(ordner) -> dict:
    """Stand des Durchgangs aus dem Buchordner (leer, wenn es keinen gibt)."""
    p = Path(ordner) / STAND_DATEI
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            pass
    return {"schritte": {}}


def speichere_stand(ordner, stand: dict) -> Path:
    """Stand in den Buchordner schreiben.

    Er liegt beim Buch, nicht beim Werkzeug: so sieht auch Wochen später und
    an einem anderen Rechner jeder, was geprüft wurde und was noch offen ist.
    """
    ordner = Path(ordner)
    ordner.mkdir(parents=True, exist_ok=True)
    p = ordner / STAND_DATEI
    p.write_text(json.dumps(stand, indent=2, ensure_ascii=False),
                 encoding="utf-8")
    return p


def vermerke_lauf(stand: dict, schritt_id: str, *, quelle: str = "",
                  dateien: list | None = None, hinweise: list | None = None) -> dict:
    """Festhalten, dass ein Schritt gelaufen ist (Häkchen bleiben unberührt)."""
    eintrag = stand.setdefault("schritte", {}).setdefault(schritt_id, {})
    eintrag["gelaufen_am"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    eintrag["quelle"] = str(quelle)
    eintrag["dateien"] = [Path(d).name for d in (dateien or [])]
    eintrag["hinweise"] = list(hinweise or [])
    eintrag.setdefault("haken", [])
    return stand


def setze_haken(stand: dict, schritt_id: str, haken: list[str]) -> dict:
    stand.setdefault("schritte", {}).setdefault(schritt_id, {})["haken"] = list(haken)
    return stand


def ist_gelaufen(stand: dict, schritt_id: str) -> bool:
    return bool(stand.get("schritte", {}).get(schritt_id, {}).get("gelaufen_am"))


def ist_abgehakt(stand: dict, schritt_id: str) -> bool:
    """Alle Häkchen des Schritts gesetzt? Nur dann geht es weiter."""
    gesetzt = set(stand.get("schritte", {}).get(schritt_id, {}).get("haken", []))
    return set(schritt(schritt_id)["checkliste"]) <= gesetzt


def ist_fertig(stand: dict, schritt_id: str) -> bool:
    return ist_gelaufen(stand, schritt_id) and ist_abgehakt(stand, schritt_id)


# ---------------------------------------------------------------------
# Die drei Schritte
# ---------------------------------------------------------------------

def schritt_cover(paar: dict, titel: str, cfg: dict, *, mit_2d=True,
                  mit_3d=True, vorlage=None, log=print) -> dict:
    """Schritt 1 — ruft cover_previews.

    ``share_pflicht=True``: liegt der Ablageort nicht vor, bricht es ab statt
    still neben die .exe zu schreiben. In einer Kette wäre ein Buchordner an
    unerwarteter Stelle fatal — die beiden Folgeschritte suchten ins Leere.
    """
    return cp.lauf(paar["pdf_pfad"], titel, cfg_cover(cfg),
                   doc=paar["doc"], reg=paar["reg"],
                   mit_2d=mit_2d, mit_3d=mit_3d, vorlage=vorlage,
                   share_pflicht=True, log=log)


def schritt_pibi(paar: dict, ordner, cfg: dict, log=print) -> list[Path]:
    """Schritt 2 — ruft pi_bi_generator, mit dem Cover aus Schritt 1."""
    ordner = Path(ordner)
    ccfg = cfg_cover(cfg)
    muster = ccfg.get("muster_2d", cp.DEFAULT_CONFIG["muster_2d"])
    daten, suffix = None, ".jpg"
    for dpi in (int(ccfg.get("dpi_print", 300)), int(ccfg.get("dpi_web", 72))):
        p = ordner / muster.format(dpi=dpi, sc=paar["sc"])
        if p.exists():
            daten, suffix = pb.lade_cover_datei(p), p.suffix
            log(f"Cover aus Schritt 1: {p.name}")
            break
    if daten is None:
        log("⚠ Kein Cover aus Schritt 1 gefunden — die .docx behalten das "
            "Platzhalter-Cover.")
    return pb.erzeuge_alle(paar["buch"], ordner, cfg_pibi(cfg),
                           cover_bytes=daten, cover_suffix=suffix, log=log)


def schritt_shop(paar: dict, ordner, cfg: dict, *, secret: str,
                 kategorien: list[str] | None = None, dry_run: bool = False,
                 ueberschreiben: bool = False, log=print) -> dict:
    """Schritt 3 — ruft shopware_publisher, Bilder aus dem Buchordner."""
    scfg = cfg_shop(cfg)
    bilder = sw.finde_bilder(paar["sc"], scfg, ordner=Path(ordner))
    return sw.veroeffentliche(paar["felder"], scfg, bilder, secret=secret,
                              dry_run=dry_run, ueberschreiben=ueberschreiben,
                              kategorien=kategorien, log=log)
