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
    # Gearbeitet wird ÖRTLICH, abgelegt wird am Ende auf dem Netzordner.
    # Der Grund ist gemessen: derselbe Cover-Schritt braucht örtlich 1,7 s und
    # über die Netzeinbindung ein Vielfaches — er bewegt rund 11 MB.
    #
    # "arbeitsordner" = die schnelle Werkbank (leer = ~/Buch_Arbeit)
    # "ablageort"     = wohin der fertige Buchordner in Schritt 4 wandert
    "arbeitsordner": "",
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
    # Webserver für die Presse-Dateien. Die Knöpfe auf der Produktseite
    # (Presseinfo, 2D, 3D, Blick ins Buch) sind KEIN Produktfeld — das Template
    # zeigt sie, wenn die Datei unter dem erwarteten Pfad liegt. Gemessen an
    # zwölf Fällen: Knopf da genau dann, wenn Datei da.
    #
    # Der Server bietet nur SSH an (21 und 990 sind zu), also SFTP. Das
    # Passwort liegt verschlüsselt daneben, wie das Shop-Secret.
    # ACHTUNG: Dateisystempfade, KEINE URL-Pfade. Die Seite liegt unter
    # verlag-regionalkultur.de/presse/… — auf der Platte aber unter
    # /var/www/shopware/public/presse. Das public/ von Shopware ist das
    # Web-Verzeichnis. Wer den URL-Pfad hier einträgt, landet in der
    # Dateisystemwurzel und scheitert an den Rechten.
    "sftp": {
        "host": "verlag-regionalkultur.de",
        "port": 22,
        "benutzer": "sftpuser",
        "presse_basis": "/var/www/shopware/public/presse",
        "newsletter_basis": "/var/www/shopware/public/newsletter_",
        "hostkey": "",          # beim ersten Verbinden gemerkt (wie SSH selbst)
    },
    "config_version": 2,
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

    # Version 2: die SFTP-Basispfade waren aus den URL-Pfaden abgeleitet
    # ("/presse") statt aus dem Dateisystem ("/var/www/shopware/public/presse").
    # Das muss auch eine gewachsene Konfiguration mitbekommen — Passwort und
    # Serverschlüssel bleiben dabei unangetastet.
    if int(cfg.get("config_version", 1)) < 2:
        z = cfg.setdefault("sftp", {})
        for k in ("presse_basis", "newsletter_basis"):
            z[k] = DEFAULT_CONFIG["sftp"][k]
        if not (z.get("benutzer") or "").strip():
            z["benutzer"] = DEFAULT_CONFIG["sftp"]["benutzer"]
        cfg["config_version"] = 2
    return cfg


def speichere_config(cfg: dict) -> None:
    CONFIG_PFAD.write_text(json.dumps(cfg, indent=2, ensure_ascii=False),
                           encoding="utf-8")


def _misch(vorgabe: dict, eigenes: dict) -> dict:
    zusammen = json.loads(json.dumps(vorgabe))      # frische Kopie
    zusammen.update(eigenes or {})
    return zusammen


def arbeitsordner(cfg: dict) -> Path:
    """Der örtliche Arbeitsordner — hier entstehen die Dateien.

    Wird angelegt, wenn es ihn nicht gibt: er liegt auf der eigenen Platte, da
    darf das Werkzeug das. Der Netzordner dagegen wird nie angelegt — den gibt
    es, oder der Pfad ist falsch.
    """
    raw = (cfg or {}).get("arbeitsordner") or ""
    p = (Path(os.path.expandvars(str(raw))).expanduser() if raw
         else Path.home() / "Buch_Arbeit")
    p.mkdir(parents=True, exist_ok=True)
    return p


def cfg_cover(cfg: dict) -> dict:
    """Konfiguration für cover_previews.

    `artikeldaten_dir` zeigt hier auf den ÖRTLICHEN Arbeitsordner, nicht auf
    den Netzordner: gearbeitet wird schnell, abgelegt wird in Schritt 4.
    """
    c = _misch(cp.DEFAULT_CONFIG, cfg.get("cover_previews", {}))
    c["artikeldaten_dir"] = str(arbeitsordner(cfg))
    return c


def cfg_pibi(cfg: dict) -> dict:
    return _misch(pb.DEFAULT_CONFIG, cfg.get("pi_bi_generator", {}))


def sftp_zugang(cfg: dict) -> dict:
    """Der SFTP-Abschnitt — **lebend**, nicht kopiert.

    Wer hier das Passwort setzt, schreibt in die Konfiguration des Durchgangs;
    ein anschließendes ``speichere_config`` behält es.
    """
    z = cfg.setdefault("sftp", {})
    for k, v in DEFAULT_CONFIG["sftp"].items():
        z.setdefault(k, v)
    return z


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
            "(optional) Buch auf aktiv gestellt",
        ],
    },
    {
        "id": "ablegen",
        "titel": "4 — Ablegen",
        "checkliste": [
            "Alle Dateien sind auf dem Netzordner angekommen",
            "Der örtliche Arbeitsordner kann weg",
        ],
    },
    {
        "id": "presse",
        "titel": "5 — Presse",
        "checkliste": [
            "Die Knöpfe auf der Produktseite sind da und öffnen das Richtige",
            "Blick ins Buch zeigt die richtigen Seiten",
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
    """Der ÖRTLICHE Ordner dieses Buchs — Namensregel wie in cover_previews."""
    return cp.ziel_ordner(sc, titel, cfg_cover(cfg))


def share_basis(cfg: dict) -> Path | None:
    """Der Netzordner selbst — None, wenn nicht erreichbar.

    Angelegt wird hier nichts: einen Netzordner, den es nicht gibt, hält man
    besser für einen falschen Pfad als für eine fehlende Ablage.
    """
    raw = (cfg or {}).get("ablageort") or ""
    if not raw:
        return None
    basis = Path(os.path.expandvars(str(raw))).expanduser()
    try:
        return basis if basis.is_dir() else None
    except OSError:                      # Netzpfad nicht erreichbar
        return None


def share_ordner(quelle, cfg: dict) -> Path | None:
    """Wohin dieser Buchordner am Ende soll — None, wenn nicht erreichbar."""
    basis = share_basis(cfg)
    return None if basis is None else basis / Path(quelle).name


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


def schritt_ablegen(quelle, cfg: dict, *, log=print) -> dict:
    """Schritt 4 — den örtlichen Buchordner auf den Netzordner legen.

    Kopiert, **prüft nach** und meldet, was ankam. Die örtliche Kopie bleibt
    stehen: geht beim Übertragen etwas schief, ist die Arbeit noch da. Wann sie
    weg kann, sagt die Checkliste.

    Zwischendateien (``_slot_*.png``, ``_mockup_*.jsx``) bleiben zurück — sie
    sind Futter für Photoshop und haben im Artikelordner nichts zu suchen. Auf
    einem Windows-Lauf räumt ``cover_previews`` sie ohnehin selbst weg; nach
    einem Trockenlauf liegen sie noch da.

    Liegen am Ziel schon gleichnamige Dateien, werden sie **nicht**
    überschrieben, sondern nach ``_alt/<Zeitstempel>/`` weggesichert — dasselbe
    Muster, das cover_previews im Artikelordner benutzt.

    ``quelle`` ist der **vorhandene** örtliche Buchordner, nicht Kurzcode und
    Titel: der Zielname wird davon abgeleitet. Sonst zeigte ein nachträglich
    geänderter Titel ins Leere — der Ordner hieße noch alt, gesucht würde neu.

    Rückgabe: {"ziel", "quelle", "kopiert", "uebersprungen", "gesichert",
               "fehler"}
    """
    import shutil

    quelle = Path(quelle)
    if not quelle.is_dir():
        raise RuntimeError(f"Es gibt keinen örtlichen Buchordner:\n{quelle}")

    basis = share_basis(cfg)
    if basis is None:
        raise RuntimeError(
            "Der Ablageort ist nicht erreichbar:\n"
            f"{cfg.get('ablageort') or '(nicht eingetragen)'}")
    ziel = basis / quelle.name

    mitnehmen = sorted(p for p in quelle.iterdir()
                       if p.is_file() and not p.name.startswith("_"))
    uebersprungen = sorted(p.name for p in quelle.iterdir()
                           if p.is_file() and p.name.startswith("_"))
    if not mitnehmen:
        raise RuntimeError(f"Im Buchordner liegt nichts zum Ablegen:\n{quelle}")

    ziel.mkdir(parents=True, exist_ok=True)
    vorhanden = [ziel / p.name for p in mitnehmen if (ziel / p.name).exists()]
    gesichert = None
    if vorhanden:
        gesichert = cp.sichere_weg(
            vorhanden, datetime.now().strftime("%Y-%m-%d_%H%M%S"))
        log(f"{len(vorhanden)} vorhandene Datei(en) nach "
            f"_alt/{gesichert.name}/ gesichert.")

    kopiert, fehler = [], []
    for p in mitnehmen:
        z = ziel / p.name
        try:
            log(f"Kopiere {p.name} …")
            # copyfile überträgt NUR die Daten. copy2 würde zusätzlich
            # Zeitstempel und Rechte setzen — das quittiert eine SMB-Freigabe
            # mit "Errno 95: Operation not supported", OBWOHL die Datei längst
            # vollständig angekommen ist. Das sah dann nach sieben Fehlern aus,
            # während in Wahrheit alles dalag.
            shutil.copyfile(p, z)
        except OSError as e:
            fehler.append(f"{p.name}: {e}")
            continue
        try:
            shutil.copystat(p, z)        # nett, aber nicht überall möglich
        except OSError:
            pass
        # Nachprüfen statt vertrauen — über das Netz bricht ein Kopiervorgang
        # gern in der Mitte ab, und eine halbe Datei sieht aus wie eine ganze.
        try:
            if z.stat().st_size != p.stat().st_size:
                fehler.append(f"{p.name}: Größe weicht ab "
                              f"({z.stat().st_size} statt {p.stat().st_size} Bytes)")
                continue
        except OSError as e:
            fehler.append(f"{p.name}: nach dem Kopieren nicht lesbar ({e})")
            continue
        kopiert.append(z)

    return {"ziel": ziel, "quelle": quelle, "kopiert": kopiert,
            "uebersprungen": uebersprungen, "gesichert": gesichert,
            "fehler": fehler}


def spiegle_stand(ordner, ziel) -> bool:
    """`durchgang.json` noch einmal ans Ziel kopieren.

    Nötig, weil der Ablege-Schritt erst NACH dem Kopieren vermerkt wird — die
    Fassung auf dem Netzordner wäre sonst immer einen Schritt alt und wüsste
    nichts davon, dass sie selbst abgelegt wurde.
    """
    import shutil
    q = Path(ordner) / STAND_DATEI
    if not q.exists():
        return False
    try:
        shutil.copyfile(q, Path(ziel) / STAND_DATEI)
        return True
    except OSError:
        return False


# ---------------------------------------------------------------------
# Shop-Zugang vom ShopwarePublisher übernehmen
# ---------------------------------------------------------------------

def publisher_config_pfade() -> list[Path]:
    """Wo die config.json des ShopwarePublisher liegen könnte.

    Ausgeliefert liegen die Werkzeuge als Geschwisterordner unter VR-Tools,
    im Quellbaum nebeneinander im Repo.
    """
    return [APP_DIR.parent / "ShopwarePublisher" / "config.json",
            APP_DIR.parent / "shopware_publisher" / "config.json",
            Path(sw.CONFIG_PFAD)]


def finde_publisher_config() -> Path | None:
    for p in publisher_config_pfade():
        try:
            if p.is_file():
                return p
        except OSError:
            continue
    return None


def uebernimm_shop_zugang(cfg: dict, pfad=None) -> str:
    """Zugangsdaten aus der config.json des ShopwarePublisher übernehmen.

    Spart das zweite Eintippen. Mitgenommen werden Shop-URL, Zugriffsschlüssel
    und das **verschlüsselte** Secret samt Salt — beides ist selbsttragend, es
    lässt sich mit demselben Master-Passwort entsperren wie dort. Das Passwort
    selbst wandert nicht mit und wird auch nie gespeichert.

    Rückgabe: eine Zeile, was übernommen wurde (für die Anzeige).
    """
    pfad = Path(pfad) if pfad else finde_publisher_config()
    if not pfad or not pfad.is_file():
        raise RuntimeError(
            "Keine config.json des ShopwarePublisher gefunden. Gesucht in:\n"
            + "\n".join(f"  {p}" for p in publisher_config_pfade()))
    try:
        fremd = json.loads(pfad.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as e:
        raise RuntimeError(f"{pfad} lässt sich nicht lesen:\n{e}")

    umgebungen = fremd.get("umgebungen") or {}
    if not umgebungen:
        raise RuntimeError(f"In {pfad} stehen keine Umgebungen.")

    eigene = cfg.setdefault("shopware_publisher", {})
    ziel = eigene.setdefault("umgebungen", {})
    uebernommen = []
    for name, umg in umgebungen.items():
        if not umg.get("shop_url"):
            continue                     # leere Umgebung bringt nichts
        ziel[name] = json.loads(json.dumps(umg))     # frische Kopie
        uebernommen.append(
            f"{name} ({'mit Secret' if umg.get('secret_enc') else 'ohne Secret'})")
    if not uebernommen:
        raise RuntimeError(f"In {pfad} ist keine Umgebung eingerichtet.")
    eigene["aktive_umgebung"] = fremd.get("aktive_umgebung", "dev")
    return f"Übernommen aus {pfad}: " + ", ".join(uebernommen)


# ---------------------------------------------------------------------
# Schritt 5 — Presse-Dateien auf den Webserver
# ---------------------------------------------------------------------

def presse_dateien(ordner, sc: str, cfg: dict, *, mit_pi: bool = True,
                   bib_pdf=None) -> list[dict]:
    """Welche Datei wohin gehört — und ob sie da ist.

    Die vier Knöpfe auf der Produktseite sind kein Produktfeld: das Template
    zeigt sie, wenn die Datei unter dem erwarteten Pfad liegt. Hochladen und
    Anzeigen sind also dasselbe. Deshalb steuert ``mit_pi`` die Presseinfo —
    einen zweiten Schalter gibt es nicht.

    ``bib_pdf`` ist die von Hand gewählte Blick-ins-Buch-Datei; sie entsteht
    nicht im Durchgang und kann alles Mögliche heißen.

    Rückgabe je Eintrag: {"art", "lokal", "fern", "da", "pflicht"}
    """
    ordner = Path(ordner)
    z = sftp_zugang(cfg)
    presse = (z.get("presse_basis") or "/presse").rstrip("/")
    news = (z.get("newsletter_basis") or "/newsletter_").rstrip("/")
    ccfg = cfg_cover(cfg)
    dpi = int(ccfg.get("dpi_print", 300))
    m2d = ccfg.get("muster_2d", cp.DEFAULT_CONFIG["muster_2d"]).format(dpi=dpi, sc=sc)
    m3d = ccfg.get("muster_3d", cp.DEFAULT_CONFIG["muster_3d"]).format(dpi=dpi, sc=sc)
    mpng = ccfg.get("muster_3d_png", cp.DEFAULT_CONFIG["muster_3d_png"]).format(sc=sc)

    liste: list[dict] = []

    def dazu(art, lokal, fern, pflicht=True):
        lokal = Path(lokal)
        liste.append({"art": art, "lokal": lokal, "fern": fern,
                      "da": lokal.is_file(), "pflicht": pflicht})

    # Die Presseinfo kommt als PDF aus Word — der Generator liefert nur .docx.
    # Erwartet wird sie unter demselben Namen im Buchordner.
    if mit_pi:
        dazu("Presseinfo", ordner / f"PI_{sc}.pdf", f"{presse}/PI/PI_{sc}.pdf")
    dazu("3D-Cover", ordner / m3d, f"{presse}/3D/{m3d}")
    dazu("2D-Cover", ordner / m2d, f"{presse}/2D/{m2d}")
    if bib_pdf:
        dazu("Blick ins Buch", bib_pdf, f"{presse}/bib/bib_{sc}.pdf")
    # Das Newsletter-Cover ist zugleich die dritte Bildquelle des Publishers.
    dazu("Newsletter-Cover", ordner / mpng, f"{news}/{mpng}", pflicht=False)
    return liste


class SftpFehler(RuntimeError):
    """Verbindung oder Übertragung ist gescheitert."""


def _sftp_verbinden(cfg: dict, passwort: str):
    """Verbindung aufbauen und den Serverschlüssel prüfen.

    Beim ersten Mal wird der Schlüssel gemerkt, danach verglichen — dasselbe
    Vorgehen wie SSH selbst. Ändert er sich, wird abgebrochen statt gefragt:
    hier werden Dateien auf einen öffentlichen Webserver geschoben, das ist
    kein guter Ort für ein Achselzucken.
    """
    import paramiko

    z = sftp_zugang(cfg)
    host, port = z.get("host") or "", int(z.get("port") or 22)
    benutzer = (z.get("benutzer") or "").strip()
    if not (host and benutzer):
        raise SftpFehler("Server oder Benutzer fehlt — bitte eintragen.")

    try:
        transport = paramiko.Transport((host, port))
        transport.connect(username=benutzer, password=passwort)
    except Exception as e:
        raise SftpFehler(f"Verbindung zu {host}:{port} gescheitert: {e}")

    schluessel = transport.get_remote_server_key()
    finger = schluessel.get_base64()
    gemerkt = (z.get("hostkey") or "").strip()
    if gemerkt and gemerkt != finger:
        transport.close()
        raise SftpFehler(
            "Der Serverschlüssel hat sich geändert. Entweder wurde der Server "
            "neu aufgesetzt — dann den Eintrag 'hostkey' in der config.json "
            "leeren — oder es antwortet jemand anderes.")
    if not gemerkt:
        z["hostkey"] = finger
    return transport, paramiko.SFTPClient.from_transport(transport)


def _sftp_mkdirs(sftp, pfad: str) -> None:
    """Fehlende Ordner anlegen — SFTP kennt kein mkdir -p."""
    teile = [t for t in pfad.strip("/").split("/") if t]
    lauf = ""
    for t in teile:
        lauf += "/" + t
        try:
            sftp.stat(lauf)
        except IOError:
            sftp.mkdir(lauf)


def sftp_pruefen(cfg: dict, passwort: str) -> dict:
    """Nachsehen, wohin man kommt — ohne etwas hochzuladen.

    Beantwortet die Fragen, die man sonst rät: Wo landet der Benutzer? Gibt es
    die Zielordner? Was liegt darin? Ein falsch geratener Basispfad fällt hier
    auf und nicht erst, wenn Dateien an der falschen Stelle liegen.

    Rückgabe: {"start", "presse", "newsletter", "hostkey_neu"}
    """
    z = sftp_zugang(cfg)
    vorher = (z.get("hostkey") or "").strip()
    transport, sftp = _sftp_verbinden(cfg, passwort)
    try:
        bericht = {"start": sftp.normalize("."),
                   "hostkey_neu": not vorher,
                   "presse": None, "newsletter": None}
        for name, pfad in (("presse", z.get("presse_basis")),
                           ("newsletter", z.get("newsletter_basis"))):
            try:
                sftp.stat(pfad)
                inhalt = sorted(sftp.listdir(pfad))
                bericht[name] = {"pfad": pfad, "da": True,
                                 "inhalt": inhalt[:12],
                                 "anzahl": len(inhalt)}
            except IOError as e:
                bericht[name] = {"pfad": pfad, "da": False, "grund": str(e),
                                 "inhalt": [], "anzahl": 0}
        return bericht
    finally:
        try:
            sftp.close(); transport.close()
        except Exception:
            pass


def schritt_presse(ordner, sc: str, cfg: dict, *, passwort: str,
                   mit_pi: bool = True, bib_pdf=None, log=print) -> dict:
    """Schritt 5 — die Presse-Dateien auf den Webserver legen.

    Übertragen wird nur, was da ist; nach jeder Datei wird die Größe am Ziel
    verglichen. Über eine Leitung bricht eine Übertragung gern in der Mitte ab,
    und eine halbe PDF sieht aus wie eine ganze — nur dass dann der Knopf auf
    der Produktseite ins Leere führt.

    Rückgabe: {"geladen", "fehlend", "fehler"}
    """
    liste = presse_dateien(ordner, sc, cfg, mit_pi=mit_pi, bib_pdf=bib_pdf)
    fehlend = [e for e in liste if not e["da"]]
    zu_laden = [e for e in liste if e["da"]]
    if not zu_laden:
        raise SftpFehler("Es gibt nichts hochzuladen — keine der erwarteten "
                         "Dateien liegt im Buchordner.")

    transport, sftp = _sftp_verbinden(cfg, passwort)
    geladen, fehler = [], []
    try:
        for e in zu_laden:
            ziel_ordner = str(Path(e["fern"]).parent).replace("\\", "/")
            try:
                _sftp_mkdirs(sftp, ziel_ordner)
                log(f"Lade hoch: {e['art']} → {e['fern']}")
                sftp.put(str(e["lokal"]), e["fern"])
                fern_gross = sftp.stat(e["fern"]).st_size
                lokal_gross = e["lokal"].stat().st_size
                if fern_gross != lokal_gross:
                    fehler.append(f"{e['art']}: Größe weicht ab "
                                  f"({fern_gross} statt {lokal_gross} Bytes)")
                    continue
            except Exception as ex:
                fehler.append(f"{e['art']}: {ex}")
                continue
            geladen.append(e)
    finally:
        try:
            sftp.close(); transport.close()
        except Exception:
            pass

    return {"geladen": geladen, "fehlend": fehlend, "fehler": fehler}
