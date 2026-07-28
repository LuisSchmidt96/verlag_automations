"""
Shopware-Publisher Core-Logik
=============================
Reine Datenlogik ohne UI-Abhängigkeiten — liest einen VLB-ONIX-XML-Datensatz,
baut daraus ein Shopware-6-Produkt und legt es über die Admin-API als
**Entwurf** (inaktiv) an.

Ablauf:
1. lade_buchfelder(xml)          -> Buchdaten (aus ONIX 3.1)
2. finde_bilder(sc, cfg)         -> Cover/3D-Bilder vom Artikeldaten-Share
3. ShopClient(...).verbinde()    -> Token (client_credentials)
4. baue_produkt(...)             -> Produkt-Payload
5. ShopClient.veroeffentliche()  -> Medien hochladen + Produkt upserten

Idempotent: die Produkt-ID wird deterministisch aus der ISBN abgeleitet, ein
zweiter Lauf aktualisiert also dasselbe Produkt statt ein zweites anzulegen.
"""

from __future__ import annotations

import base64
import hashlib
import html
import json
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path


# ---------------------------------------------------------------------
# Ordner / Konfiguration — Muster wie die anderen Tools (neben der .exe)
# ---------------------------------------------------------------------

def _base_dir() -> Path:
    """Ordner neben der .exe (PyInstaller-Build) bzw. neben dem Tool-Code."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


APP_DIR = _base_dir()
APP_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_PFAD = APP_DIR / "config.json"


# Alles, was je Shop unterschiedlich ist, steckt in einer Umgebung ("dev"/"prod").
# Wichtig: tax_id, currency_id, category_id & Co. sind **UUIDs des jeweiligen
# Shops** — dieselbe Kategorie hat auf dev und prod verschiedene IDs. Sie dürfen
# deshalb nicht global stehen, sonst schreibt man dev-IDs in den Produktivshop.
DEFAULT_UMGEBUNG = {
    # --- Zugang (Admin -> Einstellungen -> System -> Integrationen) ----------
    "shop_url": "",                 # z. B. https://shop.verlag-regionalkultur.de
    "access_key_id": "",            # wie ein Benutzername — nicht geheim
    # Das Secret liegt NIE im Klartext auf der Platte: es wird mit einem
    # Schlüssel verschlüsselt, der aus dem Master-Passwort abgeleitet wird
    # (scrypt + AES-GCM). Wer die config.json kopiert, hat nur Chiffretext.
    "secret_enc": "",               # base64(nonce + ciphertext)
    "kdf_salt": "",                 # base64(salt)
    # Dev-Store mit nicht überprüfbarem Zertifikat: auf false setzen.
    # Im Produktivshop true lassen!
    "tls_pruefen": True,

    # --- Zuordnungen aus diesem Shop (beim "Verbinden" befüllt) --------------
    "tax_id": "",                   # Steuersatz-ID (Bücher: 7 %)
    "tax_rate": 7.0,                # zugehöriger Satz, für die Netto-Rechnung
    "currency_id": "",              # EUR
    "category_id": "",              # Zielkategorie (optional)
    "manufacturer_id": "",          # Hersteller/Verlag (optional)
    # Ohne Sales-Channel-Sichtbarkeit ist das Produkt im Shop UNSICHTBAR —
    # auch aktiv geschaltet. 30 = überall sichtbar (wie im Bestand).
    "sales_channel_id": "",
    "visibility": 30,
    # Layout der Produktseite; wird beim Verbinden aus einem vorhandenen
    # Produkt übernommen, damit neue Produkte gleich aussehen.
    "cms_page_id": "",
}

DEFAULT_CONFIG = {
    # Wird bei jeder Änderung der Verlags-Vorgaben (unten) erhöht, damit
    # bestehende config.json-Dateien die neuen Werte übernehmen (siehe
    # _migriere). Ohne das blieben alte Werte per setdefault eingefroren.
    "config_version": 2,

    # --- Umgebungen ---------------------------------------------------------
    "aktive_umgebung": "dev",
    "umgebungen": {
        "dev": dict(DEFAULT_UMGEBUNG),
        "prod": dict(DEFAULT_UMGEBUNG),
    },

    # --- Produkt-Voreinstellungen (shop-unabhängig) --------------------------
    # Entwurf: Produkt wird angelegt, ist aber im Shop nicht sichtbar.
    "aktiv": False,
    # Lieferbarkeit — für alle Bücher gleich (Vorgabe des Verlags):
    "default_stock": 9999,          # Lagerbestand
    "is_closeout": True,            # Abverkauf AN
    "restock_time": None,           # Wiederauffüllzeit leer lassen
    "min_purchase": 1,              # Mindestabnahme
    "purchase_steps": 1,            # Staffelung
    "shipping_free": False,         # Versandkostenfrei aus
    # Maximalabnahme und Lieferzeit bleiben bewusst leer.

    # --- customFields (SW5-Migration; das Theme zeigt sie an) ----------------
    "custom_fields": {
        "untertitel": "migration_Shopware5_product_attr1",
        "autor_link": "migration_Shopware5_product_attr2",
        "format":     "migration_Shopware5_product_attr12",
    },
    "autoren_basis_url": "/autoren-herausgeber",

    # --- Bilder (kommen vom cover_previews-Tool auf dem Artikeldaten-Share) --
    "artikeldaten_dir": r"\\C019\d\Online\Webseite\Artikeldaten",
    "muster_2d": "2D_{dpi}_{sc}.jpg",
    "muster_3d": "3D_{dpi}_{sc}.jpg",
    "dpi_web": 72,
    "dpi_print": 300,

    # --- Buchdaten ----------------------------------------------------------
    "verlag_name": "verlag regionalkultur",
    "isbn_prefix": "978-3-95505",
    "einband_map": {
        "BB": "fester Einband",
        "BC": "kartoniert",
        "BE": "Klappenbroschur",
        "PB": "kartoniert",
        "BZ": "Leinen",
    },
    "last_input_dir": "",
}

# Schlüssel, die früher flach in der config.json standen (eine Umgebung).
_ALTE_FLACHE_SCHLUESSEL = tuple(DEFAULT_UMGEBUNG) + ("secret_access_key",)

# Verlags-Vorgaben (für alle Bücher gleich): werden bei einer Versionserhöhung
# aus DEFAULT_CONFIG aufgefrischt, damit alte config.json-Werte nicht per
# setdefault eingefroren bleiben (z. B. default_stock 0 -> 9999).
_VORGABE_SCHLUESSEL = ("aktiv", "default_stock", "is_closeout", "restock_time",
                       "min_purchase", "purchase_steps", "shipping_free",
                       "custom_fields", "autoren_basis_url")


def _migriere(cfg: dict) -> dict:
    """Alte, flache Config in die Umgebungs-Struktur heben und Verlags-Vorgaben
    bei einer Versionserhöhung auffrischen."""
    if "umgebungen" not in cfg:
        alt = {k: cfg.pop(k) for k in _ALTE_FLACHE_SCHLUESSEL if k in cfg}
        umg = dict(DEFAULT_UMGEBUNG)
        umg.update({k: v for k, v in alt.items() if k in DEFAULT_UMGEBUNG})
        if alt.get("secret_access_key"):      # Klartext-Altlast mitnehmen
            umg["secret_access_key"] = alt["secret_access_key"]
        cfg["umgebungen"] = {"dev": umg, "prod": dict(DEFAULT_UMGEBUNG)}
        cfg["aktive_umgebung"] = "dev"

    if cfg.get("config_version", 1) < DEFAULT_CONFIG["config_version"]:
        for k in _VORGABE_SCHLUESSEL:
            cfg[k] = json.loads(json.dumps(DEFAULT_CONFIG[k]))   # frische Kopie
        cfg["config_version"] = DEFAULT_CONFIG["config_version"]
    return cfg


def lade_config() -> dict:
    if CONFIG_PFAD.exists():
        with open(CONFIG_PFAD, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        cfg = _migriere(cfg)
        for k, v in DEFAULT_CONFIG.items():
            cfg.setdefault(k, v)
        for name, umg in cfg.get("umgebungen", {}).items():
            for k, v in DEFAULT_UMGEBUNG.items():
                umg.setdefault(k, v)
        return cfg
    with open(CONFIG_PFAD, "w", encoding="utf-8") as f:
        json.dump(DEFAULT_CONFIG, f, indent=2, ensure_ascii=False)
    return json.loads(json.dumps(DEFAULT_CONFIG))     # tiefe Kopie


# ---------------------------------------------------------------------
# Umgebungen (dev / prod)
# ---------------------------------------------------------------------

def umgebungs_namen(cfg: dict) -> list[str]:
    return list(cfg.get("umgebungen", {}))


def aktive_umgebung(cfg: dict) -> str:
    name = cfg.get("aktive_umgebung") or "dev"
    if name not in cfg.get("umgebungen", {}):
        name = (umgebungs_namen(cfg) or ["dev"])[0]
    return name


def umgebung(cfg: dict, name: str | None = None) -> dict:
    """Die (aktive) Umgebung — Zugang + shopspezifische Zuordnungen."""
    name = name or aktive_umgebung(cfg)
    umgs = cfg.setdefault("umgebungen", {})
    return umgs.setdefault(name, dict(DEFAULT_UMGEBUNG))


def ist_produktiv(name: str) -> bool:
    """Heuristik für die Warnfarbe im GUI."""
    return name.lower().startswith(("prod", "live"))


def effektiv(cfg: dict) -> dict:
    """Globale Einstellungen + aktive Umgebung zu einer flachen Sicht
    zusammenlegen — so arbeiten die Bau-Funktionen unverändert weiter."""
    flach = {k: v for k, v in cfg.items()
             if k not in ("umgebungen", "aktive_umgebung")}
    flach.update(umgebung(cfg))
    return flach


def speichere_config(cfg: dict) -> None:
    with open(CONFIG_PFAD, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)


# ---------------------------------------------------------------------
# Master-Passwort: Secret verschlüsselt ablegen
# ---------------------------------------------------------------------
#
# Ein reiner Passwort-Dialog wäre wirkungslos, solange das Secret im Klartext
# in der config.json steht — man liest die Datei einfach auf. Das Passwort ist
# deshalb der *Schlüssel*: daraus wird per scrypt ein AES-Schlüssel abgeleitet,
# mit dem das Secret verschlüsselt gespeichert wird. AES-GCM ist authentifiziert,
# ein falsches Passwort scheitert also sauber statt Datenmüll zu liefern.

SCRYPT_N, SCRYPT_R, SCRYPT_P = 2 ** 14, 8, 1


class PasswortFehler(RuntimeError):
    pass


def _kdf(passwort: str, salt: bytes) -> bytes:
    return hashlib.scrypt(passwort.encode("utf-8"), salt=salt,
                          n=SCRYPT_N, r=SCRYPT_R, p=SCRYPT_P, dklen=32)


def _aesgcm(schluessel: bytes):
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    return AESGCM(schluessel)


def hat_secret(umg: dict) -> bool:
    """Liegt für diese Umgebung ein verschlüsseltes Secret vor?"""
    return bool(umg.get("secret_enc") and umg.get("kdf_salt"))


def setze_secret(umg: dict, secret: str, passwort: str) -> None:
    """Secret dieser Umgebung mit dem Master-Passwort verschlüsselt ablegen.
    Jede Umgebung bekommt ein eigenes Salt — dev und prod sind unabhängig."""
    if not passwort:
        raise PasswortFehler("Master-Passwort darf nicht leer sein.")
    salt = os.urandom(16)
    nonce = os.urandom(12)
    ct = _aesgcm(_kdf(passwort, salt)).encrypt(nonce, secret.encode("utf-8"), None)
    umg["secret_enc"] = base64.b64encode(nonce + ct).decode()
    umg["kdf_salt"] = base64.b64encode(salt).decode()
    umg.pop("secret_access_key", None)      # evtl. Altlast im Klartext entfernen


def hole_secret(umg: dict, passwort: str) -> str:
    """Secret dieser Umgebung entschlüsseln. Falsches Passwort -> PasswortFehler."""
    if not hat_secret(umg):
        raise PasswortFehler("Für diese Umgebung ist noch kein Secret hinterlegt.")
    roh = base64.b64decode(umg["secret_enc"])
    salt = base64.b64decode(umg["kdf_salt"])
    try:
        klar = _aesgcm(_kdf(passwort, salt)).decrypt(roh[:12], roh[12:], None)
    except Exception as e:                  # InvalidTag u. Ä.
        raise PasswortFehler("Falsches Master-Passwort.") from e
    return klar.decode("utf-8")


def klartext_secret_vorhanden(umg: dict) -> str:
    """Altlast: Secret aus einer früheren Version, das noch im Klartext steht.
    Wird beim ersten Start mit Passwort verschlüsselt."""
    return umg.get("secret_access_key") or ""


# ---------------------------------------------------------------------
# ONIX-Auswertung (VLB, ONIX 3.1, Kurz-Tags ohne Namespace)
# ---------------------------------------------------------------------

_INTRO_LABELS = ("Aus dem Vorwort",)


def _txt(el) -> str:
    return (el.text or "").strip() if el is not None else ""


def _join_und(namen: list[str]) -> str:
    """['A','B','C'] -> 'A, B und C'."""
    namen = [n for n in namen if n]
    if not namen:
        return ""
    if len(namen) == 1:
        return namen[0]
    return ", ".join(namen[:-1]) + " und " + namen[-1]


def _kontributoren(product, rolle: str) -> list[dict]:
    """[{'name': 'Ruth Birkle', 'vorname': 'Ruth', 'nachname': 'Birkle'}, …]

    b039/b040 (Vor-/Nachname) sind für den Autor-Link nötig — b036 ist nur der
    zusammengesetzte Name.
    """
    dd = product.find("descriptivedetail")
    if dd is None:
        return []
    beitraege = []
    for c in dd.findall("contributor"):
        if _txt(c.find("b035")) == rolle:
            seq = _txt(c.find("b034"))
            name = _txt(c.find("b036"))
            vor, nach = _txt(c.find("b039")), _txt(c.find("b040"))
            if not name and nach:
                name = f"{vor} {nach}".strip()
            if not nach and name:               # Rückfall: am letzten Leerzeichen
                teile = name.rsplit(" ", 1)
                vor, nach = (teile[0], teile[1]) if len(teile) == 2 else ("", name)
            if name:
                beitraege.append((int(seq) if seq.isdigit() else 999,
                                  {"name": name, "vorname": vor, "nachname": nach}))
    return [n for _, n in sorted(beitraege, key=lambda x: x[0])]


def _namen(kontributoren: list[dict]) -> list[str]:
    return [k["name"] for k in kontributoren]


_UMLAUTE = {"ä": "ae", "ö": "oe", "ü": "ue", "ß": "ss",
            "Ä": "ae", "Ö": "oe", "Ü": "ue"}


def slug(text: str) -> str:
    """'Sophie Brandes' -> 'sophie-brandes' (Umlaute ausgeschrieben)."""
    t = "".join(_UMLAUTE.get(c, c) for c in (text or ""))
    t = re.sub(r"[^A-Za-z0-9]+", "-", t).strip("-").lower()
    return t


def autor_link(k: dict, basis: str) -> str:
    """<a href='/autoren-herausgeber/b/brandes-sophie/'>Sophie Brandes</a>

    Konvention aus den bestehenden Produkten: Nachname-Vorname, einsortiert
    unter dem Anfangsbuchstaben des Nachnamens.
    """
    nach, vor = k.get("nachname", ""), k.get("vorname", "")
    if not nach:
        return html.escape(k.get("name", ""))
    pfad = slug(f"{nach}-{vor}") if vor else slug(nach)
    initial = slug(nach)[:1]
    url = f"{basis.rstrip('/')}/{initial}/{pfad}/"
    return f"<a href='{url}'>{html.escape(k.get('name', ''))}</a>"


def _zahl_de(x: float) -> str:
    """24.4 -> '24,4'; 24.0 -> '24'."""
    s = f"{x:g}"
    return s.replace(".", ",")


def _entferne_intro(absaetze: list[str]) -> list[str]:
    """Entfernt ein führendes Label wie 'Aus dem Vorwort:'."""
    if not absaetze:
        return absaetze
    for label in _INTRO_LABELS:
        pat = re.compile(rf"^\s*{re.escape(label)}\s*[:\-–—]?\s*", re.IGNORECASE)
        m = pat.match(absaetze[0])
        if m:
            rest = absaetze[0][m.end():].strip()
            return ([rest] if rest else []) + absaetze[1:]
    return absaetze


def shortcode_aus_isbn(isbn13: str) -> str:
    """9783955055721 -> '05-572-1' (wie in den anderen Tools)."""
    e = re.sub(r"\D", "", isbn13)
    return f"{e[7:9]}-{e[9:12]}-{e[12]}"


def lade_buchfelder(xml_pfad, cfg: dict | None = None) -> dict:
    """Liest die für den Shop nötigen Felder aus einer VLB-ONIX-XML.

    Bewusst schlanker als der pi_bi_generator: nur was das Produkt braucht.
    """
    cfg = cfg or DEFAULT_CONFIG
    einband_map = cfg.get("einband_map", DEFAULT_CONFIG["einband_map"])
    isbn_prefix = cfg.get("isbn_prefix", DEFAULT_CONFIG["isbn_prefix"])

    root = ET.parse(str(xml_pfad)).getroot()
    product = root.find("product")
    if product is None:
        raise ValueError("Kein <product>-Element in der XML gefunden — "
                         "ist das eine VLB-ONIX-Datei?")
    dd = product.find("descriptivedetail")

    # ISBN-13 (b221 = 03 oder 15)
    isbn13 = ""
    for pid in product.findall("productidentifier"):
        if _txt(pid.find("b221")) in ("03", "15"):
            isbn13 = _txt(pid.find("b244"))
            break
    if not isbn13 or len(isbn13) != 13 or not isbn13.isdigit():
        raise ValueError(f"Keine gültige ISBN-13 gefunden (gelesen: {isbn13!r}).")

    # Titel + Untertitel (Produktebene, nicht die Collection-Kopie).
    # Achtung: das b029 unter collection/ ist der Untertitel der REIHE.
    titel, untertitel = "", ""
    if dd is not None:
        for td in dd.findall("titledetail"):
            if _txt(td.find("b202")) == "01":
                for te in td.findall("titleelement"):
                    if _txt(te.find("x409")) == "01":
                        titel = _txt(te.find("b203"))
                        untertitel = _txt(te.find("b029"))
                        break
            if titel:
                break

    # Serie / Band
    serientitel, band = "", ""
    if dd is not None:
        coll = dd.find("collection")
        if coll is not None:
            for te in coll.findall("titledetail/titleelement"):
                lvl = _txt(te.find("x409"))
                if lvl == "02" and not serientitel:
                    serientitel = _txt(te.find("b203"))
                if lvl == "01" and not band:
                    band = _txt(te.find("x410"))

    herausgeber = _kontributoren(product, "B01")
    autoren = _kontributoren(product, "A01")

    seiten = _txt(dd.find("extent/b219")) if dd is not None else ""
    einband_code = _txt(dd.find("b012")) if dd is not None else ""
    einband = einband_map.get(einband_code, einband_code)

    verlag = _txt(product.find("publishingdetail/publisher/b081")) \
        or cfg.get("verlag_name", "")

    # Preis (Deutschland)
    preis_brutto, waehrung = 0.0, "EUR"
    for price in product.findall("productsupply/supplydetail/price"):
        terr = price.find("territory")
        if terr is not None and _txt(terr.find("x449")) == "DE":
            try:
                preis_brutto = float(_txt(price.find("j151")))
            except ValueError:
                preis_brutto = 0.0
            waehrung = _txt(price.find("j152")) or "EUR"
            break

    roh_datum = _txt(product.find("publishingdetail/publishingdate/b306"))
    datum, datum_iso = "", ""
    if len(roh_datum) == 8 and roh_datum.isdigit():
        datum = f"{roh_datum[6:8]}.{roh_datum[4:6]}.{roh_datum[0:4]}"
        datum_iso = (f"{roh_datum[0:4]}-{roh_datum[4:6]}-{roh_datum[6:8]}"
                     f"T00:00:00.000+00:00")

    # Maße (measure): x315 01=Höhe, 02=Breite, 03=Dicke, 08=Gewicht
    masse: dict[str, float] = {}
    _MASS = {"01": "hoehe_cm", "02": "breite_cm", "03": "dicke_cm",
             "08": "gewicht_kg"}
    if dd is not None:
        for m in dd.findall("measure"):
            feld = _MASS.get(_txt(m.find("x315")))
            wert, einheit = _txt(m.find("c094")), _txt(m.find("c095")).lower()
            if not feld or not wert:
                continue
            try:
                w = float(wert.replace(",", "."))
            except ValueError:
                continue
            if feld == "gewicht_kg" and einheit in ("gr", "g"):
                w = w / 1000.0                  # Shopware rechnet in kg
            elif feld != "gewicht_kg" and einheit == "mm":
                w = w / 10.0                    # Shopware-Felder sind cm-Werte
            masse[feld] = w

    cover_url = ""
    for sr in product.findall("collateraldetail/supportingresource"):
        if _txt(sr.find("x436")) == "01":
            cover_url = _txt(sr.find("resourceversion/x435"))
            break

    d104 = _txt(product.find("collateraldetail/textcontent/d104"))
    werbetext = [ln.strip() for ln in d104.split("\n") if ln.strip()]
    werbetext = _entferne_intro(werbetext)

    e = isbn13
    return {
        "isbn13": isbn13,
        "isbn13_formatiert": f"{isbn_prefix}-{e[9:12]}-{e[12]}",
        "shortcode": shortcode_aus_isbn(isbn13),
        "titel": titel,
        "untertitel": untertitel,
        "serientitel": serientitel,
        "band": band,
        "autoren": _namen(autoren),
        "herausgeber": _namen(herausgeber),
        "autoren_teile": autoren,          # inkl. Vor-/Nachname (für Autor-Link)
        "herausgeber_teile": herausgeber,
        "seiten": seiten,
        "einband": einband,
        "verlag": verlag,
        "preis_brutto": preis_brutto,
        "waehrung": waehrung,
        "datum": datum,
        "datum_iso": datum_iso,
        "cover_url": cover_url,
        "werbetext_absaetze": werbetext,
        **masse,
    }


# ---------------------------------------------------------------------
# Bilder (vom cover_previews-Tool auf dem Artikeldaten-Share)
# ---------------------------------------------------------------------

def artikeldaten_dir(cfg: dict) -> Path | None:
    raw = (cfg or {}).get("artikeldaten_dir") or ""
    if not raw:
        return None
    p = Path(os.path.expandvars(str(raw))).expanduser()
    try:
        return p if p.is_dir() else None
    except OSError:
        return None


def finde_artikel_ordner(sc: str, cfg: dict) -> Path | None:
    """Vorhandenen Ordner zum Kurzcode suchen ('05-597-4_Oberkirch')."""
    basis = artikeldaten_dir(cfg)
    if not basis or not sc:
        return None
    for p in sorted(basis.glob(f"{sc}*")):
        if p.is_dir():
            return p
    return None


def finde_bilder(sc: str, cfg: dict, ordner: Path | None = None) -> dict:
    """Sucht die von cover_previews erzeugten Bilder zum Kurzcode.

    Rückgabe: {"cover": Path|None, "galerie": [Path], "ordner": Path|None}
    Cover = 2D-Web-JPEG, Galerie = 3D-Mockup (falls vorhanden).
    """
    ordner = ordner or finde_artikel_ordner(sc, cfg)
    ergebnis = {"cover": None, "galerie": [], "ordner": ordner}
    if not ordner or not ordner.is_dir():
        return ergebnis

    m2d = cfg.get("muster_2d", DEFAULT_CONFIG["muster_2d"])
    m3d = cfg.get("muster_3d", DEFAULT_CONFIG["muster_3d"])
    dpi_web = int(cfg.get("dpi_web", 72))
    dpi_print = int(cfg.get("dpi_print", 300))

    cover = ordner / m2d.format(dpi=dpi_web, sc=sc)
    if not cover.exists():                       # Rückfall: Druckauflösung
        cover = ordner / m2d.format(dpi=dpi_print, sc=sc)
    if cover.exists():
        ergebnis["cover"] = cover

    for dpi in (dpi_web, dpi_print):             # 3D: Web bevorzugt
        p3 = ordner / m3d.format(dpi=dpi, sc=sc)
        if p3.exists():
            ergebnis["galerie"].append(p3)
            break
    return ergebnis


# ---------------------------------------------------------------------
# Produkt-Payload
# ---------------------------------------------------------------------

def produkt_id(isbn13: str) -> str:
    """Deterministische Shopware-ID (32 Hex) aus der ISBN — macht den Upsert
    idempotent: derselbe Titel landet immer auf demselben Produkt."""
    return hashlib.md5(re.sub(r"\D", "", isbn13).encode()).hexdigest()


def _media_id(isbn13: str, zweck: str) -> str:
    return hashlib.md5(f"{isbn13}:{zweck}".encode()).hexdigest()


def produkt_name(f: dict) -> str:
    name = f.get("titel", "").strip()
    if f.get("band"):
        name = f"{name} – Band {f['band']}"
    return name


def baue_beschreibung(f: dict) -> str:
    """Beschreibung im Hausstil des Verlags:

    Werbetext-Absätze (durch <br><br> getrennt), danach ggf. die Mitautoren
    („Mit Beiträgen von …“) und ein KURSIVER Fakten-Block mit Zeilenumbrüchen
    (<br>) — kein <p>, kein <hr>, keine '·'-Trennzeichen. So wie die bestehenden
    Produkte im Shop (siehe beispiel_produkt.json). Herausgeber/Autor stehen im
    eigenen „Autor“-Feld (customFields), nicht hier.
    """
    bloecke: list[str] = []
    absaetze = [html.escape(a) for a in f.get("werbetext_absaetze", []) if a.strip()]
    if absaetze:
        bloecke.append("<br>\n<br>\n".join(absaetze))

    # Mitautoren als eigener Absatz (die Herausgeber stehen im Autor-Feld)
    if f.get("autoren"):
        bloecke.append(html.escape(
            "Mit Beiträgen von " + _join_und(f["autoren"]) + "."))

    # Kursiver Fakten-Block: Umfang/Einband, Verlag + Jahr, ISBN + Preis
    zeilen: list[str] = []
    umfang = []
    if f.get("seiten"):
        umfang.append(f"{f['seiten']} Seiten")
    if f.get("einband"):
        umfang.append(f["einband"])
    if umfang:
        zeilen.append(", ".join(umfang) + ".")
    jahr = (f.get("datum") or "")[-4:]
    if f.get("verlag") and jahr:
        zeilen.append(f"{f['verlag']}. {jahr}.")
    elif f.get("verlag"):
        zeilen.append(f["verlag"] + ".")
    isbn_preis = ""
    if f.get("isbn13_formatiert"):
        isbn_preis = f"ISBN {f['isbn13_formatiert']}."
    if f.get("preis_brutto"):
        betrag = f"{f['preis_brutto']:.2f}".replace(".", ",")
        isbn_preis = (isbn_preis + " " if isbn_preis else "") \
            + f"{f.get('waehrung', 'EUR')} {betrag}."
    if isbn_preis:
        zeilen.append(isbn_preis.strip())
    if zeilen:
        bloecke.append("<i>" + "<br>\n".join(html.escape(z) for z in zeilen)
                       + "</i>")

    return "<br>\n<br>\n".join(bloecke)


def netto(brutto: float, satz: float) -> float:
    return round(float(brutto) / (1.0 + float(satz) / 100.0), 4)


def baue_customfields(f: dict, cfg: dict) -> dict:
    """SW5-Migrations-Felder, wie sie die bestehenden Produkte nutzen:
    attr1 = Untertitel, attr12 = Format + Einband, attr2 = Autor-Feld."""
    schluessel = cfg.get("custom_fields", DEFAULT_CONFIG["custom_fields"])
    cf: dict = {}

    # Untertitel: bei einer Reihe die Bandangabe ("Band 3"), sonst der
    # ONIX-Untertitel (auf Produktebene).
    untertitel = f"Band {f['band']}" if f.get("band") else f.get("untertitel", "")
    if untertitel and schluessel.get("untertitel"):
        cf[schluessel["untertitel"]] = untertitel

    # "24 x 17 cm, fester Einband"
    teile = []
    if f.get("hoehe_cm") and f.get("breite_cm"):
        teile.append(f"{_zahl_de(f['hoehe_cm'])} x {_zahl_de(f['breite_cm'])} cm")
    if f.get("einband"):
        teile.append(f["einband"])
    if teile and schluessel.get("format"):
        cf[schluessel["format"]] = ", ".join(teile)

    # Autor-Feld: die HERAUSGEBER (mit Link, " / "-getrennt, "(Hrsg.)") — so wie
    # im Bestand. Die eigentlichen Autoren sind nur Mitautoren und gehören hier
    # NICHT rein (die stehen in der Beschreibung). Ohne Herausgeber (rein
    # autorschaftliches Buch) steht hier der/die Autor(en).
    basis = cfg.get("autoren_basis_url", DEFAULT_CONFIG["autoren_basis_url"])
    if f.get("herausgeber_teile") and schluessel.get("autor_link"):
        links = " / ".join(autor_link(k, basis) for k in f["herausgeber_teile"])
        cf[schluessel["autor_link"]] = links + " (Hrsg.)"
    elif f.get("autoren_teile") and schluessel.get("autor_link"):
        cf[schluessel["autor_link"]] = ", ".join(
            autor_link(k, basis) for k in f["autoren_teile"])
    return cf


def baue_produkt(f: dict, cfg: dict, medien: list[dict] | None = None,
                 bestehende_id: str | None = None) -> dict:
    """Baut den Shopware-Produkt-Payload (upsert).

    ``medien``        = [{"media_id": .., "cover": bool}, ...] (schon hochgeladen)
    ``bestehende_id`` = ID eines bereits vorhandenen Produkts mit derselben
                        Artikelnummer. Muss übernommen werden — productNumber ist
                        eindeutig, ein neuer Datensatz würde sonst abgelehnt.
    """
    isbn = f["isbn13"]
    # Artikelnummer wie im Bestand: ISBN MIT Bindestrichen (978-3-95505-559-2)
    nummer = f.get("isbn13_formatiert") or isbn
    satz = float(cfg.get("tax_rate", 7.0))
    brutto = float(f.get("preis_brutto") or 0.0)

    payload = {
        "id": bestehende_id or produkt_id(isbn),
        "productNumber": nummer,
        "ean": nummer,
        "name": produkt_name(f),
        "active": bool(cfg.get("aktiv", False)),   # Entwurf: inaktiv
        "description": baue_beschreibung(f),
        "taxId": cfg.get("tax_id") or None,
        "price": [{
            "currencyId": cfg.get("currency_id") or None,
            "gross": round(brutto, 2),
            "net": netto(brutto, satz),
            "linked": True,        # Shopware rechnet netto selbst nach
        }],
        # Lieferbarkeit — Vorgabe des Verlags, für alle Bücher gleich.
        # Maximalabnahme und Lieferzeit bleiben leer (nicht mitschicken).
        "stock": int(cfg.get("default_stock", 9999)),
        "isCloseout": bool(cfg.get("is_closeout", True)),
        "minPurchase": int(cfg.get("min_purchase", 1)),
        "purchaseSteps": int(cfg.get("purchase_steps", 1)),
        "shippingFree": bool(cfg.get("shipping_free", False)),
    }

    # Wiederauffüllzeit: nur setzen, wenn konfiguriert (sonst leer lassen)
    if cfg.get("restock_time") is not None:
        payload["restockTime"] = int(cfg["restock_time"])

    if f.get("datum_iso"):
        payload["releaseDate"] = f["datum_iso"]

    # Maße/Gewicht aus der ONIX (Shopware: cm bzw. kg)
    for onix, sw in (("hoehe_cm", "height"), ("breite_cm", "width"),
                     ("dicke_cm", "length"), ("gewicht_kg", "weight")):
        if f.get(onix):
            payload[sw] = f[onix]

    # SEO
    payload["metaTitle"] = produkt_name(f)[:255]
    text = " ".join(f.get("werbetext_absaetze") or [])
    if text:
        payload["metaDescription"] = text[:255]
    stichworte = [k["nachname"] for k in (f.get("autoren_teile") or [])
                  if k.get("nachname")]
    stichworte += [k["nachname"] for k in (f.get("herausgeber_teile") or [])
                   if k.get("nachname")]
    if f.get("titel"):
        stichworte.append(f["titel"])
    if f.get("serientitel"):
        stichworte.append(f["serientitel"])
    if stichworte:
        payload["keywords"] = ", ".join(dict.fromkeys(stichworte))[:255]

    if cfg.get("manufacturer_id"):
        payload["manufacturerId"] = cfg["manufacturer_id"]
    if cfg.get("category_id"):
        payload["categories"] = [{"id": cfg["category_id"]}]
    if cfg.get("cms_page_id"):
        payload["cmsPageId"] = cfg["cms_page_id"]

    # Sichtbarkeit: OHNE diesen Eintrag taucht das Produkt im Shop NICHT auf —
    # auch nicht, wenn es aktiv geschaltet wird.
    if cfg.get("sales_channel_id"):
        payload["visibilities"] = [{
            "id": _media_id(isbn, "vis:" + cfg["sales_channel_id"]),
            "salesChannelId": cfg["sales_channel_id"],
            "visibility": int(cfg.get("visibility", 30)),
        }]

    cf = baue_customfields(f, cfg)
    if cf:
        payload["customFields"] = cf

    if medien:
        eintraege, cover_pmid = [], None
        for i, m in enumerate(medien):
            # coverId zeigt auf die product_media-Verknüpfung, NICHT auf media.id
            pmid = _media_id(isbn, f"pm{i}")
            eintraege.append({"id": pmid, "mediaId": m["media_id"],
                              "position": i})
            if m.get("cover") and cover_pmid is None:
                cover_pmid = pmid
        payload["media"] = eintraege
        payload["coverId"] = cover_pmid or eintraege[0]["id"]
    return payload


# ---------------------------------------------------------------------
# Shopware Admin-API-Client (nur stdlib)
# ---------------------------------------------------------------------

class ShopFehler(RuntimeError):
    pass


class ProduktExistiert(ShopFehler):
    """Das Buch ist im Shop schon vorhanden und darf nicht ohne ausdrückliche
    Bestätigung überschrieben werden (schützt gepflegte Bestandsdaten)."""

    def __init__(self, nummer: str, produkt_id: str):
        self.nummer = nummer
        self.produkt_id = produkt_id
        super().__init__(f"Produkt {nummer} existiert bereits.")


def normalisiere_url(url: str) -> str:
    """'dev.example.de/' -> 'https://dev.example.de'. Ohne Schema kann urllib
    die Adresse nicht auflösen ('unknown url type')."""
    u = (url or "").strip().rstrip("/")
    if u and not re.match(r"^https?://", u, re.IGNORECASE):
        u = "https://" + u
    return u


class ShopClient:
    """Minimaler Admin-API-Client (OAuth client_credentials + Sync + Medien)."""

    def __init__(self, shop_url: str, key: str, secret: str, timeout: float = 30.0,
                 tls_pruefen: bool = True):
        self.base = normalisiere_url(shop_url)
        self.key = key
        self.secret = secret
        self.timeout = timeout
        # Dev-Store hinter Caddy: dessen interne CA kennt Python nicht. Dann
        # kann die Zertifikatsprüfung hier abgeschaltet werden (nur für Dev!).
        self._ssl = None
        if not tls_pruefen:
            import ssl
            ctx = ssl.create_default_context()
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE
            self._ssl = ctx
        self._token: str | None = None

    # -- HTTP-Grundlagen ------------------------------------------------
    def _url(self, pfad: str) -> str:
        if not self.base:
            raise ShopFehler("Keine Shop-URL konfiguriert.")
        return f"{self.base}{pfad}"

    def _roh(self, method: str, pfad: str, body: bytes | None,
             content_type: str | None, auth: bool = True) -> bytes:
        req = urllib.request.Request(self._url(pfad), data=body, method=method)
        req.add_header("Accept", "application/json")
        if content_type:
            req.add_header("Content-Type", content_type)
        if auth:
            req.add_header("Authorization", f"Bearer {self.token()}")
        try:
            with urllib.request.urlopen(req, timeout=self.timeout,
                                        context=self._ssl) as r:
                return r.read()
        except urllib.error.HTTPError as e:
            # Basic-Auth vor dem Shop (Apache/nginx/Caddy) blockt die API: Basic
            # und Bearer teilen sich denselben Authorization-Header — beides
            # gleichzeitig geht nicht. Das muss im Webserver gelöst werden.
            wa = (e.headers.get("WWW-Authenticate") or "") if e.headers else ""
            if e.code == 401 and wa.lower().startswith("basic"):
                raise ShopFehler(
                    "Der Shop steht hinter einer Basic-Auth (Webserver) — die "
                    "Admin-API ist so nicht erreichbar: Basic und Bearer nutzen "
                    "beide den Authorization-Header.\n\n"
                    "Lösung: im Dev-vHost /api von der Basic-Auth ausnehmen, "
                    "am besten nur für die eigene IP (siehe README).") from e
            detail = e.read().decode("utf-8", "replace")[:800]
            raise ShopFehler(f"{method} {pfad} -> HTTP {e.code}\n{detail}") from e
        except urllib.error.URLError as e:
            grund = str(getattr(e, "reason", e))
            if "CERTIFICATE_VERIFY_FAILED" in grund:
                raise ShopFehler(
                    "TLS-Zertifikat nicht überprüfbar — bei einem Dev-Store "
                    "hinter Caddy ist das meist dessen interne CA.\n\n"
                    "Lösung: in der config.json \"tls_pruefen\": false setzen "
                    "(nur für den Dev-Store!) oder Caddys Root-CA im System "
                    "vertrauen.") from e
            raise ShopFehler(f"{method} {pfad} -> nicht erreichbar: {grund}") from e

    def _json(self, method: str, pfad: str, daten: dict | None = None,
              auth: bool = True) -> dict:
        body = json.dumps(daten).encode("utf-8") if daten is not None else None
        roh = self._roh(method, pfad, body,
                        "application/json" if body else None, auth=auth)
        if not roh:
            return {}
        try:
            return json.loads(roh)
        except json.JSONDecodeError:
            return {}

    # -- Auth -----------------------------------------------------------
    def token(self) -> str:
        """Access-Token per client_credentials holen.

        Shopware 6.7 hat league/oauth2-server angehoben und weist Anfragen ab,
        die nicht OAuth-konform sind. Konform ist laut RFC 6749
        `application/x-www-form-urlencoded` — deshalb wird so gesendet; ältere
        Stände, die nur JSON mochten, werden als Rückfall bedient.
        Ein `scope` wird bewusst nicht mitgeschickt: Shopware setzt dann den
        Standard-Scope der Integration.
        """
        if self._token:
            return self._token

        felder = {
            "grant_type": "client_credentials",
            "client_id": self.key,
            "client_secret": self.secret,
        }
        try:
            roh = self._roh("POST", "/api/oauth/token",
                            urllib.parse.urlencode(felder).encode("utf-8"),
                            "application/x-www-form-urlencoded", auth=False)
            antwort = json.loads(roh or b"{}")
        except ShopFehler:
            # Rückfall: JSON-Body (so machten es ältere Shopware-Stände)
            antwort = self._json("POST", "/api/oauth/token", felder, auth=False)

        tok = antwort.get("access_token")
        if not tok:
            raise ShopFehler("Kein access_token erhalten — Zugangsdaten prüfen.")
        self._token = tok
        return tok

    def verbinde(self) -> dict:
        """Token holen + Shop-Version lesen (Verbindungstest)."""
        self._token = None
        self.token()
        return self._json("GET", "/api/_info/version")

    # -- Nachschlagen (für die Zuordnungs-Auswahl im GUI) ---------------
    def _liste(self, pfad: str, limit: int = 100) -> list[dict]:
        antwort = self._json("GET", f"{pfad}?limit={limit}")
        return antwort.get("data", []) or []

    def steuersaetze(self) -> list[dict]:
        return self._liste("/api/tax")

    def waehrungen(self) -> list[dict]:
        return self._liste("/api/currency")

    def kategorien(self) -> list[dict]:
        return self._liste("/api/category", limit=500)

    def hersteller(self) -> list[dict]:
        return self._liste("/api/product-manufacturer")

    def sales_channels(self) -> list[dict]:
        return self._liste("/api/sales-channel")

    def produkt_id_zu_nummer(self, nummer: str) -> str | None:
        """ID eines bereits vorhandenen Produkts mit dieser Artikelnummer.

        productNumber ist eindeutig: gibt es das Buch schon (z. B. aus der
        SW5-Migration), MUSS dessen ID wiederverwendet werden — sonst lehnt
        Shopware den Upsert als Duplikat ab.
        """
        treffer = self.suche("product", {
            "limit": 1,
            "filter": [{"type": "equals", "field": "productNumber",
                        "value": nummer}],
        })
        return treffer[0]["id"] if treffer else None

    def vorlage_vom_bestand(self) -> dict:
        """Verkaufskanal, Seiten-Layout und Hersteller aus einem vorhandenen
        Produkt lesen.

        Alle drei sind im Verlagsshop für jedes Buch gleich (Hersteller =
        „Standard“) — es lohnt nicht, das von Hand zu pflegen. Ohne
        Verkaufskanal wäre ein neues Produkt im Shop sogar unsichtbar.
        """
        treffer = self.suche("product", {
            "limit": 1,
            "associations": {
                "visibilities": {"associations": {"salesChannel": {}}},
                "manufacturer": {},
            },
        })
        if not treffer:
            return {}
        p = treffer[0]
        sicht = (p.get("visibilities") or [{}])[0]
        kanal = sicht.get("salesChannel") or {}
        hersteller = p.get("manufacturer") or {}
        return {
            "cms_page_id": p.get("cmsPageId") or "",
            "sales_channel_id": sicht.get("salesChannelId") or "",
            "sales_channel_name": kanal.get("name") or "",
            "visibility": sicht.get("visibility") or 30,
            "manufacturer_id": p.get("manufacturerId") or "",
            "manufacturer_name": hersteller.get("name") or "",
        }

    # -- Medien ---------------------------------------------------------
    def produkt_medien_ordner(self) -> str | None:
        """ID des Standard-Medienordners für Produktbilder (für Thumbnails).
        Best effort — ohne Ordner funktioniert der Upload auch, dann werden
        aber keine Thumbnails erzeugt."""
        try:
            antwort = self._json(
                "GET", "/api/media-folder"
                       "?filter[defaultFolder.entity]=product&limit=1")
            daten = antwort.get("data") or []
            return daten[0]["id"] if daten else None
        except ShopFehler:
            return None

    def freier_dateiname(self, name: str, ext: str, media_id: str) -> str:
        """Kollisionsfreien Dateinamen besorgen. Ohne das schlägt der Upload
        fehl, wenn schon eine (andere) Datei so heißt."""
        try:
            q = urllib.parse.urlencode(
                {"fileName": name, "extension": ext, "mediaId": media_id})
            antwort = self._json("GET", f"/api/_action/media/provide-name?{q}")
            return antwort.get("fileName") or name
        except ShopFehler:
            return name

    def medium_hochladen(self, datei: Path, media_id: str, dateiname: str,
                         ordner_id: str | None = None) -> str:
        """Legt (idempotent) ein Medium an und lädt die Bilddatei hoch."""
        datei = Path(datei)
        ext = datei.suffix.lstrip(".").lower() or "jpg"
        mime = {"jpg": "image/jpeg", "jpeg": "image/jpeg",
                "png": "image/png"}.get(ext, "application/octet-stream")

        # 1) Medien-Datensatz anlegen (upsert über feste ID -> idempotent)
        eintrag: dict = {"id": media_id}
        if ordner_id:
            eintrag["mediaFolderId"] = ordner_id
        self.sync("media", [eintrag])

        # 2) Dateinamen absichern und Binärdaten hochladen
        name = self.freier_dateiname(dateiname, ext, media_id)
        q = urllib.parse.urlencode({"extension": ext, "fileName": name})
        self._roh("POST", f"/api/_action/media/{media_id}/upload?{q}",
                  datei.read_bytes(), mime)
        return media_id

    # -- Lesen ----------------------------------------------------------
    def suche(self, entity: str, koerper: dict) -> list[dict]:
        """POST /api/search/<entity> — erlaubt Filter + Assoziationen."""
        antwort = self._json("POST", f"/api/search/{entity}", koerper)
        return antwort.get("data") or []

    # -- Schreiben ------------------------------------------------------
    def sync(self, entity: str, payload: list[dict]) -> dict:
        return self._json("POST", "/api/_action/sync", {
            f"write-{entity}": {
                "entity": entity, "action": "upsert", "payload": payload,
            }
        })

    def admin_link(self, isbn13: str) -> str:
        return f"{self.base}/admin#/sw/product/detail/{produkt_id(isbn13)}"


# ---------------------------------------------------------------------
# Gesamtablauf
# ---------------------------------------------------------------------

def veroeffentliche(f: dict, cfg: dict, bilder: dict | None = None,
                    secret: str = "", dry_run: bool = False,
                    ueberschreiben: bool = False, log=print) -> dict:
    """Lädt die Bilder hoch und legt/aktualisiert das Produkt als Entwurf.

    ``secret`` ist das entschlüsselte Shopware-Secret — es kommt bewusst als
    Parameter und wird nie aus der Config gelesen (dort liegt nur Chiffretext).
    ``ueberschreiben`` muss True sein, um ein **bestehendes** Buch zu ändern;
    sonst wird ``ProduktExistiert`` geworfen, BEVOR irgendetwas gesendet oder
    ein Bild hochgeladen wird — so bleiben gepflegte Bestandsdaten geschützt.
    dry_run=True baut nur den Payload (nichts wird gesendet).
    Rückgabe: {"payload": .., "medien": [..], "admin_url": ..}
    """
    isbn = f["isbn13"]
    bilder = bilder or {"cover": None, "galerie": []}
    eff = effektiv(cfg)          # globale Einstellungen + aktive Umgebung

    dateien: list[tuple[Path, bool]] = []
    if bilder.get("cover"):
        dateien.append((Path(bilder["cover"]), True))
    for g in bilder.get("galerie", []):
        dateien.append((Path(g), False))

    if dry_run:
        medien = [{"media_id": _media_id(isbn, p.stem), "cover": ist_cover,
                   "datei": str(p)} for p, ist_cover in dateien]
        payload = baue_produkt(f, eff, medien)
        log(f"Dry-Run: Payload für {payload['productNumber']} gebaut "
            f"({len(medien)} Bild(er)) — nichts gesendet.")
        return {"payload": payload, "medien": medien, "admin_url": ""}

    if not secret:
        raise ShopFehler("Kein Secret entsperrt — bitte Master-Passwort eingeben.")
    client = ShopClient(eff.get("shop_url", ""), eff.get("access_key_id", ""),
                        secret, tls_pruefen=bool(eff.get("tls_pruefen", True)))
    fehlend = [k for k in ("tax_id", "currency_id") if not eff.get(k)]
    if fehlend:
        raise ShopFehler("Zuordnung fehlt: " + ", ".join(fehlend) +
                         " — bitte einmal 'Verbinden' und auswählen.")

    # Gibt es das Buch schon (z. B. aus der SW5-Migration)? Diese Prüfung läuft
    # ZUERST — noch vor dem Bild-Upload —, damit ein bestehendes Produkt ohne
    # ausdrückliche Freigabe garantiert unangetastet bleibt.
    nummer = f.get("isbn13_formatiert") or isbn
    bestehende_id = client.produkt_id_zu_nummer(nummer)
    if bestehende_id and not ueberschreiben:
        raise ProduktExistiert(nummer, bestehende_id)
    if bestehende_id:
        log(f"Vorhandenes Produkt {nummer} wird überschrieben (bestätigt).")

    ordner_id = client.produkt_medien_ordner() if dateien else None
    medien = []
    for p, ist_cover in dateien:
        mid = _media_id(isbn, p.stem)
        log(f"Lade Bild hoch: {p.name} …")
        client.medium_hochladen(p, mid, p.stem, ordner_id)
        medien.append({"media_id": mid, "cover": ist_cover, "datei": str(p)})

    payload = baue_produkt(f, eff, medien, bestehende_id)
    log(f"{'Aktualisiere' if bestehende_id else 'Lege an'}: {nummer} "
        f"({aktive_umgebung(cfg)}) …")
    client.sync("product", [payload])
    return {"payload": payload, "medien": medien, "neu": not bestehende_id,
            "admin_url": f"{client.base}/admin#/sw/product/detail/{payload['id']}"}
