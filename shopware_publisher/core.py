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
import datetime
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
# Wichtig: tax_id, currency_id, manufacturer_id & Co. sind **UUIDs des
# jeweiligen Shops** — dieselbe Kategorie hat auf dev und prod verschiedene IDs.
# Sie dürfen deshalb nicht global stehen, sonst schreibt man dev-IDs in den
# Produktivshop. Die Kategorien selbst stehen aus demselben Grund in gar keiner
# Config mehr: sie werden je Buch in der Oberfläche gewählt.
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
    "config_version": 3,

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
    "restock_time": 1,              # wie im Bestand (dort steht überall 1)
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
    # Zuordnungstabellen für das Kategorie-Raten. Beide nennen nur NAMEN —
    # ob es die Kategorie im Shop wirklich gibt, wird nachgesehen; eine
    # Kategorie erfindet das Werkzeug nie.
    #
    # Warengruppe (ONIX Schema 26) -> Sachkategorie. Verglichen wird der Teil
    # NACH dem " / ", also "Geschichte/Regionalgeschichte" statt
    # "Hardcover, Softcover / Geschichte/Regionalgeschichte".
    # VOM VERLAG ZU PRÜFEN — das ist ein Vorschlag, keine Setzung.
    "warengruppe_kategorien": {
        "Geschichte/Regionalgeschichte": ["Ortsgeschichte"],
        "Geschichte": ["Ortsgeschichte"],
        "Kunst/Architektur": ["Kunst und Musik"],
        "Kinder- und Jugendbücher": ["Kinder- u. Jugendliteratur"],
        "Reise": ["Reise-Kultur-Naturführer"],
        "Natur": ["Natur und Ökologie"],
        "Wirtschaft": ["Technik und Wirtschaft"],
    },
    # Ort (aus den freien Schlagwörtern, Schema 20) -> Region.
    # EBENFALLS ZU PRÜFEN.
    "ort_regionen": {
        "Mannheim": ["Kurpfalz"],
        "Heidelberg": ["Kurpfalz"],
        "Ludwigshafen": ["Kurpfalz"],
        "Speyer": ["Kurpfalz"],
        "Weinheim": ["Kurpfalz"],
        "Schwetzingen": ["Kurpfalz"],
        "Stuttgart": ["Württemberg"],
        "Ludwigsburg": ["Württemberg"],
        "Heilbronn": ["Württemberg"],
        "Tübingen": ["Württemberg"],
        "Karlsruhe": ["Karlsruhe-Region"],
        "Bruchsal": ["Karlsruhe-Region"],
        "Ettlingen": ["Karlsruhe-Region"],
        "Bretten": ["Karlsruhe-Region"],
    },

    # --- Bilder (kommen vom cover_previews-Tool auf dem Artikeldaten-Share) --
    "artikeldaten_dir": r"\\C019\d\Online\Webseite\Artikeldaten",
    "muster_2d": "2D_{dpi}_{sc}.jpg",
    "muster_3d": "3D_{dpi}_{sc}.jpg",
    "dpi_web": 72,
    "dpi_print": 300,
    # Dritte Bildquelle: die Cover, die der Newsletter nutzt, liegen offen auf
    # dem eigenen Webserver — volle Auflösung, ohne Share erreichbar.
    "newsletter_basis_url": "https://verlag-regionalkultur.de/newsletter_/",
    "cover_timeout": 20,

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

    # --- Gewichtsschätzung -------------------------------------------------
    # Wird beim Verbinden aus dem Bestand gelernt (siehe lerne_gewichtsmodell)
    # und hier abgelegt, damit die Schätzung auch ohne Shop-Zugriff steht.
    "gewichtsmodell": {},
    "gewichtsmodell_max_tage": 30,       # danach beim Verbinden neu lernen
}

# Schlüssel, die früher flach in der config.json standen (eine Umgebung).
_ALTE_FLACHE_SCHLUESSEL = tuple(DEFAULT_UMGEBUNG) + ("secret_access_key",)

# Verlags-Vorgaben (für alle Bücher gleich): werden bei einer Versionserhöhung
# aus DEFAULT_CONFIG aufgefrischt, damit alte config.json-Werte nicht per
# setdefault eingefroren bleiben (z. B. default_stock 0 -> 9999).
_VORGABE_SCHLUESSEL = ("aktiv", "default_stock", "is_closeout", "restock_time",
                       "min_purchase", "purchase_steps", "shipping_free",
                       "custom_fields", "autoren_basis_url",
                       "newsletter_basis_url", "cover_timeout")


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
        # Kategorien werden seit Version 3 je Buch gewählt. Der alte globale
        # Schlüssel muss weg, sonst bleibt in einer gewachsenen config.json eine
        # dev-Kategorie stehen, die niemand mehr sieht.
        for umg in (cfg.get("umgebungen") or {}).values():
            umg.pop("category_id", None)
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
        if _txt(c.find("b035")) != rolle:
            continue
        seq = _txt(c.find("b034"))
        # b047 = Körperschaft ("Stiftung Geißstraße"). Die hat keinen Vor- und
        # Nachnamen und darf vor allem NICHT am letzten Leerzeichen zerlegt
        # werden — sonst wird "Geißstraße" zum Nachnamen. Ohne diesen Zweig
        # fiel so ein Herausgeber bisher ersatzlos weg.
        koerperschaft = _txt(c.find("b047"))
        if koerperschaft:
            beitraege.append((int(seq) if seq.isdigit() else 999,
                              {"name": koerperschaft, "vorname": "",
                               "nachname": "", "koerperschaft": True}))
            continue
        name = _txt(c.find("b036"))
        vor, nach = _txt(c.find("b039")), _txt(c.find("b040"))
        if not name and nach:
            name = f"{vor} {nach}".strip()
        if not nach and name:               # Rückfall: am letzten Leerzeichen
            teile = name.rsplit(" ", 1)
            vor, nach = (teile[0], teile[1]) if len(teile) == 2 else ("", name)
        if name:
            beitraege.append((int(seq) if seq.isdigit() else 999,
                              {"name": name, "vorname": vor, "nachname": nach,
                               "koerperschaft": False}))
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

    Körperschaften stehen unter ihrem GANZEN Namen — auch die haben im Shop
    eine Autorenseite (/autoren-herausgeber/s/stiftung-geissstrasse/, geprüft).
    """
    name = k.get("name", "")
    nach, vor = k.get("nachname", ""), k.get("vorname", "")
    if nach:
        pfad = slug(f"{nach}-{vor}") if vor else slug(nach)
        initial = slug(nach)[:1]
    elif name:
        pfad = slug(name)
        initial = pfad[:1]
    else:
        return ""
    url = f"{basis.rstrip('/')}/{initial}/{pfad}/"
    return f"<a href='{url}'>{html.escape(name)}</a>"


# Satzende: Punkt/Ruf-/Fragezeichen vor einem Leerzeichen — aber NICHT nach
# einer Ziffer ("zum ausgehenden 18. Jahrhundert") und nicht nach einem
# einzelnen Buchstaben ("z. B."). Ohne diese zwei Ausnahmen endete die
# Meta-Beschreibung mitten in einer Jahrhundertangabe, was schlimmer aussieht
# als ein sauberes Auslassungszeichen.
_SATZENDE = re.compile(r"(?<![0-9])(?<!\s[A-Za-zÄÖÜäöü])[.!?](?=\s)")


def kuerze(text: str, grenze: int = 255) -> str:
    """Auf `grenze` Zeichen kürzen, ohne mitten im Wort abzubrechen.

    Shopware nimmt für metaTitle/metaDescription/keywords 255 Zeichen. Ein
    harter Schnitt endet mitten im Wort („… verfolgt auch kritisch die Ve“) —
    das steht so bei den migrierten Bestandsprodukten und sieht in der
    Google-Vorschau aus wie ein Fehler. Deshalb: erst am Satzende versuchen
    (dann steht dort ein ganzer Gedanke, ohne Auslassungszeichen), sonst an
    der letzten Wortgrenze mit „…“.
    """
    text = " ".join((text or "").split())       # Zeilenumbrüche raus
    if len(text) <= grenze:
        return text

    schnitt = text[:grenze - 1]                 # Platz für das „…“
    # Satzende nur nehmen, wenn danach noch genug Text steht — sonst bliebe
    # von einer langen ersten Passage bloss ein Halbsatz übrig.
    treffer = list(_SATZENDE.finditer(schnitt))
    if treffer and treffer[-1].start() >= grenze * 0.6:
        return schnitt[:treffer[-1].start() + 1]
    wort = schnitt.rfind(" ")
    # Auch den Punkt mit abschneiden — sonst steht dort "18.…".
    return (schnitt[:wort] if wort > 0 else schnitt).rstrip(" .,;:–-") + "…"


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


# VLB liefert dieselben Daten in zwei Schreibweisen: mit KURZ-Tags (<b012>) und
# mit REFERENZ-Tags (<ProductForm>). Welche man bekommt, haengt allein am
# Exportdialog — die Referenzfassung heisst dann "onix3Ref_….xml". Statt jeden
# Lesezugriff zu verdoppeln, wird der Baum einmalig auf die Kurzform gebracht;
# darunter bleibt der gesamte Lesecode unveraendert.
#
# Nur die Tags, die dieses Werkzeug wirklich liest — die ONIX-Liste ist ein
# Vielfaches davon. Wer hier ein Feld ergaenzt, traegt es auch hier ein.
REFERENZ_ZU_KURZ = {
    "Product": "product",
    "ProductIdentifier": "productidentifier",
    "ProductIDType": "b221",
    "IDValue": "b244",
    "DescriptiveDetail": "descriptivedetail",
    "TitleDetail": "titledetail",
    "TitleType": "b202",
    "TitleElement": "titleelement",
    "TitleElementLevel": "x409",
    "TitleText": "b203",
    "Subtitle": "b029",
    "Collection": "collection",
    "PartNumber": "x410",
    "Extent": "extent",
    "ExtentValue": "b219",
    "ProductForm": "b012",
    "NumberOfIllustrations": "b125",
    "IllustrationsNote": "b062",
    "PublishingDetail": "publishingdetail",
    "Publisher": "publisher",
    "PublisherName": "b081",
    "ProductSupply": "productsupply",
    "SupplyDetail": "supplydetail",
    "Price": "price",
    "Territory": "territory",
    "CountriesIncluded": "x449",
    "PriceAmount": "j151",
    "CurrencyCode": "j152",
    "PublishingDate": "publishingdate",
    "Date": "b306",
    "Measure": "measure",
    "MeasureType": "x315",
    "Measurement": "c094",
    "MeasureUnitCode": "c095",
    "CollateralDetail": "collateraldetail",
    "SupportingResource": "supportingresource",
    "ResourceContentType": "x436",
    "ResourceVersion": "resourceversion",
    "ResourceLink": "x435",
    "TextContent": "textcontent",
    "Text": "d104",
    "Subject": "subject",
    "SubjectSchemeIdentifier": "b067",
    "SubjectCode": "b069",
    "SubjectHeadingText": "b070",
    "Contributor": "contributor",
    "SequenceNumber": "b034",
    "ContributorRole": "b035",
    "PersonName": "b036",
    "CorporateName": "b047",
    "NamesBeforeKey": "b039",
    "KeyNames": "b040",
}


def _normalisiere_tags(root) -> None:
    """Referenzfassung auf Kurz-Tags umschreiben (siehe REFERENZ_ZU_KURZ).

    Namensraeume werden vorher abgestreift: die VLB-Dateien haben keinen, aber
    ONIX erlaubt ihn, und ein Namensraum würde jedes find() ins Leere laufen
    lassen — lautlos, was der schlimmste Fall waere.
    """
    for el in root.iter():
        if isinstance(el.tag, str):
            if "}" in el.tag:
                el.tag = el.tag.split("}", 1)[1]
            el.tag = REFERENZ_ZU_KURZ.get(el.tag, el.tag)


def lade_buchfelder(xml_pfad, cfg: dict | None = None) -> dict:
    """Liest die für den Shop nötigen Felder aus einer VLB-ONIX-XML.

    Bewusst schlanker als der pi_bi_generator: nur was das Produkt braucht.
    """
    cfg = cfg or DEFAULT_CONFIG
    einband_map = cfg.get("einband_map", DEFAULT_CONFIG["einband_map"])
    isbn_prefix = cfg.get("isbn_prefix", DEFAULT_CONFIG["isbn_prefix"])

    root = ET.parse(str(xml_pfad)).getroot()
    _normalisiere_tags(root)
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
    # "448 Seiten mit 390 Farb- und Schwarz-Weiß-Abbildungen, fester Einband."
    # — die Zahl (b125) und die Art (b062) stehen so im Shop und kommen beide
    # aus der ONIX. Ohne sie fehlt der Umfangzeile die Hälfte.
    abbildungen = _txt(dd.find("b125")) if dd is not None else ""
    abbildungsart = _txt(dd.find("b062")) if dd is not None else ""

    # Schlagwörter und Warengruppe. Schema 20 sind die freien Schlagwörter
    # ("Mannheim", "Kurpfalz", "Stadtgeschichte"), Schema 26 die
    # Warengruppen-Systematik ("Hardcover, Softcover / Geschichte").
    schlagworte: list[str] = []
    warengruppe = ""
    if dd is not None:
        for sub in dd.findall("subject"):
            schema = _txt(sub.find("b067"))
            text = _txt(sub.find("b070"))
            if not text:
                continue
            if schema == "20":
                schlagworte.append(text)
            elif schema == "26" and not warengruppe:
                warengruppe = text
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
        "abbildungen": abbildungen,
        "abbildungsart": abbildungsart,
        "schlagworte": schlagworte,
        "warengruppe": warengruppe,
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


def finde_bilder(sc: str, cfg: dict, ordner: Path | None = None,
                 xml_pfad=None) -> dict:
    """Sucht die Bilder zum Kurzcode — an mehreren Orten.

    Rückgabe: {"cover": Path|None, "galerie": [Path], "ordner": Path|None,
               "quelle": str}

    Reihenfolge, und warum sie so ist:

    1. **Artikeldaten-Share** — dort erzeugt ``cover_previews`` die Dateien.
       Das ist das Original, alles andere sind Kopien davon.
    2. **Neben der ONIX-Datei** — derselbe Griff, den der pi_bi_generator tut.
       Hilft, wenn der Share gerade nicht erreichbar ist, die Bilder aber
       schon lokal liegen.

    Die dritte Quelle (Webserver) ist bewusst NICHT hier: sie lädt herunter
    und gehört damit nicht in eine Funktion, die nur nachsieht. Dafür gibt es
    ``hole_cover_web``.

    ``quelle`` wird mitgegeben, damit die Vorschau sagen kann, WOHER das Bild
    kommt — bei drei möglichen Ablagen ist das keine Nebensache.
    """
    ergebnis = {"cover": None, "galerie": [], "ordner": None, "quelle": ""}
    if not sc:
        return ergebnis

    m2d = cfg.get("muster_2d", DEFAULT_CONFIG["muster_2d"])
    m3d = cfg.get("muster_3d", DEFAULT_CONFIG["muster_3d"])
    dpi_web = int(cfg.get("dpi_web", 72))
    dpi_print = int(cfg.get("dpi_print", 300))

    # -- Quelle 1: Artikeldaten-Share --------------------------------------
    ordner = ordner or finde_artikel_ordner(sc, cfg)
    if ordner and ordner.is_dir():
        ergebnis["ordner"] = ordner
        for dpi in (dpi_web, dpi_print):        # Web bevorzugt, Druck als Rückfall
            p = ordner / m2d.format(dpi=dpi, sc=sc)
            if p.exists():
                ergebnis["cover"] = p
                break
        for dpi in (dpi_web, dpi_print):
            p3 = ordner / m3d.format(dpi=dpi, sc=sc)
            if p3.exists():
                ergebnis["galerie"].append(p3)
                break
        if ergebnis["cover"]:
            ergebnis["quelle"] = "Artikeldaten-Share"
            return ergebnis

    # -- Quelle 2: neben der ONIX-Datei ------------------------------------
    if xml_pfad:
        nachbar = Path(xml_pfad).parent
        kandidaten = [m2d.format(dpi=dpi_print, sc=sc),
                      m2d.format(dpi=dpi_web, sc=sc),
                      f"{sc}.jpg", f"{sc}.png"]
        for name in kandidaten:
            p = nachbar / name
            if p.exists():
                ergebnis["cover"] = p
                ergebnis["ordner"] = ergebnis["ordner"] or nachbar
                ergebnis["quelle"] = "neben der ONIX-Datei"
                break
        if ergebnis["cover"] and not ergebnis["galerie"]:
            for dpi in (dpi_print, dpi_web):
                p3 = nachbar / m3d.format(dpi=dpi, sc=sc)
                if p3.exists():
                    ergebnis["galerie"].append(p3)
                    break

    return ergebnis


def cover_web_url(sc: str, cfg: dict) -> str:
    basis = cfg.get("newsletter_basis_url",
                    DEFAULT_CONFIG["newsletter_basis_url"])
    return f"{basis}{sc}.png"


def hole_cover_web(sc: str, cfg: dict) -> Path | None:
    """Dritte Quelle: das Cover vom eigenen Webserver holen.

    Dort liegt die Datei, die auch der Newsletter verwendet — volle Auflösung
    (rund 1500 x 2400 px), erreichbar ohne den Artikeldaten-Share. Sie wird in
    einen Zwischenspeicher neben dem Werkzeug gelegt, damit sie nicht bei
    jedem Blick erneut geladen wird.

    Rückgabe: Pfad zur Datei, oder None wenn es sie dort nicht gibt.
    """
    if not sc:
        return None
    cache = APP_DIR / "cover_cache"
    ziel = cache / f"{sc}.png"
    if ziel.exists() and ziel.stat().st_size > 0:
        return ziel
    url = cover_web_url(sc, cfg)
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "verlag-shopware-publisher"})
        with urllib.request.urlopen(
                req, timeout=float(cfg.get("cover_timeout", 20))) as r:
            daten = r.read()
    except Exception:
        return None                 # keine Quelle ist kein Fehler, nur ein Mangel
    if not daten:
        return None
    cache.mkdir(parents=True, exist_ok=True)
    ziel.write_bytes(daten)
    return ziel


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
    """Nur der Titel — ohne Bandangabe.

    Im Shop heißt der Band 14 der Schriftenreihe MARCHIVUM schlicht "Die GBG in
    Mannheim": weder "Band 14" noch der Reihentitel stehen irgendwo auf der
    Produktseite. Der Untertitel gehört ins Untertitelfeld und in die
    Zitatzeile der Beschreibung, die Reihe nirgendwohin.
    """
    return f.get("titel", "").strip()


def baue_beschreibung(f: dict) -> str:
    """Beschreibung im Hausstil des Verlags.

    Aufbau, abgelesen an den gepflegten Produkten im Livesystem:

        <Werbetext, Absätze durch <br><br> getrennt>

        <i>Heiko Brohm, Harald Stockert, Die GBG in Mannheim. 100 Jahre in 100 Geschichten.<br>
        448 Seiten mit 390 Farb- und Schwarz-Weiß-Abbildungen, fester Einband.<br>
        ISBN 978-3-95505-607-0. EUR 29,80.</i>

    Kein <p>, kein <hr>, keine '·'-Trennzeichen.

    Zwei Eigenheiten, die man sonst wieder "korrigiert":

    * Der EIGENE Verlag wird nicht genannt. Die Zeile "verlag regionalkultur.
      2026." steht bei keinem der aktuellen Produkte — sie taucht nur bei
      Fremdimprints auf ("Edition Guderjahn. 1997.").
    * Der Untertitel steht in der Zitatzeile, nicht im Namen. Die Reihe
      ("Schriftenreihe MARCHIVUM, Band 14") steht überhaupt nicht im Shop.
    """
    bloecke: list[str] = []
    absaetze = [html.escape(a) for a in f.get("werbetext_absaetze", []) if a.strip()]
    if absaetze:
        bloecke.append("<br>\n<br>\n".join(absaetze))

    # Mitautoren als eigener Absatz — die Herausgeber stehen in der Zitatzeile
    if f.get("autoren"):
        bloecke.append(html.escape(
            "Mit Beiträgen von " + _join_und(f["autoren"]) + "."))

    zeilen: list[str] = []

    # Zitatzeile: "Beitragende, Titel. Untertitel."
    beitragende = f.get("herausgeber") or f.get("autoren") or []
    zitat = ", ".join(beitragende)
    if f.get("titel"):
        zitat = (zitat + ", " if zitat else "") + f["titel"] + "."
    if f.get("untertitel"):
        zitat = (zitat + " " if zitat else "") + f["untertitel"] + "."
    if zitat:
        zeilen.append(zitat)

    # Umfangzeile: "448 Seiten mit 390 Farb- und Schwarz-Weiß-Abbildungen,
    # fester Einband." Ohne Abbildungsangabe fällt der mittlere Teil weg.
    umfang = ""
    if f.get("seiten"):
        umfang = f"{f['seiten']} Seiten"
        if f.get("abbildungen") and f.get("abbildungsart"):
            umfang += f" mit {f['abbildungen']} {f['abbildungsart']}"
    if f.get("einband"):
        umfang = (umfang + ", " if umfang else "") + f["einband"]
    if umfang:
        zeilen.append(umfang + ".")

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

    # Untertitel: IMMER der echte aus der ONIX. Früher stand hier bei
    # Reihenbänden "Band 5" — damit ging der Untertitel ersatzlos verloren
    # ("100 Jahre in 100 Geschichten"), und die Bandangabe steht im Shop
    # ohnehin nirgends. Das Theme zeigt dieses Feld unter der Überschrift.
    untertitel = f.get("untertitel", "")
    if untertitel and schluessel.get("untertitel"):
        cf[schluessel["untertitel"]] = untertitel

    # "17 x 23,5 cm, fester Einband" — BREITE mal HÖHE, so steht es im Shop
    # (die ONIX liefert beides getrennt: Höhe 23,5 / Breite 17). Die
    # Shopware-Felder height/width behalten davon unberührt ihre echte
    # Bedeutung; hier geht es nur um die Anzeigezeile.
    teile = []
    if f.get("hoehe_cm") and f.get("breite_cm"):
        teile.append(f"{_zahl_de(f['breite_cm'])} x {_zahl_de(f['hoehe_cm'])} cm")
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


def _kat_punkte(kat_name: str, person: dict) -> int:
    """Wie gut passt eine Kategorie zu einer Person? 2 = sicher, 1 = nur
    Nachname, 0 = gar nicht."""
    name = (kat_name or "").strip().lower()
    nach = (person.get("nachname") or "").strip().lower()
    vor = (person.get("vorname") or "").strip().lower()
    ganz = (person.get("name") or "").strip().lower()
    if not name:
        return 0
    if not nach:                                  # Körperschaft
        return 2 if ganz and ganz in name else 0
    if nach not in name:
        return 0
    return 2 if vor and vor in name else 1


def warengruppe_sachteil(text: str) -> str:
    """"Hardcover, Softcover / Kunst/Architektur" -> "Kunst/Architektur".

    Die Warengruppe nennt vorn den Einband und dahinter die Sachgruppe; nur
    letztere sagt etwas über den Inhalt. Getrennt wird am " / " MIT
    Leerzeichen — der Schrägstrich in "Geschichte/Regionalgeschichte" gehört
    zum Namen und darf nicht zerschnitten werden.
    """
    text = (text or "").strip()
    return text.split(" / ", 1)[1].strip() if " / " in text else text


def kategorien_cache_pfad(umgebung: str) -> Path:
    """Je Umgebung eine eigene Datei — die IDs von dev und prod sind
    verschieden, ein gemeinsamer Zwischenspeicher wäre Gift."""
    sicher = re.sub(r"[^0-9A-Za-z_-]", "_", umgebung or "unbekannt")
    return APP_DIR / f"kategorien_{sicher}.json"


def schreibe_kategorien_cache(umgebung: str, shop_url: str,
                              kategorien: list[dict]) -> Path:
    """Den Kategoriebaum neben der Config ablegen.

    Zweck ist nicht Geschwindigkeit, sondern **Nachvollziehbarkeit**: ohne
    Zugang zum Shop lässt sich sonst nicht prüfen, ob eine Kategorie fehlt
    oder nur nicht gefunden wurde.
    """
    ziel = kategorien_cache_pfad(umgebung)
    ziel.write_text(json.dumps({
        "umgebung": umgebung,
        "shop_url": shop_url,
        "geholt_am": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "anzahl": len(kategorien),
        "kategorien": sorted(kategorien, key=lambda k: k.get("name", "")),
    }, indent=2, ensure_ascii=False), encoding="utf-8")
    return ziel


def lade_kategorien_cache(umgebung: str) -> list[dict]:
    p = kategorien_cache_pfad(umgebung)
    if not p.is_file():
        return []
    try:
        return json.loads(p.read_text(encoding="utf-8")).get("kategorien") or []
    except (OSError, json.JSONDecodeError):
        return []


def cache_sucher(kategorien: list[dict]):
    """Eine `suche(text)`-Funktion, die im Zwischenspeicher nachsieht.

    Passt in dieselbe Stelle wie `ShopClient.kategorien_suchen` — das Raten
    merkt nicht, woher die Kandidaten kommen.
    """
    def suche(text: str) -> list[dict]:
        t = (text or "").strip().lower()
        return [k for k in kategorien if t and t in (k.get("name") or "").lower()]
    return suche


def kategorie_vorschlaege(f: dict, suche, cfg: dict | None = None,
                          log=None) -> tuple[list[dict], list[str]]:
    """Je beteiligter Person eine Kategorie vorschlagen.

    Der Shop führt **eine Kategorie pro Person**, benannt „Nachname, Vorname"
    (gemessen: Suche nach 'Wiegand' liefert „Wiegand, Anna", „Wiegand, Hermann",
    „Wiegand, Lutz"). Ein Buch mit vier Herausgebern gehört entsprechend in
    mehrere — es gibt KEINE Sammelkategorie, die alle vier nennt. (Der
    Breadcrumb im Shop zeigt zwar „Brohm / Stockert (Hrsg.)", das ist aber das
    Autorenfeld des Produkts, nicht der Kategoriename.)

    ``suche`` ist eine Funktion ``suche(text) -> list[dict]``; sie fragt den
    Shop. Der Kategoriebaum ist zu groß, um ihn zu laden.

    Rückgabe: (gefundene Kategorien, Namen ohne Treffer). Beides wird angezeigt
    — auch wer NICHT gefunden wurde, denn dann muss ein Mensch ran. Neue
    Kategorien legt das Werkzeug nie an.

    Mehrdeutiges wird bewusst nicht geraten: passen mehrere Kategorien nur über
    den Nachnamen, wird keine vorgeschlagen.
    """
    beteiligte = f.get("herausgeber_teile") or f.get("autoren_teile") or []
    treffer: list[dict] = []
    fehlend: list[str] = []
    gesehen: set[str] = set()

    def merke(kat: dict) -> None:
        if kat["id"] not in gesehen:
            gesehen.add(kat["id"])
            treffer.append(kat)

    for person in beteiligte:
        begriff = (person.get("nachname") or person.get("name") or "").strip()
        if not begriff:
            continue
        kandidaten = suche(begriff)
        # Körperschaften stehen mitunter unter einem Teil ihres Namens
        if not kandidaten and " " in begriff:
            kandidaten = suche(begriff.split()[-1])

        bewertet = [(_kat_punkte(k.get("name", ""), person), k) for k in kandidaten]
        bewertet = [(pkt, k) for pkt, k in bewertet if pkt > 0]
        if not bewertet:
            fehlend.append(person.get("name") or begriff)
            continue
        beste = max(pkt for pkt, _ in bewertet)
        gleichauf = [k for pkt, k in bewertet if pkt == beste]
        if beste == 1 and len(gleichauf) > 1:      # mehrdeutig -> nicht raten
            fehlend.append(person.get("name") or begriff)
            continue
        merke(gleichauf[0])

    # Sachkategorien. Gesucht wird nach NAMEN, und übernommen wird nur, was
    # im Shop wirklich so heißt — verglichen wird der GANZE Name, nicht ein
    # Teil davon. "Geschichte" steckt sonst gleichzeitig in "Ortsgeschichte",
    # "Kirchengeschichte" und "Jüdische Geschichte" und schaufelte drei
    # falsche Kategorien herein.
    cfg = cfg or DEFAULT_CONFIG
    namen: list[str] = []

    # 1. die freien Schlagwörter selbst ("Kurpfalz" IST eine Kategorie)
    schlagworte = [w.strip() for w in (f.get("schlagworte") or []) if w and w.strip()]
    namen += schlagworte

    # 2. Ort -> Region ("Stuttgart" ist keine Kategorie, "Württemberg" schon)
    orte = cfg.get("ort_regionen", DEFAULT_CONFIG["ort_regionen"])
    klein_orte = {k.lower(): v for k, v in orte.items()}
    for wort in schlagworte:
        namen += klein_orte.get(wort.lower(), [])

    # 3. Warengruppe -> Sachkategorie (hat jedes Buch, trifft also immer)
    wg = cfg.get("warengruppe_kategorien", DEFAULT_CONFIG["warengruppe_kategorien"])
    klein_wg = {k.lower(): v for k, v in wg.items()}
    # Die Warengruppe kann mehrere Sachgruppen nennen, durch Komma getrennt:
    # "Geschichte/Regionalgeschichte, Ländergeschichte". Deshalb erst das
    # Ganze versuchen, dann jeden Teil einzeln — sonst greift die Tabelle nur
    # bei den kurzen Angaben.
    sachteil = warengruppe_sachteil(f.get("warengruppe", ""))
    if sachteil:
        stuecke = [sachteil] + [t.strip() for t in sachteil.split(",")]
        for stueck in dict.fromkeys(t for t in stuecke if t):
            namen += klein_wg.get(stueck.lower(), [])

    gesucht = list(dict.fromkeys(n for n in namen if len(n) >= 3))
    ohne_treffer = []
    for name in gesucht:
        gefunden = False
        for kat in suche(name):
            if (kat.get("name") or "").strip().lower() == name.lower():
                merke(kat)
                gefunden = True
        if not gefunden:
            ohne_treffer.append(name)

    # Nachvollziehbar machen, WONACH gesucht wurde. Ohne das steht am Ende nur
    # "keine Kategorie" da, und niemand weiß, ob die Namen falsch waren oder
    # der Shop sie nicht führt.
    if log and gesucht:
        log(f"   Sachkategorien gesucht nach: {', '.join(gesucht)}")
        if ohne_treffer:
            log(f"   davon im Shop nicht vorhanden: {', '.join(ohne_treffer)}")

    return treffer, fehlend


# ---------------------------------------------------------------------
# Gewicht schätzen — gelernt an den Produkten, die schon im Shop stehen
# ---------------------------------------------------------------------
#
# Die VLB-ONIX liefert das Gewicht praktisch nie mit (ein <Measure> vom Typ 08
# fehlt). Im Shop ist es aber gepflegt: die Bestandsprodukte haben Gewicht,
# Breite und Höhe, und die Seitenzahl steht in der Umfangzeile ihrer
# Beschreibung ("448 Seiten mit 390 ..."). Damit lässt sich das Gewicht eines
# neuen Buchs aus dem eigenen Bestand ableiten, statt es zu raten.
#
# Geschätzt wird mit dem physikalischen Modell, nicht mit einem Mittelwert:
#
#     Gewicht = Fläche * (Papier_g_qm * Blattzahl + Einband_g_qm)
#
# Ein Buch ist ein Stapel Blätter (Seitenzahl / 2) eines Papiers mit festem
# Flächengewicht, dazu ein Einband, der ebenfalls mit der Fläche skaliert.
# Beide Größen sind ablesbare Zahlen in g/m² — geht eine Schätzung daneben,
# sieht man am Modell, woran es liegt. Ein bloßer Mittelwert oder ein
# Gramm-pro-Seite-Faktor könnte das nicht: er würde ein großformatiges
# dünnes Buch genauso schwer machen wie einen kleinen dicken Band.
#
# Gelernt wird je Einbandart getrennt — ein fester Einband wiegt bei gleicher
# Fläche ein Vielfaches einer Broschur.

# Grenzen, innerhalb derer ein Bestandsprodukt als Buch durchgeht. Im Shop
# stehen auch Nicht-Bücher (Karten, Poster) und Datensätze mit Tippfehlern;
# die dürfen das Modell nicht verziehen.
_GEW_MIN_KG, _GEW_MAX_KG = 0.02, 10.0
_FLAECHE_MIN_QM, _FLAECHE_MAX_QM = 0.005, 0.25      # ca. A6 bis doppelt A2
_SEITEN_MIN, _SEITEN_MAX = 8, 2000

# Ab so vielen Proben wird eine Einbandgruppe eigenständig geschätzt.
MIN_PROBEN = 8
# Schnitt der Gruppen bzw. Rechenweg; siehe gewichtsmodell_veraltet.
GEWICHTSMODELL_VERSION = 2
# Obergrenze für die Paarbildung (Theil-Sen ist quadratisch in der Punktzahl).
_MAX_PUNKTE = 200
# Ab dieser Abweichung gilt ein Bestandsgewicht als verdächtig: es wird beim
# Lernen weggelassen und in der Gegenprobe genannt. 40 % trennt am Bestand
# sauber zwischen Streuung (Papiersorte, Abbildungsteil) und Datenmüll.
_VERDACHT_AB = 0.40

_SEITEN_RE = re.compile(r"(\d{1,4})\s*Seiten", re.IGNORECASE)


def seiten_aus_beschreibung(text: str) -> int:
    """Seitenzahl aus der Umfangzeile eines Bestandsprodukts ("448 Seiten").

    Die Seitenzahl hat in Shopware kein eigenes Feld — sie steht nur im Text.
    Genommen wird der ERSTE Treffer: die Umfangzeile steht im Hausstil hinter
    dem Werbetext, im Werbetext selbst kommt "Seiten" kaum vor.
    """
    m = _SEITEN_RE.search(text or "")
    return int(m.group(1)) if m else 0


def _einband_aus_format(formatzeile: str) -> str:
    """Einbandart aus dem Format-Feld ("17 x 23,5 cm, fester Einband").

    Das ist dieselbe Zeile, die baue_customfields schreibt — der Einband steht
    darin hinter dem letzten Komma. Fehlt das Komma, ist keine Einbandart
    hinterlegt; dann zaehlt die Probe nur in die Gesamtgruppe.
    """
    text = (formatzeile or "").strip()
    if "," not in text:
        return ""
    return text.rsplit(",", 1)[1].strip()


# Einbandarten, wie sie im Shop stehen und wie die ONIX sie liefert, auf die
# zwei Klassen bringen, die fürs Gewicht zählen. Gemessen am Bestand: dieselbe
# Fläche und Blattzahl wiegt fest gebunden rund das Doppelte einer Broschur —
# das ist der grosse Unterschied, alles darunter geht im Rauschen unter.
#
# Nötig ist die Normalisierung, weil beide Seiten andere Wörter benutzen: die
# ONIX sagt "kartoniert" (b012 = BC), das Format-Feld im Shop sagt "Broschur",
# und daneben steht dort Freitext ("fester Einband im repräsentativen
# Großformat", "fester Einabnd", "Klappenbroschur mit Fadenheftung"). Ohne
# Normalisierung fand ein kartoniertes Buch gar keine Gruppe und wurde mit dem
# Mischmodell geschätzt — gemessen 16,4 % daneben statt 10,4 %.
_EINBAND_FEST = re.compile(
    r"fest|leinen|gebunden|hardcover|schuber|einabnd", re.IGNORECASE)
_EINBAND_WEICH = re.compile(
    r"broschur|kartoniert|paperback|flexib|geheftet|klapp", re.IGNORECASE)


def einband_klasse(text: str) -> str:
    """"fester Einband" / "Broschur" / "" (unbekannt)."""
    t = text or ""
    # Reihenfolge: "Klappenbroschur mit festem Rücken" gibt es nicht, aber
    # "fester Einband mit Klappen" schon — fest schlägt deshalb weich.
    if _EINBAND_FEST.search(t):
        return "fester Einband"
    if _EINBAND_WEICH.search(t):
        return "Broschur"
    return ""


def gewichtsprobe(p: dict, cfg: dict) -> dict | None:
    """Ein Bestandsprodukt in eine Lernprobe umrechnen — oder None."""
    def zahl(x) -> float:
        try:
            return float(x)
        except (TypeError, ValueError):
            return 0.0

    uebersetzt = p.get("translated") or {}
    gewicht = zahl(p.get("weight"))
    breite, hoehe = zahl(p.get("width")), zahl(p.get("height"))
    # Aus der SW5-Migration stehen bei einzelnen Produkten Millimeter in den
    # cm-Feldern. Ein 100 cm breites Buch gibt es nicht — dann sind es mm.
    if breite > 100:
        breite /= 10.0
    if hoehe > 100:
        hoehe /= 10.0

    seiten = seiten_aus_beschreibung(
        p.get("description") or uebersetzt.get("description") or "")
    flaeche = breite * hoehe / 10000.0          # cm² -> m²

    if not (_GEW_MIN_KG <= gewicht <= _GEW_MAX_KG):
        return None
    if not (_FLAECHE_MIN_QM <= flaeche <= _FLAECHE_MAX_QM):
        return None
    if not (_SEITEN_MIN <= seiten <= _SEITEN_MAX):
        return None

    cf = p.get("customFields") or uebersetzt.get("customFields") or {}
    schluessel = (cfg.get("custom_fields")
                  or DEFAULT_CONFIG["custom_fields"]).get("format")
    einband_roh = _einband_aus_format(cf.get(schluessel, "") if schluessel else "")

    # Das echte Gewicht und die Maße bleiben dabei — nur so lässt sich die
    # Schätzung hinterher gegen den Bestand halten (kreuzvalidierung).
    return {
        "blatt": seiten / 2.0,
        "g_qm": gewicht * 1000.0 / flaeche,     # Gewicht je Quadratmeter
        "einband": einband_klasse(einband_roh),
        "einband_roh": einband_roh,
        "nummer": p.get("productNumber") or "",
        "seiten": seiten,
        "breite_cm": round(breite, 1),
        "hoehe_cm": round(hoehe, 1),
        "flaeche_qm": flaeche,
        "gewicht_kg": gewicht,
    }


def _median(werte: list[float]) -> float:
    w = sorted(werte)
    n = len(w)
    if not n:
        return 0.0
    return w[n // 2] if n % 2 else (w[n // 2 - 1] + w[n // 2]) / 2.0


def _theil_sen(punkte: list[tuple[float, float]]) -> tuple[float, float]:
    """Robuste Gerade y = a*x + b als Median aller Paar-Steigungen.

    Kleinste Quadrate wären hier falsch: ein einziges Produkt mit falsch
    gepflegtem Gewicht (Gramm statt Kilogramm) oder ein Poster, das durch die
    Plausibilitätsgrenzen gerutscht ist, zöge die Gerade sichtbar weg. Der
    Median verträgt rund ein Viertel solcher Ausreißer, ohne zu kippen.
    """
    steigungen = []
    for i, (x1, y1) in enumerate(punkte):
        for x2, y2 in punkte[i + 1:]:
            if x1 != x2:
                steigungen.append((y2 - y1) / (x2 - x1))
    if not steigungen:
        return 0.0, _median([y for _, y in punkte])
    a = _median(steigungen)
    return a, _median([y - a * x for x, y in punkte])


def _gerade(proben: list[dict]) -> tuple[float, float] | None:
    """Ein Durchgang Theil-Sen über die Proben einer Gruppe."""
    punkte = sorted((pr["blatt"], pr["g_qm"]) for pr in proben)
    if len(punkte) > _MAX_PUNKTE:
        # Gleichmäßig ausdünnen — sortiert, damit der ganze Bereich von
        # dünn bis dick erhalten bleibt und nicht nur ein Ausschnitt.
        schritt = len(punkte) / _MAX_PUNKTE
        punkte = [punkte[int(i * schritt)] for i in range(_MAX_PUNKTE)]

    papier, einband = _theil_sen(punkte)
    if papier <= 0:
        return None             # ohne Papiergewicht ist das Modell wertlos
    # Eine leicht negative Konstante ist Rauschen bei Broschuren (der Umschlag
    # steckt dann im Papierwert). Physikalisch kann sie nicht negativ sein.
    return papier, max(einband, 0.0)


def _abweichungen(proben: list[dict], papier: float, einband: float) -> list[float]:
    return [abs(papier * pr["blatt"] + einband - pr["g_qm"]) / pr["g_qm"]
            for pr in proben]


def _schaetze_gruppe(proben: list[dict]) -> dict | None:
    """Papier- und Einbandgewicht (g/m²) für eine Gruppe von Proben.

    Zwei Durchgänge: der erste dient nur dazu, die Datensätze zu finden, deren
    Gewicht im Shop nicht stimmen KANN (gemessen: 51 von 459 liegen über 40 %
    daneben — "0,200 kg" bei 132 Seiten in 22 x 28 cm gibt es nicht, das ist
    ein Platzhalter). Der zweite Durchgang lernt ohne sie. Theil-Sen verträgt
    zwar Ausreißer, aber es gibt keinen Grund, sie mitzuschleppen.
    """
    erste = _gerade(proben)
    if not erste:
        return None

    behalten = [pr for pr, ab in zip(proben, _abweichungen(proben, *erste))
                if ab <= _VERDACHT_AB]
    verworfen = len(proben) - len(behalten)
    if len(behalten) >= MIN_PROBEN:
        zweite = _gerade(behalten) or erste
    else:
        behalten, zweite, verworfen = proben, erste, 0

    papier, einband = zweite
    return {
        "papier_g_qm": round(papier, 1),
        "einband_g_qm": round(einband, 1),
        "proben": len(behalten),
        "verworfen": verworfen,
        # Streuung an den Proben, aus denen gelernt wurde. lerne_gewichte_vom_shop
        # ersetzt sie durch den ehrlichen, kreuzvalidierten Wert.
        "abweichung_prozent": round(
            100.0 * _median(_abweichungen(behalten, papier, einband)), 1),
    }


def lerne_gewichtsmodell(proben: list[dict],
                         min_proben: int = MIN_PROBEN) -> dict:
    """Aus Bestandsproben ein Gewichtsmodell je Einbandart bauen.

    Die Gruppe "*" (alle Einbandarten zusammen) ist der Rückfall für Bücher,
    deren Einbandart im Bestand zu selten vorkommt.
    """
    gruppen: dict[str, list[dict]] = {"*": list(proben)}
    for pr in proben:
        if pr.get("einband"):
            gruppen.setdefault(pr["einband"], []).append(pr)

    ergebnis = {}
    for name, teil in gruppen.items():
        if len(teil) < min_proben:
            continue
        modell = _schaetze_gruppe(teil)
        if modell:
            ergebnis[name] = modell

    return {
        # Die Version steigt, wenn sich Gruppierung oder Rechnung ändern —
        # ein gespeichertes Modell aus der Zeit davor wäre sonst still weiter
        # in Gebrauch, obwohl seine Gruppen anders geschnitten sind.
        "version": GEWICHTSMODELL_VERSION,
        "stand": datetime.date.today().isoformat(),
        "proben": len(proben),
        "gruppen": ergebnis,
    }


def gewichtsmodell_veraltet(cfg: dict) -> bool:
    """Fehlt das Modell oder ist es älter als `gewichtsmodell_max_tage`?

    Der Bestand wächst; ein Modell von vor einem Jahr kennt die zuletzt
    gepflegten Bücher nicht. Neu gelernt wird beim Verbinden, weil dann
    ohnehin eine Verbindung steht.
    """
    modell = cfg.get("gewichtsmodell") or {}
    if not modell.get("gruppen"):
        return True
    if modell.get("version") != GEWICHTSMODELL_VERSION:
        return True
    try:
        stand = datetime.date.fromisoformat(modell.get("stand", ""))
    except ValueError:
        return True
    tage = int(cfg.get("gewichtsmodell_max_tage",
                       DEFAULT_CONFIG["gewichtsmodell_max_tage"]))
    return (datetime.date.today() - stand).days > tage


def _gruppe_fuer(gruppen: dict, einband: str) -> tuple[dict | None, str]:
    """Passende Einbandgruppe — sonst die Gruppe über alle Einbandarten."""
    if gruppen.get(einband):
        return gruppen[einband], einband
    return gruppen.get("*"), "alle Einbandarten"


def kreuzvalidierung(proben: list[dict], faltungen: int = 5) -> dict:
    """Die Schätzung gegen die echten Gewichte im Bestand halten.

    Ein Modell an denselben Produkten zu messen, aus denen es gelernt hat,
    schmeichelt sich selbst. Deshalb k-fache Kreuzvalidierung: das Modell
    lernt an vier Fünfteln des Bestands und schätzt das letzte Fünftel, das es
    nie gesehen hat — so wie es später ein neues Buch schätzt.

    Zurück kommen die Abweichungen (Median, 90 %-Wert, Trefferquoten) je
    Einbandart und die größten Ausreißer. Die Ausreißer sind doppelt nützlich:
    sie zeigen entweder die Grenzen des Modells oder ein Produkt, dessen
    Gewicht im Shop falsch gepflegt ist.
    """
    # Fest nach Artikelnummer aufteilen — derselbe Bestand ergibt dieselben
    # Zahlen, sonst wäre nicht zu sehen, ob eine Änderung etwas gebracht hat.
    sortiert = sorted(proben, key=lambda pr: pr.get("nummer", ""))
    einzeln: list[dict] = []
    for i in range(max(2, faltungen)):
        test = sortiert[i::max(2, faltungen)]
        lern = [pr for j, pr in enumerate(sortiert)
                if j % max(2, faltungen) != i]
        gruppen = (lerne_gewichtsmodell(lern) or {}).get("gruppen") or {}
        for pr in test:
            gruppe, quelle = _gruppe_fuer(gruppen, pr.get("einband") or "")
            if not gruppe:
                continue
            kg = pr["flaeche_qm"] * (gruppe["papier_g_qm"] * pr["blatt"]
                                     + gruppe["einband_g_qm"]) / 1000.0
            einzeln.append({
                **pr,
                "geschaetzt_kg": round(kg, 3),
                "gruppe": quelle,
                "abweichung": (kg - pr["gewicht_kg"]) / pr["gewicht_kg"],
            })

    def kennzahlen(teil: list[dict]) -> dict:
        if not teil:
            return {}
        rel = sorted(abs(e["abweichung"]) for e in teil)
        return {
            "n": len(teil),
            "median_prozent": round(100 * _median(rel), 1),
            "p90_prozent": round(100 * rel[int(0.9 * (len(rel) - 1))], 1),
            "im_10er_band": round(100 * sum(r <= 0.10 for r in rel) / len(rel)),
            "im_20er_band": round(100 * sum(r <= 0.20 for r in rel) / len(rel)),
            # Vorzeichen: schätzt das Modell im Schnitt zu schwer oder zu leicht?
            "schlagseite_prozent": round(
                100 * _median([e["abweichung"] for e in teil]), 1),
        }

    je_einband: dict[str, dict] = {}
    for e in einzeln:
        je_einband.setdefault(e.get("einband") or "— ohne Angabe —", []).append(e)

    return {
        "gesamt": kennzahlen(einzeln),
        "je_einband": {k: kennzahlen(v) for k, v in sorted(je_einband.items())},
        "einzeln": sorted(einzeln, key=lambda e: -abs(e["abweichung"])),
    }


def lerne_und_pruefe(proben: list[dict]) -> tuple[dict, dict]:
    """Modell lernen UND gleich gegen den Bestand halten.

    Die Streuung, die beim Lernen anfällt, schmeichelt sich selbst — das
    Modell kennt diese Bücher ja. Im Modell steht deshalb der Wert aus der
    Kreuzvalidierung: so weit lag es bei Büchern daneben, die es NICHT kannte.
    Das ist die Zahl, die für ein neues Buch gilt, und die deshalb auch im
    Werkzeug am Gewichtsfeld steht.
    """
    modell = lerne_gewichtsmodell(proben)
    if not modell.get("gruppen"):
        return modell, {"gesamt": {}, "je_einband": {}, "einzeln": []}

    pruefung = kreuzvalidierung(proben)
    for name, gruppe in modell["gruppen"].items():
        gemessen = (pruefung["gesamt"] if name == "*"
                    else pruefung["je_einband"].get(name))
        if gemessen:
            gruppe["abweichung_prozent"] = gemessen["median_prozent"]
            gruppe["gemessen"] = True
    modell["verdaechtig"] = sum(1 for e in pruefung["einzeln"]
                                if abs(e["abweichung"]) > _VERDACHT_AB)
    return modell, pruefung


def schaetze_gewicht(f: dict, modell: dict | None) -> tuple[float, str]:
    """Gewicht in kg schätzen. Gibt (0.0, Grund) zurück, wenn es nicht geht.

    Der Grund wird angezeigt: eine stille 0 wäre schlimmer als gar keine
    Schätzung — dann stünde im Shop ein falsches Versandgewicht.
    """
    gruppen = ((modell or {}).get("gruppen") or {})
    if not gruppen:
        # Nicht "beim Verbinden" sagen: der Buchdurchgang kennt keinen
        # Verbinden-Knopf, dort wird beim ersten Anlegen gelernt.
        return 0.0, "noch kein Modell — wird aus dem Bestand gelernt, sobald " \
                    "eine Verbindung zum Shop steht"

    try:
        seiten = int(str(f.get("seiten") or "0").strip() or 0)
    except ValueError:
        seiten = 0
    breite = float(f.get("breite_cm") or 0.0)
    hoehe = float(f.get("hoehe_cm") or 0.0)
    if not seiten:
        return 0.0, "keine Seitenzahl in der ONIX"
    if not (breite and hoehe):
        return 0.0, "kein Format (Breite/Höhe) in der ONIX"

    gruppe, quelle = _gruppe_fuer(gruppen, einband_klasse(f.get("einband")))
    if not gruppe:
        return 0.0, "keine passende Vergleichsgruppe im Bestand"

    flaeche = breite * hoehe / 10000.0
    kg = flaeche * (gruppe["papier_g_qm"] * seiten / 2.0
                    + gruppe["einband_g_qm"]) / 1000.0
    if not (_GEW_MIN_KG <= kg <= _GEW_MAX_KG):
        return 0.0, f"unplausibles Ergebnis ({kg:.2f} kg)"

    # "typisch daneben" ist der kreuzvalidierte Median: so weit lag die
    # Schätzung bei Büchern daneben, die das Modell nicht kannte. Ohne diese
    # Angabe liest sich eine Schätzung wie eine Messung.
    return round(kg, 3), (
        f"geschätzt aus {gruppe['proben']} vergleichbaren Produkten ({quelle})"
        f" — typisch ±{_zahl_de(gruppe['abweichung_prozent'])} % daneben")


def baue_produkt(f: dict, cfg: dict, medien: list[dict] | None = None,
                 bestehende_id: str | None = None,
                 kategorien: list[str] | None = None) -> dict:
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
    # metaTitle bleibt der blosse Titel — so steht es im Bestand, auch bei
    # Büchern mit Untertitel. Der Untertitel steht im eigenen Feld (attr1).
    payload["metaTitle"] = kuerze(produkt_name(f))
    text = " ".join(f.get("werbetext_absaetze") or [])
    if text:
        payload["metaDescription"] = kuerze(text)
    # Schlagwort je Beteiligtem ist der NACHNAME (wie im Bestand: "Brandes,
    # Kinderbuch, Grünes Gras erzähl mir was, Dilsberg"). Eine Körperschaft
    # („Stiftung Geißstraße“) hat keinen Nachnamen — dann der ganze Name,
    # sonst fiel sie hier ersatzlos weg und es blieb nur der Titel stehen.
    def _stichwort(k: dict) -> str:
        return k.get("nachname") or k.get("name") or ""

    stichworte = [_stichwort(k) for k in (f.get("autoren_teile") or [])]
    stichworte += [_stichwort(k) for k in (f.get("herausgeber_teile") or [])]
    stichworte = [x for x in stichworte if x]
    if f.get("titel"):
        stichworte.append(f["titel"])
    if f.get("serientitel"):
        stichworte.append(f["serientitel"])
    if stichworte:
        payload["keywords"] = kuerze(", ".join(dict.fromkeys(stichworte)))

    # Such-Schlagwörter — das Feld, das im Admin unter „Kategorie" steht.
    # NICHT zu verwechseln mit `keywords` oben: das sind die SEO-Meta-Wörter
    # auf einem anderen Reiter. Genau diese Verwechslung führte dazu, dass die
    # Maske leer blieb, obwohl „keywords" gefüllt war.
    #
    # Im Bestand steht darin (Beispiel „Grünes Gras"): Titel, Untertitel, der
    # Beteiligte mit vollem Namen, die ISBN — dazu von Hand Sachbegriffe
    # („Kinder", „Jugendliche"), die aus der ONIX nicht abzuleiten sind und
    # deshalb hier auch nicht erfunden werden. Der Reihentitel kommt dazu:
    # wer „Schriftenreihe MARCHIVUM" sucht, soll Band 14 finden.
    suchworte = [f.get("titel"), f.get("untertitel"), f.get("serientitel")]
    suchworte += [k.get("name") for k in (f.get("herausgeber_teile") or [])]
    suchworte += [k.get("name") for k in (f.get("autoren_teile") or [])]
    suchworte += list(f.get("schlagworte") or [])   # aus der ONIX, Schema 20
    suchworte.append(f.get("isbn13_formatiert"))
    suchworte = [t.strip() for t in suchworte if t and str(t).strip()]
    if suchworte:
        payload["customSearchKeywords"] = list(dict.fromkeys(suchworte))

    if cfg.get("manufacturer_id"):
        payload["manufacturerId"] = cfg["manufacturer_id"]
    # Kategorien kommen je Buch aus der Oberfläche — nicht mehr eine globale
    # aus der Config. Ohne Kategorie hat das Produkt im Shop keinen Breadcrumb.
    if kategorien:
        payload["categories"] = [{"id": k} for k in dict.fromkeys(kategorien)]
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
        """Die ersten 500 Kategorien — NUR als grobe Übersicht.

        Zum Suchen taugt das nicht: der Shop führt je Autorenkonstellation eine
        eigene Kategorie, das sind weit mehr als 500. Wer hier filtert, filtert
        auf einem willkürlichen Ausschnitt. Dafür ``kategorien_suchen``.
        """
        return self._liste("/api/category", limit=500)

    def alle_kategorien(self, log=None) -> list[dict]:
        """Den GANZEN Kategoriebaum holen, seitenweise.

        `/api/category?limit=500` gibt genau 500 zurück und verschweigt, dass
        es mehr gibt — deshalb wird geblättert, bis eine Seite nicht mehr voll
        ist. Zurück kommen nur id und name; mehr braucht das Raten nicht, und
        die Datei bleibt klein.
        """
        alle: dict[str, str] = {}
        seite, pro_seite = 1, 500
        while True:
            antwort = self._json(
                "POST", "/api/search/category",
                {"limit": pro_seite, "page": seite,
                 "sort": [{"field": "name", "order": "ASC"}]})
            daten = antwort.get("data") or []
            for k in daten:
                if k.get("id"):
                    alle[k["id"]] = (k.get("name") or "").strip()
            if log:
                log(f"   Kategorien geholt: {len(alle)}")
            if len(daten) < pro_seite:
                break
            seite += 1
            if seite > 40:              # Notbremse, 20.000 wären absurd
                break
        return [{"id": i, "name": n} for i, n in alle.items() if n]

    def kategorien_suchen(self, text: str, limit: int = 50) -> list[dict]:
        """Kategorien im Shop suchen — server-seitig.

        Gemessen am Livesystem: ``/api/category`` gibt mit limit=500 genau 500
        zurück, also abgeschnitten. Alles zu laden und örtlich zu filtern
        verfehlt damit zuverlässig genau die Kategorie, die man sucht. Deshalb
        sucht der Shop selbst.
        """
        koerper: dict = {"limit": limit,
                         "sort": [{"field": "name", "order": "ASC"}]}
        if text:
            koerper["filter"] = [{"type": "contains", "field": "name",
                                  "value": text}]
        return self.suche("category", koerper)

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

    def bestand_gewichte(self, max_produkte: int = 2000) -> list[dict]:
        """Gepflegte Produkte mit Gewicht und Maßen holen (Lernstoff).

        Nur die vier Felder, die die Schätzung braucht — `includes` hält die
        Antwort klein, sonst kaeme der halbe Produktkatalog mit Preisen und
        Medien über die Leitung. Gefiltert wird server-seitig auf Werte > 0;
        das schließt die vielen Produkte ohne gepflegtes Gewicht aus.
        """
        raus: list[dict] = []
        seite = 1
        limit = 500
        while len(raus) < max_produkte:
            treffer = self.suche("product", {
                "limit": limit,
                "page": seite,
                "filter": [
                    {"type": "range", "field": "weight",
                     "parameters": {"gt": 0}},
                    {"type": "range", "field": "width",
                     "parameters": {"gt": 0}},
                    {"type": "range", "field": "height",
                     "parameters": {"gt": 0}},
                ],
                "includes": {"product": ["productNumber", "weight", "width",
                                         "height", "description",
                                         "customFields"]},
            })
            raus += treffer
            if len(treffer) < limit:
                break
            seite += 1
        return raus[:max_produkte]

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

def lerne_gewichte_vom_shop(client: "ShopClient", cfg: dict) -> dict:
    """Bestand holen, Modell lernen und in der Config ablegen.

    Die Config ist der richtige Ort: so steht die Schätzung auch beim
    nächsten Start bereit, ohne dass wieder tausend Produkte geladen werden.
    """
    proben = [gewichtsprobe(p, cfg) for p in client.bestand_gewichte()]
    modell, _ = lerne_und_pruefe([p for p in proben if p])
    cfg["gewichtsmodell"] = modell
    return modell


def _gewicht_ergaenzen(f: dict, cfg: dict, client: "ShopClient | None" = None,
                       log=print) -> dict:
    """Fehlendes Gewicht schätzen — notfalls erst das Modell dafür lernen.

    Hier und nicht nur in der Oberfläche, damit JEDER Weg in den Shop ein
    Versandgewicht bekommt: der Buchdurchgang ruft ``veroeffentliche`` direkt
    auf und hat eine eigene config.json, in der nie jemand „Verbinden“ gedrückt
    hat. Ohne das eigene Nachlernen bliebe dort für immer „kein Modell“.

    Der übergebene Datensatz bleibt unverändert — der Aufrufer soll keine Zahl
    untergeschoben bekommen, die er nie gesehen hat. Das Modell dagegen landet
    in ``cfg``: der Aufrufer speichert seine Konfiguration ohnehin, und beim
    nächsten Buch steht es dann schon da.
    """
    if f.get("gewicht_kg"):
        return f

    if client is not None and gewichtsmodell_veraltet(cfg):
        try:
            log("Lerne Gewichte aus dem Bestand (einmalig) …")
            lerne_gewichte_vom_shop(client, cfg)
        except Exception as fehler:
            # Ein Buch ohne Gewicht ist ärgerlich, ein abgebrochener Upload
            # schlimmer. Weiter geht es in jedem Fall.
            log(f"Gewichtsmodell nicht gelernt: {fehler}")

    kg, grund = schaetze_gewicht(f, cfg.get("gewichtsmodell"))
    if kg:
        log(f"Gewicht geschätzt: {_zahl_de(kg)} kg — {grund}")
        return {**f, "gewicht_kg": kg}
    log(f"Ohne Gewicht: {grund}")
    return f


def veroeffentliche(f: dict, cfg: dict, bilder: dict | None = None,
                    secret: str = "", dry_run: bool = False,
                    ueberschreiben: bool = False, log=print,
                    kategorien: list[str] | None = None) -> dict:
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


    # (Datei, Rolle). Die ROLLE bestimmt die Medien-ID, nicht der Dateiname:
    # dasselbe Cover heisst je nach Quelle anders ("2D_72_05-607-0" vom Share,
    # "05-607-0" vom Webserver). Haengt die ID am Namen, entsteht bei jedem
    # Quellenwechsel ein zweiter Mediendatensatz, und der alte bleibt als
    # Waise in der Mediathek liegen. An der Rolle festgemacht ist derselbe
    # Platz desselben Buchs immer derselbe Datensatz — er bekommt nur neue
    # Bytes. Erst dadurch REPARIERT ein erneuter Lauf ein korrigiertes Cover.
    dateien: list[tuple[Path, str]] = []
    if bilder.get("cover"):
        dateien.append((Path(bilder["cover"]), "cover"))
    for i, g in enumerate(bilder.get("galerie", [])):
        dateien.append((Path(g), f"galerie{i}"))

    if dry_run:
        f = _gewicht_ergaenzen(f, cfg, log=log)      # ohne Shop: nur rechnen
        medien = [{"media_id": _media_id(isbn, rolle), "cover": rolle == "cover",
                   "datei": str(p)} for p, rolle in dateien]
        payload = baue_produkt(f, eff, medien, kategorien=kategorien)
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

    f = _gewicht_ergaenzen(f, cfg, client=client, log=log)

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
    for p, rolle in dateien:
        mid = _media_id(isbn, rolle)
        log(f"Lade Bild hoch: {p.name} …")
        client.medium_hochladen(p, mid, p.stem, ordner_id)
        medien.append({"media_id": mid, "cover": rolle == "cover",
                       "datei": str(p)})

    payload = baue_produkt(f, eff, medien, bestehende_id, kategorien)
    log(f"{'Aktualisiere' if bestehende_id else 'Lege an'}: {nummer} "
        f"({aktive_umgebung(cfg)}) …")
    client.sync("product", [payload])
    return {"payload": payload, "medien": medien, "neu": not bestehende_id,
            "admin_url": f"{client.base}/admin#/sw/product/detail/{payload['id']}"}
