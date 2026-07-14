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


DEFAULT_CONFIG = {
    # --- Shopware-Zugang (Admin -> Einstellungen -> System -> Integrationen) --
    # Das Secret steht im Klartext in dieser config.json. Sie liegt neben der
    # .exe und ist per .gitignore vom Repo ausgeschlossen.
    "shop_url": "",                 # z. B. https://shop.verlag-regionalkultur.de
    "access_key_id": "",
    "secret_access_key": "",

    # --- Zuordnungen aus dem Shop (werden beim "Verbinden" befüllt) ----------
    "tax_id": "",                   # Steuersatz-ID (Bücher: 7 %)
    "tax_rate": 7.0,                # zugehöriger Satz, für die Netto-Rechnung
    "currency_id": "",              # EUR
    "category_id": "",              # Zielkategorie (optional)
    "manufacturer_id": "",          # Hersteller/Verlag (optional)
    "default_stock": 0,

    # --- Produkt-Voreinstellungen -------------------------------------------
    # Entwurf: Produkt wird angelegt, ist aber im Shop nicht sichtbar.
    "aktiv": False,

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


def _kontributoren(product, rolle: str) -> list[str]:
    dd = product.find("descriptivedetail")
    if dd is None:
        return []
    beitraege = []
    for c in dd.findall("contributor"):
        if _txt(c.find("b035")) == rolle:
            seq = _txt(c.find("b034"))
            name = _txt(c.find("b036"))
            if name:
                beitraege.append((int(seq) if seq.isdigit() else 999, name))
    return [n for _, n in sorted(beitraege)]


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

    # Titel (Produktebene, nicht die Collection-Kopie)
    titel = ""
    if dd is not None:
        for td in dd.findall("titledetail"):
            if _txt(td.find("b202")) == "01":
                for te in td.findall("titleelement"):
                    if _txt(te.find("x409")) == "01":
                        titel = _txt(te.find("b203"))
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
    datum = ""
    if len(roh_datum) == 8 and roh_datum.isdigit():
        datum = f"{roh_datum[6:8]}.{roh_datum[4:6]}.{roh_datum[0:4]}"

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
        "serientitel": serientitel,
        "band": band,
        "autoren": autoren,
        "herausgeber": herausgeber,
        "seiten": seiten,
        "einband": einband,
        "verlag": verlag,
        "preis_brutto": preis_brutto,
        "waehrung": waehrung,
        "datum": datum,
        "cover_url": cover_url,
        "werbetext_absaetze": werbetext,
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
    """Werbetext (ONIX d104) als Absätze + kompakter Fakten-Block."""
    teile = [f"<p>{html.escape(a)}</p>" for a in f.get("werbetext_absaetze", [])]

    fakten = []
    if f.get("herausgeber"):
        fakten.append(f"{_join_und(f['herausgeber'])} (Hrsg.)")
    if f.get("autoren"):
        fakten.append(_join_und(f["autoren"]))
    umfang = []
    if f.get("seiten"):
        umfang.append(f"{f['seiten']} Seiten")
    if f.get("einband"):
        umfang.append(f["einband"])
    if umfang:
        fakten.append(", ".join(umfang))
    if f.get("isbn13_formatiert"):
        fakten.append(f"ISBN {f['isbn13_formatiert']}")
    if f.get("datum"):
        fakten.append(f"Erschienen: {f['datum']}")
    if f.get("preis_brutto"):
        betrag = f"{f['preis_brutto']:.2f}".replace(".", ",")
        fakten.append(f"{betrag} {f.get('waehrung', 'EUR')}")

    if fakten:
        teile.append("<hr />")
        teile.append("<p>" + " &middot; ".join(html.escape(x) for x in fakten)
                     + "</p>")
    return "\n".join(teile)


def netto(brutto: float, satz: float) -> float:
    return round(float(brutto) / (1.0 + float(satz) / 100.0), 4)


def baue_produkt(f: dict, cfg: dict, medien: list[dict] | None = None) -> dict:
    """Baut den Shopware-Produkt-Payload (upsert).

    ``medien`` = [{"media_id": .., "cover": bool}, ...] (bereits hochgeladen).
    """
    isbn = f["isbn13"]
    satz = float(cfg.get("tax_rate", 7.0))
    brutto = float(f.get("preis_brutto") or 0.0)

    payload = {
        "id": produkt_id(isbn),
        "productNumber": isbn,
        "ean": isbn,
        "name": produkt_name(f),
        "active": bool(cfg.get("aktiv", False)),   # Entwurf: inaktiv
        "stock": int(cfg.get("default_stock", 0)),
        "description": baue_beschreibung(f),
        "taxId": cfg.get("tax_id") or None,
        "price": [{
            "currencyId": cfg.get("currency_id") or None,
            "gross": round(brutto, 2),
            "net": netto(brutto, satz),
            "linked": True,
        }],
    }
    if cfg.get("manufacturer_id"):
        payload["manufacturerId"] = cfg["manufacturer_id"]
    if cfg.get("category_id"):
        payload["categories"] = [{"id": cfg["category_id"]}]

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


class ShopClient:
    """Minimaler Admin-API-Client (OAuth client_credentials + Sync + Medien)."""

    def __init__(self, shop_url: str, key: str, secret: str, timeout: float = 30.0):
        self.base = (shop_url or "").rstrip("/")
        self.key = key
        self.secret = secret
        self.timeout = timeout
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
            with urllib.request.urlopen(req, timeout=self.timeout) as r:
                return r.read()
        except urllib.error.HTTPError as e:
            detail = e.read().decode("utf-8", "replace")[:800]
            raise ShopFehler(f"{method} {pfad} -> HTTP {e.code}\n{detail}") from e
        except urllib.error.URLError as e:
            raise ShopFehler(f"{method} {pfad} -> nicht erreichbar: {e.reason}") from e

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
                    dry_run: bool = False, log=print) -> dict:
    """Lädt die Bilder hoch und legt/aktualisiert das Produkt als Entwurf.

    dry_run=True baut nur den Payload (nichts wird gesendet).
    Rückgabe: {"payload": .., "medien": [..], "admin_url": ..}
    """
    isbn = f["isbn13"]
    bilder = bilder or {"cover": None, "galerie": []}

    dateien: list[tuple[Path, bool]] = []
    if bilder.get("cover"):
        dateien.append((Path(bilder["cover"]), True))
    for g in bilder.get("galerie", []):
        dateien.append((Path(g), False))

    if dry_run:
        medien = [{"media_id": _media_id(isbn, p.stem), "cover": ist_cover,
                   "datei": str(p)} for p, ist_cover in dateien]
        payload = baue_produkt(f, cfg, medien)
        log(f"Dry-Run: Payload für {payload['productNumber']} gebaut "
            f"({len(medien)} Bild(er)) — nichts gesendet.")
        return {"payload": payload, "medien": medien, "admin_url": ""}

    client = ShopClient(cfg.get("shop_url", ""), cfg.get("access_key_id", ""),
                        cfg.get("secret_access_key", ""))
    fehlend = [k for k in ("tax_id", "currency_id") if not cfg.get(k)]
    if fehlend:
        raise ShopFehler("Zuordnung fehlt: " + ", ".join(fehlend) +
                         " — bitte einmal 'Verbinden' und auswählen.")

    ordner_id = client.produkt_medien_ordner() if dateien else None
    medien = []
    for p, ist_cover in dateien:
        mid = _media_id(isbn, p.stem)
        log(f"Lade Bild hoch: {p.name} …")
        client.medium_hochladen(p, mid, p.stem, ordner_id)
        medien.append({"media_id": mid, "cover": ist_cover, "datei": str(p)})

    payload = baue_produkt(f, cfg, medien)
    log(f"Lege Produkt an/aktualisiere: {payload['productNumber']} …")
    client.sync("product", [payload])
    return {"payload": payload, "medien": medien,
            "admin_url": client.admin_link(isbn)}
