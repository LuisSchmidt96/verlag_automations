"""
PI/BI-Generator Core-Logik
===========================
Reine Datenlogik ohne UI-Abhängigkeiten — liest einen VLB-ONIX-XML-Datensatz
und erzeugt daraus Presse- (PI) und Buchinformation (BI) als docx + html.

Ablauf:
1. lade_buchdaten(xml)          -> Buchdaten (aus ONIX 3.1)
2. lade_cover_datei / _web      -> Cover-Bytes (lokal oder vom Webserver)
3. generiere_docx / generiere_html(vorlage, buch, ...) -> gefüllte Dokumente

Die docx-Vorlagen (vorlagen/*.docx) werden einmalig mit baue_docx_vorlagen()
aus den Muster-Dokumenten in beispiele/ erzeugt und mitgeliefert; jede
buchspezifische Textzeile steckt darin als Platzhalter ({{TITEL}} usw.) in
genau einem Run, damit das Ersetzen zur Laufzeit robust ist.
"""

from __future__ import annotations

import html
import io
import json
import os
import re
import sys
import urllib.request
import xml.etree.ElementTree as ET
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from PIL import Image


# ---------------------------------------------------------------------
# Ordner (Config neben der .exe, Vorlagen als Ressourcen)
# ---------------------------------------------------------------------

def _base_dir() -> Path:
    """Ordner neben der .exe (PyInstaller-Build) bzw. neben dem Tool-Code."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


# App-mutierte Laufzeitdaten (config.json, cover_cache/, pi_bi_output/) liegen
# direkt neben der .exe — wie bei booxpress_etiketten, kein data-Unterordner.
APP_DIR = _base_dir()
APP_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_PFAD = APP_DIR / "config.json"


def _vorlagen_dir() -> Path:
    """Nur-Lese-Vorlagen (docx/html). Im PyInstaller-Build liegen sie unter
    sys._MEIPASS (per datas gebündelt), sonst neben diesem Modul."""
    if getattr(sys, "frozen", False):
        p = Path(getattr(sys, "_MEIPASS", "")) / "pi_bi_generator" / "vorlagen"
        if p.exists():
            return p
        return _base_dir() / "vorlagen"
    return Path(__file__).parent / "vorlagen"


VORLAGEN_DIR = _vorlagen_dir()


# ---------------------------------------------------------------------
# Konfiguration
# ---------------------------------------------------------------------

DEFAULT_CONFIG = {
    "output_dir": "pi_bi_output",
    "last_input_dir": "",
    "verlag_name": "verlag regionalkultur",
    # Fester ISBN-Präfix des Verlags (978-3-95505-…) für Formatierung/Kurzcode.
    "isbn_prefix": "978-3-95505",
    # Basis-URLs für die HTML-Fassung (Cover-Thumbnail + Blick-ins-Buch-PDF).
    # Diese Web-Assets werden vom Nutzer separat hochgeladen.
    "newsletter_base_url": "https://verlag-regionalkultur.de/newsletter_/",
    "bib_base_url": "https://verlag-regionalkultur.de/presse/bib/",
    # Fallback-Link auf die Webshop-Detailseite (nicht aus der XML ableitbar).
    "detail_fallback_url": "https://verlag-regionalkultur.de/",
    # Einband-Code (ONIX b012) -> Klartext, so wie der Verlag ihn schreibt.
    # BC/PB hiessen hier "kartoniert" — das Wort benutzt der Verlag nicht;
    # in PI, BI und im Shop heisst der Softcover "Broschur" (im Shopbestand
    # 124 x "Broschur" gegen 1 x "kartoniert").
    "einband_map": {
        "BB": "fester Einband",
        "BC": "Broschur",
        "BE": "Klappenbroschur",
        "PB": "Broschur",
        "BZ": "Leinen",
    },
    # Beugung nach „hrsg. von“: Name der Körperschaft -> Dativform samt
    # Artikel. Aus den Daten ist sie nicht abzuleiten — „der Deutschen
    # Waldenservereinigung“, aber „dem Stadtarchiv“, je nach Geschlecht des
    # Namens. Geraten wird sie deshalb nicht: einmal hier eintragen, dann
    # stimmt sie für jeden weiteren Band der Reihe.
    "koerperschaft_dativ": {
        "Deutsche Waldenservereinigung e.V. Ötisheim-Schönenberg":
            "der Deutschen Waldenservereinigung e.V. Ötisheim-Schönenberg",
    },
    "cover_timeout": 15,
    # Wird erhöht, wenn eine der Tabellen oben sich ändert — sonst bliebe in
    # einer gewachsenen config.json (sie liegt neben der .exe) für immer der
    # alte Wert stehen, z. B. „kartoniert“ statt „Broschur“.
    "config_version": 2,
}

# Rein technische Tabellen: die gehören dem Code, nicht dem Bediener, und
# werden bei einer Versionserhöhung aufgefrischt.
_VORGABE_SCHLUESSEL = ("einband_map",)


def _migriere(cfg: dict) -> dict:
    """Vorgaben auffrischen, ohne eigene Einträge zu verlieren."""
    if cfg.get("config_version", 0) < DEFAULT_CONFIG["config_version"]:
        for k in _VORGABE_SCHLUESSEL:
            cfg[k] = json.loads(json.dumps(DEFAULT_CONFIG[k]))
        cfg["config_version"] = DEFAULT_CONFIG["config_version"]
    # Die Dativ-Tabelle wächst beim Verlag weiter — hier nur ergänzen, was
    # noch fehlt, sonst wären eigene Einträge nach einem Update weg.
    tabelle = cfg.setdefault("koerperschaft_dativ", {})
    for k, v in DEFAULT_CONFIG["koerperschaft_dativ"].items():
        tabelle.setdefault(k, v)
    return cfg


def lade_config() -> dict:
    if CONFIG_PFAD.exists():
        with open(CONFIG_PFAD, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        for k, v in DEFAULT_CONFIG.items():
            cfg.setdefault(k, v)
        cfg = _migriere(cfg)
        speichere_config(cfg)          # damit die Auffrischung nicht verpufft
        return cfg
    with open(CONFIG_PFAD, "w", encoding="utf-8") as f:
        json.dump(DEFAULT_CONFIG, f, indent=2, ensure_ascii=False)
    return dict(DEFAULT_CONFIG)


def speichere_config(cfg: dict) -> None:
    with open(CONFIG_PFAD, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)


# ---------------------------------------------------------------------
# Datenmodell
# ---------------------------------------------------------------------

class Buchdaten:
    __slots__ = (
        "isbn13", "isbn13_formatiert", "shortcode",
        "titel", "untertitel", "serientitel", "band", "band_text",
        "herausgeber", "koerperschaften", "autoren",
        # Kopf = die Zeilen über dem Werbetext, Kasten = der Rahmen darunter.
        # Beide zeigen dieselben Namen, aber anders getrennt (siehe unten).
        "editoren_kopf", "editoren_kasten", "kopf_zusatz",
        "mitwirkende", "reihe_zeile", "kasten_zusatz",
        "umfang_zeile", "titel_band",
        "verlag", "preis", "verlag_isbn_preis", "datum",
        "cover_url", "werbetext_absaetze",
    )

    def __init__(self, **kw):
        for name in self.__slots__:
            setattr(self, name, kw.get(name, ""))
        for liste in ("herausgeber", "koerperschaften", "autoren",
                      "kasten_zusatz"):
            if not getattr(self, liste):
                setattr(self, liste, [])
        if not self.werbetext_absaetze:
            self.werbetext_absaetze = []


# ---------------------------------------------------------------------
# ONIX-Parser (3.1, Kurz-Tags, ohne Namespace)
# ---------------------------------------------------------------------

def _txt(el) -> str:
    return (el.text or "").strip() if el is not None else ""


# Einleitungs-Label, die aus dem Werbetext entfernt werden sollen.
_INTRO_LABELS = ("Aus dem Vorwort",)


def _entferne_intro(absaetze: list[str]) -> list[str]:
    """Entfernt ein führendes Label wie 'Aus dem Vorwort:' — egal ob als
    eigener Absatz oder als Prefix des ersten Absatzes."""
    if not absaetze:
        return absaetze
    for label in _INTRO_LABELS:
        pat = re.compile(rf"^\s*{re.escape(label)}\s*[:\-–—]?\s*", re.IGNORECASE)
        m = pat.match(absaetze[0])
        if m:
            rest = absaetze[0][m.end():].strip()
            return ([rest] if rest else []) + absaetze[1:]
    return absaetze


def _join_und(namen: list[str]) -> str:
    """['A','B','C'] -> 'A, B und C'."""
    namen = [n for n in namen if n]
    if not namen:
        return ""
    if len(namen) == 1:
        return namen[0]
    return ", ".join(namen[:-1]) + " und " + namen[-1]


def _kontributoren(product, rolle: str) -> list[dict]:
    """[{'name': 'Albert de Lange', 'koerperschaft': False}, …]

    b047 = Körperschaft („Deutsche Waldenservereinigung e.V. Ötisheim-
    Schönenberg“). Gelesen wurde bisher nur b036, der Personenname — ein
    körperschaftlicher Herausgeber fiel damit stillschweigend aus PI und BI
    heraus. Bei den Waldenserstudien fehlte so die herausgebende Vereinigung
    ganz, und in der Kopfzeile stand nur einer von mehreren Herausgebern.
    """
    dd = product.find("descriptivedetail")
    if dd is None:
        return []
    beitraege = []
    for c in dd.findall("contributor"):
        if _txt(c.find("b035")) != rolle:
            continue
        seq = int(_txt(c.find("b034"))) if _txt(c.find("b034")).isdigit() else 999
        koerperschaft = _txt(c.find("b047"))
        name = koerperschaft or _txt(c.find("b036"))
        if name:
            beitraege.append((seq, {"name": name,
                                    "koerperschaft": bool(koerperschaft)}))
    return [k for _, k in sorted(beitraege, key=lambda x: x[0])]


# VLB liefert dieselben Daten in zwei Schreibweisen: mit KURZ-Tags (<b012>) und
# mit REFERENZ-Tags (<ProductForm>). Welche man bekommt, haengt allein am
# Exportdialog — die Referenzfassung heisst dann "onix3Ref_….xml". Statt jeden
# Lesezugriff zu verdoppeln, wird der Baum einmalig auf die Kurzform gebracht;
# darunter bleibt der gesamte Lesecode unveraendert.
#
# Dieselbe Tabelle steht im shopware_publisher. Das ist bewusste Dopplung
# (siehe CLAUDE.md: kein geteilter Code zwischen den Werkzeugen) — jedes
# Werkzeug bleibt fuer sich lauffaehig. Wer hier ein Feld ergaenzt, sollte
# dort nachsehen.
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
    ONIX erlaubt ihn, und ein Namensraum wuerde jedes find() ins Leere laufen
    lassen — lautlos, was der schlimmste Fall waere.
    """
    for el in root.iter():
        if isinstance(el.tag, str):
            if "}" in el.tag:
                el.tag = el.tag.split("}", 1)[1]
            el.tag = REFERENZ_ZU_KURZ.get(el.tag, el.tag)

def lade_buchdaten(xml_pfad, cfg: dict | None = None) -> Buchdaten:
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

    # -- ISBN / Kurzcode ------------------------------------------------
    isbn13 = ""
    for pid in product.findall("productidentifier"):
        if _txt(pid.find("b221")) in ("03", "15"):
            isbn13 = _txt(pid.find("b244"))
            break
    if not isbn13 or len(isbn13) != 13 or not isbn13.isdigit():
        raise ValueError(f"Keine gültige ISBN-13 gefunden (gelesen: {isbn13!r}).")
    e = isbn13
    isbn13_formatiert = f"{isbn_prefix}-{e[9:12]}-{e[12]}"
    shortcode = f"{e[7:9]}-{e[9:12]}-{e[12]}"

    # -- Titel + Untertitel (Produktebene, NICHT die Collection-Kopie) --
    # b029 (Untertitel) wurde bisher gar nicht gelesen. Er gehört im Kasten
    # hinter den Titel und im Kopf unter den Titel — sonst fehlt bei einem
    # Tagungsband die halbe Aussage („Zwischen Bild und Wirklichkeit in
    # England, Piemont, Nordamerika und Württemberg (1600–1900)“).
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

    # -- Serie / Band ---------------------------------------------------
    serientitel = ""
    band = ""
    if dd is not None:
        coll = dd.find("collection")
        if coll is not None:
            for te in coll.findall("titledetail/titleelement"):
                lvl = _txt(te.find("x409"))
                if lvl == "02" and not serientitel:
                    serientitel = _txt(te.find("b203"))
                if lvl == "01" and not band:
                    band = _txt(te.find("x410"))
    band_text = f"Band {band}" if band else ""

    # -- Beteiligte -----------------------------------------------------
    hrsg = _kontributoren(product, "B01")
    herausgeber = [k["name"] for k in hrsg if not k["koerperschaft"]]
    koerperschaften = [k["name"] for k in hrsg if k["koerperschaft"]]
    autoren = [k["name"] for k in _kontributoren(product, "A01")]
    mitwirkende = (f"Mit Beiträgen von {_join_und(autoren)}."
                   if autoren else "")

    # -- Reihenzeile ----------------------------------------------------
    # Die Reihe kommt nur dann in den Kasten, wenn sie NICHT der Buchtitel
    # selbst ist: bei den „Bausteinen“ heisst die Reihe wie das Buch, dort
    # stünde sonst zweimal dasselbe (so hält es auch das Musterdokument).
    #
    # Die herausgebende Körperschaft steht in dieser Zeile, nicht bei den
    # Personen — so korrigiert der Verlag es von Hand. Gibt es keine
    # Reihenzeile, gehört sie zu den Herausgebern, sonst verschwände sie
    # wieder.
    reihe_zeile = ""
    if serientitel and serientitel.strip() != titel.strip():
        stuecke = [serientitel]
        if koerperschaften:
            # „hrsg. von der Deutschen Waldenservereinigung e.V.“ — Artikel
            # und Beugung stehen in der Tabelle koerperschaft_dativ. Fehlt
            # ein Name dort, steht die ungebeugte Form da und das Werkzeug
            # sagt Bescheid; geraten wird nicht.
            dativ = cfg.get("koerperschaft_dativ", {})
            gebeugt = []
            for k in koerperschaften:
                if k in dativ:
                    gebeugt.append(dativ[k])
                else:
                    gebeugt.append(k)
                    print(f"⚠ Körperschaft {k!r} steht nicht in "
                          f"koerperschaft_dativ (config.json) — in der "
                          f"Reihenzeile fehlt die Beugung ('der …'/'dem …').")
            stuecke.append("hrsg. von " + _join_und(gebeugt))
        if band:
            stuecke.append(f"Bd. {band}")
        reihe_zeile = ", ".join(stuecke)

    # Kopf: „A, B und C (Hrsg.)“ — Kasten: „A, B, C (Hrsg.)“. Beides so
    # vorgegeben; im alten Musterdokument war es genau andersherum (Kopf mit
    # Schrägstrichen), das war nicht gewollt.
    namen = herausgeber + ([] if reihe_zeile else koerperschaften)
    editoren_kopf = (_join_und(namen) + " (Hrsg.)") if namen else ""
    editoren_kasten = (", ".join(namen) + " (Hrsg.)") if namen else ""

    # -- Umfang / Einband ----------------------------------------------
    seiten = _txt(dd.find("extent/b219")) if dd is not None else ""
    anzahl_abb = _txt(dd.find("b125")) if dd is not None else ""
    abb_text = _txt(dd.find("b062")) if dd is not None else ""
    einband_code = _txt(dd.find("b012")) if dd is not None else ""
    einband = einband_map.get(einband_code, "")
    if not einband and einband_code:
        einband = einband_code  # unbekannter Code: roh anzeigen
        print(f"⚠ Unbekannter Einband-Code b012={einband_code!r} — "
              f"bitte einband_map in config.json ergänzen.")
    teile = []
    if seiten:
        teile.append(f"{seiten} Seiten")
    if anzahl_abb and abb_text:
        teile.append(f"mit {anzahl_abb} {abb_text}")
    umfang_zeile = " ".join(teile)
    if einband:
        umfang_zeile = (umfang_zeile + ", " if umfang_zeile else "") + einband
    if umfang_zeile:
        umfang_zeile += "."

    # Kasten-Titelzeile: „Titel. Untertitel.“ Die Bandangabe steht hier nur
    # noch, wenn es keine eigene Reihenzeile gibt (Muster: „Bausteine …
    # Band 5.“) — sonst stünde sie zweimal.
    stuecke = [titel]
    if untertitel:
        stuecke.append(untertitel)
    if band_text and not reihe_zeile:
        stuecke.append(band_text)
    titel_band = ". ".join(t.rstrip(". ") for t in stuecke if t) + "."

    # Kopfzeile unter dem Titel: der Untertitel, wenn es einen gibt, sonst
    # die Bandangabe (Muster: „Band 5“). Beides zugleich stand dort nie.
    kopf_zusatz = untertitel or band_text

    # Zusatzzeilen im Kasten, in dieser Reihenfolge. Leere fallen weg —
    # der Absatz wird dann entfernt, statt eine Leerzeile zu hinterlassen.
    kasten_zusatz = [z for z in (reihe_zeile, mitwirkende) if z]

    # -- Verlag / Preis / Datum ----------------------------------------
    verlag = _txt(product.find("publishingdetail/publisher/b081")) \
        or cfg.get("verlag_name", "")
    preis = ""
    for price in product.findall("productsupply/supplydetail/price"):
        terr = price.find("territory")
        x449 = _txt(terr.find("x449")) if terr is not None else ""
        if x449 == "DE":
            betrag = _txt(price.find("j151"))
            waehrung = _txt(price.find("j152")) or "EUR"
            if betrag:
                try:
                    preis = f"{waehrung} {float(betrag):.2f}".replace(".", ",")
                except ValueError:
                    preis = f"{waehrung} {betrag}"
            break
    verlag_isbn_preis = f"{verlag}, ISBN {isbn13_formatiert}."
    if preis:
        verlag_isbn_preis += f" {preis}."

    roh_datum = _txt(product.find("publishingdetail/publishingdate/b306"))
    datum = ""
    if len(roh_datum) == 8 and roh_datum.isdigit():
        datum = f"{roh_datum[6:8]}.{roh_datum[4:6]}.{roh_datum[0:4]}"

    # -- Cover-URL (Vorderseite x436=01) -------------------------------
    cover_url = ""
    for sr in product.findall("collateraldetail/supportingresource"):
        if _txt(sr.find("x436")) == "01":
            cover_url = _txt(sr.find("resourceversion/x435"))
            break

    # -- Werbetext (d104, mehrzeilig) ----------------------------------
    d104 = _txt(product.find("collateraldetail/textcontent/d104"))
    werbetext_absaetze = [ln.strip() for ln in d104.split("\n") if ln.strip()]
    werbetext_absaetze = _entferne_intro(werbetext_absaetze)

    return Buchdaten(
        isbn13=isbn13, isbn13_formatiert=isbn13_formatiert, shortcode=shortcode,
        titel=titel, untertitel=untertitel,
        serientitel=serientitel, band=band, band_text=band_text,
        herausgeber=herausgeber, koerperschaften=koerperschaften,
        autoren=autoren,
        editoren_kopf=editoren_kopf, editoren_kasten=editoren_kasten,
        kopf_zusatz=kopf_zusatz, reihe_zeile=reihe_zeile,
        kasten_zusatz=kasten_zusatz,
        mitwirkende=mitwirkende, umfang_zeile=umfang_zeile, titel_band=titel_band,
        verlag=verlag, preis=preis, verlag_isbn_preis=verlag_isbn_preis,
        datum=datum, cover_url=cover_url, werbetext_absaetze=werbetext_absaetze,
    )


# ---------------------------------------------------------------------
# Cover laden (lokal oder vom Webserver)
# ---------------------------------------------------------------------

def lade_cover_datei(pfad) -> bytes:
    return Path(pfad).read_bytes()


def web_cover_url(cfg: dict, buch: Buchdaten) -> str:
    return f"{cfg.get('newsletter_base_url', '')}{buch.shortcode}.png"


def lade_cover_web(url: str, timeout: float = 15.0) -> bytes:
    req = urllib.request.Request(
        url, headers={"User-Agent": "PiBiGenerator/1.0"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


# ---------------------------------------------------------------------
# docx-Platzhalter (buchspezifische Zeilen)
# ---------------------------------------------------------------------

TOKEN_WERBETEXT = "{{WERBETEXT}}"


# Die Platzhalternamen stammen aus den Musterdokumenten und sind historisch —
# sie sagen, WO etwas steht, nicht mehr, wie es getrennt ist:
#
#   {{EDITOREN}}      Kopf, über dem Titel      -> „A, B und C (Hrsg.)“
#   {{BAND}}          Kopf, unter dem Titel     -> Untertitel, sonst „Band 5“
#   {{EDITOREN_UND}}  Kasten, erste Zeile       -> „A, B, C (Hrsg.)“
#   {{TITEL_BAND}}    Kasten, Titelzeile        -> „Titel. Untertitel.“
#   {{MITWIRKENDE}}   Kasten, Zusatzzeilen      -> Reihe + „Mit Beiträgen von“
#
# Umbenennen hiesse, die beiden .docx-Vorlagen anzufassen; der Gewinn wäre
# rein kosmetisch, das Risiko nicht.
def _mapping(buch: Buchdaten) -> dict:
    """Einzeilige Platzhalter -> Wert.

    Werbetext und Kasten-Zusatzzeilen sind mehrzeilig und stehen deshalb
    nicht hier, sondern in _fuelle_mehrzeilig.
    """
    return {
        "{{EDITOREN}}": buch.editoren_kopf,
        "{{TITEL}}": buch.titel,
        "{{BAND}}": buch.kopf_zusatz,
        "{{EDITOREN_UND}}": buch.editoren_kasten,
        "{{TITEL_BAND}}": buch.titel_band,
        "{{UMFANG_ZEILE}}": buch.umfang_zeile,
        "{{VERLAG_ISBN_PREIS}}": buch.verlag_isbn_preis,
    }


def _textruns(para):
    """Runs eines Absatzes ohne Zeichnungs-Runs (Cover-Bild etc.)."""
    return [r for r in para.runs if r._r.find(qn("w:drawing")) is None]


def _ersetze_in_absaetzen(absaetze, mapping):
    for p in absaetze:
        runs = _textruns(p)
        if not runs:
            continue
        full = "".join(r.text for r in runs)
        if "{{" not in full:
            continue
        neu = full
        for k, v in mapping.items():
            neu = neu.replace(k, v)
        if neu != full:
            runs[0].text = neu
            for r in runs[1:]:
                r.text = ""


def _alle_absaetze(doc):
    """Alle Absätze aus Body + Kopf-/Fußzeilen + Tabellen."""
    yield from doc.paragraphs
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                yield from cell.paragraphs
    for sec in doc.sections:
        for hf in (sec.header, sec.first_page_header, sec.even_page_header,
                   sec.footer, sec.first_page_footer, sec.even_page_footer):
            if hf is not None:
                yield from hf.paragraphs


def _ersetze_platzhalter(doc, mapping):
    _ersetze_in_absaetzen(list(_alle_absaetze(doc)), mapping)


def _fuelle_mehrzeilig(doc, token: str, zeilen: list[str],
                       leer_behalten: bool = False):
    """Ersetzt einen Platzhalter-Absatz durch je einen Absatz pro Zeile,
    unter Beibehaltung von Formatierung und Absatzstil der Vorlage.

    ``leer_behalten`` steuert den Fall ohne Inhalt: der Werbetext behält
    seinen (leeren) Absatz, die Zusatzzeilen im Kasten nicht — dort bliebe
    sonst eine Leerzeile zwischen Titel und Umfangzeile stehen.
    """
    ziel = None
    for p in doc.paragraphs:
        if p.text.strip() == token:
            ziel = p
            break
    if ziel is None:
        return
    for line in (zeilen or ([""] if leer_behalten else [])):
        neu = deepcopy(ziel._p)
        para = Paragraph(neu, ziel._parent)
        runs = _textruns(para)
        if runs:
            runs[0].text = line
            for r in runs[1:]:
                r.text = ""
        ziel._p.addprevious(neu)
    ziel._p.getparent().remove(ziel._p)


# ---------------------------------------------------------------------
# Cover-Bild im docx austauschen
# ---------------------------------------------------------------------

def _cover_zeichnung(doc):
    """Das <w:drawing>, das ein eingebettetes Bild (a:blip) referenziert."""
    for p in doc.paragraphs:
        for r in p.runs:
            drawing = r._r.find(qn("w:drawing"))
            if drawing is not None and drawing.find(".//" + qn("a:blip")) is not None:
                return drawing
    return None


def _setze_cover_groesse(drawing, breite_beibehalten: bool, w: int, h: int):
    """Passt die Bildbox an das Seitenverhältnis des neuen Covers an
    (Breite bleibt, Höhe wird nachgezogen), damit nichts verzerrt wird."""
    if not w or not h:
        return
    ext = drawing.find(".//" + qn("wp:extent"))
    if ext is None or not ext.get("cx"):
        return
    cx = int(ext.get("cx"))
    cy = round(cx * h / w)
    ext.set("cy", str(cy))
    for aext in drawing.findall(".//" + qn("a:ext")):
        if aext.get("cx"):  # nur die größenrelevanten a:ext (nicht extLst-URI)
            aext.set("cx", str(cx))
            aext.set("cy", str(cy))


def _tausche_cover(doc, cover_bytes: bytes) -> bool:
    """Überschreibt das Cover-Bild (word/media/image1.jpeg) und zieht die
    Bildbox auf das echte Seitenverhältnis. Gibt True bei Erfolg zurück."""
    cover_rel = None
    for rel in doc.part.rels.values():
        if "image" in rel.reltype and rel.target_ref.endswith("image1.jpeg"):
            cover_rel = rel
            break
    if cover_rel is None:
        return False
    cover_rel.target_part._blob = cover_bytes
    try:
        im = Image.open(io.BytesIO(cover_bytes))
        drawing = _cover_zeichnung(doc)
        if drawing is not None:
            _setze_cover_groesse(drawing, True, im.size[0], im.size[1])
    except Exception as ex:  # Größenanpassung ist optional
        print(f"⚠ Cover-Größe konnte nicht angepasst werden: {ex}")
    return True


# ---------------------------------------------------------------------
# docx erzeugen
# ---------------------------------------------------------------------

def generiere_docx(vorlage_pfad, buch: Buchdaten, cover_bytes: bytes | None,
                   ziel_pfad) -> None:
    doc = Document(str(vorlage_pfad))
    _fuelle_mehrzeilig(doc, TOKEN_WERBETEXT, buch.werbetext_absaetze,
                       leer_behalten=True)
    _fuelle_mehrzeilig(doc, "{{MITWIRKENDE}}", buch.kasten_zusatz)
    _ersetze_platzhalter(doc, _mapping(buch))
    if cover_bytes:
        _tausche_cover(doc, cover_bytes)
    ziel_pfad = Path(ziel_pfad)
    ziel_pfad.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(ziel_pfad))


# ---------------------------------------------------------------------
# html erzeugen
# ---------------------------------------------------------------------

def _html_mapping(buch: Buchdaten, detail_url: str, cfg: dict) -> dict:
    nl = cfg.get("newsletter_base_url", "")
    bib = cfg.get("bib_base_url", "")
    werbetext_html = "<br />\n".join(
        html.escape(a) for a in buch.werbetext_absaetze)
    # Eine Zeile je Zusatzangabe, mit derselben Auszeichnung wie der Rest des
    # Kastens. Ohne Zusatzangaben bleibt der Platz leer statt eine
    # Leerzeile zu erzeugen — deshalb steht das Markup hier und nicht in der
    # Vorlage.
    zusatz_html = "".join(
        f'<font size="2" style="font-style:italic">{html.escape(z)}</font>'
        "<br />\n\t\t\t" for z in buch.kasten_zusatz)
    return {
        "{{EDITOREN}}": html.escape(buch.editoren_kopf),
        "{{EDITOREN_UND}}": html.escape(buch.editoren_kasten),
        "{{TITEL_HTML}}": html.escape(buch.titel),
        "{{TITEL_BAND}}": html.escape(buch.titel_band),
        "{{KASTEN_ZUSATZ_HTML}}": zusatz_html,
        "{{BAND}}": html.escape(buch.band),
        "{{BAND_TEXT}}": html.escape(buch.kopf_zusatz),
        "{{DATUM}}": html.escape(buch.datum),
        "{{DETAIL_URL}}": html.escape(detail_url, quote=True),
        "{{COVER_THUMB_URL}}": html.escape(f"{nl}{buch.shortcode}.png", quote=True),
        "{{BIB_URL}}": html.escape(f"{bib}bib_{buch.shortcode}.pdf", quote=True),
        "{{WERBETEXT_HTML}}": werbetext_html,
        "{{UMFANG_ZEILE}}": html.escape(buch.umfang_zeile),
        "{{VERLAG_ISBN_PREIS}}": html.escape(buch.verlag_isbn_preis),
        "{{TITEL}}": html.escape(buch.titel),
    }


def generiere_html(vorlage_pfad, buch: Buchdaten, detail_url: str, cfg: dict,
                   ziel_pfad) -> None:
    text = Path(vorlage_pfad).read_text(encoding="utf-8")
    for k, v in _html_mapping(buch, detail_url, cfg).items():
        text = text.replace(k, v)
    ziel_pfad = Path(ziel_pfad)
    ziel_pfad.parent.mkdir(parents=True, exist_ok=True)
    ziel_pfad.write_text(text, encoding="utf-8")


# ---------------------------------------------------------------------
# Vorlagen-Erzeugung (Entwickler-Routine, einmalig)
# ---------------------------------------------------------------------

# Absatz-Index -> Platzhalter im Muster-docx (PI und BI identisch bis P16).
_VORLAGE_TOKENS = {
    1: "{{EDITOREN}}",
    3: "{{TITEL}}",
    5: "{{BAND}}",              # nur Text-Run; Cover-Zeichnung bleibt
    7: TOKEN_WERBETEXT,         # P8, P9 werden gelöscht
    12: "{{EDITOREN_UND}}",
    13: "{{TITEL_BAND}}",
    14: "{{MITWIRKENDE}}",
    15: "{{UMFANG_ZEILE}}",
    16: "{{VERLAG_ISBN_PREIS}}",
}


def _setze_token(para, token: str):
    runs = _textruns(para)
    if not runs:
        return
    runs[0].text = token
    for r in runs[1:]:
        r._r.getparent().remove(r._r)


def baue_docx_vorlage(muster_pfad, ziel_pfad) -> None:
    doc = Document(str(muster_pfad))
    paras = doc.paragraphs

    # Plausibilitätsprüfung: erwartete Ankertexte
    if "Hrsg." not in paras[1].text or "ISBN" not in paras[16].text:
        raise ValueError(f"Unerwartete Absatzstruktur in {muster_pfad} — "
                         f"Vorlagen-Mapping passt nicht.")

    # Blurb-Folgeabsätze (P8, P9) merken und später entfernen
    zu_loeschen = [paras[8]._p, paras[9]._p]
    for idx, token in _VORLAGE_TOKENS.items():
        _setze_token(paras[idx], token)
    for p in zu_loeschen:
        p.getparent().remove(p)

    # Cover-Zuschnitt (srcRect) entfernen -> Laufzeit setzt echtes Verhältnis
    drawing = _cover_zeichnung(doc)
    if drawing is not None:
        for sr in drawing.findall(".//" + qn("a:srcRect")):
            sr.getparent().remove(sr)

    ziel_pfad = Path(ziel_pfad)
    ziel_pfad.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(ziel_pfad))


def baue_docx_vorlagen(beispiel_dir=None, ziel_dir=None) -> None:
    """Erzeugt vorlagen/pi_vorlage.docx und bi_vorlage.docx aus den Mustern."""
    beispiel_dir = Path(beispiel_dir or (_base_dir() / "beispiele"
                        / "Bausteine Bruchsal, Bd. 5"))
    ziel_dir = Path(ziel_dir or VORLAGEN_DIR)
    baue_docx_vorlage(beispiel_dir / "PI_05-559-2.docx",
                      ziel_dir / "pi_vorlage.docx")
    baue_docx_vorlage(beispiel_dir / "BI_05-559-2.docx",
                      ziel_dir / "bi_vorlage.docx")
    print(f"Vorlagen geschrieben nach {ziel_dir}")


# ---------------------------------------------------------------------
# Kompletter Lauf ohne Oberfläche
# ---------------------------------------------------------------------

def erzeuge_alle(buch: "Buchdaten", ziel_ordner, cfg: dict, *,
                 cover_bytes: bytes | None = None,
                 cover_suffix: str = ".jpg",
                 detail_url: str = "",
                 log=print) -> list[Path]:
    """PI und BI erzeugen — als .docx UND .html, dazu das Cover.

    Bis hierher stand diese Reihenfolge samt Dateinamen in ``app._do_run`` und
    war damit von außen nicht anstoßbar. Jetzt hier, und die Oberfläche ruft
    dieselbe Funktion — sonst laufen die beiden Wege mit der Zeit auseinander.

    PI = Presseinformation (Anschreiben + Rezensionsexemplar),
    BI = Buchinformation (Buchhandel, mit Konditionen und Bestellformular).

    Ohne ``cover_bytes`` behalten die .docx das Platzhalter-Cover der Vorlage;
    erzeugt wird trotzdem.

    Rückgabe: Liste der geschriebenen Dateien.
    """
    ziel = Path(ziel_ordner)
    ziel.mkdir(parents=True, exist_ok=True)
    sc = buch.shortcode
    detail_url = detail_url or cfg.get("detail_fallback_url", "")

    v = VORLAGEN_DIR
    geschrieben: list[Path] = []

    # Namen mit Unterstrich — früher hießen die HTML-Dateien "PI 05-559-2.html"
    # (mit Leerzeichen), die docx aber "PI_05-559-2.docx". Ein Buchordner, in
    # dem dieselbe Sache zweimal anders heißt, ist eine Stolperfalle.
    for vorlage, name in (("pi_vorlage.docx", f"PI_{sc}.docx"),
                          ("bi_vorlage.docx", f"BI_{sc}.docx")):
        ziel_datei = ziel / name
        log(f"Erzeuge {name} …")
        generiere_docx(v / vorlage, buch, cover_bytes, ziel_datei)
        geschrieben.append(ziel_datei)

    for vorlage, name in (("pi_vorlage.html", f"PI_{sc}.html"),
                          ("bi_vorlage.html", f"BI_{sc}.html")):
        ziel_datei = ziel / name
        log(f"Erzeuge {name} …")
        generiere_html(v / vorlage, buch, detail_url, cfg, ziel_datei)
        geschrieben.append(ziel_datei)

    if cover_bytes:
        ziel_datei = ziel / f"cover_{buch.isbn13}{cover_suffix or '.jpg'}"
        ziel_datei.write_bytes(cover_bytes)
        geschrieben.append(ziel_datei)
    else:
        log("⚠ Kein Cover — die .docx behalten das Platzhalter-Cover.")

    return geschrieben
