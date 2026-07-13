"""
Cover-Previews Core-Logik
=========================
Reine Datenlogik ohne UI-Abhängigkeiten — liest einen druckfertigen
Umschlag-PDF (Rückseite | Rücken | Vorderseite, plus Einschlag/Beschnitt),
erkennt die Schnitt-/Falzmarken und erzeugt daraus Vorschau-Bilder.

Ablauf:
1. oeffne_pdf(pfad)                       -> fitz.Document
2. finde_schnittlinien(doc)               -> Regionen (Marken-Auswertung)
3. rendere_seite(doc, dpi)                -> PIL.Image (Rasterung)
4. extrahiere(bild, box_px)               -> Ausschnitt (Vorderseite etc.)
5. speichere_2d(...) / erzeuge_3d_photoshop(...)

Die Schnittmarken stehen im PDF als Vektor-Striche. Erkennungsregel:
Eine vertikale Marke, die den oberen/unteren Blattrand berührt, markiert eine
X-Schnittlinie; eine horizontale Marke am linken/rechten Blattrand eine
Y-Schnittlinie. Innenliegende Element-Marken (Logos o. Ä.) berühren keinen
Blattrand und werden dadurch ignoriert.
"""

from __future__ import annotations

import io
import json
import os
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

import fitz  # PyMuPDF
from PIL import Image, ImageDraw

PT_PRO_ZOLL = 72.0


# ---------------------------------------------------------------------
# Ordner (Config, Ausgabe) — Muster wie die anderen Tools
# ---------------------------------------------------------------------

def _base_dir() -> Path:
    """Ordner neben der .exe (PyInstaller-Build) bzw. neben dem Tool-Code."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


# App-mutierte Laufzeitdaten (config.json, cover_output/) liegen direkt neben
# der .exe — wie bei den anderen Tools, kein data-Unterordner.
APP_DIR = _base_dir()
APP_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_PFAD = APP_DIR / "config.json"


# ---------------------------------------------------------------------
# Konfiguration
# ---------------------------------------------------------------------

DEFAULT_CONFIG = {
    # Ablage der fertigen Bilder: je Buch ein Ordner unter artikeldaten_dir
    # (Konvention dort: "<Kurzcode>_<Titel>", z. B. "05-597-4_Oberkirch").
    # Leer / nicht erreichbar -> Rückfall auf output_dir neben der .exe.
    "artikeldaten_dir": r"\\C019\d\Online\Webseite\Artikeldaten",
    "output_dir": "cover_output",
    "last_input_dir": "",
    # Einband: "hardcover" zeigt die Falz-Ebene (Rille am Buchdeckel),
    # "softcover" blendet sie aus.
    "einband": "hardcover",
    # ISBN-Präfix des Verlags (für Kurzcode-Ableitung, falls im PDF gefunden).
    "isbn_prefix": "978-3-95505",
    # DPI-Ziele
    "dpi_print": 300,
    "dpi_web": 72,
    "web_max_px": 800,
    # Reihenfolge des Spreads: "back_front" = Rückseite links, Vorderseite
    # rechts (deutscher Standard). "front_back" spiegelt die Zuordnung.
    "spread_reihenfolge": "back_front",
    # Marken-Erkennung
    "marken_band_pt": 60.0,    # wie weit reicht der äußere Rand, in dem Marken
                               # liegen dürfen (manche Setzer setzen sie versetzt
                               # ab statt bis an die Blattkante zu ziehen)
    "marken_max_rel": 0.15,    # eine Marke ist kurz gegenüber der Seite
    "kante_tol_pt": 3.0,       # nur noch für den Rückfall: Marke berührt die Kante
    "cluster_tol_pt": 4.0,     # Marken mit ~gleicher Position zusammenfassen
    # Dateinamen-Muster (Konvention des Verlags — auf dem Artikeldaten-Share
    # sind 344 von 345 2D-Bildern .jpg, nicht .jpeg; mit .jpeg würden die
    # neuen Dateien neben den alten liegen statt sie zu ersetzen)
    "muster_2d": "2D_{dpi}_{sc}.jpg",
    "muster_3d": "3D_{dpi}_{sc}.jpg",
    "muster_3d_png": "{sc}.png",
    "jpeg_qualitaet": 95,
    # 3D-Mockup-Vorlagen (ein PSD je Buchformat, im Ordner _NEU_Vorlage)
    "vorlagen_dir": "",            # leer -> mitgelieferter/lokaler _NEU_Vorlage-Ordner
    "vorlagen_tol_cm": 1.0,        # Toleranz beim Format-Abgleich
    "rand_cm": 1.5,               # Rand oben/links/rechts ums Buch
    "rand_unten_cm": 0.0,         # unten: 0 = das Bild endet, wo die Spiegelung
                                  # ausläuft (sonst schwebt das Buch über Weiß)
    # weißer Umschlag: Anteil heller, blasser Randpixel, ab dem korrigiert wird
    "weiss_schwellwert": 0.85,
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
# 3D-Vorlagen-Registry (Buchformat -> Mockup-PSD)
# ---------------------------------------------------------------------

def _asset_pfad(name: str) -> Path:
    """Nur-Lese-Asset (z. B. vorlagen_map.json) neben dem Modul bzw. im
    PyInstaller-Bundle (sys._MEIPASS)."""
    if getattr(sys, "frozen", False):
        p = Path(getattr(sys, "_MEIPASS", "")) / "cover_previews" / name
        if p.exists():
            return p
    return Path(__file__).parent / name


def lade_vorlagen_map() -> dict:
    """Vorab erzeugte Zuordnung Vorlage -> {format_cm, spine, slots, hide_bg}.
    Die Layer-IDs darin stammen aus einer einmaligen PSD-Analyse."""
    p = _asset_pfad("vorlagen_map.json")
    if not p.exists():
        return {}
    with open(p, "r", encoding="utf-8") as f:
        return json.load(f)


def vorlagen_dir(cfg: dict) -> Path:
    """Ordner mit den Mockup-PSDs (_NEU_Vorlage). Per config einstellbar, sonst
    neben der .exe/dem Modul."""
    raw = (cfg or {}).get("vorlagen_dir") or ""
    if raw:
        return Path(os.path.expandvars(raw)).expanduser()
    basis = _base_dir() if getattr(sys, "frozen", False) else Path(__file__).parent
    return basis / "_NEU_Vorlage"


def vorlagen_liste(cfg: dict) -> list[dict]:
    """Alle bekannten Vorlagen mit Format + Pfad (inkl. Nicht-Auto wie EBOOK).

    ``book_cm`` sind die wahren Buchmaße (Breite, Höhe); ``format_cm`` ist nur der
    Name der Vorlage und mal B×H, mal H×B (siehe psd_analyse.py). Für Vergleiche
    also immer ``book_cm`` nehmen.
    """
    vd = vorlagen_dir(cfg)
    liste = []
    for name, t in lade_vorlagen_map().items():
        liste.append({
            "name": name, "pfad": vd / name, "format_cm": t.get("format_cm"),
            "book_cm": t.get("book_cm"), "px_pro_cm": t.get("content_px_per_cm"),
            "slots": t.get("slots", []), "hide_bg": t.get("hide_bg", []),
            "auto": bool(t.get("auto")), "spine": bool(t.get("spine")),
        })
    return liste


def front_masse_cm(reg: "Regionen") -> tuple[float, float] | None:
    if reg.front is None:
        return None
    x0, y0, x1, y1 = reg.front
    return ((x1 - x0) / PT_PRO_ZOLL * 2.54, (y1 - y0) / PT_PRO_ZOLL * 2.54)


def spine_dicke_cm(reg: "Regionen") -> float | None:
    """Echte Rückendicke aus den Schnittlinien (cm)."""
    if reg.spine is None:
        return None
    x0, _, x1, _ = reg.spine
    return (x1 - x0) / PT_PRO_ZOLL * 2.54


def waehle_vorlage(reg: "Regionen", cfg: dict) -> dict | None:
    """Wählt anhand des erkannten Vorderseiten-Trims (in cm) die nächstliegende
    Auto-Vorlage. Rückgabe: Listeneintrag ergänzt um 'dist_cm' und
    'im_toleranz' — oder None, wenn keine Auto-Vorlage vorhanden ist."""
    masse = front_masse_cm(reg)
    if masse is None:
        return None
    w_cm, h_cm = masse
    tol = float(cfg.get("vorlagen_tol_cm", 1.0)) * (2 ** 0.5)
    best, best_d = None, None
    for eintrag in vorlagen_liste(cfg):
        if not eintrag["auto"] or not eintrag["book_cm"]:
            continue
        fw, fh = eintrag["book_cm"]
        d = ((fw - w_cm) ** 2 + (fh - h_cm) ** 2) ** 0.5
        if best_d is None or d < best_d:
            best, best_d = dict(eintrag), d
    if best is not None:
        best["dist_cm"] = round(best_d, 2)
        best["im_toleranz"] = best_d <= tol
    return best


# ---------------------------------------------------------------------
# PDF öffnen / rastern
# ---------------------------------------------------------------------

def oeffne_pdf(pfad) -> fitz.Document:
    return fitz.open(str(pfad))


def seite_masse(doc: fitz.Document) -> tuple[float, float]:
    """Seitengröße der ersten Seite in Punkt (Breite, Höhe)."""
    r = doc[0].rect
    return (r.width, r.height)


def rendere_seite(doc: fitz.Document, dpi: int) -> Image.Image:
    """Rastert die erste Seite bei gegebener DPI zu einem PIL-Bild (RGB)."""
    pix = doc[0].get_pixmap(dpi=dpi, alpha=False)
    return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)


# ---------------------------------------------------------------------
# Schnittmarken-Erkennung
# ---------------------------------------------------------------------

@dataclass
class Regionen:
    """Alle Maße in PDF-Punkt (1/72 Zoll), Ursprung oben links."""
    seite_pt: tuple[float, float]
    x_cuts: list[float] = field(default_factory=list)   # vertikale Schnitte (x)
    y_cuts: list[float] = field(default_factory=list)   # horizontale Schnitte (y)
    back: tuple | None = None      # (x0, y0, x1, y1)
    spine: tuple | None = None
    front: tuple | None = None
    content: tuple | None = None   # gesamter Trim-Spread (Rücken/Vorder/Rück)

    def box_px(self, name: str, dpi: int) -> tuple[int, int, int, int]:
        """Region als Pixel-Box (left, top, right, bottom) bei gegebener DPI."""
        box = getattr(self, name)
        if box is None:
            raise ValueError(f"Region '{name}' nicht bestimmt")
        s = dpi / PT_PRO_ZOLL
        x0, y0, x1, y1 = box
        return (round(x0 * s), round(y0 * s), round(x1 * s), round(y1 * s))


def _liniensegmente(page: fitz.Page) -> list[tuple[float, float, float, float]]:
    segs = []
    for dr in page.get_drawings():
        for it in dr["items"]:
            if it[0] == "l":
                a, b = it[1], it[2]
                segs.append((a.x, a.y, b.x, b.y))
            elif it[0] == "re":
                r = it[1]
                segs.extend([
                    (r.x0, r.y0, r.x1, r.y0), (r.x0, r.y1, r.x1, r.y1),
                    (r.x0, r.y0, r.x0, r.y1), (r.x1, r.y0, r.x1, r.y1),
                ])
    return segs


def _cluster(werte: list[float], tol: float) -> list[float]:
    """Nahe beieinander liegende Positionen zu je einem Mittelwert bündeln."""
    if not werte:
        return []
    werte = sorted(werte)
    gruppen = [[werte[0]]]
    for v in werte[1:]:
        if v - gruppen[-1][-1] <= tol:
            gruppen[-1].append(v)
        else:
            gruppen.append([v])
    return [sum(g) / len(g) for g in gruppen]


def _breiteste_luecken(cuts: list[float]) -> list[int]:
    """Indizes der Lücken zwischen aufeinanderfolgenden Schnitten, absteigend
    nach Breite (Lücke i liegt zwischen cuts[i] und cuts[i+1])."""
    luecken = [(cuts[i + 1] - cuts[i], i) for i in range(len(cuts) - 1)]
    luecken.sort(reverse=True)
    return [i for _, i in luecken]


def _gepaart(a: list[float], b: list[float], tol: float) -> list[float]:
    """Positionen, die in BEIDEN Listen vorkommen (gemittelt).

    Eine echte Schnittmarke steht immer doppelt: oben und unten (bzw. links und
    rechts) auf derselben Höhe. Grafik im Umschlag tut das so gut wie nie —
    darum ist die Paarung das Merkmal, das Marken von Motiv trennt.
    """
    treffer = []
    for v in a:
        partner = [w for w in b if abs(v - w) <= tol]
        if partner:
            treffer.append((v + sum(partner) / len(partner)) / 2)
    return treffer


def _marken(segs, W: float, H: float, cfg: dict) -> tuple[list[float], list[float]]:
    """Schnittmarken aus den Vektor-Strichen lesen.

    Marken liegen im äußeren Rand des Blattes (außerhalb des Anschnitts) und sind
    kurz. Ob sie den Blattrand berühren, ist NICHT verlässlich: manche Setzer
    ziehen sie bis an die Kante (US_Kleindenkmale), andere setzen sie versetzt ab
    (US_Kellerkinder — dort beginnt die Marke erst 9 pt unter der Blattkante).

    Es reicht aber auch nicht, einfach alles am Rand zu nehmen: der Barcode sitzt
    bei US_Kellerkinder rund 100 pt über der Unterkante und bestünde diese Prüfung.
    Ausschlaggebend ist die Paarung oben/unten (siehe _gepaart).
    """
    band = float(cfg.get("marken_band_pt", 60.0))
    max_rel = float(cfg.get("marken_max_rel", 0.15))
    ctol = float(cfg.get("cluster_tol_pt", 4.0))

    oben, unten, links, rechts = [], [], [], []
    for x0, y0, x1, y1 in segs:
        senkrecht = abs(x0 - x1) < 0.5 and abs(y0 - y1) >= 0.5
        waagerecht = abs(y0 - y1) < 0.5 and abs(x0 - x1) >= 0.5
        if senkrecht:
            a, b = min(y0, y1), max(y0, y1)
            if b - a > H * max_rel:
                continue                       # langer Strich: Rahmen, keine Marke
            if b <= band:
                oben.append(x0)
            elif a >= H - band:
                unten.append(x0)
        elif waagerecht:
            a, b = min(x0, x1), max(x0, x1)
            if b - a > W * max_rel:
                continue
            if b <= band:
                links.append(y0)
            elif a >= W - band:
                rechts.append(y0)

    x_cuts = _gepaart(_cluster(oben, ctol), _cluster(unten, ctol), ctol)
    y_cuts = _gepaart(_cluster(links, ctol), _cluster(rechts, ctol), ctol)
    return x_cuts, y_cuts


def _marken_am_blattrand(segs, W: float, H: float,
                         cfg: dict) -> tuple[list[float], list[float]]:
    """Rückfall: nur Marken, die den Blattrand wirklich berühren."""
    tol = float(cfg.get("kante_tol_pt", 3.0))
    ctol = float(cfg.get("cluster_tol_pt", 4.0))
    x_marks, y_marks = [], []
    for x0, y0, x1, y1 in segs:
        senkrecht = abs(x0 - x1) < 0.5
        waagerecht = abs(y0 - y1) < 0.5
        if senkrecht and not waagerecht:
            if min(y0, y1) <= tol or max(y0, y1) >= H - tol:
                x_marks.append(x0)
        elif waagerecht and not senkrecht:
            if min(x0, x1) <= tol or max(x0, x1) >= W - tol:
                y_marks.append(y0)
    return _cluster(x_marks, ctol), _cluster(y_marks, ctol)


def finde_schnittlinien(doc: fitz.Document, cfg: dict | None = None) -> Regionen:
    """Wertet die Schnitt-/Falzmarken aus und leitet die Regionen ab.

    Zwei Setzweisen kommen vor: Marken bis an die Blattkante gezogen, oder mit
    Abstand davor abgesetzt. _marken() deckt beide ab; findet es keine vier
    senkrechten Schnitte, greift die alte, strengere Regel als Rückfall.
    """
    cfg = cfg or DEFAULT_CONFIG
    page = doc[0]
    W, H = page.rect.width, page.rect.height
    segs = _liniensegmente(page)

    x_cuts, y_cuts = _marken(segs, W, H, cfg)
    if len(x_cuts) < 4 or len(y_cuts) < 2:
        alt_x, alt_y = _marken_am_blattrand(segs, W, H, cfg)
        if len(alt_x) >= len(x_cuts) and len(alt_y) >= len(y_cuts):
            x_cuts, y_cuts = alt_x, alt_y
    return regionen_aus_cuts(x_cuts, y_cuts, (W, H), cfg)


def regionen_aus_cuts(x_cuts: list[float], y_cuts: list[float],
                      seite_pt: tuple[float, float],
                      cfg: dict | None = None) -> Regionen:
    """Leitet Back/Spine/Front/Content aus den (ggf. per GUI korrigierten)
    Schnittlisten ab. Getrennt von der Marken-Erkennung, damit die GUI nach
    einem Nudge neu rechnen kann."""
    cfg = cfg or DEFAULT_CONFIG
    W, H = seite_pt
    x_cuts = sorted(x_cuts)
    y_cuts = sorted(y_cuts)
    reg = Regionen(seite_pt=(W, H), x_cuts=x_cuts, y_cuts=y_cuts)

    # Y: größte Lücke = Inhaltshöhe (oberer/unterer Trim).
    if len(y_cuts) >= 2:
        gy = _breiteste_luecken(y_cuts)[0]
        cy0, cy1 = y_cuts[gy], y_cuts[gy + 1]
    else:
        cy0, cy1 = 0.0, H

    # X: die zwei breitesten Lücken = die beiden Buchdeckel; dazwischen Rücken.
    if len(x_cuts) >= 4:
        g = _breiteste_luecken(x_cuts)
        i, j = sorted(g[:2])          # linke und rechte Deckel-Lücke
        links = (x_cuts[i], x_cuts[i + 1])
        rechts = (x_cuts[j], x_cuts[j + 1])
        spine = (x_cuts[i + 1], x_cuts[j])   # zwischen den Deckeln
        if cfg.get("spread_reihenfolge", "back_front") == "back_front":
            back_x, front_x = links, rechts
        else:
            back_x, front_x = rechts, links
        reg.back = (back_x[0], cy0, back_x[1], cy1)
        reg.front = (front_x[0], cy0, front_x[1], cy1)
        reg.spine = (spine[0], cy0, spine[1], cy1)
        reg.content = (links[0], cy0, rechts[1], cy1)
    return reg


# ---------------------------------------------------------------------
# Ausschnitt / Overlay
# ---------------------------------------------------------------------

def extrahiere(bild: Image.Image, box_px: tuple[int, int, int, int]) -> Image.Image:
    return bild.crop(box_px)

def overlay_marken(bild: Image.Image, reg: Regionen, dpi: int) -> Image.Image:
    """Zeichnet die erkannten Schnittlinien + Regionen zur Kontrolle ein."""
    s = dpi / PT_PRO_ZOLL
    out = bild.convert("RGB").copy()
    d = ImageDraw.Draw(out)
    for x in reg.x_cuts:
        d.line([(x * s, 0), (x * s, out.height)], fill=(0, 150, 255), width=2)
    for y in reg.y_cuts:
        d.line([(0, y * s), (out.width, y * s)], fill=(0, 150, 255), width=2)
    farben = {"back": (0, 200, 0), "spine": (255, 140, 0), "front": (220, 0, 0)}
    for name, farbe in farben.items():
        box = getattr(reg, name)
        if box:
            d.rectangle([c * s for c in box], outline=farbe, width=4)
    return out


# ---------------------------------------------------------------------
# Weißer-Umschlag-Erkennung
# ---------------------------------------------------------------------

def ist_weisser_umschlag(front: Image.Image, cfg: dict | None = None) -> bool:
    """Heuristik: ist die Vorderseite überwiegend (nahezu) weiß? Dann braucht
    das 3D-Mockup eine leichte Grau-Korrektur, damit sich das Buch vom weißen
    Hintergrund abhebt. Gemessen wird ein Randstreifen der Vorderseite: hohe
    Helligkeit + geringe Sättigung = weiß."""
    cfg = cfg or DEFAULT_CONFIG
    schwelle = float(cfg.get("weiss_schwellwert", 0.85))
    im = front.convert("RGB")
    w, h = im.size
    klein = im.resize((max(1, w // 4), max(1, h // 4)), Image.BILINEAR)
    w2, h2 = klein.size
    r2 = max(1, int(min(w2, h2) * 0.06))
    px = klein.load()
    hell = n = 0
    for y in range(h2):
        for x in range(w2):
            if x < r2 or x >= w2 - r2 or y < r2 or y >= h2 - r2:
                r, g, b = px[x, y]
                mx, mn = max(r, g, b), min(r, g, b)
                sat = 0.0 if mx == 0 else (mx - mn) / mx
                if mx >= 235 and sat <= 0.06:
                    hell += 1
                n += 1
    return n > 0 and (hell / n) >= schwelle


# ---------------------------------------------------------------------
# 2D-Ausgabe (JPEG; 300 & 72 dpi = gleiche Pixel, nur DPI-Eintrag)
# ---------------------------------------------------------------------

def speichere_2d(front: Image.Image, out_dir: Path, sc: str, cfg: dict) -> list[Path]:
    """Speichert die 2D-Vorderseite als JPEG in Druck- und Web-DPI. Wie im
    manuellen Ablauf ('neu berechnen aus') sind beide Dateien pixelgleich —
    nur der eingebettete DPI-Wert unterscheidet sich."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    muster = cfg.get("muster_2d", "2D_{dpi}_{sc}.jpeg")
    q = int(cfg.get("jpeg_qualitaet", 95))
    rgb = front.convert("RGB")
    pfade = []
    for dpi in (int(cfg.get("dpi_print", 300)), int(cfg.get("dpi_web", 72))):
        p = out_dir / muster.format(dpi=dpi, sc=sc)
        rgb.save(p, "JPEG", quality=q, dpi=(dpi, dpi))
        pfade.append(p)
    return pfade


# ---------------------------------------------------------------------
# ISBN / Kurzcode
# ---------------------------------------------------------------------

def extrahiere_isbn(doc: fitz.Document) -> str | None:
    txt = doc[0].get_text()
    m = re.search(r"ISBN[\s:]*((?:97[89][\s-]?)(?:\d[\s-]?){9}\d)", txt)
    if not m:
        m = re.search(r"\b(97[89](?:[\s-]?\d){10})\b", txt)
    if not m:
        return None
    ziffern = re.sub(r"\D", "", m.group(1))
    return ziffern if len(ziffern) == 13 else None


def shortcode_aus_isbn(isbn13: str) -> str:
    """z. B. 9783955055721 -> '05-572-1' (wie im pi_bi_generator)."""
    e = re.sub(r"\D", "", isbn13)
    return f"{e[7:9]}-{e[9:12]}-{e[12]}"


# ---------------------------------------------------------------------
# Zielordner (Artikeldaten-Share)
# ---------------------------------------------------------------------

def artikeldaten_dir(cfg: dict) -> Path | None:
    """Basis-Ordner auf dem Share — None, wenn nicht gesetzt oder nicht erreichbar."""
    raw = (cfg or {}).get("artikeldaten_dir") or ""
    if not raw:
        return None
    p = Path(os.path.expandvars(raw))
    try:
        return p if p.is_dir() else None
    except OSError:          # Netzpfad nicht erreichbar
        return None


def finde_artikel_ordner(sc: str, cfg: dict) -> Path | None:
    """Vorhandenen Ordner zum Kurzcode suchen ("05-597-4_Oberkirch")."""
    basis = artikeldaten_dir(cfg)
    if not basis or not sc:
        return None
    for p in sorted(basis.glob(f"{sc}*")):
        if p.is_dir():
            return p
    return None


def ordner_name(sc: str, titel: str = "") -> str:
    """Ordnername nach Konvention: Kurzcode, optional mit Titel."""
    titel = re.sub(r'[<>:"/\\|?*]', "", (titel or "").strip()).strip(" ._")
    return f"{sc}_{titel}" if titel else sc


def ziel_ordner(sc: str, titel: str, cfg: dict) -> tuple[Path, bool]:
    """Zielordner aus Kurzcode + Titel + ob er schon existiert.

    Maßgeblich ist immer der eingegebene Name — wer den Titel ändert, bekommt
    auch einen anderen Ordner. (Ein vorhandener Ordner zum Kurzcode wird beim
    Einlesen nur als Vorschlag ins Titelfeld übernommen, siehe
    finde_artikel_ordner; er überstimmt die Eingabe nicht.)

    Ist der Share nicht erreichbar, fällt es auf output_dir neben der .exe
    zurück, damit die Arbeit nicht verloren geht.
    """
    basis = artikeldaten_dir(cfg)
    if basis is None:
        return APP_DIR / cfg.get("output_dir", "cover_output"), False
    p = basis / ordner_name(sc, titel)
    return p, p.is_dir()


def ausgabe_namen(sc: str, cfg: dict, mit_2d: bool, mit_3d: bool) -> list[str]:
    """Dateinamen, die ein Lauf schreiben würde (Konvention des Verlags)."""
    namen = []
    dpi_p, dpi_w = int(cfg.get("dpi_print", 300)), int(cfg.get("dpi_web", 72))
    if mit_2d:
        m = cfg.get("muster_2d", "2D_{dpi}_{sc}.jpeg")
        namen += [m.format(dpi=d, sc=sc) for d in (dpi_p, dpi_w)]
    if mit_3d:
        m = cfg.get("muster_3d", "3D_{dpi}_{sc}.jpg")
        namen += [m.format(dpi=d, sc=sc) for d in (dpi_p, dpi_w)]
        namen.append(cfg.get("muster_3d_png", "{sc}.png").format(sc=sc))
    return namen


def kollisionen(out_dir: Path, namen: list[str]) -> list[Path]:
    """Welche der Zieldateien liegen dort schon?"""
    out_dir = Path(out_dir)
    return [out_dir / n for n in namen if (out_dir / n).exists()]


def sichere_weg(dateien: list[Path], zeitstempel: str) -> Path:
    """Vorhandene Dateien nach _alt/<Zeitstempel>/ verschieben.

    So bleiben die neuen Dateien konventionsgerecht benannt und die alten sind
    trotzdem nicht weg. ``zeitstempel`` wird übergeben (nicht hier gebildet),
    damit ein Lauf alles in denselben Unterordner legt.
    """
    if not dateien:
        return None
    ziel = Path(dateien[0]).parent / "_alt" / zeitstempel
    ziel.mkdir(parents=True, exist_ok=True)
    for f in dateien:
        f = Path(f)
        f.replace(ziel / f.name)
    return ziel


# ---------------------------------------------------------------------
# 3D-Mockup über Photoshop (Windows, COM)
# ---------------------------------------------------------------------

def region_box(reg: Regionen, name: str) -> tuple:
    """Punkt-Box für einen Regionsnamen inkl. Kombi-Region 'front_spine'."""
    einfach = {"front": reg.front, "spine": reg.spine,
               "back": reg.back, "content": reg.content, "full": reg.content}
    if name in einfach:
        box = einfach[name]
        if box is None:
            raise ValueError(f"Region '{name}' nicht bestimmt")
        return box
    if name == "front_spine":
        if not (reg.front and reg.spine):
            raise ValueError("front/spine nicht bestimmt")
        fx0, fy0, fx1, fy1 = reg.front
        sx0, _, sx1, _ = reg.spine
        return (min(fx0, sx0), fy0, max(fx1, sx1), fy1)
    raise ValueError(f"Unbekannte Region: {name}")


def _ausschnitt(reg: Regionen, bild_hi: Image.Image, region: str,
                dpi: int) -> Image.Image:
    s = dpi / PT_PRO_ZOLL
    box_px = tuple(round(c * s) for c in region_box(reg, region))
    return extrahiere(bild_hi, box_px)


def _stage_cover(reg: Regionen, bild_hi: Image.Image, dpi: int, out_dir: Path,
                 sc: str, size_px: tuple[int, int], so_dpi: float) -> Path:
    """Vorderseite exakt auf die Slot-Größe bringen.

    Photoshop bildet den Inhalt eines Smart-Objekts über einen festen Transform
    auf die 3D-Fläche ab; die Maße des Inhalts sind dabei der Maßstab. Ein Inhalt
    in Slot-Größe füllt die Fläche also genau — auch wenn der PDF-Trim leicht vom
    Vorlagenformat abweicht (die Verzerrung liegt typisch unter 2 %).

    ``so_dpi`` ist die Auflösung der Vorlage: Photoshop setzt den Inhalt nach
    PHYSISCHER Größe ein (px / dpi), nicht nach Pixelzahl. Ein 300-dpi-Tag in
    einer 367,59-dpi-Vorlage landet um Faktor 300/367,59 daneben.
    """
    img = _ausschnitt(reg, bild_hi, "front", dpi).resize(tuple(size_px), Image.LANCZOS)
    p = Path(out_dir) / f"_slot_{sc}_front_{size_px[0]}x{size_px[1]}.png"
    img.save(p, dpi=(so_dpi, so_dpi))
    return p


def _stage_spine(reg: Regionen, bild_hi: Image.Image, dpi: int, out_dir: Path,
                 sc: str, size_px: tuple[int, int], k: float, anchor: str,
                 so_dpi: float, log=print) -> Path:
    """Rücken maßgetreu in die volle Slot-Leinwand setzen.

    Die Rücken-Slots der Vorlagen sind bewusst ÜBERBREIT (bei 17x24 z. B. 10 cm)
    und tragen ihr Motiv nur in einem schmalen, an der Cover-Kante ausgerichteten
    Streifen — der Rest ist transparent. Die Leinwand nimmt damit jede Rückendicke
    auf. Genau so füllen wir sie: das Motiv ``Dicke × k`` px breit, an der
    ``anchor``-Kante (der Kante zum Cover), Rest transparent.

    Der Inhalt behält dadurch seine Slot-Größe — der Transform des Smart-Objekts
    bleibt unverändert und der Rücken sitzt weiter bündig am Buch. Würde man
    stattdessen nur den schmalen Streifen liefern, skalierte Photoshop den
    Transform auf dessen Maße und der Rücken löste sich vom Cover.
    """
    breite, hoehe = int(size_px[0]), int(size_px[1])
    dicke = spine_dicke_cm(reg) or 0.0
    motiv_px = max(1, round(dicke * k))
    if motiv_px > breite:
        log(f"Warnung: Rücken {dicke:.2f} cm passt nicht in den Slot "
            f"({breite / k:.2f} cm) — wird beschnitten.")
        motiv_px = breite

    motiv = _ausschnitt(reg, bild_hi, "spine", dpi).resize((motiv_px, hoehe),
                                                           Image.LANCZOS)
    leinwand = Image.new("RGBA", (breite, hoehe), (0, 0, 0, 0))
    x = breite - motiv_px if anchor == "right" else 0
    leinwand.paste(motiv.convert("RGBA"), (x, 0))

    p = Path(out_dir) / f"_slot_{sc}_spine_{breite}x{hoehe}.png"
    leinwand.save(p, dpi=(so_dpi, so_dpi))
    log(f"Rücken {dicke:.2f} cm -> {motiv_px} px Motiv in {breite} px Leinwand "
        f"({k:.1f} px/cm, bündig {anchor}).")
    return p


def _stage_slots(reg: Regionen, bild_hi: Image.Image, eintrag: dict, dpi: int,
                 out_dir: Path, sc: str, braucht_spine: bool,
                 log=print) -> list[tuple[int, Path]]:
    """Legt je Smart-Object-Slot ein PNG in dessen Slot-Größe an.

    Jeder Inhalt behält exakt die Größe seines Slots — nur so bleibt der Transform
    des Smart-Objekts unverändert und das Buch zusammen. Der Inhaltsraum hat je
    Vorlage eine eigene Auflösung ``k`` px/cm (162–310 dpi, siehe psd_analyse.py).

    Buch und Spiegelung sind eigene Slots und können unterschiedliche Größen haben
    (29x22 hat zwei Cover-Maße) — gleiche Zielgrößen werden nur einmal gerendert.
    """
    k = eintrag.get("content_px_per_cm")
    so_dpi = eintrag.get("content_dpi")
    if not k or not so_dpi:
        raise ValueError(
            "Vorlage ohne content_px_per_cm/content_dpi — vorlagen_map.json mit "
            "'python -m cover_previews.psd_analyse' neu erzeugen.")

    cache: dict[tuple[str, int, int], Path] = {}
    slots = []
    for slot in eintrag.get("slots", []):
        rolle = slot["role"]
        if rolle == "spine" and not braucht_spine:
            continue
        size = slot.get("size_px")
        if not size:
            raise ValueError(
                f"Slot {slot['layer_id']} hat keine size_px — vorlagen_map.json "
                "mit 'python -m cover_previews.psd_analyse' neu erzeugen.")

        schluessel = (rolle, int(size[0]), int(size[1]))
        if schluessel not in cache:
            if rolle == "spine":
                cache[schluessel] = _stage_spine(
                    reg, bild_hi, dpi, out_dir, sc, size, k,
                    slot.get("anchor", "right"), so_dpi, log)
            else:
                cache[schluessel] = _stage_cover(reg, bild_hi, dpi, out_dir, sc,
                                                 size, so_dpi)
        slots.append((slot["layer_id"], cache[schluessel]))
    return slots


def _jp(p) -> str:
    """Pfad für ExtendScript (Forward-Slashes, Anführungszeichen escapen)."""
    return str(p).replace("\\", "/").replace('"', '\\"')


def _baue_jsx(psd_pfad, eintrag: dict, slot_pngs: list[tuple[int, Path]],
              weiss: bool, png_transp, jpg_pfade: list[tuple[int, Path]],
              rand_cm: float, rand_unten_cm: float = 0.0,
              hardcover: bool = True) -> str:
    """ExtendScript für das 3D-Mockup:
    - Smart-Objekte per Layer-ID ersetzen; ``slot_pngs`` ist die Liste
      (layer_id, PNG) — jedes PNG exakt in der Größe seines Slots, sodass der
      Transform des Smart-Objekts unverändert bleibt;
    - Falz-Ebenen (Rille am Buchdeckel) nur beim Hardcover einblenden;
    - optional Weiß-Korrektur (Selektive Farbkorrektur: Weiß +10 % Schwarz);
    - Hintergrund ausblenden, transparent zuschneiden -> PNG;
    - Hintergrund wieder ein, flach, JPEG in 300 & 72 dpi (nur DPI-Tag).

    Die Vorlage wird direkt geöffnet (nicht kopiert) und im ``finally`` immer
    mit DONOTSAVECHANGES geschlossen — auch wenn ein Schritt scheitert. Sonst
    bliebe das Master-PSD mit ersetzten Smart-Objekten offen und ungespeichert
    in Photoshop stehen, wo ein beiläufiges Strg+S es überschreiben würde.
    """
    ersetzungen = [f'    replaceById({layer_id}, new File("{_jp(datei)}"));'
                   for layer_id, datei in slot_pngs]
    ersetzen = "\n".join(ersetzungen)
    # Sichtbarkeit der Falz IMMER explizit setzen: die Vorlagen sind sich uneinig,
    # wie sie ausgeliefert werden (bei 16x16 und 29x22 ist sie aus, sonst an).
    falz_ids = ", ".join(str(i) for i in eintrag.get("falz", []))
    hide_ids = ", ".join(str(i) for i in eintrag.get("hide_bg", []))
    jpg_zeilen = "\n".join(
        f'    saveJpg(fertig, "{_jp(p)}", {dpi});' for dpi, p in jpg_pfade)

    return f'''#target photoshop
app.preferences.rulerUnits = Units.PIXELS;
(function () {{
  var doc = app.open(new File("{_jp(psd_pfad)}"));
  var HIDE = [{hide_ids}];

  function selectById(id) {{
    var ref = new ActionReference();
    ref.putIdentifier(charIDToTypeID("Lyr "), id);
    var d = new ActionDescriptor();
    d.putReference(charIDToTypeID("null"), ref);
    executeAction(charIDToTypeID("slct"), d, DialogModes.NO);
  }}
  function replaceById(id, file) {{
    selectById(id);
    var d = new ActionDescriptor();
    d.putPath(charIDToTypeID("null"), file);
    executeAction(stringIDToTypeID("placedLayerReplaceContents"), d, DialogModes.NO);
  }}
  function setVis(id, vis) {{ selectById(id); doc.activeLayer.visible = vis; }}
  function weissKorrektur(pct) {{
    var desc = new ActionDescriptor();
    var ref = new ActionReference();
    ref.putClass(stringIDToTypeID("adjustmentLayer"));
    desc.putReference(charIDToTypeID("null"), ref);
    var adj = new ActionDescriptor();
    adj.putEnumerated(charIDToTypeID("Mthd"), charIDToTypeID("CrMthd"), charIDToTypeID("Rltv"));
    var list = new ActionList();
    var c = new ActionDescriptor();
    c.putEnumerated(charIDToTypeID("Clrs"), charIDToTypeID("Clrs"), charIDToTypeID("Whts"));
    c.putUnitDouble(charIDToTypeID("Blck"), charIDToTypeID("#Prc"), pct);
    list.putObject(charIDToTypeID("Clrs"), c);
    adj.putList(charIDToTypeID("Clrs"), list);
    desc.putObject(charIDToTypeID("Usng"), stringIDToTypeID("selectiveColor"), adj);
    executeAction(charIDToTypeID("Mk  "), desc, DialogModes.NO);
  }}
  function saveJpg(d, path, dpi) {{
    d.resizeImage(undefined, undefined, dpi, ResampleMethod.NONE);
    var jo = new JPEGSaveOptions(); jo.quality = 12;
    d.saveAs(new File(path), jo, true, Extension.LOWERCASE);
  }}

  var fertig = null;   // zusammengefasstes Ergebnis-Dokument

  try {{
    // 1) Cover/Rücken einsetzen
{ersetzen}

    // 1b) Falz (Rille am Buchdeckel): nur beim Hardcover sichtbar
    var FALZ = [{falz_ids}];
    for (var i = 0; i < FALZ.length; i++) {{
      try {{ setVis(FALZ[i], {str(bool(hardcover)).lower()}); }} catch (e) {{}}
    }}

    // 2) Weiß-Korrektur (nur bei weißem Umschlag)
    {"weissKorrektur(10);" if weiss else "// keine Weiß-Korrektur"}

    // 3) Hintergrund aus, transparent zuschneiden
    for (var i = 0; i < HIDE.length; i++) {{ try {{ setVis(HIDE[i], false); }} catch (e) {{}} }}
    doc.trim(TrimType.TRANSPARENT);

    // Ergebnis in ein FRISCHES Dokument kopieren, statt die Leinwand einfach
    // wieder zu vergrößern. Grund: trim() verkleinert nur die Leinwand, die
    // Ebenen behalten ihre Pixel außerhalb davon. Ein resizeCanvas() legt die
    // wieder frei — und weil Photoshop Ebenenmasken beim Vergrößern mit WEISS
    // (= sichtbar) fortsetzt, käme die ausmaskierte untere Hälfte der Spiegelung
    // als harter Block zurück. Auf der hellen Vorlage sieht man das nicht, bei
    // einem dunklen Umschlag klebt ein grauer Klotz unter dem Buch.
    // Das neue Dokument enthält nur die sichtbaren Pixel — außerhalb ist nichts.
    // Rand oben/links/rechts, unten aber nur so viel wie gewünscht: das Bild soll
    // dort enden, wo die Spiegelung ausläuft — sonst schwebt das Buch über einer
    // leeren weißen Fläche. Der Zuschnitt oben endet ohnehin genau am Motiv.
    var m = Math.round({rand_cm} / 2.54 * doc.resolution);
    var mu = Math.round({rand_unten_cm} / 2.54 * doc.resolution);
    var W = doc.width.value + 2 * m, H = doc.height.value + m + mu;
    doc.selection.selectAll();
    doc.selection.copy(true);             // auf Basis aller sichtbaren Ebenen
    doc.selection.deselect();

    fertig = app.documents.add(W, H, doc.resolution, "cover_3d",
                               NewDocumentMode.RGB, DocumentFill.TRANSPARENT);
    fertig.paste();                       // wird mittig eingefügt ...
    // ... darum nach oben schieben, bis oben genau m Rand steht.
    var dy = m - (m + mu) / 2;
    if (dy) fertig.activeLayer.translate(0, dy);
    fertig.selection.deselect();

    var o = new ExportOptionsSaveForWeb();
    o.format = SaveDocumentType.PNG; o.PNG8 = false; o.transparency = true;
    fertig.exportDocument(new File("{_jp(png_transp)}"), ExportType.SAVEFORWEB, o);

    // 4) Dieselben Pixel auf Weiß -> JPEG in 300 & 72 dpi (nur DPI-Tag anders).
    // Der Hintergrund der Vorlagen ist reines Weiß.
    var bg = fertig.artLayers.add();
    var weissF = new SolidColor();
    weissF.rgb.red = 255; weissF.rgb.green = 255; weissF.rgb.blue = 255;
    fertig.selection.selectAll();
    fertig.selection.fill(weissF);
    fertig.selection.deselect();
    bg.move(fertig, ElementPlacement.PLACEATEND);
    fertig.flatten();
{jpg_zeilen}
  }} finally {{
    // Immer schließen: sonst bliebe das Master-PSD mit ersetzten Smart-Objekten
    // offen und ungespeichert stehen, wenn oben etwas scheitert.
    if (fertig !== null) {{
      try {{ fertig.close(SaveOptions.DONOTSAVECHANGES); }} catch (e) {{}}
    }}
    doc.close(SaveOptions.DONOTSAVECHANGES);
  }}
}})();
'''


def erzeuge_3d_photoshop(reg: Regionen, bild_hi: Image.Image, cfg: dict,
                         out_dir: Path, sc: str, vorlage_name: str,
                         weiss: bool | None = None, dry_run: bool = False,
                         log=print) -> list[Path]:
    """Erzeugt das 3D-Mockup: Photoshop (per COM) ersetzt im format-passenden
    Mockup-PSD die Cover-/Rücken-Smart-Objekte und exportiert transparentes PNG
    + flache 3D-JPEGs (300 & 72 dpi).

    ``vorlage_name`` ist der Dateiname der Vorlage (Schlüssel in vorlagen_map).
    dry_run=True schreibt nur Slot-PNGs + das .jsx (ohne Photoshop).
    """
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    eintrag = lade_vorlagen_map().get(vorlage_name)
    if eintrag is None:
        raise ValueError(f"Unbekannte Vorlage: {vorlage_name}")
    if not eintrag.get("slots"):
        raise ValueError(f"Vorlage '{vorlage_name}' hat keine Smart-Objekte "
                         "— sie muss manuell in Photoshop befüllt werden.")
    psd = vorlagen_dir(cfg) / vorlage_name
    if not dry_run and not psd.exists():
        raise FileNotFoundError(f"Vorlage nicht gefunden: {psd}")

    dpi_print = int(cfg.get("dpi_print", 300))
    dpi_web = int(cfg.get("dpi_web", 72))
    braucht_spine = eintrag.get("spine") and reg.spine is not None
    slot_pngs = _stage_slots(reg, bild_hi, eintrag, dpi_print, out_dir, sc,
                             braucht_spine, log)
    if weiss is None:
        weiss = ist_weisser_umschlag(extrahiere(
            bild_hi, tuple(round(c * dpi_print / PT_PRO_ZOLL)
                           for c in region_box(reg, "front"))), cfg)

    png_transp = out_dir / cfg.get("muster_3d_png", "{sc}.png").format(sc=sc)
    muster_3d = cfg.get("muster_3d", "3D_{dpi}_{sc}.jpg")
    jpg_pfade = [(dpi_print, out_dir / muster_3d.format(dpi=dpi_print, sc=sc)),
                 (dpi_web, out_dir / muster_3d.format(dpi=dpi_web, sc=sc))]

    hardcover = str(cfg.get("einband", "hardcover")).lower() != "softcover"
    jsx = _baue_jsx(psd, eintrag, slot_pngs, bool(weiss), png_transp, jpg_pfade,
                    float(cfg.get("rand_cm", 1.5)),
                    float(cfg.get("rand_unten_cm", 0.0)), hardcover=hardcover)
    jsx_pfad = out_dir / f"_mockup_{sc}.jsx"
    jsx_pfad.write_text(jsx, encoding="utf-8")

    if dry_run:
        log(f"Dry-Run: Slot-PNG(s) + {jsx_pfad.name} geschrieben "
            f"(Vorlage {vorlage_name}, weiß={bool(weiss)}).")
        return [jsx_pfad] + sorted({p for _, p in slot_pngs})

    log(f"Starte Photoshop (COM) — Vorlage {vorlage_name} …")
    import win32com.client  # nur unter Windows verfügbar; bewusst lokal importiert
    ps = win32com.client.Dispatch("Photoshop.Application")
    ps.DoJavaScript(jsx)
    fehlend = [p for _, p in jpg_pfade if not p.exists()] + \
              ([png_transp] if not png_transp.exists() else [])
    if fehlend:
        # Zwischendateien stehen lassen — mit ihnen lässt sich der Photoshop-
        # Schritt nachvollziehen bzw. das .jsx von Hand ausführen.
        raise RuntimeError("Photoshop hat nicht alle Ausgaben erzeugt: "
                           + ", ".join(p.name for p in fehlend))

    raeume_auf(out_dir, sc)
    return [png_transp, *[p for _, p in jpg_pfade]]


def raeume_auf(out_dir: Path, sc: str) -> list[Path]:
    """Zwischendateien des 3D-Laufs entfernen (_slot_*.png, _mockup_*.jsx).

    Sie sind nur Futter für Photoshop; im Artikelordner haben sie nichts zu
    suchen. Bei einem Fehler bleiben sie liegen (siehe erzeuge_3d_photoshop).
    """
    weg = []
    for p in sorted(Path(out_dir).glob(f"_slot_{sc}_*.png")) + \
            sorted(Path(out_dir).glob(f"_mockup_{sc}.jsx")):
        try:
            p.unlink()
            weg.append(p)
        except OSError:
            pass
    return weg
