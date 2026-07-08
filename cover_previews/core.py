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
# Daten-Ordner (Config, Ausgabe) — Muster wie die anderen Tools
# ---------------------------------------------------------------------

def _base_dir() -> Path:
    """Ordner neben der .exe (PyInstaller-Build) bzw. neben dem Tool-Code."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


APP_DIR = _base_dir() / "data"
APP_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_PFAD = APP_DIR / "config.json"


# ---------------------------------------------------------------------
# Konfiguration
# ---------------------------------------------------------------------

DEFAULT_CONFIG = {
    "output_dir": "cover_output",
    "last_input_dir": "",
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
    "kante_tol_pt": 3.0,       # wie nah muss eine Marke am Blattrand liegen
    "cluster_tol_pt": 4.0,     # Marken mit ~gleicher Position zusammenfassen
    # Photoshop-3D
    "mockup_psd_pfad": "",
    "mockup_slots": [{"layer": "COVER", "region": "front_spine"}],
    "dateiname_muster": "{typ}_{dpi}_{sc}.png",
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


def finde_schnittlinien(doc: fitz.Document, cfg: dict | None = None) -> Regionen:
    """Wertet die Schnitt-/Falzmarken aus und leitet die Regionen ab.

    Regel: eine Marke zählt nur, wenn sie den zugehörigen Blattrand berührt
    (vertikale Marke oben/unten, horizontale Marke links/rechts). Dadurch
    fallen innenliegende Element-Marken (z. B. Logo-Rahmen) heraus.
    """
    cfg = cfg or DEFAULT_CONFIG
    tol = float(cfg.get("kante_tol_pt", 3.0))
    ctol = float(cfg.get("cluster_tol_pt", 4.0))

    page = doc[0]
    W, H = page.rect.width, page.rect.height
    segs = _liniensegmente(page)

    x_marks, y_marks = [], []
    for x0, y0, x1, y1 in segs:
        senkrecht = abs(x0 - x1) < 0.5
        waagerecht = abs(y0 - y1) < 0.5
        if senkrecht and not waagerecht:
            # vertikale Marke: zählt, wenn sie oben ODER unten den Rand berührt
            if min(y0, y1) <= tol or max(y0, y1) >= H - tol:
                x_marks.append(x0)
        elif waagerecht and not senkrecht:
            if min(x0, x1) <= tol or max(x0, x1) >= W - tol:
                y_marks.append(y0)

    x_cuts = _cluster(x_marks, ctol)
    y_cuts = _cluster(y_marks, ctol)
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
# 2D-Ausgabe
# ---------------------------------------------------------------------

def _skaliere_lange_kante(bild: Image.Image, max_px: int) -> Image.Image:
    w, h = bild.size
    lang = max(w, h)
    if lang <= max_px:
        return bild
    f = max_px / lang
    return bild.resize((round(w * f), round(h * f)), Image.LANCZOS)


def speichere_2d(front: Image.Image, out_dir: Path, sc: str, cfg: dict) -> list[Path]:
    """Speichert die 2D-Vorderseite in Druck- und Web-Auflösung."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    muster = cfg.get("dateiname_muster", "{typ}_{dpi}_{sc}.png")
    dpi_print = int(cfg.get("dpi_print", 300))
    dpi_web = int(cfg.get("dpi_web", 72))
    web_max = int(cfg.get("web_max_px", 800))

    pfade = []
    p_print = out_dir / muster.format(typ="2D", dpi=dpi_print, sc=sc)
    front.save(p_print, dpi=(dpi_print, dpi_print))
    pfade.append(p_print)

    web = _skaliere_lange_kante(front, web_max)
    p_web = out_dir / muster.format(typ="2D", dpi=dpi_web, sc=sc)
    web.save(p_web, dpi=(dpi_web, dpi_web))
    pfade.append(p_web)
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


def _staging_pngs(reg: Regionen, bild_hi: Image.Image, cfg: dict,
                  out_dir: Path, sc: str) -> list[tuple[str, Path]]:
    """Schneidet je Slot die passende Region aus und legt sie als PNG ab.
    Rückgabe: Liste (layer_name, png_pfad) für das JSX."""
    dpi = int(cfg.get("dpi_print", 300))
    slots = cfg.get("mockup_slots") or []
    ergebnis = []
    for slot in slots:
        layer = slot["layer"]
        box_pt = region_box(reg, slot.get("region", "front"))
        s = dpi / PT_PRO_ZOLL
        box_px = tuple(round(c * s) for c in box_pt)
        img = extrahiere(bild_hi, box_px)
        p = Path(out_dir) / f"_slot_{sc}_{layer}.png"
        img.save(p, dpi=(dpi, dpi))
        ergebnis.append((layer, p))
    return ergebnis


def _baue_jsx(psd_pfad: str, slots: list[tuple[str, Path]], out_png: Path) -> str:
    """ExtendScript, das im Mockup-PSD die benannten Smart-Objekte durch die
    Slot-PNGs ersetzt und das Ergebnis als PNG exportiert. Das Original-PSD
    wird nur geöffnet und ohne Speichern geschlossen (bleibt unverändert)."""
    def jp(p) -> str:                       # Pfad für ExtendScript (Forward-Slash)
        return str(p).replace("\\", "/").replace('"', '\\"')

    ersetzen = "\n".join(
        f'  replaceSO("{layer}", "{jp(png)}");' for layer, png in slots)
    return f'''#target photoshop
(function () {{
  var doc = app.open(new File("{jp(psd_pfad)}"));
  function find(container, name) {{
    for (var i = 0; i < container.layers.length; i++) {{
      var l = container.layers[i];
      if (l.name == name) return l;
      if (l.typename == "LayerSet") {{ var r = find(l, name); if (r) return r; }}
    }}
    return null;
  }}
  function replaceSO(name, file) {{
    var l = find(doc, name);
    if (!l) {{ throw new Error("Smart-Object-Ebene nicht gefunden: " + name); }}
    doc.activeLayer = l;
    var d = new ActionDescriptor();
    d.putPath(charIDToTypeID("null"), new File(file));
    executeAction(stringIDToTypeID("placedLayerReplaceContents"), d, DialogModes.NO);
  }}
{ersetzen}
  var opts = new ExportOptionsSaveForWeb();
  opts.format = SaveDocumentType.PNG; opts.PNG8 = false; opts.quality = 100;
  doc.exportDocument(new File("{jp(out_png)}"), ExportType.SAVEFORWEB, opts);
  doc.close(SaveOptions.DONOTSAVECHANGES);
}})();
'''


def erzeuge_3d_photoshop(reg: Regionen, bild_hi: Image.Image, cfg: dict,
                         out_dir: Path, sc: str, dry_run: bool = False,
                         log=print) -> list[Path]:
    """Erzeugt das 3D-Mockup, indem Photoshop (per COM) im Mockup-PSD die
    Cover-Smart-Objekte austauscht und exportiert.

    dry_run=True schreibt nur die Slot-PNGs + das erzeugte .jsx (zum Prüfen,
    ohne Photoshop). Rückgabe: Liste erzeugter Dateien.
    """
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    psd = cfg.get("mockup_psd_pfad") or ""
    if not psd:
        raise ValueError("Kein Mockup-PSD in der Konfiguration hinterlegt.")

    slots = _staging_pngs(reg, bild_hi, cfg, out_dir, sc)
    muster = cfg.get("dateiname_muster", "{typ}_{dpi}_{sc}.png")
    dpi_print = int(cfg.get("dpi_print", 300))
    out_png = out_dir / muster.format(typ="3D", dpi=dpi_print, sc=sc)
    jsx = _baue_jsx(psd, slots, out_png)
    jsx_pfad = out_dir / f"_mockup_{sc}.jsx"
    jsx_pfad.write_text(jsx, encoding="utf-8")

    if dry_run:
        log(f"Dry-Run: {len(slots)} Slot-PNG(s) + {jsx_pfad.name} geschrieben.")
        return [jsx_pfad, *[p for _, p in slots]]

    log("Starte Photoshop (COM) …")
    import win32com.client  # nur unter Windows verfügbar; bewusst lokal importiert
    ps = win32com.client.Dispatch("Photoshop.Application")
    ps.DoJavaScript(jsx)
    if not out_png.exists():
        raise RuntimeError("Photoshop hat keine Ausgabedatei erzeugt — "
                           "Smart-Object-Ebenenname prüfen.")
    erzeugt = [out_png]

    # Web-Fassung aus dem 3D-Export herunterrechnen.
    dpi_web = int(cfg.get("dpi_web", 72))
    web = _skaliere_lange_kante(Image.open(out_png).convert("RGB"),
                                int(cfg.get("web_max_px", 800)))
    p_web = out_dir / muster.format(typ="3D", dpi=dpi_web, sc=sc)
    web.save(p_web, dpi=(dpi_web, dpi_web))
    erzeugt.append(p_web)
    return erzeugt
