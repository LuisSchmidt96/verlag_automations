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
    # Dateinamen-Muster (Konvention des Verlags)
    "muster_2d": "2D_{dpi}_{sc}.jpeg",
    "muster_3d": "3D_{dpi}_{sc}.jpg",
    "muster_3d_png": "{sc}.png",
    "jpeg_qualitaet": 95,
    # 3D-Mockup-Vorlagen (ein PSD je Buchformat, im Ordner _NEU_Vorlage)
    "vorlagen_dir": "",            # leer -> mitgelieferter/lokaler _NEU_Vorlage-Ordner
    "vorlagen_tol_cm": 1.0,        # Toleranz beim Format-Abgleich
    "rand_cm": 1.5,               # Rand ums Buch beim transparenten Zuschnitt
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
    """Alle bekannten Vorlagen mit Format + Pfad (inkl. Nicht-Auto wie EBOOK)."""
    vd = vorlagen_dir(cfg)
    liste = []
    for name, t in lade_vorlagen_map().items():
        liste.append({
            "name": name, "pfad": vd / name, "format_cm": t.get("format_cm"),
            "auto": bool(t.get("auto")), "spine": bool(t.get("spine")),
        })
    return liste


def front_masse_cm(reg: "Regionen") -> tuple[float, float] | None:
    if reg.front is None:
        return None
    x0, y0, x1, y1 = reg.front
    return ((x1 - x0) / PT_PRO_ZOLL * 2.54, (y1 - y0) / PT_PRO_ZOLL * 2.54)


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
        if not eintrag["auto"] or not eintrag["format_cm"]:
            continue
        fw, fh = eintrag["format_cm"]
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


def _stage_region(reg: Regionen, bild_hi: Image.Image, region: str, dpi: int,
                  out_dir: Path, sc: str) -> Path:
    """Schneidet eine Region (front/spine) aus und legt sie als PNG ab."""
    s = dpi / PT_PRO_ZOLL
    box_px = tuple(round(c * s) for c in region_box(reg, region))
    img = extrahiere(bild_hi, box_px)
    p = Path(out_dir) / f"_slot_{sc}_{region}.png"
    img.save(p, dpi=(dpi, dpi))
    return p


def _jp(p) -> str:
    """Pfad für ExtendScript (Forward-Slashes, Anführungszeichen escapen)."""
    return str(p).replace("\\", "/").replace('"', '\\"')


def _baue_jsx(psd_pfad, eintrag: dict, cover_png, spine_png, weiss: bool,
              png_transp, jpg_pfade: list[tuple[int, Path]], rand_cm: float) -> str:
    """ExtendScript für das 3D-Mockup:
    - Smart-Objekte per Layer-ID durch Cover/Rücken ersetzen (Reflexion ist mit
      derselben SO-Quelle verknüpft und aktualisiert sich mit);
    - optional Weiß-Korrektur (Selektive Farbkorrektur: Weiß +10 % Schwarz);
    - Hintergrund ausblenden, transparent zuschneiden -> PNG;
    - Hintergrund wieder ein, flach, JPEG in 300 & 72 dpi (nur DPI-Tag).
    Original-PSD wird nur geöffnet und ohne Speichern geschlossen.
    """
    ersetzungen = []
    for slot in eintrag.get("slots", []):
        datei = cover_png if slot["role"] == "cover" else spine_png
        if datei is None:
            continue
        ersetzungen.append(f'  replaceById({slot["layer_id"]}, '
                           f'new File("{_jp(datei)}"));')
    ersetzen = "\n".join(ersetzungen)
    hide_ids = ", ".join(str(i) for i in eintrag.get("hide_bg", []))
    jpg_zeilen = "\n".join(
        f'  saveJpg("{_jp(p)}", {dpi});' for dpi, p in jpg_pfade)

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
  function saveJpg(path, dpi) {{
    doc.resizeImage(undefined, undefined, dpi, ResampleMethod.NONE);
    var jo = new JPEGSaveOptions(); jo.quality = 12;
    doc.saveAs(new File(path), jo, true, Extension.LOWERCASE);
  }}

  // 1) Cover/Rücken einsetzen
{ersetzen}

  // 2) Weiß-Korrektur (nur bei weißem Umschlag)
  {"weissKorrektur(10);" if weiss else "// keine Weiß-Korrektur"}

  // 3) Transparent zuschneiden und als PNG exportieren
  for (var i = 0; i < HIDE.length; i++) {{ try {{ setVis(HIDE[i], false); }} catch (e) {{}} }}
  doc.trim(TrimType.TRANSPARENT);
  var m = Math.round({rand_cm} / 2.54 * doc.resolution);
  doc.resizeCanvas(new UnitValue(doc.width.value + 2 * m, "px"),
                   new UnitValue(doc.height.value + 2 * m, "px"),
                   AnchorPosition.MIDDLECENTER);
  var o = new ExportOptionsSaveForWeb();
  o.format = SaveDocumentType.PNG; o.PNG8 = false; o.transparency = true;
  doc.exportDocument(new File("{_jp(png_transp)}"), ExportType.SAVEFORWEB, o);

  // 4) Hintergrund wieder ein, flach, JPEG in 300 & 72 dpi (gleiche Pixel)
  for (var i = 0; i < HIDE.length; i++) {{ try {{ setVis(HIDE[i], true); }} catch (e) {{}} }}
  doc.flatten();
{jpg_zeilen}
  doc.close(SaveOptions.DONOTSAVECHANGES);
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
    cover_png = _stage_region(reg, bild_hi, "front", dpi_print, out_dir, sc)
    spine_png = (_stage_region(reg, bild_hi, "spine", dpi_print, out_dir, sc)
                 if braucht_spine else None)
    if weiss is None:
        weiss = ist_weisser_umschlag(extrahiere(
            bild_hi, tuple(round(c * dpi_print / PT_PRO_ZOLL)
                           for c in region_box(reg, "front"))), cfg)

    png_transp = out_dir / cfg.get("muster_3d_png", "{sc}.png").format(sc=sc)
    muster_3d = cfg.get("muster_3d", "3D_{dpi}_{sc}.jpg")
    jpg_pfade = [(dpi_print, out_dir / muster_3d.format(dpi=dpi_print, sc=sc)),
                 (dpi_web, out_dir / muster_3d.format(dpi=dpi_web, sc=sc))]

    jsx = _baue_jsx(psd, eintrag, cover_png, spine_png, bool(weiss),
                    png_transp, jpg_pfade, float(cfg.get("rand_cm", 1.5)))
    jsx_pfad = out_dir / f"_mockup_{sc}.jsx"
    jsx_pfad.write_text(jsx, encoding="utf-8")

    if dry_run:
        log(f"Dry-Run: Slot-PNG(s) + {jsx_pfad.name} geschrieben "
            f"(Vorlage {vorlage_name}, weiß={bool(weiss)}).")
        return [jsx_pfad, cover_png] + ([spine_png] if spine_png else [])

    log(f"Starte Photoshop (COM) — Vorlage {vorlage_name} …")
    import win32com.client  # nur unter Windows verfügbar; bewusst lokal importiert
    ps = win32com.client.Dispatch("Photoshop.Application")
    ps.DoJavaScript(jsx)
    fehlend = [p for _, p in jpg_pfade if not p.exists()] + \
              ([png_transp] if not png_transp.exists() else [])
    if fehlend:
        raise RuntimeError("Photoshop hat nicht alle Ausgaben erzeugt: "
                           + ", ".join(p.name for p in fehlend))
    return [png_transp, *[p for _, p in jpg_pfade]]
