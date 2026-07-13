#!/usr/bin/env python3
"""Analysiert die Mockup-Vorlagen und schreibt vorlagen_map.json.

Entwickler-Werkzeug (Windows + Photoshop), kein Teil der .exe. Einmal je
PSD-Revision laufen lassen — Slot-Größen und Anker ändern sich mit den Vorlagen:

    python -m cover_previews.psd_analyse              # alle Vorlagen
    python -m cover_previews.psd_analyse 17x24.psd    # nur eine

Ausgelesen wird je Smart-Objekt-Slot:

* ``size_px``  — Originalgröße des Smart-Objekt-Inhalts. Photoshop passt beim
  Ersetzen den neuen Inhalt NICHT wieder an das ursprüngliche Viereck an, sondern
  behält den Transform und wendet ihn auf die Maße des neuen Inhalts an. Die
  Slot-Größe ist damit der Maßstab, in dem wir liefern müssen.
* ``anchor``   — welche Canvas-Seite des Rücken-Slots am Cover liegt. Ein
  schmalerer Rücken schrumpft sonst zur Außenkante hin und löst sich vom Buch.

Daraus je Vorlage abgeleitet:

* ``book_cm``  — wahre Buchmaße (Breite, Höhe). Die Vorlagennamen sind uneinheitlich:
  ``17x24`` ist Breite×Höhe, ``21x13,5`` dagegen Höhe×Breite. Welche Zuordnung
  stimmt, verrät das Seitenverhältnis des Cover-Smart-Objekts.
* ``content_px_per_cm`` — Auflösung des Inhaltsraums (``SO_cover_breite / Buchbreite``).
  Sie schwankt je Vorlage zwischen ~162 und ~310 dpi; nur darin darf gestaged werden.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

from cover_previews import core

# ExtendScript kennt kein JSON -> Zeilen mit "|" als Trenner zurückgeben.
JSX = r'''
#target photoshop
(function () {
  var doc = app.open(new File("%s"));
  var out = ["DOC|" + doc.resolution];

  function info(id) {
    var ref = new ActionReference();
    ref.putIdentifier(charIDToTypeID("Lyr "), id);
    var d = executeActionGet(ref);
    var name = d.getString(charIDToTypeID("Nm  ")).replace(/\|/g, "/");
    try {
      var som = d.getObjectValue(stringIDToTypeID("smartObjectMore"));
      var size = som.getObjectValue(stringIDToTypeID("size"));
      var w = size.getDouble(stringIDToTypeID("width"));
      var h = size.getDouble(stringIDToTypeID("height"));
      var t = som.getList(stringIDToTypeID("transform"));
      var pts = [];
      for (var i = 0; i < t.count; i++) pts.push(t.getDouble(i));
      return "SLOT|" + id + "|" + name + "|SO|" + w + "|" + h + "|" + pts.join(",");
    } catch (e) {
      return "SLOT|" + id + "|" + name + "|NOSO|0|0|";
    }
  }
  var IDS = [%s];
  for (var i = 0; i < IDS.length; i++) out.push(info(IDS[i]));

  // Falz-Ebenen (Buchdeckel-Rille am Ruecken) — nur beim Hardcover sichtbar.
  // Buch und Spiegelung haben je eine ("Falz", "Falz Kopie").
  function walk(set) {
    for (var i = 0; i < set.layers.length; i++) {
      var L = set.layers[i];
      if (L.typename == "LayerSet") { walk(L); continue; }
      if (/^falz/i.test(L.name)) out.push("FALZ|" + L.id + "|" + L.name);
    }
  }
  walk(doc);

  doc.close(SaveOptions.DONOTSAVECHANGES);
  return out.join("\n");
})();
'''


def _lies_psd(ps, psd: Path,
              layer_ids: list[int]) -> tuple[dict[int, dict], list[int], float]:
    """Smart-Object-Größe + Transform je Slot, Falz-Ebenen-IDs, Dokument-dpi."""
    js = JSX % (str(psd).replace("\\", "/"), ", ".join(str(i) for i in layer_ids))
    slots, falz, dpi = {}, [], 300.0
    for zeile in ps.DoJavaScript(js).strip().splitlines():
        t = zeile.split("|")
        if t[0] == "DOC":
            dpi = float(t[1])
            continue
        if t[0] == "FALZ":
            falz.append(int(t[1]))
            continue
        lid = int(t[1])
        if t[3] != "SO":
            slots[lid] = {"name": t[2], "so": False}
            continue
        werte = [float(x) for x in t[6].split(",")]
        slots[lid] = {
            "name": t[2], "so": True, "w": float(t[4]), "h": float(t[5]),
            # Content-Ecken (0,0) (W,0) (W,H) (0,H) -> Canvas
            "quad": [(werte[i], werte[i + 1]) for i in range(0, 8, 2)],
        }
    return slots, falz, dpi


def _book_cm(format_cm: list[float], so_w: float, so_h: float) -> list[float]:
    """Buchmaße in wahrer Orientierung: die Zuordnung wählen, deren Seitenverhältnis
    zum Cover-Smart-Objekt passt (die Vorlagennamen sind mal B×H, mal H×B)."""
    a, b = format_cm
    return list(min([(a, b), (b, a)], key=lambda c: abs(so_w / so_h - c[0] / c[1])))


def _anchor(spine_quad, cover_quad) -> str:
    """Welche Seite des Rücken-Vierecks liegt am Cover? Diese Kante muss beim
    Ersetzen stehen bleiben, sonst löst sich der Rücken vom Buch."""
    cover_x = [p[0] for p in cover_quad]
    spine_x = [p[0] for p in spine_quad]
    # Rücken links vom Cover -> seine rechte Kante ist die gemeinsame.
    return "right" if sum(spine_x) / 4 < sum(cover_x) / 4 else "left"


def main(argv: list[str]) -> int:
    import win32com.client

    map_pfad = Path(__file__).parent / "vorlagen_map.json"
    mapping = json.loads(map_pfad.read_text(encoding="utf-8"))
    cfg = core.lade_config()
    vd = core.vorlagen_dir(cfg)

    ziel = argv or list(mapping)
    ps = win32com.client.Dispatch("Photoshop.Application")
    print(f"Photoshop {ps.Version} — Vorlagen aus {vd}\n")

    for name in ziel:
        eintrag = mapping.get(name)
        if eintrag is None:
            print(f"{name}: nicht in vorlagen_map.json — übersprungen")
            continue
        slots = eintrag.get("slots", [])
        if not slots:
            print(f"{name:<14} keine Smart-Objekte — übersprungen")
            continue
        psd = vd / name
        if not psd.exists():
            print(f"{name:<14} PSD fehlt: {psd}")
            continue

        daten, falz, dpi = _lies_psd(ps, psd, [s["layer_id"] for s in slots])
        cover = next((daten[s["layer_id"]] for s in slots
                      if s["role"] == "cover" and daten[s["layer_id"]].get("so")), None)
        if cover is None:
            print(f"{name:<14} kein Cover-Smart-Objekt gefunden — übersprungen")
            continue
        eintrag["falz"] = sorted(falz)
        # Photoshop setzt ersetzten Inhalt nach PHYSISCHER Größe ein (px / dpi).
        # Die Staging-PNGs müssen also mit der Auflösung der Vorlage gespeichert
        # werden — 21x13,5 läuft z. B. auf 367,59 dpi statt 300.
        eintrag["content_dpi"] = round(dpi, 2)

        if eintrag.get("format_cm"):
            buch = _book_cm(eintrag["format_cm"], cover["w"], cover["h"])
            k = cover["w"] / buch[0]
            eintrag["book_cm"] = [round(v, 2) for v in buch]
            eintrag["content_px_per_cm"] = round(k, 2)
        else:
            k = None   # EBOOK: kein Buchformat

        for slot in slots:
            L = daten[slot["layer_id"]]
            if not L.get("so"):
                print(f"  ! {name}: Layer {slot['layer_id']} ({L['name']}) "
                      f"ist kein Smart-Objekt")
                slot.pop("size_px", None)
                slot.pop("anchor", None)
                continue
            slot["size_px"] = [int(round(L["w"])), int(round(L["h"]))]
            if slot["role"] == "spine":
                slot["anchor"] = _anchor(L["quad"], cover["quad"])
            else:
                slot.pop("anchor", None)

        kk = f"{k:6.1f} px/cm" if k else "(kein Buchformat)"
        buch_s = f"{eintrag['book_cm'][0]}x{eintrag['book_cm'][1]}" if k else "-"
        print(f"  {name:<14} Buch {buch_s:<10} {kk:<12} "
              f"{eintrag['content_dpi']:>6.2f} dpi   "
              f"{len(slots)} Slots, Falz: {eintrag['falz'] or '—'}")

    map_pfad.write_text(json.dumps(mapping, indent=2, ensure_ascii=False) + "\n",
                        encoding="utf-8")
    print(f"\n-> {map_pfad}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
