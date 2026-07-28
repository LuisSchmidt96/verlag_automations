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

  // Ebenenbaum: id | parent | ist-Gruppe | top | bottom | name. Daraus leitet
  // Python die Falz-Ebenen (Name ~ "Falz") UND die Spiegelungs-Gruppe her.
  function walk(set, parentId) {
    for (var i = 0; i < set.layers.length; i++) {
      var L = set.layers[i];
      var id = -1; try { id = L.id; } catch (e) {}
      var grp = (L.typename == "LayerSet") ? 1 : 0;
      var b = L.bounds;
      out.push("TREE|" + id + "|" + parentId + "|" + grp + "|" +
               Math.round(b[1].as("px")) + "|" + Math.round(b[3].as("px")) + "|" +
               L.name.replace(/\|/g, "/"));
      if (grp) walk(L, id);
    }
  }
  walk(doc, -1);

  doc.close(SaveOptions.DONOTSAVECHANGES);
  return out.join("\n");
})();
'''


def _lies_psd(ps, psd: Path, layer_ids: list[int]
              ) -> tuple[dict[int, dict], dict[int, dict], float]:
    """Smart-Object-Größe/Transform je Slot, Ebenenbaum, Dokument-dpi."""
    js = JSX % (str(psd).replace("\\", "/"), ", ".join(str(i) for i in layer_ids))
    slots, baum, dpi = {}, {}, 300.0
    for zeile in ps.DoJavaScript(js).strip().splitlines():
        t = zeile.split("|")
        if t[0] == "DOC":
            dpi = float(t[1])
        elif t[0] == "TREE":
            lid = int(t[1])
            baum[lid] = {"parent": int(t[2]), "grp": t[3] == "1",
                         "top": int(t[4]), "bot": int(t[5]), "name": t[6]}
        elif t[0] == "SLOT":
            lid = int(t[1])
            if t[3] != "SO":
                slots[lid] = {"name": t[2], "so": False}
            else:
                werte = [float(x) for x in t[6].split(",")]
                slots[lid] = {
                    "name": t[2], "so": True, "w": float(t[4]), "h": float(t[5]),
                    # Content-Ecken (0,0) (W,0) (W,H) (0,H) -> Canvas
                    "quad": [(werte[i], werte[i + 1]) for i in range(0, 8, 2)],
                }
    return slots, baum, dpi


def _falz_ids(baum: dict[int, dict]) -> list[int]:
    """Falz-Ebenen aus dem Baum (Name beginnt mit „Falz")."""
    return sorted(i for i, n in baum.items()
                  if not n["grp"] and n["name"].lower().startswith("falz"))


def _spiegelung_ids(baum: dict[int, dict], slot_ids: list[int],
                    falz_ids: list[int]) -> list[int]:
    """Die auszublendende Spiegelungs-Gruppe herleiten.

    Jede Vorlage enthält das Buch zweimal: aufrecht und als Spiegelung darunter.
    Buch- und Spiegelungs-Ebenen (Slots + Falz) trennen sich sauber an der
    GRÖSSTEN Lücke ihrer ``top``-Werte — die Spiegelung sitzt deutlich tiefer.
    Zu jeder Spiegelungs-Ebene die äußerste Vorfahr-Gruppe suchen, die KEINE
    Buch-Ebene enthält; das ist die Gruppe, die fürs freigestellte TIF aus muss.
    """
    marker = [i for i in (slot_ids + falz_ids) if i in baum]
    if len(marker) < 2:
        return []                                  # z. B. EBOOK: keine Spiegelung
    tops = sorted((baum[i]["top"], i) for i in marker)
    luecken = [(tops[k + 1][0] - tops[k][0], k) for k in range(len(tops) - 1)]
    breite, k = max(luecken)
    # Ohne echte zweite Hälfte (eine Vorlage ohne Spiegelung) keine Gruppe.
    seite_h = max(n["bot"] for n in baum.values())
    if breite < seite_h * 0.15:
        return []
    buch = {i for _, i in tops[:k + 1]}
    spiegel = {i for _, i in tops[k + 1:]}

    def vorfahren(i):
        kette, p = [], baum[i]["parent"]
        while p in baum:
            kette.append(p)
            p = baum[p]["parent"]
        return kette                               # innen -> außen

    buch_gruppen = {g for i in buch for g in vorfahren(i)}
    gruppen = set()
    for i in spiegel:
        aussen = None
        for g in vorfahren(i):                      # äußerste nicht-Buch-Gruppe
            if g not in buch_gruppen:
                aussen = g
        gruppen.add(aussen if aussen is not None else i)
    return sorted(gruppen)


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

        slot_ids = [s["layer_id"] for s in slots]
        daten, baum, dpi = _lies_psd(ps, psd, slot_ids)
        cover = next((daten[s["layer_id"]] for s in slots
                      if s["role"] == "cover" and daten[s["layer_id"]].get("so")), None)
        if cover is None:
            print(f"{name:<14} kein Cover-Smart-Objekt gefunden — übersprungen")
            continue
        falz = _falz_ids(baum)
        eintrag["falz"] = falz
        # Gruppe, die fürs freigestellte CMYK-TIF ausgeblendet wird (nur Buch).
        eintrag["spiegelung"] = _spiegelung_ids(baum, slot_ids, falz)
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
              f"{eintrag['content_dpi']:>6.2f} dpi   {len(slots)} Slots, "
              f"Falz: {eintrag['falz'] or '—'}, "
              f"Spiegelung: {eintrag['spiegelung'] or '—'}")

    map_pfad.write_text(json.dumps(mapping, indent=2, ensure_ascii=False) + "\n",
                        encoding="utf-8")
    print(f"\n-> {map_pfad}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
