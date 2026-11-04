#!/usr/bin/env python3
"""
Kopfloser Probelauf — Entwicklungshilfe, nicht Teil der .exe
=============================================================

Zwei Teile:

1. **Sollwerte**: erfundene Kontakte, drei Testbücher, feste Erwartung. Prüft
   die Auswahlregel samt Obergebieten — schlägt laut fehl, wenn eine Änderung
   an `core.grund` oder an der Startliste die Regel verbiegt.
2. **Echter Export** (wenn ein Pfad angegeben ist): legt den Verteiler im
   Scratch-Ordner neu an und zählt, wie viele Startwerte aus Access übernommen
   wurden und was liegen blieb. Die eigene Mappe neben dem Tool bleibt
   unangetastet.

Aufruf aus dem Repo-Wurzelordner:

    python -m presseverteiler.probelauf [mailing_list_updater/Access_Export.xlsx]
"""

from __future__ import annotations

import sys
import tempfile
from collections import Counter
from pathlib import Path

from presseverteiler import core


def _kontakt(access_id: int, institution: str, mail: str = "") -> core.Kontakt:
    return core.Kontakt(access_id, {"ID": access_id, "Institution": institution,
                                    "eMail": mail, "Presse/ZS": "Presse"})


def _geklickt(schlagworte, *namen) -> set[str]:
    """Was im Regionen-Baum des Buchs angekreuzt ist, wenn man diese
    Schlagworte anklickt: die Oberfläche nimmt die Teilgebiete jeweils mit.
    Bei den Themen gilt einfach, was angeklickt ist."""
    return core.mit_unterbegriffen(set(namen), schlagworte)


def sollwerte() -> int:
    v = core.Verteiler(schlagworte=core.standard_schlagworte())
    sw = v.schlagworte
    faelle = [
        # (ID, Institution, Regionen, Themen, immer)
        (1, "Kurier Bretten", {"Kraichgau"}, set(), False),
        (2, "BNN Karlsruhe", {"Karlsruhe (Region)"}, set(), False),
        (3, "Eisenbahn-Kurier", set(), {"Eisenbahn"}, False),
        (4, "Badische Heimat", {"Baden"}, {"Geschichte"}, False),
        (5, "dpa Südwest", set(), set(), True),
        (6, "nie eingeordnet", set(), set(), False),
        (7, "Mannheimer Morgen", {"Mannheim"}, set(), False),
        (8, "Krimi-Couch", set(), {"Krimis"}, False),
        (9, "Literaturbeilage", set(), {"Belletristik / Literatur"}, False),
        (10, "Die Rheinpfalz", {"Pfalz"}, set(), False),
        (11, "SWR Landesschau", {"Baden-Württemberg"}, set(), False),
        (12, "Archiv für Sozialgeschichte", set(), {"Geschichte"}, False),
        (13, "Heimatblätter Hardt", {"Landkreis Karlsruhe"},
         {"Ortsgeschichte"}, False),
        (14, "Stuttgarter Zeitung", {"Württemberg"}, set(), False),
    ]
    kontakte = []
    for i, inst, regionen, themen, immer in faelle:
        kontakte.append(_kontakt(i, inst, f"redaktion{i}@example.org"))
        v.eintraege[i] = core.Eintrag(i, regionen, themen, immer)

    buecher = [
        # Den Kraichgau selbst angeklickt: der liegt auch im Landkreis
        # Heilbronn, also gehört die Stuttgarter Zeitung (Württemberg) dazu.
        ("Eisenbahn im Kraichgau", _geklickt(sw, "Kraichgau"),
         {"Eisenbahn"}, {1, 2, 3, 5, 11, 14}),
        # Östringen: Ortsgeschichte, im Shop ohne Region — von Hand
        # „Landkreis Karlsruhe“ angeklickt, der Kraichgau kommt als Teilgebiet
        # mit. Also bekommt es der Kurier Bretten, die Stuttgarter Zeitung
        # aber NICHT: der mitgenommene Kraichgau zieht nicht über Heilbronn
        # nach Württemberg (siehe core.erreicht_vom_buch). Badische Heimat und
        # Archiv für Sozialgeschichte auch nicht — Ortsgeschichte hängt
        # absichtlich nicht unter Geschichte.
        ("Ortsgeschichte Östringen", _geklickt(sw, "Landkreis Karlsruhe"),
         {"Ortsgeschichte"}, {1, 2, 5, 11, 13}),
        # Ganz Baden angeklickt: alle badischen Lokalblätter bekommen es.
        ("Kirchengeschichte Baden", _geklickt(sw, "Baden"),
         {"Kirchengeschichte"}, {1, 2, 4, 5, 7, 11, 12}),
        ("Krimi aus Mannheim", _geklickt(sw, "Mannheim"), {"Krimis"},
         {5, 7, 8, 9, 11}),
        # Themen nehmen nichts mit: ein allgemeines Technikbuch geht nicht an
        # das Eisenbahn-Fachblatt.
        ("Technik allgemein", set(), {"Technik und Wirtschaft"}, {5}),
        # IfU hängt unter Technik UND Geschichte — selbst angeklickt, erreicht
        # es beide.
        ("Festschrift IfU", set(), {"IfU"}, {5, 12}),
    ]
    fehler = 0
    for titel, regionen, themen, soll in buecher:
        treffer = core.auswahl(v, kontakte, regionen, themen)
        ist = {k.access_id for k, _ in treffer}
        zeichen = "ok " if ist == soll else "FEHLER"
        print(f"  {zeichen} {titel}: {sorted(ist)}")
        if ist != soll:
            fehler += 1
            print(f"         erwartet {sorted(soll)}")
        for k, g in treffer:
            print(f"         {k.institution:20} {g}")

    meldungen = core.pruefe_verteiler(v)
    if meldungen:
        fehler += 1
        print("  FEHLER Startliste der Schlagworte ist nicht stimmig:")
        for m in meldungen:
            print("         " + m)

    # Doppelt (Groß/klein), mehrere Adressen mit Komma, ein Feld ohne @.
    text, ohne, doppelt = core.mailadressen(
        [_kontakt(1, "a", "x@y.de"), _kontakt(2, "b", "X@y.de"),
         _kontakt(3, "c"), _kontakt(4, "d", "r@z.de, info@z.de"),
         _kontakt(5, "e", "palmyra-verlag")])
    if (text, ohne, doppelt) != ("x@y.de; r@z.de; info@z.de", 2, 1):
        fehler += 1
        print(f"  FEHLER mailadressen: {(text, ohne, doppelt)}")
    return fehler


def echter_export(pfad: str) -> None:
    kontakte, uebergangen = core.lade_presse(pfad)
    print(f"  Pressekontakte: {len(kontakte)} "
          f"(übergangen, weil verzogen/verstorben o. Ä.: {uebergangen})")
    v = core.neuer_verteiler(kontakte)
    eingeordnet = [e for e in v.eintraege.values() if e.eingeordnet]
    print(f"  mit Startwerten aus Access: {len(eingeordnet)}")
    print(f"    nur Region: {sum(1 for e in eingeordnet if e.regionen and not e.themen)}"
          f", nur Thema: {sum(1 for e in eingeordnet if e.themen and not e.regionen)}"
          f", beides: {sum(1 for e in eingeordnet if e.themen and e.regionen)}")
    print(f"  ohne E-Mail: {sum(1 for k in kontakte if not k.mail)}")
    print("  nicht übersetzte Altwerte:")
    for (feld, wert), n in sorted(v.altwerte.items(), key=lambda x: -x[1]):
        print(f"    {n:4}  {feld:12} {wert}")
    for m in core.pruefe_verteiler(v):
        print("  Meldung: " + m)

    # Die Felder der Einschätzung einmal belegen, damit der Test sie mitprüft
    # — sonst fiele eine vergessene Spalte erst beim Pressereferenten auf.
    probe = v.eintraege[kontakte[0].access_id]
    probe.begruendung, probe.unsicher = "Probe: Lokalzeitung", True
    probe.medienart, probe.ressort = core.MEDIENARTEN[0], core.RESSORTS[0]
    felder = ("regionen", "themen", "immer", "geprueft", "bemerkung",
              "begruendung", "unsicher", "medienart", "ressort")
    with tempfile.TemporaryDirectory() as tmp:
        ziel = core.speichere_verteiler(v, kontakte,
                                        Path(tmp) / "Presseverteiler.xlsx")
        wieder = core.lade_verteiler(ziel)
        gleich = all(
            all(getattr(a, f) == getattr(b, f) for f in felder)
            for a, b in ((v.eintraege[i], wieder.eintraege[i])
                         for i in v.eintraege)) and wieder.altwerte == v.altwerte
        print(f"  Mappe hin und zurück: {len(wieder.eintraege)} Einträge, "
              f"{'unverändert' if gleich else 'ABWEICHUNG'}")

        sw = wieder.schlagworte
        for titel, regionen, themen in (
                ("Eisenbahn im Kraichgau", {"Kraichgau"}, {"Eisenbahn"}),
                ("Ortsgeschichte Östringen", {"Landkreis Karlsruhe"},
                 {"Ortsgeschichte"}),
                ("Krimi aus Mannheim", {"Mannheim"}, {"Krimis"})):
            treffer = core.auswahl(wieder, kontakte,
                                   _geklickt(sw, *regionen), themen)
            gruende = Counter(g.split(":")[0] for _, g in treffer)
            print(f"  {titel}: {len(treffer)} Empfänger "
                  f"({', '.join(f'{n}× {g}' for g, n in gruende.most_common())})")


def main(argv: list[str]) -> int:
    print("Sollwerte:")
    fehler = sollwerte()
    if len(argv) > 1:
        print(f"\nEchter Export ({argv[1]}):")
        echter_export(argv[1])
    print(f"\n{'Alles in Ordnung.' if not fehler else f'{fehler} Fehler.'}")
    return 1 if fehler else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
