#!/usr/bin/env python3
"""
Kopfloser Probelauf — Entwicklungshilfe, nicht Teil der .exe
=============================================================

Zwei Teile:

1. **Sollwerte**: erfundene Kontakte und Bücher, feste Erwartung. Prüft den
   Vorschlag aus ähnlichen Büchern — schlägt laut fehl, wenn eine Änderung an
   `core.vorschlag` oder an der Startliste ihn verbiegt.
2. **Echter Export** (wenn ein Pfad angegeben ist): legt den Verteiler im
   Scratch-Ordner neu an und zählt, wie viele Startwerte aus Access übernommen
   wurden und was liegen blieb. Die eigene Mappe neben dem Tool bleibt
   unangetastet.

Wie gut der Vorschlag die echten Listen trifft, misst `blindtest.py`.

Aufruf aus dem Repo-Wurzelordner:

    python -m presseverteiler.probelauf [mailing_list_updater/Access_Export.xlsx]
"""

from __future__ import annotations

import os
import sys
import tempfile
import time
from datetime import date
from pathlib import Path

import openpyxl

from presseverteiler import core


def _kontakt(access_id: int, institution: str, mail: str = "") -> core.Kontakt:
    return core.Kontakt(access_id, {"ID": access_id, "Institution": institution,
                                    "eMail": mail, "Presse/ZS": "Presse"})


def _geklickt(schlagworte, *namen) -> set[str]:
    """Was im Regionen-Baum des Buchs angekreuzt ist, wenn man diese
    Schlagworte anklickt: die Oberfläche nimmt die Teilgebiete jeweils mit.
    Bei den Themen gilt einfach, was angeklickt ist."""
    return core.mit_unterbegriffen(set(namen), schlagworte)


def _pruefe(name: str, ist, soll) -> int:
    gleich = ist == soll
    print(f"  {'ok ' if gleich else 'FEHLER'} {name}: {ist}")
    if not gleich:
        print(f"         erwartet {soll}")
    return 0 if gleich else 1


def ausgabe() -> int:
    """Ein Ordner je Buch, die Listen je Versandweg, erneutes Speichern."""
    def person(i, institution, mail, strasse):
        return core.Kontakt(i, {"ID": i, "Institution": institution,
                                "eMail": mail, "Straße": strasse,
                                "Ort": "Ubstadt", "Presse/ZS": "Presse"})

    mail_und_post = person(1, "Zeitung", "a@z.de, b@z.de", "Weg 1")
    doppelte_mail = person(2, "Ärztezeitung", "A@z.de", "Weg 2")
    nur_post = person(3, "Blatt", "", "Weg 3")
    nur_mail = person(4, "Portal", "c@p.de", "")
    weder_noch = person(5, "Niemand", "", "")
    alle = [(k, "Grund") for k in (mail_und_post, doppelte_mail, nur_post,
                                   nur_mail, weder_noch)]

    def spalte(pfad, nummer):
        ws = openpyxl.load_workbook(pfad).worksheets[0]
        return [z[nummer] for z in ws.iter_rows(min_row=2, values_only=True)]

    fehler = 0
    with tempfile.TemporaryDirectory() as tmp:
        ziel = core.speichere_versand("Ein: Titel", "1", {"Baden"}, set(), alle,
                                      [(nur_post, "Grund")], tmp)
        fehler += _pruefe("Ausgabe: Ordner samt vier Dateien", (
            ziel.parent == Path(tmp), sorted(p.name for p in ziel.iterdir())),
            (True, sorted((core.DATEI_EMPFAENGER, core.DATEI_MAIL,
                           core.DATEI_POST, core.DATEI_BCC))))
        fehler += _pruefe("Ausgabe: alle Empfänger, alphabetisch", spalte(
            ziel / core.DATEI_EMPFAENGER, 5),
            ["Ärztezeitung", "Blatt", "Niemand", "Portal", "Zeitung"])
        # Wer dieselbe Adresse teilt, bekommt sie nur einmal; „a@z.de“ steht
        # bei der Ärztezeitung, weil sie alphabetisch vorn liegt.
        fehler += _pruefe("Ausgabe: je Adresse eine E-Mail-Zeile", spalte(
            ziel / core.DATEI_MAIL, 3), ["A@z.de", "c@p.de", "b@z.de"])
        fehler += _pruefe("Ausgabe: BCC-Text wie die Liste", (
            ziel / core.DATEI_BCC).read_text(encoding="utf-8"),
            "A@z.de; c@p.de; b@z.de")
        fehler += _pruefe("Ausgabe: Post mit Anschrift, Vermerk E-Mail", list(zip(
            spalte(ziel / core.DATEI_POST, 5), spalte(ziel / core.DATEI_POST, 14))),
            [("Ärztezeitung", "ja"), ("Blatt", "nein"), ("Zeitung", "ja")])

        # Nun fällt die Post weg: Ihre alte Liste darf nicht stehen bleiben.
        core.speichere_versand("Ein: Titel", "1", {"Baden"}, set(),
                               [(nur_mail, "Grund")], [], tmp)
        fehler += _pruefe("Ausgabe: erneutes Speichern ersetzt alles", (
            len(list(Path(tmp).iterdir())), spalte(ziel / core.DATEI_POST, 5)),
            (1, []))

        # Gleicher Titel an einem anderen Tag, etwa das Jahrbuch des Vorjahrs:
        # Dessen Liste darf nicht überschrieben werden.
        gestern = time.time() - 86400
        os.utime(ziel / core.DATEI_EMPFAENGER, (gestern, gestern))
        zweiter = core.speichere_versand("Ein: Titel", "1", {"Baden"}, set(),
                                         alle, [], tmp)
        fehler += _pruefe("Ausgabe: anderer Tag bekommt eigenen Ordner", (
            zweiter.name, spalte(ziel / core.DATEI_EMPFAENGER, 5)),
            (f"Ein- Titel {date.today().isoformat()}", ["Portal"]))
        erneut = core.speichere_versand("Ein: Titel", "1", {"Baden"}, set(),
                                        alle, [], tmp)
        fehler += _pruefe("Ausgabe: derselbe Tag ersetzt auch dort",
                          erneut == zweiter, True)
        fehler += _pruefe("Ausgabe: ohne Titel zählt die ISBN",
                          core._buch_ordner(Path(tmp), "", "978-3").name, "978-3")
        zweiter.rename(Path(tmp) / "weg")
        os.utime(ziel / core.DATEI_EMPFAENGER)

        (ziel / f"~${core.DATEI_POST}").touch()
        try:
            core.speichere_versand("Ein: Titel", "1", {"Baden"}, set(), alle,
                                   [], tmp)
            gesperrt = False
        except OSError:
            gesperrt = True
        fehler += _pruefe("Ausgabe: offene Datei stoppt vor dem Schreiben", (
            gesperrt, spalte(ziel / core.DATEI_EMPFAENGER, 5)),
            (True, ["Portal"]))
    return fehler


def sollwerte() -> int:
    v = core.Verteiler(schlagworte=core.standard_schlagworte())
    sw = v.schlagworte
    faelle = [
        # (ID, Institution, Regionen laut Einschätzung, immer)
        (1, "Kulturredaktion Karlsruhe", set(), False),
        (2, "Kurier Bretten", {"Kraichgau"}, False),
        (3, "Eisenbahn-Kurier", set(), False),
        (4, "Mannheimer Morgen", {"Mannheim"}, False),
        (5, "dpa Südwest", set(), True),
        (6, "nie gewählt, aus dem Kraichgau", {"Kraichgau"}, False),
        (7, "nie gewählt, sonst woher", set(), False),
        (8, "einmal dabei", set(), False),
    ]
    kontakte = []
    for i, inst, regionen, immer in faelle:
        kontakte.append(_kontakt(i, inst, f"redaktion{i}@example.org"))
        v.eintraege[i] = core.Eintrag(i, regionen, set(), immer)

    heute = date(2026, 6, 1)
    historie = [
        core.Buch(1, date(2026, 1, 10), "Bretten", "", "", {"Kraichgau"},
                  {"Ortsgeschichte"}, {1, 2}),
        core.Buch(2, date(2026, 2, 10), "Östringen", "", "", {"Kraichgau"},
                  {"Ortsgeschichte"}, {1, 2}),
        core.Buch(3, date(2026, 3, 10), "Bahn im Kraichgau", "", "",
                  {"Kraichgau"}, {"Eisenbahn"}, {1, 2, 3}),
        core.Buch(4, date(2026, 3, 11), "Eppingen", "", "", {"Kraichgau"},
                  {"Ortsgeschichte"}, {1, 2}),
        core.Buch(5, date(2026, 3, 12), "Kraichtal", "", "", {"Kraichgau"},
                  {"Ortsgeschichte"}, {1, 2}),
        core.Buch(6, date(2026, 3, 13), "Sulzfeld", "", "", {"Kraichgau"},
                  {"Ortsgeschichte"}, {1, 2}),
        core.Buch(7, date(2026, 3, 14), "Zaisenhausen", "", "", {"Kraichgau"},
                  {"Ortsgeschichte"}, {1, 2}),
        core.Buch(8, date(2026, 3, 15), "Oberöwisheim", "", "", {"Kraichgau"},
                  {"Ortsgeschichte"}, {1, 2, 8}),
        core.Buch(9, date(2026, 4, 10), "Krimi Mannheim", "", "", {"Mannheim"},
                  {"Krimis"}, {1, 4}),
        core.Buch(10, date(2026, 4, 11), "Krimi Neckarau", "", "", {"Mannheim"},
                  {"Krimis"}, {1, 4}),
        core.Buch(11, date(2026, 4, 12), "Krimi Lindenhof", "", "",
                  {"Mannheim"}, {"Krimis"}, {1, 4}),
        # Liegt nach dem Stichtag des Tests unten: darf nie zählen.
        core.Buch(12, date(2026, 7, 1), "Zukunft", "", "", {"Kraichgau"},
                  {"Ortsgeschichte"}, {7}),
    ]

    def frage(regionen, themen, stichtag=heute, **werte):
        vs = core.vorschlag(regionen, themen, historie, v, kontakte, stichtag,
                            **werte)
        an = {t.kontakt.access_id for t in vs.treffer if t.vorgeschlagen}
        return vs, an

    fehler = 0
    vs, an = frage({"Kraichgau"}, {"Ortsgeschichte"})
    fehler += _pruefe("Ortsbuch im Kraichgau: vorangekreuzt", sorted(an),
                      [1, 2, 5])
    nach_id = {t.kontakt.access_id: t for t in vs.treffer}
    fehler += _pruefe("Zukunft zählt nicht", 7 in nach_id, False)
    fehler += _pruefe("selten dabei: unten, nicht angekreuzt",
                      (8 in nach_id, 8 in an), (True, False))
    fehler += _pruefe("Gegend passt: unten, nicht angekreuzt",
                      (6 in nach_id, 6 in an), (True, False))
    reihe = [t.kontakt.access_id for t in vs.treffer]
    fehler += _pruefe("Angekreuzte stehen oben", set(reihe[:3]), {1, 2, 5})
    fehler += _pruefe("Begründung", nach_id[2].warum,
                      f"in {nach_id[2].buecher} von {len(vs.nachbarn)} "
                      f"ähnlichen Büchern")

    _, an = frage({"Mannheim"}, {"Krimis"})
    fehler += _pruefe("Krimi Mannheim: anderes Buch, andere Liste", sorted(an),
                      [1, 4, 5])
    vs, an = frage(set(), set())
    fehler += _pruefe("Nichts angekreuzt: nur „immer“",
                      (sorted(an), vs.nachbarn), ([5], []))
    # Am Erscheinungstag selbst zählt nur, was davor lag.
    vs, _ = frage({"Mannheim"}, {"Krimis"}, stichtag=date(2026, 4, 10))
    fehler += _pruefe("Liste vom selben Tag zählt nicht",
                      9 in {b.nummer for _, b in vs.nachbarn}, False)

    # Mitgenommene Teilgebiete verändern das Buch nicht: Der Landkreis
    # Karlsruhe bringt den Kraichgau mit, gemeint ist aber der Landkreis.
    fehler += _pruefe("Spitzen", core.spitzen(
        _geklickt(sw, "Landkreis Karlsruhe"), sw), {"Landkreis Karlsruhe"})

    neu = list(historie)
    b1 = core.buch_merken(neu, "Neu", "1", {"Kraichgau"}, set(), {1}, heute)
    b2 = core.buch_merken(neu, "Neu", "1", {"Kraichgau"}, set(), {1, 2}, heute)
    fehler += _pruefe("Am selben Tag erneut gespeichert ersetzt",
                      (b1 is b2, len(neu), b2.empfaenger, b2.nummer),
                      (True, len(historie) + 1, {1, 2}, 13))
    with tempfile.TemporaryDirectory() as tmp:
        ziel = core.speichere_historie(neu, kontakte,
                                       Path(tmp) / "Pressehistorie.xlsx")
        wieder = core.lade_historie(ziel)
    felder = ("nummer", "datum", "titel", "isbn", "gattung", "regionen",
              "themen", "empfaenger")
    fehler += _pruefe("Historie hin und zurück", [
        tuple(getattr(b, f) for f in felder) for b in wieder] == [
        tuple(getattr(b, f) for f in felder)
        for b in sorted(neu, key=lambda b: (b.datum, b.nummer))], True)
    fehler += _pruefe("Ohne Mappe leer", core.lade_historie(
        Path(tmp) / "gibt-es-nicht.xlsx"), [])

    fehler += _pruefe("Umlaute sortieren wie Grundbuchstaben", sorted(
        ["Zürich", "Überlingen", "Aalen"], key=core.alphabetisch),
        ["Aalen", "Überlingen", "Zürich"])
    fehler += ausgabe()

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
