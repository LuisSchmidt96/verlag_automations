#!/usr/bin/env python3
"""
Kopfloser Probelauf — Entwicklungshilfe, nicht Teil der .exe
=============================================================

Fährt den Abgleich ohne GUI und gibt aus, wie sich die Fälle verteilen und
wie die Bewertung im Grenzbereich aussieht. Damit lässt sich die Punktevergabe
am echten Material beurteilen, bevor eine einzige Tkinter-Zeile existiert.

Aufruf aus dem Repo-Wurzelordner:

    python -m mailing_list_updater.probelauf \\
        mailing_list_updater/Access_Export.xlsx \\
        mailing_list_updater/Leware-Kunden_2026_13.08.2026.xlsx \\
        mailing_list_updater/Aufträge_2026.xlsx [weitere Aufträge …]
"""

from __future__ import annotations

import sys
from collections import Counter
from datetime import date

from mailing_list_updater import core


def _kurz(lex: dict) -> str:
    teile = [
        str(lex.get("Kd.-Nr") or ""),
        " ".join(x for x in (str(lex.get("Vorname") or ""),
                             str(lex.get("Name") or "")) if x),
        str(lex.get("Firma") or ""),
        f'{lex.get("Plz") or ""} {lex.get("Ort") or ""}'.strip(),
        f'{lex.get("Straße") or ""} {lex.get("Haus Nr.") or ""}'.strip(),
    ]
    return " | ".join(t for t in teile if t)


def _kurz_access(acc: dict) -> str:
    teile = [
        f'ID {acc.get("ID")}',
        " ".join(x for x in (str(acc.get("Vorname") or ""),
                             str(acc.get("Name") or "")) if x),
        str(acc.get("Institution") or ""),
        f'{acc.get("PLZ") or ""} {acc.get("Ort") or ""}'.strip(),
        f'{acc.get("Straße") or ""} {acc.get("Hausnummer") or ""}'.strip(),
    ]
    return " | ".join(t for t in teile if t)


def main(argv: list[str]) -> int:
    if len(argv) < 4:
        print(__doc__)
        return 2

    access_pfad, kunden_pfad, *auftrag_pfade = argv[1:]

    print("Einlesen …")
    access_spalten, access = core.lade_access(access_pfad)
    _kunden_spalten, kunden = core.lade_lexware_kunden(kunden_pfad)
    auftraege = core.lade_auftraege(auftrag_pfade)
    print(f"  Access:    {len(access):6d} Sätze, {len(access_spalten)} Spalten")
    print(f"  Lexware:   {len(kunden):6d} Kunden")
    print(f"  Aufträge:  {len(auftraege):6d} Zeilen aus {len(auftrag_pfade)} Datei(en)")

    vorher = len(kunden)
    kunden = core.ergaenze_kunden_aus_auftraegen(kunden, auftraege)
    if len(kunden) > vorher:
        print(f"  + {len(kunden) - vorher} Bestandskunden aus den "
              f"Adressspalten der Aufträge nachgebildet")
    else:
        print("  (Aufträge tragen keine brauchbaren Adressspalten — "
              "Bestandskunden fehlen)")

    lage = core.bestelljahre(auftraege)
    print(f"  davon mit Kauf (RG > 0): {len(lage.kaufjahr)} Kunden, "
          f"mit irgendeinem Beleg: {len(lage.hat_beleg)}")
    if lage.fruehestes_datum is not None:
        print(f"  Auftragszeitraum: {str(lage.fruehestes_datum)[:10]} bis "
              f"{str(lage.spaetestes_datum)[:10]}")

    print()
    print("=" * 72)
    print("Plausibilitätsprüfung der Eingaben")
    print("=" * 72)
    for zeile in core.pruefe_spalten(access_spalten, _kunden_spalten,
                                     core.auftrags_spalten(auftraege)):
        print("  Hinweis: " + zeile)
    warnung = core.pruefe_access_vollstaendigkeit(access, date.today().year)
    if warnung:
        print("  !! WARNUNG !!  " + warnung)
    befund = core.pruefe_zeitraeume(kunden, lage)
    print(("  !! WARNUNG !!  " if befund.verdaechtig else "  in Ordnung:  ")
          + befund.text)

    print()
    print("=" * 72)
    print("Abgleich")
    print("=" * 72)
    zuordnungen = core.gleiche_alle_ab(kunden, access, lage)

    zaehler = Counter(z.fall for z in zuordnungen)
    for fall in (core.FALL_NEU, core.FALL_AKTUALISIEREN,
                 core.FALL_UNKLAR, core.FALL_OHNE_AUFTRAG):
        print(f"  {zaehler.get(fall, 0):5d}  {fall}")
    print(f"  {'-' * 5}")
    print(f"  {len(zuordnungen):5d}  gesamt")

    ohne_jahr = sum(1 for z in zuordnungen
                    if z.fall != core.FALL_OHNE_AUFTRAG and z.bestelljahr is None)
    print(f"\n  davon Freiexemplare ohne Bestelljahr: {ohne_jahr}")
    jahre = Counter(z.bestelljahr for z in zuordnungen if z.bestelljahr)
    print("  Bestelljahre:", dict(sorted(jahre.items())))

    print()
    print("=" * 72)
    print("Punkteverteilung des jeweils besten Kandidaten")
    print("=" * 72)
    stufen = [(100, 101), (95, 100), (92, 95), (85, 92), (75, 85),
              (65, 75), (60, 65)]
    for a, b in stufen:
        n = sum(1 for z in zuordnungen
                if z.bester and a <= z.bester.punkte < b)
        marke = ""
        if a == 92:
            marke = "  <- SCHWELLE_SICHER"
        elif a == 75:
            marke = "  <- SCHWELLE_UNKLAR"
        elif a == 60:
            marke = "  <- SCHWELLE_MINDEST"
        print(f"  {a:3d}–{b - 1:3d}: {n:5d}{marke}")
    print(f"  ohne Kandidat: "
          f"{sum(1 for z in zuordnungen if not z.bester):5d}")

    print()
    print("=" * 72)
    print("Stichproben")
    print("=" * 72)
    for fall in (core.FALL_AKTUALISIEREN, core.FALL_UNKLAR,
                 core.FALL_NEU, core.FALL_OHNE_AUFTRAG):
        gruppe = [z for z in zuordnungen if z.fall == fall]
        print(f"\n--- {fall} ({len(gruppe)}) " + "-" * 40)
        for z in gruppe[:6]:
            print(f"  LEX  {_kurz(z.lexware)}")
            if z.bestelljahr:
                print(f"       Bestelljahr {z.bestelljahr}")
            for k in z.kandidaten[:3]:
                sperre = f"  [GESPERRT: {k.sperrgrund}]" if k.gesperrt else ""
                print(f"  ACC  {k.punkte:5.1f}  {_kurz_access(k.access)}"
                      f"   ({k.begruendung}){sperre}")
            for h in z.hinweise:
                print(f"       Hinweis: {h}")
            print()

    print()
    print("=" * 72)
    print("Grenzbereich: die 12 Fälle direkt unter SCHWELLE_SICHER")
    print("=" * 72)
    grenz = sorted(
        (z for z in zuordnungen
         if z.bester and core.SCHWELLE_UNKLAR <= z.bester.punkte < core.SCHWELLE_SICHER),
        key=lambda z: -z.bester.punkte,
    )
    for z in grenz[:12]:
        print(f"  {z.bester.punkte:5.1f}  LEX {_kurz(z.lexware)}")
        print(f"         ACC {_kurz_access(z.bester.access)}"
              f"   ({z.bester.begruendung})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
