#!/usr/bin/env python3
"""Blindtest des Vorschlags gegen die echten Presselisten (Entwicklerskript).

Für jedes Buch tut der Test so, als erschiene es gerade. Das Tool bekommt nur
Regionen und Themen des Buchs und die Listen von *davor*. Erst danach
wird der Vorschlag mit der echten Liste verglichen. Geprüft wird core.vorschlag
selbst, also genau der Code, den die Oberfläche benutzt.

    python -m presseverteiler.blindtest --abstimmen     # Werte an Listen bis 2024
    python -m presseverteiler.blindtest                 # blind ab 2025
    python -m presseverteiler.blindtest --ab 2023 --buecher

`--abstimmen` schaut nur auf 2023–2024, damit die Werte in core.py nicht an den
Büchern eingestellt sind, an denen sie danach gemessen werden.
"""

import random
import statistics as st
import sys
from collections import Counter
from dataclasses import replace
from datetime import date
from itertools import product

from presseverteiler import core
from presseverteiler.einschaetzung import _export_pfad, _option


def gefunden_richtig(auswahl: set, echt: set) -> tuple[float, float, float]:
    treffer = len(auswahl & echt)
    gefunden = treffer / len(echt) if echt else 0.0
    richtig = treffer / len(auswahl) if auswahl else 0.0
    f1 = 2 * treffer / (len(auswahl) + len(echt)) if auswahl or echt else 0.0
    return gefunden, richtig, f1


def wie_angekreuzt(b: core.Buch, v: core.Verteiler) -> set:
    """So stehen die Regionen im Baum: samt mitgenommenen Teilgebieten."""
    return core.mit_unterbegriffen(b.regionen, v.schlagworte)


class Lauf:
    def __init__(self, argumente):
        self.kontakte, _ = core.lade_presse(_export_pfad(argumente))
        self.ids = {k.access_id for k in self.kontakte}
        self.v = core.lade_verteiler()
        self.historie = core.lade_historie(
            _option(argumente, "--historie", str(core.HISTORIE_PFAD)))

    def echt(self, b: core.Buch) -> set:
        return b.empfaenger & self.ids

    def fragen(self, b: core.Buch, historie=None, **werte) -> core.Vorschlag:
        """Der Vorschlag für ein Buch, das noch keine Liste hat."""
        return core.vorschlag(wie_angekreuzt(b, self.v), b.themen,
                              self.historie if historie is None else historie,
                              self.v, self.kontakte, stichtag=b.datum, **werte)

    def faelle(self, von: int, bis: int = 9999) -> list[core.Buch]:
        return [b for b in self.historie
                if von <= b.datum.year <= bis and len(self.echt(b)) >= 5]


def blindheit_pruefen(lauf: Lauf, faelle: list[core.Buch], wie_viele=25) -> int:
    """Belegt, dass der Vorschlag nichts vom Buch und der Zukunft weiß.

    Die eigene Liste und alle späteren werden durch Unsinn ersetzt. Bleibt der
    Vorschlag gleich, kann er sie nicht benutzt haben.
    """
    fehler = 0
    zufall = random.Random(7)
    ids = sorted(lauf.ids)
    for b in zufall.sample(faelle, min(wie_viele, len(faelle))):
        verfaelscht = [
            replace(m, empfaenger=set(zufall.sample(ids, 40)))
            if m.datum >= b.datum else m for m in lauf.historie]
        a = [(t.kontakt.access_id, t.stimmen) for t in lauf.fragen(b).treffer]
        c = [(t.kontakt.access_id, t.stimmen)
             for t in lauf.fragen(b, verfaelscht).treffer]
        if a != c:
            fehler += 1
            print(f"  NICHT BLIND: {b.titel[:50]}")
    return fehler


def abstimmen(lauf: Lauf):
    faelle = lauf.faelle(2023, 2024)
    print(f"Abstimmung an {len(faelle)} Büchern aus 2023–2024 (rollend)\n")
    ergebnis = []
    for anzahl, alterung in product((5, 10, 15, 25, 40), (2, 4, 8)):
        vorschlaege = [(b, lauf.fragen(b, anzahl=anzahl, alterung=alterung,
                                       schwelle=0.0)) for b in faelle]
        for schwelle in (.2, .25, .3, .35, .4):
            werte = [gefunden_richtig(
                {t.kontakt.access_id for t in vs.treffer
                 if t.stimmen >= schwelle}, lauf.echt(b))
                for b, vs in vorschlaege]
            ergebnis.append((st.mean(w[2] for w in werte), anzahl, alterung,
                             schwelle, st.mean(w[0] for w in werte),
                             st.mean(w[1] for w in werte)))
    ergebnis.sort(reverse=True)
    print("F1    Nachbarn  Alterung  Schwelle  gefunden  richtig")
    for f1, anzahl, alterung, schwelle, gef, ri in ergebnis[:8]:
        print(f"{f1:.3f}  {anzahl:8}  {alterung:8}  {schwelle:8.2f}  "
              f"{100*gef:7.0f}%  {100*ri:6.0f}%")
    print(f"\nEingestellt in core.py: Nachbarn {core.NACHBARN}, Alterung "
          f"{core.ALTERUNG_JAHRE}, Schwelle {core.SCHWELLE}")


def bericht(lauf: Lauf, ab: int, buecher_zeigen: bool):
    faelle = lauf.faelle(ab)
    print(f"Blindtest: {len(faelle)} Bücher ab {ab}, je Buch nur Listen von "
          f"davor\n")
    fehler = blindheit_pruefen(lauf, faelle)
    print(f"Blindheit: {min(25, len(faelle))} Stichproben mit verfälschter "
          f"Zukunft, {fehler} Abweichungen\n")

    zeilen, oben = [], {n: [] for n in (30, 50, 75, 100, 150, 200)}
    fest, ohne_nachbarn, geschwister = [], 0, []
    for b in faelle:
        echt = lauf.echt(b)
        vs = lauf.fragen(b)
        if not vs.nachbarn:
            ohne_nachbarn += 1
        auswahl = {t.kontakt.access_id for t in vs.treffer if t.vorgeschlagen}
        reihe = [t.kontakt.access_id for t in vs.treffer]
        gef, ri, f1 = gefunden_richtig(auswahl, echt)
        zeilen.append((f1, gef, ri, len(auswahl), len(echt), b, vs))
        for n, w in oben.items():
            w.append(len(set(reihe[:n]) & echt) / len(echt))
        # Vergleich: die 60 bis dahin Häufigsten, ohne jedes Wissen übers Buch
        bisher = Counter(i for m in lauf.historie if m.datum < b.datum
                         for i in m.empfaenger & lauf.ids)
        fest.append(gefunden_richtig({i for i, _ in bisher.most_common(60)},
                                     echt))
        # Listen, die fast gleich früher schon da waren, herausnehmen
        ohne = [m for m in lauf.historie
                if gefunden_richtig(m.empfaenger, b.empfaenger)[2] < .9]
        vo = lauf.fragen(b, ohne)
        geschwister.append(gefunden_richtig(
            {t.kontakt.access_id for t in vo.treffer if t.vorgeschlagen},
            echt)[2])

    def zeile(name, n, gef, ri, f1):
        print(f"  {name:42} {n:5.0f}  {100*gef:7.0f}%  {100*ri:6.0f}%  {f1:.2f}")

    print(f"  {'':42} Empf.  gefunden  richtig  F1")
    zeile("Vorschlag (vorangekreuzt)", st.median(z[3] for z in zeilen),
          st.mean(z[1] for z in zeilen), st.mean(z[2] for z in zeilen),
          st.mean(z[0] for z in zeilen))
    zeile("immer die 60 bisher Häufigsten", 60, st.mean(w[0] for w in fest),
          st.mean(w[1] for w in fest), st.mean(w[2] for w in fest))
    print(f"  {'echte Listen':42} {st.median(z[4] for z in zeilen):5.0f}")
    print(f"\n  ohne fast gleiche frühere Listen: F1 {st.mean(geschwister):.2f}")
    print(f"  Bücher ohne jedes ähnliche Buch: {ohne_nachbarn}")
    verteilung = sorted(z[0] for z in zeilen)
    print(f"  F1 je Buch: schlechtestes Viertel bis "
          f"{verteilung[len(verteilung)//4]:.2f}, Median "
          f"{st.median(verteilung):.2f}, bestes Viertel ab "
          f"{verteilung[3*len(verteilung)//4]:.2f}")

    print("\n  Anteil seiner Liste unter den ersten N der Vorschlagsliste:")
    print("  " + "".join(f"{'Top ' + str(n):>9}" for n in oben))
    print("  " + "".join(f"{100*st.mean(w):8.0f}%" for w in oben.values()))

    print("\n  Je Gattung:")
    gattungen = sorted({z[5].gattung for z in zeilen})
    for g in gattungen:
        m = [z for z in zeilen if z[5].gattung == g]
        print(f"    {g or '—':24} n={len(m):3}  F1 {st.mean(z[0] for z in m):.2f}")
    print("\n  Je Jahr:")
    for j in sorted({z[5].datum.year for z in zeilen}):
        m = [z for z in zeilen if z[5].datum.year == j]
        print(f"    {j}  n={len(m):3}  F1 {st.mean(z[0] for z in m):.2f}")

    print("\n  Wenn er schludrig ankreuzt (F1):")
    for name, aendern in (
            ("ohne Themen", lambda b: replace(b, themen=set())),
            ("ohne Regionen", lambda b: replace(b, regionen=set()))):
        f = []
        for b in faelle:
            vs = lauf.fragen(aendern(b))
            f.append(gefunden_richtig(
                {t.kontakt.access_id for t in vs.treffer if t.vorgeschlagen},
                lauf.echt(b))[2])
        print(f"    {name:18} {st.mean(f):.2f}")

    if buecher_zeigen:
        print("\n  Je Buch (schlechteste und beste):")
        zeilen.sort(key=lambda z: z[0])
        for z in zeilen[:8] + zeilen[-8:]:
            print(f"    {z[0]:.2f} vorgeschlagen {z[3]:3}, echt {z[4]:3}, "
                  f"{z[5].gattung[:14]:14} {z[5].titel[:60]}")


def main(argumente: list[str]):
    lauf = Lauf(argumente)
    print(f"{len(lauf.historie)} Bücher in der Historie\n")
    if "--abstimmen" in argumente:
        abstimmen(lauf)
    else:
        bericht(lauf, int(_option(argumente, "--ab", "2025")),
                "--buecher" in argumente)


if __name__ == "__main__":
    main(sys.argv[1:])
