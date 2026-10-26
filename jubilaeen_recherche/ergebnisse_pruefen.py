"""Ergebnisse der Prüfpakete kontrollieren und gültige Zeilen in nachtraege.csv übernehmen.

Aufruf:
    python ergebnisse_pruefen.py                 # nur prüfen, Bericht ausgeben
    python ergebnisse_pruefen.py --uebernehmen   # gültige Zeilen anhängen

Erwartet pakete/paket_<nr>.md und pakete/ergebnis_<nr>.jsonl (siehe pruefanleitung.md).
Eine Zeile gilt nur, wenn
  * ihre QID im Paket steht und genau einmal vorkommt,
  * das Jahr wörtlich im Beleg steht (Agenten neigen dazu, aus „175-jähriges
    Bestehen 2012“ ein Gründungsjahr auszurechnen – das ist kein Beleg),
  * der Beleg wörtlich im Paket steht. Geprüft wird gegen das ganze Paket, nicht nur
    den eigenen Block: der Beleg eines Ortsteils steht oft im Block seiner Gemeinde.
    Kürzungen mit „…“, „...“ oder „[...]“ sind erlaubt; jedes Stück muss im Paket stehen.
Bereits in nachtraege.csv vorhandene QIDs werden nicht noch einmal angehängt.
"""

import csv
import json
import re
import sys
from pathlib import Path

ORDNER = Path(__file__).resolve().parent
PAKETE = ORDNER / "pakete"
NACHTRAEGE = ORDNER / "nachtraege.csv"
KUERZUNG = re.compile(r"\s*(?:\[\.\.\.\]|\[…\]|\.\.\.|…)\s*")


def norm(s):
    s = s.replace("“", "\"").replace("„", "\"").replace("”", "\"")
    return re.sub(r"\s+", " ", s).strip()


def bloecke(pfad):
    text = pfad.read_text(encoding="utf-8")
    return {b[3:].split("|")[0].strip(): norm(b) for b in re.split(r"\n(?=## Q)", text)[1:]}


def pruefe(nr):
    paket = bloecke(PAKETE / f"paket_{nr}.md")
    gesamt = " ".join(paket.values())
    pfad = PAKETE / f"ergebnis_{nr}.jsonl"
    if not pfad.exists():
        return None, [f"{nr}: keine Ergebnisdatei"]
    zeilen, fehler, gesehen = [], [], set()
    for i, roh in enumerate(pfad.read_text(encoding="utf-8").splitlines(), 1):
        if not roh.strip():
            continue
        try:
            z = json.loads(roh)
        except ValueError as e:
            fehler.append(f"{nr} Zeile {i}: kein JSON ({e})")
            continue
        q = z.get("qid")
        if q not in paket:
            fehler.append(f"{nr} Zeile {i}: {q} steht nicht im Paket")
            continue
        if q in gesehen:
            fehler.append(f"{nr}: {q} doppelt")
            continue
        gesehen.add(q)
        jahr, beleg = z.get("jahr"), (z.get("beleg") or "").strip()
        if jahr is not None:
            if not isinstance(jahr, int):
                fehler.append(f"{nr}: {q} Jahr ist keine Zahl: {jahr!r}")
                continue
            if str(jahr) not in beleg:
                fehler.append(f"{nr}: {q} ({z.get('label')}) Jahr {jahr} steht nicht im Beleg: {beleg[:80]!r}")
                continue
            stuecke = [s for s in KUERZUNG.split(norm(beleg)) if len(s) > 3]
            if not stuecke or not all(s in gesamt for s in stuecke):
                fehler.append(f"{nr}: {q} ({z.get('label')}) Beleg nicht wörtlich im Paket: {beleg[:80]!r}")
                continue
        zeilen.append(z)
    fehlend = set(paket) - gesehen
    if fehlend:
        fehler.append(f"{nr}: {len(fehlend)} Items ohne Ergebnis")
    return zeilen, fehler


def main():
    uebernehmen = "--uebernehmen" in sys.argv
    vorhanden = {z["qid"] for z in csv.DictReader(open(NACHTRAEGE, encoding="utf-8-sig"), delimiter=";")}
    alle, alle_fehler = [], []
    for p in sorted(PAKETE.glob("paket_*.md")):
        nr = p.stem.split("_", 1)[1]
        if nr == "kontrolle":
            continue
        zeilen, fehler = pruefe(nr)
        if zeilen is None:
            continue  # Paket ohne Ergebnis (noch in Arbeit) – nicht als Fehler melden
        neu = [z for z in zeilen if z["qid"] not in vorhanden]
        if neu:
            print(f"paket {nr}: {len(zeilen)} gültig, {sum(1 for z in zeilen if z.get('jahr') is not None)} mit Jahr, "
                  f"{len(neu)} neu, {len(fehler)} Probleme")
        alle += neu
        alle_fehler += fehler
    print(f"\nNeu: {len(alle)} Zeilen, {sum(1 for z in alle if z.get('jahr') is not None)} mit Jahr")
    if alle_fehler:
        print("Probleme (alle Pakete):\n  " + "\n  ".join(alle_fehler[:80]))
    if uebernehmen and alle:
        with open(NACHTRAEGE, "a", newline="", encoding="utf-8") as f:
            w = csv.writer(f, delimiter=";")
            for z in alle:
                w.writerow([z["qid"], z.get("label", ""), z.get("klasse", ""), z.get("art", ""),
                            "" if z.get("jahr") is None else z["jahr"],
                            "" if z.get("praezision") is None else z["praezision"],
                            int(bool(z.get("unsicher"))), z.get("beleg", ""), z.get("bemerkung", ""),
                            "claude-sonnet"])
        print(f"{len(alle)} Zeilen an nachtraege.csv angehängt")


if __name__ == "__main__":
    main()
