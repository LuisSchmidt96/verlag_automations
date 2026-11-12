#!/usr/bin/env python3
"""Baut Pressehistorie.xlsx aus den alten Presselisten (Entwicklerskript).

Die Listen liegen in training_data/ (je Buch ein Ordner, Access-Export mit
Spalte ID, personenbezogen und nicht eingecheckt). Region, Thema und Gattung je
Buch stammen aus einschaetzung/lernen/etikett_*.jsonl. Dort hat ein Modell die
Bücher nach Titel und Shop-Text verschlagwortet, ohne die Empfänger zu sehen.

    python -m presseverteiler.historie_aufbauen [Access-Export] [--ziel Pfad]

Läuft einmal. Danach wächst die Mappe von selbst, weil das Programm jede
gespeicherte Liste aufnimmt. Eine vorhandene Mappe wird nur mit --ersetzen
überschrieben, sonst gingen die seitdem gespeicherten Listen verloren.
"""

import json
import sys
import warnings
from datetime import datetime
from pathlib import Path

import openpyxl

from presseverteiler import core
from presseverteiler.einschaetzung import _export_pfad, _option

ORDNER = Path(__file__).resolve().parent
LISTEN = ORDNER / "training_data" / "Presselisten für Rezensionsversand"
ETIKETTEN = ORDNER / "einschaetzung" / "lernen"

# Über 600 Zeilen ist es keine Auswahl, sondern die ganze Presseliste: Der
# Pressereferent kopiert sie je Buch und löscht, was nicht passt.
GESAMTLISTE_AB = 600


def lies_liste(pfad: Path, access_ids: set[int]) -> dict | None:
    """Empfänger einer Liste als {ID: (Institution, Ort)} samt Datum."""
    wb = openpyxl.load_workbook(pfad, data_only=True)
    zeilen = [list(z) for z in wb.worksheets[0].iter_rows(values_only=True)]
    geaendert = wb.properties.modified
    wb.close()
    zeilen = [z for z in zeilen if any(c not in (None, "") for c in z)]
    if not zeilen or len(zeilen) > GESAMTLISTE_AB:
        return None

    # Die ID-Spalte heißt nicht immer so (acht Listen haben „<“, eine hat gar
    # keine Kopfzeile). Sie ist die Spalte, in der die meisten Access-IDs stehen.
    breite = max(len(z) for z in zeilen)
    spalte, beste = None, 0
    for j in range(breite):
        n = sum(1 for z in zeilen
                if j < len(z) and core._ganzzahl(z[j]) in access_ids)
        if n > beste:
            spalte, beste = j, n
    if spalte is None:
        return None
    kopf = [core._text(c) for c in zeilen[0]]
    hat_kopf = core._ganzzahl(zeilen[0][spalte]) is None

    def wert(z, name):
        return (core._text(z[kopf.index(name)])
                if hat_kopf and name in kopf and kopf.index(name) < len(z)
                else "")

    anzeige = {}
    for z in zeilen[1 if hat_kopf else 0:]:
        access_id = core._ganzzahl(z[spalte]) if spalte < len(z) else None
        if access_id is not None:
            anzeige[access_id] = (wert(z, "Institution"),
                                  " ".join(t for t in (wert(z, "PLZ"),
                                                       wert(z, "Ort")) if t))
    datum = (geaendert or datetime.fromtimestamp(pfad.stat().st_mtime)).date()
    return {"datum": datum, "anzeige": anzeige}


def main(argumente: list[str]):
    warnings.filterwarnings("ignore")   # openpyxl stört sich an alten Formaten
    ziel = Path(_option(argumente, "--ziel", str(core.HISTORIE_PFAD)))
    if ziel.exists() and "--ersetzen" not in argumente:
        print(f"{ziel} gibt es schon. Mit --ersetzen überschreiben (gehen "
              f"seit dem Aufbau gespeicherte Listen verloren).")
        sys.exit(1)
    alle = core.lade_tabelle(_export_pfad(argumente), core.ACCESS_PFLICHT)
    access_ids = {core._ganzzahl(s.get("ID")) for s in alle} - {None}
    kontakte, _ = core.lade_presse(_export_pfad(argumente))

    etiketten = {}
    for pfad in sorted(ETIKETTEN.glob("etikett_*.jsonl")):
        for zeile in pfad.read_text(encoding="utf-8").splitlines():
            if zeile.strip():
                d = json.loads(zeile)
                etiketten[d["pfad"]] = d

    buecher, gesehen, ohne = [], set(), 0
    for pfad in sorted(LISTEN.rglob("*.xlsx")):
        if pfad.name.startswith(("._", "~$")) or pfad.parent == LISTEN:
            continue
        schluessel = str(pfad.relative_to(LISTEN))
        d = etiketten.get(schluessel)
        if d is None:
            ohne += 1
            continue
        liste = lies_liste(pfad, access_ids)
        if liste is None:
            continue
        # Dieselbe Liste zweimal im Ordner desselben Buchs ist eine Kopie.
        doppelt = (pfad.parent.name, frozenset(liste["anzeige"]))
        if doppelt in gesehen:
            continue
        gesehen.add(doppelt)
        buecher.append(core.Buch(
            0, liste["datum"], d.get("titel") or pfad.parent.name,
            d.get("isbn", ""), d.get("gattung", ""), set(d["regionen"]),
            set(d["themen"]), set(liste["anzeige"]), schluessel,
            liste["anzeige"]))
    buecher.sort(key=lambda b: (b.datum, b.herkunft))
    for nummer, b in enumerate(buecher, start=1):
        b.nummer = nummer
    core.speichere_historie(buecher, kontakte, ziel)
    print(f"{len(buecher)} Bücher, {sum(len(b.empfaenger) for b in buecher)} "
          f"Empfänger-Zeilen in {ziel}. Ohne Etikett übergangen: {ohne}.")


if __name__ == "__main__":
    main(sys.argv[1:])
