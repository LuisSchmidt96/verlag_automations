"""Erste Einschätzung der Pressekontakte (Entwicklerskript, nicht in der .exe).

Von 1.903 Pressekontakten waren nach dem Umschlüsseln der Access-Felder 980
gar nicht eingeordnet, und die übrigen nur grob. Das alles von Hand
einzuordnen, wäre die Arbeit, die das Programm dem Pressereferenten ersparen
soll. Deshalb schätzen Agenten jeden Kontakt einmal ein: Was ist das für ein
Medium, wo erscheint es, worüber berichtet es? Das Ergebnis kommt als
UNGEPRÜFTER Vorschlag mit Begründung in die Mappe. Bestätigt wird im
Programm (Reiter „Einordnen“).

    python -m presseverteiler.einschaetzung pakete [Access-Export] [--anzahl 10]
    python -m presseverteiler.einschaetzung pruefen
    python -m presseverteiler.einschaetzung uebernehmen [Access-Export] [--mappe Pfad]

`pakete` schreibt nach einschaetzung/ die Schlagwortliste und je Paket eine
Datei paket_<nr>.md. Ein Agent arbeitet ein Paket nach
einschaetzanleitung.md ab und schreibt ergebnis_<nr>.jsonl.
`pruefen` kontrolliert die Ergebnisse, `uebernehmen` trägt sie in die Mappe
ein. Geprüfte Einträge rührt es nicht an.

Pakete und Ergebnisse nennen Redaktionen samt Bemerkungen aus Access, also
personenbezogene Daten. Deshalb liegen sie in einschaetzung/ und sind
nicht eingecheckt.
"""

import json
import shutil
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

from presseverteiler import core

ORDNER = Path(__file__).resolve().parent / "einschaetzung"
ANLEITUNG = Path(__file__).resolve().parent / "einschaetzanleitung.md"
STANDARD_EXPORT = (Path(__file__).resolve().parent.parent
                   / "mailing_list_updater" / "Access_Export.xlsx")

# Was ein Agent je Kontakt sieht. Die Personennamen fehlen absichtlich: Sie
# helfen beim Einordnen nicht, und so steht in den Paketen nicht mehr
# Persönliches als nötig. Die Bemerkungen bleiben drin — in ihnen steht, für
# welches Blatt eine freie Journalistin schreibt oder wer „alle
# Kraichgau-Titel“ möchte.
FELDER = [("Institution", "Institution"), ("Abteilung", "Abteilung"),
          ("Unterabteilung", "Unterabteilung"), ("Land", "Land"),
          ("Medium", "Medium (Access)"), ("Presse/ZS", "Presse/ZS")]


def _export_pfad(argumente: list[str]) -> Path:
    for a in argumente:
        if not a.startswith("--"):
            return Path(a)
    konfig = core.CONFIG_PFAD
    if konfig.exists():
        pfad = json.loads(konfig.read_text(encoding="utf-8")).get("access")
        if pfad:
            return Path(pfad)
    return STANDARD_EXPORT


def _option(argumente: list[str], name: str, standard: str) -> str:
    for i, a in enumerate(argumente):
        if a == name and i + 1 < len(argumente):
            return argumente[i + 1]
        if a.startswith(name + "="):
            return a.split("=", 1)[1]
    return standard


# ---------------------------------------------------------------------
# Pakete schreiben
# ---------------------------------------------------------------------

def schlagwort_text(schlagworte: dict[str, core.Schlagwort]) -> str:
    """Die Schlagwortliste für die Agenten: je Achse ein eingerückter Baum.

    Ein Schlagwort mit mehreren Obergebieten wird nur unter dem ersten voll
    ausgeschrieben, unter den anderen steht ein Verweis. Sonst stünde der
    Rhein-Neckar-Kreis samt Kraichgau, Bergstraße und Odenwald zweimal da,
    und die Liste würde doppelt so lang, ohne mehr zu sagen.
    """
    zeilen = ["# Schlagworte", "",
              "Nur diese Schlagworte sind erlaubt, genau so geschrieben. "
              "Eingerückt = gehört zu dem darüber. Wer für ein Gebiet "
              "eingetragen ist, bekommt auch alles darunter.", ""]
    for achse in core.ACHSEN:
        eigene = {n: s for n, s in schlagworte.items() if s.achse == achse}
        kinder: dict[str, list[str]] = {n: [] for n in eigene}
        wurzeln = []
        for n, s in eigene.items():
            eltern = [p for p in s.gehoert_zu if p in eigene]
            if not eltern:
                wurzeln.append(n)
            for p in eltern:
                kinder[p].append(n)
        zeilen += [f"## {achse}", ""]
        geschrieben: set[str] = set()

        def schreibe(name: str, tiefe: int):
            einzug = "  " * tiefe
            oben = eigene[name].gehoert_zu
            if name in geschrieben:
                zeilen.append(f"{einzug}- {name} (siehe oben)")
                return
            geschrieben.add(name)
            zusatz = (f"  [gehört zu: {'; '.join(oben)}]"
                      if len(oben) > 1 else "")
            zeilen.append(f"{einzug}- {name}{zusatz}")
            for k in kinder[name]:
                schreibe(k, tiefe + 1)

        for w in wurzeln:
            schreibe(w, 0)
        zeilen.append("")
    zeilen += ["## Medienart", "", *(f"- {m}" for m in core.MEDIENARTEN), "",
               "## Ressort", "", *(f"- {r}" for r in core.RESSORTS),
               "- (leer, wenn nicht erkennbar)", ""]
    return "\n".join(zeilen)


def kontakt_zeile(k: core.Kontakt) -> str:
    teile = [f"ID {k.access_id}"]
    for feld, titel in FELDER:
        wert = k.feld(feld)
        if wert:
            teile.append(f"{titel}: {wert}")
    if not k.institution:
        teile.append("(keine Institution — Einzelperson)")
    if k.ort:
        teile.append(f"Ort: {k.ort}")
    domains = sorted({m.split("@", 1)[1].lower() for m in k.mails})
    if domains:
        teile.append(f"Mail-Domain: {', '.join(domains)}")
    for felder, titel in ((core.REGION_FELDER, "Access-Region"),
                          (core.FACH_FELDER, "Access-Fachgebiet")):
        werte = [k.feld(f) for f in felder if k.feld(f)]
        if werte:
            teile.append(f"{titel}: {', '.join(werte)}")
    bemerkung = " ".join(k.feld("Bemerkungen").split())
    if bemerkung:
        teile.append(f"Bemerkung: {bemerkung}")
    return " | ".join(teile)


def pakete(argumente: list[str]):
    export = _export_pfad(argumente)
    anzahl = int(_option(argumente, "--anzahl", "10"))
    kontakte, _ = core.lade_presse(export)
    ORDNER.mkdir(exist_ok=True)
    (ORDNER / "schlagworte.md").write_text(
        schlagwort_text(core.standard_schlagworte()), encoding="utf-8")

    # Kontakte derselben Institution bleiben in einem Paket: Ein Agent, der
    # alle 23 RNZ-Redaktionen sieht, ordnet sie einheitlich ein.
    ziel = -(-len(kontakte) // anzahl)
    gruppen: list[list[core.Kontakt]] = [[]]
    for k in kontakte:
        aktuell = gruppen[-1]
        if (len(aktuell) >= ziel and len(gruppen) < anzahl
                and aktuell[-1].institution.lower() != k.institution.lower()):
            gruppen.append([])
        gruppen[-1].append(k)

    for alt in ORDNER.glob("paket_*.md"):
        alt.unlink()
    for nr, gruppe in enumerate(gruppen, start=1):
        kopf = [f"# Paket {nr:02d} — {len(gruppe)} Kontakte", "",
                f"Ergebnis nach: einschaetzung/ergebnis_{nr:02d}.jsonl "
                f"(eine Zeile je ID, siehe einschaetzanleitung.md)", ""]
        (ORDNER / f"paket_{nr:02d}.md").write_text(
            "\n".join(kopf + [kontakt_zeile(k) for k in gruppe]) + "\n",
            encoding="utf-8")
        print(f"paket_{nr:02d}.md: {len(gruppe)} Kontakte")
    print(f"{len(kontakte)} Kontakte aus {export.name} in {len(gruppen)} "
          f"Paketen, Schlagworte in einschaetzung/schlagworte.md")


# ---------------------------------------------------------------------
# Ergebnisse lesen und prüfen
# ---------------------------------------------------------------------

def paket_ids() -> dict[str, list[int]]:
    ids = {}
    for pfad in sorted(ORDNER.glob("paket_*.md")):
        nr = pfad.stem.split("_", 1)[1]
        ids[nr] = [int(z.split(" | ", 1)[0][3:])
                   for z in pfad.read_text(encoding="utf-8").splitlines()
                   if z.startswith("ID ")]
    return ids


def lade_ergebnisse() -> tuple[dict[int, dict], list[str]]:
    """Alle ergebnis_*.jsonl als {ID: Zeile}, dazu die Fehler beim Lesen."""
    ergebnisse, fehler = {}, []
    for pfad in sorted(ORDNER.glob("ergebnis_*.jsonl")):
        for nr, zeile in enumerate(pfad.read_text(encoding="utf-8")
                                   .splitlines(), start=1):
            if not zeile.strip():
                continue
            try:
                daten = json.loads(zeile)
                access_id = int(daten["id"])
            except (ValueError, KeyError, TypeError) as e:
                fehler.append(f"{pfad.name}:{nr}: unlesbar ({e})")
                continue
            if access_id in ergebnisse:
                fehler.append(f"{pfad.name}:{nr}: ID {access_id} doppelt")
            ergebnisse[access_id] = daten
    return ergebnisse, fehler


def pruefe_zeile(daten: dict, schlagworte: dict[str, core.Schlagwort]
                 ) -> list[str]:
    fehler = []
    for schluessel, achse in (("regionen", core.REGION),
                              ("themen", core.THEMA)):
        werte = daten.get(schluessel, [])
        if not isinstance(werte, list):
            fehler.append(f"„{schluessel}“ ist keine Liste")
            continue
        for w in werte:
            s = schlagworte.get(w)
            if s is None or s.achse != achse:
                fehler.append(f"{achse} „{w}“ gibt es nicht")
    for schluessel, erlaubt in (("medienart", core.MEDIENARTEN),
                                ("ressort", core.RESSORTS)):
        wert = daten.get(schluessel) or ""
        if wert and wert not in erlaubt:
            fehler.append(f"{schluessel} „{wert}“ ist nicht erlaubt")
    if not (daten.get("grund") or "").strip():
        fehler.append("Begründung fehlt")
    return fehler


def pruefen(argumente: list[str] = ()) -> bool:
    schlagworte = core.standard_schlagworte()
    erwartet = paket_ids()
    ergebnisse, fehler = lade_ergebnisse()
    alle = {i for ids in erwartet.values() for i in ids}

    for nr, ids in erwartet.items():
        fehlend = [i for i in ids if i not in ergebnisse]
        datei = ORDNER / f"ergebnis_{nr}.jsonl"
        if not datei.exists():
            print(f"Paket {nr}: noch kein Ergebnis")
        elif fehlend:
            print(f"Paket {nr}: {len(fehlend)} von {len(ids)} fehlen "
                  f"(z. B. {fehlend[:5]})")
        else:
            print(f"Paket {nr}: vollständig ({len(ids)})")
    fremd = sorted(set(ergebnisse) - alle)
    if fremd:
        fehler.append(f"{len(fremd)} IDs stehen in keinem Paket: {fremd[:10]}")
    for access_id, daten in sorted(ergebnisse.items()):
        for f in pruefe_zeile(daten, schlagworte):
            fehler.append(f"ID {access_id}: {f}")

    gueltig = [d for i, d in ergebnisse.items() if i in alle]
    if gueltig:
        regionen = Counter(r for d in gueltig for r in d.get("regionen", []))
        themen = Counter(t for d in gueltig for t in d.get("themen", []))
        leer = sum(1 for d in gueltig if not (d.get("regionen")
                                              or d.get("themen")
                                              or d.get("immer")))
        print(f"\n{len(gueltig)} eingeschätzt, davon "
              f"{sum(1 for d in gueltig if d.get('unsicher'))} unsicher, "
              f"{leer} leer (bekommen nie etwas), "
              f"{sum(1 for d in gueltig if d.get('immer'))} „immer“")
        print("Regionen:", ", ".join(f"{n} {c}" for n, c
                                     in regionen.most_common()))
        print("Themen:", ", ".join(f"{n} {c}" for n, c
                                   in themen.most_common()))
        for schluessel in ("medienart", "ressort"):
            zahl = Counter(d.get(schluessel) or "—" for d in gueltig)
            print(f"{schluessel}:", ", ".join(f"{n} {c}" for n, c
                                              in zahl.most_common()))
    if fehler:
        print(f"\n{len(fehler)} Fehler:")
        for f in fehler:
            print("  " + f)
    return not fehler


# ---------------------------------------------------------------------
# In die Mappe übernehmen
# ---------------------------------------------------------------------

def schlagworte_angleichen(v: core.Verteiler) -> list[str]:
    """Bringt das Blatt „Schlagworte“ auf den Stand der Startliste im Code.

    Nur für diesen einen Schritt: Die Agenten haben nach der Startliste
    eingeordnet, also muss die Mappe jedes ihrer Schlagworte kennen.
    Fehlende kommen dazu, abweichendes „gehört zu“ wird überschrieben —
    und gemeldet, falls das jemand in Excel absichtlich geändert hatte.
    Eigene Schlagworte der Mappe bleiben, hinten angehängt.
    """
    meldungen = []
    neu: dict[str, core.Schlagwort] = {}
    for name, s in core.standard_schlagworte().items():
        alt = v.schlagworte.get(name)
        if alt is None:
            meldungen.append(f"neu: {name} ({s.achse})")
        elif alt.gehoert_zu != s.gehoert_zu:
            meldungen.append(f"geändert: {name} gehört zu "
                             f"„{'; '.join(alt.gehoert_zu)}“ → "
                             f"„{'; '.join(s.gehoert_zu)}“")
        neu[name] = s
    for name, s in v.schlagworte.items():
        if name not in neu:
            neu[name] = s
    v.schlagworte = neu
    return meldungen


def uebernehmen(argumente: list[str]):
    if not pruefen():
        print("\nNichts übernommen — erst die Fehler beheben.")
        sys.exit(1)
    mappe = Path(_option(argumente, "--mappe", str(core.VERTEILER_PFAD)))
    sperre = core.pruefe_sperre(mappe)
    if sperre or (mappe.parent / f".~lock.{mappe.name}#").exists():
        print(sperre or f"{mappe.name} ist in LibreOffice geöffnet.")
        sys.exit(1)

    kontakte, _ = core.lade_presse(_export_pfad(argumente))
    v = core.lade_verteiler(mappe)
    sicherung = ORDNER / (f"{mappe.stem}_vor_einschaetzung_"
                          f"{datetime.now():%Y-%m-%d_%H%M}.xlsx")
    shutil.copy2(mappe, sicherung)
    print(f"\nSicherung: {sicherung}")

    for m in schlagworte_angleichen(v):
        print("Schlagwort " + m)
    neu, _ = core.abgleich(v, kontakte)
    if neu:
        print(f"{len(neu)} neue Pressekontakte aus Access aufgenommen")

    ergebnisse, _ = lade_ergebnisse()
    gesetzt = geprueft = ohne = 0
    for access_id, daten in ergebnisse.items():
        e = v.eintraege.get(access_id)
        if e is None:
            ohne += 1
            continue
        if e.geprueft:
            geprueft += 1
            continue
        e.regionen = set(daten.get("regionen", []))
        e.themen = set(daten.get("themen", []))
        e.immer = bool(daten.get("immer"))
        e.unsicher = bool(daten.get("unsicher"))
        e.medienart = daten.get("medienart") or ""
        e.ressort = daten.get("ressort") or ""
        e.begruendung = " ".join((daten.get("grund") or "").split())
        gesetzt += 1
    core.speichere_verteiler(v, kontakte, mappe)
    print(f"\n{gesetzt} Vorschläge eingetragen (ungeprüft), {geprueft} schon "
          f"geprüfte nicht angerührt, {ohne} ohne Eintrag in der Mappe.")
    meldungen = core.pruefe_verteiler(v)
    for m in meldungen:
        print("Warnung: " + m)


def main(argumente: list[str]):
    befehle = {"pakete": pakete, "pruefen": pruefen,
               "uebernehmen": uebernehmen}
    if not argumente or argumente[0] not in befehle:
        print(__doc__)
        sys.exit(2)
    ergebnis = befehle[argumente[0]](argumente[1:])
    if argumente[0] == "pruefen" and not ergebnis:
        sys.exit(1)


if __name__ == "__main__":
    main(sys.argv[1:])
