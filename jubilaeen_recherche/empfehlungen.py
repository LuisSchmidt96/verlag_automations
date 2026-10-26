"""Empfehlungsliste für ein Zieljahr: alle runden Jubiläen, nach Punkten sortiert, mit Kommentar.

Aufruf: python empfehlungen.py 2029 [2030 …]   → empfehlungen_<jahr>.csv neben dem Skript

Die Punkte sind eine Faustregel, kein Modell: Wer kauft am ehesten ein Buch
beim Verlag? Gemeinden vor Ortsteilen vor Firmen vor Vereinen, große
Jubiläumszahlen vor kleinen, das Verlagsgebiet (Nordbaden, Pfalz) vor dem
Rest. Abzüge für unsichere Jahre und für Fälle, die erfahrungsgemäß nicht
kaufen (Verlage, aufgelöste Firmen). Die Kommentare sagen, warum — damit man
die Reihenfolge prüfen und übergehen kann.
"""

import csv
import re
import sqlite3
import sys
from pathlib import Path

ORDNER = Path(__file__).resolve().parent

KERNGEBIET = {  # Nordbaden und Pfalz: dort sitzt der Verlag, dort kennt man ihn
    "Baden-Baden", "Heidelberg", "Karlsruhe", "Mannheim", "Pforzheim", "Landkreis Calw", "Enzkreis",
    "Landkreis Freudenstadt", "Landkreis Karlsruhe", "Neckar-Odenwald-Kreis", "Landkreis Rastatt",
    "Rhein-Neckar-Kreis", "Landkreis Bad Dürkheim", "Donnersbergkreis", "Landkreis Germersheim",
    "Landkreis Kaiserslautern", "Landkreis Kusel", "Landkreis Südliche Weinstraße", "Rhein-Pfalz-Kreis",
    "Landkreis Südwestpfalz", "Frankenthal (Pfalz)", "Frankenthal", "Kaiserslautern", "Landau in der Pfalz",
    "Ludwigshafen am Rhein", "Neustadt an der Weinstraße", "Pirmasens", "Speyer", "Zweibrücken"}

KLASSE_PUNKTE = {"gemeinde": 50, "ortsteil": 38, "unternehmen": 30, "verein": 12}
KLASSE_NAME = {"gemeinde": "Gemeinde/Stadt", "ortsteil": "Ortsteil", "unternehmen": "Unternehmen/Einrichtung",
               "verein": "Verein"}

# Handnotizen aus der Durchsicht (Stand Oktober 2026, aus dem Gedächtnis —
# vor einer Anfrage prüfen). Schlüssel: Label wie in der DB.
NOTIZEN = {
    "Markgröningen": "Stadt, 1250 Jahre – Favorit",
    "Vaihingen an der Enz": "Stadt, 1250 Jahre – Artikel sagt nur „könnte“, Jahr vorher prüfen",
    "Bad Wimpfen": "Stadt, 1200 Jahre – Favorit",
    "Adelsheim": "Stadt, 1250 Jahre, im Verlagsgebiet – Favorit",
    "Großsachsen": "Urkunde 779 gilt für Groß-, Lützel- und Hohensachsen – evtl. gemeinsames Fest der Gemeinde Hirschberg",
    "Witzenmann-Gruppe": "Familienunternehmen, Pforzheim",
    "Carl Roth": "Familienunternehmen, Karlsruhe",
    "Staatliche Akademie der Bildenden Künste Karlsruhe": "Festschrift so gut wie sicher, Land zahlt",
    "SC Freiburg": "großer Einzelfall – Chronik sicher, Verlag offen",
    "Universität Stuttgart": "groß, plant vermutlich mit eigenen Historikern",
    "Koch, Neff und Volckmar": "2023 insolvent, an Zeitfracht – eher nicht",
    "KNO Verlagsauslieferung": "2023 insolvent, an Zeitfracht – eher nicht",
    "Ipsen (Germany)": "Gründung des französischen Mutterkonzerns – deutsche Tochter bestellt kaum",
    "Erlebnispark Tripsdrill": "Familienbetrieb mit viel Geschichte",
    "Mack Rides GmbH & Co KG": "Familie hinter dem Europa-Park, 250 Jahre – Favorit",
    "Bembé": "gilt als älteste Parkettfabrik",
    "Lamy": "seit 2024 japanischer Eigentümer, Firmengeschichte ist Markenkern",
    "Debeka": "großer Versicherer, Wikidata-Mitarbeiterzahl veraltet",
    "1. FSV Mainz 05": "Chronik sicher, Verlag offen",
    "Karolinen-Gymnasium": "250 Jahre Schule – Festschrift so gut wie gesetzt",
    "Badische Maschinenfabrik Durlach": "direkt im Verlagsgebiet",
}

SCHULE = re.compile(r"schule|gymnasium|realschule|hochschule|akademie|universität|internat|pädagog", re.I)
BANK = re.compile(r"volksbank|raiffeisenbank|sparkasse|bankhaus|\bbank\b", re.I)
VERLAG = re.compile(r"verlag", re.I)
KLEINVEREIN = re.compile(r"narren|zunft|fastnacht|narrhalla|tennis|schach|modellflug|reiter", re.I)


def jubilaeums_punkte(jub):
    if jub >= 1000:
        return 35
    if jub >= 500:
        return 28
    if jub % 100 == 0:
        return 22
    if jub in (125, 150, 175, 250):
        return 15
    if jub == 75:
        return 5
    if jub == 50:
        return 3
    return 0


def einwohner_punkte(ew):
    """Größe zählt: ein 100-Seelen-Dorf finanziert kein Festbuch wie eine Kleinstadt."""
    if ew is None:
        return 0
    if ew < 500:
        return -8
    if ew < 2000:
        return 0
    if ew < 10000:
        return 6
    if ew < 30000:
        return 12
    return 18


def hole_einwohner(qids):
    """Einwohnerzahl (P1082, jüngster Wert) aus Wikidata – über den Cache des Hauptskripts."""
    sys.path.insert(0, str(ORDNER))
    import jubilaeen as j
    ergebnis = {}
    for stapel in j.in_stapeln(qids):
        for b in j.sparql(f"""SELECT ?item ?ew ?stand WHERE {{
  VALUES ?item {{ {j.werte_liste(stapel)} }}
  ?item p:P1082 ?a . ?a ps:P1082 ?ew . OPTIONAL {{ ?a pq:P585 ?stand }}
}}"""):
            q, ew, stand = j.qid(b, "item"), float(j.wert(b, "ew")), j.wert(b, "stand") or ""
            if q not in ergebnis or stand > ergebnis[q][1]:
                ergebnis[q] = (int(ew), stand)
    return {q: ew for q, (ew, _) in ergebnis.items()}


def bewerte(z):
    punkte = KLASSE_PUNKTE[z["klasse"]] + jubilaeums_punkte(z["jubilaeum"]) + einwohner_punkte(z.get("einwohner"))
    notizen = []
    if z["landkreis"] in KERNGEBIET:
        punkte += 15
        notizen.append("Verlagsgebiet")
    elif z["bundesland"] == "Baden-Württemberg":
        punkte += 5
    if z["wikipedia_de"]:
        punkte += 3
    name = z["label"] or ""
    if z["klasse"] == "gemeinde":
        notizen.append("eigener Haushalt – Festbuch über Gemeinde/Heimatverein")
    elif z["klasse"] == "ortsteil":
        notizen.append(f"Ortsteil von {z['gemeinde'] or '?'} – Ansprechpartner Ortschaftsrat/Heimatverein, Geld von der Gemeinde")
    elif z["klasse"] == "unternehmen":
        if SCHULE.search(name):
            punkte += 4
            notizen.append("Schule/Hochschule – Festschrift üblich, kleines Budget")
        elif BANK.search(name):
            punkte += 6
            notizen.append("Bank – klassischer Festschrift-Kunde, vorher Fusion prüfen")
        elif VERLAG.search(name):
            punkte -= 25
            notizen.append("selbst Verlag – macht sein Buch selbst")
        else:
            notizen.append("prüfen, ob es die Firma in dieser Form noch gibt")
    elif z["klasse"] == "verein":
        if KLEINVEREIN.search(name):
            punkte -= 8
            notizen.append("kleiner Verein – eher Broschüre")
        else:
            notizen.append("Verein – Festschrift meist ehrenamtlich, über Anzeigen")
    if z["unsicher"]:
        punkte -= 10
        notizen.append("Jahr unsicher")
    if z["mehrdeutig"]:
        punkte -= 5
        notizen.append("mehrere Jahre im Umlauf")
    if z["quelle"] == "wikidata" and z["klasse"] in ("gemeinde", "ortsteil"):
        notizen.append("Jahr nur aus Wikidata, ohne Beleg")
    if z["quelle"] == "label":
        punkte -= 5
        notizen.append("Jahr nur aus dem Vereinsnamen")
    if name in NOTIZEN:
        notizen.insert(0, NOTIZEN[name])
    return punkte, notizen


def empfehlungen(db, jahr):
    db.row_factory = sqlite3.Row
    zeilen = db.execute("""
        SELECT j.qid, max(j.jubilaeum) AS jubilaeum, i.label, i.klasse, i.gemeinde, i.landkreis, i.bundesland,
               i.website, i.wikipedia_de, i.mehrdeutig, d.jahr AS gruendungsjahr, d.quelle, d.eigenschaft,
               d.wert, coalesce(d.bemerkung, '') AS bemerkung
        FROM jubilaeen j JOIN items i USING (qid) JOIN daten d ON d.id = j.daten_id
        WHERE j.jahr = ? AND j.bevorzugt = 1 AND i.aufgeloest IS NULL
        GROUP BY j.qid""", (jahr,)).fetchall()
    einwohner = hole_einwohner({r["qid"] for r in zeilen if r["klasse"] in ("gemeinde", "ortsteil")})
    ergebnis = []
    for r in zeilen:
        z = dict(r)
        z["einwohner"] = einwohner.get(z["qid"])
        z["unsicher"] = z["bemerkung"].startswith("unsicher")
        punkte, notizen = bewerte(z)
        if z["bemerkung"]:
            notizen.append("Nachtrag: " + z["bemerkung"])
        ergebnis.append((punkte, z, notizen))
    ergebnis.sort(key=lambda t: (-t[0], -t[1]["jubilaeum"], t[1]["label"] or ""))
    return ergebnis


def main():
    db = sqlite3.connect(ORDNER / "jubilaeen.sqlite")
    for jahr in [int(a) for a in sys.argv[1:]] or [2029]:
        pfad = ORDNER / f"empfehlungen_{jahr}.csv"
        liste = empfehlungen(db, jahr)
        # Semikolon + BOM: öffnet so per Doppelklick in einem deutschen Excel
        with open(pfad, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["Rang", "Punkte", "Jubiläum", "Jahre", "seit", "Name", "Art", "Einwohner", "Gemeinde", "Landkreis",
                        "Bundesland", "Quelle", "Beleg", "Kommentar", "Website", "Wikipedia", "QID"])
            for rang, (punkte, z, notizen) in enumerate(liste, 1):
                quelle = {"wikidata": "Wikidata", "wikipedia": "Wikipedia (geprüft)",
                          "label": "Vereinsname"}.get(z["quelle"], z["quelle"])
                beleg = z["wert"] if z["quelle"] == "wikipedia" else ""
                w.writerow([rang, punkte, jahr, z["jubilaeum"], z["gruendungsjahr"], z["label"] or z["qid"],
                            KLASSE_NAME[z["klasse"]], z["einwohner"] or "", z["gemeinde"] or "", z["landkreis"] or "",
                            z["bundesland"] or "", quelle, beleg, "; ".join(notizen),
                            z["website"] or "", z["wikipedia_de"] or "", z["qid"]])
        print(f"{pfad.name}: {len(liste)} Einträge")
        for punkte, z, notizen in liste[:40]:
            print(f"  {punkte:4} {z['jubilaeum']:5}  {(z['label'] or '')[:34]:34} {str(z['einwohner'] or ''):>7} {(z['landkreis'] or '')[:24]:24} {notizen[0][:55]}")


if __name__ == "__main__":
    main()
