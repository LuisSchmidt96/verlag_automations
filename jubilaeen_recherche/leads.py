"""Lead-Listen je Zieljahr: Stufen und Kommentare von Hand, Kontaktdaten aus kontakte/.

Aufruf: python leads.py 2028 2029 2030

Je Jahr braucht es
  empfehlungen_<jahr>.csv   (python empfehlungen.py <jahr>) – Treffer, Belege, Einwohner
  einschaetzung_<jahr>.py   – die Durchsicht von Hand:
      E          QID → (Stufe, Kurzkommentar) für jede Zeile
      KOMMENTAR  QID → ausführlicher Kommentar für S, A, B in Verlags- und Nachbargebiet
      REIHENFOLGE  Liste von QIDs: Reihenfolge innerhalb der Stufen (sonst nach Rang)
Keine Punkteformel: die Stufe ist eine Einzelentscheidung, das Skript sammelt sie nur ein.

Stufen:
  S  Top-Lead: große runde Zahl, Träger mit Geld, gute Chance auf ein Buch
  A  guter Lead: aktiv ansprechen
  B  möglich: mitnehmen, wenn ohnehin Kontakt besteht
  C  unwahrscheinlich: kleines Budget, kleine Zahl oder weit weg vom Programm
  X  kein Lead: aufgelöst, insolvent, selbst Verlag, Gebäude, Datenfehler

Ausgabe je Jahr:
  leads_<jahr>.csv                  alle Treffer mit Stufe
  leads_<jahr>_verlagsgebiet.csv    S, A, B in BW und der Pfalz, mit Kommentar und Kontakt
  leads_<jahr>_nachbargebiete.csv   dasselbe für die Nachbargebiete (Südhessen,
                                    Rheinhessen/Nahe, bayerische Grenzkreise)
"""

import csv
import importlib
import json
import re
import sys
from collections import Counter
from pathlib import Path

ORDNER = Path(__file__).resolve().parent

PFALZ = {"Landkreis Bad Dürkheim", "Donnersbergkreis", "Landkreis Germersheim", "Landkreis Kaiserslautern",
         "Landkreis Kusel", "Landkreis Südliche Weinstraße", "Rhein-Pfalz-Kreis", "Landkreis Südwestpfalz",
         "Frankenthal (Pfalz)", "Frankenthal", "Kaiserslautern", "Landau in der Pfalz", "Ludwigshafen am Rhein",
         "Neustadt an der Weinstraße", "Pirmasens", "Speyer", "Zweibrücken"}
# Nachbargebiete: Ring um BW und Pfalz, nachgetragen wie das Verlagsgebiet.
# Kreisfreie Städte stehen in den Daten oft ohne Landkreis – daher auch über die Gemeinde.
NACHBAR_KREISE = {
    # Südhessen
    "Landkreis Bergstraße", "Landkreis Darmstadt-Dieburg", "Odenwaldkreis", "Landkreis Groß-Gerau",
    # Rheinhessen/Nahe
    "Landkreis Mainz-Bingen", "Landkreis Alzey-Worms", "Landkreis Bad Kreuznach",
    # bayerische Grenzkreise
    "Landkreis Lindau", "Landkreis Oberallgäu", "Landkreis Unterallgäu", "Landkreis Neu-Ulm",
    "Landkreis Günzburg", "Landkreis Dillingen an der Donau", "Landkreis Donau-Ries", "Landkreis Ansbach",
    "Landkreis Neustadt an der Aisch-Bad Windsheim", "Landkreis Würzburg", "Landkreis Main-Spessart",
    "Landkreis Miltenberg",
}
NACHBAR_STAEDTE = {"Darmstadt", "Mainz", "Worms", "Memmingen", "Neu-Ulm"}

STUFEN = {"S": 0, "A": 1, "B": 2, "C": 3, "X": 4}
NARREN = re.compile(r"narr|zunft|hexen|geister|teufel|dämone|fastnacht|wölf", re.I)


def gebiet(z):
    if z["Bundesland"] == "Baden-Württemberg" or z["Landkreis"] in PFALZ:
        return "Verlagsgebiet"
    if z["Landkreis"] in NACHBAR_KREISE or z["Gemeinde"] in NACHBAR_STAEDTE:
        return "Nachbargebiet"
    return "außerhalb"


def lies_kontakte():
    """Kontaktdaten aus kontakte/ergebnis_*.jsonl samt Gegenprobe (kontakte/pruefung.json)."""
    kontakte = {}
    for pfad in sorted((ORDNER / "kontakte").glob("ergebnis_*.jsonl")):
        for zeile in pfad.read_text(encoding="utf-8").splitlines():
            if zeile.strip():
                k = json.loads(zeile)
                kontakte[k["qid"]] = k
    pruef = ORDNER / "kontakte" / "pruefung.json"
    if pruef.exists():
        for qid, status in json.loads(pruef.read_text(encoding="utf-8")).items():
            if qid in kontakte and kontakte[qid].get("email"):  # nach Korrektur geleerte Adressen nicht
                kontakte[qid]["geprueft"] = status
    return kontakte


def domain_passt(email, *websites):
    """Grobe Plausibilität: gehört die Mail-Domain zur Website? Sonst von Hand gegenlesen."""
    if not email or "@" not in email:
        return True
    maildomain = email.rsplit("@", 1)[1].lower()
    for w in websites:
        host = re.sub(r"^https?://(www\.)?", "", (w or "").lower()).split("/")[0]
        if host and (maildomain.endswith(host) or host.endswith(maildomain)
                     or maildomain.split(".")[-2:] == host.split(".")[-2:]):
            return True
    return False


def stufe(z, e):
    if z["QID"] in e.E:
        return e.E[z["QID"]]
    # Nur für Zeilen, die noch niemand durchgesehen hat – leads() meldet sie laut.
    # Vorläufig C, bis die Durchsicht nachgeholt ist.
    if NARREN.search(z["Name"]):
        return "C", "Narrenzunft – Festheft, kein Buch"
    if z["Art"] == "Verein":
        return "C", f"{z['Jahre']} Jahre, kleiner Verein"
    return "C", f"{z['Jahre']} Jahre – zu klein für ein Jubiläumsbuch"


def schreibe_liste(pfad, zeilen, e, kontakte):
    fehlt = [z["Name"] for z in zeilen if z["QID"] not in e.KOMMENTAR]
    if fehlt:
        raise SystemExit(f"{pfad.name}: Kommentar fehlt für {fehlt}")
    with open(pfad, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Nr", "Stufe", "Name", "Art", "Jubiläum", "Ort", "Landkreis", "Bundesland", "Einwohner",
                    "Kommentar", "E-Mail", "E-Mail geprüft", "Telefon", "Anschrift", "Ansprechpartner",
                    "Kontakt-Quelle", "Kontakt-Hinweis", "Beleg", "Wikipedia"])
        for nr, z in enumerate(zeilen, 1):
            k = kontakte.get(z["QID"], {})
            hinweis = k.get("bemerkung") or ""
            if k.get("email") and not domain_passt(k["email"], z["Website"], k.get("quelle")):
                hinweis = "; ".join(x for x in ("Mail-Domain passt nicht zur Website – gegenlesen", hinweis) if x)
            ort = z["Gemeinde"] if z["Gemeinde"] != z["Name"] else ""
            w.writerow([nr, z["Stufe"], z["Name"], z["Art"], f"{z['Jahre']} Jahre (seit {z['seit']})", ort,
                        z["Landkreis"] or z["Gemeinde"], z["Bundesland"], z["Einwohner"], e.KOMMENTAR[z["QID"]],
                        k.get("email") or "", k.get("geprueft") or "", k.get("telefon") or "",
                        k.get("anschrift") or "", k.get("ansprechpartner") or "", k.get("quelle") or "",
                        hinweis, z["Beleg"], z["Wikipedia"]])


def leads(jahr, kontakte):
    e = importlib.import_module(f"einschaetzung_{jahr}")
    zeilen = list(csv.DictReader(open(ORDNER / f"empfehlungen_{jahr}.csv", encoding="utf-8-sig"), delimiter=";"))
    alle = []
    for z in zeilen:
        s, kurz = stufe(z, e)
        alle.append({**z, "Stufe": s, "Einschätzung": kurz, "Gebiet": gebiet(z)})
    offen = [z["Name"] for z in zeilen if z["QID"] not in e.E]
    if offen:
        print(f"!! {jahr}: {len(offen)} Treffer noch nicht von Hand durchgesehen (Regel-Stufe C): {offen[:15]}")
    alle.sort(key=lambda z: (STUFEN[z["Stufe"]], int(z["Rang"])))
    spalten = ["Stufe", "Einschätzung", "Name", "Art", "Jubiläum", "Jahre", "seit", "Einwohner", "Gemeinde",
               "Landkreis", "Bundesland", "Gebiet", "Quelle", "Beleg", "Website", "Wikipedia", "QID"]
    with open(ORDNER / f"leads_{jahr}.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=spalten, delimiter=";", extrasaction="ignore")
        w.writeheader()
        w.writerows(alle)

    reihenfolge = getattr(e, "REIHENFOLGE", [])
    def ordnung(z):
        return STUFEN[z["Stufe"]], reihenfolge.index(z["QID"]) if z["QID"] in reihenfolge else 10_000 + int(z["Rang"])
    for name in ("Verlagsgebiet", "Nachbargebiet"):
        auswahl = sorted((z for z in alle if z["Stufe"] in ("S", "A", "B") and z["Gebiet"] == name), key=ordnung)
        datei = "verlagsgebiet" if name == "Verlagsgebiet" else "nachbargebiete"
        schreibe_liste(ORDNER / f"leads_{jahr}_{datei}.csv", auswahl, e, kontakte)
        print(f"{jahr} {name}: {len(auswahl)} Leads {dict(Counter(z['Stufe'] for z in auswahl))}")
    print(f"{jahr} gesamt:", dict(sorted(Counter(z["Stufe"] for z in alle).items(), key=lambda t: STUFEN[t[0]])))


def main():
    sys.path.insert(0, str(ORDNER))
    kontakte = lies_kontakte()
    for jahr in [int(a) for a in sys.argv[1:]] or [2029]:
        leads(jahr, kontakte)


if __name__ == "__main__":
    main()
