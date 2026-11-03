"""Jubiläums-Recherche über Wikidata.

Holt Unternehmen, Vereine, Gemeinden und Ortsteile in Baden-Württemberg,
Rheinland-Pfalz, Saarland und Hessen samt Gründungs- bzw. Ersterwähnungsdatum
vom Wikidata Query Service, schreibt alles in eine SQLite-DB plus CSV und
rechnet die runden Jubiläen aus.

Kein Tool mit GUI wie die anderen Ordner, sondern ein Recherche-Skript für die
Kommandozeile — es läuft selten und liefert Listen, keine Druckdaten.

Aufruf:
    python jubilaeen.py --kreis Q8178            # Probelauf: nur Landkreis Karlsruhe
    python jubilaeen.py                          # alle vier Länder
    python jubilaeen.py --land BW --land HE      # nur einzelne Länder
    python jubilaeen.py --jahr 2029              # zusätzlich: Treffer für 2029
    python jubilaeen.py --nur-auswerten --jahr 2029   # nichts abfragen, nur DB lesen
    python jubilaeen.py --kreis Q8178 --pruefliste    # Wikipedia-Sätze für Items ohne Datum

Nachtragen: die Prüfliste nach pruefanleitung.md abarbeiten, Ergebnisse in
nachtraege.csv; der nächste Lauf (aus dem Cache, also schnell) liest sie ein.

Wiederholungsläufe kosten fast nichts: jede Rohantwort (auch ein Timeout) liegt
in cache/ und wird nicht noch einmal abgefragt. Die DB wird bei jedem Lauf aus
dem Cache neu aufgebaut, damit Schemaänderungen nie eine halbe DB hinterlassen.
"""

import argparse
import csv
import datetime
import hashlib
import json
import re
import sqlite3
import sys
import time
from collections import defaultdict, deque
from pathlib import Path

import requests

ENDPUNKT = "https://query.wikidata.org/sparql"
# Wikimedia sperrt anonyme Standard-User-Agents (python-requests/…) gern ohne
# Vorwarnung; die Policy verlangt einen Namen und eine Kontaktmöglichkeit.
USER_AGENT = "VerlagJubilaeumBot/0.1 (kontakt@cyber-schmidt.de) python-requests"

ORDNER = Path(__file__).resolve().parent
CACHE_ORDNER = ORDNER / "cache"

LAENDER = {
    "BW": ("Q985", "Baden-Württemberg"),
    "RP": ("Q1200", "Rheinland-Pfalz"),
    "SL": ("Q1201", "Saarland"),
    "HE": ("Q1199", "Hessen"),
    "BY": ("Q980", "Bayern"),
}
# Ganz abgefragt werden nur diese vier; von Bayern nur die Kreise an der Grenze
# zu BW (Wunsch der Verlagsleitung, November 2026) – ganz Bayern wäre ein
# Vielfaches an Abfragen und weit weg vom Verlagsgebiet.
STANDARD_LAENDER = ["BW", "RP", "SL", "HE"]
ZUSATZ_KREISE = {
    # an BW grenzend laut Wikidata P47 (die Stadt Neu-Ulm steckt im Landkreis)
    "BY": ["Q14825", "Q10402", "Q10401", "Q10408", "Q10410", "Q10412", "Q10417", "Q10418",
           "Q10454", "Q10456", "Q10458", "Q10469", "Q10472"],
}

# Klassen der Suche. „ort“ heißt: das Item liegt über P159/P131 irgendwo im
# Gebiet. „selbst“ heißt: das Item IST eine Gebietseinheit und hängt selbst
# per P131* am Gebiet.
KLASSEN = {
    "unternehmen": {"oberklassen": ["Q4830453"], "verortung": "ort"},
    "verein": {"oberklassen": ["Q48204", "Q847017"], "verortung": "ort"},
    "gemeinde": {"oberklassen": ["Q262166"], "verortung": "selbst"},
    # Wohnviertel und Stadtbezirk dazu, weil die Stadtteile der Großstädte
    # nicht als Ortsteil geführt sind: Frankfurt-Höchst ist „Stadtteil von
    # Frankfurt am Main“ ⊂ Wohnviertel, Bad Cannstatt „Stadtbezirk in
    # Deutschland“ — beides alte Orte mit eigener Ersterwähnung. Kostet etwas
    # Rauschen (Neubausiedlungen), die bleiben eben ohne Datum.
    "ortsteil": {"oberklassen": ["Q253019", "Q123705", "Q2740635"], "verortung": "selbst"},
}
# Fällt ein Item in mehrere Klassen (ein Sportverein, der zugleich als
# Gewerbebetrieb geführt ist; eine Stadt, die auch Ortsteil-Typ hat), gilt
# die erste Klasse dieser Liste als Hauptklasse. Alle Klassen stehen
# trotzdem in item_klassen.
KLASSEN_VORRANG = ["gemeinde", "ortsteil", "verein", "unternehmen"]
# Ortsjubiläen hängen fast immer an der Ersterwähnung (Lorscher Codex & Co.),
# nicht an einer „Gründung“ — die gibt es bei Dörfern meist gar nicht.
ORTS_KLASSEN = {"gemeinde", "ortsteil"}

KREIS_KLASSEN = ["Q106658", "Q22865"]  # Landkreis, kreisfreie Stadt
GEMEINDE_KLASSE = "Q262166"

# Präzision der Wikidata-Zeitwerte: 9 = Jahr, 10 = Monat, 11 = Tag,
# 8 = Jahrzehnt, 7 = Jahrhundert. Unter Jahresgenauigkeit lässt sich kein
# Jubiläumsjahr ausrechnen.
MIN_PRAEZISION = 9

LABEL_JAHR = re.compile(r"\b(1[6-9]\d\d)\b")

WP_API = "https://de.wikipedia.org/w/api.php"
# Von Hand bzw. von Claude nachgetragene Jahre aus Wikipedia-Artikeln. Liegt
# bewusst außerhalb der DB: die wird bei jedem Lauf neu aus dem Cache gebaut,
# Nachträge darin wären beim nächsten Lauf weg.
NACHTRAEGE = ORDNER / "nachtraege.csv"
NACHTRAEGE_SPALTEN = ["qid", "label", "klasse", "art", "jahr", "praezision", "unsicher",
                      "beleg", "bemerkung", "geprueft_am"]

PAUSE_S = 1.0          # zwischen zwei Anfragen — WDQS ist ein geteilter Dienst
CLIENT_TIMEOUT_S = 90  # WDQS bricht selbst nach 60 s ab; etwas Luft für die Antwort
MAX_VERSUCHE = 5
STAPEL = 150           # QIDs je VALUES-Block


# ---------------------------------------------------------------------------
# Abfragen und Cache
# ---------------------------------------------------------------------------

class Zeitueberschreitung(Exception):
    """WDQS hat die Abfrage wegen der 60-s-Grenze abgebrochen."""


def _cache_pfad(abfrage):
    return CACHE_ORDNER / (hashlib.sha1(abfrage.encode("utf-8")).hexdigest() + ".json")


_sitzung = None
_letzte_anfrage = 0.0
statistik_abfragen = {"cache": 0, "netz": 0, "timeout": 0}


def sparql(abfrage):
    """Führt eine Abfrage aus (oder liest sie aus dem Cache) und liefert die Bindings.

    Auch Timeouts werden gecacht: sonst würde jeder Wiederholungslauf erst
    wieder 60 s auf die Landesabfrage warten, bevor er auf Kreise ausweicht.
    """
    pfad = _cache_pfad(abfrage)
    if pfad.exists():
        eintrag = json.loads(pfad.read_text(encoding="utf-8"))
        statistik_abfragen["cache"] += 1
        if eintrag["status"] == "zeitueberschreitung":
            raise Zeitueberschreitung()
        return eintrag["antwort"]["results"]["bindings"]

    abgerissen = 0
    for versuch in range(1, MAX_VERSUCHE + 1):
        _warte_pause()
        statistik_abfragen["netz"] += 1
        try:
            # POST, weil die VALUES-Blöcke mit 150 QIDs zu lang für eine URL werden
            antwort = _hole_sitzung().post(
                ENDPUNKT, data={"query": abfrage}, timeout=CLIENT_TIMEOUT_S,
                headers={"Accept": "application/sparql-results+json"})
        except requests.Timeout:
            _merke_timeout(pfad, abfrage)
        except requests.exceptions.ChunkedEncodingError:
            # Dritte Form des 60-s-Abbruchs: WDQS kappt den Strom mitten in
            # der Antwort. Kann auch ein Netzaussetzer sein — deshalb erst
            # beim zweiten Mal als Timeout werten (und dann aufteilen).
            abgerissen += 1
            if abgerissen >= 2:
                _merke_timeout(pfad, abfrage)
            _backoff(versuch, "Antwort abgerissen")
            continue
        except requests.ConnectionError as fehler:
            _backoff(versuch, f"Verbindungsfehler: {fehler}")
            continue

        text = antwort.text
        # WDQS meldet den Abbruch auf zwei Arten: 500 mit TimeoutException im
        # Text, oder das Gateway gibt nach ~65 s mit 504 auf.
        if antwort.status_code == 504 or (
                antwort.status_code == 500 and "TimeoutException" in text):
            _merke_timeout(pfad, abfrage)
        if antwort.status_code == 429:
            nachgabe = antwort.headers.get("Retry-After", "")
            sekunden = int(nachgabe) if nachgabe.isdigit() else 30 * versuch
            print(f"    429 – warte {sekunden} s", file=sys.stderr)
            time.sleep(sekunden)
            continue
        if antwort.status_code >= 500:
            _backoff(versuch, f"HTTP {antwort.status_code}")
            continue
        if antwort.status_code != 200:
            raise RuntimeError(f"HTTP {antwort.status_code} bei Abfrage:\n{abfrage}\n\n{text[:2000]}")
        try:
            daten = antwort.json()
        except ValueError:
            # Ergebnisse werden gestreamt; läuft die Abfrage während des
            # Streamens in die 60 s, kommt 200 mit abgeschnittenem JSON.
            _merke_timeout(pfad, abfrage)
        CACHE_ORDNER.mkdir(exist_ok=True)
        pfad.write_text(json.dumps({"status": "ok", "abfrage": abfrage, "antwort": daten},
                                   ensure_ascii=False), encoding="utf-8")
        return daten["results"]["bindings"]
    raise RuntimeError(f"Nach {MAX_VERSUCHE} Versuchen aufgegeben:\n{abfrage}")


def _hole_sitzung():
    global _sitzung
    if _sitzung is None:
        _sitzung = requests.Session()
        _sitzung.headers["User-Agent"] = USER_AGENT
    return _sitzung


def _warte_pause():
    global _letzte_anfrage
    warte = PAUSE_S - (time.monotonic() - _letzte_anfrage)
    if warte > 0:
        time.sleep(warte)
    _letzte_anfrage = time.monotonic()


def _merke_timeout(pfad, abfrage):
    CACHE_ORDNER.mkdir(exist_ok=True)
    pfad.write_text(json.dumps({"status": "zeitueberschreitung", "abfrage": abfrage},
                               ensure_ascii=False), encoding="utf-8")
    statistik_abfragen["timeout"] += 1
    raise Zeitueberschreitung()


def _backoff(versuch, grund):
    sekunden = 5 * 2 ** (versuch - 1)
    print(f"    {grund} – neuer Versuch in {sekunden} s", file=sys.stderr)
    time.sleep(sekunden)


def qid(binding, name):
    wert = binding.get(name, {}).get("value")
    if wert and wert.startswith("http://www.wikidata.org/entity/"):
        return wert.rsplit("/", 1)[1]
    return wert


def wert(binding, name):
    return binding.get(name, {}).get("value")


def werte_liste(qids):
    return " ".join(f"wd:{q}" for q in qids)


def in_stapeln(liste, groesse=STAPEL):
    """Stapel für VALUES-Blöcke.

    QIDs werden in feste Nummernfenster (je 5 Mio.) sortiert und erst darin
    gestapelt. Bei einfach sortierten Stapeln verschob jedes neue Item alle
    Stapel dahinter — und jeder verschobene Stapel ist eine neue Abfrage,
    also ein Cache-Fehlschlag. So trifft ein neues Item nur sein Fenster.
    """
    liste = sorted(liste)
    if not all(re.fullmatch(r"Q\d+", x) for x in liste):  # z. B. Wikipedia-Titel
        for i in range(0, len(liste), groesse):
            yield liste[i:i + groesse]
        return
    fenster = defaultdict(list)
    for q in sorted(liste, key=lambda q: int(q[1:])):
        fenster[int(q[1:]) // 5_000_000].append(q)
    for f in sorted(fenster):
        teil = fenster[f]
        for i in range(0, len(teil), groesse):
            yield teil[i:i + groesse]


# ---------------------------------------------------------------------------
# Gebiete: Kreise und Gemeinden als Teilgebiete für den Split
# ---------------------------------------------------------------------------

def teilgebiete(region, ebene):
    """Teilgebiete für den Split plus die „Zwischenebene“ darüber.

    Land → Kreise, Kreis → Gemeinden. Gesucht wird von oben nach unten, zwei
    Stufen tief: direkt unter dem Gebiet, und unter allem dort, was keine
    Gemeinde ist (Regierungsbezirk, Verbandsgemeinde). Der Weg von unten —
    alle Landkreise Deutschlands, dann P131+ bis zum Land — riss für BW die
    60 s.

    Auf Landesebene zählen auch Gemeinden als Teilgebiet, die direkt am Land
    oder an einem Regierungsbezirk hängen: Darmstadt und Wiesbaden sind in
    Wikidata nur „Stadt in Deutschland“, nicht „kreisfreie Stadt“, und fielen
    sonst samt allen Firmen dort aus dem Split.

    Eine kreisfreie Stadt hat keine Gemeinden unter sich — dort wird nach
    Stadtteilen/Stadtbezirken geteilt. Sonst fielen alle Firmen weg, deren
    Sitz ein Stadtteil ist (Porsche: Zuffenhausen); die Unternehmensabfrage
    für ganz Stuttgart reißt die 60 s.

    Die Zwischenebene (das Gebiet selbst, Regierungsbezirke, Verbands-
    gemeinden) wird eigens abgefragt — dort hängen Items, die in keinem
    Teilgebiet auftauchen.
    """
    abfrage = f"""SELECT DISTINCT ?einheit ?eltern ?istKreis ?istGemeinde ?istOrtsteil WHERE {{
  hint:Query hint:optimizer "None" .
  {{ ?einheit wdt:P131 wd:{region} . BIND(wd:{region} AS ?eltern) }}
  UNION
  {{ ?eltern wdt:P131 wd:{region} .
     FILTER NOT EXISTS {{ ?eltern wdt:P31/wdt:P279* wd:{GEMEINDE_KLASSE} }}
     ?einheit wdt:P131 ?eltern . }}
  FILTER NOT EXISTS {{ ?einheit wdt:P576 [] }}
  # OPTIONAL statt BIND(EXISTS …): mit optimizer "None" bleibt EXISTS bei
  # WDQS ungebunden, und der FILTER wirft dann jede Zeile weg.
  OPTIONAL {{ VALUES ?kk {{ {werte_liste(KREIS_KLASSEN)} }}
             ?einheit wdt:P31/wdt:P279* ?kk . BIND("true" AS ?istKreis) }}
  OPTIONAL {{ ?einheit wdt:P31/wdt:P279* wd:{GEMEINDE_KLASSE} . BIND("true" AS ?istGemeinde) }}
  OPTIONAL {{ VALUES ?ok {{ {werte_liste(KLASSEN["ortsteil"]["oberklassen"])} }}
             ?einheit wdt:P31/wdt:P279* ?ok . BIND("true" AS ?istOrtsteil) }}
  FILTER(BOUND(?istKreis) || BOUND(?istGemeinde) || BOUND(?istOrtsteil))
}}"""
    zeilen = [(qid(b, "einheit"), qid(b, "eltern"), wert(b, "istKreis") == "true",
               wert(b, "istGemeinde") == "true", wert(b, "istOrtsteil") == "true")
              for b in sparql(abfrage)]
    if ebene == "land":
        kreise = {e for e, _, k, _, _ in zeilen if k}
        # Gemeinden unter einem Kreis (Saarland, RLP: Kreis hängt direkt am
        # Land, seine Gemeinden sind dann Stufe zwei) sind kein Teilgebiet.
        teile = kreise | {e for e, el, k, g, _ in zeilen if g and not k and el not in kreise}
    else:
        # Gemeinden und Stadtteile/-bezirke immer zusammen: unter Stuttgart
        # hängt per Fehleintrag eine einzelne Gemeinde, und „erst Gemeinden,
        # nur wenn keine da sind Stadtteile“ ließ dann alle Bezirke liegen.
        # Bei Landkreisen kommen so nur die wenigen Ortsteile dazu, die direkt
        # am Kreis hängen. Ein Stadtteil unter einem Stadtbezirk steckt schon
        # in dessen P131*-Abfrage.
        ortsteile = {e for e, _, _, _, o in zeilen if o}
        teile = ({e for e, _, _, g, _ in zeilen if g}
                 | {e for e, el, _, _, o in zeilen if o and el not in ortsteile})
    zwischen = ({region} | {el for e, el, *_ in zeilen if e in teile}) - teile
    return teile, zwischen


def abfrage_items(klasse, region=None, orte=None, typ_zuerst=False, endziffer=None):
    """SPARQL für eine Klasse in einem Gebiet (oder an genau diesen Orten).

    ``endziffer`` nimmt nur Items, deren QID auf diese Ziffer endet — für
    Orte, an denen so viel direkt hängt (Frankfurt am Main), dass schon die
    Typprüfung für alle zusammen die 60 s reißt.
    """
    ober = werte_liste(KLASSEN[klasse]["oberklassen"])
    typ_pruefung = f"VALUES ?oberklasse {{ {ober} }}\n  ?item wdt:P31 ?typ . ?typ wdt:P279* ?oberklasse ."
    if KLASSEN[klasse]["verortung"] == "selbst":
        if orte is not None:
            gebiet = f"VALUES ?item {{ {werte_liste(orte)} }}"
        else:
            gebiet = f"?item wdt:P131* wd:{region} ."
        if typ_zuerst:
            # Landesebene: Gemeinden gibt es deutschlandweit nur ~11 000, das
            # ist billiger als alles aufzuzählen, was in einem Land liegt.
            koerper = f"""VALUES ?oberklasse {{ {ober} }}
  ?typ wdt:P279* ?oberklasse .
  ?item wdt:P31 ?typ .
  {gebiet}"""
        else:
            koerper = f"{gebiet}\n  {typ_pruefung}"
        return f"""SELECT DISTINCT ?item WHERE {{
  hint:Query hint:optimizer "None" .
  {koerper}
}}"""
    # Verortung über einen Ort: Hauptsitz (P159) oder Lage (P131). Erst den
    # Ort einschränken, dann die Typprüfung — andersherum müsste WDQS alle
    # Gewerbebetriebe der Welt anfassen, und P279* unter „business“ ist riesig.
    if orte is not None:
        ortsbedingung = f"VALUES ?ort {{ {werte_liste(orte)} }}"
    else:
        ortsbedingung = f"?ort wdt:P131* wd:{region} ."
    # Nur einfügen, wenn gebraucht: jedes Zeichen mehr im Abfragetext ist ein
    # anderer Cache-Schlüssel — eine leere Zeile hat schon einmal den ganzen
    # Cache entwertet.
    endfilter = f'\n  FILTER(STRENDS(STR(?item), "{endziffer}"))' if endziffer is not None else ""
    return f"""SELECT DISTINCT ?item ?ort ?ueber WHERE {{
  hint:Query hint:optimizer "None" .
  {ortsbedingung}
  {{ ?item wdt:P159 ?ort . BIND("P159" AS ?ueber) }}
  UNION
  {{ ?item wdt:P131 ?ort . BIND("P131" AS ?ueber) }}{endfilter}
  {typ_pruefung}
}}"""


def treffer_lesen(bindings):
    ergebnis = defaultdict(set)  # qid -> {(ort, über)}
    for b in bindings:
        item = qid(b, "item")
        if b.get("ort"):
            ergebnis[item].add((qid(b, "ort"), wert(b, "ueber")))
        else:
            ergebnis[item].add((item, "selbst"))
    return ergebnis


class Sammler:
    """Hält die Treffer aller Abfragen und protokolliert jede Abfrage."""

    def __init__(self):
        self.treffer = defaultdict(lambda: defaultdict(set))  # qid -> klasse -> {(ort, über)}
        self.land_kontext = {}  # qid -> Ländercode der Abfrage, die es fand
        self.protokoll = []     # (land, ebene, region, klasse, status, anzahl)

    def aufnehmen(self, klasse, land, gefunden):
        for item, orte in gefunden.items():
            self.treffer[item][klasse] |= orte
            self.land_kontext.setdefault(item, land)


def hole_region(sammler, klasse, land, region, ebene, einrueckung=""):
    """Fragt eine Klasse in einer Region ab; bei Timeout eine Ebene tiefer.

    Ebenen: land → kreis → gemeinde. Unterhalb der Gemeinde wird nicht mehr
    geteilt; was dort noch reißt, steht als „fehlgeschlagen“ im Protokoll.
    """
    abfrage = abfrage_items(klasse, region=region,
                            typ_zuerst=(ebene == "land"))
    try:
        gefunden = treffer_lesen(sparql(abfrage))
        sammler.aufnehmen(klasse, land, gefunden)
        sammler.protokoll.append((land, ebene, region, klasse, "ok", len(gefunden)))
        print(f"{einrueckung}{klasse:12} {ebene:8} {region:10} {len(gefunden):6} Items")
        return
    except Zeitueberschreitung:
        print(f"{einrueckung}{klasse:12} {ebene:8} {region:10} Timeout → aufteilen")

    if ebene in ("land", "kreis"):
        try:
            teile, zwischen = teilgebiete(region, ebene)
        except Zeitueberschreitung:
            sammler.protokoll.append((land, ebene, region, klasse, "fehlgeschlagen", 0))
            print(f"{einrueckung}  !! Teilgebiete von {region}: Timeout – fehlt im Ergebnis")
            return
        naechste = "kreis" if ebene == "land" else "gemeinde"
    else:
        sammler.protokoll.append((land, ebene, region, klasse, "fehlgeschlagen", 0))
        print(f"{einrueckung}  !! {region} lässt sich nicht weiter teilen – fehlt im Ergebnis")
        return

    print(f"{einrueckung}  {len(teile)} Teilgebiete, {len(zwischen)} Zwischenebenen")
    if not teile:
        # Kreisfreie Stadt: sie ist selbst die Gemeinde, darunter gibt es nichts
        # zu teilen. Dann fehlen alle Items, die nur an einem Stadtteil hängen —
        # das soll in der Statistik auffallen, nicht untergehen.
        sammler.protokoll.append((land, naechste, region, klasse, "nur-zwischenebene", 0))
        print(f"{einrueckung}  !! {region}: keine Teilgebiete – Items an Stadtteilen fehlen")
    for teil in sorted(teile):
        hole_region(sammler, klasse, land, teil, naechste, einrueckung + "  ")
    # Items, die direkt am Gebiet oder an einer Zwischenebene hängen (z. B.
    # P131 = Regierungsbezirk), fallen durch alle Teilgebiete — eigens holen.
    try:
        gefunden = treffer_lesen(sparql(abfrage_items(klasse, orte=sorted(zwischen))))
        sammler.aufnehmen(klasse, land, gefunden)
        sammler.protokoll.append((land, ebene + "-rest", region, klasse, "ok", len(gefunden)))
    except Zeitueberschreitung:
        if KLASSEN[klasse]["verortung"] != "ort":
            sammler.protokoll.append((land, ebene + "-rest", region, klasse, "fehlgeschlagen", 0))
            print(f"{einrueckung}  !! Zwischenebene von {region}: Timeout")
            return
        # Zu viel hängt direkt dran (Frankfurt am Main: alle Firmen mit Sitz
        # „Frankfurt“). Je Ort und je Endziffer der QID — ein Zehntel der
        # Typprüfungen pro Abfrage.
        print(f"{einrueckung}  Zwischenebene von {region}: Timeout → je Ort und Endziffer")
        for ort in sorted(zwischen):
            for ziffer in "0123456789":
                try:
                    gefunden = treffer_lesen(sparql(abfrage_items(klasse, orte=[ort], endziffer=ziffer)))
                    sammler.aufnehmen(klasse, land, gefunden)
                    sammler.protokoll.append((land, ebene + "-rest", f"{ort}/{ziffer}", klasse, "ok", len(gefunden)))
                except Zeitueberschreitung:
                    sammler.protokoll.append((land, ebene + "-rest", f"{ort}/{ziffer}", klasse, "fehlgeschlagen", 0))
                    print(f"{einrueckung}  !! {ort}, Endziffer {ziffer}: Timeout")


# ---------------------------------------------------------------------------
# Details, Daten, Verortung
# ---------------------------------------------------------------------------

def hole_details(qids):
    """Label, Website, Koordinaten, de-Wikipedia, Auflösung (P576)."""
    details = {}
    for stapel in in_stapeln(qids):
        abfrage = f"""SELECT ?item ?label ?sprache ?web ?koord ?wp ?aufgeloest WHERE {{
  VALUES ?item {{ {werte_liste(stapel)} }}
  OPTIONAL {{ ?item rdfs:label ?label . BIND(LANG(?label) AS ?sprache)
             FILTER(?sprache IN ("de", "mul", "en")) }}
  OPTIONAL {{ ?item wdt:P856 ?web }}
  OPTIONAL {{ ?item wdt:P625 ?koord }}
  OPTIONAL {{ ?wp schema:about ?item ; schema:isPartOf <https://de.wikipedia.org/> }}
  OPTIONAL {{ ?item wdt:P576 ?aufgeloest }}
}}"""
        for b in sparql(abfrage):
            d = details.setdefault(qid(b, "item"), {
                "labels": {}, "web": None, "koord": None, "wp": None, "aufgeloest": None})
            if b.get("label"):
                d["labels"][wert(b, "sprache")] = wert(b, "label")
            for schluessel, feld in (("web", "web"), ("koord", "koord"), ("wp", "wp"),
                                     ("aufgeloest", "aufgeloest")):
                # Mehrere Websites/Koordinaten: die kleinste nehmen, damit das
                # Ergebnis bei jedem Lauf dasselbe ist. Für die Jubiläen egal.
                if b.get(feld) and (d[schluessel] is None or wert(b, feld) < d[schluessel]):
                    d[schluessel] = wert(b, feld)
    return details


def hole_daten(qids):
    """Alle Werte von P571 und P1249 mit Präzision und Rang — keiner wird verworfen."""
    daten = defaultdict(list)
    for stapel in in_stapeln(qids):
        abfrage = f"""SELECT ?item ?eig ?aussage ?zeit ?praez ?rang WHERE {{
  VALUES ?item {{ {werte_liste(stapel)} }}
  VALUES (?p ?psv ?eig) {{ (p:P571 psv:P571 "P571") (p:P1249 psv:P1249 "P1249") }}
  ?item ?p ?aussage .
  ?aussage ?psv ?wert ; wikibase:rank ?rang .
  ?wert wikibase:timeValue ?zeit ; wikibase:timePrecision ?praez .
}}"""
        for b in sparql(abfrage):
            rang = wert(b, "rang").rsplit("#", 1)[1].replace("Rank", "").lower()
            daten[qid(b, "item")].append({
                "eigenschaft": wert(b, "eig"),
                "aussage": wert(b, "aussage").rsplit("/", 1)[1],
                "wert": wert(b, "zeit"),
                "praezision": int(wert(b, "praez")),
                "rang": rang,  # preferred / normal / deprecated
            })
    return daten


def hole_verortung(orte):
    """P131-Kanten oberhalb jedes Orts, mit Merkmal Gemeinde/Kreis/aufgelöst.

    Es kommen Kanten statt einer flachen Vorfahrenliste zurück, weil P131*
    keine Reihenfolge kennt: erst mit den Kanten lässt sich die *nächste*
    Gemeinde über einem Ortsteil bestimmen (und nicht ein historischer
    Vorgänger, der über einen alten P131-Wert ebenfalls erreichbar ist).
    """
    kanten = defaultdict(set)
    merkmale = {}
    for stapel in in_stapeln(orte, 100):
        abfrage = f"""SELECT DISTINCT ?knoten ?eltern ?istGemeinde ?istKreis ?aufgeloest WHERE {{
  VALUES ?ort {{ {werte_liste(stapel)} }}
  ?ort wdt:P131* ?knoten .
  OPTIONAL {{ ?knoten wdt:P131 ?eltern }}
  BIND(EXISTS {{ ?knoten wdt:P31/wdt:P279* wd:{GEMEINDE_KLASSE} }} AS ?istGemeinde)
  BIND(EXISTS {{ VALUES ?kk {{ {werte_liste(KREIS_KLASSEN)} }} ?knoten wdt:P31/wdt:P279* ?kk }} AS ?istKreis)
  BIND(EXISTS {{ ?knoten wdt:P576 [] }} AS ?aufgeloest)
}}"""
        for b in sparql(abfrage):
            k = qid(b, "knoten")
            if b.get("eltern"):
                kanten[k].add(qid(b, "eltern"))
            merkmale[k] = {
                "gemeinde": wert(b, "istGemeinde") == "true",
                "kreis": wert(b, "istKreis") == "true",
                "aufgeloest": wert(b, "aufgeloest") == "true",
            }
    return kanten, merkmale


def naechste(start, kanten, merkmale, art):
    """Breitensuche nach oben: erste aktuelle Einheit der Art, sonst erste historische."""
    gesehen, schlange, historisch = {start}, deque([start]), None
    while schlange:
        k = schlange.popleft()
        m = merkmale.get(k)
        if m and m[art]:
            if not m["aufgeloest"]:
                return k
            historisch = historisch or k
        for e in sorted(kanten.get(k, ())):
            if e not in gesehen:
                gesehen.add(e)
                schlange.append(e)
    return historisch


def land_von(start, kanten):
    land_qids = {q: code for code, (q, _) in LAENDER.items()}
    gesehen, schlange = {start}, deque([start])
    while schlange:
        k = schlange.popleft()
        if k in land_qids:
            return land_qids[k]
        for e in sorted(kanten.get(k, ())):
            if e not in gesehen:
                gesehen.add(e)
                schlange.append(e)
    return None


# ---------------------------------------------------------------------------
# Wikipedia: Prüfliste und Nachträge
#
# Wikidata kennt die Ersterwähnung fast keines Ortes — im Wikipedia-Artikel
# steht sie fast immer, aber nur im Fließtext („wurde 767 im Lorscher Codex
# erstmals erwähnt“). Ein Regex entscheidet dort zu oft falsch (Jahr einer
# Kirche, eines Ortsteils, „erstmals seit 1974“). Deshalb teilt sich die
# Arbeit: das Skript schneidet die Kandidatensätze aus (grobes Sieb, lieber
# zu viel als zu wenig), ein Mensch oder Claude entscheidet nach
# pruefanleitung.md und trägt das Ergebnis mit Belegsatz in nachtraege.csv ein.
# ---------------------------------------------------------------------------

SCHLUESSELWOERTER = re.compile(
    r"erwähn|erstmal|urkundl|gründ|codex|kodex|genannt|nennung|bezeugt|belegt|nachgewiesen"
    r"|ältest|datiert|zurückverfolg",
    re.I)
JAHRESZAHL = re.compile(r"\b\d{3,4}\b")
# Abkürzungen, nach deren Punkt kein neuer Satz beginnt („im 8. Jh. erwähnt“)
ABKUERZUNG = re.compile(r"(?:\be|\bJh|\bJhd|\bJhs|\bca|\bbzw|\bv\. Chr|\bn\. Chr|\bChr|\bSt|\bNr|\bz\. B|\bu\. a|\bggf|\bsog|\bvgl)\.$")
INFOBOX_GRUENDUNG = re.compile(
    r"^\s*\|\s*(Gründungsdatum|Gründung|gegründet|Gegründet)\s*=\s*(.+)$", re.M)


def _wp_cache_pfad(titel):
    return CACHE_ORDNER / ("wp_" + hashlib.sha1(titel.encode("utf-8")).hexdigest() + ".json")


def hole_wikipedia(titel):
    """Wikitext je Artikeltitel; jeder Titel einzeln gecacht, 50 je Anfrage geholt."""
    texte, offen = {}, []
    for t in titel:
        pfad = _wp_cache_pfad(t)
        if pfad.exists():
            texte[t] = json.loads(pfad.read_text(encoding="utf-8"))["text"]
        else:
            offen.append(t)
    for stapel in in_stapeln(offen, 50):
        normalisiert, weiter, inhalte = {}, {}, {}
        fortsetzung = {}
        while True:
            daten = _wikipedia_anfrage({
                "action": "query", "prop": "revisions", "rvprop": "content",
                "rvslots": "main", "titles": "|".join(stapel), "redirects": 1,
                "format": "json", "formatversion": 2, **fortsetzung})
            abfrage = daten["query"]
            normalisiert.update({x["from"]: x["to"] for x in abfrage.get("normalized", [])})
            weiter.update({x["from"]: x["to"] for x in abfrage.get("redirects", [])})
            for s in abfrage["pages"]:
                if "revisions" in s:
                    inhalte[s["title"]] = s["revisions"][0]["slots"]["main"]["content"]
            # Wird die Antwort zu groß, liefert die API einen Teil der Texte
            # erst nach „continue“ — ohne diese Schleife fehlen sie stumm.
            if "continue" not in daten:
                break
            fortsetzung = daten["continue"]
        # Die API normalisiert Titel und folgt Weiterleitungen; zurück auf den
        # angefragten Titel abbilden, sonst findet der Cache ihn nicht wieder.
        for t in stapel:
            n = normalisiert.get(t, t)
            text = inhalte.get(weiter.get(n, n), "")
            CACHE_ORDNER.mkdir(exist_ok=True)
            _wp_cache_pfad(t).write_text(json.dumps({"titel": t, "text": text}, ensure_ascii=False),
                                         encoding="utf-8")
            texte[t] = text
    return texte


def _wikipedia_anfrage(parameter):
    for versuch in range(1, MAX_VERSUCHE + 1):
        _warte_pause()
        statistik_abfragen["netz"] += 1
        antwort = _hole_sitzung().get(WP_API, params=parameter, timeout=CLIENT_TIMEOUT_S)
        if antwort.status_code == 429 or antwort.status_code >= 500:
            _backoff(versuch, f"Wikipedia HTTP {antwort.status_code}")
            continue
        antwort.raise_for_status()
        return antwort.json()
    raise RuntimeError("Wikipedia antwortet nicht")


def wikitext_zu_text(wt):
    """Grob lesbarer Fließtext: Vorlagen, Einzelnachweise, Bilder und Tabellen raus."""
    wt = re.sub(r"<!--.*?-->", "", wt, flags=re.S)
    wt = re.sub(r"<ref[^>]*/>", "", wt)
    wt = re.sub(r"<ref[^>]*>.*?</ref>", "", wt, flags=re.S)
    vorher = None
    while vorher != wt:  # verschachtelte Vorlagen von innen nach außen
        vorher = wt
        wt = re.sub(r"\{\{[^{}]*\}\}", "", wt)
    wt = re.sub(r"\{\|.*?\|\}", "", wt, flags=re.S)
    wt = re.sub(r"\[\[(?:Datei|Bild|File|Image|Kategorie):(?:[^\[\]]|\[\[[^\]]*\]\])*\]\]", "", wt, flags=re.I)
    wt = re.sub(r"\[\[(?:[^|\]]*\|)?([^\]]*)\]\]", r"\1", wt)
    wt = re.sub(r"\[https?://\S+ ([^\]]*)\]", r"\1", wt)
    wt = re.sub(r"'{2,}", "", wt)
    wt = re.sub(r"<[^>]+>", "", wt)
    return wt.replace("&nbsp;", " ")


def _saetze(zeile):
    teile = re.split(r"(?<=[a-zäöüß)\]\"“'][.!?])\s+", zeile)
    saetze = []
    for teil in teile:
        if saetze and ABKUERZUNG.search(saetze[-1]):
            saetze[-1] += " " + teil
        else:
            saetze.append(teil)
    return [s.strip() for s in saetze if s.strip()]


def kandidatensaetze(wikitext, label=None, hoechstens=20):
    """(Abschnitt, Satz) für alle Sätze mit Jahreszahl und einem Schlüsselwort.

    Der Abschnitt gehört dazu, weil in Artikeln über zusammengelegte Gemeinden
    die Ersterwähnungen unter den Ortsteilen stehen („=== Wössingen ===“) —
    ohne Überschrift sähe das Jahr aus, als gehöre es der Gemeinde.

    Dazu kommen immer der Anfang der Einleitung (dort steht, ob die Gemeinde
    ein Zusammenschluss ist oder früher anders hieß — Philippsburg „bis 1623
    Udenheim“) und, falls es ihn gibt, der Abschnitt mit dem Namen des Items:
    viele Ortsteil-Artikel leiten auf die Gemeinde weiter, und ihr Abschnitt
    nennt die Ersterwähnung nicht immer mit einem der Schlüsselwörter.
    """
    name = re.sub(r"\s*\(.*\)$", "", label or "").strip()
    treffer = [("Infobox", f"{m.group(1)} = {wikitext_zu_text(m.group(2)).strip()}")
               for m in INFOBOX_GRUENDUNG.finditer(wikitext)]
    abschnitte, rueckfall, einleitung, eigener = [], [], [], []
    for zeile in wikitext_zu_text(wikitext).split("\n"):
        kopf = re.match(r"^(=+)\s*(.*?)\s*=+\s*$", zeile)
        if kopf:
            abschnitte = abschnitte[:len(kopf.group(1)) - 2] + [kopf.group(2)]
            continue
        ort = " > ".join(abschnitte) or "Einleitung"
        if zeile.strip():
            if not abschnitte:
                einleitung.append(zeile.strip())
            elif name and name in abschnitte[-1]:
                eigener.append(zeile.strip())
        for satz in _saetze(zeile):
            if JAHRESZAHL.search(satz) and SCHLUESSELWOERTER.search(satz):
                treffer.append((ort, satz[:400]))
            elif "Geschichte" in ort and len(" ".join(rueckfall)) < 1200:
                rueckfall.append(satz)
    if len(treffer) > hoechstens:
        treffer = treffer[:hoechstens] + [("…", f"{len(treffer) - hoechstens} weitere Sätze abgeschnitten")]
    # Kein Kandidat: dann den Anfang des Geschichtsteils, damit der Prüfer
    # nicht ins Leere schaut, sondern sieht, dass dort wirklich nichts steht.
    if not treffer and rueckfall:
        treffer = [("Geschichte (Anfang, kein Kandidat)", " ".join(rueckfall)[:1200])]
    vorspann = []
    if einleitung:
        vorspann.append(("Einleitung (Anfang)", " ".join(einleitung)[:300]))
    if eigener:
        vorspann.append((f"Abschnitt „{name}“", " ".join(eigener)[:700]))
    return vorspann + treffer


def lies_nachtraege():
    nachtraege = defaultdict(list)
    if NACHTRAEGE.exists():
        with open(NACHTRAEGE, encoding="utf-8-sig", newline="") as f:
            for zeile in csv.DictReader(f, delimiter=";"):
                nachtraege[zeile["qid"].strip()].append(
                    {k: (zeile.get(k) or "").strip() for k in NACHTRAEGE_SPALTEN})
    return nachtraege


def schreibe_pruefliste(db, pfad):
    """Alle Items, denen das bevorzugte Datum fehlt und die einen Artikel haben.

    Orte stehen drauf, solange ihnen eine Ersterwähnung fehlt — auch wenn
    Wikidata ein P571 hat, denn das ist bei zusammengelegten Gemeinden das
    Jahr der Gebietsreform, nicht der Anlass für ein Ortsjubiläum.
    """
    erledigt = set(lies_nachtraege())
    zeilen = db.execute("""
        SELECT i.qid, i.label, i.klasse, i.gemeinde, i.wikipedia_de,
               (SELECT group_concat(coalesce(d.eigenschaft, 'Name') || ' ' || d.jahr
                       || CASE WHEN d.verwendbar = 0 THEN ' (ungenau)' ELSE '' END, ', ')
                  FROM daten d WHERE d.qid = i.qid) AS vorhanden
        FROM items i
        WHERE i.wikipedia_de IS NOT NULL AND (
          (i.klasse IN ('gemeinde', 'ortsteil') AND NOT EXISTS (
             SELECT 1 FROM daten d WHERE d.qid = i.qid AND d.eigenschaft = 'P1249' AND d.verwendbar = 1))
          OR (i.klasse NOT IN ('gemeinde', 'ortsteil') AND i.datum_status IN ('label', 'ungenau', 'keins')))
        ORDER BY i.gemeinde, i.klasse != 'gemeinde', i.klasse, i.label""").fetchall()
    # Nach Gemeinde sortiert, Gemeinde zuerst: so stehen Gemeinde und ihre
    # Ortsteile im selben Block — auch wenn die Liste später in Pakete für
    # mehrere Prüfer geteilt wird. Der Gemeindeartikel ist oft der einzige
    # Beleg für einen Ortsteil.
    zeilen = [z for z in zeilen if z[0] not in erledigt]
    titel = {z[0]: requests.utils.unquote(z[4].rsplit("/wiki/", 1)[1]).replace("_", " ")
             for z in zeilen}
    texte = hole_wikipedia(sorted(set(titel.values())))
    with open(pfad, "w", encoding="utf-8") as f:
        f.write(f"# Prüfliste – {len(zeilen)} Items\n\n"
                "Regeln: pruefanleitung.md. Ergebnis je Item eine Zeile in nachtraege.csv.\n\n")
        for q, label, klasse, gemeinde, url, vorhanden in zeilen:
            f.write(f"## {q} | {label} | {klasse} | Gemeinde: {gemeinde or '?'}"
                    f" | Wikidata: {vorhanden or '–'} | {url}\n")
            for abschnitt, satz in (kandidatensaetze(texte.get(titel[q], ""), label)
                                    or [("–", "Artikel leer oder nicht gefunden")]):
                f.write(f"- [{abschnitt}] {satz}\n")
            f.write("\n")
    return len(zeilen)


# ---------------------------------------------------------------------------
# Jubiläen
# ---------------------------------------------------------------------------

def ist_rund(alter):
    """Runde Jubiläen: alle 25 Jahre bis 200, danach volle Hunderter und 250/750/1250.

    250, 750 und 1250 stehen extra drin, weil genau diese Ortsjubiläen in der
    Region ständig gefeiert werden (1250 Jahre = Ersterwähnung im Lorscher
    Codex um 770). 1100, 1200 … ergeben sich aus den vollen Hundertern und
    betreffen in der Praxis nur Gemeinden.
    """
    if alter <= 0:
        return False
    return (alter <= 200 and alter % 25 == 0) or alter % 100 == 0 or alter % 250 == 0


def jahr_aus_zeit(zeitwert):
    """'1904-01-01T00:00:00Z' → 1904; vorchristliche Jahre → None."""
    treffer = re.match(r"^(-?)\+?(\d+)-", zeitwert or "")
    if not treffer or treffer.group(1) == "-":
        return None
    return int(treffer.group(2))


# ---------------------------------------------------------------------------
# Datenbank
# ---------------------------------------------------------------------------

SCHEMA = """
CREATE TABLE items (
    qid TEXT PRIMARY KEY,
    label TEXT,
    klasse TEXT,              -- Hauptklasse nach KLASSEN_VORRANG
    bundesland TEXT,
    landkreis_qid TEXT, landkreis TEXT,
    gemeinde_qid TEXT, gemeinde TEXT,
    ort_qid TEXT, ort_ueber TEXT,  -- P159 / P131 / selbst
    website TEXT, breite REAL, laenge REAL, wikipedia_de TEXT,
    aufgeloest TEXT,          -- P576, falls gesetzt (z. B. eingemeindeter Ort)
    datum_status TEXT,        -- Quelle des bevorzugten Jahres: wikidata / wikipedia / label;
                              -- sonst ungenau / keins
    mehrdeutig INTEGER        -- 1 = mehrere gleichrangige Jahre im bevorzugten Feld
);
CREATE TABLE item_klassen (qid TEXT, klasse TEXT, PRIMARY KEY (qid, klasse));
CREATE TABLE daten (
    id INTEGER PRIMARY KEY,
    qid TEXT,
    eigenschaft TEXT,         -- P571 (Gründung) / P1249 (Ersterwähnung) — auch bei
                              -- quelle = wikipedia als Art des Datums; NULL bei label
    quelle TEXT,              -- wikidata / wikipedia (aus nachtraege.csv) / label
    aussage TEXT,
    wert TEXT,                -- Zeitwert, bei wikipedia der Belegsatz, bei label der Name
    jahr INTEGER, praezision INTEGER, rang TEXT,
    verwendbar INTEGER,       -- Jahrespräzision oder besser, nicht deprecated
    bevorzugt INTEGER,        -- wird für die Jubiläen in erster Linie herangezogen
    bemerkung TEXT            -- nur bei wikipedia: Unsicherheit, Bezug auf einen Ortsteil …
);
CREATE TABLE jubilaeen (
    qid TEXT, daten_id INTEGER, jahr INTEGER, jubilaeum INTEGER, bevorzugt INTEGER
);
CREATE TABLE abfragen (land TEXT, ebene TEXT, region TEXT, klasse TEXT, status TEXT, anzahl INTEGER);
CREATE INDEX jub_jahr ON jubilaeen (jahr);
CREATE INDEX daten_qid ON daten (qid);

-- SQLite-Views kennen keine Parameter; „jubilaeen_in_jahr(X)“ ist daher
-- diese View mit WHERE jahr = X, oder auf der Kommandozeile --jahr X.
CREATE VIEW jubilaeen_in_jahr AS
SELECT j.jahr, j.jubilaeum, i.label, i.klasse, i.bundesland, i.landkreis, i.gemeinde,
       d.jahr AS gruendungsjahr, d.eigenschaft, d.quelle, d.praezision, j.bevorzugt,
       i.website, i.wikipedia_de, i.qid
FROM jubilaeen j JOIN items i USING (qid) JOIN daten d ON d.id = j.daten_id;
"""


def bevorzugte_daten(klasse, daten):
    """Markiert, welche Werte für die Jubiläen in erster Linie gelten.

    Gemeinden/Ortsteile: P1249 vor P571, alle anderen umgekehrt. Hat das
    bevorzugte Feld mehrere Werte, gewinnt ein Wert mit Rang „preferred“;
    sonst bleiben alle bevorzugt und das Item gilt als mehrdeutig — die
    Entscheidung soll ein Mensch treffen, nicht die Reihenfolge der Antwort.
    """
    reihenfolge = ["P1249", "P571"] if klasse in ORTS_KLASSEN else ["P571", "P1249"]
    verwendbar = [d for d in daten if d["verwendbar"]]
    for eig in reihenfolge:
        kandidaten = [d for d in verwendbar if d["eigenschaft"] == eig]
        if kandidaten:
            bevorzugt = [d for d in kandidaten if d["rang"] == "preferred"] or kandidaten
            for d in bevorzugt:
                d["bevorzugt"] = 1
            return len({d["jahr"] for d in bevorzugt}) > 1
    return False


def baue_datenbank(sammler, db_pfad, horizont_von, horizont_bis):
    qids = sorted(sammler.treffer)
    nachtraege = lies_nachtraege()
    print(f"\nDetails für {len(qids)} Items …")
    details = hole_details(qids)
    print("Gründungs- und Ersterwähnungsdaten …")
    rohdaten = hole_daten(qids)

    # Für die Verortung pro Item einen Ort wählen: Hauptsitz vor Lage, weil
    # P131 bei Firmen oft auf eine Niederlassung oder den Kreis zeigt.
    ort_wahl = {}
    for q, klassen in sammler.treffer.items():
        orte = set().union(*klassen.values())
        for ueber in ("selbst", "P159", "P131"):
            passend = sorted(o for o, u in orte if u == ueber)
            if passend:
                ort_wahl[q] = (passend[0], ueber)
                break
    print("Verortung (Gemeinde, Landkreis) …")
    kanten, merkmale = hole_verortung({o for o, _ in ort_wahl.values()})
    einheiten = {e for e in (
        naechste(o, kanten, merkmale, art) for o, _ in ort_wahl.values()
        for art in ("gemeinde", "kreis")) if e}
    namen = hole_details(einheiten - set(details))
    alle_details = {**namen, **details}

    def label(q):
        if not q:
            return None
        lb = alle_details.get(q, {}).get("labels", {})
        return lb.get("de") or lb.get("mul") or lb.get("en")

    if db_pfad.exists():
        db_pfad.unlink()
    db = sqlite3.connect(db_pfad)
    db.executescript(SCHEMA)
    db.executemany("INSERT INTO abfragen VALUES (?,?,?,?,?,?)", sammler.protokoll)

    for q in qids:
        klassen = sammler.treffer[q]
        hauptklasse = next(k for k in KLASSEN_VORRANG if k in klassen)
        d = details.get(q, {"labels": {}})
        name = label(q)
        ort, ueber = ort_wahl[q]
        gemeinde = naechste(ort, kanten, merkmale, "gemeinde")
        kreis = naechste(ort, kanten, merkmale, "kreis")
        land = land_von(ort, kanten) or sammler.land_kontext.get(q)
        breite = laenge = None
        koord = re.match(r"Point\(([-\d.eE]+) ([-\d.eE]+)\)", d.get("koord") or "")
        if koord:
            laenge, breite = float(koord.group(1)), float(koord.group(2))

        daten = []
        for r in rohdaten.get(q, []):
            jahr = jahr_aus_zeit(r["wert"])
            daten.append({**r, "quelle": "wikidata", "jahr": jahr, "bevorzugt": 0,
                          "verwendbar": int(jahr is not None
                                            and r["praezision"] >= MIN_PRAEZISION
                                            and r["rang"] != "deprecated")})
        for n in nachtraege.get(q, []):
            if not n["jahr"]:
                continue  # nachgesehen, nichts gefunden — hält das Item nur von der Prüfliste fern
            praezision = int(n["praezision"] or MIN_PRAEZISION)
            daten.append({
                "eigenschaft": "P1249" if n["art"] == "ersterwaehnung" else "P571",
                "quelle": "wikipedia", "aussage": None, "wert": n["beleg"],
                "jahr": int(n["jahr"]), "praezision": praezision, "rang": "normal",
                "verwendbar": int(praezision >= MIN_PRAEZISION), "bevorzugt": 0,
                "bemerkung": "; ".join(x for x in (
                    "unsicher" if n["unsicher"] == "1" else "", n["bemerkung"]) if x) or None})
        mehrdeutig = bevorzugte_daten(hauptklasse, daten)
        quellen = {x["quelle"] for x in daten if x["bevorzugt"]}
        if quellen:
            # Widersprechen sich Wikidata und Wikipedia, steht das Item als
            # mehrdeutig da; der Status nennt dann Wikidata.
            status = "wikidata" if "wikidata" in quellen else "wikipedia"
        elif (hauptklasse == "verein" and name and (treffer := LABEL_JAHR.search(name))
              and int(treffer.group(1)) <= horizont_von):
            # „TSV 1904 Musterdorf“: das Jahr im Namen ist fast immer das
            # Gründungsjahr. Als eigene Quelle markiert, damit man es
            # nachprüft, bevor man jemandem zum 125-jährigen gratuliert.
            daten.append({"eigenschaft": None, "quelle": "label", "aussage": None,
                          "wert": name, "jahr": int(treffer.group(1)),
                          "praezision": MIN_PRAEZISION, "rang": None,
                          "verwendbar": 1, "bevorzugt": 1})
            status = "label"
        else:
            status = "ungenau" if daten else "keins"

        db.execute("INSERT INTO items VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", (
            q, name, hauptklasse, LAENDER[land][1] if land else None,
            kreis, label(kreis), gemeinde, label(gemeinde), ort, ueber,
            d.get("web"), breite, laenge, d.get("wp"), d.get("aufgeloest"),
            status, int(mehrdeutig)))
        db.executemany("INSERT INTO item_klassen VALUES (?,?)", [(q, k) for k in klassen])
        for x in daten:
            cur = db.execute(
                "INSERT INTO daten (qid, eigenschaft, quelle, aussage, wert, jahr, praezision,"
                " rang, verwendbar, bevorzugt, bemerkung) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                (q, x["eigenschaft"], x["quelle"], x["aussage"], x["wert"], x["jahr"],
                 x["praezision"], x["rang"], x["verwendbar"], x["bevorzugt"],
                 x.get("bemerkung")))
            if x["verwendbar"]:
                db.executemany("INSERT INTO jubilaeen VALUES (?,?,?,?,?)", [
                    (q, cur.lastrowid, j, j - x["jahr"], x["bevorzugt"])
                    for j in range(horizont_von, horizont_bis + 1) if ist_rund(j - x["jahr"])])
    db.commit()
    return db


# ---------------------------------------------------------------------------
# Ausgabe
# ---------------------------------------------------------------------------

def schreibe_csv(db, pfad, abfrage, parameter=()):
    cur = db.execute(abfrage, parameter)
    # Semikolon + BOM: so öffnet ein deutsches Excel die Datei per Doppelklick
    # mit Spalten und Umlauten, ohne Importassistent.
    with open(pfad, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow([c[0] for c in cur.description])
        w.writerows(cur)


ITEMS_CSV = """
SELECT i.*,
  (SELECT group_concat(DISTINCT d.jahr) FROM daten d
    WHERE d.qid = i.qid AND d.bevorzugt = 1) AS jahr_bevorzugt,
  (SELECT min(j.jahr) || ' (' || j.jubilaeum || ')' FROM jubilaeen j
    WHERE j.qid = i.qid AND j.bevorzugt = 1 AND j.jahr >= :heute) AS naechstes_jubilaeum
FROM items i ORDER BY bundesland, klasse, label
"""


def statistik(db):
    print("\nCoverage je Klasse und Bundesland (Quelle des bevorzugten Jahres)")
    print("  wikidata  = Jahr aus P571/P1249 (≥ Jahrespräzision)")
    print("  wikipedia = aus dem Artikel nachgetragen (nachtraege.csv)")
    print("  label     = nur Jahr aus dem Vereinsnamen")
    print("  ungenau   = Datum vorhanden, aber nur Jahrzehnt/Jahrhundert o. ä.\n")
    kopf = (f"{'Bundesland':20} {'Klasse':12} {'gesamt':>7} {'wikidata':>9} {'wikipedia':>10}"
            f" {'label':>6} {'ungenau':>8} {'keins':>7} {'% mit Jahr':>10}")
    print(kopf)
    print("-" * len(kopf))
    zeilen = db.execute("""
        SELECT coalesce(bundesland, '?'), klasse, count(*),
               sum(datum_status = 'wikidata'), sum(datum_status = 'wikipedia'),
               sum(datum_status = 'label'), sum(datum_status = 'ungenau'),
               sum(datum_status = 'keins')
        FROM items GROUP BY 1, 2 ORDER BY 1, 2""").fetchall()
    for land, klasse, n, wd, wp, lb, ung, keins in zeilen:
        print(f"{land:20} {klasse:12} {n:7} {wd:9} {wp:10} {lb:6} {ung:8} {keins:7}"
              f" {100 * (wd + wp + lb) / n:10.0f}")
    pro_eig = db.execute("""
        SELECT i.klasse, d.quelle, d.eigenschaft, count(DISTINCT d.qid) FROM daten d JOIN items i USING (qid)
        WHERE d.verwendbar = 1 AND d.quelle != 'label' GROUP BY 1, 2, 3 ORDER BY 1, 2, 3""").fetchall()
    print("\nItems mit verwendbarem Wert je Quelle und Eigenschaft:")
    for klasse, quelle, eig, n in pro_eig:
        print(f"  {klasse:12} {quelle:10} {eig:6} {n}")
    mehr = db.execute("SELECT klasse, count(*) FROM items WHERE mehrdeutig = 1 GROUP BY 1").fetchall()
    if mehr:
        print("\nMehrdeutig (mehrere gleichrangige Jahre im bevorzugten Feld): "
              + ", ".join(f"{k} {n}" for k, n in mehr))
    fehl = db.execute("SELECT land, ebene, region, klasse, status FROM abfragen WHERE status != 'ok'").fetchall()
    if fehl:
        print("\n!! Unvollständig (fehlgeschlagen = Timeout auf unterster Ebene,"
              " nur-zwischenebene = Stadtteile einer kreisfreien Stadt fehlen):")
        for f in fehl:
            print("   ", *f)


def zeige_jahr(db, jahr, alle, ausgabe_pfad):
    bedingung = "" if alle else " AND d.bevorzugt = 1"
    abfrage = f"""
        SELECT :jahr - d.jahr AS jubilaeum, i.label, i.klasse, i.bundesland, i.landkreis,
               i.gemeinde, d.jahr AS gruendungsjahr, d.eigenschaft AS feld, d.quelle,
               d.bevorzugt, i.mehrdeutig, d.bemerkung, i.website, i.wikipedia_de, i.qid
        FROM daten d JOIN items i USING (qid)
        WHERE d.verwendbar = 1{bedingung}
        GROUP BY i.qid, d.jahr
        ORDER BY 1 DESC, i.klasse, i.label"""
    # Direkt aus den Daten gerechnet statt aus der Tabelle jubilaeen, damit
    # --jahr auch außerhalb des vorberechneten Horizonts funktioniert.
    db.create_function("ist_rund", 1, lambda a: int(ist_rund(a)))
    abfrage = abfrage.replace("GROUP BY", "AND ist_rund(:jahr - d.jahr) GROUP BY")
    zeilen = db.execute(abfrage, {"jahr": jahr}).fetchall()
    print(f"\nJubiläen {jahr}: {len(zeilen)} Treffer"
          + ("" if alle else " (nur bevorzugte Daten; --alle zeigt auch die übrigen)"))
    for z in zeilen[:40]:
        print(f"  {z[0]:5}  {z[6]:>4}  {z[8]:9} {z[2]:11} {(z[1] or z[14])[:45]:45} {z[5] or ''}")
    if len(zeilen) > 40:
        print(f"  … {len(zeilen) - 40} weitere in {ausgabe_pfad.name}")
    schreibe_csv(db, ausgabe_pfad, abfrage, {"jahr": jahr})


def main():
    p = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    p.add_argument("--land", action="append", choices=sorted(LAENDER),
                   help="nur diese Länder (mehrfach möglich); Standard: alle vier")
    p.add_argument("--kreis", help="nur diesen Kreis (QID), z. B. Q8178 = Landkreis Karlsruhe")
    p.add_argument("--klasse", action="append", choices=list(KLASSEN))
    p.add_argument("--jahr", type=int, help="alle Jubiläen in diesem Jahr ausgeben")
    p.add_argument("--alle", action="store_true",
                   help="bei --jahr auch nicht bevorzugte Daten (z. B. P571 neben P1249)")
    p.add_argument("--nur-auswerten", action="store_true",
                   help="nichts abfragen, nur die vorhandene DB auswerten")
    p.add_argument("--pruefliste", action="store_true",
                   help="Wikipedia-Kandidatensätze für Items ohne Datum ausgeben (zum Nachtragen)")
    p.add_argument("--horizont", type=int, default=30,
                   help="Jahre ab heute, für die die Tabelle jubilaeen gefüllt wird")
    args = p.parse_args()

    name = f"jubilaeen_{args.kreis}" if args.kreis else "jubilaeen"
    db_pfad = ORDNER / f"{name}.sqlite"
    heute = datetime.date.today().year

    if args.nur_auswerten:
        db = sqlite3.connect(db_pfad)
    else:
        sammler = Sammler()
        klassen = args.klasse or list(KLASSEN)
        if args.kreis:
            for k in klassen:
                hole_region(sammler, k, None, args.kreis, "kreis")
        else:
            for code in args.land or STANDARD_LAENDER:
                print(f"\n== {LAENDER[code][1]} ==")
                for k in klassen:
                    hole_region(sammler, k, code, LAENDER[code][0], "land")
            if not args.land:
                for code, kreise in ZUSATZ_KREISE.items():
                    print(f"\n== {LAENDER[code][1]} (Grenzkreise) ==")
                    for kreis in kreise:
                        for k in klassen:
                            hole_region(sammler, k, code, kreis, "kreis")
        db = baue_datenbank(sammler, db_pfad, heute, heute + args.horizont)
        schreibe_csv(db, ORDNER / f"{name}_items.csv", ITEMS_CSV, {"heute": heute})
        schreibe_csv(db, ORDNER / f"{name}_daten.csv",
                     "SELECT d.*, i.label, i.klasse FROM daten d JOIN items i USING (qid) ORDER BY qid, id")
        print(f"\nAbfragen: {statistik_abfragen['netz']} übers Netz, "
              f"{statistik_abfragen['cache']} aus dem Cache, {statistik_abfragen['timeout']} Timeouts")
        print(f"Geschrieben: {db_pfad.name}, {name}_items.csv, {name}_daten.csv")

    statistik(db)
    if args.pruefliste:
        pfad = ORDNER / f"{name}_pruefliste.md"
        anzahl = schreibe_pruefliste(db, pfad)
        print(f"\nPrüfliste: {anzahl} Items ohne Datum in {pfad.name}")
    if args.jahr:
        zeige_jahr(db, args.jahr, args.alle, ORDNER / f"{name}_jahr_{args.jahr}.csv")


if __name__ == "__main__":
    main()
