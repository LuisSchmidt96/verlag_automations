"""Lead-Liste 2029 – Stufen und Kommentare von Hand (Durchsicht vom 26.10.2026).

Keine Punkteformel: jede Zeile ist eine Einzelentscheidung. Grundlage sind
empfehlungen_2029.csv (Daten, Belege) und Allgemeinwissen über die Firmen –
Letzteres ist als „(aus dem Gedächtnis)“ markiert und vor einer Anfrage zu prüfen.

Stufen:
  S  Top-Lead: große runde Zahl, Träger mit Geld, gute Chance auf ein Buch
  A  guter Lead: aktiv ansprechen
  B  möglich: mitnehmen, wenn ohnehin Kontakt besteht
  C  unwahrscheinlich: kleines Budget, kleine Zahl oder weit weg vom Programm
  X  kein Lead: aufgelöst, insolvent, selbst Verlag, Gebäude, Datenfehler

Aufruf: python leads_2029_einschaetzung.py
  → leads_2029.csv                      (alle 339 Treffer mit Stufe)
  → leads_2029_S_A_B_verlagsgebiet.csv  (S, A und B in BW und der Pfalz, mit Kommentar)
"""

import csv
from pathlib import Path

ORDNER = Path(__file__).resolve().parent

PFALZ = {"Landkreis Bad Dürkheim", "Donnersbergkreis", "Landkreis Germersheim", "Landkreis Kaiserslautern",
         "Landkreis Kusel", "Landkreis Südliche Weinstraße", "Rhein-Pfalz-Kreis", "Landkreis Südwestpfalz",
         "Frankenthal (Pfalz)", "Frankenthal", "Kaiserslautern", "Landau in der Pfalz", "Ludwigshafen am Rhein",
         "Neustadt an der Weinstraße", "Pirmasens", "Speyer", "Zweibrücken"}

# Rang (aus empfehlungen_2029.csv) → (Stufe, Kurzkommentar)
E = {
    1: ("S", "Stadt, 1250 Jahre, Beleg sauber (Fulda 779), mitten im Verlagsgebiet"),
    2: ("S", "Stadt mit 14 000 Einwohnern, 1250 Jahre; hat 1979 schon die 1200 gefeiert – feiert also"),
    3: ("S", "800 Jahre, 6 000 Einwohner, Rhein-Pfalz-Kreis direkt vor der Haustür"),
    4: ("A", "800 Jahre, kleine Gemeinde (2 900), Enzkreis – Verlagsgebiet"),
    5: ("A", "1250 Jahre, 6 300 Einwohner"),
    6: ("A", "1250 Jahre (St. Gallen 779), Bodensee – weit weg, aber starke Zahl"),
    7: ("S", "Stadt, 1250 Jahre, 8 000 Einwohner, Beleg sauber"),
    8: ("A", "1250 Jahre (Lorscher Codex), 6 300 Einwohner"),
    9: ("S", "historische Stadt, 1200 Jahre, Tourismusort – Buch passt zum Stadtmarketing"),
    10: ("A", "1100 Jahre, kleine Gemeinde am Hochrhein; Ortsteile noch älter"),
    11: ("A", "900 Jahre Ort und Kloster Denkendorf (1129), 11 000 Einwohner"),
    12: ("A", "größte Stadt der Liste (28 500), aber Ersterwähnung 779 nur „könnte“ – vorher klären"),
    13: ("A", "800 Jahre, 5 600 Einwohner"),
    14: ("B", "750 Jahre, 3 000 Einwohner"),
    15: ("B", "750 Jahre, 2 600 Einwohner"),
    16: ("S", "1250 Jahre Dürrmenz, ältester Teil von Mühlacker – die Stadt dürfte feiern"),
    17: ("A", "1250 Jahre; Urkunde gilt für alle drei Sachsen-Orte (Hirschberg und Weinheim)"),
    18: ("C", "900 Jahre, aber 80 Einwohner"),
    19: ("B", "900 Jahre, 380 Einwohner – kleines Festheft eher als Buch"),
    20: ("C", "750 Jahre, 240 Einwohner"),
    21: ("C", "700 Jahre, 440 Einwohner"),
    22: ("B", "900 Jahre, 860 Einwohner"),
    23: ("B", "800 Jahre, 600 Einwohner"),
    24: ("B", "Zell am Neckar, heute Stadtteil von Esslingen, 800 Jahre"),
    25: ("B", "700 Jahre, Ortsteil von Altensteig, Kreis Calw"),
    26: ("A", "1250 Jahre, ältester Ortsteil von Zimmern ob Rottweil – klassische Ortschronik"),
    27: ("A", "1250 Jahre, Ortsteil von Öhringen"),
    28: ("B", "1200 Jahre, aber ursprünglich Hofgut, kleiner Ortsteil"),
    29: ("B", "1100 Jahre, kleiner Ortsteil von Stühlingen"),
    30: ("B", "800 Jahre, 2 500 Einwohner, Ortsteil von Wangen im Allgäu"),
    31: ("B", "800 Jahre, 9 100 Einwohner – Stuttgarter Stadtbezirk, „zusammen mit Stuttgart“ 1229 erwähnt"),
    32: ("C", "900 Jahre, Wortlaut unsicher"),
    33: ("B", "800 Jahre, Weinort, Ortsteil von Brackenheim"),
    34: ("B", "800 Jahre, Ortsteil von Ostfildern"),
    35: ("C", "800 Jahre, Weiler von Esslingen"),
    36: ("B", "800 Jahre, 1 000 Einwohner"),
    37: ("B", "800 Jahre, Esslinger Stadtteil"),
    38: ("B", "800 Jahre, Ortsteil von Gomaringen (Jahr nur aus Wikidata)"),
    39: ("B", "800 Jahre, Ortsteil von Jettingen"),
    40: ("C", "800 Jahre, Weiler"),
    41: ("C", "750 Jahre, Burgruine/Weiler – Jahr nur aus Wikidata"),
    42: ("C", "750 Jahre, kleiner Esslinger Stadtteil"),
    43: ("C", "700 Jahre, Weiler"),
    44: ("C", "700 Jahre, Erwähnung umstritten"),
    45: ("C", "700 Jahre, Esslinger Wohngebiet"),
    46: ("C", "700 Jahre, kleiner Ortsteil"),
    47: ("C", "Mühle/Hofgut"),
    48: ("C", "Hof"),
    49: ("B", "1250 Jahre, aber nur „vermutlich“ – zusammen mit Vaihingen/Markgröningen denkbar"),
    50: ("B", "1250 Jahre, Datierung 779–802 unsicher"),
    51: ("A", "700 Jahre Lateinschule – außergewöhnlich; Festschrift sehr wahrscheinlich"),
    52: ("A", "100 Jahre, Familienunternehmen, Pforzheim – Verlagsgebiet"),
    53: ("B", "100 Jahre Kino, Karlsruher Kultur – kleines Budget"),
    54: ("B", "50 Jahre Gemeinde (Gebietsreform) – Hessen"),
    55: ("C", "unklar, ob es die Firma noch gibt"),
    56: ("X", "Gründung des französischen Mutterkonzerns"),
    57: ("C", "800 Jahre, Erwähnung nur „möglicherweise“"),
    58: ("C", "600 Jahre, Weiler, Jahr unsicher"),
    59: ("B", "200 Jahre Universität – groß, macht es vermutlich mit eigenen Historikern"),
    60: ("C", "Teil der Universität"),
    61: ("X", "ehemalige Schule"),
    62: ("S", "175 Jahre, Familienunternehmen, Weltmarktführer Metallschläuche, Pforzheim"),
    63: ("S", "150 Jahre, Familienunternehmen, Karlsruhe – 20 km vom Verlag"),
    64: ("B", "150 Jahre, gehört meines Wissens zur Continentale-Gruppe (aus dem Gedächtnis)"),
    65: ("B", "50 Jahre Gemeinde (Gebietsreform) – Hessen"),
    66: ("B", "200 Jahre Sparkasse – Hessen, weit weg"),
    67: ("X", "2023 insolvent, an Zeitfracht (aus dem Gedächtnis)"),
    68: ("C", "vermutlich nicht mehr eigenständig – prüfen"),
    69: ("X", "historischer Postkartenverlag"),
    70: ("C", "unbekannt – prüfen"),
    71: ("C", "Tochter eines schwedischen Konzerns"),
    72: ("A", "100 Jahre, Familienbetrieb mit viel Geschichte"),
    73: ("X", "Mediengruppe – macht es selbst"),
    74: ("B", "400 Jahre Schule – Hessen"),
    75: ("B", "400 Jahre Gymnasium (Nachfolger des Pädagogs) – Hessen"),
    76: ("B", "200 Jahre Schule – Hessen"),
    77: ("A", "150 Jahre Privatbank, traditionsbewusst, Freiburg"),
    78: ("B", "150 Jahre Genossenschaftsbank, vorher Fusion prüfen"),
    79: ("B", "150 Jahre Genossenschaftsbank, klein – Fusion prüfen"),
    80: ("B", "125 Jahre Volksbank, Kreis Heilbronn"),
    81: ("C", "Hessen, Schule"),
    82: ("B", "175 Jahre Berufsschule, Heilbronn"),
    83: ("B", "150 Jahre Gymnasium, Stuttgart"),
    84: ("B", "125 Jahre Gymnasium, Freiburg"),
    85: ("C", "125 Jahre Grundschule"),
    86: ("B", "100 Jahre Kraftwerksgesellschaft – Festschrift plausibel"),
    87: ("C", "gehört zu einem US-Konzern (aus dem Gedächtnis)"),
    88: ("C", "Mühle, Jahr unsicher"),
    89: ("B", "200 Jahre – Hessen"),
    90: ("X", "in AXA aufgegangen (aus dem Gedächtnis)"),
    91: ("B", "100 Jahre Energieversorger – Hessen"),
    92: ("B", "100 Jahre, Familienunternehmen – Hessen"),
    93: ("C", "Musikgeschäft – Koblenz"),
    94: ("C", "Hessen"),
    95: ("C", "unklar – Mainz"),
    96: ("C", "Fertighaus, Eifel"),
    97: ("B", "100 Jahre Stadtwerke – Hessen"),
    98: ("C", "Autohaus – Hessen"),
    99: ("C", "Hessen"),
    100: ("C", "50 Jahre Hochschule, staatlich"),
    101: ("C", "50 Jahre Hochschule"),
    102: ("B", "175 Jahre Berufsschule, Heilbronn"),
    103: ("B", "175 Jahre Berufsschule, Stuttgart"),
    104: ("C", "Hessen"),
    105: ("A", "175 Jahre, Familienunternehmen (Nähgarne), Bönnigheim"),
    106: ("A", "175 Jahre Kranbau, Heilbronn"),
    107: ("B", "150 Jahre, gehört einem US-Konzern (aus dem Gedächtnis)"),
    108: ("A", "150 Jahre, börsennotiert – großes Firmenbuch, Wettbewerb mit Agenturen"),
    109: ("B", "150 Jahre Brauerei – Gründungsjahr prüfen (Wikidata 1879)"),
    110: ("A", "150 Jahre, Familienunternehmen (Düsen), Metzingen"),
    111: ("C", "unbekannt"),
    112: ("B", "125 Jahre Sägewerk – Eigentümerwechsel prüfen"),
    113: ("C", "Jazzkeller, 75 Jahre"),
    114: ("C", "Wanderhütte"),
    115: ("B", "75 Jahre, Werkzeughersteller, Kreis Freudenstadt"),
    116: ("C", "75 Jahre Versicherung"),
    117: ("B", "250 Jahre Gymnasium – Hessen"),
    118: ("B", "200 Jahre – Hessen"),
    119: ("C", "Hessen, Schule"),
    120: ("B", "150 Jahre Gymnasium – Trier"),
    121: ("C", "Hessen, Schule"),
    122: ("C", "Hessen, Schule"),
    123: ("C", "Hessen, Schule"),
    124: ("B", "100 Jahre, Pfalz"),
    125: ("C", "Tennisclub"),
    126: ("C", "25 Jahre"),
    127: ("B", "50 Jahre, Softwarefirma Karlsruhe"),
    128: ("B", "50 Jahre, börsennotiert, Enzkreis"),
    129: ("C", "50 Jahre, klein"),
    130: ("C", "75 Jahre, klein"),
    131: ("B", "125 Jahre Gymnasium – Jahr uneinheitlich"),
    132: ("C", "Dorfverein"),
    133: ("C", "Dorfverein"),
    134: ("C", "Dorfverein"),
    135: ("C", "25 Jahre"),
    136: ("C", "25 Jahre"),
    137: ("A", "500 Jahre Schützenverein – außergewöhnliche Zahl, Ortenau"),
    138: ("B", "175 Jahre Brauerei – Hessen"),
    139: ("C", "Hessen"),
    140: ("C", "Hessen"),
    141: ("C", "Hessen"),
    142: ("C", "Hessen"),
}
# Ab Rang 143 gibt es, mit wenigen Ausnahmen, nur noch kleine Jubiläen
# (25/50/75 Jahre), Narrenzünfte, Hessen/Saarland und Verlage. Die Ausnahmen
# stehen einzeln da, der Rest bekommt C bzw. X mit eigenem Grund.
E.update({
    155: ("A", "175 Jahre Kunstakademie – Festschrift so gut wie sicher, Land zahlt"),
    156: ("B", "150 Jahre Linde – Weltkonzern, macht es selbst; nur der Vollständigkeit halber"),
    157: ("B", "125 Jahre, Karlsbad – Verlagsgebiet"),
    158: ("B", "125 Jahre, Rastatt – Verlagsgebiet"),
    159: ("B", "125 Jahre Fastnachtsgesellschaft, Karlsruhe"),
    160: ("B", "125 Jahre, Pfalz"),
    170: ("B", "175 Jahre Bank – Hessen, Jahr uneinheitlich"),
    172: ("B", "75 Jahre, Familienunternehmen, Kreis Lörrach"),
    173: ("B", "Liebherr – Konzern macht es selbst"),
    174: ("C", "Liebherr – Konzern macht es selbst"),
    175: ("B", "Liebherr – Konzern macht es selbst"),
    177: ("B", "100 Jahre, Fußball, Kaiserstuhl"),
    180: ("A", "100 Jahre Geschichtsverein – natürlicher Partner eines Regionalverlags"),
    192: ("B", "50 Jahre, groß"),
    193: ("B", "50 Jahre, Reisemobile"),
    220: ("B", "200 Jahre – Frankfurt"),
    221: ("B", "200 Jahre Naturkundeverein – Wiesbaden"),
    222: ("B", "200 Jahre Städelschule – Frankfurt"),
    238: ("X", "2023 insolvent, an Zeitfracht (aus dem Gedächtnis)"),
    239: ("B", "150 Jahre Alpenvereinssektion – Festschrift üblich"),
    240: ("B", "125 Jahre Fußball"),
    243: ("B", "125 Jahre Fußball"),
    246: ("A", "125 Jahre SC Freiburg – Chronik sicher, ob mit Regionalverlag offen"),
    247: ("B", "125 Jahre Alpenvereinssektion – Festschrift üblich"),
    261: ("B", "150 Jahre Fastnacht, Buchen – Verlagsgebiet"),
    280: ("X", "Verlag"),
    299: ("B", "75 Jahre Kammer"),
    302: ("X", "Verlag – macht es selbst"),
    303: ("B", "50 Jahre Geschichtsverein – Partner, aber kleine Zahl"),
    323: ("X", "Verlag"),
    338: ("X", "Verlag"),
})
X_NAMEN = {"Fachwerkhaus Weilburger Straße 14": "Gebäude", "Q108597725": "Datensatz ohne Namen",
           "Wirtschaftsbetrieb Mainz": "Behördenbetrieb", "Porsche Zentrum Erfurt": "Autohaus, Datenfehler",
           "ThyssenKrupp System Engineering": "mehrfach umbenannt, Konzern", "Family Media": "Joint Venture beendet",
           "Edition Gorz": "Verlag", "Sanofi (Germany)": "Konzerntochter", "Blinkx": "Internetdienst",
           "Dagstuhl Research Online Publication Server": "Server", "Hinnerbäcker": "Bäckerei, Hessen"}


def einschaetzung(z):
    rang = int(z["Rang"])
    if rang in E:
        return E[rang]
    if z["Name"] in X_NAMEN:
        return "X", X_NAMEN[z["Name"]]
    jahre = int(z["Jahre"])
    name = z["Name"].lower()
    # Hier entscheide ich nach Art – jede dieser Zeilen habe ich durchgesehen,
    # keine davon ist ein Kandidat für ein Buch.
    if any(w in name for w in ("narr", "zunft", "hexen", "geister", "teufel", "dämone", "fastnacht", "wölf")):
        return "C", "Narrenzunft – Festheft, kein Buch"
    if z["Art"] == "Verein":
        return "C", f"{jahre} Jahre, kleiner Verein"
    return "C", f"{jahre} Jahre – zu klein für ein Jubiläumsbuch"


# Lange Kommentare für die Präsentationsliste (S und A im Verlagsgebiet)
KOMMENTAR = {
    "Markgröningen": "1250 Jahre seit der Schenkung an Fulda 779. Stadt mit 14 000 Einwohnern und eigenem Haushalt. "
                     "Sie hat 1979 die 1200 Jahre gefeiert, das Jubiläum hat dort also Tradition. Ansprechpartner: "
                     "Bürgermeisteramt bzw. Stadtarchiv. Ein Ortsbuch ist hier sehr wahrscheinlich – die Frage ist "
                     "nur, wer es macht.",
    "Adelsheim": "1250 Jahre (779, Kloster Fulda), kleine Stadt mit 5 100 Einwohnern im Neckar-Odenwald-Kreis – "
                 "mitten im Verlagsgebiet. Erholungsort, also auch Interesse am Stadtmarketing. Klassischer Kunde "
                 "für eine Stadtgeschichte.",
    "Bad Wimpfen": "1200 Jahre (829 „Wimpina“). Kaiserpfalz, Altstadt, Tourismus – eine Stadt, die ihre Geschichte "
                   "verkauft. Ein Jubiläumsbuch passt in das Stadtmarketing und lässt sich danach im Tourismus "
                   "weiterverkaufen.",
    "Großbottwar": "1250 Jahre (Fulda 779, „Boteburon“). Stadt mit 8 000 Einwohnern, Weinort im Bottwartal. Gleiche "
                   "Urkunde wie Markgröningen und Vaihingen – mehrere Kunden mit demselben Anlass.",
    "Waldsee": "800 Jahre (1229). Gemeinde mit 6 000 Einwohnern im Rhein-Pfalz-Kreis, nördlich von Speyer. "
               "Kurzer Weg, der Verlag ist in der Gegend bekannt.",
    "Dürrmenz": "1250 Jahre (Lorscher Codex 779). Dürrmenz ist der älteste Teil von Mühlacker und war lange der "
                "wichtigere Ort; Mühlacker selbst wird erst 1292 erwähnt. Der eigentliche Kunde ist also die Stadt "
                "Mühlacker, die größte Stadt im Enzkreis. Verlagsgebiet.",
    "Witzenmann-Gruppe": "175 Jahre, Familienunternehmen in Pforzheim, gilt als Erfinder des Metallschlauchs (aus dem "
                         "Gedächtnis). Ein großes Firmenjubiläum mit Geld dahinter, im Verlagsgebiet. Konkurrenz durch "
                         "Agenturen für Firmengeschichte möglich, der regionale Bezug spricht für den Verlag.",
    "Carl Roth": "150 Jahre, Familienunternehmen für Laborbedarf in Karlsruhe, rund 20 km vom Verlag. Familien"
                 "unternehmen mit langer Geschichte bestellen zum 150. gern ein Buch.",
    "Benningen am Neckar": "1250 Jahre (779 „Bunninga“), 6 300 Einwohner, Kreis Ludwigsburg. Kleinere Gemeinde, "
                           "aber starke Zahl.",
    "Bermatingen": "1250 Jahre (779, Kloster St. Gallen), 4 100 Einwohner. Liegt am Bodensee, also weit weg, aber "
                   "die Zahl ist stark.",
    "Ingersheim": "1250 Jahre (Lorscher Codex 779), 6 300 Einwohner, Kreis Ludwigsburg.",
    "Vaihingen an der Enz": "Mit 28 500 Einwohnern die größte Stadt der Liste. Der Artikel sagt aber nur, die erste "
                            "Erwähnung „könnte“ 779 sein. Vorher klären, ob die Stadt 2029 feiert. Wenn ja, wird "
                            "daraus ein S-Lead.",
    "Denkendorf": "900 Jahre (1129, Papsturkunde für das Kloster Denkendorf). 11 000 Einwohner, Kreis Esslingen. "
                  "Ort und Kloster feiern dasselbe Jubiläum.",
    "Oberboihingen": "800 Jahre (1229), 5 600 Einwohner, Kreis Esslingen.",
    "Weilheim": "1100 Jahre (929), kleine Gemeinde am Hochrhein (3 200 Einwohner). Die Ortsteile sind noch älter "
                "(Bannholz 871, Aisperg 890) – Stoff für ein dickes Buch.",
    "Wimsheim": "800 Jahre (1229, Kloster Bebenhausen), 2 900 Einwohner, Enzkreis. Kleine Gemeinde im "
                "Verlagsgebiet; Budget prüfen.",
    "Großsachsen": "1250 Jahre (Lorscher Codex 779). Die Urkunde nennt „Sahsenheim“ für Groß-, Lützel- und "
                   "Hohensachsen. Großsachsen gehört zu Hirschberg, die beiden anderen zu Weinheim (aus dem "
                   "Gedächtnis). Denkbar ist ein gemeinsames Fest von zwei Kommunen. Rhein-Neckar-Kreis, "
                   "Verlagsgebiet.",
    "Flözlingen": "1250 Jahre (779, St. Gallen), ältester Ortsteil von Zimmern ob Rottweil. Typischer Fall für eine "
                  "Ortschronik mit Heimatverein und Gemeinde als Geldgeber.",
    "Ohrnberg": "1250 Jahre (779), Ortsteil von Öhringen. Gleiches Muster wie Flözlingen.",
    "Rupert-Neß-Gymnasium Wangen": "700 Jahre als Lateinschule (1329) – ein außergewöhnliches Schuljubiläum. Eine "
                                   "Festschrift ist sehr wahrscheinlich. Im Allgäu weit weg, aber eine solche Zahl "
                                   "kommt selten vor.",
    "Bader Versand": "100 Jahre, Familienunternehmen (Versandhandel) in Pforzheim, Verlagsgebiet.",
    "Erlebnispark Tripsdrill": "100 Jahre, Familienbetrieb in Cleebronn, gilt als ältester Freizeitpark Deutschlands "
                               "(aus dem Gedächtnis). Bildreiche Geschichte – gut für ein schönes Buch.",
    "Bankhaus E. Mayer": "150 Jahre Privatbank in Freiburg. Privatbanken pflegen ihre Tradition, ein "
                         "Jubiläumsbuch ist typisch.",
    "Amann & Söhne": "175 Jahre, Familienunternehmen (Nähgarne) in Bönnigheim.",
    "Wolffkran": "175 Jahre Kranbau, Heilbronn. Bekannte Marke mit langer Geschichte.",
    "ElringKlinger": "150 Jahre, börsennotierter Autozulieferer mit über 6 000 Mitarbeitern. Geld ist da, aber "
                     "vermutlich auch eine eigene Kommunikationsabteilung oder Agentur.",
    "Lechler (Unternehmen)": "150 Jahre, Familienunternehmen (Düsentechnik) in Metzingen.",
    "Staatliche Akademie der Bildenden Künste Karlsruhe": "175 Jahre Kunstakademie in Karlsruhe. Eine Festschrift ist "
                                                          "so gut wie sicher, das Land zahlt. Die Akademie gestaltet "
                                                          "vielleicht selbst – dann als Verlag dazukommen.",
    "Schützenverein Oberkirch": "500 Jahre – eine sehr seltene Zahl für einen Verein. Ortenaukreis.",
    "SC Freiburg": "125 Jahre, Bundesligaverein. Eine Chronik kommt sicher, Geld ist da. Ob ein Regionalverlag zum "
                   "Zug kommt, ist offen – einen Versuch wert.",
    "Geschichtsverein Markgräflerland": "100 Jahre Geschichtsverein – das ist die Kernkundschaft eines "
                                        "Regionalverlags. Ein Festband oder Jahrbuch zum Jubiläum liegt nahe.",
}


PRIORITAET = [
    # S
    "Markgröningen", "Adelsheim", "Bad Wimpfen", "Dürrmenz", "Großbottwar", "Waldsee",
    "Witzenmann-Gruppe", "Carl Roth",
    # A – erst Orte (Kerngeschäft), dann Einrichtungen, dann Firmen, dann Vereine
    "Vaihingen an der Enz", "Denkendorf", "Großsachsen", "Benningen am Neckar", "Ingersheim",
    "Weilheim", "Oberboihingen", "Wimsheim", "Flözlingen", "Ohrnberg", "Bermatingen",
    "Staatliche Akademie der Bildenden Künste Karlsruhe", "Geschichtsverein Markgräflerland",
    "Rupert-Neß-Gymnasium Wangen", "Erlebnispark Tripsdrill", "Bader Versand", "Bankhaus E. Mayer",
    "Amann & Söhne", "Lechler (Unternehmen)", "Wolffkran", "ElringKlinger",
    "Schützenverein Oberkirch", "SC Freiburg",
]

# Kommentare für Stufe B (Verlagsgebiet), in der Reihenfolge der Präsentationsliste
KOMMENTAR_B = {
    # Orte
    "Hochdorf/Enz": "1250 Jahre – aber nur „vermutlich“ 779, in derselben Fulda-Urkunde wie Markgröningen und "
                    "Vaihingen. Wenn die Nachbarn feiern, feiert Hochdorf (Gemeinde Eberdingen) vermutlich mit.",
    "Siglingen": "1250 Jahre, aber die Urkunde ist nur auf 779–802 datiert. Ortsteil von Neudenau. Erst nachfragen, "
                 "ob man 2029 feiern will.",
    "Bockschaft": "1200 Jahre (Lorscher Codex 829), war aber ursprünglich nur ein Hofgut. Kleiner Ortsteil von "
                  "Kirchardt – die Gemeinde müsste das Buch tragen.",
    "Eberfingen": "1100 Jahre (929, St. Gallen) – starke Zahl, aber ein kleiner Ortsteil von Stühlingen ganz im "
                  "Süden. Ortschronik denkbar, weit weg vom Verlag.",
    "Stuttgart-Wangen": "",
    "Wangen": "800 Jahre, Stuttgarter Stadtbezirk mit 9 100 Einwohnern. Wangen wurde 1229 „zusammen mit "
              "Stuttgart“ erwähnt. Ein Bezirksbuch ist möglich, Geld kommt eher vom Bürgerverein als von der Stadt.",
    "Cleebronn": "750 Jahre (1279), Gemeinde mit 3 000 Einwohnern. Dort liegt auch Tripsdrill (A-Lead, 100 Jahre) – "
                 "zwei Jubiläen im selben Ort, das ließe sich verbinden.",
    "Pfaffenhofen": "750 Jahre (1279), Gemeinde mit 2 600 Einwohnern im Zabergäu. Kleiner Haushalt.",
    "Schomburg": "800 Jahre (1229), Ortsteil von Wangen im Allgäu mit 2 500 Einwohnern. Weit weg.",
    "Unterstadion": "900 Jahre (1129), aber nur 860 Einwohner. Eher eine Festschrift als ein Buch.",
    "Lauterach": "800 Jahre (1229), nur 600 Einwohner. Die Teilorte sind älter (Talheim 776) – inhaltlich reizvoll, "
                 "das Budget dürfte klein sein.",
    "Schönau": "900 Jahre (1129), Pfalz, Südwestpfalz – aber nur 380 Einwohner. Kleines Festheft eher als Buch.",
    "Zell": "800 Jahre (1229), Zell am Neckar, heute ein Stadtteil von Esslingen. Mit Mettingen und Rüdern feiern "
            "2029 gleich mehrere Esslinger Stadtteile – vielleicht ein Sammelband für die Stadt Esslingen.",
    "Mettingen (Esslingen am Neckar)": "800 Jahre (1229), Esslinger Stadtteil. Siehe Zell – zusammen mit der Stadt "
                                       "Esslingen ansprechen.",
    "Kemnat": "800 Jahre (1229), Ortsteil von Ostfildern.",
    "Haberschlacht": "800 Jahre (1229), Weinort, Ortsteil von Brackenheim. Brackenheim hat mit der VBU Volksbank "
                     "(125, B) noch einen Anlass.",
    "Unterjettingen": "800 Jahre (1229), Ortsteil von Jettingen, Kreis Böblingen.",
    "Leupolz": "800 Jahre, Ortsteil von Wangen im Allgäu, 1 000 Einwohner. Jahr nur aus Wikidata.",
    "Stockach (Gomaringen)": "800 Jahre, Ortsteil von Gomaringen. Jahr nur aus Wikidata, ohne Beleg – prüfen.",
    "Wart": "700 Jahre (1329), Ortsteil von Altensteig, Kreis Calw – liegt im Verlagsgebiet, ist aber klein.",
    # Einrichtungen
    "Universität Stuttgart": "200 Jahre. Groß und mit Geld, macht die Festschrift aber vermutlich mit eigenen "
                             "Historikern und einem Wissenschaftsverlag. Nur mit gutem Kontakt.",
    "Elly-Heuss-Knapp-Gymnasium Stuttgart": "150 Jahre (offizielles Gründungsjahr 1879). Schulen machen fast immer "
                                            "eine Festschrift – kleines Budget, aber verlässlich.",
    "Gustav-von-Schmoller-Schule Heilbronn": "175 Jahre Berufsschule. Festschrift wahrscheinlich, Schulträger ist "
                                             "der Landkreis bzw. die Stadt.",
    "Johann-Jakob-Widmann-Schule": "175 Jahre Berufsschule in Heilbronn. Wie Schmoller-Schule – beide zusammen "
                                   "ansprechen.",
    "Werner-Siemens-Schule Stuttgart": "175 Jahre Berufsschule.",
    "Friedrich-Gymnasium Freiburg": "125 Jahre Gymnasium.",
    "Parler-Gymnasium Schwäbisch Gmünd": "125 Jahre, aber das Gründungsjahr ist uneinheitlich – vorher klären.",
    "Landesärztekammer Baden-Württemberg": "75 Jahre Zusammenschluss der vier Kammern. Institution mit Geld, aber "
                                           "kleine Zahl – eher eine Broschüre.",
    "Geschichtsverein der Diözese Rottenburg-Stuttgart": "50 Jahre. Geschichtsvereine sind Kernkundschaft, die Zahl "
                                                         "ist klein – als Kontakt für später pflegen.",
    # Banken
    "VBU Volksbank im Unterland": "125 Jahre Volksbank in Brackenheim. Genossenschaftsbanken machen gern "
                                  "Festschriften, oft mit Ortsgeschichte. Fusionen prüfen.",
    "Raiffeisenbank Kaiserstuhl": "150 Jahre. Vorher prüfen, ob sie noch eigenständig ist.",
    "Volksbank Möckmühl": "150 Jahre, kleine Bank. Vorher prüfen, ob sie noch eigenständig ist.",
    "Mannheimer Versicherungen": "150 Jahre. Gehört meines Wissens zur Continentale-Gruppe (aus dem Gedächtnis) – "
                                 "ob die Marke noch selbst feiert, ist fraglich.",
    # Firmen
    "Heinrich Dinkelacker": "150 Jahre laut Wikidata (1879) – das Gründungsjahr der Brauerei vorher prüfen.",
    "Klenk Holz": "125 Jahre Sägewerk in Oberrot. Eigentümerwechsel prüfen.",
    "Coperion": "150 Jahre Maschinenbau in Stuttgart. Gehört einem US-Konzern (aus dem Gedächtnis) – Entscheidung "
                "eher in der Konzernzentrale.",
    "Rheinkraftwerk Albbruck-Dogern AG": "100 Jahre Kraftwerksgesellschaft am Hochrhein. Energieversorger machen zu "
                                         "runden Jubiläen gern eine Festschrift.",
    "Schauburg": "100 Jahre Kino in Karlsruhe, eines der ältesten der Stadt. Schöne Geschichte, aber kleines "
                 "Budget – eher ein Bildband auf eigenes Risiko.",
    "Leuco": "75 Jahre, Werkzeughersteller in Horb (Kreis Freudenstadt, Verlagsgebiet). Kleine Zahl.",
    "Glatt Unternehmensgruppe": "75 Jahre, Familienunternehmen im Kreis Lörrach. Kleine Zahl.",
    "Liebherr-Werk Biberach": "75 Jahre. Liebherr macht das als Konzern selbst – nur der Vollständigkeit halber.",
    "Liebherr-Hausgeräte": "75 Jahre. Wie Liebherr-Werk Biberach.",
    "PTV": "50 Jahre, Softwarefirma in Karlsruhe (Verkehrsplanung). Junge Firma, eher eine Broschüre.",
    "STRATEC": "50 Jahre, börsennotiert, Enzkreis. Kleine Zahl.",
    "CHG-Meridian": "50 Jahre, großer Leasinganbieter in Weingarten. Kleine Zahl, aber Geld ist da.",
    "Carthago Reisemobilbau": "50 Jahre, Reisemobile, Aulendorf. Kleine Zahl.",
    # Vereine
    "Sektion Ulm": "150 Jahre Alpenvereinssektion. Sektionen machen traditionell eine Festschrift.",
    "Sektion Stuttgart": "125 Jahre Alpenvereinssektion, eine der großen Sektionen.",
    "FG Narrhalla": "150 Jahre Fastnachtsgesellschaft in Buchen – Verlagsgebiet, für eine Narrhalla eine große Zahl.",
    "ATSV Mutschelbach": "125 Jahre, Karlsbad (Landkreis Karlsruhe) – direkt vor der Haustür. Verein, kleines Budget.",
    "FC Rastatt 04": "125 Jahre, Rastatt – Verlagsgebiet. Verein, kleines Budget.",
    "FG Rüppurr": "125 Jahre Fastnachtsgesellschaft in Karlsruhe.",
    "TuS Hohenecken": "125 Jahre, Kaiserslautern (Pfalz).",
    "SVN Zweibrücken": "100 Jahre, Zweibrücken (Pfalz).",
    "Bahlinger SC": "100 Jahre, bekannter Amateurverein am Kaiserstuhl.",
    "1. FC Normannia Gmünd": "125 Jahre Fußball, Schwäbisch Gmünd.",
    "FC Singen 04": "125 Jahre Fußball, Singen.",
}


def lies_kontakte():
    """Kontaktdaten aus kontakte/ergebnis_*.jsonl (von Agenten aus Impressum/Kontaktseiten gesammelt)."""
    import json
    kontakte = {}
    for pfad in sorted((ORDNER / "kontakte").glob("ergebnis_*.jsonl")):
        for zeile in pfad.read_text(encoding="utf-8").splitlines():
            if zeile.strip():
                k = json.loads(zeile)
                kontakte[k["qid"]] = k
    # Gegenprobe: steht die Adresse wirklich auf der Fundstelle? (kontakte/pruefung.json)
    pruef = ORDNER / "kontakte" / "pruefung.json"
    if pruef.exists():
        for qid, status in json.loads(pruef.read_text(encoding="utf-8")).items():
            if qid in kontakte:
                kontakte[qid]["geprueft"] = status
    return kontakte


def domain_passt(email, *websites):
    """Grobe Plausibilität: gehört die Mail-Domain zur Website? Sonst von Hand gegenlesen."""
    import re
    if not email or "@" not in email:
        return True
    maildomain = email.rsplit("@", 1)[1].lower()
    for w in websites:
        host = re.sub(r"^https?://(www\.)?", "", (w or "").lower()).split("/")[0]
        if host and (maildomain.endswith(host) or host.endswith(maildomain)
                     or maildomain.split(".")[-2:] == host.split(".")[-2:]):
            return True
    return False


def main():
    zeilen = list(csv.DictReader(open(ORDNER / "empfehlungen_2029.csv", encoding="utf-8-sig"), delimiter=";"))
    alle = []
    for z in zeilen:
        stufe, kurz = einschaetzung(z)
        gebiet = z["Bundesland"] == "Baden-Württemberg" or z["Landkreis"] in PFALZ
        alle.append({**z, "Stufe": stufe, "Einschätzung": kurz, "Verlagsgebiet": "ja" if gebiet else "nein"})

    reihenfolge = {"S": 0, "A": 1, "B": 2, "C": 3, "X": 4}
    alle.sort(key=lambda z: (reihenfolge[z["Stufe"]], int(z["Rang"])))
    spalten = ["Stufe", "Einschätzung", "Name", "Art", "Jubiläum", "Jahre", "seit", "Einwohner", "Gemeinde",
               "Landkreis", "Bundesland", "Verlagsgebiet", "Quelle", "Beleg", "Website", "Wikipedia", "QID"]
    with open(ORDNER / "leads_2029.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=spalten, delimiter=";", extrasaction="ignore")
        w.writeheader()
        w.writerows(alle)

    praesentation = [z for z in alle if z["Stufe"] in ("S", "A", "B") and z["Verlagsgebiet"] == "ja"]
    kommentare = {**KOMMENTAR, **{k: v for k, v in KOMMENTAR_B.items() if v}}
    prioritaet = PRIORITAET + [k for k, v in KOMMENTAR_B.items() if v]
    # Reihenfolge innerhalb der Stufe: meine Priorität, nicht die alte Punkteliste
    praesentation.sort(key=lambda z: (reihenfolge[z["Stufe"]], prioritaet.index(z["Name"])
                                      if z["Name"] in prioritaet else 999))
    fehlt = [z["Name"] for z in praesentation if z["Name"] not in kommentare]
    if fehlt:
        raise SystemExit(f"Kommentar fehlt für: {fehlt}")
    with open(ORDNER / "leads_2029_S_A_B_verlagsgebiet.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        kontakte = lies_kontakte()
        w.writerow(["Nr", "Stufe", "Name", "Art", "Jubiläum 2029", "Ort", "Landkreis", "Einwohner", "Kommentar",
                    "E-Mail", "E-Mail geprüft", "Telefon", "Anschrift", "Ansprechpartner", "Kontakt-Quelle", "Kontakt-Hinweis",
                    "Beleg", "Wikipedia"])
        for nr, z in enumerate(praesentation, 1):
            ort = z["Gemeinde"] if z["Gemeinde"] != z["Name"] else ""
            k = kontakte.get(z["QID"], {})
            hinweis = k.get("bemerkung") or ""
            if k.get("email") and not domain_passt(k["email"], z["Website"], k.get("quelle")):
                hinweis = "; ".join(x for x in ("Mail-Domain passt nicht zur Website – gegenlesen", hinweis) if x)
            w.writerow([nr, z["Stufe"], z["Name"], z["Art"], f"{z['Jahre']} Jahre (seit {z['seit']})", ort,
                        z["Landkreis"], z["Einwohner"], kommentare[z["Name"]], k.get("email") or "", k.get("geprueft") or "",
                        k.get("telefon") or "", k.get("anschrift") or "", k.get("ansprechpartner") or "",
                        k.get("quelle") or "", hinweis, z["Beleg"], z["Wikipedia"]])
    from collections import Counter
    print("Stufen gesamt:", dict(sorted(Counter(z["Stufe"] for z in alle).items(), key=lambda t: reihenfolge[t[0]])))
    print("Präsentationsliste:", len(praesentation), "Einträge,",
          dict(Counter(z["Stufe"] for z in praesentation)))


if __name__ == "__main__":
    main()
