"""Durchsicht der Jubiläen 2029 von Hand (Oktober/November 2026).

Stufen siehe leads.py. Kommentare mit „(aus dem Gedächtnis)“ vor einer Anfrage prüfen.
"""

E = {
    'Q355435': ('S', 'Stadt, 1250 Jahre, Beleg sauber (Fulda 779), mitten im Verlagsgebiet'),  # Adelsheim
    'Q13896': ('S', 'Stadt mit 14 000 Einwohnern, 1250 Jahre; hat 1979 schon die 1200 gefeiert – feiert also'),  # Markgröningen
    'Q23023': ('S', '800 Jahre, 6 000 Einwohner, Rhein-Pfalz-Kreis direkt vor der Haustür'),  # Waldsee
    'Q634084': ('S', '1200 Jahre, 12 700 Einwohner, Königspfalz Tribur – große Geschichte, große Gemeinde'),  # Trebur
    'Q61830': ('S', 'Stadt, 1250 Jahre, 8 000 Einwohner, Beleg sauber'),  # Großbottwar
    'Q491535': ('S', 'historische Stadt, 1200 Jahre, Tourismusort – Buch passt zum Stadtmarketing'),  # Bad Wimpfen
    'Q74763163': ('S', '1250 Jahre Dürrmenz, ältester Teil von Mühlacker – die Stadt dürfte feiern'),  # Dürrmenz
    'Q2586953': ('S', '175 Jahre, Familienunternehmen, Weltmarktführer Metallschläuche, Pforzheim'),  # Witzenmann-Gruppe
    'Q17122614': ('S', '150 Jahre, Familienunternehmen, Karlsruhe – 20 km vom Verlag'),  # Carl Roth
    'Q80511': ('A', '800 Jahre, kleine Gemeinde (2 900), Enzkreis – Verlagsgebiet'),  # Wimsheim
    'Q61653': ('A', '1250 Jahre, 6 300 Einwohner'),  # Benningen am Neckar
    'Q507274': ('A', '1250 Jahre (St. Gallen 779), Bodensee – weit weg, aber starke Zahl'),  # Bermatingen
    'Q61705': ('A', '1250 Jahre (Lorscher Codex), 6 300 Einwohner'),  # Ingersheim
    'Q515323': ('A', '1100 Jahre, kleine Gemeinde am Hochrhein; Ortsteile noch älter'),  # Weilheim
    'Q80809': ('A', '900 Jahre Ort und Kloster Denkendorf (1129), 11 000 Einwohner'),  # Denkendorf
    'Q61928': ('A', 'größte Stadt der Liste (28 500), aber Ersterwähnung 779 nur „könnte“ – vorher klären'),  # Vaihingen an der Enz
    'Q80606': ('A', '800 Jahre, 5 600 Einwohner'),  # Oberboihingen
    'Q43273976': ('A', '1250 Jahre; Urkunde gilt für alle drei Sachsen-Orte (Hirschberg und Weinheim)'),  # Großsachsen
    'Q20179999': ('A', '1250 Jahre, ältester Ortsteil von Zimmern ob Rottweil – klassische Ortschronik'),  # Flözlingen
    'Q2016767': ('A', '1250 Jahre, Ortsteil von Öhringen'),  # Ohrnberg
    'Q16006232': ('A', '700 Jahre Lateinschule – außergewöhnlich; Festschrift sehr wahrscheinlich'),  # Rupert-Neß-Gymnasium Wangen
    'Q798951': ('A', '100 Jahre, Familienunternehmen, Pforzheim – Verlagsgebiet'),  # Bader Versand
    'Q1998041': ('A', '900 Jahre; hat 1979 schon die 850 gefeiert – feiert also'),  # Nordheim
    'Q520675': ('A', '100 Jahre, Familienbetrieb mit viel Geschichte'),  # Erlebnispark Tripsdrill
    'Q1874094': ('A', '400 Jahre Gymnasium (Pädagog 1629) – Festschrift so gut wie sicher'),  # Ludwig-Georgs-Gymnasium
    'Q1746794': ('A', '150 Jahre Privatbank, traditionsbewusst, Freiburg'),  # Bankhaus E. Mayer
    'Q16002754': ('A', '175 Jahre, Familienunternehmen (Nähgarne), Bönnigheim'),  # Amann & Söhne
    'Q1485474': ('A', '175 Jahre Kranbau, Heilbronn'),  # Wolffkran
    'Q1333317': ('A', '150 Jahre, börsennotiert – großes Firmenbuch, Wettbewerb mit Agenturen'),  # ElringKlinger
    'Q1811343': ('A', '150 Jahre, Familienunternehmen (Düsen), Metzingen'),  # Lechler (Unternehmen)
    'Q19930480': ('A', '500 Jahre Schützenverein – außergewöhnliche Zahl, Ortenau'),  # Schützenverein Oberkirch
    'Q322938': ('A', '175 Jahre Kunstakademie – Festschrift so gut wie sicher, Land zahlt'),  # Staatliche Akademie der Bildenden Künste Karlsruhe
    'Q19963065': ('A', '100 Jahre Geschichtsverein – natürlicher Partner eines Regionalverlags'),  # Geschichtsverein Markgräflerland
    'Q106394': ('A', '125 Jahre SC Freiburg – Chronik sicher, ob mit Regionalverlag offen'),  # SC Freiburg
    'Q546393': ('B', '750 Jahre, 3 000 Einwohner'),  # Cleebronn
    'Q537465': ('B', '750 Jahre, 2 600 Einwohner'),  # Pfaffenhofen
    'Q373283': ('B', '900 Jahre, 380 Einwohner – kleines Festheft eher als Buch'),  # Schönau
    'Q540770': ('B', '900 Jahre, 860 Einwohner'),  # Unterstadion
    'Q632934': ('B', '800 Jahre, 600 Einwohner'),  # Lauterach
    'Q1308178': ('B', 'Zell am Neckar, heute Stadtteil von Esslingen, 800 Jahre'),  # Zell
    'Q2549872': ('B', '700 Jahre, Ortsteil von Altensteig, Kreis Calw'),  # Wart
    'Q889245': ('B', '1200 Jahre, aber ursprünglich Hofgut, kleiner Ortsteil'),  # Bockschaft
    'Q18019848': ('B', '1100 Jahre, kleiner Ortsteil von Stühlingen'),  # Eberfingen
    'Q19969752': ('B', '800 Jahre, 2 500 Einwohner, Ortsteil von Wangen im Allgäu'),  # Schomburg
    'Q1469016': ('B', '800 Jahre, 9 100 Einwohner – Stuttgarter Stadtbezirk, „zusammen mit Stuttgart“ 1229 erwähnt'),  # Wangen
    'Q1566803': ('B', '800 Jahre, Weinort, Ortsteil von Brackenheim'),  # Haberschlacht
    'Q1738511': ('B', '800 Jahre, Ortsteil von Ostfildern'),  # Kemnat
    'Q17577331': ('B', '800 Jahre, 1 000 Einwohner'),  # Leupolz
    'Q1925995': ('B', '800 Jahre, Esslinger Stadtteil'),  # Mettingen (Esslingen am Neckar)
    'Q2350567': ('B', '800 Jahre, Ortsteil von Gomaringen (Jahr nur aus Wikidata)'),  # Stockach (Gomaringen)
    'Q33510632': ('B', '800 Jahre, Ortsteil von Jettingen'),  # Unterjettingen
    'Q47796855': ('B', '1250 Jahre, aber nur „vermutlich“ – zusammen mit Vaihingen/Markgröningen denkbar'),  # Hochdorf/Enz
    'Q2284859': ('B', '1250 Jahre, Datierung 779–802 unsicher'),  # Siglingen
    'Q2231851': ('B', '100 Jahre Kino, Karlsruher Kultur – kleines Budget'),  # Schauburg
    'Q1798208': ('B', '700 Jahre, Ortsteil von Michelstadt (Odenwald)'),  # Steinbuch
    'Q2582423': ('B', '700 Jahre, Ortsteil von Bensheim (Bergstraße)'),  # Wilmshausen
    'Q634106': ('B', '50 Jahre Gemeinde (Gebietsreform) – Hessen'),  # Wettenberg
    'Q122453': ('B', '200 Jahre Universität – groß, macht es vermutlich mit eigenen Historikern'),  # Universität Stuttgart
    'Q1890715': ('B', '150 Jahre, gehört meines Wissens zur Continentale-Gruppe (aus dem Gedächtnis)'),  # Mannheimer Versicherungen
    'Q622428': ('B', '50 Jahre Gemeinde (Gebietsreform) – Hessen'),  # Lahnau
    'Q15128978': ('B', '200 Jahre Sparkasse – Hessen, weit weg'),  # Sparkasse Waldeck-Frankenberg
    'Q1665423': ('B', '200 Jahre Schule – Hessen'),  # Viktoriaschule
    'Q17477739': ('B', '150 Jahre Genossenschaftsbank, vorher Fusion prüfen'),  # Raiffeisenbank Kaiserstuhl
    'Q15134090': ('B', '150 Jahre Genossenschaftsbank, klein – Fusion prüfen'),  # Volksbank Möckmühl
    'Q2505471': ('B', '125 Jahre Volksbank, Kreis Heilbronn'),  # VBU Volksbank im Unterland
    'Q1555507': ('B', '175 Jahre Berufsschule, Heilbronn'),  # Gustav-von-Schmoller-Schule Heilbronn
    'Q89491135': ('B', '150 Jahre Gymnasium, Stuttgart'),  # Elly-Heuss-Knapp-Gymnasium Stuttgart
    'Q15688375': ('B', '125 Jahre Gymnasium, Freiburg'),  # Friedrich-Gymnasium Freiburg
    'Q125578831': ('B', '100 Jahre Kraftwerksgesellschaft – Festschrift plausibel'),  # Rheinkraftwerk Albbruck-Dogern AG
    'Q2109610': ('B', '200 Jahre – Hessen'),  # Priesterseminar Limburg
    'Q17486810': ('B', '100 Jahre Energieversorger – Hessen'),  # EAM
    'Q21036277': ('B', '100 Jahre, Familienunternehmen – Hessen'),  # HEWI
    'Q1380025': ('B', '100 Jahre Stadtwerke – Hessen'),  # Städtische Werke
    'Q133932437': ('B', '175 Jahre Berufsschule, Heilbronn'),  # Johann-Jakob-Widmann-Schule
    'Q134068397': ('B', '175 Jahre Berufsschule, Stuttgart'),  # Werner-Siemens-Schule Stuttgart
    'Q1131085': ('B', '150 Jahre, gehört einem US-Konzern (aus dem Gedächtnis)'),  # Coperion
    'Q1597044': ('B', '150 Jahre Brauerei – Gründungsjahr prüfen (Wikidata 1879)'),  # Heinrich Dinkelacker
    'Q1774105': ('B', '125 Jahre Sägewerk – Eigentümerwechsel prüfen'),  # Klenk Holz
    'Q18023330': ('B', '75 Jahre, Werkzeughersteller, Kreis Freudenstadt'),  # Leuco
    'Q1464247': ('B', '250 Jahre Gymnasium – Hessen'),  # Friedrichsgymnasium
    'Q30264093': ('B', '200 Jahre – Hessen'),  # Glasbau Hahn (Germany)
    'Q1284907': ('B', '175 Jahre Schule, Darmstadt'),  # Edith-Stein-Schule Darmstadt
    'Q766044': ('B', '150 Jahre Gymnasium – Trier'),  # Auguste-Viktoria-Gymnasium Trier
    'Q325249': ('B', '100 Jahre, Pfalz'),  # SVN Zweibrücken
    'Q2044808': ('B', '50 Jahre, Softwarefirma Karlsruhe'),  # PTV
    'Q22909101': ('B', '50 Jahre, börsennotiert, Enzkreis'),  # STRATEC
    'Q139863925': ('B', '125 Jahre Gymnasium – Jahr uneinheitlich'),  # Parler-Gymnasium Schwäbisch Gmünd
    'Q204880': ('B', '175 Jahre Brauerei – Hessen'),  # Licher Privatbrauerei
    'Q17326366': ('B', '150 Jahre Stadtbibliothek – Festschrift plausibel, städtisches Budget'),  # Stadtbibliothek Darmstadt
    'Q72927795': ('B', '150 Jahre Linde – Weltkonzern, macht es selbst; nur der Vollständigkeit halber'),  # Linde AG
    'Q120968418': ('B', '125 Jahre, Karlsbad – Verlagsgebiet'),  # ATSV Mutschelbach
    'Q319548': ('B', '125 Jahre, Rastatt – Verlagsgebiet'),  # FC Rastatt 04
    'Q1388226': ('B', '125 Jahre Fastnachtsgesellschaft, Karlsruhe'),  # FG Rüppurr
    'Q2458740': ('B', '125 Jahre, Pfalz'),  # TuS Hohenecken
    'Q795936': ('B', '175 Jahre Bank – Hessen, Jahr uneinheitlich'),  # ODDO BHF
    'Q1529839': ('B', '75 Jahre, Familienunternehmen, Kreis Lörrach'),  # Glatt Unternehmensgruppe
    'Q23788252': ('B', 'Liebherr – Konzern macht es selbst'),  # Liebherr-Hausgeräte
    'Q23826343': ('B', 'Liebherr – Konzern macht es selbst'),  # Liebherr-Werk Biberach
    'Q176850': ('B', '100 Jahre, Fußball, Kaiserstuhl'),  # Bahlinger SC
    'Q1023617': ('B', '50 Jahre, groß'),  # CHG-Meridian
    'Q1045931': ('B', '50 Jahre, Reisemobile'),  # Carthago Reisemobilbau
    'Q1444994': ('B', '200 Jahre – Frankfurt'),  # Frankfurter Kunstverein
    'Q1965958': ('B', '200 Jahre Naturkundeverein – Wiesbaden'),  # Nassauischer Verein für Naturkunde
    'Q875722': ('B', '200 Jahre Städelschule – Frankfurt'),  # Staatliche Hochschule für Bildende Künste – Städelschule
    'Q116895440': ('B', '150 Jahre Alpenvereinssektion – Festschrift üblich'),  # Sektion Ulm
    'Q162349': ('B', '125 Jahre Fußball'),  # 1. FC Normannia Gmünd
    'Q319868': ('B', '125 Jahre Fußball'),  # FC Singen 04
    'Q39562085': ('B', '125 Jahre Alpenvereinssektion – Festschrift üblich'),  # Sektion Stuttgart
    'Q109769788': ('B', '150 Jahre Fastnacht, Buchen – Verlagsgebiet'),  # FG Narrhalla
    'Q39080113': ('B', '75 Jahre Kammer'),  # Landesärztekammer Baden-Württemberg
    'Q1250760': ('B', '50 Jahre Geschichtsverein – Partner, aber kleine Zahl'),  # Geschichtsverein der Diözese Rottenburg-Stuttgart
    'Q649456': ('C', '900 Jahre, aber 80 Einwohner'),  # Hirschthal
    'Q562832': ('C', '750 Jahre, 240 Einwohner'),  # Heinzenhausen
    'Q563008': ('C', '700 Jahre, 440 Einwohner'),  # Neunkirchen am Potzberg
    'Q112151869': ('C', '900 Jahre, Wortlaut unsicher'),  # Broggingen
    'Q1790172': ('C', '800 Jahre, Weiler von Esslingen'),  # Krummenacker
    'Q63871491': ('C', '800 Jahre, Weiler'),  # Weiler
    'Q13142752': ('C', '750 Jahre, Burgruine/Weiler – Jahr nur aus Wikidata'),  # Landeck
    'Q2203516': ('C', '750 Jahre, kleiner Esslinger Stadtteil'),  # Rüdern
    'Q1586935': ('C', '700 Jahre, Weiler'),  # Harsberg
    'Q1146049': ('C', '700 Jahre, Erwähnung umstritten'),  # Ingersheim an der Jagst
    'Q2596006': ('C', '700 Jahre, Esslinger Wohngebiet'),  # Wäldenbronn
    'Q29567459': ('C', '700 Jahre, kleiner Ortsteil'),  # Zoznegg
    'Q58915406': ('C', 'Mühle/Hofgut'),  # Klinglesmühle
    'Q110892450': ('C', 'Hof'),  # Hirschhof
    'Q100344792': ('C', 'unklar, ob es die Firma noch gibt'),  # Georg Reichard
    'Q59434929': ('C', '800 Jahre, Erwähnung nur „möglicherweise“'),  # Dambach (Stödtlen)
    'Q112776368': ('C', '600 Jahre, Weiler, Jahr unsicher'),  # Altmannsweiler
    'Q1616201': ('C', 'Teil der Universität'),  # Universitätsbibliothek Stuttgart
    'Q1952619': ('C', '200 Jahre, Weiler'),  # Mummenroth
    'Q1394683': ('C', '200 Jahre, Weiler am Odenwaldkamm'),  # Tromm
    'Q107118168': ('C', 'vermutlich nicht mehr eigenständig – prüfen'),  # Enzinger Union Werke
    'Q109939644': ('C', 'unbekannt – prüfen'),  # Gathmann
    'Q30269260': ('C', 'Tochter eines schwedischen Konzerns'),  # Getinge (Germany)
    'Q47495905': ('C', 'Hessen, Schule'),  # Freie Waldorfschule Kassel
    'Q66098763': ('C', '125 Jahre Grundschule'),  # Heusteigschule
    'Q2587819': ('C', 'gehört zu einem US-Konzern (aus dem Gedächtnis)'),  # Wohlhaupter
    'Q1226535': ('C', 'Mühle, Jahr unsicher'),  # Dinkelmühle Graf
    'Q66439845': ('C', 'Musikgeschäft – Koblenz'),  # Musik-Klein
    'Q1451815': ('C', 'Hessen'),  # Precision Mechatronics
    'Q1609884': ('C', 'unklar – Mainz'),  # Riga Mainz
    'Q1676048': ('C', 'Fertighaus, Eifel'),  # Streif
    'Q25345785': ('C', 'Autohaus – Hessen'),  # Taunus-Auto
    'Q130367700': ('C', 'Hessen'),  # YMOS
    'Q1262599': ('C', '50 Jahre Hochschule, staatlich'),  # Duale Hochschule Baden-Württemberg Karlsruhe
    'Q1622186': ('C', '50 Jahre Hochschule'),  # Hochschule für Jüdische Studien
    'Q18027093': ('C', 'Hessen'),  # Raiffeisenbank Werratal-Landeck
    'Q300222': ('C', 'unbekannt'),  # AWG
    'Q1051811': ('C', 'Jazzkeller, 75 Jahre'),  # Cave 54
    'Q130241437': ('C', 'Wanderhütte'),  # Hirzeck-Haus
    'Q1961387': ('C', '75 Jahre Versicherung'),  # NRV-Rechtsschutz
    'Q1330577': ('C', 'Hessen, Schule'),  # Elisabethschule Marburg
    'Q2315885': ('C', 'Hessen, Schule'),  # St.-Angela-Schule (Königstein)
    'Q189884': ('C', 'Hessen, Schule'),  # Zentgrafenschule
    'Q1652777': ('C', 'Tennisclub'),  # TC Rüppurr Karlsruhe 1929
    'Q1733921': ('C', '25 Jahre'),  # Karlshochschule International University
    'Q139916607': ('C', '50 Jahre, klein'),  # Tria Technologies
    'Q111296742': ('C', '75 Jahre, klein'),  # Dossmann GmbH
    'Q112662540': ('C', 'Dorfverein'),  # FC Fortuna 1929 Lohrbach e.V.
    'Q112798369': ('C', 'Dorfverein'),  # Sportverein 1929 Schollbrunn e.V.
    'Q112843126': ('C', 'Dorfverein'),  # Sportverein 1929 Wagenschwend e.V.
    'Q16909390': ('C', '25 Jahre'),  # Fakultät für Rechtswissenschaft und Volkswirtschaftslehre (Universität Mannheim)
    'Q15242252': ('C', '25 Jahre'),  # Leonardo da Vinci Gymnasium
    'Q2208030': ('C', 'Hessen'),  # STI Group
    'Q802425': ('C', 'Hessen'),  # Freien Grunder Eisenbahn
    'Q1531851': ('C', 'Hessen'),  # Glockenbäckerei
    'Q111541235': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # Anwaltskanzlei Achim Edelmann
    'Q111470716': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # KWM Weisshaar
    'Q111686134': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # Tannenhof Beuchert
    'Q862856': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Billiger.de
    'Q1349176': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Erdgas Südwest
    'Q1774640': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Klinikum Mittelbaden
    'Q1650863': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Rhein-Neckar-Verkehr
    'Q2454048': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Triport
    'Q1605964': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Tucano
    'Q56477200': ('C', '75 Jahre – zu klein für ein Jubiläumsbuch'),  # Karl-Schubert-Schule Waldorfschule
    'Q16054243': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # Akademie für Weltmission
    'Q15819794': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # ISA Internationales Stadtbauatelier
    'Q138042824': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # Realschule Schenkensee
    'Q132447426': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Betriebsgesellschaft der Multifunktionsarena Mannheim
    'Q111898030': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # GMEINDER LOKOMOTIVEN GmbH
    'Q50038738': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Novaliq (Germany)
    'Q126924018': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # elprotek GmbH
    'Q2208645': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # SV Sparkassenversicherung Holding
    'Q2307638': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Sparkasse Rhein-Haardt
    'Q921885': ('C', '75 Jahre – zu klein für ein Jubiläumsbuch'),  # DIY Element System
    'Q23905207': ('C', 'Liebherr – Konzern macht es selbst'),  # Liebherr-Mischtechnik
    'Q1636515': ('C', '75 Jahre – zu klein für ein Jubiläumsbuch'),  # Sioux
    'Q1275885': ('C', '100 Jahre, kleiner Verein'),  # ESC Ulm
    'Q1387714': ('C', '100 Jahre, kleiner Verein'),  # FC Teningen
    'Q30988435': ('C', '100 Jahre, kleiner Verein'),  # Polizeisportverein Stuttgart
    'Q1589234': ('C', '100 Jahre, kleiner Verein'),  # SV 1930 Hockenheim
    'Q20754335': ('C', '75 Jahre – zu klein für ein Jubiläumsbuch'),  # Elly Heuss-Knapp-Schule
    'Q1622175': ('C', '75 Jahre – zu klein für ein Jubiläumsbuch'),  # Hochschule für Forstwirtschaft Rottenburg
    'Q1156271': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # DZB Bank
    'Q56476872': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # Freie Waldorfschule Ludwigsburg
    'Q67192745': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # Volkshochschule Mittleres Taubertal
    'Q1375788': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Landesgymnasium für Hochbegabte Schwäbisch Gmünd
    'Q185923': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Zeitenspiegel-Reportageschule Reutlingen
    'Q787854': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # Automobiltechnik Walther
    'Q918086': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # Britax Römer
    'Q122363477': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # Güde
    'Q1811552': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # Lederer Ragnarsdóttir Oei
    'Q114937173': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # Orgelbau Wiedenmann
    'Q27938128': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # Schwimmbad Musikclub
    'Q50038830': ('C', '75 Jahre – zu klein für ein Jubiläumsbuch'),  # Ivoclar Vivadent (Germany)
    'Q74110760': ('C', '100 Jahre, kleiner Verein'),  # FC Epfendorf 1929
    'Q2307146': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Sparda-Bank Hessen
    'Q1615776': ('C', '75 Jahre – zu klein für ein Jubiläumsbuch'),  # Hessen Mobil – Straßen- und Verkehrsmanagement
    'Q579048': ('C', '75 Jahre – zu klein für ein Jubiläumsbuch'),  # Korth
    'Q29015552': ('C', '75 Jahre – zu klein für ein Jubiläumsbuch'),  # Leppermühle
    'Q2017108': ('C', '75 Jahre – zu klein für ein Jubiläumsbuch'),  # Okapia
    'Q2250899': ('C', '75 Jahre – zu klein für ein Jubiläumsbuch'),  # Schuldorf Bergstraße
    'Q2252487': ('C', '75 Jahre – zu klein für ein Jubiläumsbuch'),  # Schwab Versand
    'Q1696671': ('C', '75 Jahre – zu klein für ein Jubiläumsbuch'),  # Sektkellerei Herres
    'Q4719455': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # Alexander M. Patch American High School
    'Q870659': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # ESB Business School
    'Q137796697': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # Gfw - Gesellschaft für Fremdenverkehrswerbung
    'Q50842356': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # MCC-Kino-Center
    'Q883110': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Blauwald
    'Q15110012': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # E-Werk Mittelbaden
    'Q2693387': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Euronics Deutschland
    'Q1620488': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Hirschmann Automation and Control
    'Q131931899': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # K2 Systems
    'Q1445025': ('C', '100 Jahre, kleiner Verein'),  # Frankfurter Psychoanalytisches Institut
    'Q427160': ('C', '100 Jahre, kleiner Verein'),  # SC Waldgirmes
    'Q15843651': ('C', '100 Jahre, kleiner Verein'),  # SV Klausen
    'Q56249664': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # Freie Waldorfschule Mainz
    'Q338517': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Accadis International School Bad Homburg
    'Q1377530': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Europäische Studienakademie Kälte-Klima-Lüftung
    'Q1806348': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # TU Darmstadt, Universitäts- und Landesbibliothek
    'Q635510': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # Archiv Frau und Musik
    'Q1200317': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # Derpart
    'Q48180528': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # Driver Reifen und KFZ-Technik
    'Q1245447': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # Kohlpharma
    'Q823197': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # TIBHAR
    'Q2399494': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # Technolit
    'Q2467185': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # Union Krankenversicherung
    'Q2539208': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # Wagner Solar GmbH
    'Q882307': ('C', '125 Jahre, kleiner Verein'),  # Blau-Weiß Wiehre Freiburg
    'Q133328779': ('C', '125 Jahre, kleiner Verein'),  # ETSV Lauda
    'Q1389911': ('C', '125 Jahre, kleiner Verein'),  # FV Bad Saulgau 04
    'Q573782': ('C', '125 Jahre, kleiner Verein'),  # Olympia Laupheim
    'Q104051104': ('C', '125 Jahre, kleiner Verein'),  # Spvgg Mössingen
    'Q1827272': ('C', '125 Jahre, kleiner Verein'),  # TSV Ofterdingen
    'Q55209501': ('C', '75 Jahre – zu klein für ein Jubiläumsbuch'),  # Deutsche Tele Medien
    'Q30269375': ('C', '75 Jahre – zu klein für ein Jubiläumsbuch'),  # SymbioPharm (Germany)
    'Q83945866': ('C', '75 Jahre – zu klein für ein Jubiläumsbuch'),  # Volkswagen Service Kassel
    'Q83938184': ('C', '75 Jahre – zu klein für ein Jubiläumsbuch'),  # Volkswagen Zentrum Kassel
    'Q6731260': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Magna PT B.V. & Co. KG
    'Q7120122': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # PPI Automotive Design
    'Q32361064': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Silverback
    'Q112916015': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Sola Media
    'Q82075822': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # SpaceTech GmbH
    'Q30285304': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Synovo
    'Q7950779': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # WINGcon
    'Q2228731': ('C', '75 Jahre – zu klein für ein Jubiläumsbuch'),  # Scala
    'Q135747726': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # Berufsbildungswerk CJD Homburg/Saar
    'Q30256279': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # FVS
    'Q7959593': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # Wagner & Co Solar Technology
    'Q357366': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Adinotec
    'Q831808': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Betriebs-Center für Banken
    'Q894377': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Bosch Thermotechnik
    'Q131785711': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Fahrzeugmanagement Region Frankfurt RheinMain
    'Q1367630': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # MainMobil
    'Q2249155': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Schott Solar
    'Q19309712': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Sybac Solar
    'Q48813999': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenzunft Aixheim
    'Q133304514': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # CIU Korntal
    'Q34192864': ('C', '50 Jahre – zu klein für ein Jubiläumsbuch'),  # Kulturzentrum Linse
    'Q162186': ('C', '125 Jahre, kleiner Verein'),  # 1. FCA Darmstadt
    'Q666059': ('C', '125 Jahre, kleiner Verein'),  # Borussia Fulda
    'Q50639122': ('C', '125 Jahre, kleiner Verein'),  # Sektion Koblenz des Deutschen Alpenvereins
    'Q15848061': ('C', '125 Jahre, kleiner Verein'),  # SpVgg Burgbrohl
    'Q112091540': ('C', '50 Jahre, kleiner Verein'),  # MFC Ikarus e.V.
    'Q111029088': ('C', '50 Jahre, kleiner Verein'),  # TC Schlierstadt 1979 e.V.
    'Q16837472': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # European Group
    'Q5577181': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Goethe Business School
    'Q43402241': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # HEAG mobiBus GmbH & Co. KG
    'Q43402515': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # HEAG mobiTram GmbH & Co. KG
    'Q30257082': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Hessen Agentur
    'Q141726583': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # JSLM AG
    'Q29439907': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # Montessori-Verein Karben-Bad Vilbel e.V.
    'Q1939300': ('C', '25 Jahre, kleiner Verein'),  # Rat für deutsche Rechtschreibung
    'Q8042840': ('C', '25 Jahre – zu klein für ein Jubiläumsbuch'),  # xaitment
    'Q22674520': ('C', '75 Jahre Kino'),  # Bali Kino
    'Q20635013': ('C', '75 Jahre, kleiner Verein'),  # Grün-Weiß Hochspeyer
    'Q110915764': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Fastnachtsgesellschaft Getzemer Narre
    'Q112245993': ('C', '75 Jahre, kleiner Verein'),  # Tennisclub Schwarz-Weiß Adelsheim e.V.
    'Q1622214': ('C', '50 Jahre, kleiner Verein'),  # Hochschule für Polizei Baden-Württemberg
    'Q23786546': ('C', '50 Jahre, kleiner Verein'),  # SC Gutach-Bleibach
    'Q47525618': ('C', '75 Jahre, kleiner Verein'),  # Sioux-West Freiburg
    'Q48920193': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenzunft Keaschmecker Betra
    'Q112919056': ('C', '50 Jahre, kleiner Verein'),  # Reiterfreunde Hüffenhardt
    'Q112898176': ('C', '50 Jahre, kleiner Verein'),  # Tennisverein Blau-Weiß '79 Reichenbuch e.V.
    'Q1008924': ('C', '75 Jahre, kleiner Verein'),  # Bundesverband Materialwirtschaft, Einkauf und Logistik
    'Q822551': ('C', '75 Jahre, kleiner Verein'),  # Rot-Weiß Darmstadt
    'Q47525612': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Friburger Schloßberggeister 1979
    'Q1181399': ('C', '25 Jahre, kleiner Verein'),  # Debattierclub Stuttgart
    'Q1676519': ('C', '25 Jahre, kleiner Verein'),  # ForAfrika Deutschland
    'Q96745617': ('C', '25 Jahre, kleiner Verein'),  # SC04 Tuttlingen
    'Q829133': ('C', '50 Jahre, kleiner Verein'),  # Berufsverband für den Rettungsdienst
    'Q1203557': ('C', '50 Jahre, kleiner Verein'),  # Deutsche Migräne- und Kopfschmerz-Gesellschaft
    'Q1575284': ('C', '50 Jahre, kleiner Verein'),  # Handwerkskammer Frankfurt-Rhein-Main
    'Q2229432': ('C', '75 Jahre, kleiner Verein'),  # Schachclub Eppingen
    'Q47784673': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Bahlinger Silberberg-Hexen 2004
    'Q50018898': ('C', '25 Jahre, kleiner Verein'),  # Kreuzritter vom Argental
    'Q50319898': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Rottenburger Schelmenteufel
    'Q141118356': ('C', '50 Jahre, kleiner Verein'),  # Türk Gücü Hanau
    'Q47516490': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenverein Rötenbach
    'Q47498652': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrhalla nelraM
    'Q47516457': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Bärenzunft Bärental
    'Q48977812': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Fell-Lädsche-Zunft Marbach
    'Q47784801': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Hafennarren Friedrichshafen
    'Q47486465': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenfreunde Heuberg
    'Q48781337': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenvereinigung Bärenthal
    'Q47498677': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenzunft Wierer Frösch
    'Q117843259': ('C', '25 Jahre, kleiner Verein'),  # Bundesverband für Ergotherapeut:innen in Deutschland
    'Q140233237': ('C', '25 Jahre, kleiner Verein'),  # Förderverein Gentoo
    'Q47519967': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenzunft Maulburg 1954
    'Q47516474': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenverein Reetiwölf
    'Q48549549': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenzunft Meerrettichdämone Urloffen
    'Q47784944': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Wolfszunft Wolfenweiler 1979
    'Q30269106': ('X', 'Gründung des französischen Mutterkonzerns'),  # Ipsen (Germany)
    'Q113620553': ('X', 'ehemalige Schule'),  # Realprogymnasium (Buchen)
    'Q1447863': ('X', '2023 insolvent, an Zeitfracht (aus dem Gedächtnis)'),  # Koch, Neff und Volckmar
    'Q115142921': ('X', 'historischer Postkartenverlag'),  # Portrait Atelier Bauer
    'Q1782491': ('X', 'Mediengruppe – macht es selbst'),  # Konradin Mediengruppe
    'Q19964300': ('X', 'Gebäude bzw. Vorläufer – Ansprechpartner ist das Ludwig-Georgs-Gymnasium'),  # Altes Pädagog
    'Q1152019': ('X', 'in AXA aufgegangen (aus dem Gedächtnis)'),  # DBV-Winterthur Holding
    'Q47530744': ('X', 'Gebäude'),  # Fachwerkhaus Weilburger Straße 14
    'Q107985183': ('X', 'Bäckerei, Hessen'),  # Hinnerbäcker
    'Q19971620': ('X', 'Verlag'),  # Edition Gorz
    'Q1395401': ('X', 'Joint Venture beendet'),  # Family Media
    'Q19963254': ('X', 'mehrfach umbenannt, Konzern'),  # ThyssenKrupp System Engineering
    'Q1718850': ('X', '2023 insolvent, an Zeitfracht (aus dem Gedächtnis)'),  # KNO Verlagsauslieferung
    'Q1488660': ('X', 'Verlag'),  # Gabler Verlag
    'Q108597725': ('X', 'Datensatz ohne Namen'),  # Q108597725
    'Q134424136': ('X', 'Internetdienst'),  # Blinkx
    'Q111308810': ('X', 'Server'),  # Dagstuhl Research Online Publication Server
    'Q83853327': ('X', 'Autohaus, Datenfehler'),  # Porsche Zentrum Erfurt
    'Q30282423': ('X', 'Konzerntochter'),  # Sanofi (Germany)
    'Q133805486': ('X', 'Behördenbetrieb'),  # Wirtschaftsbetrieb Mainz
    'Q2645334': ('X', 'Verlag – macht es selbst'),  # Alfred Kröner Verlag
    'Q1384477': ('X', 'Verlag'),  # Expert Verlag
    'Q1735714': ('X', 'Verlag'),  # Westend Verlag
    # Bayerische Grenzkreise (Nachtrag November 2026)
    'Q573628': ('S', '1250 Jahre, 9 700 Einwohner, Rokokoschloss der Würzburger Fürstbischöfe'),  # Veitshöchheim
    'Q55523887': ('S', '250 Jahre, Familienunternehmen, Kreuzwertheim gegenüber Wertheim'),  # Kurtz Ersa
    'Q65179937': ('A', '175 Jahre Diakonissenanstalt Neuendettelsau (Löhe 1854) – großer Träger'),  # DIAKONEO
    'Q636150': ('A', '1250 Jahre, tagesgenau (14. 10. 779), 2 700 Einwohner'),  # Kist
    'Q536407': ('A', '1250 Jahre, 4 500 Einwohner, bei Würzburg'),  # Kürnach
    'Q508445': ('A', '1250 Jahre, 9 500 Einwohner – Jahr 779/780 nicht eindeutig'),  # Höchberg
    'Q1773361': ('A', '175 Jahre Sparkasse – Festschrift typisch'),  # Sparkasse Neu-Ulm - Illertissen
    'Q1943544': ('A', '100 Jahre, Familienmolkerei, groß'),  # Molkerei Gropper
    'Q1627195': ('A', '150 Jahre, Familienunternehmen (Logistik), Neu-Ulm'),  # Honold Logistik Gruppe
    'Q2038417': ('A', '150 Jahre, Familienunternehmen (Waschanlagen), Memmingen'),  # Otto Christ
    'Q21039847': ('B', '1250 Jahre, Ortsteil von Giebelstadt'),  # Herchsheim
    'Q505756': ('B', '900 Jahre, 1 800 Einwohner – Beleg nennt einen Ritter „von Dahspach“'),  # Dachsbach
    'Q506190': ('B', '800 Jahre, 2 200 Einwohner'),  # Arberg
    'Q528727': ('B', '800 Jahre, 2 500 Einwohner'),  # Weidenbach
    'Q508471': ('B', '1250 Jahre, aber 1 100 Einwohner'),  # Gössenheim
    'Q17322305': ('B', '900 Jahre, 1 400 Einwohner, Stadtteil von Ichenhausen'),  # Hochwang
    'Q127388879': ('B', '175 Jahre Schule – Jahr nur aus Wikidata'),  # Laurentius-Gymnasium
    'Q1526374': ('B', '175 Jahre, Klebebänder (aus dem Gedächtnis) – Jahr nur aus Wikidata'),  # Monta
    'Q47767582': ('C', 'Urkunde 779/803 – Jahr unsicher, 160 Einwohner'),  # Urfersheim
    'Q15106662': ('C', '750 Jahre, kleiner Ortsteil von Roden'),  # Ansbach
    'Q17461472': ('C', '750 Jahre, Weiler'),  # Betzenhausen
    'Q66818329': ('C', 'Mühle'),  # Wiedergeltinger Mühle
    'Q13726820': ('C', '600 Jahre, Weiler'),  # Hinterstaufen
    'Q133933922': ('C', '600 Jahre, Weiler'),  # Westerhofen
    'Q1941936': ('C', '900 Jahre, aber 57 Einwohner'),  # Modelshausen
    'Q114355862': ('C', 'Einöde, 5 Einwohner'),  # Schnattern
    'Q111350243': ('C', 'Weiler, 18 Einwohner'),  # Gerstland
    'Q1235109': ('C', 'Mühle, 5 Einwohner'),  # Pilsmühle
    'Q2649693': ('C', 'Hof, 12 Einwohner'),  # Almhof
    'Q1552376': ('C', 'Berghaus'),  # Grüntenhaus
    'Q50038495': ('C', '50 Jahre – kleine Zahl'),  # Nuvisan (Germany)
    'Q134442509': ('C', '125 Jahre, Sportverein – Festheft'),  # Sportgemeinde Simmerberg 1904
}

KOMMENTAR = {
    'Q13896': '1250 Jahre seit der Schenkung an Fulda 779. Stadt mit 14 000 Einwohnern und eigenem Haushalt. Sie hat 1979 die 1200 Jahre gefeiert, das Jubiläum hat dort also Tradition. Ansprechpartner: Bürgermeisteramt bzw. Stadtarchiv. Ein Ortsbuch ist hier sehr wahrscheinlich – die Frage ist nur, wer es macht.',  # Markgröningen
    'Q355435': '1250 Jahre (779, Kloster Fulda), kleine Stadt mit 5 100 Einwohnern im Neckar-Odenwald-Kreis – mitten im Verlagsgebiet. Erholungsort, also auch Interesse am Stadtmarketing. Klassischer Kunde für eine Stadtgeschichte.',  # Adelsheim
    'Q491535': '1200 Jahre (829 „Wimpina“). Kaiserpfalz, Altstadt, Tourismus – eine Stadt, die ihre Geschichte verkauft. Ein Jubiläumsbuch passt in das Stadtmarketing und lässt sich danach im Tourismus weiterverkaufen.',  # Bad Wimpfen
    'Q61830': '1250 Jahre (Fulda 779, „Boteburon“). Stadt mit 8 000 Einwohnern, Weinort im Bottwartal. Gleiche Urkunde wie Markgröningen und Vaihingen – mehrere Kunden mit demselben Anlass.',  # Großbottwar
    'Q23023': '800 Jahre (1229). Gemeinde mit 6 000 Einwohnern im Rhein-Pfalz-Kreis, nördlich von Speyer. Kurzer Weg, der Verlag ist in der Gegend bekannt.',  # Waldsee
    'Q74763163': '1250 Jahre (Lorscher Codex 779). Dürrmenz ist der älteste Teil von Mühlacker und war lange der wichtigere Ort; Mühlacker selbst wird erst 1292 erwähnt. Der eigentliche Kunde ist also die Stadt Mühlacker, die größte Stadt im Enzkreis. Verlagsgebiet.',  # Dürrmenz
    'Q2586953': '175 Jahre, Familienunternehmen in Pforzheim, gilt als Erfinder des Metallschlauchs (aus dem Gedächtnis). Ein großes Firmenjubiläum mit Geld dahinter, im Verlagsgebiet. Konkurrenz durch Agenturen für Firmengeschichte möglich, der regionale Bezug spricht für den Verlag.',  # Witzenmann-Gruppe
    'Q17122614': '150 Jahre, Familienunternehmen für Laborbedarf in Karlsruhe, rund 20 km vom Verlag. Familienunternehmen mit langer Geschichte bestellen zum 150. gern ein Buch.',  # Carl Roth
    'Q61653': '1250 Jahre (779 „Bunninga“), 6 300 Einwohner, Kreis Ludwigsburg. Kleinere Gemeinde, aber starke Zahl.',  # Benningen am Neckar
    'Q507274': '1250 Jahre (779, Kloster St. Gallen), 4 100 Einwohner. Liegt am Bodensee, also weit weg, aber die Zahl ist stark.',  # Bermatingen
    'Q61705': '1250 Jahre (Lorscher Codex 779), 6 300 Einwohner, Kreis Ludwigsburg.',  # Ingersheim
    'Q61928': 'Mit 28 500 Einwohnern die größte Stadt der Liste. Der Artikel sagt aber nur, die erste Erwähnung „könnte“ 779 sein. Vorher klären, ob die Stadt 2029 feiert. Wenn ja, wird daraus ein S-Lead.',  # Vaihingen an der Enz
    'Q80809': '900 Jahre (1129, Papsturkunde für das Kloster Denkendorf). 11 000 Einwohner, Kreis Esslingen. Ort und Kloster feiern dasselbe Jubiläum.',  # Denkendorf
    'Q80606': '800 Jahre (1229), 5 600 Einwohner, Kreis Esslingen.',  # Oberboihingen
    'Q515323': '1100 Jahre (929), kleine Gemeinde am Hochrhein (3 200 Einwohner). Die Ortsteile sind noch älter (Bannholz 871, Aisperg 890) – Stoff für ein dickes Buch.',  # Weilheim
    'Q80511': '800 Jahre (1229, Kloster Bebenhausen), 2 900 Einwohner, Enzkreis. Kleine Gemeinde im Verlagsgebiet; Budget prüfen.',  # Wimsheim
    'Q43273976': '1250 Jahre (Lorscher Codex 779). Die Urkunde nennt „Sahsenheim“ für Groß-, Lützel- und Hohensachsen. Großsachsen gehört zu Hirschberg, die beiden anderen zu Weinheim (aus dem Gedächtnis). Denkbar ist ein gemeinsames Fest von zwei Kommunen. Rhein-Neckar-Kreis, Verlagsgebiet.',  # Großsachsen
    'Q20179999': '1250 Jahre (779, St. Gallen), ältester Ortsteil von Zimmern ob Rottweil. Typischer Fall für eine Ortschronik mit Heimatverein und Gemeinde als Geldgeber.',  # Flözlingen
    'Q2016767': '1250 Jahre (779), Ortsteil von Öhringen. Gleiches Muster wie Flözlingen.',  # Ohrnberg
    'Q16006232': '700 Jahre als Lateinschule (1329) – ein außergewöhnliches Schuljubiläum. Eine Festschrift ist sehr wahrscheinlich. Im Allgäu weit weg, aber eine solche Zahl kommt selten vor.',  # Rupert-Neß-Gymnasium Wangen
    'Q798951': '100 Jahre, Familienunternehmen (Versandhandel) in Pforzheim, Verlagsgebiet.',  # Bader Versand
    'Q520675': '100 Jahre, Familienbetrieb in Cleebronn, gilt als ältester Freizeitpark Deutschlands (aus dem Gedächtnis). Bildreiche Geschichte – gut für ein schönes Buch.',  # Erlebnispark Tripsdrill
    'Q1746794': '150 Jahre Privatbank in Freiburg. Privatbanken pflegen ihre Tradition, ein Jubiläumsbuch ist typisch.',  # Bankhaus E. Mayer
    'Q16002754': '175 Jahre, Familienunternehmen (Nähgarne) in Bönnigheim.',  # Amann & Söhne
    'Q1485474': '175 Jahre Kranbau, Heilbronn. Bekannte Marke mit langer Geschichte.',  # Wolffkran
    'Q1333317': '150 Jahre, börsennotierter Autozulieferer mit über 6 000 Mitarbeitern. Geld ist da, aber vermutlich auch eine eigene Kommunikationsabteilung oder Agentur.',  # ElringKlinger
    'Q1811343': '150 Jahre, Familienunternehmen (Düsentechnik) in Metzingen.',  # Lechler (Unternehmen)
    'Q322938': '175 Jahre Kunstakademie in Karlsruhe. Eine Festschrift ist so gut wie sicher, das Land zahlt. Die Akademie gestaltet vielleicht selbst – dann als Verlag dazukommen.',  # Staatliche Akademie der Bildenden Künste Karlsruhe
    'Q19930480': '500 Jahre – eine sehr seltene Zahl für einen Verein. Ortenaukreis.',  # Schützenverein Oberkirch
    'Q106394': '125 Jahre, Bundesligaverein. Eine Chronik kommt sicher, Geld ist da. Ob ein Regionalverlag zum Zug kommt, ist offen – einen Versuch wert.',  # SC Freiburg
    'Q19963065': '100 Jahre Geschichtsverein – das ist die Kernkundschaft eines Regionalverlags. Ein Festband oder Jahrbuch zum Jubiläum liegt nahe.',  # Geschichtsverein Markgräflerland
    'Q47796855': '1250 Jahre – aber nur „vermutlich“ 779, in derselben Fulda-Urkunde wie Markgröningen und Vaihingen. Wenn die Nachbarn feiern, feiert Hochdorf (Gemeinde Eberdingen) vermutlich mit.',  # Hochdorf/Enz
    'Q2284859': '1250 Jahre, aber die Urkunde ist nur auf 779–802 datiert. Ortsteil von Neudenau. Erst nachfragen, ob man 2029 feiern will.',  # Siglingen
    'Q889245': '1200 Jahre (Lorscher Codex 829), war aber ursprünglich nur ein Hofgut. Kleiner Ortsteil von Kirchardt – die Gemeinde müsste das Buch tragen.',  # Bockschaft
    'Q18019848': '1100 Jahre (929, St. Gallen) – starke Zahl, aber ein kleiner Ortsteil von Stühlingen ganz im Süden. Ortschronik denkbar, weit weg vom Verlag.',  # Eberfingen
    'Q1469016': '800 Jahre, Stuttgarter Stadtbezirk mit 9 100 Einwohnern. Wangen wurde 1229 „zusammen mit Stuttgart“ erwähnt. Ein Bezirksbuch ist möglich, Geld kommt eher vom Bürgerverein als von der Stadt.',  # Wangen
    'Q546393': '750 Jahre (1279), Gemeinde mit 3 000 Einwohnern. Dort liegt auch Tripsdrill (A-Lead, 100 Jahre) – zwei Jubiläen im selben Ort, das ließe sich verbinden.',  # Cleebronn
    'Q537465': '750 Jahre (1279), Gemeinde mit 2 600 Einwohnern im Zabergäu. Kleiner Haushalt.',  # Pfaffenhofen
    'Q19969752': '800 Jahre (1229), Ortsteil von Wangen im Allgäu mit 2 500 Einwohnern. Weit weg.',  # Schomburg
    'Q540770': '900 Jahre (1129), aber nur 860 Einwohner. Eher eine Festschrift als ein Buch.',  # Unterstadion
    'Q632934': '800 Jahre (1229), nur 600 Einwohner. Die Teilorte sind älter (Talheim 776) – inhaltlich reizvoll, das Budget dürfte klein sein.',  # Lauterach
    'Q373283': '900 Jahre (1129), Pfalz, Südwestpfalz – aber nur 380 Einwohner. Kleines Festheft eher als Buch.',  # Schönau
    'Q1308178': '800 Jahre (1229), Zell am Neckar, heute ein Stadtteil von Esslingen. Mit Mettingen und Rüdern feiern 2029 gleich mehrere Esslinger Stadtteile – vielleicht ein Sammelband für die Stadt Esslingen.',  # Zell
    'Q1925995': '800 Jahre (1229), Esslinger Stadtteil. Siehe Zell – zusammen mit der Stadt Esslingen ansprechen.',  # Mettingen (Esslingen am Neckar)
    'Q1738511': '800 Jahre (1229), Ortsteil von Ostfildern.',  # Kemnat
    'Q1566803': '800 Jahre (1229), Weinort, Ortsteil von Brackenheim. Brackenheim hat mit der VBU Volksbank (125, B) noch einen Anlass.',  # Haberschlacht
    'Q33510632': '800 Jahre (1229), Ortsteil von Jettingen, Kreis Böblingen.',  # Unterjettingen
    'Q17577331': '800 Jahre, Ortsteil von Wangen im Allgäu, 1 000 Einwohner. Jahr nur aus Wikidata.',  # Leupolz
    'Q2350567': '800 Jahre, Ortsteil von Gomaringen. Jahr nur aus Wikidata, ohne Beleg – prüfen.',  # Stockach (Gomaringen)
    'Q2549872': '700 Jahre (1329), Ortsteil von Altensteig, Kreis Calw – liegt im Verlagsgebiet, ist aber klein.',  # Wart
    'Q122453': '200 Jahre. Groß und mit Geld, macht die Festschrift aber vermutlich mit eigenen Historikern und einem Wissenschaftsverlag. Nur mit gutem Kontakt.',  # Universität Stuttgart
    'Q89491135': '150 Jahre (offizielles Gründungsjahr 1879). Schulen machen fast immer eine Festschrift – kleines Budget, aber verlässlich.',  # Elly-Heuss-Knapp-Gymnasium Stuttgart
    'Q1555507': '175 Jahre Berufsschule. Festschrift wahrscheinlich, Schulträger ist der Landkreis bzw. die Stadt.',  # Gustav-von-Schmoller-Schule Heilbronn
    'Q133932437': '175 Jahre Berufsschule in Heilbronn. Wie Schmoller-Schule – beide zusammen ansprechen.',  # Johann-Jakob-Widmann-Schule
    'Q134068397': '175 Jahre Berufsschule.',  # Werner-Siemens-Schule Stuttgart
    'Q15688375': '125 Jahre Gymnasium.',  # Friedrich-Gymnasium Freiburg
    'Q139863925': '125 Jahre, aber das Gründungsjahr ist uneinheitlich – vorher klären.',  # Parler-Gymnasium Schwäbisch Gmünd
    'Q39080113': '75 Jahre Zusammenschluss der vier Kammern. Institution mit Geld, aber kleine Zahl – eher eine Broschüre.',  # Landesärztekammer Baden-Württemberg
    'Q1250760': '50 Jahre. Geschichtsvereine sind Kernkundschaft, die Zahl ist klein – als Kontakt für später pflegen.',  # Geschichtsverein der Diözese Rottenburg-Stuttgart
    'Q2505471': '125 Jahre Volksbank in Brackenheim. Genossenschaftsbanken machen gern Festschriften, oft mit Ortsgeschichte. Fusionen prüfen.',  # VBU Volksbank im Unterland
    'Q17477739': '150 Jahre. Vorher prüfen, ob sie noch eigenständig ist.',  # Raiffeisenbank Kaiserstuhl
    'Q15134090': '150 Jahre, kleine Bank. Vorher prüfen, ob sie noch eigenständig ist.',  # Volksbank Möckmühl
    'Q1890715': '150 Jahre. Gehört meines Wissens zur Continentale-Gruppe (aus dem Gedächtnis) – ob die Marke noch selbst feiert, ist fraglich.',  # Mannheimer Versicherungen
    'Q1597044': '150 Jahre laut Wikidata (1879) – das Gründungsjahr der Brauerei vorher prüfen.',  # Heinrich Dinkelacker
    'Q1774105': '125 Jahre Sägewerk in Oberrot. Eigentümerwechsel prüfen.',  # Klenk Holz
    'Q1131085': '150 Jahre Maschinenbau in Stuttgart. Gehört einem US-Konzern (aus dem Gedächtnis) – Entscheidung eher in der Konzernzentrale.',  # Coperion
    'Q125578831': '100 Jahre Kraftwerksgesellschaft am Hochrhein. Energieversorger machen zu runden Jubiläen gern eine Festschrift.',  # Rheinkraftwerk Albbruck-Dogern AG
    'Q2231851': '100 Jahre Kino in Karlsruhe, eines der ältesten der Stadt. Schöne Geschichte, aber kleines Budget – eher ein Bildband auf eigenes Risiko.',  # Schauburg
    'Q18023330': '75 Jahre, Werkzeughersteller in Horb (Kreis Freudenstadt, Verlagsgebiet). Kleine Zahl.',  # Leuco
    'Q1529839': '75 Jahre, Familienunternehmen im Kreis Lörrach. Kleine Zahl.',  # Glatt Unternehmensgruppe
    'Q23826343': '75 Jahre. Liebherr macht das als Konzern selbst – nur der Vollständigkeit halber.',  # Liebherr-Werk Biberach
    'Q23788252': '75 Jahre. Wie Liebherr-Werk Biberach.',  # Liebherr-Hausgeräte
    'Q2044808': '50 Jahre, Softwarefirma in Karlsruhe (Verkehrsplanung). Junge Firma, eher eine Broschüre.',  # PTV
    'Q22909101': '50 Jahre, börsennotiert, Enzkreis. Kleine Zahl.',  # STRATEC
    'Q1023617': '50 Jahre, großer Leasinganbieter in Weingarten. Kleine Zahl, aber Geld ist da.',  # CHG-Meridian
    'Q1045931': '50 Jahre, Reisemobile, Aulendorf. Kleine Zahl.',  # Carthago Reisemobilbau
    'Q116895440': '150 Jahre Alpenvereinssektion. Sektionen machen traditionell eine Festschrift.',  # Sektion Ulm
    'Q39562085': '125 Jahre Alpenvereinssektion, eine der großen Sektionen.',  # Sektion Stuttgart
    'Q109769788': '150 Jahre Fastnachtsgesellschaft in Buchen – Verlagsgebiet, für eine Narrhalla eine große Zahl.',  # FG Narrhalla
    'Q120968418': '125 Jahre, Karlsbad (Landkreis Karlsruhe) – direkt vor der Haustür. Verein, kleines Budget.',  # ATSV Mutschelbach
    'Q319548': '125 Jahre, Rastatt – Verlagsgebiet. Verein, kleines Budget.',  # FC Rastatt 04
    'Q1388226': '125 Jahre Fastnachtsgesellschaft in Karlsruhe.',  # FG Rüppurr
    'Q2458740': '125 Jahre, Kaiserslautern (Pfalz).',  # TuS Hohenecken
    'Q325249': '100 Jahre, Zweibrücken (Pfalz).',  # SVN Zweibrücken
    'Q176850': '100 Jahre, bekannter Amateurverein am Kaiserstuhl.',  # Bahlinger SC
    'Q162349': '125 Jahre Fußball, Schwäbisch Gmünd.',  # 1. FC Normannia Gmünd
    'Q319868': '125 Jahre Fußball, Singen.',  # FC Singen 04
    'Q634084': '1200 Jahre (829, Urkunde Ludwigs des Frommen). Gemeinde mit 12 700 Einwohnern im Kreis Groß-Gerau. Trebur war karolingische Königspfalz (Tribur) – Stoff für ein großes Geschichtsbuch, und die Gemeinde ist groß genug, es zu bezahlen. Bester Lead in Südhessen und Rheinhessen.',  # Trebur
    'Q1998041': '900 Jahre (1129, Schenkung Bischof Burchards von Worms). Ortsteil von Biblis. Hat 1979 schon das 850-jährige Bestehen gefeiert – die Tradition ist da.',  # Nordheim
    'Q1874094': '400 Jahre (1629 als Pädagog gegründet). Eine der ältesten Schulen Hessens – zu so einer Zahl gibt es eine Festschrift. Darmstadt, gleich an der Bergstraße.',  # Ludwig-Georgs-Gymnasium
    'Q1665423': '200 Jahre Schule in Darmstadt. Festschrift üblich, kleines Budget.',  # Viktoriaschule
    'Q1284907': '175 Jahre Schule in Darmstadt.',  # Edith-Stein-Schule Darmstadt
    'Q17326366': '150 Jahre. Bibliotheken machen zu runden Jubiläen gern einen Band – Träger ist die Stadt.',  # Stadtbibliothek Darmstadt
    'Q1798208': '700 Jahre (1329), Ortsteil von Michelstadt im Odenwald.',  # Steinbuch
    'Q2582423': '700 Jahre (1329), kleiner Ortsteil von Bensheim an der Bergstraße.',  # Wilmshausen
    # Bayerische Grenzkreise
    'Q573628': '1250 Jahre (779 erstmals erwähnt). Gemeinde mit 9 700 Einwohnern direkt bei Würzburg, ehemalige Sommerresidenz der Würzburger Fürstbischöfe mit berühmtem Rokokogarten. Große Zahl, wohlhabende Stadtrandgemeinde, Schloss und Hofgarten als Stoff – Kandidat für ein großes Ortsbuch.',  # Veitshöchheim
    'Q55523887': '250 Jahre laut Wikidata (1779). Familienunternehmen in Kreuzwertheim, direkt gegenüber von Wertheim – praktisch am Rand des Verlagsgebiets. Maschinenbau mit 250-jähriger Familiengeschichte (Anfang als Eisenhammer im Haslochtal, aus dem Gedächtnis): der klassische Fall für ein Firmenbuch, wie Friedr. Dick 2028.',  # Kurtz Ersa
    'Q65179937': '175 Jahre (9. Mai 1854, Diakonissenanstalt Neuendettelsau, gegründet von Wilhelm Löhe). Großer diakonischer Träger im Kreis Ansbach. Gründungstag und Gründer sind klar belegt – eine Jubiläumschronik ist sehr wahrscheinlich, mit Kirchen- und Sozialgeschichte als Stoff.',  # DIAKONEO
    'Q636150': '1250 Jahre (14. Oktober 779, erste Würzburger Markbeschreibung, „Chistesbrunno“). 2 700 Einwohner, Kreis Würzburg. Tagesgenauer Beleg – die Gemeinde kann den Tag feiern. Klein, aber eine starke Zahl.',  # Kist
    'Q536407': '1250 Jahre (779 „Quirnaha“). 4 500 Einwohner, Stadtrandgemeinde bei Würzburg. Das Jahr steht im Artikel schon in der Einleitung – die Gemeinde kennt ihr Jubiläum.',  # Kürnach
    'Q508445': '1250 Jahre nach der zweiten Würzburger Markbeschreibung (779/780) – das Jahr ist nicht eindeutig, die Gemeinde könnte auch 2030 feiern. 9 500 Einwohner, Nachbar von Würzburg. Vor der Anfrage klären, welches Jahr Höchberg zählt.',  # Höchberg
    'Q1773361': '175 Jahre laut Wikidata (1854). Sparkasse direkt an der Grenze zu Ulm. Sparkassen begehen solche Zahlen typischerweise mit Festschrift – wie die Kreissparkasse Biberach 2028.',  # Sparkasse Neu-Ulm - Illertissen
    'Q1943544': '100 Jahre laut Wikidata (1929). Familienmolkerei in Bissingen (Kreis Dillingen), heute ein großer Hersteller (aus dem Gedächtnis). Familie und Wachstum – Firmenbuch gut denkbar.',  # Molkerei Gropper
    'Q1627195': '150 Jahre laut Wikidata (1879). Logistik-Familienunternehmen in Neu-Ulm, gegenüber von Ulm. Lange Familiengeschichte – Firmenbuch möglich.',  # Honold Logistik Gruppe
    'Q2038417': '150 Jahre laut Wikidata (1879). Hersteller von Autowaschanlagen in Memmingen, Familienunternehmen (aus dem Gedächtnis). Gute Zahl, Firma mit Geld.',  # Otto Christ
    'Q21039847': '1250 Jahre (779), Ortsteil von Giebelstadt (Kreis Würzburg). Ansprechpartner wäre die Gemeinde Giebelstadt.',  # Herchsheim
    'Q505756': '900 Jahre (1129, ein Ritter Heinrich von Dahspach verkauft ein Gut an St. Michael in Bamberg). Gemeinde mit 1 800 Einwohnern im Aischgrund. Als Ortsnennung üblich, aber fragen, ob Dachsbach selbst ab 1129 zählt.',  # Dachsbach
    'Q506190': '800 Jahre (1229 „Arberch“), Gemeinde mit 2 200 Einwohnern im Kreis Ansbach.',  # Arberg
    'Q528727': '800 Jahre (1229 „Widenbach“), Gemeinde mit 2 500 Einwohnern im Kreis Ansbach.',  # Weidenbach
    'Q508471': '1250 Jahre (779), aber nur 1 100 Einwohner im Kreis Main-Spessart. Eher ein Festheft als ein Buch.',  # Gössenheim
    'Q17322305': '900 Jahre (1129), 1 400 Einwohner, Stadtteil von Ichenhausen (Kreis Günzburg). Ansprechpartner die Stadt.',  # Hochwang
    'Q127388879': '175 Jahre laut Wikidata (1854). Gymnasium in Neuendettelsau im Umfeld der Diakonie. Festschrift denkbar, kleines Budget.',  # Laurentius-Gymnasium
    'Q1526374': '175 Jahre laut Wikidata (1854). Unternehmen in Immenstadt im Allgäu (Klebebänder, aus dem Gedächtnis). Jahr und Eigentümer vor der Anfrage prüfen.',  # Monta
}

REIHENFOLGE = [
    'Q13896',  # Markgröningen
    'Q355435',  # Adelsheim
    'Q491535',  # Bad Wimpfen
    'Q74763163',  # Dürrmenz
    'Q61830',  # Großbottwar
    'Q23023',  # Waldsee
    'Q2586953',  # Witzenmann-Gruppe
    'Q17122614',  # Carl Roth
    'Q61928',  # Vaihingen an der Enz
    'Q80809',  # Denkendorf
    'Q43273976',  # Großsachsen
    'Q61653',  # Benningen am Neckar
    'Q61705',  # Ingersheim
    'Q515323',  # Weilheim
    'Q80606',  # Oberboihingen
    'Q80511',  # Wimsheim
    'Q20179999',  # Flözlingen
    'Q2016767',  # Ohrnberg
    'Q507274',  # Bermatingen
    'Q322938',  # Staatliche Akademie der Bildenden Künste Karlsruhe
    'Q19963065',  # Geschichtsverein Markgräflerland
    'Q16006232',  # Rupert-Neß-Gymnasium Wangen
    'Q520675',  # Erlebnispark Tripsdrill
    'Q798951',  # Bader Versand
    'Q1746794',  # Bankhaus E. Mayer
    'Q16002754',  # Amann & Söhne
    'Q1811343',  # Lechler (Unternehmen)
    'Q1485474',  # Wolffkran
    'Q1333317',  # ElringKlinger
    'Q19930480',  # Schützenverein Oberkirch
    'Q106394',  # SC Freiburg
    'Q47796855',  # Hochdorf/Enz
    'Q2284859',  # Siglingen
    'Q889245',  # Bockschaft
    'Q18019848',  # Eberfingen
    'Q1469016',  # Wangen
    'Q546393',  # Cleebronn
    'Q537465',  # Pfaffenhofen
    'Q19969752',  # Schomburg
    'Q540770',  # Unterstadion
    'Q632934',  # Lauterach
    'Q373283',  # Schönau
    'Q1308178',  # Zell
    'Q1925995',  # Mettingen (Esslingen am Neckar)
    'Q1738511',  # Kemnat
    'Q1566803',  # Haberschlacht
    'Q33510632',  # Unterjettingen
    'Q17577331',  # Leupolz
    'Q2350567',  # Stockach (Gomaringen)
    'Q2549872',  # Wart
    'Q122453',  # Universität Stuttgart
    'Q89491135',  # Elly-Heuss-Knapp-Gymnasium Stuttgart
    'Q1555507',  # Gustav-von-Schmoller-Schule Heilbronn
    'Q133932437',  # Johann-Jakob-Widmann-Schule
    'Q134068397',  # Werner-Siemens-Schule Stuttgart
    'Q15688375',  # Friedrich-Gymnasium Freiburg
    'Q139863925',  # Parler-Gymnasium Schwäbisch Gmünd
    'Q39080113',  # Landesärztekammer Baden-Württemberg
    'Q1250760',  # Geschichtsverein der Diözese Rottenburg-Stuttgart
    'Q2505471',  # VBU Volksbank im Unterland
    'Q17477739',  # Raiffeisenbank Kaiserstuhl
    'Q15134090',  # Volksbank Möckmühl
    'Q1890715',  # Mannheimer Versicherungen
    'Q1597044',  # Heinrich Dinkelacker
    'Q1774105',  # Klenk Holz
    'Q1131085',  # Coperion
    'Q125578831',  # Rheinkraftwerk Albbruck-Dogern AG
    'Q2231851',  # Schauburg
    'Q18023330',  # Leuco
    'Q1529839',  # Glatt Unternehmensgruppe
    'Q23826343',  # Liebherr-Werk Biberach
    'Q23788252',  # Liebherr-Hausgeräte
    'Q2044808',  # PTV
    'Q22909101',  # STRATEC
    'Q1023617',  # CHG-Meridian
    'Q1045931',  # Carthago Reisemobilbau
    'Q116895440',  # Sektion Ulm
    'Q39562085',  # Sektion Stuttgart
    'Q109769788',  # FG Narrhalla
    'Q120968418',  # ATSV Mutschelbach
    'Q319548',  # FC Rastatt 04
    'Q1388226',  # FG Rüppurr
    'Q2458740',  # TuS Hohenecken
    'Q325249',  # SVN Zweibrücken
    'Q176850',  # Bahlinger SC
    'Q162349',  # 1. FC Normannia Gmünd
    'Q319868',  # FC Singen 04
    'Q634084',  # Trebur
    'Q1998041',  # Nordheim
    'Q1874094',  # Ludwig-Georgs-Gymnasium
    'Q1665423',  # Viktoriaschule
    'Q1284907',  # Edith-Stein-Schule Darmstadt
    'Q17326366',  # Stadtbibliothek Darmstadt
    'Q1798208',  # Steinbuch
    'Q2582423',  # Wilmshausen
    'Q573628',  # Veitshöchheim
    'Q55523887',  # Kurtz Ersa
    'Q65179937',  # DIAKONEO
    'Q636150',  # Kist
    'Q536407',  # Kürnach
    'Q508445',  # Höchberg
    'Q1773361',  # Sparkasse Neu-Ulm - Illertissen
    'Q1943544',  # Molkerei Gropper
    'Q1627195',  # Honold Logistik Gruppe
    'Q2038417',  # Otto Christ
    'Q21039847',  # Herchsheim
    'Q505756',  # Dachsbach
    'Q506190',  # Arberg
    'Q528727',  # Weidenbach
    'Q508471',  # Gössenheim
    'Q17322305',  # Hochwang
    'Q127388879',  # Laurentius-Gymnasium
    'Q1526374',  # Monta
]
