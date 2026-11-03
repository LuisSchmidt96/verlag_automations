"""Durchsicht der Jubiläen 2028 von Hand (November 2026).

Stufen siehe leads.py. Kommentare mit „(aus dem Gedächtnis)“ vor einer Anfrage prüfen.
Zeilen mit Stufe C ohne eigenen Satz: in der Durchsicht gesehen und als Kleinkram verworfen;
der Grund steht kurz dabei.
"""

E = {
    'Q22901': ('S', 'Stadt, 1250 Jahre (Lorscher Codex 778), 19 000 Einwohner, Pfalz – Verlagsgebiet'),  # Bad Dürkheim
    'Q574314': ('S', 'Große Kreisstadt, 1250 Jahre, 23 000 Einwohner'),  # Laupheim
    'Q25430': ('S', '1000 Jahre, Kurort – Buch passt ins Kurmarketing'),  # Badenweiler
    'Q1457656': ('S', '250 Jahre, Familienunternehmen (Messer), Deizisau'),  # Friedr. Dick
    'Q656053': ('A', '1250 Jahre, 2 600 Einwohner, Kreis Germersheim – Pfalz'),  # Leimersheim
    'Q527910': ('A', '750 Jahre, 13 700 Einwohner, Bodenseekreis'),  # Meckenbeuren
    'Q536492': ('A', '1200 Jahre, Stadtteil von Kraichtal – Verlagsgebiet'),  # Münzesheim
    'Q80444': ('A', '800 Jahre, 5 000 Einwohner'),  # Heiningen
    'Q80717': ('A', '800 Jahre, 5 500 Einwohner'),  # Kuchen
    'Q811217': ('A', '1250 Jahre, Stadtteil von Bretten – Verlagsgebiet'),  # Bauerbach
    'Q2910': ('A', '175 Jahre Gemeinde Ludwigshafen (1853) – Großstadt, Jahr im Artikel nicht einheitlich'),  # Ludwigshafen am Rhein
    'Q609223': ('A', '1250 Jahre, Weinort an der Nahe – Nachbargebiet'),  # Monzingen
    'Q1784438': ('A', '1250 Jahre, Stadtteil von Kehl'),  # Kork
    'Q19279881': ('A', '1200 Jahre, Kernort der Stadt Rheinau'),  # Freistett
    'Q2568657': ('A', '750 Jahre Schule (Lateinschule) – Festschrift sehr wahrscheinlich'),  # Wieland-Gymnasium
    'Q52205046': ('A', '100 Jahre, Maschinenbau in Mühlacker – Verlagsgebiet'),  # Elumatec
    'Q761708': ('A', '150 Jahre Hochschule Karlsruhe – Festschrift wahrscheinlich'),  # Hochschule Karlsruhe
    'Q107661243': ('A', '300 Jahre Familienbrauerei'),  # Lammbrauerei Hilsenbeck
    'Q1412465': ('A', '200 Jahre Württembergische Versicherung – großes Haus, viel Geschichte'),  # Württembergische Versicherung
    'Q1744206': ('A', '100 Jahre, Familienunternehmen (Saunen)'),  # Klafs
    'Q1926291': ('A', '100 Jahre, Familienunternehmen (Wäsche)'),  # Mey
    'Q1787769': ('A', '175 Jahre Kreissparkasse – Festschrift typisch'),  # Kreissparkasse Biberach
    'Q1787801': ('A', '175 Jahre Kreissparkasse – Festschrift typisch'),  # Kreissparkasse Reutlingen
    'Q139585624': ('A', '300 Jahre Weingut, Rheinhessen – Nachbargebiet'),  # Weingut Dreissigacker
    'Q22674595': ('A', '150 Jahre, Familienunternehmen (Füllstoffe), weltweit tätig'),  # J. Rettenmaier & Söhne
    'Q1544329': ('A', '125 Jahre, Familienunternehmen (Süßwaren), Offenburg'),  # Vivil
    'Q565836': ('B', '1200 Jahre, aber nur 940 Einwohner – Pfalz'),  # Wilgartswiesen
    'Q566917': ('B', '750 Jahre, 625 Einwohner – Pfalz'),  # Battweiler
    'Q632790': ('B', '900 Jahre, 2 200 Einwohner'),  # Dogern
    'Q61715': ('B', '800 Jahre, 2 300 Einwohner'),  # Assamstadt
    'Q81107': ('B', '800 Jahre, 3 000 Einwohner'),  # Bühlertann
    'Q61645': ('B', '800 Jahre, 2 600 Einwohner'),  # Inzlingen
    'Q610873': ('B', '900 Jahre, 2 100 Einwohner, Rheinhessen'),  # Gau-Bickelheim
    'Q759220': ('B', '750 Jahre, Ortsteil von Karlsbad – Verlagsgebiet'),  # Auerbach
    'Q104528295': ('B', '750 Jahre, Ortsteil von Remchingen – Verlagsgebiet'),  # Darmsbach
    'Q1259455': ('B', '700 Jahre, Ortsteil von Pfinztal – Verlagsgebiet'),  # Kleinsteinbach
    'Q534139': ('B', '900 Jahre, Jahr unsicher'),  # Steinhausen an der Rottum
    'Q1267985': ('B', '1250 Jahre, Ortsteil von Bad Friedrichshall'),  # Duttenberg
    'Q23559654': ('B', '1250 Jahre, Jahr unsicher – Pfalz'),  # Gronau (Rödersheim-Gronau)
    'Q1815426': ('B', '1250 Jahre, kleiner Ortsteil von Geisingen'),  # Leipferdingen
    'Q1948864': ('B', '1200 Jahre, Ortsteil von Schaafheim, Südhessen'),  # Mosbach (Schaafheim)
    'Q892649': ('B', '900 Jahre, Ortsteil von Berkheim'),  # Bonlanden
    'Q1538589': ('B', '900 Jahre, Ortsteil von Bad Wurzach'),  # Gospoldshofen
    'Q1453319': ('B', '900 Jahre, Ortsteil von Essingen'),  # Lauterburg
    'Q851104': ('B', '800 Jahre, Ortsteil von Singen'),  # Beuren an der Aach
    'Q1498709': ('B', '800 Jahre, Ortsteil von Mulfingen'),  # Jagstberg
    'Q1892378': ('B', '800 Jahre, Ortsteil von Herbertingen'),  # Marbach
    'Q21162566': ('B', '750 Jahre, Ortsteil von Weinstadt'),  # Endersbach
    'Q1804975': ('B', '750 Jahre, Ortsteil von Schopfheim'),  # Langenau (Schopfheim)
    'Q1955890': ('B', '750 Jahre, Jahr unsicher, Karlsbad'),  # Mutschelbach
    'Q23787857': ('B', '100 Jahre, Kaiserslautern – Pfalz'),  # Torpedo-Gruppe
    'Q1823343': ('B', '800 Jahre, Ortsteil von Fischbachtal, Südhessen'),  # Lichtenberg
    'Q1633039': ('B', '150 Jahre Berufsschule, Wiesloch'),  # Hubert-Sternberg-Schule Wiesloch
    'Q127785718': ('B', '200 Jahre Gymnasium – Jahr uneinheitlich'),  # Albert-Schweitzer-Gymnasium
    'Q124412231': ('B', '150 Jahre Gewerbeschule'),  # Gewerbeschule Mudau
    'Q1453964': ('B', '100 Jahre Hochschule (Waldorf-Lehrerbildung)'),  # Freie Hochschule Stuttgart
    'Q2327042': ('B', '125 Jahre Stadtbücherei – Pfalz'),  # Stadtbücherei Zweibrücken
    'Q134075575': ('B', '200 Jahre Berufsschule'),  # Philipp-Matthäus-Hahn-Schule Nürtingen
    'Q141183014': ('B', '125 Jahre, Holzschutz, Sinzheim'),  # Dr. Wolman GmbH
    'Q2639562': ('B', '100 Jahre, Fruchtsäfte'),  # Albi
    'Q2693774': ('B', '100 Jahre, Elektrotechnik'),  # FRAKO
    'Q61597151': ('B', '100 Jahre, Medizintechnik'),  # Kauth Group
    'Q1629440': ('B', '100 Jahre, Medizintechnik'),  # Leiber Group
    'Q127920014': ('B', '100 Jahre, Getriebe'),  # Neugart Holding
    'Q2023445': ('B', '100 Jahre, Busunternehmen'),  # Omnibus Müller
    'Q18027734': ('B', '100 Jahre Schlosskellerei'),  # Schlosskellerei Affaltrach Dr. Reinhold Baumann
    'Q2244959': ('B', '100 Jahre Schluchseewerk – Energieversorger, Festschrift plausibel'),  # Schluchseewerk
    'Q1480075': ('B', '100 Jahre, Unternehmensgruppe'),  # Unternehmensgruppe Dr. Eckert
    'Q2541699': ('B', '100 Jahre, Leuchten'),  # Waldmann Group
    'Q15135139': ('B', '100 Jahre Winzergenossenschaft'),  # Winzergenossenschaft Waldulm
    'Q2597189': ('B', '100 Jahre Mineralquelle'),  # Wüteria Mineralquellen
    'Q2127570': ('B', '125 Jahre Raiffeisenbank – Fusion prüfen'),  # Raiffeisenbank im Breisgau
    'Q1483069': ('B', '250 Jahre Gymnasium'),  # Fürstenberg-Gymnasium Donaueschingen
    'Q2232368': ('B', '175 Jahre Gymnasium'),  # Scheffel-Gymnasium Bad Säckingen
    'Q126118233': ('B', '150 Jahre Gymnasium'),  # Gymnasium Kenzingen
    'Q15128001': ('B', '150 Jahre Gymnasium'),  # Schubart-Gymnasium Ulm
    'Q133880153': ('B', '100 Jahre Mineralwasser'),  # Dietenbronner Mineralwasser
    'Q457503': ('B', '200 Jahre, Familienunternehmen (Handel) – Saarland'),  # Globus Holding
    'Q15117664': ('B', '150 Jahre Kreissparkasse – Hessen'),  # Kreissparkasse Gelnhausen
    'Q451423': ('B', '175 Jahre WMF – Konzern (aus dem Gedächtnis: ausländischer Eigentümer)'),  # Württembergische Metallwarenfabrik
    'Q38063667': ('B', '150 Jahre Klinikum'),  # Klinikum Crailsheim
    'Q1021191': ('B', '125 Jahre Brauerei'),  # Bürgerliches Brauhaus Ravensburg-Lindau
    'Q1054584': ('B', '125 Jahre, Keramik – Konzern'),  # CeramTec
    'Q459029': ('B', '125 Jahre Burda – selbst Medienhaus, macht es eher selbst'),  # Hubert Burda Media Holding
    'Q2326871': ('B', '125 Jahre Stadtbibliothek'),  # Stadtbibliothek Heilbronn
    'Q124612865': ('B', '150 Jahre Brauerei, Wertheim'),  # Brauerei Georg Dehner Lindelbach
    'Q466029': ('B', '150 Jahre Konservatorium – Frankfurt'),  # Dr. Hoch’s Konservatorium
    'Q520174': ('B', '150 Jahre Brauerei – Saarland'),  # Karlsberg Holding
    'Q1751176': ('B', '150 Jahre Stadtbibliothek Worms'),  # Stadtbibliothek Worms
    'Q1942540': ('B', '125 Jahre, Gewürzwerk Mainz'),  # Moguntia (Gewürzwerk)
    'Q1875427': ('B', '150 Jahre Ruderverein – Pfalz'),  # Ludwigshafener Ruderverein
    'Q300187': ('B', '125 Jahre Ringerverein – Pfalz'),  # AV 03 Speyer
    'Q481377': ('B', '125 Jahre, Pfalz'),  # Arminia Ludwigshafen
    'Q700053': ('B', '125 Jahre, Pfalz'),  # FK Pirmasens
    'Q96998207': ('B', '125 Jahre, Ladenburg – Verlagsgebiet'),  # FV 03 Ladenburg
    'Q134580377': ('B', '150 Jahre Feuerwehr – Festschrift üblich'),  # Freiwillige Feuerwehr Freiburg im Breisgau
    'Q1965395': ('B', '125 Jahre Narrenzunft Rottweil – bekannteste Fasnet'),  # Narrenzunft Rottweil
    'Q47488922': ('B', '175 Jahre Narrenzunft – große Zahl'),  # Narrenzunft Frohsinn 1853 Donaueschingen
    'Q560270': ('C', '900 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Bisterschied
    'Q674420': ('C', '900 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Oberndorf
    'Q561699': ('C', '900 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Duchroth
    'Q680934': ('C', '900 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Traisen
    'Q561616': ('C', '1250 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Dintesheim
    'Q136327662': ('C', '200 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Lauerhof
    'Q20191080': ('C', '700 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Mülben
    'Q17577333': ('C', '800 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Levertsweiler
    'Q1348370': ('C', '750 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Erbisreute
    'Q56286159': ('C', '750 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Härtsfeldhausen
    'Q116780258': ('C', '750 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Sießen
    'Q136173585': ('C', '750 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Stöckenhof
    'Q130474046': ('C', '600 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Kaltenberg
    'Q134284184': ('C', '600 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Sirnitz
    'Q137125333': ('C', '600 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Ödernhardt
    'Q2009244': ('C', '500 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Oberbrüden
    'Q128062705': ('C', '500 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Stegmühle
    'Q334653': ('C', '900 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Abtweiler
    'Q1462024': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Staatliches Seminar für Didaktik und Lehrerbildung Karlsruhe
    'Q1777363': ('C', '600 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Knoden
    'Q105336876': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Jugendherberge Walldürn
    'Q692314': ('C', '750 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Bürchau
    'Q691684': ('C', '750 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Elbenschwand
    'Q133836071': ('C', '750 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Kühlenbronn
    'Q700088': ('C', '750 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Neuenweg
    'Q255489': ('C', '800 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Äpfingen
    'Q26271375': ('C', '750 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Friedingen
    'Q1825916': ('C', '700 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Lindach
    'Q1559332': ('C', '250 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Göllheimer Häuschen
    'Q19964929': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Unsere Volksbank St. Wendeler Land
    'Q29122974': ('C', 'japanische Konzerntochter'),  # IHI Corporation (Germany)
    'Q30272335': ('C', '125 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Grieshaber
    'Q1532037': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # HVB Wiest + Schürmann
    'Q137293970': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Maico
    'Q100989456': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Ulrich-von-Hutten-Gymnasium
    'Q15787245': ('C', '175 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Bayerische Brauerei Kaiserslautern
    'Q124738157': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Bausparkasse Deutsche Bau-Gemeinschaft
    'Q1697086': ('C', '125 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Johannes-Gutenberg-Schule
    'Q50039015': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Trautwein (Germany)
    'Q1391225': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Hochschule für Rechtspflege Schwetzingen
    'Q115161973': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Nikolaus-von-Weis-Gymnasium
    'Q2312852': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Sportschule Schöneck
    'Q27886770': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Volkshochschule im Landkreis Karlsruhe
    'Q124515033': ('C', '750 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Glasern
    'Q786168': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Autobus Sippel
    'Q28549952': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Bäckerei Pappert
    'Q1024572': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # CVJM-Kolleg
    'Q1182873': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Degeto Film
    'Q1341466': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Energieversorgung Mittelrhein
    'Q1596383': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Heinrich-Haus
    'Q10575468': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Maizena
    'Q108913752': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # OKAL
    'Q2092055': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Picard
    'Q2250672': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Schufa
    'Q25389080': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Zickenheiner
    'Q27525485': ('C', '250 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Ramselegut
    'Q1439043': ('C', '125 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Fortuna-Werke
    'Q2223917': ('C', '125 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Sanwald
    'Q100348461': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Orgelbau Wolfgang Scherpf
    'Q2099972': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Pneuhage
    'Q1429774': ('C', '175 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Albert-Schweitzer-Schule
    'Q2217674': ('C', '175 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Samson-Raphael-Hirsch-Schule
    'Q2571498': ('C', '175 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Wilhelm-Hofmann-Gymnasium
    'Q821960': ('C', '150 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Berlitz Sprachschulen
    'Q17321222': ('C', '125 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Diesterwegschule
    'Q189649': ('C', '125 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Zenner International
    'Q107088141': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Hessische Kommunale Gasfernversorgung
    'Q1492202': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Popakademie Baden-Württemberg
    'Q1545677': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Grenke AG
    'Q25698371': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Pfitzenmeier
    'Q112253052': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Schaller & Partner
    'Q22674693': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Sinus-Institut
    'Q1311508': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Sunrise Medical
    'Q2546896': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Walter services
    'Q134619919': ('C', '150 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Backhaus
    'Q113764059': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Ingenieurbüro Sack & Partner
    'Q6348521': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Kaiserslautern High School
    'Q112612619': ('C', '100 Jahre, kleiner Verein'),  # FC 1928 Asbach e.V.
    'Q112571921': ('C', '100 Jahre, kleiner Verein'),  # SV Rot-Weiß Waldhausen 1928 e.V.
    'Q18630954': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Sparkasse Ulm
    'Q56476970': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Freie Interkulturelle Waldorfschule Mannheim
    'Q47160231': ('C', '250 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Brohl Wellpappe
    'Q20828783': ('C', '125 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Marienhaus
    'Q125554406': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Boos-Druck
    'Q113885318': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # DE - Demke Electronic GmbH
    'Q6716393': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # MK-Motorsport
    'Q33127861': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Cosmopop
    'Q1278072': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Eurobaustoff
    'Q287560': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Gameforge
    'Q1924032': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Messe Karlsruhe
    'Q2112040': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # profine GmbH
    'Q474049': ('C', '75 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Amorbach
    'Q15854774': ('C', '75 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Windeck
    'Q88705764': ('C', '125 Jahre, kleiner Verein'),  # FC Germania Friedrichsfeld
    'Q18288274': ('C', '125 Jahre, kleiner Verein'),  # MSV Ludwigshafen
    'Q130532341': ('C', '125 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Palasthotel
    'Q2179171': ('C', '125 Jahre, kleiner Verein'),  # SpVgg Sandhofen
    'Q82557851': ('C', '125 Jahre, kleiner Verein'),  # TSV Lingenfeld
    'Q495467': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Andreae-Gymnasium Herrenberg
    'Q1262617': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Duale Hochschule Baden-Württemberg Ravensburg
    'Q1749624': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Otto-Hahn-Gymnasium
    'Q23305709': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Axonic Informationssysteme GmbH
    'Q109450250': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Chrono24
    'Q136425130': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Forum (Historisches Museum der Pfalz Speyer)
    'Q129717763': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Karlsruher Schieneninfrastruktur-Gesellschaft
    'Q138456224': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Pfalzwind
    'Q422600': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Aktionsgemeinschaft Soziale Marktwirtschaft
    'Q831958': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Betriebskrankenkasse Mahle
    'Q894366': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Bosch BKK
    'Q27313706': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # CHT Gruppe
    'Q138465560': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Hechinger Unternehmensgruppe
    'Q1653105': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # IBM Deutschland Research & Development
    'Q1882023': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # MULAG Fahrzeugwerk
    'Q1958759': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Müller
    'Q2109095': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Prettl
    'Q836847': ('C', '100 Jahre, kleiner Verein'),  # FC Denzlingen
    'Q15126760': ('C', '100 Jahre, kleiner Verein'),  # SV Oberachern
    'Q1023782': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # CJD Jugenddorf-Christophorusschule Oberurff
    'Q1565108': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Holzfachschule Bad Wildungen
    'Q56476199': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Freie Waldorfschule Esslingen
    'Q56476628': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Freie Waldorfschule Heilbronn
    'Q56477082': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Freie Waldorfschule Schopfheim
    'Q56477139': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Freie Waldorfschule Wahlwies
    'Q56477246': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Rudolf-Steiner-Schule Freie Waldorfschule Villingen-Schwenningen
    'Q135069144': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Gymnasium Remchingen
    'Q115518222': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # IB Berufliche Schulen Waiblingen
    'Q191426': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Zeppelin Universität
    'Q1575402': ('C', '75 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Hanfertal
    'Q1604590': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Helukabel
    'Q2339825': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Steinbach sprechende Bücher
    'Q118363805': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Vogt-Plastic
    'Q50842536': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Albrecht Kino Waldshut
    'Q30272278': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Endress Hauser Messtechnik (Germany)
    'Q105086530': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Carl-von-Ossietzky-Schule
    'Q1558197': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Gymnasium Nieder-Olm
    'Q1391239': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Hochschule Worms
    'Q1204498': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Deutsche WertpapierService Bank
    'Q1428203': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Vereinigte Volksbank Maingau
    'Q1463086': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Braas (Unternehmen)
    'Q22691786': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Enzianhütte
    'Q1615896': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Hessische Landeszentrale für politische Bildung
    'Q23788176': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Otto Heuss
    'Q1769435': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Willy Scharnow-Stiftung für Touristik
    'Q116257713': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Q116257713
    'Q100956184': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # MIMOT (Germany)
    'Q137273084': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Rale Engineering
    'Q131446172': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # herrmann+bosch architekten
    'Q131617947': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # ART Engineering
    'Q892895': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Bonus-Markt
    'Q24355693': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Evangelisches Bildungszentrum für Gesundheitsberufe Stuttgart
    'Q1798489': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Habila GmbH
    'Q2002862': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Novalung
    'Q1800693': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Rentschler Biopharma SE
    'Q17164356': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Studio Soi
    'Q1435798': ('C', '100 Jahre, kleiner Verein'),  # Fontana Finthen
    'Q97924152': ('C', '100 Jahre, kleiner Verein'),  # VfR Alsheim
    'Q56230270': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Freie Waldorfschule Darmstadt
    'Q1667895': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Internatsschule Schloss Hansenberg
    'Q1741300': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Landesmusikakademie Rheinland-Pfalz
    'Q873773': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Rhein-Main International Montessori School
    'Q4777995': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Anvis Group
    'Q15891084': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Berge & Meer Touristik GmbH
    'Q1254073': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Dr. Theiss Naturwaren
    'Q39099877': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Massa Haus
    'Q1965954': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Nassauische Verkehrs-GmbH
    'Q17102982': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Sauer Flugmotoren
    'Q20182979': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Stadtwerke Hanau
    'Q316899': ('C', '125 Jahre, kleiner Verein'),  # FC Emmendingen
    'Q1387514': ('C', '125 Jahre, kleiner Verein'),  # FC Radolfzell
    'Q951552': ('C', '125 Jahre, kleiner Verein'),  # Lahrer FV 03
    'Q870052': ('C', '125 Jahre, kleiner Verein'),  # SV 03 Tübingen
    'Q120914239': ('C', '125 Jahre, kleiner Verein'),  # TSV Wolfschlugen
    'Q1202811': ('C', '75 Jahre, kleiner Verein'),  # Deutsche Gesellschaft für Anästhesiologie und Intensivmedizin
    'Q129176543': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Hastamat Verpackungstechnik
    'Q133825969': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Rex
    'Q140333875': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Emergent Actio
    'Q22101806': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Skyway Products
    'Q50038983': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Ypsomed (Germany)
    'Q15056701': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Hochtaunusschule Oberursel
    'Q20819656': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Riegl Laser Measurement Systems
    'Q28196057': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Telespazio VEGA Deutschland
    'Q139504920': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Westermo Eltec
    'Q451851': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Amadeus FiRe
    'Q135010393': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Brain-Media.de
    'Q1022704': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # CABB
    'Q1152109': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # DB Cargo Deutschland
    'Q29017339': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Energie Waldeck-Frankenberg
    'Q19968197': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Gruenrekorder
    'Q1390846': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Hessische Staatsweingüter Kloster Eberbach
    'Q18628527': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # LevoBank
    'Q1123387': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Nexi Germany
    'Q2107022': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Powerland
    'Q1359446': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Sell GmbH
    'Q24395214': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Treofan
    'Q26794585': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # United Power Technology
    'Q2511648': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Vectus Verkehrsgesellschaft
    'Q2516289': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Verkehrsgesellschaft Mittelhessen
    'Q985222': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Volkswagen Original Teile Logistik
    'Q114976625': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # IB Berufliche Schulen Eislingen
    'Q47498669': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenzunft Germania Sasbachwalden 1928
    'Q122573850': ('C', '150 Jahre, kleiner Verein'),  # Oberhessischer Geschichtsverein
    'Q135714581': ('C', '150 Jahre, kleiner Verein'),  # TV Homburg
    'Q157463': ('C', '125 Jahre, kleiner Verein'),  # 1. FC Saarbrücken
    'Q162391': ('C', '125 Jahre, kleiner Verein'),  # 1. FC Sobernheim
    'Q602922': ('C', '125 Jahre, kleiner Verein'),  # FVgg Mombach 03
    'Q442066': ('C', '125 Jahre, kleiner Verein'),  # Germania Wiesbaden
    'Q469346': ('C', '125 Jahre, kleiner Verein'),  # Spvgg. 03 Neu-Isenburg
    'Q47498654': ('C', '50 Jahre, kleiner Verein'),  # Niesatzer Hurzle
    'Q112583527': ('C', '50 Jahre, kleiner Verein'),  # TSV Hollerbach
    'Q481266': ('C', '25 Jahre, kleiner Verein'),  # Amtsrichterverband
    'Q107517906': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Autarcon GmbH
    'Q113385135': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Coreo AG
    'Q2147493': ('C', '25 Jahre, kleiner Verein'),  # Mannheim Bandits
    'Q21121072': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Samsung Electronics Holding
    'Q46644435': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # To Europe
    'Q140770893': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # print-ID GmbH & Co. KG
    'Q107170713': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Schramm Lack- und Farbenfabriken
    'Q470917': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Synovate
    'Q134963458': ('C', '25 Jahre, kleiner Verein'),  # Förderverein Psychosoziale Notfallversorgung im Neckar-Odenwald Kreis e.V.
    'Q111026315': ('C', '25 Jahre, kleiner Verein'),  # Rotary Club Neckar-Odenwald-Kreis
    'Q113253240': ('C', '25 Jahre, kleiner Verein'),  # SK Springer Binau 2003
    'Q51485830': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenzunft Kirrlach
    'Q47498662': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenzunft Otterschwierer Leimewängscht 1953
    'Q1518526': ('C', '50 Jahre, kleiner Verein'),  # Gesellschaft für Technische Kommunikation
    'Q1797145': ('C', '50 Jahre, kleiner Verein'),  # Künstlerhaus Stuttgart
    'Q2011478': ('C', '50 Jahre, kleiner Verein'),  # Oberwelt
    'Q1722266': ('C', '50 Jahre, kleiner Verein'),  # Ohne Rüstung Leben
    'Q47525625': ('C', '75 Jahre, kleiner Verein'),  # Freiburger Westhansele
    'Q48922029': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenzunft Salzstetten
    'Q112583630': ('C', '50 Jahre, kleiner Verein'),  # Tennisclub Reinhardsachsen 1978 e.V.
    'Q111609277': ('C', '50 Jahre, kleiner Verein'),  # Tennisclub Waldstadt e.V.
    'Q13487149': ('C', '75 Jahre, kleiner Verein'),  # 1. Wiesbadener BC
    'Q1205529': ('C', '75 Jahre, kleiner Verein'),  # Deutsches Aktieninstitut
    'Q25894736': ('C', '75 Jahre, kleiner Verein'),  # Gesellschaft für Arzneipflanzen- und Naturstoff-Forschung
    'Q15823229': ('C', '75 Jahre, kleiner Verein'),  # Katholisches Filmwerk
    'Q1802759': ('C', '75 Jahre, kleiner Verein'),  # Landeszahnärztekammer Hessen
    'Q1270251': ('C', '75 Jahre, kleiner Verein'),  # Marburger Kunstverein
    'Q1321086': ('C', '75 Jahre, kleiner Verein'),  # TSV Schott Mainz
    'Q135864109': ('C', '75 Jahre, kleiner Verein'),  # TTC Grün-Weiß Staffel
    'Q51717097': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Burghexen Schwandorf
    'Q47525577': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Gigili-Geister Freiburg-Munzingen
    'Q635268': ('C', '25 Jahre, kleiner Verein'),  # Architekturbild
    'Q1203570': ('C', '25 Jahre, kleiner Verein'),  # Deutsche Model United Nations
    'Q1250598': ('C', '25 Jahre, kleiner Verein'),  # Projekt Chance
    'Q50319835': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenzunft Waldhufenstromer Schömberg
    'Q1620847': ('C', '50 Jahre, kleiner Verein'),  # Historische Eisenbahn Frankfurt
    'Q50090876': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Brandweinhexen Kohlstetten 2003
    'Q162386': ('C', '25 Jahre, kleiner Verein'),  # 1. FC Schwalmstadt
    'Q141807': ('C', '25 Jahre, kleiner Verein'),  # Deutsch-Pennsylvanischer Arbeitskreis
    'Q47498637': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenzunft Backstein und Matrazenbourg 1953
    'Q48975845': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenzunft Hungrige Stühlinger
    'Q48976231': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenzunft Karsau
    'Q48976257': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenzunft Murg
    'Q48976013': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Räuberzunft Gündelwangen
    'Q47525574': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Zunft der Feuer-Narre
    'Q48977643': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenverein Stetten unter Holstein 1978 Rauchkatzen
    'Q47486469': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenvereinigung Kleggau
    'Q47516461': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenzunft Dauchingen
    'Q47498650': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenzunft Höllwaldteufel Lautenbach
    'Q48986464': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenzunft Sattelbach City
    'Q48817344': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenzunft Wurmlinger Knöpfle
    'Q47516500': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Schanzelzunft
    'Q110484110': ('C', '25 Jahre, kleiner Verein'),  # Hessischer Literaturrat
    'Q110736693': ('C', '25 Jahre, kleiner Verein'),  # Psychotherapeutenkammer des Saarlandes
    'Q136658087': ('C', '25 Jahre, kleiner Verein'),  # Verband hessischer Kommunalarchivarinnen und -archivare e.V.
    'Q47519978': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenzunft Rheinwaldhexen Oberhausen 1953
    'Q48704031': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrnezunft Wasaschomberler Einsingen
    'Q49281029': ('X', 'Doppelter Eintrag, gleicher Ort wie Gemeinde Assamstadt'),  # Assamstadt
    'Q107138851': ('X', 'historisches Unternehmen'),  # Ettlingen-Maxau Papier- und Zellstoffwerke
    'Q107139023': ('X', 'historisches Unternehmen'),  # Holzhydrolyse
    'Q104889293': ('X', 'historisches Unternehmen'),  # Vereinigte Korkindustrie AG
    'Q130283594': ('X', 'Gebäude'),  # Neulinger Straße 2, Ölbronn
    'Q134121876': ('X', 'Gebäude'),  # Schlierbacher Landstraße 136
    'Q67206361': ('X', 'in der Württembergischen Versicherung aufgegangen'),  # Württembergische Feuerversicherung
    'Q1676321': ('X', 'Verlag – macht es selbst'),  # J. G. Oncken Nachf.
    'Q871620': ('X', 'historisches Unternehmen'),  # Flugzeugbau Max Gerner
    'Q1396951': ('X', 'historisches Unternehmen'),  # Albert Nestler Zeichentechnik
    'Q1285325': ('X', 'Verlag – macht es selbst'),  # Edition Steingräber
    'Q107538935': ('X', 'historisches Unternehmen'),  # Heddernheimer Kupferwerke
    'Q135240630': ('X', 'historisches Kaufhaus'),  # Kaufhaus M. Levy
    'Q107718143': ('X', 'historisches Unternehmen'),  # Adolph Fürst & Sohn
    'Q107695489': ('X', 'historisches Unternehmen'),  # Höchster Gießereien L. Scriba
    'Q110549553': ('X', 'historische Schule, in Akademie aufgegangen'),  # Badische Kunstgewerbeschule Karlsruhe
    'Q47497604': ('X', 'Gebäude'),  # Dreiseithof
    'Q47493722': ('X', 'Gebäude'),  # Zweiseithof
    'Q558781': ('X', 'Verlag – macht es selbst'),  # Reclam-Verlag
    'Q111389980': ('X', 'Gebäude'),  # Bayer-Hochhaus
    'Q1468453': ('X', 'Verlag – macht es selbst'),  # S. Hirzel Verlag
    'Q1372591': ('X', 'Verlag – macht es selbst'),  # Eugen Heinz Druck- und Verlagsgesellschaft
    'Q1662434': ('X', 'Verlag – macht es selbst'),  # Info Verlag
    'Q591905': ('X', '1968 geschlossen'),  # Hochschule für Gestaltung Ulm
    'Q119626855': ('X', 'Verlag – macht es selbst'),  # Juventa Verlag
    'Q102210843': ('X', 'Verlag – macht es selbst'),  # Draupadi Verlag
    'Q873700': ('X', 'Verlag – macht es selbst'),  # Pons-Verlag
    'Q61467153': ('X', 'Verlag – macht es selbst'),  # Gunter Narr Verlag
    'Q33093382': ('X', 'Verlag – macht es selbst'),  # Pop Verlag
    'Q54380029': ('X', 'Verlag – macht es selbst'),  # R.G. Fischer Verlag
    'Q1258558': ('X', 'Verlag – macht es selbst'),  # Trotzdem Verlag
    'Q28823623': ('X', 'Verlag – macht es selbst'),  # Wernersche Verlagsgesellschaft
    'Q19971795': ('X', 'Verlag – macht es selbst'),  # Marix Verlag
    # Bayerische Grenzkreise (Nachtrag November 2026)
    'Q14825': ('S', 'Stadt, 900 Jahre (1128), 44 000 Einwohner – kreisfreie Stadt direkt an der Grenze zu BW'),  # Memmingen
    'Q32066': ('A', '900 Jahre (1128, Gründung Kloster Oberzell), 4 300 Einwohner, bei Würzburg'),  # Zell am Main
    'Q889673': ('A', '500 Jahre Schule (1528 laut Wikidata) – Festschrift sehr wahrscheinlich'),  # Bodensee-Gymnasium Lindau
    'Q510813': ('B', '800 Jahre, 940 Einwohner; Ortsteil Irsingen noch älter (1053)'),  # Gerolfingen
    'Q124409312': ('B', '125 Jahre Genossenschaftsbank – Datum eines Vorgängers, Bank aus Fusion'),  # VR SüdBank
    'Q2794077': ('B', '100 Jahre, Keramikhersteller am Untermain (aus dem Gedächtnis)'),  # Scheurich
    'Q1879680': ('B', '125 Jahre, Schulzentrum in Neuendettelsau – Jahr nur aus Wikidata'),  # Löhe-Campus
    'Q508851': ('C', '50 Jahre Gemeinde aus der Gebietsreform – Festheft'),  # Triefenstein
    'Q1566727': ('C', '900 Jahre, aber 150 Einwohner, Jahr nur aus Wikidata'),  # Habelsee
    'Q1373858': ('C', '900 Jahre, 270 Einwohner, Jahr nur aus Wikidata'),  # Schweinsdorf
    'Q2146502': ('C', '750 Jahre, 59 Einwohner'),  # Retzendorf
    'Q106515781': ('C', '750 Jahre, Weiler'),  # Beuren
    'Q1331976': ('C', '750 Jahre, kleiner Ortsteil von Holzheim'),  # Ellerbach
    'Q106368936': ('C', '750 Jahre, Weiler'),  # Oberthalhofen
    'Q109046748': ('C', '600 Jahre, Weiler am Bodensee'),  # Ebnet
    'Q131479550': ('C', 'Alpe, kein Ort'),  # Gaisalpe
    'Q20826924': ('C', 'Einöde'),  # Schachen
    'Q114877628': ('C', '300 Jahre, Weiler'),  # Eschachthal
    'Q105795043': ('C', 'Mühle, 2 Einwohner'),  # Zellesmühle
    'Q23559617': ('C', 'bei Rosenheim – außerhalb des Gebiets'),  # Aising
    'Q574736': ('C', '50 Jahre, Yachtwerft – kleine Zahl'),  # Bavaria Yachtbau
    'Q2008131': ('X', 'deutsche Tochter eines US-Konzerns'),  # OSI Food Solutions Germany
    'Q124462789': ('C', '75 Jahre, wenig bekannt'),  # EZV
    'Q83491915': ('C', '50 Jahre – kleine Zahl'),  # Regner
    'Q300805': ('C', '25 Jahre'),  # Aaglander
    'Q141622758': ('C', '25 Jahre'),  # AOT GmbH
    'Q1920014': ('C', '25 Jahre – Neugründung, die Brautradition ist älter'),  # Memminger Brauerei
    'Q125150784': ('C', '100 Jahre, Ruderverein – Festheft'),  # Ruder-Club Karlstadt 1928
    'Q112874283': ('C', '50 Jahre, Sportverein'),  # SG 1978 Waldbrunn
    'Q49022034': ('C', 'Narrenverein – Festheft, kein Buch'),  # Weißensberger Narrenverein
}

KOMMENTAR = {
    'Q22901': '1250 Jahre seit der ersten Erwähnung im Lorscher Codex (1. Juni 778 als „Turnesheim“). Kreisstadt mit 19 000 Einwohnern, bekannt durch Wurstmarkt und Weinbau – eine Stadt, die feiert. Pfalz, also mitten im Verlagsgebiet. Bester Ort-Lead 2028.',  # Bad Dürkheim
    'Q574314': '1250 Jahre (778 „Loubhaim“). Große Kreisstadt mit 23 000 Einwohnern im Kreis Biberach. Größe und Zahl stimmen – ein Stadtbuch ist realistisch.',  # Laupheim
    'Q25430': '1000 Jahre Badenweiler (1028, Kaiser Konrad II.). Kurort im Markgräflerland, 4 700 Einwohner. Römerbad und Kurgeschichte – ein schöner Stoff, und ein Kurort hat Interesse am Marketing. Die Ortsteile sind noch älter (774).',  # Badenweiler
    'Q1457656': '250 Jahre, Familienunternehmen für Messer und Feilen in Deizisau (Kreis Esslingen). Eine so lange Firmengeschichte mit Familie dahinter ist der klassische Fall für ein Firmenbuch.',  # Friedr. Dick
    'Q656053': '1250 Jahre (778, Schenkung an das Kloster Lorsch). Gemeinde mit 2 600 Einwohnern am Rhein, Kreis Germersheim – Pfalz und nah am Verlag. Kleiner Haushalt, aber eine starke Zahl.',  # Leimersheim
    'Q527910': '750 Jahre (1278 „Meckenburron“, laut Artikel gesichert). Gemeinde mit 13 700 Einwohnern im Bodenseekreis. Groß genug für ein Buch.',  # Meckenbeuren
    'Q536492': '1200 Jahre (Lorscher Codex 828). Stadtteil von Kraichtal mit 2 700 Einwohnern, Landkreis Karlsruhe – direkt im Verlagsgebiet. Typischer Fall für eine Ortschronik.',  # Münzesheim
    'Q80444': '800 Jahre (1228, Urkunde in St. Gallen). Gemeinde mit 5 000 Einwohnern im Kreis Göppingen.',  # Heiningen
    'Q80717': '800 Jahre (1228 „Cuchin“). Gemeinde mit 5 500 Einwohnern im Kreis Göppingen, Nachbar von Heiningen – beide feiern 2028.',  # Kuchen
    'Q811217': '1250 Jahre (Lorscher Codex, 26. April 778). Stadtteil von Bretten, Verlagsgebiet. Stand schon in der Liste für 2029 – das Jubiläum ist 2028.',  # Bauerbach
    'Q2910': '175 Jahre Gemeinde Ludwigshafen (1853). Großstadt mit 177 000 Einwohnern in der Pfalz. Achtung: Ein anderer Artikel nennt 1843 – vorher klären, was die Stadt feiert. Wenn sie feiert, ist es ein großer Auftrag.',  # Ludwigshafen am Rhein
    'Q1784438': '1250 Jahre (778). Kork ist ein Stadtteil von Kehl – Ortsverwaltung und Stadt ansprechen.',  # Kork
    'Q19279881': '1200 Jahre (828 „Fregistat“). Freistett ist der Kernort der Stadt Rheinau (Ortenaukreis) – die Stadt dürfte mitfeiern.',  # Freistett
    'Q2568657': '750 Jahre Wieland-Gymnasium Biberach (Lateinschule seit 1278). Ein außergewöhnliches Schuljubiläum, Festschrift sehr wahrscheinlich.',  # Wieland-Gymnasium
    'Q52205046': '100 Jahre, Maschinenbau für Fenster- und Fassadenbau in Mühlacker – Enzkreis, Verlagsgebiet.',  # Elumatec
    'Q761708': '150 Jahre Hochschule Karlsruhe (1878). Große Hochschule in Karlsruhe, nah am Verlag. Eine Festschrift ist wahrscheinlich, vielleicht mit eigenem Team – dann als Verlag dazukommen.',  # Hochschule Karlsruhe
    'Q107661243': '300 Jahre Lammbrauerei Hilsenbeck in Gruibingen (Kreis Göppingen). Familienbrauerei mit sehr langer Geschichte.',  # Lammbrauerei Hilsenbeck
    'Q1412465': '200 Jahre Württembergische Versicherung (Stuttgart). Großes Haus mit langer Geschichte – Geld ist da, aber vermutlich auch eine eigene Kommunikation.',  # Württembergische Versicherung
    'Q1744206': '100 Jahre Klafs, Familienunternehmen (Saunen und Wellness) in Schwäbisch Hall. Bekannte Marke.',  # Klafs
    'Q1926291': '100 Jahre Mey, Familienunternehmen (Wäsche) in Albstadt.',  # Mey
    'Q1787769': '175 Jahre Kreissparkasse Biberach. Sparkassen machen zu runden Jubiläen gern eine Festschrift, oft mit Regionalgeschichte.',  # Kreissparkasse Biberach
    'Q1787801': '175 Jahre Kreissparkasse Reutlingen – wie Biberach.',  # Kreissparkasse Reutlingen
    'Q22674595': '150 Jahre J. Rettenmaier & Söhne in Rosenberg (Ostalbkreis), Familienunternehmen und weltweit tätig.',  # J. Rettenmaier & Söhne
    'Q1544329': '125 Jahre Vivil in Offenburg, Familienunternehmen (Pastillen). Bekannte Marke.',  # Vivil
    'Q609223': '1250 Jahre (778, Schenkung von Weinbergen an Lorsch). Weinort an der Nahe mit 1 600 Einwohnern. Kleiner Ort, aber eine starke Zahl. Nachbargebiet.',  # Monzingen
    'Q139585624': '300 Jahre Weingut Dreissigacker in Bechtheim (Rheinhessen). Ein bekanntes Weingut mit einer solchen Zahl – ein schönes Buch über Familie und Wein. Nachbargebiet.',  # Weingut Dreissigacker
    'Q565836': '1200 Jahre (828, Kloster Hornbach). Nur 940 Einwohner im Pfälzerwald – starke Zahl, kleiner Haushalt.',  # Wilgartswiesen
    'Q566917': '750 Jahre, 625 Einwohner, Südwestpfalz. Eher ein Festheft.',  # Battweiler
    'Q632790': '900 Jahre (1128), 2 200 Einwohner am Hochrhein.',  # Dogern
    'Q61715': '800 Jahre (1228), 2 300 Einwohner, Main-Tauber-Kreis.',  # Assamstadt
    'Q81107': '800 Jahre (1228), 3 000 Einwohner, Kreis Schwäbisch Hall.',  # Bühlertann
    'Q61645': '800 Jahre (1228), 2 600 Einwohner, Kreis Lörrach.',  # Inzlingen
    'Q759220': '750 Jahre (1278), Ortsteil von Karlsbad, Landkreis Karlsruhe. Stand schon in der Liste 2029 – das Jubiläum ist 2028.',  # Auerbach
    'Q104528295': '750 Jahre (1278), Ortsteil von Remchingen, Enzkreis.',  # Darmsbach
    'Q1259455': '700 Jahre (1328), Ortsteil von Pfinztal, Landkreis Karlsruhe.',  # Kleinsteinbach
    'Q534139': '900 Jahre, aber das Jahr ist unsicher. Gemeinde mit 2 200 Einwohnern.',  # Steinhausen an der Rottum
    'Q1267985': '1250 Jahre (778), Ortsteil von Bad Friedrichshall.',  # Duttenberg
    'Q23559654': '1250 Jahre, aber unsicher. Ortsteil von Rödersheim-Gronau, Pfalz.',  # Gronau (Rödersheim-Gronau)
    'Q1815426': '1250 Jahre (778, St. Gallen), kleiner Ortsteil von Geisingen.',  # Leipferdingen
    'Q892649': '900 Jahre (1128), Ortsteil von Berkheim.',  # Bonlanden
    'Q1538589': '900 Jahre (1128), Ortsteil von Bad Wurzach.',  # Gospoldshofen
    'Q1453319': '900 Jahre (1128), Ortsteil von Essingen.',  # Lauterburg
    'Q851104': '800 Jahre (1228), Ortsteil von Singen.',  # Beuren an der Aach
    'Q1498709': '800 Jahre (1228), Ortsteil von Mulfingen.',  # Jagstberg
    'Q1892378': '800 Jahre (1228), Ortsteil von Herbertingen.',  # Marbach
    'Q21162566': '750 Jahre (1278), Ortsteil von Weinstadt – Weinort im Remstal.',  # Endersbach
    'Q1804975': '750 Jahre (1278), Ortsteil von Schopfheim.',  # Langenau (Schopfheim)
    'Q1955890': '750 Jahre, Jahr nur „zurückverfolgt“, Ortsteil von Karlsbad.',  # Mutschelbach
    'Q23787857': '100 Jahre Torpedo-Gruppe in Kaiserslautern (Pfalz).',  # Torpedo-Gruppe
    'Q1633039': '150 Jahre Berufsschule in Wiesloch.',  # Hubert-Sternberg-Schule Wiesloch
    'Q127785718': '200 Jahre Gymnasium Gernsbach – Gründungsjahr uneinheitlich, vorher klären.',  # Albert-Schweitzer-Gymnasium
    'Q124412231': '150 Jahre Gewerbeschule Mudau, Neckar-Odenwald-Kreis.',  # Gewerbeschule Mudau
    'Q1453964': '100 Jahre Freie Hochschule Stuttgart (Ausbildung von Waldorflehrern).',  # Freie Hochschule Stuttgart
    'Q2327042': '125 Jahre Stadtbücherei Zweibrücken (Pfalz).',  # Stadtbücherei Zweibrücken
    'Q134075575': '200 Jahre Berufsschule in Nürtingen.',  # Philipp-Matthäus-Hahn-Schule Nürtingen
    'Q141183014': '125 Jahre Dr. Wolman (Holzschutz) in Sinzheim.',  # Dr. Wolman GmbH
    'Q2639562': '100 Jahre Albi (Fruchtsäfte) in Berghülen.',  # Albi
    'Q2693774': '100 Jahre FRAKO (Kondensatoren) in Teningen.',  # FRAKO
    'Q61597151': '100 Jahre Kauth (Medizintechnik) in Denkingen.',  # Kauth Group
    'Q1629440': '100 Jahre Leiber (Medizintechnik) in Emmingen-Liptingen.',  # Leiber Group
    'Q127920014': '100 Jahre Neugart (Getriebe) in Kippenheim.',  # Neugart Holding
    'Q2023445': '100 Jahre Omnibus Müller in Schwäbisch Hall.',  # Omnibus Müller
    'Q18027734': '100 Jahre Schlosskellerei Affaltrach.',  # Schlosskellerei Affaltrach Dr. Reinhold Baumann
    'Q2244959': '100 Jahre Schluchseewerk. Energieversorger machen gern Festschriften.',  # Schluchseewerk
    'Q1480075': '100 Jahre Unternehmensgruppe Dr. Eckert in Stuttgart.',  # Unternehmensgruppe Dr. Eckert
    'Q2541699': '100 Jahre Waldmann (Leuchten) in Villingen-Schwenningen.',  # Waldmann Group
    'Q15135139': '100 Jahre Winzergenossenschaft Waldulm.',  # Winzergenossenschaft Waldulm
    'Q2597189': '100 Jahre Wüteria Mineralquellen in Gemmingen.',  # Wüteria Mineralquellen
    'Q2127570': '125 Jahre Raiffeisenbank im Breisgau. Fusion prüfen.',  # Raiffeisenbank im Breisgau
    'Q1483069': '250 Jahre Fürstenberg-Gymnasium Donaueschingen.',  # Fürstenberg-Gymnasium Donaueschingen
    'Q2232368': '175 Jahre Scheffel-Gymnasium Bad Säckingen.',  # Scheffel-Gymnasium Bad Säckingen
    'Q126118233': '150 Jahre Gymnasium Kenzingen.',  # Gymnasium Kenzingen
    'Q15128001': '150 Jahre Schubart-Gymnasium Ulm.',  # Schubart-Gymnasium Ulm
    'Q133880153': '100 Jahre Dietenbronner Mineralwasser.',  # Dietenbronner Mineralwasser
    'Q451423': '175 Jahre WMF in Geislingen. Großer Name, gehört aber meines Wissens einem ausländischen Konzern (aus dem Gedächtnis).',  # Württembergische Metallwarenfabrik
    'Q38063667': '150 Jahre Klinikum Crailsheim.',  # Klinikum Crailsheim
    'Q1021191': '125 Jahre Bürgerliches Brauhaus Ravensburg-Lindau.',  # Bürgerliches Brauhaus Ravensburg-Lindau
    'Q1054584': '125 Jahre CeramTec in Plochingen – Konzern.',  # CeramTec
    'Q459029': '125 Jahre Burda. Selbst ein Medienhaus, macht das Buch vermutlich selbst.',  # Hubert Burda Media Holding
    'Q2326871': '125 Jahre Stadtbibliothek Heilbronn.',  # Stadtbibliothek Heilbronn
    'Q124612865': '150 Jahre Brauerei Dehner in Wertheim-Lindelbach.',  # Brauerei Georg Dehner Lindelbach
    'Q1875427': '150 Jahre Ludwigshafener Ruderverein (Pfalz).',  # Ludwigshafener Ruderverein
    'Q300187': '125 Jahre AV 03 Speyer (Ringen), Pfalz.',  # AV 03 Speyer
    'Q481377': '125 Jahre Arminia Ludwigshafen (Pfalz).',  # Arminia Ludwigshafen
    'Q700053': '125 Jahre FK Pirmasens (Pfalz).',  # FK Pirmasens
    'Q96998207': '125 Jahre FV 03 Ladenburg – Verlagsgebiet.',  # FV 03 Ladenburg
    'Q134580377': '150 Jahre Freiwillige Feuerwehr Freiburg. Feuerwehren machen zum Jubiläum fast immer eine Festschrift.',  # Freiwillige Feuerwehr Freiburg im Breisgau
    'Q1965395': '125 Jahre Narrenzunft Rottweil – die bekannteste Fasnet im Land, ein bildstarkes Thema.',  # Narrenzunft Rottweil
    'Q47488922': '175 Jahre Narrenzunft Frohsinn Donaueschingen – für eine Zunft eine große Zahl.',  # Narrenzunft Frohsinn 1853 Donaueschingen
    'Q610873': '900 Jahre (1128), 2 100 Einwohner, Rheinhessen. Nachbargebiet.',  # Gau-Bickelheim
    'Q1948864': '1200 Jahre (828, Einhard). Ortsteil von Schaafheim in Südhessen. Nachbargebiet.',  # Mosbach (Schaafheim)
    'Q1823343': '800 Jahre (1228), Ortsteil von Fischbachtal in Südhessen. Nachbargebiet.',  # Lichtenberg
    'Q1751176': '150 Jahre Stadtbibliothek Worms. Nachbargebiet.',  # Stadtbibliothek Worms
    'Q1942540': '125 Jahre Moguntia (Gewürze) in Mainz. Nachbargebiet.',  # Moguntia (Gewürzwerk)
    # Bayerische Grenzkreise
    'Q14825': '900 Jahre seit der ersten Erwähnung als „Mammingin“ (1128, Urkunde aus dem Kloster Ochsenhausen). Kreisfreie Stadt mit 44 000 Einwohnern an der Iller, gegenüber dem Kreis Biberach. Ehemalige Reichsstadt mit großer Geschichte (Zwölf Artikel 1525) – ein Stadtjubiläum dieser Größe bekommt fast sicher ein Buch; die Frage ist eher, ob schon ein Verlag gebunden ist. Bester Lead in den bayerischen Grenzkreisen 2028.',  # Memmingen
    'Q32066': '900 Jahre (1128, erste Erwähnung bei der Gründung des Klosters Oberzell). Gemeinde mit 4 300 Einwohnern direkt vor Würzburg. Kloster und Ort teilen sich das Jubiläum – guter Stoff für ein Ortsbuch.',  # Zell am Main
    'Q889673': '500 Jahre laut Wikidata (Gründung 1528). Gymnasium in Lindau. Schulen begehen so eine Zahl fast immer mit Festschrift – das Jahr vor der Anfrage auf der Schulseite prüfen, es stammt nur aus Wikidata.',  # Bodensee-Gymnasium Lindau
    'Q510813': '800 Jahre Gerolfingen (1228 „Gerlifingen“), 940 Einwohner im Kreis Ansbach. Klein; die Ortsteile Irsingen (1053) und Aufkirchen (1188) haben eigene, ältere Daten. Eher ein Festheft, aber anfragen, wenn man ohnehin in der Gegend ist.',  # Gerolfingen
    'Q124409312': '125 Jahre nach dem Gründungsdatum in der Infobox (1. April 1903). Genossenschaftsbank mit Sitz in Lindau; das Datum gehört zu einem Vorgängerinstitut, die heutige Bank entstand durch Fusion. Banken leisten sich Festschriften – aber erst klären, welches Jubiläum sie selbst zählen.',  # VR SüdBank
    'Q2794077': '100 Jahre laut Wikidata (1928). Keramikhersteller am Untermain (Kreis Miltenberg). Firmenbuch möglich; Jahr und heutige Lage des Unternehmens vor der Anfrage prüfen.',  # Scheurich
    'Q1879680': '125 Jahre laut Wikidata (1903). Schulzentrum in Neuendettelsau (Kreis Ansbach) im Umfeld der Diakonie. Festschrift denkbar, aber kein großes Budget.',  # Löhe-Campus
}

REIHENFOLGE = [
    'Q22901',  # Bad Dürkheim
    'Q574314',  # Laupheim
    'Q25430',  # Badenweiler
    'Q1457656',  # Friedr. Dick
    'Q536492',  # Münzesheim
    'Q811217',  # Bauerbach
    'Q656053',  # Leimersheim
    'Q2910',  # Ludwigshafen am Rhein
    'Q527910',  # Meckenbeuren
    'Q19279881',  # Freistett
    'Q1784438',  # Kork
    'Q80444',  # Heiningen
    'Q80717',  # Kuchen
    'Q2568657',  # Wieland-Gymnasium
    'Q761708',  # Hochschule Karlsruhe
    'Q1787769',  # Kreissparkasse Biberach
    'Q1787801',  # Kreissparkasse Reutlingen
    'Q1412465',  # Württembergische Versicherung
    'Q107661243',  # Lammbrauerei Hilsenbeck
    'Q22674595',  # J. Rettenmaier & Söhne
    'Q1744206',  # Klafs
    'Q1926291',  # Mey
    'Q1544329',  # Vivil
    'Q52205046',  # Elumatec
    'Q609223',  # Monzingen
    'Q139585624',  # Weingut Dreissigacker
    'Q14825',  # Memmingen
    'Q32066',  # Zell am Main
    'Q889673',  # Bodensee-Gymnasium Lindau
    'Q510813',  # Gerolfingen
    'Q124409312',  # VR SüdBank
    'Q2794077',  # Scheurich
    'Q1879680',  # Löhe-Campus
]
