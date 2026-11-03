"""Durchsicht der Jubiläen 2030 von Hand (November 2026).

Stufen siehe leads.py. Kommentare mit „(aus dem Gedächtnis)“ vor einer Anfrage prüfen.
Zeilen mit Stufe C ohne eigenen Satz: in der Durchsicht gesehen und als Kleinkram verworfen;
der Grund steht kurz dabei.
"""

E = {
    'Q3758': ('S', 'Großstadt, 1200 Jahre (um 830 „Luthra“), Pfalz – Jahr nur „um“'),  # Kaiserslautern
    'Q522165': ('S', 'Stadt, 750 Jahre, 18 000 Einwohner, Pfalz'),  # Wörth am Rhein
    'Q16102': ('S', 'Große Kreisstadt, 1250 Jahre (um 780), 34 000 Einwohner'),  # Kornwestheim
    'Q816360': ('S', '250 Jahre Bembé, älteste Parkettfabrik'),  # Bembé
    'Q566832': ('S', '250 Jahre Mack (Europa-Park-Familie)'),  # Mack Rides GmbH & Co KG
    'Q545371': ('A', '1200 Jahre, 3 400 Einwohner, Landkreis Rastatt – Verlagsgebiet'),  # Au am Rhein
    'Q552671': ('A', '1200 Jahre, 5 100 Einwohner, Pfalz'),  # Waldmohr
    'Q442426': ('A', '1250 Jahre, 1 700 Einwohner, Pfalz'),  # Altleiningen
    'Q662258': ('A', '100 Jahre Gemeinde (Gründung 1930), 11 600 Einwohner, Pfalz'),  # Limburgerhof
    'Q61713': ('A', 'Stadt, 1250 Jahre, 4 900 Einwohner'),  # Niederstetten
    'Q2973': ('A', '700 Jahre Stadtrecht Darmstadt – Nachbargebiet'),  # Darmstadt
    'Q60136': ('A', '900 Jahre, 5 200 Einwohner'),  # Erlenbach
    'Q503260': ('A', '800 Jahre, 5 100 Einwohner'),  # Talheim
    'Q23559805': ('A', '1250 Jahre, Ortsteil von Böhl-Iggelheim – Pfalz'),  # Böhl
    'Q435044': ('A', '1200 Jahre, Stadtteil von Calw – Verlagsgebiet'),  # Altburg (Calw)
    'Q656289': ('A', '1200 Jahre (830–850), 3 900 Einwohner, Rheinhessen – Nachbargebiet'),  # Guntersblum
    'Q1933065': ('A', '1300 Jahre (730/31), Stadtteil von Rastatt – Verlagsgebiet'),  # Plittersdorf
    'Q10764833': ('A', '1250 Jahre (um 780), Weinort, Stadtteil von Neustadt – Pfalz'),  # Mußbach an der Weinstraße
    'Q1550425': ('A', '500 Jahre Schule – außergewöhnliche Zahl'),  # Donauschule Nendingen
    'Q637707': ('A', '100 Jahre Lamy, Heidelberg'),  # Lamy
    'Q19921577': ('A', '250 Jahre Gymnasium, Frankenthal – Pfalz'),  # Karolinen-Gymnasium
    'Q13487057': ('A', '1200 Jahre (830–850), Stadtteil von Mörfelden-Walldorf – Nachbargebiet'),  # Mörfelden
    'Q15785111': ('A', '175 Jahre, Landkreis Karlsruhe – direkt vor der Haustür'),  # Badische Maschinenfabrik Durlach
    'Q284672': ('A', '150 Jahre Bilfinger, Mannheim'),  # Bilfinger SE
    'Q1959244': ('A', '200 Jahre, Familienunternehmen (Chemie), Abstatt'),  # Münzing Chemie
    'Q712034': ('A', '100 Jahre Schwarz-Gruppe (Lidl, Kaufland)'),  # Schwarz-Gruppe
    'Q1441291': ('A', '100 Jahre Schwarzwaldmilch, Genossenschaft'),  # Schwarzwaldmilch
    'Q37068918': ('A', '175 Jahre Kreissparkasse – Festschrift typisch'),  # Kreissparkasse Heidenheim
    'Q151137': ('A', '150 Jahre Steiff'),  # Steiff
    'Q1929771': ('A', '125 Jahre, Holzbearbeitungsmaschinen, Tauberbischofsheim'),  # Michael Weinig AG
    'Q2110996': ('A', '250 Jahre Familienbrauerei, Odenwald – Nachbargebiet'),  # Privat-Brauerei Schmucker
    'Q107368376': ('A', '175 Jahre A. Racke, Bingen – Nachbargebiet'),  # A. Racke
    'Q105254': ('A', '125 Jahre Mainz 05 – Nachbargebiet'),  # 1. FSV Mainz 05
    'Q566942': ('B', '800 Jahre, 3 300 Einwohner, Pfalz'),  # Harthausen
    'Q2865': ('B', '700 Jahre Kassel – weit weg'),  # Kassel
    'Q653270': ('B', '800 Jahre, 710 Einwohner, Pfalz'),  # Mittelbrunn
    'Q567227': ('B', '700 Jahre, 735 Einwohner, Pfalz'),  # Erlenbach bei Kandel
    'Q544758': ('B', '800 Jahre, 2 500 Einwohner'),  # Dörzbach
    'Q232718': ('B', '800 Jahre, 1 800 Einwohner'),  # Zweiflingen
    'Q1623840': ('B', '900 Jahre, Stadtteil von Pforzheim – Verlagsgebiet'),  # Hohenwart
    'Q1559963': ('B', '750 Jahre, Stadtteil von Buchen – Verlagsgebiet'),  # Götzingen
    'Q1366967': ('B', '1250 Jahre, Ortsteil von Obersulm'),  # Eschenau
    'Q64839429': ('B', '1200 Jahre, Stadtteil von Bad Wildbad – Jahr unsicher'),  # Calmbach
    'Q13142778': ('B', '1200 Jahre, kleiner Ortsteil von Efringen-Kirchen'),  # Maugenhard
    'Q1729708': ('B', '1200 Jahre (um 830), Stadtteil von Calw'),  # Stammheim (Calw)
    'Q1894293': ('B', '1100 Jahre, Ortsteil von Meßkirch'),  # Rohrdorf
    'Q654131': ('B', '1250 Jahre, aber 330 Einwohner – Rheinhessen'),  # Bermersheim
    'Q112161327': ('B', '900 Jahre, Ortsteil von Bad Bellingen'),  # Bamlach
    'Q67182807': ('B', '900 Jahre, Stadtteil von Korntal-Münchingen'),  # Münchingen
    'Q1589756': ('B', '900 Jahre, Stadtteil von Bad Krozingen'),  # Schlatt
    'Q1391491': ('B', '800 Jahre, Stadtteil von Aalen'),  # Fachsenfeld
    'Q70223970': ('B', '800 Jahre, Ortsteil von Wallhausen'),  # Hengstfeld
    'Q50445466': ('B', '800 Jahre, Stadtteil von Lauda-Königshofen'),  # Sachsenflur
    'Q105847639': ('B', '750 Jahre, Stadtteil von Weil am Rhein'),  # Friedlingen
    'Q798693': ('B', '700 Jahre, Kurort, Ortsteil von Bad Peterstal-Griesbach'),  # Bad Griesbach im Schwarzwald
    'Q1694501': ('B', '100 Jahre Schule, Heidelberg'),  # St. Raphael-Schulen Heidelberg
    'Q19980223': ('B', '1250 Jahre (um 780), Stadtteil von Tauberbischofsheim'),  # Hochhausen
    'Q1262922': ('B', '100 Jahre, Versandhandel, Pforzheim'),  # Mona Versand
    'Q1743604': ('B', '900 Jahre, Stadtteil von Heppenheim'),  # Kirschhausen
    'Q2008838': ('B', '900 Jahre, Ortsteil von Mörlenbach'),  # Ober-Mumbach
    'Q26794510': ('B', '150 Jahre VR-Bank – Fusion prüfen'),  # VR-Bank in Mittelbaden
    'Q16857824': ('B', '150 Jahre Volksbank – Fusion prüfen'),  # Volksbank Kirnau
    'Q1569797': ('B', '150 Jahre Gymnasium, Neustadt – Pfalz'),  # Kurfürst-Ruprecht-Gymnasium Neustadt
    'Q1324579': ('B', '125 Jahre Goldschmiedeschule, Pforzheim'),  # Goldschmiedeschule mit Uhrmacherschule Pforzheim
    'Q682489': ('B', '125 Jahre Gymnasium, Pforzheim'),  # Reuchlin-Gymnasium Pforzheim
    'Q106560674': ('B', '100 Jahre, Mudau'),  # AURORA Konrad G. Schulz GmbH & Co. KG
    'Q107102723': ('B', '100 Jahre, Gernsbach'),  # Katz & Klumpp
    'Q2641625': ('B', '200 Jahre Gymnasium, Konstanz'),  # Alexander-von-Humboldt-Gymnasium Konstanz
    'Q106905625': ('B', '100 Jahre Gymnasium'),  # Karl-Maybach-Gymnasium
    'Q1645492': ('B', '100 Jahre Internatsschule'),  # Urspringschule
    'Q18677538': ('B', '125 Jahre Hotel, Buchen'),  # Bahnhof-Hotel Buchen
    'Q116779466': ('B', '125 Jahre, Enzkreis'),  # Felss Group
    'Q2307319': ('B', '200 Jahre Sparkasse Bensheim'),  # Sparkasse Bensheim
    'Q2307691': ('B', '200 Jahre Sparkasse Starkenburg'),  # Sparkasse Starkenburg
    'Q107139026': ('B', '125 Jahre Rheinschifffahrt, Karlsruhe'),  # Karlsruher Schiffahrts-Aktiengesellschaft
    'Q21039726': ('B', '100 Jahre'),  # Bott
    'Q18287000': ('B', '100 Jahre Galerie'),  # Galerie Valentien
    'Q74821641': ('B', '100 Jahre'),  # Gottlob Brodbeck GmbH & Co. KG
    'Q473339': ('B', '100 Jahre Modellbau'),  # Graupner Modellbau
    'Q1343869': ('B', '100 Jahre Stadtwerke'),  # Stadtwerke Tübingen
    'Q2584554': ('B', '100 Jahre'),  # Winkels
    'Q101131482': ('B', '150 Jahre Volksbank – Fusion prüfen'),  # Volksbank Sulmtal
    'Q111731279': ('B', '250 Jahre Gartenbauschule'),  # Staatsschule für Gartenbau
    'Q135653480': ('B', '175 Jahre Gymnasium'),  # Erasmus-Widmann-Gymnasium
    'Q317125': ('B', '175 Jahre Hochschule'),  # Hochschule Reutlingen
    'Q121337900': ('B', '150 Jahre Gymnasium'),  # Eugen-Bolz-Gymnasium Rottenburg
    'Q1558156': ('B', '150 Jahre Gymnasium'),  # Gymnasium Korntal-Münchingen
    'Q1515507': ('B', '125 Jahre Berufsschule'),  # Gertrud-Luckner-Gewerbeschule
    'Q65239480': ('B', '125 Jahre Berufsschule'),  # Kaufmännische Schule 1
    'Q130973425': ('B', '100 Jahre'),  # RAPID Maschinenbau
    'Q106676426': ('B', '100 Jahre'),  # Rüdinger Spedition
    'Q138000577': ('B', '100 Jahre'),  # W. Gessmann GmbH
    'Q29051466': ('B', '100 Jahre Thomas-Gruppe – Hunsrück'),  # Thomas-Gruppe
    'Q105728710': ('B', '150 Jahre'),  # Sülzle Gruppe
    'Q1782066': ('B', '125 Jahre'),  # Konrad Hornschuch AG
    'Q1914433': ('B', '125 Jahre'),  # Mayer & Cie.
    'Q850250': ('B', '75 Jahre Betty Barclay'),  # Betty Barclay
    'Q1739924': ('B', '75 Jahre Nolte-Möbel – Pfalz'),  # Nolte-Möbel
    'Q331492': ('B', '75 Jahre'),  # Weidenhammer Packaging Group
    'Q21559758': ('B', '75 Jahre'),  # Weisenburger Bau
    'Q850105': ('B', '175 Jahre Schule – Frankfurt'),  # Bettinaschule
    'Q1301613': ('B', '175 Jahre Gymnasium – Koblenz'),  # Eichendorff-Gymnasium
    'Q2073433': ('B', '175 Jahre Gymnasium – Ahrtal'),  # Peter-Joerres-Gymnasium
    'Q24288521': ('B', '150 Jahre Gymnasium, Alzey'),  # Staatliches Aufbaugymnasium Alzey
    'Q131819633': ('B', '125 Jahre Kunsthaus'),  # Kunsthaus Bühler
    'Q140285668': ('B', '125 Jahre'),  # REVEN
    'Q702674': ('B', '150 Jahre Wella – Konzern'),  # Wella
    'Q1181452': ('B', '125 Jahre Debeka – weit weg'),  # Debeka
    'Q1742261': ('B', '125 Jahre Kinopolis'),  # Kinopolis
    'Q2384949': ('B', '150 Jahre TSG Haßloch – Pfalz'),  # TSG Haßloch
    'Q120793245': ('B', '150 Jahre Turnverein, Mannheim'),  # TV 1880 Käfertal
    'Q2640881': ('B', '125 Jahre, Verlagsgebiet'),  # Alemannia Eggenstein
    'Q442055': ('B', '125 Jahre, Pfalz'),  # SG 05 Pirmasens
    'Q690818': ('B', '125 Jahre, Verlagsgebiet'),  # TV Busenbach
    'Q2521431': ('B', '125 Jahre, Karlsruhe'),  # VfB 05 Knielingen
    'Q71147104': ('B', '125 Jahre, Pfalz'),  # VfR Friesenheim
    'Q799100': ('B', '75 Jahre'),  # Badische Stahlwerke
    'Q1205813': ('B', '75 Jahre Literaturarchiv'),  # Deutsches Literaturarchiv Marbach
    'Q19968199': ('B', '75 Jahre Bibliothek'),  # Hegau Bibliothek
    'Q833959': ('B', '75 Jahre'),  # Sto
    'Q50194071': ('B', '125 Jahre Alpenvereinssektion – Festschrift üblich'),  # Sektion Reutlingen des Deutschen Alpenvereins
    'Q114045662': ('B', '125 Jahre Geschichtsverein – Partner'),  # Historischer Verein Ingelheim
    'Q370194': ('C', '900 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Breunigweiler
    'Q122850': ('C', '600 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Herchweiler
    'Q132024391': ('C', 'Bauwerk/Wohnplatz, keine Ersterwähnung'),  # Odinstal
    'Q1365854': ('C', '1000 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Silwingen
    'Q19963091': ('C', '700 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Kaltenbrunn
    'Q1644765': ('C', '800 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Igelsberg
    'Q788829': ('C', '800 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Kottspiel
    'Q15846673': ('C', '750 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Schollach
    'Q130729959': ('C', '700 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Berwinkel
    'Q15110436': ('C', '700 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Eschenau (Sankt Julian)
    'Q1262242': ('C', '700 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Happenbach
    'Q2009198': ('C', '700 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Oberbiegelhof
    'Q105700003': ('C', '700 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Reichenbach
    'Q2597093': ('C', '700 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Wüstenhausen
    'Q2346273': ('C', '900 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Stetten (Hechingen)
    'Q1230400': ('C', '900 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Streichen
    'Q65183907': ('C', '900 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Vogelbach
    'Q28468877': ('C', '750 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Gremmelsbach
    'Q2569725': ('C', '750 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Wiflingshausen
    'Q116436197': ('C', '700 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Blankenhöfe
    'Q28913718': ('C', '700 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Blindheim
    'Q561172': ('C', '800 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Rehbach
    'Q862882': ('C', '600 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Billings
    'Q1457727': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Friedrich-Ebert-Schule (Frankfurt am Main)
    'Q47507450': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Geschwister-Scholl-Schule
    'Q829459': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # H World International
    'Q21039796': ('C', '200 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Fürstenberg
    'Q140753461': ('C', '200 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Backhaus
    'Q29065474': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Hermann-Kiderlen-Hütte
    'Q56230610': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Freie Waldorfschule Kassel Bildung.Erziehung.Betreuung
    'Q116679995': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Auto Bach Gruppe
    'Q1276039': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # ESWE Versorgung
    'Q123367912': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Glinicke Automobil Holding
    'Q1262616': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Duale Hochschule Baden-Württemberg Mosbach
    'Q117784723': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Integrierte Gesamtschule Ernst Bloch
    'Q116621229': ('C', '150 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Volksbank Grebenhain
    'Q828744': ('C', '175 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Berufliche Schulen Korbach und Bad Arolsen
    'Q1014983': ('C', '125 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Burggymnasium Friedberg
    'Q1357123': ('C', '125 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Ernst-Ludwig-Schule Bad Nauheim
    'Q1813815': ('C', '125 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Leibnizschule Wiesbaden
    'Q83781359': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Glinicke Automobilguppe
    'Q134452042': ('C', '100 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Rudolf-Steiner-Institut
    'Q1667438': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Internationale Schule Neustadt
    'Q1140927': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Cronimet
    'Q2266946': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Seitenbacher
    'Q112601114': ('C', '100 Jahre, kleiner Verein'),  # K.K.Schützenverein Allfeld 1930 e.V.
    'Q112685626': ('C', '100 Jahre, kleiner Verein'),  # KK-Schützenverein Neckarelz-Diedesheim 1930 e.V.
    'Q112282089': ('C', '100 Jahre, kleiner Verein'),  # Schützenverein Hainstadt 1930 e.V.
    'Q15781482': ('C', '150 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Adler Real Estate
    'Q2506405': ('C', '150 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # VR-Bankverein Bad Hersfeld-Rotenburg
    'Q3600443': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Abas Software AG
    'Q55739873': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Goethe-Institut Mannheim
    'Q124627714': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Goldschmitt techmobil
    'Q140644946': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # TUP GmbH & Co. KG
    'Q98382768': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Alfatraining Bildungszentrum
    'Q716947': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Arvato Infoscore
    'Q140290468': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Exxeta
    'Q27938189': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Gigamot
    'Q835897': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Mannheim Business School
    'Q1502951': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # OSADL
    'Q107359734': ('C', '200 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Alte Leipziger Lebensversicherungsgesellschaft aG
    'Q97065048': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Albert-Schweitzer-Gymnasium Neckarsulm
    'Q60724761': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Justinus-Kerner-Gymnasium Heilbronn
    'Q107177296': ('C', '150 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Actien-Zuckerfabrik Wabern
    'Q107378852': ('C', '150 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Hako-Schuh
    'Q138605997': ('C', '125 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Masa GmbH
    'Q39214298': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Christy-Brown-Schule
    'Q1380177': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Kurt-von-Marval-Schule Gemeinschaftsschule Nordheim
    'Q2258863': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Schönbuch-Gymnasium Holzgerlingen
    'Q30268730': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Apogenix (Germany)
    'Q64143675': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # CIP Chemisches Institut Pforzheim GmbH
    'Q112709370': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Galerie Renaissance
    'Q136714846': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Heid Energieberatung
    'Q136706792': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Heid Immobilienbewertung
    'Q30261294': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Technologie-Initiative SmartFactory KL
    'Q46552026': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # medpex
    'Q138094791': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Realschule Bernhausen
    'Q798891': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Baden-Württembergische Bank
    'Q2506422': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # VR-Bank Ludwigsburg
    'Q1525588': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Michael Becker Orgelbau
    'Q2420071': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Pfleiderer Leutkirch
    'Q117455171': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Waiko Möbelwerke
    'Q134593928': ('C', '150 Jahre, kleiner Verein'),  # Freiwillige Feuerwehr Bad Wildbad Abteilung Calmbach
    'Q24562921': ('C', '100 Jahre, kleiner Verein'),  # SV Alberweiler
    'Q1609681': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Herderschule Kassel
    'Q1755658': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Max-von-Laue-Gymnasium
    'Q129261824': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # vhs Volkshochschule Metzingen-Ermstal
    'Q191217': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Kampfmittelabwehrschule
    'Q48922666': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenzunft Bettschoner 1930 Weitingen
    'Q218580': ('C', '75 Jahre, Weiler/Kleinstort oder Jahr unsicher'),  # Zollberg
    'Q759176': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Auer Weber
    'Q907089': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Brekina
    'Q1688522': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Bucher Automation
    'Q1547348': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Groninger
    'Q2071253': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Perkins Park
    'Q56192315': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # RAMPF-Gruppe
    'Q14395717': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # SL Rasch
    'Q325302': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Schlaich Bergermann Partner
    'Q21036522': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Zimmer Group
    'Q1024117': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # msg life ag
    'Q112841662': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Hauff-Technik
    'Q760638': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # August-Hermann-Francke-Schule Gießen
    'Q1391274': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Fachhochschule für Verwaltung des Saarlandes
    'Q1615884': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Hessische Hochschule für Finanzen und Rechtspflege
    'Q21224996': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Taiji Akademie
    'Q55669972': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Volkshochschule Horn
    'Q15107731': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Flatex Bank
    'Q56476071': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Freie Waldorfschule Biberach
    'Q857373': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Bieling Architekten
    'Q98533651': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Busch
    'Q312517': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Condor Flugdienst
    'Q21036498': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Esterer
    'Q1377491': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Europäische Rundfunk- und Fernseh-GmbH
    'Q99355544': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Filmkreis
    'Q440408': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Hager SE
    'Q29019803': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Haus am Maiberg
    'Q896102': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Hessische Landesbahn
    'Q100956179': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # IEF-Werner GmbH
    'Q738437': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Assa Abloy Sicherheitstechnik
    'Q15789039': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Blue Ocean Entertainment
    'Q2939019': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Carl Zeiss Vision
    'Q1121314': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Compassio
    'Q25389001': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # DWK Life Sciences
    'Q1513916': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # German Graduate School of Management and Law
    'Q1724726': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # MAG IAS
    'Q1254047': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # SBB GmbH
    'Q19269854': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # SchwabenQuellen
    'Q2326971': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Stadtbus Crailsheim
    'Q61740994': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # TeamViewer SE
    'Q1202997': ('C', '100 Jahre, kleiner Verein'),  # Deutsche Gesellschaft für Sprechwissenschaft und Sprecherziehung
    'Q56249881': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Freie Waldorfschule Trier
    'Q2496396': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Universitätsklinikum Gießen und Marburg
    'Q471190': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Amigo
    'Q30120813': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Controlware
    'Q1341471': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Energieversorgung Offenbach
    'Q1532505': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Glyn
    'Q1676681': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # JSK Architekten
    'Q30499582': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # MHK Group
    'Q1387867': ('C', '125 Jahre, kleiner Verein'),  # FC Wangen 05
    'Q1691626': ('C', '125 Jahre, kleiner Verein'),  # NLV Stuttgart-Vaihingen
    'Q819367': ('C', '125 Jahre, kleiner Verein'),  # SSV Reutlingen 05
    'Q2521570': ('C', '125 Jahre, kleiner Verein'),  # VfL Ulm/Neu-Ulm
    'Q135999436': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Hager
    'Q16984549': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Hoffmann Propeller
    'Q130342816': ('C', '75 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Unilux
    'Q30289732': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # BioFluidix (Germany)
    'Q137291545': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # SECANDA AG
    'Q12071344': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Q12071344
    'Q62479109': ('C', '50 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # FlugBlatt-Presse
    'Q125778105': ('C', '50 Jahre, kleiner Verein'),  # Kunstverein Villa Streccius
    'Q62577307': ('C', '50 Jahre, kleiner Verein'),  # SV Karlsruhe-Beiertheim
    'Q1152124': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # DB Vertrieb
    'Q123244476': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Data Center Group
    'Q1347673': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Eprimo
    'Q1654166': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Ineos Melamines
    'Q669722': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Keen Games
    'Q21036873': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Kuenstlerhaus43
    'Q1921489': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Mercurius (Unternehmen)
    'Q46998964': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Operational services
    'Q1290502': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # SEB Asset Management
    'Q97102037': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Trendtours Touristik
    'Q2516293': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Verkehrsgesellschaft Oberhessen
    'Q2521737': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Vias (Unternehmen)
    'Q1033911': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # cantus Verkehrsgesellschaft
    'Q111020377': ('C', '75 Jahre, kleiner Verein'),  # Freiwillige Feuerwehr Buchen Abteilung Hettigenbeuern
    'Q111027607': ('C', '75 Jahre, kleiner Verein'),  # Freiwillige Feuerwehr Elztal Abteilung Neckarburken
    'Q111028122': ('C', '75 Jahre, kleiner Verein'),  # Schützenverein 1955 Altheim e.V.
    'Q115498222': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # IB Berufliche Schulen Reutlingen
    'Q47525594': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Herdermer Lalli-Zunft 1930
    'Q55788424': ('C', '150 Jahre, kleiner Verein'),  # Freiwillige Feuerwehr Beselich-Obertiefenbach
    'Q1256570': ('C', '150 Jahre, kleiner Verein'),  # Kasteler Ruder- und Kanu-Gesellschaft
    'Q479194': ('C', '150 Jahre, kleiner Verein'),  # SC Frankfurt 1880
    'Q819382': ('C', '125 Jahre, kleiner Verein'),  # Borussia Neunkirchen
    'Q647893': ('C', '125 Jahre, kleiner Verein'),  # Eintracht Trier
    'Q639569': ('C', '125 Jahre, kleiner Verein'),  # Eintracht Wetzlar
    'Q549565': ('C', '125 Jahre, kleiner Verein'),  # SV Saar 05 Saarbrücken
    'Q7572132': ('C', '125 Jahre, kleiner Verein'),  # Spvgg 05 Frankfurt-Oberrad
    'Q2386051': ('C', '125 Jahre, kleiner Verein'),  # TV 05 Waldgirmes
    'Q639454': ('C', '125 Jahre, kleiner Verein'),  # VfB Marburg
    'Q187653': ('C', '125 Jahre, kleiner Verein'),  # Zellcheming
    'Q113279345': ('C', '50 Jahre, kleiner Verein'),  # Delta Club Mosbach e.V.
    'Q112918999': ('C', '50 Jahre, kleiner Verein'),  # TC Grün-Weiß Hüffenhardt e.V. 80
    'Q83793628': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Audi Service Baunatal
    'Q100105751': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Autismus-Stiftung
    'Q23893325': ('C', '25 Jahre, kleiner Verein'),  # FC Portus Pforzheim
    'Q30254722': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # GenXPro (Germany)
    'Q140385918': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Grapevine
    'Q122750594': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Great Movies
    'Q1654576': ('C', '25 Jahre, kleiner Verein'),  # ISC Mannheim
    'Q140470828': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Ihre Telefonzentrale
    'Q46526521': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # United Internet Investments Holding
    'Q83793582': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Volkswagen Nutzfahrzeuge Service Baunatal
    'Q83793565': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Volkswagen Service Baunatal
    'Q141162561': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # müllerritzrow GmbH
    'Q83794164': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # Škoda Service Baunatal
    'Q7850490': ('C', '125 Jahre, kleiner Verein'),  # TuS Dehrn
    'Q116499317': ('C', '25 Jahre, kleiner Verein'),  # SCL Baden-Baden
    'Q129330409': ('C', '75 Jahre, kleiner Verein'),  # Architektenkammer Baden-Württemberg
    'Q1322816': ('C', '75 Jahre, kleiner Verein'),  # Gospel Forum
    'Q117212397': ('C', '75 Jahre, kleiner Verein'),  # Sektion Ludwigsburg des Deutschen Alpenvereins
    'Q111029128': ('C', '50 Jahre, kleiner Verein'),  # Fischerei- und Sportanglerverein Schefflenz 1980 e.V.
    'Q49161829': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Hexenzunft Bühler Hexen 1955
    'Q15428295': ('C', '50 Jahre, kleiner Verein'),  # Hilfe für Brüder
    'Q122231851': ('C', '50 Jahre, kleiner Verein'),  # Kunstverein Brackenheim
    'Q2505122': ('C', '25 Jahre – zu klein oder zu weit weg für ein Jubiläumsbuch'),  # V-Bus
    'Q47525619': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Schwarzwälder Tannenzapfen
    'Q47498672': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Sundheimer Hexen 1955
    'Q47498631': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenzunft Bühler Hemglunkerle
    'Q116765729': ('C', '50 Jahre, kleiner Verein'),  # Schachclub BG Buchen 1980 e. V.
    'Q19968139': ('C', '75 Jahre, kleiner Verein'),  # Deutscher Verband für Podologie
    'Q2206368': ('C', '75 Jahre, kleiner Verein'),  # SKG Frankfurt
    'Q47538735': ('C', '50 Jahre, kleiner Verein'),  # Baronengilde Lackendorf
    'Q50319100': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Betzinger Krautskräga 1980
    'Q48610081': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Eschbacher Hexen-Kessel
    'Q15804888': ('C', '25 Jahre, kleiner Verein'),  # Debattierclub Freiburg
    'Q1566147': ('C', '25 Jahre, kleiner Verein'),  # HV Stuttgarter Kickers
    'Q2598026': ('C', '25 Jahre, kleiner Verein'),  # Xertifix
    'Q627483': ('C', '50 Jahre, kleiner Verein'),  # Arbeitsgemeinschaft Dokumentarfilm
    'Q1481465': ('C', '50 Jahre, kleiner Verein'),  # Hessische Hochschule für öffentliches Management und Sicherheit
    'Q28668449': ('C', '25 Jahre, kleiner Verein'),  # Crailsheim Titans
    'Q48610021': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Dennenberghexen Neustadt
    'Q50319136': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Hägeles-Hexen Rottenburg
    'Q162437': ('C', '25 Jahre, kleiner Verein'),  # 1. FFC Montabaur
    'Q15851420': ('C', '25 Jahre, kleiner Verein'),  # Travel Industry Club
    'Q47525565': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenzunft Zähringer Burgnarren
    'Q47784716': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Bocknarrenzunft Boxberg 1957
    'Q48975809': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenverein Waldlütle Eberfingen
    'Q48551668': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenzunft Mühlenbach
    'Q48705796': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenzunft Narrhalla Ringingen
    'Q48705899': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenzunft Ringingen
    'Q98280617': ('C', '25 Jahre, kleiner Verein'),  # Förderverein Schwimmbäder Reinhardshagen e.V.
    'Q108345257': ('C', '25 Jahre, kleiner Verein'),  # Lehrstellwerk Kornwestheim
    'Q47784823': ('C', 'Narrenzunft – Festheft, kein Buch'),  # 1. Narrenzunft Großbettlingen Stoiadler 2006
    'Q50319934': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenfreunde Frech-Dachs Rottenburg 2005
    'Q32188807': ('X', 'Doppelter Eintrag, gleicher Ort wie Stadt Niederstetten'),  # Niederstetten
    'Q1445741': ('X', 'Burg (Jugendherberge) – kein Käufer'),  # Veitsburg
    'Q141447360': ('X', 'Gebäude'),  # Wiesenweg 145
    'Q327188': ('X', 'Konzernteil (Mahle)'),  # Mahle Behr
    'Q135683971': ('X', 'Gebäude'),  # Gießhalle
    'Q47497580': ('X', 'Gebäude'),  # Dreiseithof mit Wirtschaftsgebäude
    'Q28979032': ('X', 'Konzerntochter'),  # BorgWarner (Germany)
    'Q47497358': ('X', 'Gebäude'),  # Vierseithof
    'Q2496475': ('X', 'Verlag – macht es selbst'),  # Universum Verlag
    'Q2516768': ('X', 'Verlag – macht es selbst'),  # Verlag Vittorio Klostermann
    'Q1633020': ('X', 'Verlag – macht es selbst'),  # Huber Verlag
    'Q1298298': ('X', 'Verlag – macht es selbst'),  # Kaden Verlag
    'Q1916525': ('X', 'Verlag – macht es selbst'),  # Medizificon Verlag
    'Q103046411': ('X', 'Verlag – macht es selbst'),  # frechverlag
    'Q1301287': ('X', 'Verlag – macht es selbst'),  # Eichborn Verlag
    'Q15824744': ('X', 'Verlag – macht es selbst'),  # Kynos Verlag
    'Q18286509': ('X', 'Verlag – macht es selbst'),  # Brighton Verlag
    'Q1741780': ('X', 'Verlag – macht es selbst'),  # Kinderbuchverlag Wolff
    'Q107176233': ('X', 'Verlag – macht es selbst'),  # Nomen Verlag
    'Q1606795': ('X', 'Verlag – macht es selbst'),  # Seeling Verlag
    # Bayerische Grenzkreise (Nachtrag November 2026)
    'Q502441': ('A', 'Stadt, 900 Jahre (1130), 6 100 Einwohner, Knabenchor'),  # Windsbach
    'Q21084038': ('A', '1000 Jahre, Hauptort des Marktes Jettingen-Scheppach'),  # Jettingen
    'Q1676521': ('A', '75 Jahre, großes Familienunternehmen (Sonnenschutz)'),  # Warema Renkhoff SE
    'Q489685': ('B', '„1600 Jahre“ nur über eine Erwähnung um 430 – Stadt zählt eher ab dem Römerkastell'),  # Günzburg
    'Q505306': ('B', '1200 Jahre, aber nur Anfang des Zeitraums 830–850; 2 500 Einwohner'),  # Obernzenn
    'Q758371': ('B', '750 Jahre, Stadtteil von Illertissen'),  # Au
    'Q22692809': ('B', '150 Jahre Schule – Jahr nur aus Wikidata'),  # Lessing-Gymnasium
    'Q511325': ('C', 'Urkunde 780–802 – kein festes Jahr, 1 000 Einwohner'),  # Ergersheim
    'Q1267012': ('C', '900 Jahre, kleiner Ortsteil'),  # Glöttweng
    'Q1226296': ('C', '800 Jahre, Weiler mit Burg'),  # Unterschönegg
    'Q109375998': ('C', '700 Jahre, Weiler'),  # Bichel
    'Q111079901': ('C', '700 Jahre, Weiler'),  # Riegen
    'Q113675594': ('C', '700 Jahre, Weiler'),  # Unterreute
    'Q114567798': ('C', '500 Jahre, Weiler'),  # Geratsried
    'Q126039825': ('C', '500 Jahre, Weiler'),  # Jauchen
    'Q124334309': ('C', 'Gründung nur „etwa 530“'),  # Kleinwindsheim
    'Q757951': ('C', '900 Jahre, 330 Einwohner'),  # Attenhausen
    'Q123122457': ('C', '750 Jahre, 250 Einwohner'),  # Altmannshausen
    'Q819533': ('C', '750 Jahre, 94 Einwohner'),  # Bergendorf
    'Q19792269': ('C', 'Mühle'),  # Ettringermühle
    'Q1568523': ('C', '750 Jahre, 45 Einwohner'),  # Hagenheim
    'Q130546943': ('C', 'Hof, 43 Einwohner'),  # Klotzenhof
    'Q2542732': ('C', '750 Jahre, 240 Einwohner'),  # Wallerdorf
    'Q2340239': ('C', '700 Jahre, 25 Einwohner'),  # Steinbächlein
    'Q2511921': ('C', 'Beleg nennt Lehensbücher 1330–1470, keine Ersterwähnung'),  # Vehlberg
    'Q33926983': ('C', 'Mühle, 5 Einwohner'),  # Finkenmühle
    'Q120866342': ('C', 'Weiler, 24 Einwohner'),  # Winneberg
    'Q114240113': ('C', '500 Jahre, Weiler'),  # Schwarzen
    'Q18286522': ('C', 'Urkunde 780–802, 195 Einwohner'),  # Buchheim
    'Q105205089': ('C', 'Urbar 830–850, 117 Einwohner'),  # Seenheim
    'Q1622340': ('C', '900 Jahre, 180 Einwohner'),  # Hochstein
    'Q109324040': ('C', 'Weiler, 12 Einwohner'),  # Tödting
    'Q34417898': ('C', '„vor 1430“, 98 Einwohner'),  # Elgersdorf
    'Q34704865': ('C', '„vor 1430“, 19 Einwohner'),  # Grieshof
    'Q34561192': ('C', '„vor 1430“, 31 Einwohner'),  # Kaltenneuses
    'Q131742409': ('X', 'Forsthaus, nur auf einer Karte von 1905 verzeichnet'),  # Forsthaus Zieglerfeld
    'Q2110560': ('C', 'Alpenvereinshütte'),  # Prinz-Luitpold-Haus
    'Q819569': ('C', '75 Jahre, wenig bekannt'),  # Berger Holding
    'Q2208576': ('C', '100 Jahre, Sportverein – Festheft'),  # SV Oberelchingen
    'Q1285023': ('C', '50 Jahre – kleine Zahl'),  # Heunisch-Guss
    'Q139487975': ('C', '50 Jahre – kleine Zahl'),  # GROHA Gropper Hallen GmbH
    'Q1349170': ('C', '25 Jahre Basketballteam – eigentlich Ehingen (Donau), Ort falsch zugeordnet'),  # Team Ehingen Urspring
    'Q50318769': ('C', 'Narrenzunft – Festheft, kein Buch'),  # Narrenzunft D'Illerstoi Senden
}

KOMMENTAR = {
    'Q3758': '1200 Jahre Kaiserslautern – die Ersterwähnung als „Villa Luthra“ ist allerdings nur auf „um 830“ datiert. Großstadt mit 100 000 Einwohnern in der Pfalz, also im Verlagsgebiet. Vorher klären, ob die Stadt 2030 feiert; wenn ja, ist das ein großes Stadtbuch.',  # Kaiserslautern
    'Q522165': '750 Jahre (1280). Stadt mit 18 000 Einwohnern am Rhein, Kreis Germersheim – Pfalz, nah am Verlag. Groß genug für ein Stadtbuch.',  # Wörth am Rhein
    'Q16102': '1250 Jahre (um 780 „Westheim“). Große Kreisstadt mit 34 000 Einwohnern im Kreis Ludwigsburg. Jahr nur „um“ – vorher klären, ob 2030 gefeiert wird.',  # Kornwestheim
    'Q566832': '250 Jahre Mack Rides in Waldkirch – die Familie hinter dem Europa-Park. Wer 250 Jahre feiert und einen Freizeitpark betreibt, macht ein Buch. Mein Favorit unter den Firmen.',  # Mack Rides GmbH & Co KG
    'Q816360': '250 Jahre Bembé in Bad Mergentheim, gilt als älteste Parkettfabrik (560 Mitarbeiter). Familiengeschichte über 250 Jahre – passt gut zu einem Regionalverlag.',  # Bembé
    'Q545371': '1200 Jahre (830 „Augia“). Gemeinde mit 3 400 Einwohnern im Landkreis Rastatt – Verlagsgebiet.',  # Au am Rhein
    'Q552671': '1200 Jahre (Lorscher Codex 830 „villa Moraha“). Gemeinde mit 5 100 Einwohnern im Kreis Kusel – Pfalz.',  # Waldmohr
    'Q442426': '1250 Jahre (Lorscher Codex 780, damals „Leiningen“). Gemeinde mit 1 700 Einwohnern im Kreis Bad Dürkheim – namensgebend für die Grafen von Leiningen, also viel Geschichte für einen kleinen Ort.',  # Altleiningen
    'Q662258': '100 Jahre Gemeinde Limburgerhof (zum 1. Januar 1930 aus Teilen von Mutterstadt und Neuhofen gebildet). 11 600 Einwohner im Rhein-Pfalz-Kreis. Eine junge Gemeinde mit erstem großen Jubiläum – gute Chance auf ein Festbuch.',  # Limburgerhof
    'Q61713': '1250 Jahre (780, Traditiones Fuldenses). Stadt mit 4 900 Einwohnern im Main-Tauber-Kreis.',  # Niederstetten
    'Q60136': '900 Jahre (Hirsauer Codex 1130). Gemeinde mit 5 200 Einwohnern im Kreis Heilbronn, Weinort.',  # Erlenbach
    'Q503260': '800 Jahre (1230). Gemeinde mit 5 100 Einwohnern im Kreis Heilbronn.',  # Talheim
    'Q23559805': '1250 Jahre (Lorscher Codex 780). Böhl ist ein Ortsteil von Böhl-Iggelheim (Rhein-Pfalz-Kreis) – die Gemeinde dürfte mitfeiern.',  # Böhl
    'Q435044': '1200 Jahre (830, Kloster Hirsau). Stadtteil von Calw – Verlagsgebiet.',  # Altburg (Calw)
    'Q1933065': '1300 Jahre (730/31, Kloster Weißenburg) – eine der höchsten Zahlen der Liste. Stadtteil von Rastatt, Verlagsgebiet. Datierung 730/31, das frühere Jahr genommen.',  # Plittersdorf
    'Q10764833': '1250 Jahre (um 780, Kloster Fulda). Weinort, Stadtteil von Neustadt an der Weinstraße – Pfalz.',  # Mußbach an der Weinstraße
    'Q1550425': '500 Jahre Donauschule Nendingen (Tuttlingen) – für eine Grundschule eine außergewöhnliche Zahl.',  # Donauschule Nendingen
    'Q637707': '100 Jahre Lamy in Heidelberg. Seit 2024 japanischer Eigentümer (aus dem Gedächtnis), aber Firmengeschichte und Design sind der Markenkern.',  # Lamy
    'Q19921577': '250 Jahre Karolinen-Gymnasium Frankenthal (Pfalz). Eine Festschrift ist so gut wie gesetzt.',  # Karolinen-Gymnasium
    'Q15785111': '175 Jahre Badische Maschinenfabrik Durlach, Landkreis Karlsruhe – direkt vor der Haustür. Vorher prüfen, ob es die Firma in dieser Form noch gibt.',  # Badische Maschinenfabrik Durlach
    'Q284672': '150 Jahre Bilfinger in Mannheim, börsennotiert. Geld ist da, vermutlich auch eine eigene Kommunikation.',  # Bilfinger SE
    'Q1959244': '200 Jahre Münzing Chemie in Abstatt (Kreis Heilbronn), Familienunternehmen.',  # Münzing Chemie
    'Q712034': '100 Jahre Schwarz-Gruppe (Lidl, Kaufland) in Neckarsulm. Riesig und mit viel Geld – ob ein Regionalverlag zum Zug kommt, ist offen, aber die Region Heilbronn ist ihr Ursprung.',  # Schwarz-Gruppe
    'Q1441291': '100 Jahre Schwarzwaldmilch in Freiburg, Genossenschaft mit regionaler Verwurzelung.',  # Schwarzwaldmilch
    'Q37068918': '175 Jahre Kreissparkasse Heidenheim. Sparkassen machen gern eine Festschrift.',  # Kreissparkasse Heidenheim
    'Q151137': '150 Jahre Steiff in Giengen. Weltbekannte Marke mit großer Geschichte.',  # Steiff
    'Q1929771': '125 Jahre Michael Weinig AG in Tauberbischofsheim.',  # Michael Weinig AG
    'Q656289': '1200 Jahre (Lorscher Codex, 830–850). Weinort mit 3 900 Einwohnern in Rheinhessen. Datierung ist eine Spanne – vorher fragen, wann gefeiert wird. Nachbargebiet.',  # Guntersblum
    'Q13487057': '1200 Jahre (Lorscher Codex, 830–850). Mörfelden ist Teil der Stadt Mörfelden-Walldorf (35 000 Einwohner). Datierung ist eine Spanne. Nachbargebiet.',  # Mörfelden
    'Q2973': '700 Jahre Darmstadt – laut Wikidata 1330, das ist das Jahr der Stadtrechte, nicht die Ersterwähnung. Großstadt mit 168 000 Einwohnern, Nachbargebiet. Vorher klären, ob die Stadt 2030 feiert.',  # Darmstadt
    'Q2110996': '250 Jahre Privat-Brauerei Schmucker in Mossautal (Odenwald). Familienbrauerei. Nachbargebiet.',  # Privat-Brauerei Schmucker
    'Q107368376': '175 Jahre A. Racke in Bingen, Familienunternehmen (Wein und Spirituosen). Nachbargebiet.',  # A. Racke
    'Q105254': '125 Jahre 1. FSV Mainz 05, Bundesliga. Eine Chronik kommt sicher – ob mit einem Regionalverlag, ist offen. Nachbargebiet.',  # 1. FSV Mainz 05
    'Q566942': '800 Jahre (1230), Gemeinde mit 3 300 Einwohnern im Rhein-Pfalz-Kreis.',  # Harthausen
    'Q653270': '800 Jahre (1230), nur 710 Einwohner, Pfalz.',  # Mittelbrunn
    'Q567227': '700 Jahre (1330), 735 Einwohner, Kreis Germersheim.',  # Erlenbach bei Kandel
    'Q544758': '800 Jahre (1230), 2 500 Einwohner, Hohenlohekreis.',  # Dörzbach
    'Q232718': '800 Jahre (1230), 1 800 Einwohner, Hohenlohekreis.',  # Zweiflingen
    'Q1623840': '900 Jahre, Stadtteil von Pforzheim – Verlagsgebiet.',  # Hohenwart
    'Q1559963': '750 Jahre, Stadtteil von Buchen (Neckar-Odenwald-Kreis) – Verlagsgebiet.',  # Götzingen
    'Q1366967': '1250 Jahre (780 „Villa Esginaha“), Ortsteil von Obersulm.',  # Eschenau
    'Q64839429': '1200 Jahre (830 „Calenbach“), Stadtteil von Bad Wildbad – laut Artikel unsicher.',  # Calmbach
    'Q13142778': '1200 Jahre (830), kleiner Ortsteil von Efringen-Kirchen.',  # Maugenhard
    'Q1729708': '1200 Jahre (um 830), Stadtteil von Calw – Verlagsgebiet.',  # Stammheim (Calw)
    'Q1894293': '1100 Jahre (930, St. Gallen), Ortsteil von Meßkirch.',  # Rohrdorf
    'Q112161327': '900 Jahre (1130), Ortsteil von Bad Bellingen.',  # Bamlach
    'Q67182807': '900 Jahre (1130), Münchingen ist Teil von Korntal-Münchingen.',  # Münchingen
    'Q1589756': '900 Jahre (1130), Ortsteil von Bad Krozingen.',  # Schlatt
    'Q1391491': '800 Jahre (1230), Stadtteil von Aalen.',  # Fachsenfeld
    'Q70223970': '800 Jahre (1230), Ortsteil von Wallhausen.',  # Hengstfeld
    'Q50445466': '800 Jahre (1230), Stadtteil von Lauda-Königshofen.',  # Sachsenflur
    'Q105847639': '750 Jahre (1280), Stadtteil von Weil am Rhein.',  # Friedlingen
    'Q798693': '700 Jahre (1330), Kurort, Ortsteil von Bad Peterstal-Griesbach.',  # Bad Griesbach im Schwarzwald
    'Q19980223': '1250 Jahre (um 780), Stadtteil von Tauberbischofsheim.',  # Hochhausen
    'Q1694501': '100 Jahre St. Raphael-Schulen Heidelberg.',  # St. Raphael-Schulen Heidelberg
    'Q1262922': '100 Jahre Mona Versand in Pforzheim.',  # Mona Versand
    'Q26794510': '150 Jahre VR-Bank in Mittelbaden. Fusionen prüfen.',  # VR-Bank in Mittelbaden
    'Q16857824': '150 Jahre Volksbank Kirnau. Fusionen prüfen.',  # Volksbank Kirnau
    'Q1569797': '150 Jahre Kurfürst-Ruprecht-Gymnasium Neustadt (Pfalz).',  # Kurfürst-Ruprecht-Gymnasium Neustadt
    'Q1324579': '125 Jahre Goldschmiedeschule mit Uhrmacherschule Pforzheim – passt zur Schmuckstadt, bildstarkes Thema.',  # Goldschmiedeschule mit Uhrmacherschule Pforzheim
    'Q682489': '125 Jahre Reuchlin-Gymnasium Pforzheim.',  # Reuchlin-Gymnasium Pforzheim
    'Q106560674': '100 Jahre AURORA in Mudau (gehört zur Indus Holding).',  # AURORA Konrad G. Schulz GmbH & Co. KG
    'Q107102723': '100 Jahre Katz & Klumpp in Gernsbach.',  # Katz & Klumpp
    'Q2641625': '200 Jahre Alexander-von-Humboldt-Gymnasium Konstanz.',  # Alexander-von-Humboldt-Gymnasium Konstanz
    'Q106905625': '100 Jahre Karl-Maybach-Gymnasium Friedrichshafen.',  # Karl-Maybach-Gymnasium
    'Q1645492': '100 Jahre Urspringschule (Internat) in Schelklingen.',  # Urspringschule
    'Q18677538': '125 Jahre Bahnhof-Hotel Buchen.',  # Bahnhof-Hotel Buchen
    'Q116779466': '125 Jahre Felss in Königsbach-Stein (Enzkreis).',  # Felss Group
    'Q107139026': '125 Jahre Karlsruher Schiffahrts-AG.',  # Karlsruher Schiffahrts-Aktiengesellschaft
    'Q21039726': '100 Jahre Bott in Gaildorf.',  # Bott
    'Q18287000': '100 Jahre Galerie Valentien in Stuttgart.',  # Galerie Valentien
    'Q74821641': '100 Jahre Gottlob Brodbeck in Metzingen.',  # Gottlob Brodbeck GmbH & Co. KG
    'Q473339': '100 Jahre Graupner Modellbau.',  # Graupner Modellbau
    'Q1343869': '100 Jahre Stadtwerke Tübingen.',  # Stadtwerke Tübingen
    'Q2584554': '100 Jahre Winkels in Sachsenheim.',  # Winkels
    'Q101131482': '150 Jahre Volksbank Sulmtal. Fusionen prüfen.',  # Volksbank Sulmtal
    'Q111731279': '250 Jahre Staatsschule für Gartenbau Stuttgart-Hohenheim.',  # Staatsschule für Gartenbau
    'Q135653480': '175 Jahre Erasmus-Widmann-Gymnasium Schwäbisch Hall.',  # Erasmus-Widmann-Gymnasium
    'Q317125': '175 Jahre Hochschule Reutlingen.',  # Hochschule Reutlingen
    'Q121337900': '150 Jahre Eugen-Bolz-Gymnasium Rottenburg.',  # Eugen-Bolz-Gymnasium Rottenburg
    'Q1558156': '150 Jahre Gymnasium Korntal-Münchingen.',  # Gymnasium Korntal-Münchingen
    'Q1515507': '125 Jahre Gertrud-Luckner-Gewerbeschule Freiburg.',  # Gertrud-Luckner-Gewerbeschule
    'Q65239480': '125 Jahre Kaufmännische Schule 1 Stuttgart.',  # Kaufmännische Schule 1
    'Q130973425': '100 Jahre RAPID Maschinenbau.',  # RAPID Maschinenbau
    'Q106676426': '100 Jahre Rüdinger Spedition in Krautheim.',  # Rüdinger Spedition
    'Q138000577': '100 Jahre W. Gessmann in Leingarten.',  # W. Gessmann GmbH
    'Q105728710': '150 Jahre Sülzle Gruppe in Rosenfeld.',  # Sülzle Gruppe
    'Q1782066': '125 Jahre Konrad Hornschuch AG.',  # Konrad Hornschuch AG
    'Q1914433': '125 Jahre Mayer & Cie. in Albstadt.',  # Mayer & Cie.
    'Q850250': '75 Jahre Betty Barclay in Nußloch – Verlagsgebiet.',  # Betty Barclay
    'Q1739924': '75 Jahre Nolte-Möbel in Germersheim – Pfalz.',  # Nolte-Möbel
    'Q331492': '75 Jahre Weidenhammer in Hockenheim – Verlagsgebiet.',  # Weidenhammer Packaging Group
    'Q21559758': '75 Jahre Weisenburger Bau in Rastatt – Verlagsgebiet.',  # Weisenburger Bau
    'Q131819633': '125 Jahre Kunsthaus Bühler in Stuttgart.',  # Kunsthaus Bühler
    'Q140285668': '125 Jahre REVEN in Sersheim.',  # REVEN
    'Q2384949': '150 Jahre TSG Haßloch (Pfalz), großer Verein.',  # TSG Haßloch
    'Q120793245': '150 Jahre TV 1880 Käfertal (Mannheim).',  # TV 1880 Käfertal
    'Q2640881': '125 Jahre Alemannia Eggenstein (Landkreis Karlsruhe).',  # Alemannia Eggenstein
    'Q442055': '125 Jahre SG 05 Pirmasens (Pfalz).',  # SG 05 Pirmasens
    'Q690818': '125 Jahre TV Busenbach (Waldbronn, Landkreis Karlsruhe).',  # TV Busenbach
    'Q2521431': '125 Jahre VfB 05 Knielingen (Karlsruhe).',  # VfB 05 Knielingen
    'Q71147104': '125 Jahre VfR Friesenheim (Ludwigshafen, Pfalz).',  # VfR Friesenheim
    'Q799100': '75 Jahre Badische Stahlwerke in Kehl.',  # Badische Stahlwerke
    'Q1205813': '75 Jahre Deutsches Literaturarchiv Marbach – Institution, macht das vermutlich selbst.',  # Deutsches Literaturarchiv Marbach
    'Q19968199': '75 Jahre Hegau-Bibliothek Singen.',  # Hegau Bibliothek
    'Q833959': '75 Jahre Sto in Stühlingen.',  # Sto
    'Q50194071': '125 Jahre Sektion Reutlingen des Alpenvereins – Sektionen machen traditionell eine Festschrift.',  # Sektion Reutlingen des Deutschen Alpenvereins
    'Q654131': '1250 Jahre (Lorscher Codex 780), aber nur 330 Einwohner. Rheinhessen, Nachbargebiet.',  # Bermersheim
    'Q1743604': '900 Jahre (Lorscher Codex 1130), Stadtteil von Heppenheim. Nachbargebiet.',  # Kirschhausen
    'Q2008838': '900 Jahre (Lorscher Codex 1130), Ortsteil von Mörlenbach. Nachbargebiet.',  # Ober-Mumbach
    'Q2307319': '200 Jahre Sparkasse Bensheim. Nachbargebiet.',  # Sparkasse Bensheim
    'Q2307691': '200 Jahre Sparkasse Starkenburg in Heppenheim. Nachbargebiet.',  # Sparkasse Starkenburg
    'Q24288521': '150 Jahre Staatliches Aufbaugymnasium Alzey. Nachbargebiet.',  # Staatliches Aufbaugymnasium Alzey
    'Q702674': '150 Jahre Wella in Darmstadt – heute Konzern, macht es vermutlich selbst. Nachbargebiet.',  # Wella
    'Q1742261': '125 Jahre Kinopolis in Darmstadt, Familienunternehmen. Nachbargebiet.',  # Kinopolis
    'Q114045662': '125 Jahre Historischer Verein Ingelheim – Geschichtsvereine sind Kernkundschaft. Nachbargebiet.',  # Historischer Verein Ingelheim
    # Bayerische Grenzkreise
    'Q502441': '900 Jahre (1130 „Windesbach“). Stadt mit 6 100 Einwohnern im Kreis Ansbach, bekannt durch den Windsbacher Knabenchor. Sauberer Beleg, eine Stadt mit Namen und Kultur – bester Orts-Lead in den bayerischen Grenzkreisen 2030; nur die Entfernung zum Verlag spricht gegen S.',  # Windsbach
    'Q21084038': '1000 Jahre (1030 erstmals belegt). Hauptort des Marktes Jettingen-Scheppach im Kreis Günzburg, der 1970 aus Jettingen und Scheppach entstand. Tausend Jahre sind ein Anlass, den ein Markt groß feiert – Ansprechpartner ist der Markt.',  # Jettingen
    'Q1676521': '75 Jahre laut Wikidata (1955). Sonnenschutz-Hersteller in Marktheidenfeld, Familienunternehmen mit mehreren tausend Mitarbeitern (aus dem Gedächtnis). 75 ist keine große Zahl, aber die Firma hat Geld und Familiengeschichte.',  # Warema Renkhoff SE
    'Q489685': '„1600 Jahre“ nur über die Erwähnung von Guntia in der Notitia dignitatum „um 430“ – kein festes Jahr, und die Stadt zählt eher ab dem Römerkastell (77/78). Große Kreisstadt mit 21 000 Einwohnern. Mitnehmen, aber nicht mit der Zahl 1600 werben.',  # Günzburg
    'Q505306': '1200 Jahre, wenn man den Anfang des Lorscher Urbars (830–850) nimmt – das Jahr ist nur der Beginn eines Zeitraums. Gemeinde mit 2 500 Einwohnern. Fragen, welches Jahr die Gemeinde selbst zählt.',  # Obernzenn
    'Q758371': '750 Jahre (1280), ehemals selbständige Gemeinde, heute Stadtteil von Illertissen (Kreis Neu-Ulm). Ansprechpartner die Stadt Illertissen.',  # Au
    'Q22692809': '150 Jahre laut Wikidata (1880). Gymnasium in Neu-Ulm. Schulen feiern solche Zahlen gern mit Festschrift; kleines Budget.',  # Lessing-Gymnasium
}

REIHENFOLGE = [
    'Q16102',  # Kornwestheim
    'Q522165',  # Wörth am Rhein
    'Q3758',  # Kaiserslautern
    'Q566832',  # Mack Rides GmbH & Co KG
    'Q816360',  # Bembé
    'Q662258',  # Limburgerhof
    'Q1933065',  # Plittersdorf
    'Q552671',  # Waldmohr
    'Q545371',  # Au am Rhein
    'Q442426',  # Altleiningen
    'Q10764833',  # Mußbach an der Weinstraße
    'Q23559805',  # Böhl
    'Q435044',  # Altburg (Calw)
    'Q61713',  # Niederstetten
    'Q60136',  # Erlenbach
    'Q503260',  # Talheim
    'Q19921577',  # Karolinen-Gymnasium
    'Q1550425',  # Donauschule Nendingen
    'Q15785111',  # Badische Maschinenfabrik Durlach
    'Q712034',  # Schwarz-Gruppe
    'Q1959244',  # Münzing Chemie
    'Q151137',  # Steiff
    'Q284672',  # Bilfinger SE
    'Q1929771',  # Michael Weinig AG
    'Q37068918',  # Kreissparkasse Heidenheim
    'Q1441291',  # Schwarzwaldmilch
    'Q637707',  # Lamy
    'Q2973',  # Darmstadt
    'Q656289',  # Guntersblum
    'Q13487057',  # Mörfelden
    'Q105254',  # 1. FSV Mainz 05
    'Q107368376',  # A. Racke
    'Q2110996',  # Privat-Brauerei Schmucker
    'Q502441',  # Windsbach
    'Q21084038',  # Jettingen
    'Q1676521',  # Warema Renkhoff SE
    'Q489685',  # Günzburg
    'Q505306',  # Obernzenn
    'Q758371',  # Au
    'Q22692809',  # Lessing-Gymnasium
]
