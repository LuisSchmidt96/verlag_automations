# Anleitung: Pressekontakte einschätzen

Gilt für jeden, der ein Paket aus `presseverteiler/einschaetzung/` abarbeitet,
ob Mensch, Claude oder Hilfs-Agent. Ablauf und Skript: `einschaetzung.py`.

## Worum es geht

Der **Verlag Regionalkultur** (Ubstadt-Weiher bei Bruchsal) macht Orts- und
Regionalgeschichte, Ortschroniken, Ortssippenbücher, Heimat- und Landeskunde,
Kunst und Kultur der Region, Natur- und Wanderführer, Mundart, Regionalkrimis
und Festschriften. Der Schwerpunkt liegt auf Baden (Karlsruhe, Kraichgau,
Kurpfalz), der Pfalz und Württemberg. Daneben kommen Südhessen, Franken und
das Elsass vor.

Für jedes neue Buch wählt der Pressereferent aus rund 1.900 Pressekontakten
die passenden aus. Dafür braucht jeder Kontakt einen **Vorschlag**, wofür er
zuständig ist. Der Pressereferent bestätigt oder korrigiert danach jeden
Vorschlag. Ein guter Vorschlag lässt sich mit einem Blick bestätigen. Die
Begründung muss ihm deshalb sagen, *was* der Kontakt ist.

## Die Auswahlregel — sie bestimmt, was „richtig“ ist

* Jeder Kontakt hat **Regionen**, **Themen** und das Feld **immer**.
* Ein Buch bekommt ebenfalls Regionen und Themen.
* Ein Kontakt bekommt das Buch, wenn **jede ausgefüllte Achse passt**. Eine
  leere Achse zählt als „egal“.
  * Nur Regionen eingetragen: bekommt jedes Buch aus diesen Regionen,
    egal zu welchem Thema.
  * Nur Themen eingetragen: bekommt jedes Buch zu diesen Themen, egal woher.
  * Beides eingetragen: bekommt nur Bücher, die in beiden passen.
  * Nichts eingetragen: bekommt **nie** etwas.
* Schlagworte gelten **nach unten**. Wer „Baden“ hat, bekommt auch Bücher über
  den Kraichgau. Wer „Kraichgau“ hat, bekommt kein Buch über Baden insgesamt.
  Für Themen gilt dasselbe: Wer „Geschichte“ hat, bekommt auch die
  Kirchengeschichte. Wer „Kirchengeschichte“ hat, bekommt nicht jedes
  Geschichtsbuch.

Daraus folgt:

* **Zu breit ist der teure Fehler.** Eine Lokalzeitung in Freiburg mit
  „Baden“ bekäme jede Ortschronik aus dem Kraichgau. Nimm das Schlagwort, das
  das Verbreitungs- oder Berichtsgebiet **am genauesten** trifft. Bei
  mehreren Gebieten trägst du mehrere ein.
* Deckt ein Medium mehrere Gebiete ab, trägst du **mehrere** ein, statt ein
  großes zu nehmen. Badische Zeitung gesamt → Freiburg / Breisgau; Ortenau;
  Landkreis Lörrach; Landkreis Waldshut. Ein Schlagwort „Südbaden“ gibt es
  nicht.
* Liegt das Gebiet **in Baden-Württemberg oder der Pfalz** und hat kein
  eigenes Schlagwort, nimm das nächstgrößere und setze `unsicher`.
  **Außerhalb** davon (Nahe, Mosel, Mittel- und Nordhessen, Bayern außer
  Franken) gilt: außerhalb des Verlagsgebiets, siehe Punkt 6.
* **Themen nur, wenn der Kontakt sich wirklich darauf beschränkt.** Eine
  Lokalzeitung berichtet über jedes Buch aus ihrer Gegend, sie bekommt nur
  Regionen. Ein landeskundliches Blatt, das fast alles aus seiner Region
  bringt, bekommt ebenfalls nur Regionen.
* Ortsgeschichte gehört **nicht** zu Geschichte. Ein Geschichtsverein vor Ort
  bekommt beides: „Geschichte“ und „Ortsgeschichte“.

## Was für welche Art von Kontakt gilt

1. **Lokal- und Regionalmedien** (Tageszeitung samt Lokalredaktion,
   Anzeigen- und Wochenblatt, Amtsblatt, Stadtmagazin, Lokalradio,
   Regionalfernsehen): **nur Regionen**, nämlich ihr Verbreitungsgebiet.
   * Eine Lokalredaktion bekommt das Gebiet **ihrer Ausgabe**, nicht das der
     ganzen Zeitung. Beispiele: BNN Redaktion Bretten → Kraichgau.
     Fränkische Nachrichten Buchen → Neckar-Odenwald-Kreis; Odenwald.
   * Zentral-, Kultur- und Chefredaktion einer Regionalzeitung bekommen das
     **ganze** Verbreitungsgebiet. Beispiele: BNN Kultur → Karlsruhe (Region).
     Die Rheinpfalz, Kultur → Pfalz.
   * Sportredaktion: Region und „Sport“. Wirtschaftsredaktion: Region und
     „Technik und Wirtschaft“. Politik-, Nachrichten- und Onlineressort:
     Region wie die Zeitung.
   * Amtsblatt einer Gemeinde: das engste Schlagwort, in dem die Gemeinde
     liegt. Die Gemeinde gehört in die Begründung („Amtsblatt Östringen“).
     Amtsblätter werden oft woanders gedruckt (Nussbaum, Wittich). Es zählt
     der Ort im Titel, nicht der Ort des Verlags.
2. **Rundfunk.** SWR-Studio → Gebiet des Studios (Studio Karlsruhe →
   Karlsruhe (Region)). Landesweite Redaktion → Baden-Württemberg bzw.
   Rheinland-Pfalz. Ein Kulturprogramm oder eine Fachredaktion (SWR2
   Literatur) bekommt zusätzlich die Themen dieser Redaktion. Privatradios
   und Regionalfernsehen werden wie Lokalmedien behandelt.
3. **Fachmedien ohne Regionsbindung** (Fachzeitschrift, Verbandsblatt,
   Rezensionsorgan, wissenschaftliche Zeitschrift): **nur Themen**, und zwar
   das engste passende. Beispiele: Eisenbahn-Kurier → Eisenbahn. Archiv für
   Sozialgeschichte → Politik und Soziales, nicht Geschichte.
   * Allgemeine Geschichtszeitschriften (Historische Zeitschrift) →
     Geschichte und `unsicher`, denn sie bekämen jedes Geschichtsbuch.
     Fremdsprachige Fachzeitschriften (Annales, BABESCH) → leer, außer die
     Bemerkung sagt etwas anderes.
   * Regionale Kultur- und Veranstaltungsmagazine (Kulturkalender,
     Kunstmagazin einer Stadt) → Region und „Kunst und Musik“, eventuell
     auch „Belletristik / Literatur“. Lifestyle- und Stadtmagazine → nur
     Region.
   * „IfU“ nicht vergeben. Was es bedeutet, ist unklar. Nimm stattdessen
     „Technik und Wirtschaft“.
4. **Regionale Fachmedien** (Geschichts- und Heimatverein, landeskundliches
   Jahrbuch, Kirchenzeitung eines Bistums oder einer Landeskirche, regionale
   Natur- oder Literaturzeitschrift): **Region und Themen**. Beispiele:
   „der pilger“ (Bistum Speyer) → Pfalz; Kirchengeschichte. Konradsblatt
   (Erzbistum Freiburg) → Baden; Kirchengeschichte. Geschichtsverein
   Bruchsal → Landkreis Karlsruhe; Geschichte; Ortsgeschichte. Genealogische
   Vereine → „Ortssippenbuch (Familienbuch)“, mit Region, wenn sie regional
   sind.
5. **Überregionale Medien** (FAZ, SZ, Zeit, Deutschlandfunk, ARD, ZDF,
   Spiegel, Publikumszeitschriften): **keine Region**. Themen nur nach dem
   Ressort: Feuilleton oder Literaturkritik → Belletristik / Literatur.
   Geschichtsmagazin → Geschichte. Reiseressort → Reise- / Kultur- /
   Naturführer. Ist das Ressort unklar (Nachrichten, Politik, Hauptstadtbüro,
   Korrespondent), bleibt alles leer. Regionalkultur interessiert
   überregionale Medien selten, deshalb im Zweifel leer statt breit.
6. **Außerhalb des Verlagsgebiets** (Lokalpresse in NRW, Niedersachsen,
   Sachsen, Österreich …): **leer**, außer es gibt einen überregionalen
   Fachbezug. Das gilt auch für **Fachblätter, die an eine fremde Region
   gebunden sind** (Archäologie im Rheinland, Bayerische
   Vorgeschichtsblätter, Westfälische Archivpflege): Ein Buch über den
   Kraichgau bespricht dort niemand. Überregionale Fachblätter (DAI,
   Aschkenas) fallen dagegen unter Punkt 3. Solche Kontakte stehen oft wegen eines einzelnen Autors in der
   Liste. Die Begründung sagt das: „Lokalzeitung Gelsenkirchen, außerhalb
   des Verlagsgebiets“. Das ist eine sichere Einschätzung, also kein
   `unsicher`.
   * Bayernweite Medien (BR landesweit, Bayerns Kulturmagazine) → Franken,
     denn Franken ist der Teil Bayerns, der zum Verlagsgebiet gehört.
   * Schweiz: Basel und Nordschweiz → Schweiz, überregionale Schweizer
     Medien nur nach Thema. Elsass (DNA, L'Alsace) → Elsass-Lothringen.
7. **Einzelpersonen** (freie Journalisten, Rezensenten): Für wen sie
   schreiben, steht meist in der Bemerkung („schreibt für RNZ“). Dann gilt
   das Gebiet dieses Mediums bzw. das Ressort. Ohne Hinweis bleibt alles leer
   und `unsicher`.
8. **Pressestellen, Ämter, Bibliotheken, Verbände** ohne eigenes Medium:
   nach Zweck. Ist nichts erkennbar: leer und `unsicher`.

## Hinweise aus Access

* **Die Bemerkung geht vor.** Dort stehen Wünsche, die nirgends sonst stehen
  („möchte alle Kraichgau-Titel an ihn nach Bretten“). Setze sie um und
  erwähne sie in der Begründung.
* Die alten Felder **Access-Region** und **Access-Fachgebiet** stammen von
  früheren Mitarbeitern. Sie sind Hinweise, keine Wahrheit.
  * Abkürzungen: his = Geschichte, nat = Natur, lit = Literatur,
    theol = Kirche, klass/mus = Musik, kunst = Kunst, bühne = Theater,
    jud = Jüdische Geschichte, päd = Schule, gen = Genealogie,
    bahn = Eisenbahn, architektur = Architektur, allg = ohne
    Themenbeschränkung, reg = regional, südwest = Südwesten/BW.
  * **Access-Region ist oft eine alte Selektion, kein Verbreitungsgebiet.**
    „Odenwald“ steht beim Amtsblatt Mannheim, bei Darmstädter
    Anzeigenblättern und beim BR in Schweinfurt. Die waren einmal Empfänger
    eines Odenwald-Buchs. Entscheidend ist, wo das Medium erscheint.
    Widerspricht die Access-Region dem Erscheinungsort, gilt der
    Erscheinungsort. Die Access-Region kann als zusätzliche Region dazukommen,
    wenn sie plausibel ist (ein Darmstädter Blatt hat Hessen, das schließt den
    Odenwald ohnehin ein).
  * „Berlin“ als Region ist der Sitz einer Hauptstadtredaktion, kein
    Berichtsgebiet.
* Die **Mail-Domain** verrät oft das Medium (bnn.de, rnz.de, swr.de).

## immer

Setze **immer** nur, wenn die Bemerkung es ausdrücklich verlangt („alle
Titel“, „jede Neuerscheinung“). Sonst nie, auch nicht für dpa oder den SWR:
„immer“ hieße auch jedes Buch über das Elsass oder Franken.

## Orte → Schlagwort (für Einheitlichkeit)

| Orte | Schlagwort |
|---|---|
| Bretten, Kraichtal, Östringen, Bad Schönborn, Sinsheim, Eppingen, Bad Rappenau, Angelbachtal | Kraichgau |
| Pfinztal, Walzbachtal, Weingarten (Baden), Remchingen, Königsbach-Stein | Pfinzgau |
| Bruchsal, Ettlingen, Stutensee, Waghäusel, Philippsburg, Graben-Neudorf, Linkenheim, Rheinstetten, Malsch, Marxzell | Landkreis Karlsruhe |
| Karlsruhe (nur die Stadt) | Karlsruhe (Stadt) |
| Zeitung für Stadt und Landkreis Karlsruhe (BNN-Hauptausgabe) | Karlsruhe (Region) |
| Rastatt, Gaggenau, Gernsbach, Bühl, Kuppenheim, Murgtal | Landkreis Rastatt |
| Baden-Baden | Baden-Baden |
| Offenburg, Lahr, Achern, Oberkirch, Kehl, Haslach, Hausach, Wolfach, Kinzigtal | Ortenau |
| Pforzheim | Pforzheim |
| Mühlacker, Neuenbürg, Keltern, Birkenfeld, Ötisheim, Maulbronn | Enzkreis |
| Calw, Nagold, Altensteig, Bad Wildbad, Bad Herrenalb | Landkreis Calw |
| Freudenstadt, Horb | Landkreis Freudenstadt |
| Heidelberg | Heidelberg |
| Mannheim | Mannheim |
| Schwetzingen, Hockenheim, Wiesloch, Walldorf, St. Leon-Rot, Leimen, Ladenburg, Neckargemünd | Rhein-Neckar-Kreis |
| Weinheim, Schriesheim, Hemsbach | Rhein-Neckar-Kreis; Bergstraße |
| Eberbach | Rhein-Neckar-Kreis; Odenwald |
| Mosbach, Osterburken | Neckar-Odenwald-Kreis |
| Buchen, Walldürn | Neckar-Odenwald-Kreis; Odenwald |
| Kurpfalz als Berichtsgebiet (Heidelberg, Mannheim, Vorderpfalz zusammen) | Metropolregion Rhein-Neckar |
| Heilbronn (nur die Stadt) | Heilbronn (Stadt) |
| Weinsberg, Lauffen, Neckarsulm, Bad Friedrichshall, Bad Wimpfen, Leingarten | Landkreis Heilbronn |
| Brackenheim, Güglingen | Zabergäu |
| Künzelsau, Öhringen | Hohenlohekreis |
| Schwäbisch Hall, Crailsheim, Gaildorf | Landkreis Schwäbisch Hall |
| Bad Mergentheim, Tauberbischofsheim, Wertheim | Tauberfranken |
| Stuttgart (nur die Stadt) | Stuttgart |
| Ludwigsburg, Bietigheim-Bissingen, Marbach, Vaihingen/Enz, Esslingen, Nürtingen, Kirchheim/Teck, Leinfelden-Echterdingen, Filderstadt, Böblingen, Sindelfingen, Herrenberg, Leonberg, Waiblingen, Backnang, Schorndorf, Göppingen, Geislingen | Region Stuttgart |
| Tübingen, Rottenburg, Mössingen | Landkreis Tübingen |
| Reutlingen, Metzingen, Pfullingen, Münsingen, Bad Urach | Landkreis Reutlingen |
| Balingen, Hechingen, Albstadt | Zollernalbkreis |
| Rottweil, Oberndorf, Schramberg | Landkreis Rottweil |
| Villingen-Schwenningen, Donaueschingen, St. Georgen, Triberg | Schwarzwald-Baar-Kreis |
| Tuttlingen, Spaichingen | Württemberg (unsicher) |
| Konstanz, Singen, Radolfzell, Stockach, Engen | Landkreis Konstanz |
| Friedrichshafen, Überlingen, Markdorf, Tettnang, Meersburg | Bodenseekreis |
| Ravensburg, Weingarten, Biberach, Bad Waldsee, Bad Saulgau, Sigmaringen, Riedlingen, Laupheim, Pfullendorf, Meßkirch | Oberschwaben |
| Ulm, Ehingen, Laichingen | Ulm / Alb-Donau |
| Aalen, Ellwangen, Schwäbisch Gmünd, Heidenheim | Ostwürttemberg |
| Freiburg, Emmendingen, Waldkirch, Breisach, Bad Krozingen, Titisee-Neustadt | Freiburg / Breisgau |
| Müllheim, Badenweiler | Markgräflerland |
| Lörrach, Weil am Rhein, Schopfheim, Rheinfelden | Landkreis Lörrach |
| Waldshut-Tiengen, Bad Säckingen, St. Blasien, Bonndorf | Landkreis Waldshut |
| Ludwigshafen | Ludwigshafen |
| Speyer | Speyer |
| Schifferstadt, Limburgerhof, Mutterstadt, Frankenthal | Rhein-Pfalz-Kreis |
| Neustadt/Weinstraße, Bad Dürkheim, Haßloch, Deidesheim, Edenkoben | Weinstraße |
| Landau, Germersheim, Kandel, Wörth, Bad Bergzabern, Herxheim, Annweiler | Südpfalz |
| Kaiserslautern, Kusel, Pirmasens, Zweibrücken, Landstuhl, Rockenhausen | Westpfalz |
| Grünstadt, Kirchheimbolanden | Pfalz |
| Die Rheinpfalz als Ganzes, pfalzweite Blätter | Pfalz |
| Mainz, Worms, Bingen, Alzey, Oppenheim | Rheinhessen |
| Bad Kreuznach, Idar-Oberstein, Koblenz, Trier, Mayen | leer (außerhalb des Verlagsgebiets) |
| landesweite Medien in Rheinland-Pfalz (SWR Mainz landesweit, Staatskanzlei) | Rheinland-Pfalz |
| Bensheim, Heppenheim, Lorsch, Viernheim, Lampertheim, Bürstadt | Bergstraße |
| Erbach, Michelstadt, Fürth (Odenwald), Wald-Michelbach, Bad König | Odenwald |
| Darmstadt, Groß-Gerau, Dieburg, Frankfurt, Wiesbaden, Offenbach, Rüsselsheim | Hessen |
| Gießen, Marburg, Fulda, Kassel | leer (außerhalb des Verlagsgebiets) |
| Würzburg, Aschaffenburg, Miltenberg, Schweinfurt, Kitzingen, Marktheidenfeld | Unterfranken |
| Nürnberg, Erlangen, Forchheim, Ansbach, Bamberg, Rothenburg | Franken |
| Saarbrücken und das übrige Saarland | Saarland |
| Strasbourg, Wissembourg, Haguenau, Colmar | Elsass-Lothringen |
| Basel, Zürich, Luzern, Chur | Schweiz |

Nicht genannte Orte ordnest du sinngemäß zu. Landesweite Medien in
Baden-Württemberg (Staatsanzeiger, Landeszentrale, dpa Südwest) bekommen
Baden-Württemberg.

## Medienart und Ressort

Je einer der Werte aus `einschaetzung/schlagworte.md`.

* **Medienart:**
  * Tageszeitung, auch ihre Lokal- und Fachredaktionen.
  * Wochenblatt: Anzeigenblatt, Wochenzeitung, Stadtteilblatt.
  * Amtsblatt: Amts- und Mitteilungsblatt einer Gemeinde.
  * Magazin: Publikums-, Stadt-, Kultur- und Mitgliedermagazin.
  * Fachzeitschrift: auch Jahrbuch, Vereins- und Verbandsblatt,
    Kirchenzeitung, wissenschaftliche Zeitschrift.
  * Hörfunk, Fernsehen.
  * Online: nur online.
  * Agentur: dpa, epd, KNA.
  * Person: freie Journalisten und Rezensenten ohne festes Medium.
  * Sonstige: Pressestelle, Behörde, Bibliothek.
* Fachzeitschrift oder Magazin? Eine Fachzeitschrift richtet sich an ein
  Fachpublikum oder an Mitglieder (wissenschaftlich, Verband, Verein,
  Jahrbuch). Ein Magazin richtet sich an alle, am Kiosk oder kostenlos. Auch
  Literatur- und Kulturzeitschriften für ein allgemeines Publikum sind
  Magazine.
* Ein Anzeigenblatt, das nach Titel und Druckerei (Nussbaum, Wittich) das
  Mitteilungsblatt einer Gemeinde ist, ist ein Amtsblatt, auch wenn Access
  „anzeigenblatt“ sagt.
* **Ressort:**
  * Lokales: Lokal- und Stadtredaktion, auch die bloße „Redaktion“ einer
    Lokalausgabe oder einer Lokalzeitung.
  * Kultur: Feuilleton, Kulturredaktion.
  * Rezensionen: Buchbesprechung, Literatur, „Einreichungen“.
  * Chefredaktion: auch Herausgeber, Redaktionsleitung, Schriftleitung.
  * Redaktion allgemein: nur „Redaktion“. Das ist der Normalfall bei
    kleinen Blättern und Zeitschriften.
  * anderes Ressort: Sport, Wirtschaft, Politik, Wissen …
  * Leer, wenn nicht erkennbar.

## unsicher

Setze `"unsicher": true`, wenn

* du das Medium nicht kennst und aus Name und Ort rätst,
* das Verbreitungsgebiet unklar ist,
* die Access-Werte deiner Einschätzung widersprechen und du nicht weißt,
  wer recht hat, oder
* es mehrere Lesarten gibt.

Nicht unsicher sind eine eindeutige Lokalzeitung, ein eindeutiges Fachblatt
und ein Kontakt, der eindeutig außerhalb des Verlagsgebiets liegt.

**Nicht erfinden.** Kennst du ein Medium nicht, schreib das in die
Begründung („Medium unbekannt, der Name deutet auf …“).

## Ergebnis

Eine Datei `einschaetzung/ergebnis_<nr>.jsonl` mit **genau einer Zeile je ID**
des Pakets, in der Reihenfolge des Pakets, UTF-8:

```
{"id": 1234, "regionen": ["Kraichgau"], "themen": [], "immer": false, "medienart": "Tageszeitung", "ressort": "Lokales", "unsicher": false, "grund": "Lokalredaktion der BNN in Bretten"}
```

* `regionen` und `themen`: Listen mit Schlagworten, exakt geschrieben wie in
  `schlagworte.md`. Die Liste darf leer sein.
* `grund`: ein kurzer Satz, höchstens etwa 120 Zeichen. Er sagt, was der
  Kontakt ist und warum diese Schlagworte. **Keine Personennamen.**
* Schreib die Datei in Teilen von etwa 50 Zeilen: den ersten Teil mit Write,
  die weiteren mit `cat >> datei <<'EOF'` angehängt. So geht bei einem
  Abbruch nichts verloren.
* Danach `python -m presseverteiler.einschaetzung pruefen` aufrufen (im
  Repo-Wurzelordner) und alle Fehler im eigenen Paket beheben.
