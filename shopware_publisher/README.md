# Shopware-Publisher

Legt aus einer **VLB-ONIX-XML** (derselbe Datensatz, den auch der
`pi_bi_generator` nutzt) ein **Shopware-6-Produkt als Entwurf** an — inklusive
Cover-Bild, Preis, Beschreibung und Fakten.

**Entwurf heißt: `active = false`.** Das Produkt ist im Shop **nicht sichtbar**,
bis es im Admin geprüft und aktiv geschaltet wird.

Getestet gegen die Admin-API von **Shopware 6.7**.

## Was übernommen wird

| Shopware-Feld  | Quelle (ONIX)                                        |
|----------------|------------------------------------------------------|
| Artikelnummer  | ISBN-13 mit Bindestrichen (auch als **EAN**)          |
| Name           | **nur der Titel** — ohne Band, ohne Reihe             |
| Untertitel (attr1) | ONIX-Untertitel                                   |
| Format (attr12)| **Breite × Höhe** + Einband („17 x 23,5 cm, fester Einband“) |
| Beschreibung   | Werbetext + kursiver Fakten-Block (siehe unten)       |
| Preis          | DE-Preis brutto; **netto** = brutto / (1 + Steuersatz)|
| Maße / Erscheinungsdatum | `Measure` bzw. `PublishingDate`             |
| Gewicht        | `Measure` 08 — fehlt fast immer, dann **geschätzt** (s. u.) |
| SEO            | metaTitle, metaDescription, keywords                   |
| Bilder         | drei Quellen (siehe unten)                            |
| Kategorien     | **je Buch** in der Oberfläche gewählt                  |
| Steuer/Währung/Hersteller | einmalig im GUI zugeordnet                 |

### Beschreibung — der Hausstil

Abgelesen an den gepflegten Produkten im Livesystem, nicht erfunden:

```
<Werbetext, Absätze durch <br><br> getrennt>

<i>Heiko Brohm, Harald Stockert, Die GBG in Mannheim. 100 Jahre in 100 Geschichten.<br>
448 Seiten mit 390 Farb- und Schwarz-Weiß-Abbildungen, fester Einband.<br>
ISBN 978-3-95505-607-0. EUR 29,80.</i>
```

Zwei Eigenheiten, die man sonst wieder „korrigiert“:

* Der **eigene Verlag wird nicht genannt**. „verlag regionalkultur. 2026.“
  steht bei keinem Bestandsprodukt — die Zeile taucht nur bei Fremdimprints auf.
* **Reihe und Band stehen nirgends** im Shop, auch nicht im Namen. Der
  Untertitel steht in der Zitatzeile und im Untertitelfeld.

**Idempotent:** Die Produkt-ID wird fest aus der ISBN abgeleitet. Ein zweiter
Lauf zum selben Buch **aktualisiert** das Produkt — es entsteht kein Duplikat.

## Einmalige Einrichtung: Integration anlegen

Im Shop-Admin:

1. **Einstellungen → System → Integrationen → „Integration hinzufügen“**
2. Name vergeben, Rolle **Administrator** zuweisen, speichern.
3. **Zugriffsschlüssel-ID** und **Geheimer Zugriffsschlüssel** kopieren
   (das Secret wird nur **einmal** angezeigt!).
4. Im Tool oben eintragen → **Verbinden**.

Danach die Auswahlfelder füllen: **Steuer** (Bücher = 7 %), **Währung** (EUR),
optional **Kategorie** und **Hersteller**. Die Auswahl landet in der
`config.json` und wird beim nächsten Start wiederverwendet.

### Master-Passwort (Secret-Schutz)

Das Secret gibt vollen Zugriff auf die Admin-API — es liegt deshalb **nie im
Klartext** auf der Platte:

- **„Secret setzen…“** fragt das Shopware-Secret ab und ein selbst gewähltes
  **Master-Passwort**. Aus dem Passwort wird per **scrypt** ein Schlüssel
  abgeleitet, mit dem das Secret **AES-GCM-verschlüsselt** in der `config.json`
  landet (`secret_enc` + `kdf_salt`).
- Bei **jedem Start** fragt das Tool das Master-Passwort und entschlüsselt das
  Secret **nur im Arbeitsspeicher**.
- Falsches Passwort → kein Shop-Zugriff. Wer die `config.json` kopiert, hat nur
  Chiffretext; das Passwort steht nirgends (auch kein Hash davon wird gebraucht —
  AES-GCM ist authentifiziert und scheitert bei falschem Schlüssel von selbst).

> **Least Privilege:** Die Integration braucht keine *Administrator*-Rolle, und
> mit ihr steht ein geleaktes Secret sofort an den Kunden- und Bestelldaten. Es
> reicht eine eigene Rolle mit
>
> * **Schreibrechten** auf `product`, `category` und `media`,
> * **Leserechten** auf `tax`, `currency`, `sales_channel`,
>   `product_manufacturer` und `media_folder`.
>
> **Kategorie braucht Schreibrechte**, nicht nur Leserechte: `kategorie_anlegen`
> legt fehlende Autoren-/Herausgeberkategorien an (im Buchdurchgang nach
> Rückfrage). Hier stand früher „Leserechte auf Steuer/Währung/Kategorie" — das
> war zu eng, seit das Werkzeug Kategorien anlegen kann.
>
> Die Rolle zuerst am **Dev-Store** ausprobieren und ein Buch als Dry-Run
> durchschieben: fehlt ein Recht, nennt der 403 die Entität, und man weitet
> genau diese.

### Zugangsdatei auf dem Share

Zugangsdaten werden **nicht mehr im Werkzeug eingetragen**. Sie stehen in einer
Datei auf dem Share:

```
\\VR-Archiv\VR-Austausch\VR-Tools\zugang.txt
```

Eine Zeile je Abschnitt, mit Strichpunkt getrennt. Leerzeilen und Zeilen ab `#`
werden übersprungen:

```
# VR-Tools Zugang. Ohne Master-Passwort nutzlos.
prod;SWIAOHHLC09YWFIWDFEZCG10SW;<blob>
dev;SWIAR3RLUMTLM2PVZGK1T2R2BW;<blob>
sftp;sftpuser;<blob>
```

| Spalte | Inhalt |
|---|---|
| 1 | Abschnitt: `dev` / `prod` (Umgebungsname) oder `sftp` |
| 2 | Kennung: Zugriffsschlüssel-ID bzw. SFTP-Benutzer — ein Benutzername, kein Geheimnis |
| 3 | Blob: `base64(salt + nonce + ciphertext)` |

**Das Salt steckt im Blob**, weil scrypt beim Entschlüsseln genau dasselbe Salt
braucht wie beim Verschlüsseln. Ein Chiffretext ohne Salt lässt sich nicht
öffnen. Statt einer vierten Spalte, die beim Weiterreichen verlorengeht, steht
alles in einem Feld: 16 Byte Salt, 12 Byte Nonce, Rest Chiffretext.

**Eine Zeile bauen** (aus der Repo-Wurzel, einmal je Abschnitt):

```
python -m shopware_publisher.zugang_zeile
```

Das Skript fragt Abschnitt, Kennung, Geheimnis und Master-Passwort ab, macht
die fertige Zeile zur Gegenprobe noch einmal auf und gibt sie aus. Für alle
Abschnitte **dasselbe** Master-Passwort nehmen — sonst muss man es je Schritt
neu eingeben.

**In der Datei steht nur, WOMIT man sich anmeldet — nie, WOHIN.** Die
Shop-Adressen stehen fest im Code. Das ist kein Schönheitsfehler: stünde
`shop_url` dort, könnte jeder, der die Datei beschreiben darf, sie auf einen
eigenen Server zeigen lassen und bekäme beim nächsten Verbinden die
Schlüssel-ID und das **entschlüsselte** Secret zugeschickt — der Token-Aufruf
geht an `shop_url`.

Weiteres Verhalten:

- **Gelesen, nicht gespiegelt.** Eine örtliche Kopie überlebt jede Änderung an
  der Freigabe. So verliert den Zugang, wer den Share nicht mehr lesen darf,
  und ein gewechseltes Secret gilt sofort für alle.
- **Fließt nicht zurück.** Beim Speichern wird alles wieder entfernt, was
  unverändert von dort kam — es landet nie in einer örtlichen `config.json`.
- **Kein Share, kein Problem.** Fehlt die Datei, startet das Werkzeug normal;
  nur der Shop- und der Presse-Schritt fehlen. Im Shop-Panel steht dann, was
  gefunden wurde und was nicht.
- **Krumme Zeilen werden überlesen**, nicht gemeldet — deshalb steht im Panel
  ein Bericht, welche Abschnitte gelesen wurden.

> **Die Zugriffsrechte auf dieser Datei sind der eigentliche Schutz**, nicht die
> Verschlüsselung — die ist die zweite Linie. Und das Master-Passwort gehört
> **nicht** auf denselben Share: sonst liegen Schloss und Schlüssel im selben
> Fach.

## Dev-Store hinter Caddy (Basic-Auth + TLS)

**Basic-Auth blockt die Admin-API.** Basic und Bearer teilen sich denselben
`Authorization`-Header — es kann nur eines von beiden gesendet werden. Steht vor
dem Dev-Store eine Basic-Auth, antwortet der Webserver mit 401, bevor Shopware
das Token überhaupt sieht. Das lässt sich **nicht** im Tool lösen, sondern nur
in der Server-Konfiguration.

**Nicht** einfach `/api/*` für alle freigeben: Dev und Produktivshop liegen auf
**demselben Server**. Was auf dem Dev-Store angreifbar ist, steht direkt neben
dem Livesystem. Die API deshalb nur von der **eigenen IP** durchlassen — im
**Dev-vHost** (Apache 2.4):

In der `.htaccess` des Dev-Stores (`public/.htaccess`, **oberhalb** der
`# BEGIN Shopware`-Marker — der Block dazwischen wird überschrieben):

```apache
SetEnvIf Request_URI ^/api(/|$) is_api=1

AuthType Basic
AuthName "Staging"
AuthUserFile /var/www/dev-shopware/.htpasswd

<RequireAny>
    <RequireAll>
        <RequireAny>
            Require env is_api
            Require env REDIRECT_is_api
        </RequireAny>
        # eigene IP / VPN-Netz
        Require ip 203.0.113.5
    </RequireAll>
    Require valid-user
</RequireAny>
```

Logik: **(ist `/api` UND aus dem eigenen Netz) ODER Basic-Auth-Benutzer.**
`Require ip` ist eine Autorisierung *ohne* Authentifizierung — von der eigenen IP
geht der Bearer-Token unangetastet durch, von überall sonst greift weiterhin die
Basic-Auth. Das Regex ist **verankert** (`^/api(/|$)`), sonst würde auch eine
Storefront-URL wie `/buecher/api-design` die Basic-Auth umgehen. `REDIRECT_is_api`
ist nötig, weil Shopwares Rewrite auf `index.php` intern umleitet und die
Umgebungsvariable dabei den Präfix bekommt.

> **Achtung:** Apache erlaubt **keine Kommentare am Zeilenende** — ein `#` hinter
> einer Direktive wird als Argument gelesen und wirft einen **500er**. Kommentare
> immer in eine eigene Zeile.

`.htaccess` wird bei jedem Request gelesen — **kein** Apache-Neustart nötig.

> **SSH-Tunnel ist hier keine gute Idee.** Dev und Prod sind namensbasierte
> vHosts auf demselben Apache — ein Tunnel auf `localhost:443` transportiert den
> Hostnamen nicht, Apache landet dann im **Default-vHost**. Im schlimmsten Fall
> schreibt man so in den **Produktivshop**.

**Immer prüfen, auf welchen Shop man schreibt.** Das Tool zeigt die Ziel-URL vor
dem Anlegen an und fragt nach. Vorher am besten einmal mit **Dry-Run** laufen.

**TLS:** Falls das Zertifikat des Dev-Stores nicht überprüfbar ist
(`CERTIFICATE_VERIFY_FAILED`, z. B. selbstsigniert), in der `config.json`

```json
"tls_pruefen": false
```

setzen — **nur für den Dev-Store**. Im Produktivshop bleibt es `true`.

## Gewicht — geschätzt am eigenen Bestand

Die VLB-ONIX liefert das Gewicht praktisch nie mit (`<Measure>` vom Typ 08
fehlt). Ohne Gewicht rechnet Shopware keine Versandkosten. Geraten wird es
deshalb nicht **frei**, sondern an den Produkten gelernt, die im Shop schon
gepflegt sind — die haben Gewicht, Breite und Höhe, und ihre Seitenzahl steht
in der Umfangzeile der Beschreibung („448 Seiten mit …“).

Gelernt wird das physikalische Modell, nicht ein Mittelwert:

```
Gewicht = Fläche × (Papier_g_qm × Blattzahl + Einband_g_qm)
```

Ein Buch ist ein Stapel Blätter (Seitenzahl / 2) plus Einband; beide Anteile
skalieren mit der Fläche. Die zwei gelernten Zahlen sind deshalb ablesbar
(z. B. „134 g/m² Papier, 976 g/m² Einband“) — geht eine Schätzung daneben,
sieht man am Modell, woran es liegt. Ein Mittelwert oder ein
Gramm-pro-Seite-Faktor könnte das nicht: er machte ein großformatiges dünnes
Buch so schwer wie einen kleinen dicken Band.

* **Je Einbandart getrennt** — gemessen am Bestand ist das der einzige
  Unterschied, der wirklich zählt: nach Einband gruppiert liegt die Schätzung
  im Median 6,0 % daneben, ohne Gruppen 8,9 %. Eine zusätzliche Trennung nach
  Format bringt nichts (6,1 %).
* **Die Einbandbezeichnung wird normalisiert** (`fester Einband` / `Broschur`).
  Das ist nicht Kosmetik: die ONIX sagt „kartoniert“, das Format-Feld im Shop
  sagt „Broschur“, daneben steht dort Freitext („fester Einband im
  repräsentativen Großformat“, „fester Einabnd“). Ohne Normalisierung fand ein
  kartoniertes Buch **gar keine** Gruppe und wurde mit dem Mischmodell
  geschätzt — gemessen 16,4 % daneben statt 10,4 %.
* **Robust gegen Ausreißer:** die Gerade ist ein Theil-Sen-Median aller
  Paar-Steigungen, keine kleinsten Quadrate (gemessen besser: 6,0 % statt
  6,7 %). Dazu ein zweiter Durchgang, der die Bestandsdatensätze weglässt, die
  im ersten über 40 % danebenlagen — das sind die falsch gepflegten Gewichte,
  nicht die schwierigen Bücher.
* **Gelernt wird beim Verbinden**, aber nur, wenn das gespeicherte Modell fehlt,
  älter als `gewichtsmodell_max_tage` (30) ist oder aus einer älteren
  Modellversion stammt (`GEWICHTSMODELL_VERSION` — ändert sich der Schnitt der
  Gruppen, wäre ein gespeichertes Modell sonst still weiter in Gebrauch). Es steht in der
  `config.json` unter `gewichtsmodell` — mit Stand, Probenzahl und der
  gemessenen Streuung je Gruppe.
* **Geschätzt wird im Veröffentlichen-Weg selbst** (`veroeffentliche`), nicht
  nur in der Oberfläche — und fehlt dort das Modell, wird es **an Ort und
  Stelle gelernt**, weil eine Verbindung ja gerade steht. Sonst bekäme der
  Buchdurchgang nie ein Gewicht: er ruft `core.veroeffentliche` direkt auf und
  hat eine eigene `config.json`, in der nie jemand „Verbinden“ gedrückt hat.
  Gelernt wird einmal, danach steht das Modell auch in dieser Config.
  Ein von Hand eingetragenes Gewicht schlägt die Schätzung immer; der
  übergebene Buchdatensatz wird dabei nicht verändert.
* **Ohne Seitenzahl oder Format wird nicht geschätzt.** Dann bleibt das Feld
  leer und daneben steht, warum — eine stille 0 wäre schlimmer als gar keine
  Angabe, weil im Shop dann ein falsches Versandgewicht stünde.

### Stimmt die Schätzung? — messen statt vermuten

Der Shop kennt zu jedem gepflegten Produkt das **echte** Gewicht. Damit lässt
sich die Schätzung prüfen:

```
python -m shopware_publisher.probelauf --gewichte
```

Das Modell lernt dabei an vier Fünfteln des Bestands und schätzt das letzte
Fünftel, das es nie gesehen hat (Kreuzvalidierung) — an denselben Produkten zu
messen, aus denen gelernt wurde, benotete sich selbst. Genau dieser Wert steht
später am Gewichtsfeld („typisch ±5,8 % daneben“); alle Zeilen landen in
`gewichte_pruefung.csv`.

**Stand vom 3.9.2026** (459 Produkte im Dev-Shop mit Gewicht, Maßen und
Seitenzahl):

| Gruppe          | Papier   | Einband    | Proben | typisch daneben |
|-----------------|----------|------------|--------|-----------------|
| fester Einband  | 125 g/m² | 6560 g/m²  | 256    | **5,8 %**       |
| Broschur        | 118 g/m² | 4040 g/m²  | 127    | **10,3 %**      |
| alle (Rückfall) | 126 g/m² | 5860 g/m²  | 400    | 7,6 %           |

Die Einbandwerte sind Buchdeckel **plus** Vorsatz und Rücken, auf die Buchfläche
gerechnet: bei 17 × 24 cm sind 6560 g/m² rund 270 g — die Größenordnung eines
gebundenen Deckels. Broschuren streuen deutlich stärker als gebundene Bücher;
das steht so auch am Gewichtsfeld und ist keine Schwäche der Rechnung, sondern
der Bestand (geheftete Hefte und dicke Klappenbroschuren in einer Gruppe).

### Nebengewinn: falsch gepflegte Gewichte im Shop

**51 der 459 Produkte** liegen über 40 % daneben — und in den meisten Fällen
irrt nicht die Schätzung, sondern der Shop: 0,150 kg bei 280 Seiten in
16,5 × 23,5 cm, 0,200 kg bei 132 Seiten in 22,2 × 28,2 cm, 9,000 kg bei
272 Seiten. Runde Platzhalterwerte (0,2 / 0,5 / 1,0 kg), die einmal eingetragen
wurden und seitdem echtes Porto kosten. `--gewichte` listet sie mit
Artikelnummer auf; sie gehören im Admin geprüft. Beim Lernen sind sie draußen.

Die **Schlagseite** in der Ausgabe ist der Median mit Vorzeichen: schätzt das
Modell im Schnitt zu schwer oder zu leicht? Die Ausreißerliste ist doppelt
nützlich — sie zeigt entweder die Grenzen des Modells oder ein Produkt, dessen
Gewicht im Shop falsch gepflegt ist.

Was das Modell gerade sagt, ohne Gegenprobe, zeigt der normale Probelauf:

```
python -m shopware_publisher.probelauf
```

## SEO-Felder

`metaTitle` ist der **blosse Titel** — so steht es im Bestand, auch bei Büchern
mit Untertitel; der Untertitel hat sein eigenes Feld (attr1). `keywords` sind
die **Nachnamen** der Beteiligten plus Titel und Reihe, wie im Bestand
(„Brandes, Kinderbuch, Grünes Gras erzähl mir was, Dilsberg“ — die Sachbegriffe
darin trägt ein Mensch nach).

Zwei Eigenheiten, die man sonst wieder einbaut:

* **Eine Körperschaft hat keinen Nachnamen.** Ist „Stiftung Geißstraße“ die
  Herausgeberin, stand in den Schlüsselwörtern vorher nur der Titel — der
  Beteiligte fiel ersatzlos weg. Jetzt greift der ganze Name.
* **Gekürzt wird an der Satz-, sonst an der Wortgrenze.** Shopware nimmt
  255 Zeichen; ein harter Schnitt endet mitten im Wort („… verfolgt auch
  kritisch die Ve“). So steht es bei den migrierten Bestandsprodukten, und in
  der Google-Vorschau sieht es aus wie ein Fehler. Passt ein ganzer Satz, endet
  die Beschreibung damit; sonst am letzten Wort mit „…“. Ein Punkt nach einer
  Ziffer zählt dabei **nicht** als Satzende, sonst bricht der Text bei
  „zum ausgehenden 18.“ ab.

## Bilder — drei Quellen

Gesucht wird in dieser Reihenfolge; die Vorschau nennt, **welche** gegriffen hat:

1. **Artikeldaten-Share** — dort erzeugt `cover_previews` die Dateien. Das ist
   das Original, alles andere sind Kopien davon.

   ```
   <Artikeldaten>\<Kurzcode>_<Titel>\
       2D_72_<sc>.jpg     -> Cover (Hauptbild)
       3D_72_<sc>.jpg     -> zusätzliches Galeriebild
   ```

2. **Neben der ONIX-Datei** — hilft, wenn der Share nicht erreichbar ist.
3. **Webserver** — `https://verlag-regionalkultur.de/newsletter_/<sc>.png`,
   dieselbe Datei, die auch der Newsletter nutzt. Volle Auflösung.

Findet keine davon etwas, **fragt das Werkzeug nach**, bevor es anlegt — ein
Buch ohne Titelbild ist im Shop kaum zu gebrauchen. Mit **„Bild wählen…“** lässt
sich jederzeit eines von der Platte nehmen.

Die Medien-ID hängt an der **Rolle** (`cover`, `galerie0`), nicht am Dateinamen.
Nur so ersetzt ein zweiter Lauf ein korrigiertes Cover, statt einen zweiten
Mediendatensatz anzulegen und den alten als Waise zurückzulassen.

> **Achtung, Schiefstand:** dasselbe Cover liegt am Ende im Share, unter
> `newsletter_/`, bei VLB und in Shopware. Diese Ablagen wissen nichts
> voneinander. Wird ein Cover korrigiert, muss es **überall** neu hoch — der
> Publisher aktualisiert nur den Shop.

## Bedienung

1. **Verbinden** (einmalig, s. o.).
2. **ONIX-XML wählen** → die Vorschau zeigt, was im Shop landet (Name,
   Artikelnummer, Preis brutto/netto, Titelbild **mit Quelle**, Kategorien,
   Beschreibung). Beide ONIX-Fassungen werden gelesen: Kurz-Tags (`<b012>`) und
   Referenz-Tags (`<ProductForm>`, Dateiname `onix3Ref_…`).
3. **Kategorien wählen** — je beteiligter Person wird die passende Kategorie
   im Shop **gesucht** und angehakt; Sachkategorien wählt ein Mensch dazu.
   Ohne Kategorie hat das Buch im Shop keinen Breadcrumb.

   Der Shop führt **eine Kategorie pro Person**, benannt `Nachname, Vorname`
   (`Wiegand, Hermann`). Ein Buch mit vier Herausgebern gehört also in vier
   Kategorien — eine Sammelkategorie „Brohm / Stockert (Hrsg.)" gibt es nicht;
   was so im Breadcrumb steht, ist das Autorenfeld des Produkts.

   Gesucht wird **server-seitig**: `/api/category` gibt mit `limit=500` genau
   500 Einträge zurück, also abgeschnitten. Alles zu laden und örtlich zu
   filtern verfehlt daher zuverlässig die gesuchte Kategorie.

   Wer **keine** eigene Kategorie hat, wird ausdrücklich genannt — dann muss
   ein Mensch ran. Angelegt wird nie eine.
4. **Gewicht** prüfen — steht es nicht in der ONIX (Normalfall), trägt das
   Werkzeug eine **Schätzung** ein und schreibt daneben, woraus sie stammt.
   Überschreiben geht jederzeit; die ONIX neu laden holt die Schätzung zurück.
5. **Dry-Run** ankreuzen, um den Payload nur anzuschauen (nichts wird gesendet).
6. **Als Entwurf anlegen** → Bilder hochladen + Produkt anlegen/aktualisieren.
   Danach lässt sich das Produkt direkt im Admin öffnen.

### Hinterher prüfen

```
python -m shopware_publisher.dump_produkt --vergleich beispiele/buch.xml
```

Hält den aus der ONIX gebauten Payload Feld für Feld gegen das Produkt im Shop
und zeigt nur, was abweicht — dazu Kategorien und Bilderzahl. Damit beantwortet
sich „ist es so angekommen, wie gedacht?“ aus Daten statt aus Erinnerung.

## Technische Notizen (Shopware 6.7)

- **Auth:** `POST /api/oauth/token`, `grant_type=client_credentials`. 6.7 hat
  `league/oauth2-server` angehoben und weist **nicht OAuth-konforme** Anfragen
  ab — der Token-Request geht daher als `application/x-www-form-urlencoded`
  (RFC 6749), nicht als JSON. Ein Rückfall auf JSON ist eingebaut.
  Ein `scope` wird bewusst nicht mitgeschickt (Shopware nimmt den Standard der
  Integration).
- **Schreiben:** `POST /api/_action/sync` im dokumentierten Format
  `{"write-product": {"entity": "product", "action": "upsert", "payload": [...]}}`.
- **Bilder:** Medium per Sync anlegen (feste ID) → Dateiname über
  `GET /api/_action/media/provide-name` kollisionsfrei machen →
  `POST /api/_action/media/{id}/upload?extension=…&fileName=…` mit den rohen
  Bytes. Die Bilder landen im **Standard-Medienordner für Produkte**, damit
  Thumbnails erzeugt werden.
- **`coverId`** zeigt auf die **`product_media`-Verknüpfung**, nicht auf
  `media.id` — eine klassische Stolperfalle.
- Keine zusätzliche Abhängigkeit: der Client nutzt nur die stdlib (`urllib`).

## Bauen (Windows, aus dem Repo-Wurzelordner)

```
pyinstaller shopware_publisher/ShopwarePublisher.spec
```

oder alle Tools zusammen: `.\tools\update_and_build.ps1`.

## Dateien

```
shopware_publisher/
├── core.py        ONIX lesen, Payload bauen, Shopware-Client
├── app.py         Tkinter-GUI
├── main.py        Einstiegspunkt für den .exe-Build
├── ShopwarePublisher.spec
└── beispiele/     ONIX-Beispiel-XML
```

Laufzeitdaten (`config.json`) legt das Tool neben der .exe ab.
