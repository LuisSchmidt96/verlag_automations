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

> **Least Privilege:** Die Integration braucht keine *Administrator*-Rolle. Eine
> eigene Rolle mit Schreibrechten auf **Produkte + Medien** und Leserechten auf
> Steuer/Währung/Kategorie reicht — dann kann ein geleaktes Secret keine Kunden-
> und Bestelldaten lesen.

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
4. **Gewicht** eintragen, falls gebraucht — VLB liefert es meist nicht mit.
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
