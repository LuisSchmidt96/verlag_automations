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
| Artikelnummer  | ISBN-13 (auch als **EAN**)                            |
| Name           | Titel (+ „– Band n“, falls Reihe)                     |
| Beschreibung   | Werbetext (`d104`) als Absätze + Fakten-Block         |
| Preis          | DE-Preis brutto; **netto** = brutto / (1 + Steuersatz)|
| Bilder         | vom `cover_previews`-Tool (siehe unten)               |
| Steuer/Währung/Kategorie/Hersteller | einmalig im GUI zugeordnet       |

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

> Die `config.json` liegt neben der .exe und enthält das **Secret im Klartext**.
> Sie ist per `.gitignore` vom Repo ausgeschlossen — bitte nicht weitergeben.

## Dev-Store hinter Caddy (Basic-Auth + TLS)

**Basic-Auth blockt die Admin-API.** Basic und Bearer teilen sich denselben
`Authorization`-Header — es kann nur eines von beiden gesendet werden. Steht vor
dem Shop eine Basic-Auth, antwortet der Proxy mit 401, bevor Shopware das Token
überhaupt sieht. Das lässt sich **nicht** im Tool lösen, sondern nur im Proxy:
`/api/*` von der Basic-Auth ausnehmen.

```caddyfile
dev.example.de {
    # alles außer der API bleibt passwortgeschützt
    @geschuetzt not path /api/*
    basic_auth @geschuetzt {
        luis $2a$14$…   # bcrypt-Hash (caddy hash-password)
    }
    reverse_proxy localhost:8000
}
```

(Bei Caddy < 2.8 heißt die Direktive `basicauth` statt `basic_auth`.)

Danach `caddy reload`. Der Admin im Browser bleibt geschützt, die API ist für
das Tool erreichbar. Alternativ das Tool direkt gegen den Shop **hinter** dem
Proxy laufen lassen (z. B. `http://localhost:8000`).

**TLS:** Nutzt Caddy für den Dev-Store seine interne CA, kennt Python das
Zertifikat nicht (`CERTIFICATE_VERIFY_FAILED`). Dann in der `config.json`

```json
"tls_pruefen": false
```

setzen — **nur für den Dev-Store**. Im Produktivshop bleibt es `true`.

## Bilder

Die Bilder erzeugt das **`cover_previews`-Tool** und legt sie auf dem
Artikeldaten-Share im Ordner des Buchs ab. Der Publisher sucht sie dort anhand
des **Kurzcodes**:

```
<Artikeldaten>\<Kurzcode>_<Titel>\
    2D_72_<sc>.jpg     -> Cover (Hauptbild)
    3D_72_<sc>.jpg     -> zusätzliches Galeriebild
```

Fehlt ein Bild, wird das Produkt trotzdem angelegt (nur ohne Bild). Der Pfad
zum Share steht in `config.json → artikeldaten_dir`.

## Bedienung

1. **Verbinden** (einmalig, s. o.).
2. **ONIX-XML wählen** → die Vorschau zeigt, was im Shop landet
   (Name, Artikelnummer, Preis brutto/netto, gefundene Bilder, Beschreibung).
3. **Dry-Run** ankreuzen, um den Payload nur anzuschauen (nichts wird gesendet).
4. **Als Entwurf anlegen** → Bilder hochladen + Produkt anlegen/aktualisieren.
   Danach lässt sich das Produkt direkt im Admin öffnen.

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
