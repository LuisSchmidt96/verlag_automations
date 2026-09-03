#!/usr/bin/env python3
"""Probelauf gegen den Dev-Shop — der ganze Weg auf einmal.

Spielt durch, was sonst per Hand geklickt wird: verbinden, alle Beispiel-ONIX
lesen, Bilder suchen, Kategorie vorschlagen, Payload bauen — und auf Wunsch die
Bücher als Entwurf anlegen und hinterher gegenprüfen.

Aufruf im Terminal (fragt nach dem Master-Passwort, es wird NICHT gespeichert):

    python -m shopware_publisher.probelauf              # nur lesen + Dry-Run
    python -m shopware_publisher.probelauf --anlegen    # legt die Bücher an
    python -m shopware_publisher.probelauf --anlegen --ueberschreiben
    python -m shopware_publisher.probelauf --anlegen --ueberschreiben --nur 9783955055592
    python -m shopware_publisher.probelauf --gewichte   # nur die Gewichtsprobe

``--gewichte`` hält die Schätzung gegen die echten Gewichte im Shop
(Kreuzvalidierung) und schreibt jede Zeile nach ``gewichte_pruefung.csv``.

``--nur`` grenzt auf ein Buch ein — damit ein ``--ueberschreiben`` nicht
versehentlich alle Bestandsbücher mitnimmt.

Läuft **nur gegen eine nicht-produktive Umgebung**. Gegen den Produktivshop
verweigert das Skript den Dienst — dafür ist die Oberfläche da, mit ihren
Rückfragen.

Dieses Dev-Skript ist NICHT Teil der .exe.
"""

from __future__ import annotations

import csv
import getpass
import sys
import traceback
from pathlib import Path

from shopware_publisher import core

BEISPIELE = Path(__file__).parent / "beispiele"


def _linie(text: str = "") -> None:
    print("\n" + "=" * 78)
    if text:
        print(text)
        print("=" * 78)


def gewichte_pruefen(client, cfg: dict) -> int:
    """Die Gewichtsschätzung gegen den Bestand halten.

    Der Shop kennt zu jedem gepflegten Produkt das ECHTE Gewicht — damit lässt
    sich messen statt vermuten, wie gut die Schätzung ist. Gelernt und geprüft
    wird getrennt (Kreuzvalidierung), sonst benotete sich das Modell selbst.
    """
    _linie("Gewichtsmodell — Gegenprobe am Bestand")
    roh = client.bestand_gewichte()
    proben = [core.gewichtsprobe(p, cfg) for p in roh]
    proben = [p for p in proben if p]
    print(f"  Produkte mit Gewicht + Maßen : {len(roh)}")
    print(f"  davon mit Seitenzahl im Text : {len(proben)}")
    if len(proben) < core.MIN_PROBEN:
        print("  ⚠ zu wenige für eine Schätzung — hier ist nichts zu holen.")
        return 1

    modell, pruefung = core.lerne_und_pruefe(proben)
    cfg["gewichtsmodell"] = modell
    core.speichere_config(cfg)
    print("\n  Gelerntes Modell:")
    for name, g in sorted(modell["gruppen"].items()):
        print(f"    {name:<18} {g['papier_g_qm']:>6} g/m² Papier  "
              f"{g['einband_g_qm']:>7} g/m² Einband   ({g['proben']} Proben, "
              f"{g.get('verworfen', 0)} verworfen)")

    def zeile(name: str, k: dict) -> None:
        if not k:
            return
        print(f"    {name:<18} n={k['n']:<5} Median {k['median_prozent']:>5} %  "
              f"90 % unter {k['p90_prozent']:>5} %   "
              f"±10 %: {k['im_10er_band']:>3} %   ±20 %: {k['im_20er_band']:>3} %"
              f"   Schlagseite {k['schlagseite_prozent']:+} %")

    print("\n  Abweichung vom echten Gewicht (Modell kannte das Buch nicht):")
    zeile("alle", pruefung["gesamt"])
    for name, k in pruefung["je_einband"].items():
        zeile(name, k)

    verdacht = [e for e in pruefung["einzeln"]
                if abs(e["abweichung"]) > core._VERDACHT_AB]
    print(f"\n  Verdächtige Bestandsgewichte: {len(verdacht)} von "
          f"{len(proben)} liegen über {core._VERDACHT_AB:.0%} daneben."
          "\n  Aus dem Lernen sind sie draußen; im Shop gehören sie geprüft —"
          "\n  ein Platzhalter wie „0,200 kg“ kostet echtes Porto:")
    print(f"    {'Artikelnummer':<20}{'Seiten':>7}{'Format':>14}"
          f"{'echt':>9}{'geschätzt':>11}{'ab':>8}   Einband")
    for e in verdacht[:15]:
        format_ = "{} x {}".format(e["breite_cm"], e["hoehe_cm"])
        print(f"    {e['nummer']:<20}{e['seiten']:>7}{format_:>14}"
              f"{e['gewicht_kg']:>9.3f}{e['geschaetzt_kg']:>11.3f}"
              f"{100 * e['abweichung']:>7.0f}%   {e['einband'] or '—'}")

    ziel = core.APP_DIR / "gewichte_pruefung.csv"
    spalten = ["nummer", "seiten", "breite_cm", "hoehe_cm", "einband",
               "einband_roh", "gewicht_kg", "geschaetzt_kg", "abweichung",
               "gruppe", "g_qm"]
    with open(ziel, "w", encoding="utf-8", newline="") as fh:
        schreiber = csv.DictWriter(fh, fieldnames=spalten, extrasaction="ignore")
        schreiber.writeheader()
        for e in sorted(pruefung["einzeln"], key=lambda x: x["nummer"]):
            schreiber.writerow({**e, "abweichung": round(e["abweichung"], 4),
                                "g_qm": round(e["g_qm"], 1)})
    print(f"\n  Alle {len(pruefung['einzeln'])} Zeilen: {ziel}")
    return 0


def main() -> int:
    anlegen = "--anlegen" in sys.argv
    ueberschreiben = "--ueberschreiben" in sys.argv
    # --nur <teil>: nur Dateien, deren Name oder ISBN das enthält. Sonst wäre
    # --ueberschreiben ein Rundumschlag über alle Bestandsbücher.
    nur = ""
    if "--nur" in sys.argv:
        i = sys.argv.index("--nur")
        if i + 1 < len(sys.argv):
            nur = sys.argv[i + 1].strip()

    cfg = core.lade_config()
    name = core.aktive_umgebung(cfg)
    eff = core.effektiv(cfg)
    url = core.normalisiere_url(eff.get("shop_url", ""))

    _linie(f"Umgebung: {name}   {url or '— keine Shop-URL —'}")

    # Ein Probelauf gehoert nicht in den Produktivshop. Lieber hier abbrechen
    # als sich auf die Aufmerksamkeit des Bedieners zu verlassen.
    if core.ist_produktiv(name):
        print("Das ist die PRODUKTIVUMGEBUNG. Der Probelauf läuft nur gegen dev.")
        print("Umschalten: in der Oberfläche, oder aktive_umgebung in config.json.")
        return 1
    if not url or not eff.get("access_key_id"):
        print("Shop-URL oder Zugriffsschlüssel fehlt — bitte im Werkzeug eintragen.")
        return 1

    umg = core.umgebung(cfg)
    if not core.hat_secret(umg):
        print("Kein geheimer Schlüssel hinterlegt — im Werkzeug unter "
              "„Geheimen Schlüssel setzen…“.")
        return 1
    try:
        secret = core.hole_secret(umg, getpass.getpass("Master-Passwort: "))
    except core.PasswortFehler as e:
        print(e)
        return 1

    # -- 1) Verbinden ---------------------------------------------------
    client = core.ShopClient(url, eff["access_key_id"], secret,
                             tls_pruefen=bool(eff.get("tls_pruefen", True)))
    version = client.verbinde().get("version", "?")
    print(f"Verbunden — Shopware {version}")

    # Nur die Gewichtsprobe: braucht weder Zuordnungen noch Beispiel-ONIX.
    if "--gewichte" in sys.argv:
        return gewichte_pruefen(client, cfg)

    # -- 2) Zuordnungen aus dem Shop ------------------------------------
    vorlage = client.vorlage_vom_bestand()
    for k in ("cms_page_id", "sales_channel_id", "visibility", "manufacturer_id"):
        if vorlage.get(k):
            umg[k] = vorlage[k]
    for schluessel, eintraege, passt in (
            ("tax_id", client.steuersaetze(),
             lambda e: abs(float(e["taxRate"]) - float(eff.get("tax_rate", 7.0))) < 0.01),
            ("currency_id", client.waehrungen(), lambda e: e["isoCode"] == "EUR")):
        if not umg.get(schluessel):
            for e in eintraege:
                if passt(e):
                    umg[schluessel] = e["id"]
                    if schluessel == "tax_id":
                        umg["tax_rate"] = float(e["taxRate"])
                    break
    eff = core.effektiv(cfg)
    print(f"Verkaufskanal      : {vorlage.get('sales_channel_name') or '⚠ keiner gefunden'}")
    print(f"Hersteller         : {vorlage.get('manufacturer_name') or '—'}")
    for pflicht in ("tax_id", "currency_id", "sales_channel_id"):
        if not eff.get(pflicht):
            print(f"⚠ {pflicht} fehlt — bitte einmal in der Oberfläche verbinden.")

    # -- 2b) Artikeldaten-Share: erreichbar? ----------------------------
    # Wenn alle Bilder vom Webserver kommen, will man wissen WARUM. Deshalb
    # hier nachsehen statt raten: gibt es den Share, was liegt darin, und
    # heissen die Ordner so, wie der Publisher sie sucht?
    _linie("Artikeldaten-Share")
    roh = eff.get("artikeldaten_dir") or ""
    basis = core.artikeldaten_dir(eff)
    print(f"  konfiguriert : {roh or '— nichts eingetragen —'}")
    if not roh:
        print("  -> Ohne Pfad kann nur der Webserver liefern.")
    elif basis is None:
        print("  ⚠ NICHT ERREICHBAR (Pfad existiert nicht oder kein Ordner).")
        print("     Prüfen: im Explorer öffnen, Netzlaufwerk verbunden?")
    else:
        try:
            ordner = sorted(x.name for x in basis.iterdir() if x.is_dir())
        except OSError as e:
            ordner = []
            print(f"  ⚠ Lesefehler: {e}")
        print(f"  erreichbar   : ja, {len(ordner)} Ordner")
        for name in ordner[:5]:
            print(f"      z. B. {name}")
        if len(ordner) > 5:
            print(f"      … und {len(ordner) - 5} weitere")

    # -- 2c) Gewichtsmodell aus dem Bestand -----------------------------
    # Das Gewicht steht in der ONIX praktisch nie; geschätzt wird es an den
    # Produkten, die im Shop schon gepflegt sind. Hier steht, WORAUS geschätzt
    # wird — geht eine Schätzung daneben, sieht man es an diesen Zahlen.
    _linie("Gewichtsmodell (aus dem Bestand gelernt)")
    try:
        modell = core.lerne_gewichte_vom_shop(client, cfg)
        core.speichere_config(cfg)
        print(f"  Proben       {modell['proben']} Produkte mit Gewicht, Maßen "
              f"und Seitenzahl")
        if not modell["gruppen"]:
            print("  ⚠ zu wenige brauchbare Produkte — es wird nicht geschätzt.")
        for name, g in sorted(modell["gruppen"].items()):
            print(f"  {name:<16} {g['papier_g_qm']:>6} g/m² Papier  "
                  f"{g['einband_g_qm']:>7} g/m² Einband  "
                  f"({g['proben']} Proben, ±{g['abweichung_prozent']} %)")
    except core.ShopFehler as e:
        print(f"  ⚠ nicht gelernt: {e}")

    # -- 3) Je Buch: lesen, Bild suchen, Kategorie vorschlagen ----------
    dateien = sorted(BEISPIELE.glob("*.xml"))
    if not dateien:
        print(f"Keine ONIX-Dateien in {BEISPIELE}")
        return 1

    for pfad in dateien:
        try:
            f = core.lade_buchfelder(pfad, eff)
        except Exception as e:
            _linie(f"{pfad.name}")
            print(f"  FEHLER beim Lesen: {e}")
            continue
        if nur and nur not in pfad.name and nur not in f["isbn13"] \
                and nur not in f["isbn13_formatiert"]:
            continue
        _linie(f"{pfad.name}")
        sc = f["shortcode"]

        bilder = core.finde_bilder(sc, eff, xml_pfad=pfad)
        if not bilder.get("cover"):
            p = core.hole_cover_web(sc, eff)
            if p:
                bilder["cover"], bilder["quelle"] = p, "Webserver (newsletter_)"

        # Kategorien: im Shop SUCHEN, je beteiligter Person eine. Den Baum zu
        # laden hilft nicht — er ist größer als eine Abfrage hergibt.
        kats, fehlend = core.kategorie_vorschlaege(f, client.kategorien_suchen, eff)
        kat_ids = [k["id"] for k in kats]

        print(f"  Name        {core.produkt_name(f)}")
        print(f"  Nummer      {f['isbn13_formatiert']}   {f['preis_brutto']:.2f} {f['waehrung']}")
        print(f"  Titelbild   {Path(bilder['cover']).name if bilder.get('cover') else '⚠ KEINES'}"
              f"   ({bilder.get('quelle') or '—'})")
        art_ordner = core.finde_artikel_ordner(sc, eff)
        if art_ordner:
            dateien = sorted(x.name for x in art_ordner.iterdir() if x.is_file())
            print(f"  Share-Ordner {art_ordner.name}  ->  "
                  f"{', '.join(dateien[:6]) or 'leer'}")
        elif basis is not None:
            print(f"  Share-Ordner ⚠ keiner, der mit {sc!r} anfängt")
        print(f"  Kategorien  {' · '.join(k.get('name','') for k in kats) or '— keine —'}")
        if fehlend:
            print(f"              ⚠ ohne eigene Kategorie: {', '.join(fehlend)}")
        if not f.get("gewicht_kg"):
            kg, grund = core.schaetze_gewicht(f, cfg.get("gewichtsmodell"))
            if kg:
                f["gewicht_kg"] = kg
            print(f"  Gewicht     {f'{kg} kg' if kg else '⚠ keines'}   ({grund})")
        else:
            print(f"  Gewicht     {f['gewicht_kg']} kg   (aus der ONIX)")

        vorhanden = client.produkt_id_zu_nummer(f["isbn13_formatiert"])
        print(f"  Im Shop     {'EXISTIERT BEREITS' if vorhanden else 'neu'}")

        if not anlegen:
            continue
        if vorhanden and not ueberschreiben:
            print("  -> übersprungen (existiert; mit --ueberschreiben erzwingen)")
            continue

        try:
            erg = core.veroeffentliche(f, cfg, bilder, secret=secret,
                                       ueberschreiben=ueberschreiben,
                                       kategorien=kat_ids,
                                       log=lambda m: print(f"     {m}"))
            print(f"  -> {'angelegt' if erg.get('neu') else 'aktualisiert'}: "
                  f"{erg['admin_url']}")
        except core.ShopFehler as e:
            print(f"  -> FEHLER: {e}")
            continue

    # -- 4) Gegenprobe --------------------------------------------------
    if anlegen:
        _linie("Gegenprobe — was ist tatsächlich angekommen?")
        for pfad in dateien:
            try:
                f = core.lade_buchfelder(pfad, eff)
            except Exception:
                continue
            if nur and nur not in pfad.name and nur not in f["isbn13"] \
                    and nur not in f["isbn13_formatiert"]:
                continue
            nummer = f["isbn13_formatiert"]
            treffer = client.suche("product", {
                "limit": 1,
                "filter": [{"type": "equals", "field": "productNumber",
                            "value": nummer}],
                "associations": {"categories": {}, "media": {}},
            })
            if not treffer:
                print(f"  {nummer}  — nicht im Shop")
                continue
            p = treffer[0]
            kats = ", ".join(k.get("name", "") for k in (p.get("categories") or []))
            print(f"  {nummer}  {(p.get('name') or '')[:34]:36} "
                  f"Bilder={len(p.get('media') or [])} "
                  f"Cover={'ja' if p.get('coverId') else 'NEIN'}  "
                  f"Kat={kats or '— keine —'}")
        print("\nEinzelheiten Feld für Feld:")
        # Pfad relativ zum Arbeitsverzeichnis ausgeben — aufgerufen wird aus
        # dem Repo-Wurzelordner, nicht aus dem Werkzeugordner.
        for pfad in dateien:
            try:
                zeig = pfad.relative_to(Path.cwd())
            except ValueError:
                zeig = pfad
            print(f"  python -m shopware_publisher.dump_produkt --vergleich {zeig}")
    else:
        print("\nNichts gesendet (Dry-Run). Zum Anlegen: --anlegen")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except core.ShopFehler as e:
        print(f"\nFEHLER: {e}")
        raise SystemExit(1)
    except KeyboardInterrupt:
        raise SystemExit(130)
    except Exception:
        traceback.print_exc()
        raise SystemExit(1)
