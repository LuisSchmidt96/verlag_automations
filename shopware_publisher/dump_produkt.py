#!/usr/bin/env python3
"""Zieht ein vorhandenes Produkt komplett aus dem Shop — zum Abgleich.

Zeigt, welche Felder Shopware bei einem *richtig gepflegten* Produkt wirklich
gesetzt hat (Sichtbarkeiten, Sales-Channel, Einheiten, SEO, Kategorien …).
Daraus lässt sich ableiten, was der Publisher noch nicht befüllt.

Aufruf im Terminal (fragt nach dem Master-Passwort — es wird nicht gespeichert):

    python -m shopware_publisher.dump_produkt
    python -m shopware_publisher.dump_produkt 9783955055592   # gezielt suchen
    python -m shopware_publisher.dump_produkt --liste         # nur auflisten

Ergebnis: shopware_publisher/beispiel_produkt.json
Dieses Dev-Skript ist NICHT Teil der .exe.
"""

from __future__ import annotations

import getpass
import json
import re
import sys
from pathlib import Path

from shopware_publisher import core

# Alles mitladen, was ein Produkt ausmacht.
ASSOZIATIONEN = {
    "tax": {},
    "manufacturer": {},
    "unit": {},
    "deliveryTime": {},
    "cover": {"associations": {"media": {}}},
    "media": {"associations": {"media": {}}},
    "categories": {},
    "mainCategories": {"associations": {"category": {}, "salesChannel": {}}},
    "prices": {},
    "visibilities": {"associations": {"salesChannel": {}}},
    "properties": {"associations": {"group": {}}},
    "options": {},
    "tags": {},
    "seoUrls": {},
    "translations": {},
    "cmsPage": {},
}


def _breadcrumb_diagnose(p: dict) -> None:
    """Die für den Breadcrumb-Fehler entscheidenden Felder auf einen Blick."""
    print("\n=== Breadcrumb-Diagnose ===")
    print("  active        :", p.get("active"))
    kats = p.get("categories") or []
    print(f"  categories    : {len(kats)}")
    for k in kats:
        print(f"      - {k.get('name')!r:40} id={k.get('id')} "
              f"type={k.get('type')} active={k.get('active')}")
    print("  categoryTree  :", p.get("categoryTree"))
    mc = p.get("mainCategories") or []
    print(f"  mainCategories: {len(mc)}")
    for m in mc:
        cat = m.get("category") or {}
        print(f"      - salesChannelId={m.get('salesChannelId')} "
              f"category={cat.get('name')!r}")
    vis = p.get("visibilities") or []
    print(f"  visibilities  : {len(vis)}")
    for v in vis:
        sc = v.get("salesChannel") or {}
        print(f"      - {sc.get('name')!r} navigationCategoryId="
              f"{sc.get('navigationCategoryId')} visibility={v.get('visibility')}")
    seo = p.get("seoUrls") or []
    print(f"  seoUrls       : {len(seo)}")
    for s in seo[:4]:
        print(f"      - {s.get('seoPathInfo')!r} canonical={s.get('isCanonical')}")


ZIEL_STANDARD = Path(__file__).parent / "beispiel_produkt.json"


def main() -> int:
    cfg = core.lade_config()
    umg = core.umgebung(cfg)
    name = core.aktive_umgebung(cfg)
    url = core.normalisiere_url(umg.get("shop_url", ""))
    print(f"Umgebung: {name}   {url}")

    if not url or not umg.get("access_key_id"):
        print("Shop-URL oder Access-Key fehlt — bitte im Tool eintragen.")
        return 1
    if not core.hat_secret(umg):
        print("Kein Secret hinterlegt — bitte im Tool „Secret setzen…“.")
        return 1

    try:
        secret = core.hole_secret(umg, getpass.getpass("Master-Passwort: "))
    except core.PasswortFehler as e:
        print(e)
        return 1

    c = core.ShopClient(url, umg["access_key_id"], secret,
                        tls_pruefen=bool(umg.get("tls_pruefen", True)))
    try:
        version = c.verbinde().get("version", "?")
        print(f"Verbunden — Shopware {version}\n")

        argumente = [a for a in sys.argv[1:] if not a.startswith("--")]
        nur_liste = "--liste" in sys.argv

        if nur_liste:
            for p in c.suche("product", {"limit": 25}):
                print(f"  {p.get('productNumber'):<20} {p.get('name')}")
            return 0

        koerper: dict = {"limit": 1, "associations": ASSOZIATIONEN}
        if argumente:
            koerper["filter"] = [{"type": "contains",
                                  "field": "productNumber",
                                  "value": argumente[0]}]
        treffer = c.suche("product", koerper)
        if not treffer:
            print("Kein Produkt gefunden. Mit --liste die vorhandenen ansehen.")
            return 1

        p = treffer[0]
        # Eigene Datei je Artikelnummer, damit die Referenz (beispiel_produkt.json)
        # erhalten bleibt und man Produkte nebeneinander vergleichen kann.
        if argumente:
            sicher = re.sub(r"[^0-9A-Za-z_-]", "_", str(p.get("productNumber")))
            ziel = Path(__file__).parent / f"dump_{sicher}.json"
        else:
            ziel = ZIEL_STANDARD
        ziel.write_text(json.dumps(p, indent=2, ensure_ascii=False),
                        encoding="utf-8")
        print(f"Produkt: {p.get('productNumber')}  —  {p.get('name')}")
        _breadcrumb_diagnose(p)
        print(f"\ngeschrieben nach: {ziel}")
        return 0
    except core.ShopFehler as e:
        print(f"\nFEHLER: {e}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
