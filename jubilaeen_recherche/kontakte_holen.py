"""Kontaktdaten ohne Agenten nachholen: Impressum/Kontaktseite selbst abrufen und E-Mail auslesen.

Aufruf: python kontakte_holen.py   → kontakte/ergebnis_skript.jsonl

Für alle Leads aus kontakte/eingabe_*.json, die bisher keine E-Mail haben (Agent
gedrosselt, Seite nicht gelesen …). Die Agenten-Suche lief beim zweiten großen
Durchgang in die Drosselung von WebFetch; dieses Skript ruft die Seiten direkt ab,
mit normaler Zertifikatsprüfung und Pause zwischen den Abrufen.

Es übernimmt nur Adressen, die auf der Seite stehen (mailto-Links oder Text, auch
„name [at] domain“). Aus mehreren Adressen wählt es nach Art des Leads die
passendste (Stadtarchiv/Hauptamt bei Gemeinden, Presse bei Firmen …). Gefundenes
überschreibt einen leeren Agenten-Eintrag, weil die Datei alphabetisch nach den
Agenten-Dateien gelesen wird (ergebnis_skript > ergebnis_j…).
"""

import glob
import html
import json
import re
import time
from pathlib import Path
from urllib.parse import urljoin, urlparse

import requests

ORDNER = Path(__file__).resolve().parent / "kontakte"
UA = "Mozilla/5.0 (VerlagJubilaeumBot/0.1; kontakt@cyber-schmidt.de)"
MAIL = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
VERSCHLEIERT = re.compile(r"([A-Za-z0-9._%+-]+)\s*[\[\(\{]\s*(?:at|ät|@)\s*[\]\)\}]\s*([A-Za-z0-9-]+(?:\s*[\[\(\{]?\s*(?:dot|punkt|\.)\s*[\]\)\}]?\s*[A-Za-z0-9-]+)+)", re.I)
TELEFON = re.compile(r"(?:Tel(?:efon)?\.?|Fon|Phone)\s*:?\s*([+0-9][0-9 /().-]{6,22}[0-9])", re.I)
PFADE = ["impressum", "kontakt", "impressum.html", "kontakt.html", "de/impressum", "de/kontakt", "imprint"]
LINKWORT = re.compile(r"impressum|kontakt|imprint|contact", re.I)
VORLIEBE = {
    "Gemeinde/Stadt": ["stadtarchiv", "archiv", "hauptamt", "buergermeister", "bürgermeister", "rathaus", "stadt",
                       "gemeinde", "poststelle", "info"],
    "Ortsteil": ["ortsverwaltung", "ortsvorsteher", "ortschaft", "archiv", "rathaus", "stadt", "gemeinde",
                 "poststelle", "info"],
    "Unternehmen/Einrichtung": ["presse", "press", "kommunikation", "communication", "marketing", "sekretariat",
                                "poststelle", "info", "kontakt", "mail"],
    "Verein": ["vorstand", "vorsitz", "geschaeftsstelle", "geschäftsstelle", "info", "kontakt"],
}
NICHT = ("example.", "sentry", "wixpress", "noreply", "no-reply", "datenschutz", "privacy", ".png", ".jpg")


def seite(sitzung, url):
    try:
        r = sitzung.get(url, timeout=15)
        if r.status_code == 200 and "html" in r.headers.get("content-type", "html"):
            return r.text
    except requests.RequestException:
        pass
    return None


def adressen(text):
    t = html.unescape(text)
    gefunden = set(m.group(0).rstrip(".") for m in MAIL.finditer(t))
    for m in re.finditer(r'mailto:([^"\'?>\s]+)', t):
        gefunden.add(m.group(1))
    for m in VERSCHLEIERT.finditer(re.sub(r"<[^>]+>", " ", t)):
        domain = re.sub(r"\s*[\[\(\{]?\s*(?:dot|punkt|\.)\s*[\]\)\}]?\s*", ".", m.group(2), flags=re.I)
        gefunden.add(f"{m.group(1)}@{domain}")
    sauber = set()
    for a in gefunden:
        a = re.sub(r"^(?:u003[ce]|%3[ce]|>|<)+", "", a.strip().lower())  # Reste aus JSON/HTML-Kodierung
        if a.count("@") == 1 and not any(n in a for n in NICHT):
            sauber.add(a)
    return sauber


ENDUNGEN = {"de", "com", "org", "net", "eu", "info", "biz", "at", "ch", "online", "bayern", "nrw", "berlin"}


def gehoert_dazu(mail, host):
    """Nur Adressen der eigenen Domain: auf Gemeindeseiten stehen auch Webagentur,
    Vorleseprogramm oder Bahn – die sind kein Kontakt für ein Ortsbuch."""
    if mail.count("@") != 1:
        return False
    domain = mail.split("@")[1]
    teile = domain.split(".")
    if len(teile) < 2 or teile[-1] not in ENDUNGEN:
        return False  # z. B. Bilddateinamen „…@216w.webp“ oder „mulfingen.de.bankverbindung“
    if not host:
        return False
    if teile[-2:] == host.split(".")[-2:]:
        return True
    kern_mail, kern_host = teile[-2], host.split(".")[-2]
    # stadt-kehl.de zu kehl.de, gemeinde-zweiflingen.de zu zweiflingen.de
    return len(kern_host) >= 4 and (kern_host in kern_mail or kern_mail in kern_host)


def waehle(mails, art, host):
    kandidaten = sorted(m for m in mails if gehoert_dazu(m, host))
    if not kandidaten:
        return None
    if art in ("Gemeinde/Stadt", "Ortsteil"):
        # Bei Verwaltungen nur allgemeine bzw. passende Stellen – nicht das
        # Ausländer- oder Bauamt, nur weil es zufällig auf der Seite steht.
        for wort in VORLIEBE[art]:
            for m in kandidaten:
                if wort in m.split("@")[0]:
                    return m
        return None
    for wort in VORLIEBE.get(art, ["info"]):
        for m in sorted(kandidaten):
            if wort in m.split("@")[0]:
                return m
    return kandidaten[0]


def hole(sitzung, lead):
    basis = lead.get("website")
    if not basis and lead["art"] in ("Gemeinde/Stadt", "Ortsteil"):
        basis = lead.get("website_gemeinde")  # Ortsteil → Gemeinde; Verein/Firma aber nie über die Stadt
    if not basis:
        return None
    if not basis.startswith("http"):
        basis = "https://" + basis
    host = urlparse(basis).netloc.lower().removeprefix("www.")
    start = seite(sitzung, basis)
    kandidaten = []
    if start:
        for m in re.finditer(r'<a[^>]+href="([^"#]+)"[^>]*>(.*?)</a>', start, re.I | re.S):
            if LINKWORT.search(m.group(1)) or LINKWORT.search(re.sub(r"<[^>]+>", "", m.group(2))):
                kandidaten.append(urljoin(basis, m.group(1)))
    kandidaten += [urljoin(basis.rstrip("/") + "/", p) for p in PFADE]
    gesehen = []
    for url in [basis] + list(dict.fromkeys(kandidaten))[:6]:
        text = start if url == basis else seite(sitzung, url)
        time.sleep(0.5)
        if not text:
            continue
        mails = adressen(text)
        if any(gehoert_dazu(m, host) for m in mails):
            gesehen.append((url, mails, text))
            if url != basis:
                break
    if not gesehen:
        return None
    url, mails, text = gesehen[-1]
    mail = waehle(mails, lead["art"], host)
    if not mail:
        return None
    tel = TELEFON.search(re.sub(r"<[^>]+>", " ", html.unescape(text)))
    privat = mail and mail.split("@")[1] in ("gmail.com", "t-online.de", "web.de", "gmx.de", "gmx.net", "freenet.de")
    return {"nr": lead["nr"], "qid": lead["qid"], "name": lead["name"], "email": mail,
            "telefon": tel.group(1).strip() if tel else "", "anschrift": "", "ansprechpartner": "",
            "quelle": url,
            "bemerkung": "automatisch aus der Seite gelesen" + ("; private Adresse – lieber über Gemeinde/Verein"
                                                                  if privat else "")
                         + (f"; weitere Adressen auf der Seite: {', '.join(sorted(mails - {mail})[:4])}"
                            if len(mails) > 1 else "")}


def main():
    leads = [l for f in sorted(glob.glob(str(ORDNER / "eingabe_*.json"))) for l in json.load(open(f, encoding="utf-8"))]
    vorhanden = {}
    for f in sorted(glob.glob(str(ORDNER / "ergebnis_*.jsonl"))):
        if f.endswith("ergebnis_skript.jsonl"):
            continue
        for z in open(f, encoding="utf-8"):
            if z.strip():
                k = json.loads(z)
                if k.get("email"):
                    vorhanden[k["qid"]] = k
    offen = {l["qid"]: l for l in leads if l["qid"] not in vorhanden}
    print(len(offen), "Leads ohne E-Mail")
    sitzung = requests.Session()
    sitzung.headers["User-Agent"] = UA
    ergebnisse = []
    for lead in offen.values():
        k = hole(sitzung, lead)
        print(f"{lead['name'][:35]:35} {(k or {}).get('email') or '–'}")
        if k and k["email"]:
            ergebnisse.append(k)
    (ORDNER / "ergebnis_skript.jsonl").write_text(
        "".join(json.dumps(k, ensure_ascii=False) + "\n" for k in ergebnisse), encoding="utf-8")
    print(len(ergebnisse), "Adressen gefunden")


if __name__ == "__main__":
    main()
