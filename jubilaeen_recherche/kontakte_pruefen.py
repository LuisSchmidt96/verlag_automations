"""Gegenprobe für gesammelte Kontaktdaten: steht die E-Mail wirklich auf der angegebenen Fundstelle?

Aufruf: python kontakte_pruefen.py   → schreibt kontakte/pruefung.json (qid → Status)

Die Agenten sollen nur Adressen übernehmen, die auf der Seite stehen. Diese Probe ruft
jede Fundstelle noch einmal selbst ab, mit normaler Zertifikatsprüfung. Status:
  bestätigt                     – Adresse steht wörtlich im Seitentext
  verschleiert, Teile gefunden  – Name und Domain stehen da, aber nicht als Adresse
                                  („name [at] domain“) – vermutlich richtig, ansehen
  nicht auf der Seite gefunden  – gegenlesen (oft PDF oder per JavaScript nachgeladen)
  Seite nicht abrufbar (…)      – gegenlesen
"""

import glob
import html
import json
import time
from pathlib import Path

import requests

ORDNER = Path(__file__).resolve().parent / "kontakte"


def main():
    ergebnis = [json.loads(z) for f in sorted(glob.glob(str(ORDNER / "ergebnis_*.jsonl")))
                for z in open(f, encoding="utf-8") if z.strip()]
    pfad = ORDNER / "pruefung.json"
    status = json.loads(pfad.read_text(encoding="utf-8")) if pfad.exists() else {}
    sitzung = requests.Session()
    sitzung.headers["User-Agent"] = "Mozilla/5.0 (VerlagJubilaeumBot/0.1; kontakt@cyber-schmidt.de)"
    for k in ergebnis:
        mail = (k.get("email") or "").lower()
        if not mail or "@" not in mail or k["qid"] in status:
            continue
        try:
            text = html.unescape(sitzung.get(k["quelle"], timeout=20).text).lower()
            lokal, domain = mail.split("@")
            if mail in text:
                status[k["qid"]] = "bestätigt"
            elif lokal in text and domain.split(".")[0] in text:
                status[k["qid"]] = "verschleiert, Teile gefunden"
            else:
                status[k["qid"]] = "nicht auf der Seite gefunden"
        except Exception as fehler:
            status[k["qid"]] = f"Seite nicht abrufbar ({type(fehler).__name__})"
        print(f"{k['name'][:35]:35} {mail:40} {status[k['qid']]}")
        time.sleep(0.5)
    pfad.write_text(json.dumps(status, ensure_ascii=False, indent=1), encoding="utf-8")


if __name__ == "__main__":
    main()
