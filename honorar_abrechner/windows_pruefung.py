#!/usr/bin/env python3
"""
Prüfung der Windows-Eigenheiten — Entwicklungshilfe, nicht Teil der .exe
=========================================================================
Alles, was sich unter Linux nicht prüfen lässt, steht hier als ausführbarer
Test. Auf dem Windows-Rechner einmal starten:

    python honorar_abrechner\\windows_pruefung.py

Geprüft wird:

1. pywin32 vorhanden und Word ansprechbar
2. COM aus einem Arbeitsthread heraus — genau so, wie die Oberfläche es tut
3. Ein Brief wird wirklich nach PDF gewandelt, und das PDF hat zwei Seiten
4. Die Sperrerkennung greift, wenn der Bestand in Excel offen ist
5. Der Bestand überlebt Öffnen und Speichern in Excel

Jeder Punkt meldet OK oder FEHLER. Am Ende steht, was zu tun ist.
"""

from __future__ import annotations

import queue
import sys
import threading
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from honorar_abrechner import core          # noqa: E402


def _kopf(nummer: int, was: str) -> None:
    print(f"\n{nummer}. {was}")
    print("   " + "-" * (len(was) + 2))


def pruefe_pywin32() -> bool:
    _kopf(1, "pywin32 und Word")
    if sys.platform != "win32":
        print("   ÜBERSPRUNGEN — dieses Skript gehört auf den Windows-Rechner.")
        return False
    try:
        import pythoncom                    # noqa: F401
        import win32com.client
    except ImportError as e:
        print(f"   FEHLER: pywin32 fehlt ({e}).")
        print("   → pip install pywin32==306")
        return False
    try:
        pythoncom.CoInitialize()
        word = win32com.client.Dispatch("Word.Application")
        version = word.Version
        word.Quit()
        pythoncom.CoUninitialize()
    except Exception as e:
        print(f"   FEHLER: Word antwortet nicht ({e}).")
        return False
    print(f"   OK — Word {version} antwortet.")
    return True


def pruefe_com_im_thread() -> bool:
    """Der eigentliche Stolperstein.

    Die Oberfläche ruft die Umwandlung aus einem Arbeitsthread auf. COM muss
    in jedem Thread einzeln angemeldet werden; fehlt das, scheitert schon das
    Dispatch. Dieser Test stellt genau diese Lage her.
    """
    _kopf(2, "COM aus einem Arbeitsthread")
    ergebnis: queue.Queue = queue.Queue()

    def arbeite():
        try:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            try:
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = False
                word.Quit()
                ergebnis.put(("ok", None))
            finally:
                pythoncom.CoUninitialize()
        except Exception as e:
            ergebnis.put(("fehler", e))

    t = threading.Thread(target=arbeite, daemon=True)
    t.start()
    t.join(timeout=90)
    if t.is_alive():
        print("   FEHLER: Word hat nicht geantwortet (hängt vermutlich an "
              "einer Rückfrage).")
        return False
    art, fehler = ergebnis.get()
    if art != "ok":
        print(f"   FEHLER: {fehler}")
        return False
    print("   OK — COM läuft im Arbeitsthread.")
    return True


def pruefe_pdf(ordner: Path) -> bool:
    _kopf(3, "Brief nach PDF wandeln")
    vorlage = core._vorlagen_dir() / "brief_vorlage.docx"
    if not vorlage.exists():
        print(f"   FEHLER: Briefvorlage fehlt unter {vorlage}")
        return False
    e = core.Empfaenger(kennung="E0001", anrede="Herr", vorname="Probe",
                        name="Leser", strasse="Bahnhofstr. 2", plz="76698",
                        ort="Ubstadt-Weiher", iban="DE00 0000 0000 0000")
    b = core.Buch(kennung="B0001", titel="Probebuch", isbn="00-000",
                  verguetungsart="Honorar",
                  kondition=core.Kondition(betrag_je_ex=1.5))
    b.jahr(2025).verkauft = 20
    e.buecher = [b]
    ab = core.rechne_empfaenger(e, 2025, dict(core.DEFAULT_CONFIG))
    ordner.mkdir(parents=True, exist_ok=True)
    docx = core.generiere_brief(vorlage, ab, dict(core.DEFAULT_CONFIG),
                                ordner / "Probe.docx")
    pdfs = core.wandle_nach_pdf([docx], log=lambda s: print(f"   {s}"))
    if not pdfs or not pdfs[0].exists():
        print("   FEHLER: Es entstand kein PDF.")
        return False
    try:
        import fitz
        seiten = fitz.open(pdfs[0]).page_count
    except ImportError:
        seiten = None
    if seiten is not None and seiten != 2:
        print(f"   FEHLER: Das PDF hat {seiten} Seiten, erwartet werden 2 "
              f"(Anschreiben hoch, Tabelle quer).")
        return False
    print(f"   OK — {pdfs[0].name}" + (f", {seiten} Seiten." if seiten else "."))
    print("   → Bitte trotzdem einmal ansehen: Querformat auf Seite 2, "
          "Umlaute im Dateinamen, Briefkopf vollständig.")
    return True


def pruefe_sperre(ordner: Path) -> bool:
    _kopf(4, "Sperrerkennung bei geöffnetem Excel")
    ordner.mkdir(parents=True, exist_ok=True)
    pfad = ordner / "Sperrprobe.xlsx"
    bestand = core.Bestand()
    bestand.empfaenger.append(core.Empfaenger(kennung="E0001", name="Probe"))
    core.speichere_bestand(bestand, pfad, mit_sicherung=False, jahr=2025)
    # Excel legt neben einer geöffneten Datei diese Sperrdatei an.
    sperre = pfad.parent / f"~${pfad.name}"
    sperre.write_text("x", encoding="utf-8")
    try:
        meldung = core.pruefe_sperre(pfad)
        if not meldung:
            print("   FEHLER: Die Sperre wurde nicht erkannt.")
            return False
        print("   OK — erkannt. Die Meldung lautet:")
        for zeile in meldung.splitlines():
            print(f"      {zeile}")
    finally:
        sperre.unlink(missing_ok=True)
    print("   → Zusätzlich von Hand: die Mappe wirklich in Excel öffnen und "
          "im Werkzeug speichern. Es muss die Meldung kommen, kein Absturz.")
    return True


def pruefe_rundlauf(ordner: Path) -> bool:
    _kopf(5, "Bestand nach Excel und zurück")
    pfad = ordner / "Rundlauf.xlsx"
    bestand = core.Bestand()
    e = core.Empfaenger(kennung="E0001", name="Probe", vorname="Anna",
                        plz="76698", ort="Ubstadt")
    b = core.Buch(kennung="B0001", titel="Probebuch", isbn="05-300",
                  kondition=core.Kondition(ladenpreis=19.9, mwst_im_preis=7.0,
                                           satz=10.0, schwelle_zehn=True))
    b.jahr(2024).verkauft = 0
    b.jahr(2025).verkauft = 50
    e.buecher = [b]
    bestand.empfaenger.append(e)
    core.speichere_bestand(bestand, pfad, mit_sicherung=False, jahr=2025)
    zurueck = core.lade_bestand(pfad)
    if zurueck.warnungen:
        print(f"   FEHLER: Warnungen beim Laden: {zurueck.warnungen}")
        return False
    b2 = zurueck.empfaenger[0].buecher[0]
    proben = [("PLZ bleibt Text", zurueck.empfaenger[0].plz, "76698"),
              ("2025 erfasst", b2.jahre[2025].verkauft, 50),
              ("2024 als echte Null", b2.jahre[2024].verkauft, 0),
              ("Schwelle", b2.kondition.schwelle_zehn, True)]
    fehler = [n for n, ist, soll in proben if ist != soll]
    if fehler:
        print(f"   FEHLER bei: {', '.join(fehler)}")
        return False
    print("   OK — nichts verrutscht.")
    print(f"   → Jetzt von Hand: {pfad.name} in Excel öffnen, im Blatt "
          f"„Abrechnung 2025“ eine Zahl ändern, speichern, und im Werkzeug "
          f"„Gespeicherten Stand neu laden“. Die Zahl muss ankommen.")
    return True


def main() -> int:
    print(__doc__.split("====\n")[1].split("Geprüft wird")[0].strip())
    ordner = core.APP_DIR / "windows_pruefung"
    ergebnisse = []
    hat_word = pruefe_pywin32()
    ergebnisse.append(("pywin32 und Word", hat_word))
    if hat_word:
        ergebnisse.append(("COM im Arbeitsthread", pruefe_com_im_thread()))
        ergebnisse.append(("Brief nach PDF", pruefe_pdf(ordner)))
    ergebnisse.append(("Sperrerkennung", pruefe_sperre(ordner)))
    ergebnisse.append(("Bestand-Rundlauf", pruefe_rundlauf(ordner)))

    print("\n" + "=" * 58)
    for was, ok in ergebnisse:
        print(f"   {'OK     ' if ok else 'FEHLER '} {was}")
    offen = [w for w, ok in ergebnisse if not ok]
    if offen:
        print(f"\n   Offen: {', '.join(offen)}")
        return 1
    print("\n   Alles durch. Bleibt von Hand: die drei Punkte, die oben mit "
          "„→“ markiert sind.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
