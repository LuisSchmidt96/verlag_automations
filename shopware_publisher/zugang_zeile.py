"""Eine Zeile für die Zugangsdatei bauen.

Die Werkzeuge lesen Zugangsdaten nur noch aus einer Datei auf dem Share
(`sw.SHARE_ZUGANG`, siehe `core.py`). Eingetragen wird dort nichts von Hand —
das Geheimnis muss verschlüsselt sein, und das erledigt dieses Skript.

Aufruf aus der Repo-Wurzel:

    python -m shopware_publisher.zugang_zeile

Es fragt Abschnitt, Kennung, Geheimnis und Master-Passwort ab und gibt die
fertige Zeile aus. Die kommt in die Zugangsdatei, eine Zeile je Abschnitt.

Bewusst ein eigenes Skript und kein Knopf im Werkzeug: die Datei wird EINMAL
angelegt, von der Person, die den Zugang einrichtet. Ein Knopf in der
Oberfläche würde nahelegen, dass das alle tun — und dann läge am Ende in
jedem Werkzeugordner eine andere Fassung.

Das Master-Passwort wird nirgends gespeichert und darf NICHT neben der
Zugangsdatei liegen: sonst liegen Schloss und Schlüssel im selben Fach.
"""

from __future__ import annotations

import getpass
import sys

from shopware_publisher import core as sw


def _frage(text: str) -> str:
    wert = input(text).strip()
    if not wert:
        sys.exit("Abgebrochen — leere Eingabe.")
    return wert


def _frage_geheim(text: str, wiederholen: bool = True) -> str:
    """Verdeckt abfragen und zur Sicherheit wiederholen lassen.

    Ein Tippfehler im Geheimnis fällt sonst erst auf, wenn der Shop-Zugriff
    scheitert — und dann sucht man den Fehler beim Passwort.
    """
    wert = getpass.getpass(text)
    if not wert:
        sys.exit("Abgebrochen — leere Eingabe.")
    if wiederholen and getpass.getpass("  noch einmal zur Bestätigung: ") != wert:
        sys.exit("Die beiden Eingaben waren nicht gleich.")
    return wert


def main() -> None:
    print(__doc__.split("Aufruf")[0].strip())
    print()
    print("Abschnitt:  'dev' oder 'prod' (Shopware) oder 'sftp'")
    abschnitt = _frage("Abschnitt: ").lower()

    if abschnitt == "sftp":
        kennung = _frage("SFTP-Benutzer (z. B. sftpuser): ")
        geheim = _frage_geheim("SFTP-Passwort: ")
    else:
        kennung = _frage("Zugriffsschlüssel-ID (SWIA…): ")
        geheim = _frage_geheim("Shopware-Secret: ")

    print()
    print("Das Master-Passwort öffnet ALLE Zeilen dieser Datei.")
    print("Für alle Abschnitte dasselbe nehmen — sonst muss man es je")
    print("Schritt neu eingeben.")
    passwort = _frage_geheim("Master-Passwort: ")

    zeile = sw.zugang_zeile(abschnitt, kennung, geheim, passwort)

    # Gegenprobe: einmal wieder aufmachen. Lieber hier scheitern als später
    # im Werkzeug, wo man die Ursache nicht mehr sieht.
    secret_enc, kdf_salt = sw.blob_teilen(zeile.split(sw.ZUGANG_TRENNER)[2])
    probe = sw.hole_secret({"secret_enc": secret_enc, "kdf_salt": kdf_salt},
                           passwort)
    if probe != geheim:
        sys.exit("Gegenprobe fehlgeschlagen — die Zeile NICHT verwenden.")

    print()
    print("Gegenprobe bestanden. Diese Zeile in die Zugangsdatei:")
    print(f"  {sw.zugang_pfad()}")
    print()
    print(zeile)


if __name__ == "__main__":
    main()
