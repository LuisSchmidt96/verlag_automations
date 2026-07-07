#!/usr/bin/env python3
"""PyInstaller-Einstiegspunkt für den BOOXpress-Etiketten-Generator.

Bauen (aus dem Repo-Wurzelordner):
    pyinstaller booxpress_etiketten/BooxpressEtiketten.spec
"""

from booxpress_etiketten.app import main

if __name__ == "__main__":
    main()
