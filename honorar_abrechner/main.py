#!/usr/bin/env python3
"""PyInstaller-Einstiegspunkt für den Honorar-Abrechner.

Bauen (aus dem Repo-Wurzelordner):
    pyinstaller honorar_abrechner/HonorarAbrechner.spec
"""

from honorar_abrechner.app import main

if __name__ == "__main__":
    main()
