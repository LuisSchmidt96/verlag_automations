#!/usr/bin/env python3
"""PyInstaller-Einstiegspunkt für den Cover-Previews-Generator.

Bauen (aus dem Repo-Wurzelordner):
    pyinstaller cover_previews/CoverPreviews.spec
"""

from cover_previews.app import main

if __name__ == "__main__":
    main()
