#!/usr/bin/env python3
"""PyInstaller-Einstiegspunkt für den PI/BI-Generator.

Bauen (aus dem Repo-Wurzelordner):
    pyinstaller pi_bi_generator/PiBiGenerator.spec
"""

from pi_bi_generator.app import main

if __name__ == "__main__":
    main()
