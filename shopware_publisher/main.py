#!/usr/bin/env python3
"""PyInstaller-Einstiegspunkt für den Shopware-Publisher.

Bauen (aus dem Repo-Wurzelordner):
    pyinstaller shopware_publisher/ShopwarePublisher.spec
"""

from shopware_publisher.app import main

if __name__ == "__main__":
    main()
