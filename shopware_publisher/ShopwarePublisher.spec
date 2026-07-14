# -*- mode: python ; coding: utf-8 -*-
import os

# Repo-Wurzel (ein Ordner über diesem Spec) muss auf den Suchpfad, damit das
# Paket `shopware_publisher` importierbar ist.
REPO_ROOT = os.path.dirname(SPECPATH)

# Keine Nur-Lese-Assets nötig: Zugangsdaten/Zuordnungen stehen in der
# config.json neben der .exe, die Bilder kommen vom Artikeldaten-Share.
datas = []
binaries = []
# Der Shopware-Client nutzt nur die stdlib (urllib) — keine Extra-Importe.
hiddenimports = []


a = Analysis(
    [os.path.join(SPECPATH, 'main.py')],
    pathex=[REPO_ROOT],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='ShopwarePublisher',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='ShopwarePublisher',
)
