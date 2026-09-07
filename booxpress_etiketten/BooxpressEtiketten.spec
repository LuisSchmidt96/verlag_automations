# -*- mode: python ; coding: utf-8 -*-
import os
from PyInstaller.utils.hooks import collect_all

# Repo-Wurzel (ein Ordner über diesem Spec) muss auf den Suchpfad, damit die
# Pakete `booxpress_etiketten` und `shared` importierbar sind.
REPO_ROOT = os.path.dirname(SPECPATH)

# Beim Bauen meldet PyInstaller: WARNING: Hidden import "jinja2" not found!
# Die Warnung ist erwartet und harmlos. Sie kommt aus dem mitgelieferten Hook
# hook-pandas.io.formats.style.py: pandas' Styler (DataFrame.style, to_html)
# braucht jinja2, das ist aber eine OPTIONALE pandas-Abhaengigkeit. Hier wird
# pandas nur fuer read_excel und isna benutzt (booxpress_etiketten/core.py),
# nie der Styler — darum steht jinja2 auch nicht in requirements.txt.
# Nicht ueber excludes stummschalten: das nimmt pandas ein Modul weg, das es
# selbst importiert, und tauscht eine Warnung gegen einen Absturz.
datas = []
binaries = []
hiddenimports = []
tmp_ret = collect_all('barcode')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


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
    name='BooxpressEtiketten',
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
    name='BooxpressEtiketten',
)
