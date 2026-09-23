# -*- mode: python ; coding: utf-8 -*-
import os

# Repo-Wurzel (ein Ordner über diesem Spec) muss auf den Suchpfad, damit das
# Paket `honorar_abrechner` importierbar ist.
REPO_ROOT = os.path.dirname(SPECPATH)

# Die Briefvorlage ist eine Nur-Lese-Beigabe und wird mitgebündelt. Das
# Zielverzeichnis muss zu core._vorlagen_dir() passen, sonst findet die .exe
# sie zur Laufzeit nicht.
datas = [
    (os.path.join(SPECPATH, 'vorlagen', 'brief_vorlage.docx'),
     'honorar_abrechner/vorlagen'),
]


a = Analysis(
    [os.path.join(SPECPATH, 'main.py')],
    pathex=[REPO_ROOT],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # windows_pruefung.py ist reine Entwicklungshilfe und gehört nicht in
    # die .exe — sie wird auf dem Windows-Rechner von Hand gestartet.
    excludes=['honorar_abrechner.windows_pruefung'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='HonorarAbrechner',
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
    name='HonorarAbrechner',
)
