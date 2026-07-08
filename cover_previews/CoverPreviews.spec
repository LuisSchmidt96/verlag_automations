# -*- mode: python ; coding: utf-8 -*-
import os

# Repo-Wurzel (ein Ordner über diesem Spec) muss auf den Suchpfad, damit das
# Paket `cover_previews` importierbar ist.
REPO_ROOT = os.path.dirname(SPECPATH)

# Es gibt keine zwingend zu bündelnden Nur-Lese-Assets: das Mockup-PSD wählt
# der Nutzer zur Laufzeit (Pfad in config.json). Liegt in mockups/ eine
# Standard-Vorlage, kann sie hier ergänzt werden, z. B.:
#   datas = [(os.path.join(SPECPATH, 'mockups', 'standard.psd'), 'cover_previews/mockups')]
datas = []
binaries = []

# PyMuPDF (fitz) wird von pyinstaller-hooks-contrib abgedeckt. pywin32 (COM)
# wird nur im 3D-Schritt und nur unter Windows benötigt — die win32-Hooks
# ziehen die passenden Module; die expliziten Angaben sind Absicherung.
hiddenimports = ['win32com', 'win32com.client', 'pythoncom', 'pywintypes']


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
    name='CoverPreviews',
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
    name='CoverPreviews',
)
