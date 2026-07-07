# -*- mode: python ; coding: utf-8 -*-
import os

# Repo-Wurzel (ein Ordner über diesem Spec) muss auf den Suchpfad, damit das
# Paket `pi_bi_generator` importierbar ist.
REPO_ROOT = os.path.dirname(SPECPATH)

# Die docx-/html-Vorlagen MÜSSEN mitgebündelt werden — ohne sie kann zur
# Laufzeit nichts erzeugt werden. Ziel-Ordner passt zu core._vorlagen_dir().
datas = [
    (os.path.join(SPECPATH, 'vorlagen', 'pi_vorlage.docx'), 'pi_bi_generator/vorlagen'),
    (os.path.join(SPECPATH, 'vorlagen', 'bi_vorlage.docx'), 'pi_bi_generator/vorlagen'),
    (os.path.join(SPECPATH, 'vorlagen', 'pi_vorlage.html'), 'pi_bi_generator/vorlagen'),
    (os.path.join(SPECPATH, 'vorlagen', 'bi_vorlage.html'), 'pi_bi_generator/vorlagen'),
]
binaries = []
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
    name='PiBiGenerator',
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
    name='PiBiGenerator',
)
