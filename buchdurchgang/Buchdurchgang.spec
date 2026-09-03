# -*- mode: python ; coding: utf-8 -*-
import os

# Repo-Wurzel (ein Ordner über diesem Spec) muss auf den Suchpfad, damit die
# Pakete `buchdurchgang`, `cover_previews`, `pi_bi_generator` und
# `shopware_publisher` importierbar sind.
REPO_ROOT = os.path.dirname(SPECPATH)

# Dieses Werkzeug BENUTZT die drei anderen (siehe buchdurchgang/core.py) und
# muss deshalb auch deren Beigaben mitnehmen:
#   * die docx-/html-Vorlagen des pi_bi_generator — ohne sie entsteht nichts
#   * vorlagen_map.json von cover_previews — die Zuordnung Format -> Mockup-PSD
#
# NICHT mitgenommen wird `_NEU_Vorlage/` (rund 480 MB Mockup-PSDs). Die liegen
# schon neben CoverPreviews.exe; der Durchgang zeigt in seiner config.json mit
# `cover_previews.vorlagen_dir` dorthin (z. B. "..\CoverPreviews\_NEU_Vorlage",
# beide Ordner liegen unter VR-Tools\ nebeneinander). Sonst läge dieselbe halbe
# Gigabyte zweimal auf dem Share.
datas = [
    (os.path.join(REPO_ROOT, 'pi_bi_generator', 'vorlagen', 'pi_vorlage.docx'),
     'pi_bi_generator/vorlagen'),
    (os.path.join(REPO_ROOT, 'pi_bi_generator', 'vorlagen', 'bi_vorlage.docx'),
     'pi_bi_generator/vorlagen'),
    (os.path.join(REPO_ROOT, 'pi_bi_generator', 'vorlagen', 'pi_vorlage.html'),
     'pi_bi_generator/vorlagen'),
    (os.path.join(REPO_ROOT, 'pi_bi_generator', 'vorlagen', 'bi_vorlage.html'),
     'pi_bi_generator/vorlagen'),
    (os.path.join(REPO_ROOT, 'cover_previews', 'vorlagen_map.json'),
     'cover_previews'),
]
binaries = []
# win32com steuert Photoshop und existiert nur unter Windows; PyInstaller
# findet den Import nicht von selbst, weil er in erzeuge_3d_photoshop() lokal
# steht (damit der Linux-Trockenlauf ohne pywin32 läuft).
hiddenimports = ['win32com', 'win32com.client']


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
    name='Buchdurchgang',
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
    name='Buchdurchgang',
)
