# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller build spec for Bible2PPT.

Windows is the only shipped target for this release (produces a single
``Bible2PPT.exe``). The spec is intentionally organised so a macOS ``.app``
branch is a small edit rather than a rewrite: the data-collection and Analysis
are OS-independent, and the OS-specific bits (icon, one-file vs .app bundle) are
isolated below behind ``sys.platform`` checks.

Build:
    pip install -r requirements-dev.txt
    pyinstaller Bible2PPT.spec
"""
import sys
from pathlib import Path

block_cipher = None
ROOT = Path(SPECPATH)

# --- bundled data (OS-independent) ---------------------------------------- #
# Everything under data/ ships read-only inside the executable; core.paths
# resolves it via sys._MEIPASS at runtime.
datas = [
    ("data/canon.json", "data"),
    ("data/ppt배경.png", "data"),
    ("data/bibles", "data/bibles"),
    ("data/i18n", "data/i18n"),
    ("data/versification", "data/versification"),
    ("data/fonts", "data/fonts"),
]

a = Analysis(
    ["main.py"],
    pathex=[str(ROOT)],
    binaries=[],
    datas=datas,
    hiddenimports=["PIL._tkinter_finder"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["numpy", "pandas", "matplotlib"],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# --- OS-specific packaging ------------------------------------------------ #
if sys.platform == "darwin":
    # macOS port target: produce an .app bundle. Provide run_icon.icns and
    # flip this branch on once a signing/notarization pipeline exists.
    icon = "run_icon.icns" if Path(ROOT / "run_icon.icns").exists() else None
else:
    icon = "run_icon.ico"

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="Bible2PPT",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon,
)

# macOS .app bundle (only assembled on macOS; harmless stub elsewhere).
if sys.platform == "darwin":
    app = BUNDLE(
        exe,
        name="Bible2PPT.app",
        icon=icon,
        bundle_identifier="ai.cognition.bible2ppt",
    )
