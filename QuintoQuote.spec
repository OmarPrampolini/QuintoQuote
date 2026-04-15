# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path


project_root = Path(SPECPATH).resolve().parent
docs_dir = project_root / "docs"

datas = []
if docs_dir.exists():
    datas.append((str(docs_dir), "docs"))


a = Analysis(
    ["quintoquote.py"],
    pathex=[str(project_root)],
    binaries=[],
    datas=datas,
    hiddenimports=["fitz"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="QuintoQuote",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="QuintoQuote",
)
