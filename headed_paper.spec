# PyInstaller spec for the Headed Paper PDF Generator.
# Build on Windows with:  pyinstaller headed_paper.spec

a = Analysis(
    ["headed_paper_pdf.py"],
    pathex=[],
    binaries=[],
    datas=[("letterhead.pdf", ".")],
    hiddenimports=[],
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
    a.binaries,
    a.datas,
    [],
    name="HeadedPaperPDF",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
)
