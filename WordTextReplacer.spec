# Build from the repository root with: pyinstaller WordTextReplacer.spec
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = collect_submodules("openpyxl") + collect_submodules("docx")

a = Analysis(
    ["src/word_text_replacer_single_with_add.py"],
    pathex=["src"],
    binaries=[],
    datas=[],
    hiddenimports=hiddenimports,
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
    name="WordTextReplacer",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
)
