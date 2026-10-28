# PyInstaller build spec for the desktop .exe.  Build with:
#     pip install -r requirements.txt -r requirements-desktop.txt
#     pyinstaller SheetJournal.spec
# Result: dist/SheetJournal.exe (single file, no console window).
#
# If an "antiword" folder (antiword.exe + its DLLs + share/ mapping files) is
# present next to this spec, it is bundled so .doc files can be read on
# machines without Git for Windows. The GitHub Actions workflow prepares it.

import os

datas = [("templates", "templates")]
if os.path.isdir("antiword"):
    datas.append(("antiword", "antiword"))

a = Analysis(
    ["desktop.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=[],
    excludes=["matplotlib", "scipy", "IPython", "pytest"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="SheetJournal",
    debug=False,
    strip=False,
    upx=False,
    console=False,
    icon=None,
)
