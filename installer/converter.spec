# PyInstaller build recipe shared by Windows and macOS.
# Run from the repository root:  pyinstaller installer/converter.spec --noconfirm
import sys
from pathlib import Path

APP_NAME = "Assessment PDF to Excel"
ROOT = Path(SPECPATH).parent

a = Analysis(
    [str(ROOT / "app.py")],
    pathex=[str(ROOT)],
    hiddenimports=["openpyxl", "pymupdf"],
    excludes=["unittest", "pydoc", "test"],
    noarchive=False,
)
pyz = PYZ(a.pure)

if sys.platform == "darwin":
    # Folder-style build wrapped in a .app bundle (starts faster than one-file).
    exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name=APP_NAME,
              console=False, upx=False)
    coll = COLLECT(exe, a.binaries, a.datas, name=APP_NAME, upx=False)
    app = BUNDLE(
        coll,
        name=f"{APP_NAME}.app",
        bundle_identifier="com.techspire.assessment-pdf-to-excel",
        info_plist={
            "CFBundleShortVersionString": "1.0.0",
            "NSHighResolutionCapable": True,
        },
    )
else:
    # Single portable .exe: copy it anywhere and double-click.
    exe = EXE(pyz, a.scripts, a.binaries, a.datas, [], name=APP_NAME,
              console=False, upx=False)
