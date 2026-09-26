# Build: .venv/Scripts/python.exe -m PyInstaller --noconfirm LoveAnna.spec
from pathlib import Path
from PyInstaller.utils.hooks import collect_data_files

root = Path(SPECPATH)
datas = [
    (str(root / 'templates'), 'templates'),
    (str(root / 'packing_rules.json'), '.'),
    (str(root / 'services' / 'asn_excel_writer.ps1'), 'services'),
] + collect_data_files('customtkinter')

a = Analysis(
    [str(root / 'main.py')], pathex=[str(root)],
    binaries=[], datas=datas,
    hiddenimports=[],
    excludes=['pytest', 'IPython', 'pymupdf', 'fitz', 'matplotlib', 'torch', 'paddle', 'tensorflow'],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, a.binaries, a.datas, [],
    name='LoveAnna', debug=False, bootloader_ignore_signals=False,
    strip=False, upx=False, console=False,
)
