"""用本机 Microsoft Excel 打开已保存的装箱单。"""
import os
from pathlib import Path
import shutil
import subprocess


def find_excel() -> Path:
    if os.name != "nt":
        raise OSError("自动打开 Microsoft Excel 目前仅支持 Windows")
    import winreg
    key_name = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe"
    for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
        for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
            try:
                with winreg.OpenKey(hive, key_name, 0, winreg.KEY_READ | view) as key:
                    value, _ = winreg.QueryValueEx(key, None)
                executable = Path(os.path.expandvars(value.strip('"')))
                if executable.is_file():
                    return executable
            except OSError:
                continue
    executable = shutil.which("excel.exe")
    if executable:
        return Path(executable)
    raise FileNotFoundError("未找到 Microsoft Excel，请确认已安装 Office")


def open_in_excel(path: Path) -> None:
    path = path.resolve()
    if not path.is_file():
        raise FileNotFoundError(f"Excel 文件不存在：{path}")
    # Argument list preserves Chinese paths and spaces without invoking a shell.
    subprocess.Popen([str(find_excel()), str(path)], shell=False)
