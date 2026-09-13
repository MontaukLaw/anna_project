"""订单目录扫描：递归查找 PDF，通过队列通知界面。"""
import os
from pathlib import Path
from queue import Queue


def scan_order_directory(directory: Path, events: Queue) -> None:
    count = 0
    warnings = 0

    def report_error(error):
        nonlocal warnings
        warnings += 1
        events.put(("pdf_scan_warning", f"无法读取：{error}"))

    try:
        if not directory.is_dir():
            raise FileNotFoundError(f"目录不存在或无法访问：{directory}")
        for root, folders, filenames in os.walk(directory, onerror=report_error, followlinks=False):
            folders.sort(key=str.casefold)
            for filename in sorted(filenames, key=str.casefold):
                path = Path(root) / filename
                if path.suffix.lower() != ".pdf":
                    continue
                count += 1
                events.put(("pdf_scan_file", (count, path, str(path.relative_to(directory)))))
        events.put(("pdf_scan_summary", (count, warnings)))
    except Exception as exc:
        events.put(("pdf_scan_error", str(exc)))
    finally:
        events.put(("pdf_scan_done", None))
