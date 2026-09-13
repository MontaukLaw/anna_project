"""后台识别任务：通过队列发送事件，不访问 Tk 控件。"""
from datetime import datetime
from pathlib import Path
from queue import Queue
from uuid import uuid4

from services.excel_service import write_table
from services.ocr_service import TableRecognitionService


def run_recognition(service: TableRecognitionService, source: Path, output_dir: Path, events: Queue):
    try:
        table = service.recognize(source)
        events.put(("table", table))
        name = f"{source.stem}_{datetime.now():%Y%m%d_%H%M%S_%f}_{uuid4().hex[:6]}.xlsx"
        try:
            destination = write_table(table, output_dir / name)
        except Exception as exc:
            events.put(("save_error", str(exc)))
        else:
            events.put(("saved", destination))
    except Exception as exc:
        events.put(("error", str(exc)))
    finally:
        events.put(("done", None))
