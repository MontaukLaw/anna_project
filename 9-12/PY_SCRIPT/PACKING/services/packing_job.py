from threading import Event
from queue import Queue
from services.order_directory_service import scan_order_directory
from services.order_pdf_service import read_order_number
from services.packing_generation_service import generate_packing_lists
from services.office_service import open_in_excel


def run_packing_job(matches, catalog, schedule, template, directory, events, pdf_directory=None):
    def log(message, level="INFO"):
        events.put(("packing_log", (message, level)))

    def choose(item, candidates, packaging_hint):
        ready = Event()
        answer = []
        events.put(("packing_choice", (item, candidates, packaging_hint, ready, answer)))
        ready.wait()
        return answer[0] if answer else None

    try:
        if pdf_directory is None:
            generated, failed, skipped = generate_packing_lists(matches, catalog, schedule, template, directory, choose, log,
                                                              on_generated=open_in_excel)
        else:
            scanned = Queue()
            scan_order_directory(pdf_directory, scanned)
            files = []
            while not scanned.empty():
                event, payload = scanned.get_nowait()
                if event == "pdf_scan_file":
                    files.append(payload[1])
                elif event == "pdf_scan_warning":
                    log(payload, "WARNING")
                elif event == "pdf_scan_error":
                    raise RuntimeError(payload)
            events.put(("packing_pdf_files", files))
            log(f"直接处理全部 PDF：共 {len(files)} 份（包含子目录）。")
            generated, failed, skipped = [], [], []
            for index, pdf in enumerate(files, 1):
                log(f"处理 PDF {index}/{len(files)}：{pdf}")
                try:
                    order = read_order_number(pdf)
                    result = generate_packing_lists({order: [pdf]}, catalog, schedule, template, directory,
                                                    choose, log, on_generated=open_in_excel)
                    generated.extend(result[0])
                    failed.extend(result[1])
                    skipped.extend(result[2])
                except Exception as exc:
                    failed.append(str(pdf))
                    log(f"PDF 未生成：{pdf}；{exc}", "ERROR")
        events.put(("packing_finished", (generated, failed, skipped)))
    except Exception as exc:
        events.put(("packing_error", str(exc)))
    finally:
        events.put(("packing_done", None))
