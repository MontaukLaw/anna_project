from threading import Event
from services.packing_generation_service import generate_packing_lists
from services.office_service import open_in_excel


def run_packing_job(matches, catalog, schedule, template, directory, events):
    def log(message, level="INFO"):
        events.put(("packing_log", (message, level)))

    def choose(item, candidates, packaging_hint):
        ready = Event()
        answer = []
        events.put(("packing_choice", (item, candidates, packaging_hint, ready, answer)))
        ready.wait()
        return answer[0] if answer else None

    try:
        generated, failed, skipped = generate_packing_lists(matches, catalog, schedule, template, directory, choose, log,
                                                          on_generated=open_in_excel)
        events.put(("packing_finished", (generated, failed, skipped)))
    except Exception as exc:
        events.put(("packing_error", str(exc)))
    finally:
        events.put(("packing_done", None))
