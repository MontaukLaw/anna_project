"""Append ASN audit details to one UTF-8 file per local calendar day."""
from datetime import datetime
from pathlib import Path
from threading import Lock
from uuid import uuid4


_write_lock = Lock()


class DailyAuditLog:
    def __init__(self, directory, now=None):
        self.directory = Path(directory)
        self.now = now or datetime.now
        self.run_id = uuid4().hex[:8]
        self.paths = []
        self.error = None

    def write(self, message, level='INFO'):
        # A logging failure must not abort the document checks or go unnoticed.
        if self.error is not None:
            return
        timestamp = self.now()
        path = self.directory / f'asn-{timestamp:%Y-%m-%d}.log'
        try:
            with _write_lock:
                self.directory.mkdir(parents=True, exist_ok=True)
                with path.open('a', encoding='utf-8') as stream:
                    stream.write(f'{timestamp:%Y-%m-%d %H:%M:%S}  [{level}]  [{self.run_id}]  {message}\n')
            if path not in self.paths:
                self.paths.append(path)
        except OSError as exc:
            self.error = f'{path}：{exc}'
