"""Persistent successful HS lookups and an audit trail of every network attempt."""
from datetime import datetime
import json
from pathlib import Path
import shutil
from tempfile import NamedTemporaryFile
from threading import Lock
from time import time_ns


_cache_lock = Lock()


class HSUnitCache:
    def __init__(self, path, log=None):
        self.path = Path(path)
        self.log = log

    def warn(self, message):
        if self.log:
            self.log(f'警告：商品编码缓存{message}（{self.path}）')

    def read(self):
        try:
            data = json.loads(self.path.read_text(encoding='utf-8-sig'))
        except FileNotFoundError:
            return {'version': 1, 'entries': {}, 'queries': []}
        if (not isinstance(data, dict) or data.get('version') != 1
                or not isinstance(data.get('entries'), dict) or not isinstance(data.get('queries'), list)):
            raise ValueError('JSON 结构无效')
        return data

    def get(self, code):
        return self.get_many([code]).get(code)

    def get_many(self, codes):
        """Read the local JSON once before any network lookup in this batch."""
        with _cache_lock:
            try:
                entries = self.read()['entries']
                return {code: entries.get(code) for code in codes}
            except (OSError, ValueError) as exc:
                self.warn(f'读取失败，将重新联网查询：{exc}')
                return {}

    def record(self, code, url, units=None, error=None):
        record = {'code': code, 'queried_at': datetime.now().astimezone().isoformat(timespec='seconds'),
                  'url': url, 'status': 'success' if units else 'error'}
        if units:
            record.update(first=units.first, second=units.second)
        else:
            record['error'] = str(error)
        temporary = None
        with _cache_lock:
            try:
                try:
                    data = self.read()
                except ValueError:
                    # Preserve malformed content for inspection before recreating the cache.
                    backup = self.path.with_name(f'{self.path.name}.invalid-{time_ns()}.bak')
                    shutil.copy2(self.path, backup)
                    self.warn(f'内容损坏，原文件已备份至 {backup.name}')
                    data = {'version': 1, 'entries': {}, 'queries': []}
                data['queries'].append(record)
                if units:
                    data['entries'][code] = record
                self.path.parent.mkdir(parents=True, exist_ok=True)
                with NamedTemporaryFile(mode='w', encoding='utf-8', dir=self.path.parent,
                                        suffix='.tmp', delete=False) as stream:
                    temporary = Path(stream.name)
                    json.dump(data, stream, ensure_ascii=False, indent=2)
                    stream.write('\n')
                temporary.replace(self.path)
            except (OSError, ValueError) as exc:
                self.warn(f'保存失败，本次查询结果仍可使用：{exc}')
            finally:
                if temporary is not None:
                    try:
                        temporary.unlink(missing_ok=True)
                    except OSError as exc:
                        self.warn(f'临时文件清理失败：{exc}')
