"""Independent factory notice tab; all Tk work stays on the main thread."""
from pathlib import Path
from queue import Empty, Queue
from threading import Event, Thread
from tkinter import filedialog, messagebox

import customtkinter as ctk

from config import PROJECT_DIR, SETTINGS_FILE, FACTORY_TEMPLATE
from services.factory_excel_service import write_factory_workbook
from services.factory_packing_service import collect_factory_rows, scan_pdfs
from services.office_service import open_in_excel
from services.product_catalog_service import read_product_catalog
from services.schedule_service import read_order_schedule
from services.settings_service import read_settings, get_saved_path, save_file_path
from ui import theme as t
from ui.components import button, label, LogPanel
from ui.packing_choice_dialog import PackingChoiceDialog


class FactoryOptionDialog(ctk.CTkToplevel):
    def __init__(self, parent, title, context, options):
        super().__init__(parent)
        self.title(title)
        self.geometry('850x600')
        self.result = None
        self.selection = ctk.IntVar(value=-1)
        self.configure(fg_color=t.BG)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        label(self, context, color=t.ACCENT, wraplength=780, justify='left').grid(
            row=0, column=0, padx=20, pady=16, sticky='ew')
        frame = ctk.CTkScrollableFrame(self, fg_color=t.PANEL)
        frame.grid(row=1, column=0, sticky='nsew', padx=20)
        frame.grid_columnconfigure(0, weight=1)
        for index, text in enumerate(options):
            option = ctk.CTkFrame(frame, fg_color=t.INSET)
            option.grid(row=index, column=0, sticky='ew', pady=6)
            ctk.CTkRadioButton(option, text=f'选择第 {index + 1} 项', variable=self.selection,
                               value=index, font=t.font(13)).pack(anchor='w', padx=12, pady=8)
            label(option, text, wraplength=720, justify='left').pack(anchor='w', padx=12, pady=(0, 10))
        actions = ctk.CTkFrame(self, fg_color='transparent')
        actions.grid(row=2, column=0, pady=16)
        button(actions, '取消本次生成', self.destroy).pack(side='left', padx=10)
        button(actions, '使用所选内容', self.confirm).pack(side='left')
        self.transient(parent.winfo_toplevel())
        self.after(100, self.grab_set)

    def confirm(self):
        if self.selection.get() < 0:
            messagebox.showwarning('请选择', '请选择需要使用的一项。', parent=self)
            return
        self.result = self.selection.get()
        self.destroy()


class FactoryTab(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, fg_color=t.BG)
        self.events = Queue()
        self.paths = {}
        self.catalog = self.schedule = None
        self.loading = set()
        self.busy = False
        self.buttons = []
        self.path_vars = {}
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)
        label(self, '自动工厂箱单生成', size=25).grid(row=0, column=0, padx=22, pady=(20, 12), sticky='w')
        sources = ctk.CTkFrame(self, fg_color=t.PANEL)
        sources.grid(row=1, column=0, padx=22, sticky='ew')
        sources.grid_columnconfigure(1, weight=1)
        for row, (key, title) in enumerate((('orders', '1. 订单 PDF 目录'),
                                          ('schedule', '2. 订单分类排期表'),
                                          ('product', '3. JP 产品资料表'),
                                          ('customers', '4. 客户订单表目录'),
                                          ('output', '输出目录'))):
            control = button(sources, title, lambda k=key: self.browse(k), width=175)
            control.grid(row=row, column=0, padx=12, pady=7)
            self.buttons.append(control)
            value = ctk.StringVar(value='尚未选择')
            self.path_vars[key] = value
            ctk.CTkEntry(sources, textvariable=value, state='readonly', height=36,
                        fg_color=t.INSET, text_color=t.TEXT, font=t.font(13)).grid(row=row, column=1, padx=(0, 12), pady=7, sticky='ew')
        toolbar = ctk.CTkFrame(self, fg_color='transparent')
        toolbar.grid(row=2, column=0, sticky='ew', padx=22, pady=14)
        self.generate_button = button(toolbar, '生成合并工厂箱单', self.generate, width=175, state='disabled')
        self.generate_button.pack(side='left')
        self.status = label(toolbar, '每个 item 一行 · 全部订单合并 · 文件名为总箱数.xlsx', color=t.MUTED)
        self.status.pack(side='left', padx=18)
        self.logs = LogPanel(self)
        self.logs.grid(row=3, column=0, sticky='nsew', padx=22, pady=(0, 18))
        self.logs.write('请选择资料和两个 PDF 目录。包含子目录；生成结果保存后自动用 Excel 打开。')
        self.after(100, self.poll)
        self.after(500, self.restore)

    def remember(self, key, path):
        try:
            save_file_path(SETTINGS_FILE, path, f'factory_{key}')
        except (OSError, ValueError) as exc:
            self.logs.write(f'路径保存失败：{exc}', 'WARNING')

    def restore(self):
        try:
            settings = read_settings(SETTINGS_FILE)
        except (OSError, ValueError) as exc:
            self.logs.write(f'读取配置失败：{exc}', 'WARNING')
            settings = {}
        for key in self.path_vars:
            path = get_saved_path(settings, SETTINGS_FILE, f'factory_{key}_file')
            if path is not None and not path.exists():
                self.logs.write(f'上次选择的路径已不存在，请重新选择：{path}', 'WARNING')
                continue
            if path is None and key == 'output':
                path = PROJECT_DIR / 'output' / 'factory_packing_lists'
            if path is not None:
                self.set_path(key, path, remember=False)
        self.update_state()

    def browse(self, key):
        if self.busy or key in self.loading:
            return
        initial = self.paths.get(key, PROJECT_DIR)
        if key in ('product', 'schedule'):
            value = filedialog.askopenfilename(parent=self, title='选择 Excel 资料表',
                initialdir=initial.parent if initial.is_file() else PROJECT_DIR,
                filetypes=[('Excel 工作簿', '*.xlsx')])
        else:
            value = filedialog.askdirectory(parent=self, title='选择' + {'orders': '订单 PDF 目录',
                'customers': '客户订单表目录', 'output': '输出目录'}[key],
                initialdir=initial if initial.is_dir() else PROJECT_DIR, mustexist=True)
        if value:
            self.set_path(key, Path(value))

    def set_path(self, key, path, remember=True):
        self.paths[key] = path
        self.path_vars[key].set(str(path))
        if key in ('product', 'schedule'):
            self.loading.add(key)
            if key == 'product':
                self.catalog = None
            else:
                self.schedule = None
            self.logs.write(f'正在自动读取：{path}')
            Thread(target=self.load, args=(key, path, remember), daemon=True).start()
        else:
            if remember:
                self.remember(key, path)
            if key in ('orders', 'customers'):
                self.loading.add(key)
                Thread(target=self.scan, args=(key, path), daemon=True).start()
        self.update_state()

    def scan(self, key, path):
        try:
            files = scan_pdfs(path)
            self.events.put(('log', f'{"订单" if key == "orders" else "客户订单表"}目录找到 {len(files)} 个 PDF：{path}'))
            for file in files:
                self.events.put(('log', f'  {file.relative_to(path)}'))
        except Exception as exc:
            self.events.put(('log', f'目录扫描失败：{exc}'))
        finally:
            self.events.put(('loaded', (key, None, path, False)))

    def load(self, key, path, remember):
        try:
            data = read_product_catalog(path) if key == 'product' else read_order_schedule(path)
            self.events.put(('loaded', (key, data, path, remember)))
        except Exception as exc:
            self.events.put(('log', f'资料读取失败，请重新选择：{path}：{exc}'))
            self.events.put(('loaded', (key, None, path, False)))

    def update_state(self):
        for (key, _), control in zip(self.path_vars.items(), self.buttons):
            control.configure(state='disabled' if self.busy or key in self.loading else 'normal')
        ready = (not self.busy and not self.loading and self.catalog is not None and self.schedule is not None
                 and all(key in self.paths for key in ('orders', 'customers', 'output')))
        self.generate_button.configure(state='normal' if ready else 'disabled',
                                       text='正在生成…' if self.busy else '生成合并工厂箱单')

    def ask_worker(self, kind, *args):
        ready, answer = Event(), []
        self.events.put((kind, (args, ready, answer)))
        ready.wait()
        return answer[0] if answer else None

    def generate(self):
        if self.busy or self.loading or self.catalog is None or self.schedule is None:
            return
        self.busy = True
        self.status.configure(text='正在读取全部订单并合并，请按提示选择冲突资料…')
        self.update_state()
        Thread(target=self.run, args=(dict(self.paths), self.catalog, self.schedule), daemon=True).start()

    def run(self, paths, catalog, schedule):
        try:
            rows = collect_factory_rows(paths['orders'], paths['customers'], catalog, schedule,
                lambda *args: self.ask_worker('product_choice', *args),
                lambda *args: self.ask_worker('option_choice', *args),
                lambda message: self.events.put(('log', message)))
            template = FACTORY_TEMPLATE
            destination = write_factory_workbook(rows, template, paths['output'])
            self.events.put(('success', (destination, len(rows), sum(row.cartons for row in rows))))
            try:
                open_in_excel(destination)
            except Exception as exc:
                self.events.put(('log', f'文件已保存，但自动打开 Excel 失败：{exc}'))
        except Exception as exc:
            self.events.put(('error', str(exc)))
        finally:
            self.events.put(('done', None))

    def poll(self):
        try:
            for _ in range(60):
                kind, payload = self.events.get_nowait()
                if kind == 'log':
                    self.logs.write(payload)
                elif kind == 'loaded':
                    key, data, path, remember = payload
                    self.loading.discard(key)
                    if key == 'product':
                        self.catalog = data
                    elif key == 'schedule':
                        self.schedule = data
                    if data is not None:
                        self.logs.write(f'资料已加载：{path}')
                        if remember:
                            self.remember(key, path)
                    self.update_state()
                elif kind in ('product_choice', 'option_choice'):
                    args, ready, answer = payload
                    try:
                        if kind == 'product_choice':
                            dialog = PackingChoiceDialog(self.winfo_toplevel(), *args, factory=True)
                        else:
                            dialog = FactoryOptionDialog(self, *args)
                        self.wait_window(dialog)
                        answer.append(dialog.result)
                    except Exception as exc:
                        self.logs.write(f'选择窗口失败：{exc}', 'ERROR')
                    finally:
                        ready.set()
                elif kind == 'success':
                    destination, count, total = payload
                    self.logs.write(f'已生成 {count} 行 / {total} 箱：{destination}')
                    self.status.configure(text=f'生成完成：{count} 行 / {total} 箱 / {destination.name}')
                elif kind == 'error':
                    self.logs.write(f'生成已停止，未输出本次箱单：{payload}', 'ERROR')
                    self.status.configure(text='生成已停止，请查看日志并处理后重试')
                    messagebox.showerror('工厂箱单生成停止', payload, parent=self)
                elif kind == 'done':
                    self.busy = False
                    self.update_state()
        except Empty:
            pass
        self.after(100, self.poll)
