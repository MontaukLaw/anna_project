"""ASN step one: select a directory and asynchronously reconcile three source documents."""
from pathlib import Path
from queue import Empty, Queue
from threading import Thread
from tkinter import filedialog
import traceback

import customtkinter as ctk

from config import PROJECT_DIR, SETTINGS_FILE, ASN_LOG_DIR, ASN_TEMPLATE
from services.asn_service import audit_directory, result_messages
from services.asn_log_service import DailyAuditLog
from services.customs_packing_service import (read_customs_packing, row_message, packing_summary,
                                             result_columns, header_message)
from services.settings_service import get_saved_path, read_settings, save_file_path, save_settings
from services.asn_generation_service import build_asn_plan, write_asn_workbook, generation_options, product_key
from ui.asn_generation_dialog import AsnGenerationDialog
from ui import theme as t
from ui.components import LogPanel, button, label


class AsnTab(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, fg_color=t.BG)
        self.events = Queue()
        self.busy = False
        self.report = None
        self.directory = None
        self.customs_path = None
        self.customs_data = None
        self.output_directory = None
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)
        header = ctk.CTkFrame(self, fg_color='transparent')
        header.grid(row=0, column=0, sticky='ew', padx=24, pady=(22, 16))
        label(header, '入仓预申报ASN生成器', size=26).pack(anchor='w')
        label(header, '三单核对 · 海关装箱单明细', size=14, color=t.ACCENT).pack(anchor='w', pady=(6, 0))
        source = ctk.CTkFrame(self, fg_color=t.PANEL, corner_radius=12,
                             border_width=1, border_color=t.BORDER)
        source.grid(row=1, column=0, sticky='ew', padx=24)
        source.grid_columnconfigure(0, weight=1)
        self.description = label(source,
            '选择包含形式发票、装箱单、香港合同的目录，自动按文件名中的合同号核对。\n'
            '三单显示核对结果；选择海关装箱单后显示每行数据，详细记录按日期保存。',
            color=t.MUTED, justify='left', anchor='w')
        self.description.grid(row=0, column=0, columnspan=2, sticky='ew', padx=18, pady=(14, 10))
        source.bind('<Configure>', lambda e: self.description.configure(wraplength=max(200, e.width - 40)))
        self.path_var = ctk.StringVar(value='尚未选择单据目录')
        ctk.CTkEntry(source, textvariable=self.path_var, state='readonly', font=t.font(13),
                     height=40, fg_color=t.INSET, border_color=t.BORDER, text_color=t.TEXT).grid(
                         row=1, column=0, sticky='ew', padx=(18, 12), pady=(0, 16))
        self.select_button = button(source, '选择三文件所在的目录', self.browse, width=196)
        self.select_button.grid(row=1, column=1, padx=(0, 18), pady=(0, 16))
        self.customs_path_var = ctk.StringVar(value='尚未选择海关装箱单')
        ctk.CTkEntry(source, textvariable=self.customs_path_var, state='readonly', font=t.font(13),
                     height=40, fg_color=t.INSET, border_color=t.BORDER, text_color=t.TEXT).grid(
                         row=2, column=0, sticky='ew', padx=(18, 12), pady=(0, 16))
        self.customs_button = button(source, '选择海关装箱单', self.browse_customs, width=156)
        self.customs_button.grid(row=2, column=1, padx=(0, 18), pady=(0, 16))
        self.output_path_var = ctk.StringVar(value='尚未选择输出目录')
        ctk.CTkEntry(source, textvariable=self.output_path_var, state='readonly', font=t.font(13),
                     height=40, fg_color=t.INSET, border_color=t.BORDER, text_color=t.TEXT).grid(
                         row=3, column=0, sticky='ew', padx=(18, 12), pady=(0, 16))
        self.output_button = button(source, '选择输出目录', self.browse_output, width=156)
        self.output_button.grid(row=3, column=1, padx=(0, 18), pady=(0, 16))
        label(source, '核对通过并读取海关装箱单后，即可生成 ASN。', color=t.MUTED).grid(
            row=4, column=0, sticky='w', padx=18, pady=(0, 16))
        self.generate_button = button(source, '生成 ASN', self.generate, width=156, state='disabled')
        self.generate_button.grid(row=4, column=1, padx=(0, 18), pady=(0, 16))
        self.status = label(self, '等待选择目录', font=t.font(17, 'bold'), color=t.ACCENT,
                            anchor='w', justify='left')
        self.status.grid(row=2, column=0, sticky='ew', padx=26, pady=16)
        self.bind('<Configure>', lambda e: self.status.configure(wraplength=max(200, e.width - 60)))
        self.logs = LogPanel(self, title='核对结果与装箱明细', subtitle='三单核对 · 海关装箱单逐行数据', show_metadata=False)
        self.logs.grid(row=3, column=0, sticky='nsew', padx=24, pady=(0, 12))
        self.log_location = label(self, '', size=12, color=t.MUTED, anchor='w')
        self.log_location.grid(row=4, column=0, sticky='ew', padx=26, pady=(0, 6))
        label(self, 'ASN 明细以海关装箱单逐行为准；未确认的字段不会自动猜填。', size=12, color=t.MUTED).grid(
            row=5, column=0, sticky='w', padx=26, pady=(0, 12))
        self.logs.write('请选择单据目录。')
        try:
            settings = read_settings(SETTINGS_FILE)
            self.directory = get_saved_path(settings, SETTINGS_FILE, 'asn_directory_file')
            if self.directory:
                self.path_var.set(str(self.directory))
            self.customs_path = get_saved_path(settings, SETTINGS_FILE, 'asn_customs_file')
            if self.customs_path:
                self.customs_path_var.set(f'上次选择（尚未读取）：{self.customs_path}')
            self.output_directory = get_saved_path(settings, SETTINGS_FILE, 'asn_output_directory')
            if self.output_directory:
                self.output_path_var.set(str(self.output_directory))
        except (ValueError, OSError) as exc:
            self.logs.write(f'上次目录读取失败：{exc}', 'WARNING')
        self.after(100, self.poll)

    def browse(self):
        if self.busy:
            return
        directory = filedialog.askdirectory(parent=self, title='选择三文件所在的目录', mustexist=True,
            initialdir=self.directory if self.directory and self.directory.is_dir() else PROJECT_DIR)
        if directory:
            self.start_check(Path(directory))

    def start_check(self, directory):
        if self.busy:
            return
        self.directory = Path(directory)
        self.path_var.set(str(self.directory))
        self.report = None
        self.logs.clear()
        self.logs.textbox.configure(wrap='word')
        self.log_location.configure(text='')
        self.busy = True
        self.update_generate_state()
        self.customs_button.configure(state='disabled')
        self.select_button.configure(state='disabled', text='正在核对…')
        self.status.configure(text='正在检查三份单据，请稍候…', text_color=t.ACCENT)
        try:
            save_file_path(SETTINGS_FILE, self.directory, 'asn_directory')
        except (ValueError, OSError) as exc:
            self.logs.write(f'目录保存失败：{exc}', 'WARNING')
        Thread(target=self.run, args=(self.directory,), daemon=True).start()

    def run(self, directory):
        journal = DailyAuditLog(ASN_LOG_DIR)
        journal.write(f'本次核对开始：{directory}', 'START')
        try:
            result = audit_directory(directory, journal.write)
            outcome = ('done', result)
        except Exception as exc:
            journal.write(traceback.format_exc(), 'ERROR')
            outcome = ('error', str(exc))
        journal.write('本次核对结束', 'END')
        self.events.put(('log_file', (journal.paths, journal.error)))
        self.events.put(outcome)

    def browse_customs(self):
        if self.busy:
            return
        initial = self.customs_path.parent if self.customs_path else (self.directory or PROJECT_DIR)
        filename = filedialog.askopenfilename(parent=self, title='选择海关装箱单',
            initialdir=initial if initial.is_dir() else PROJECT_DIR, filetypes=[('Excel 海关装箱单', '*.xlsx')])
        if filename:
            self.start_customs_read(Path(filename))

    def start_customs_read(self, path):
        if self.busy:
            return
        self.customs_path = Path(path)
        self.customs_path_var.set(str(self.customs_path))
        self.customs_data = None  # Never leave old rows available after a failed replacement.
        self.logs.clear()
        self.logs.textbox.configure(wrap='none')
        self.logs.textbox.xview('moveto', 0)
        self.log_location.configure(text='')
        self.busy = True
        self.update_generate_state()
        self.select_button.configure(state='disabled')
        self.customs_button.configure(state='disabled', text='正在读取…')
        self.status.configure(text='正在读取海关装箱单明细…', text_color=t.ACCENT)
        Thread(target=self.run_customs, args=(self.customs_path,), daemon=True).start()

    def run_customs(self, path):
        journal = DailyAuditLog(ASN_LOG_DIR)
        journal.write(f'读取海关装箱单：{path}', 'START')
        try:
            data = read_customs_packing(path)
            summary = packing_summary(data)
            journal.write(summary, 'INFO')
            self.events.put(('customs_message', (summary, 'TITLE')))
            columns = result_columns(data)
            header = header_message(columns)
            journal.write(header)
            self.events.put(('customs_message', (header, 'INFO')))
            for index, row in enumerate(data.rows, 1):
                message = row_message(index, row, columns)
                journal.write(message)
                journal.write(f'{row.sheet} 第 {row.source_row} 行原始值：{row.raw_values!r}')
                self.events.put(('customs_message', (message, 'INFO')))
            for warning in data.warnings:
                journal.write(warning, 'WARNING')
                self.events.put(('customs_message', (warning, 'WARNING')))
            try:
                save_file_path(SETTINGS_FILE, path, 'asn_customs')
            except (OSError, ValueError) as exc:
                journal.write(f'文件路径保存失败：{exc}', 'WARNING')
                self.events.put(('customs_message', ('海关装箱单已读取，但文件路径保存失败。', 'WARNING')))
            outcome = ('customs_done', data)
        except Exception as exc:
            journal.write(traceback.format_exc(), 'ERROR')
            outcome = ('customs_error', str(exc))
        journal.write('海关装箱单读取结束', 'END')
        self.events.put(('log_file', (journal.paths, journal.error)))
        self.events.put(outcome)

    def browse_output(self):
        if self.busy:
            return False
        initial = self.output_directory or (self.customs_path.parent if self.customs_path else self.directory)
        directory = filedialog.askdirectory(parent=self, title='选择 ASN 输出目录', mustexist=True,
            initialdir=initial if initial and initial.is_dir() else PROJECT_DIR)
        if not directory:
            return False
        selected = Path(directory).resolve()
        if not selected.is_dir():
            self.logs.write('输出目录不存在，请重新选择。', 'ERROR')
            return False
        self.output_directory = selected
        self.output_path_var.set(str(selected))
        try:
            save_settings(SETTINGS_FILE, {'asn_output_directory': str(selected)})
        except (OSError, ValueError) as exc:
            self.logs.write(f'本次输出目录已选择，但配置保存失败：{exc}', 'WARNING')
        return True

    def update_generate_state(self):
        self.output_button.configure(state='disabled' if self.busy else 'normal')
        ready = (not self.busy and self.report is not None and self.report.passed
                 and self.customs_data is not None and not self.customs_data.warnings)
        self.generate_button.configure(state='normal' if ready else 'disabled')

    def generate(self):
        if self.busy or self.report is None or not self.report.passed or self.customs_data is None:
            return
        try:
            options = generation_options(self.report, self.customs_data)
            if not options['contract'] or any(product_key(row) not in options['product_mapping'] for row in self.customs_data.rows):
                dialog = AsnGenerationDialog(self, self.report, self.customs_data)
                self.wait_window(dialog)
                if dialog.result is None:
                    return
                options = dialog.result
            build_asn_plan(self.report, self.customs_data, ASN_TEMPLATE, options)
            options['source_items'] = [(item.name, str(item.price), item.code)
                                       for item in self.report.documents[options['contract']]['装箱单'].items]
            if self.output_directory is None or not self.output_directory.is_dir():
                if not self.browse_output():
                    return
            output = self.output_directory
        except (OSError, ValueError, KeyError) as exc:
            self.status.configure(text='ASN 未生成，请处理下方问题', text_color='#FF8585')
            self.logs.write(f'无法准备生成：{exc}', 'ERROR')
            journal = DailyAuditLog(ASN_LOG_DIR)
            journal.write(f'ASN 生成前检查失败：{exc}', 'ERROR')
            self.events.put(('log_file', (journal.paths, journal.error)))
            return
        self.busy = True
        self.update_generate_state()
        self.select_button.configure(state='disabled')
        self.customs_button.configure(state='disabled')
        self.generate_button.configure(text='正在生成…')
        self.logs.clear()
        self.logs.textbox.configure(wrap='word')
        self.status.configure(text='正在重新核对源文件并生成 ASN…', text_color=t.ACCENT)
        Thread(target=self.run_generation, args=(self.directory, self.customs_path, Path(output), options), daemon=True).start()

    def run_generation(self, directory, customs_path, output, options):
        journal = DailyAuditLog(ASN_LOG_DIR)
        journal.write(f'生成 ASN：三单目录={directory}；PL={customs_path}；输出={output}', 'START')
        try:
            # Fresh reads prevent a workbook changed since selection from being exported using stale data.
            report = audit_directory(directory, journal.write)
            packing = read_customs_packing(customs_path)
            plan = build_asn_plan(report, packing, ASN_TEMPLATE, options)
            journal.write(f'境外收货人（海关装箱单 Customer）：{plan["headers"]["C8"]}')
            journal.write(f'合同及商品对应：{options!r}；单位来自发票；第二数量=净重/千克；P/O=CUSTOMER ORDER NO；申报日期=当天；车牌留空')
            destination = write_asn_workbook(plan, output)
            journal.write(f'ASN 已保存：{destination}；明细 {len(plan["rows"])} 行', 'SUCCESS')
            outcome = ('generation_done', (destination, len(plan['rows'])))
        except Exception as exc:
            journal.write(traceback.format_exc(), 'ERROR')
            outcome = ('generation_error', str(exc))
        journal.write('ASN 生成结束', 'END')
        self.events.put(('log_file', (journal.paths, journal.error)))
        self.events.put(outcome)

    def poll(self):
        try:
            # Bound each batch so very large audits don't starve the Tk event loop.
            for _ in range(100):
                kind, value = self.events.get_nowait()
                if kind == 'log_file':
                    paths, error = value
                    if error:
                        self.logs.write(f'详细日志保存失败：{error}', 'WARNING')
                    if paths:
                        self.log_location.configure(text='详细日志：' + '；'.join(str(path) for path in paths))
                elif kind == 'customs_message':
                    self.logs.write(*value)
                elif kind in ('generation_done', 'generation_error'):
                    self.busy = False
                    self.select_button.configure(state='normal')
                    self.customs_button.configure(state='normal')
                    self.generate_button.configure(text='生成 ASN')
                    if kind == 'generation_error':
                        self.status.configure(text='ASN 未生成，请处理下方问题', text_color='#FF8585')
                        self.logs.write(f'生成失败：{value}', 'ERROR')
                    else:
                        path, count = value
                        self.status.configure(text=f'ASN 已生成 · {count} 行明细', text_color='#83DDB2')
                        self.logs.write(f'已保存：{path}', 'SUCCESS')
                    self.update_generate_state()
                elif kind in ('customs_done', 'customs_error'):
                    self.busy = False
                    self.select_button.configure(state='normal')
                    self.customs_button.configure(state='normal', text='选择海关装箱单')
                    if kind == 'customs_error':
                        self.logs.write(f'海关装箱单读取失败：{value}', 'ERROR')
                        self.status.configure(text='读取失败，请重新选择海关装箱单', text_color='#FF8585')
                    else:
                        self.customs_data = value
                        text = f'海关装箱单已读取 · {len(value.rows)} 行明细'
                        if value.warnings:
                            text += '，部分数据需要确认'
                        self.status.configure(text=text, text_color='#FFD18A' if value.warnings else '#83DDB2')
                        self.logs.textbox.yview('moveto', 0)
                        self.logs.textbox.xview('moveto', 0)
                    self.update_generate_state()
                elif kind in ('done', 'error'):
                    self.busy = False
                    self.select_button.configure(state='normal', text='选择三文件所在的目录')
                    self.customs_button.configure(state='normal')
                    if kind == 'error':
                        self.logs.write(f'核对失败：{value}', 'ERROR')
                        self.status.configure(text='核对未完成，请处理下方问题后重试', text_color='#FF8585')
                    else:
                        self.report = value
                        for message, level in result_messages(value):
                            self.logs.write(message, level)
                        text = (f'核对通过 · {len(value.documents)} 个合同'
                                if value.passed else '核对未通过，请处理下方问题')
                        self.status.configure(text=text, text_color='#83DDB2' if value.passed else '#FF8585')
                    self.update_generate_state()
        except Empty:
            pass
        self.after(100, self.poll)
