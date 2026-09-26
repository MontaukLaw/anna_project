from pathlib import Path
from queue import Empty, Queue
from threading import Thread
from tkinter import filedialog, messagebox
import customtkinter as ctk

from config import APP_NAME, WORKSPACE_DIR, SETTINGS_FILE, DEFAULT_JP_FILE, PROJECT_DIR, PACKING_RULES_FILE, RESOURCE_DIR
from services.settings_service import read_settings, get_product_path, find_local_product_files, ensure_settings
from services.product_catalog_service import load_product_catalog_job
from services.settings_service import find_local_schedule_files, get_saved_path, save_file_path
from services.schedule_service import load_order_schedule_job
from config import PACKING_OUTPUT_DIR, PACKING_TEMPLATE
from services.packing_job import run_packing_job
from ui.packing_choice_dialog import PackingChoiceDialog
from services.order_directory_service import scan_order_directory
from ui import theme as t
from ui.components import LogPanel, Section, button, label
from ui.factory_tab import FactoryTab
from ui.asn_tab import AsnTab
from ui.font_controls import FontControls


class PackingApp(ctk.CTk):
    def __init__(self):
        ctk.set_appearance_mode("dark")
        super().__init__()
        settings_error = None
        try:
            ensure_settings(SETTINGS_FILE, {'ui_font_size': t.DEFAULT_FONT_SIZE})
            ensure_settings(PACKING_RULES_FILE, read_settings(RESOURCE_DIR / 'packing_rules.json') or {'dimension_tolerance_percent': 4})
            t.set_font_size(read_settings(SETTINGS_FILE).get('ui_font_size', t.DEFAULT_FONT_SIZE))
        except (OSError, ValueError) as exc:
            t.set_font_size(t.DEFAULT_FONT_SIZE)
            settings_error = str(exc)
        self.title("LOVE ANNA, 加油!")
        self.geometry("1180x880")
        self.minsize(960, 760)
        self.configure(fg_color=t.BG)
        self.order_search_directory: Path | None = None
        self.order_pdf_files: list[Path] = []
        self.directory_scan_busy = False
        self.directory_scan_ready = False
        self.directory_scan_warnings = 0
        self.excel_output_directory = PACKING_OUTPUT_DIR
        self.events = Queue()
        self.product_catalog = None
        self.product_loading = False
        self.order_schedule = None
        self.schedule_loading = False
        self.packing_busy = False
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.tabs = ctk.CTkTabview(self, fg_color=t.BG)
        self.tabs._segmented_button.configure(font=t.font(13))
        self.tabs.grid(row=0, column=0, sticky='nsew', padx=8, pady=8)
        self.workspace = self.tabs.add('自动装箱生成')
        self.workspace.grid_columnconfigure(0, weight=1)
        self.workspace.grid_rowconfigure(2, weight=1, uniform='workspace_panels')
        self.workspace.grid_rowconfigure(3, weight=3, uniform='workspace_panels')
        factory_parent = self.tabs.add('自动工厂箱单生成')
        self.factory_tab = FactoryTab(factory_parent)
        self.factory_tab.pack(fill='both', expand=True)
        asn_parent = self.tabs.add('入仓预申报ASN生成器')
        self.asn_tab = AsnTab(asn_parent)
        self.asn_tab.pack(fill='both', expand=True)
        self.tabs.set('入仓预申报ASN生成器')
        self.font_controls = FontControls(self, SETTINGS_FILE, self.asn_tab.logs.write)
        self.font_controls.grid(row=1, column=0, sticky='e', padx=18, pady=(0, 10))
        if settings_error:
            self.asn_tab.logs.write(f'字号配置读取失败，已使用默认字号：{settings_error}', 'WARNING')
        self._build_header()
        self._build_source()
        self._build_table()
        self.logs = LogPanel(self.workspace)
        self.logs.grid(row=3, column=0, sticky="nsew", padx=28, pady=(0, 12))
        label(self.workspace, "本地工作空间  /  订单 PDF → 装箱单 Excel", size=11, color=t.MUTED).grid(
            row=4, column=0, sticky="w", padx=30, pady=(0, 14))
        self.logs.write("系统已就绪。请选择订单 PDF 目录，加载资料表后直接处理全部订单。")
        self.logs.write(f"装箱单输出目录：{self.excel_output_directory}")
        self.after(100, self._poll_events)
        # Restore PDF-only resources when that page is first opened, so an
        # unrelated product file picker cannot interrupt the new ASN home page.
        self._packing_restored = False
        self.tabs.configure(command=self._on_tab_change)

    def _on_tab_change(self):
        if self.tabs.get() == '自动装箱生成' and not self._packing_restored:
            self._packing_restored = True
            self._restore_product_catalog()
            self._restore_order_schedule()
            self._restore_last_selections()

    def _build_header(self):
        header = ctk.CTkFrame(self.workspace, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=30, pady=(24, 18))
        label(header, APP_NAME, size=27).pack(anchor="w")
        label(header, "PACKING WORKSPACE    /    箱单数据工作台", size=12, color=t.MUTED).pack(anchor="w", pady=(5, 0))
        badge = label(header, "  PDF 订单 · 本地运行  ", size=12, color=t.ACCENT,
                      fg_color=t.PANEL, corner_radius=6)
        badge.place(relx=1, y=8, anchor="ne")

    def _build_source(self):
        panel = Section(self.workspace, "01", "订单目录与资料", "选择 PDF 目录、产品资料、订单排期和输出目录")
        panel.grid(row=1, column=0, sticky="ew", padx=28, pady=(0, 16))
        controls = ctk.CTkFrame(panel, fg_color="transparent")
        controls.grid(row=1, column=0, sticky="ew", padx=20, pady=(2, 10))
        for column in (2, 4, 6, 8):
            controls.grid_columnconfigure(column, weight=1, uniform="source_paths")
        self.order_directory_path = ctk.StringVar(value="尚未选择订单搜索目录")
        self.order_directory_entry = ctk.CTkEntry(
            controls, textvariable=self.order_directory_path, state="readonly",
            height=38, fg_color=t.INSET, border_color=t.BORDER,
            text_color=t.MUTED, font=t.font(12))
        self.order_directory_entry.grid(row=0, column=2, sticky="ew", padx=(0, 8))
        self.order_directory_button = button(
            controls, "搜索订单目录", self.browse_order_directory, width=110)
        self.order_directory_button.grid(row=0, column=3, padx=(0, 12))
        self.product_path = ctk.StringVar(value="JP 产品资料表：尚未选择")
        self.product_entry = ctk.CTkEntry(
            controls, textvariable=self.product_path, state="readonly", height=38,
            fg_color=t.INSET, border_color=t.BORDER, text_color=t.MUTED, font=t.font(12))
        self.product_entry.grid(row=0, column=4, sticky="ew", padx=(0, 8))
        self.product_button = button(controls, "选择 JP 产品资料表", self.browse_product_catalog, width=144)
        self.product_button.grid(row=0, column=5, padx=(0, 12))
        self.schedule_path = ctk.StringVar(value="订单排期表：尚未选择")
        self.schedule_entry = ctk.CTkEntry(
            controls, textvariable=self.schedule_path, state="readonly", height=38,
            fg_color=t.INSET, border_color=t.BORDER, text_color=t.MUTED, font=t.font(12))
        self.schedule_entry.grid(row=0, column=6, sticky="ew", padx=(0, 8))
        self.schedule_button = button(controls, "选择订单排期表", self.browse_order_schedule, width=120)
        self.schedule_button.grid(row=0, column=7, padx=(0, 12))
        self.output_directory_path = ctk.StringVar(value=str(self.excel_output_directory))
        self.output_directory_entry = ctk.CTkEntry(
            controls, textvariable=self.output_directory_path, state="readonly", height=38,
            fg_color=t.INSET, border_color=t.BORDER, text_color=t.MUTED, font=t.font(12))
        self.output_directory_entry.grid(row=0, column=8, sticky="ew", padx=(0, 8))
        self.output_directory_button = button(controls, "输出目录", self.browse_output_directory, width=88)
        self.output_directory_button.grid(row=0, column=9)

    def _build_table(self):
        panel = Section(self.workspace, "02", "订单 PDF 列表", "包含子目录中的全部 PDF，逐份生成装箱单")
        panel.grid(row=2, column=0, sticky="nsew", padx=28, pady=(0, 16))
        panel.grid_rowconfigure(2, weight=1)
        toolbar = ctk.CTkFrame(panel, fg_color="transparent")
        toolbar.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 10))
        self.all_pdf_button = button(toolbar, "直接处理所有pdf订单", self.generate_all_pdf_packing, width=180)
        self.all_pdf_button.pack(side="left", padx=(0, 10))
        self.table_info = label(toolbar, "0 个 PDF", size=12, color=t.MUTED)
        self.table_info.pack(side="right")
        self.table_text = ctk.CTkTextbox(panel, fg_color=t.INSET, text_color=t.TEXT,
                                        font=t.font(14), wrap="none", height=60)
        self.table_text.grid(row=2, column=0, sticky="nsew", padx=20, pady=(0, 16))
        self._set_table_text("请选择订单 PDF 目录。\n\n扫描后在此列出全部 PDF；加载产品资料表和订单排期表后，点击“直接处理所有pdf订单”。")

    def _remember_selection(self, prefix, path):
        try:
            save_file_path(SETTINGS_FILE, path, prefix)
        except (OSError, ValueError) as exc:
            self.logs.write(f"选择路径保存失败：{exc}", "ERROR")

    def _restore_last_selections(self):
        try:
            settings = read_settings(SETTINGS_FILE)
        except (OSError, ValueError) as exc:
            self.logs.write(f"无法恢复上次的文件选择：{exc}", "WARNING")
            return
        paths = {}
        for prefix, title in (("output_directory", "输出目录"), ("order_pdf_directory", "订单 PDF 目录")):
            try:
                path = get_saved_path(settings, SETTINGS_FILE, f"{prefix}_file")
                if path is not None:
                    valid = path.is_dir()
                    if valid:
                        paths[prefix] = path
                    else:
                        self.logs.write(f"上次选择的{title}已不存在，请重新选择：{path}", "WARNING")
            except (OSError, ValueError) as exc:
                self.logs.write(f"无法恢复{title}：{exc}", "WARNING")
        # Restore only the directories used by the PDF workflow.
        if 'output_directory' in paths:
            self.excel_output_directory = paths['output_directory']
            self.output_directory_path.set(str(self.excel_output_directory))
            self.logs.write(f"已恢复输出目录：{self.excel_output_directory}")
        if 'order_pdf_directory' in paths:
            self.order_search_directory = paths['order_pdf_directory']
            self.order_directory_path.set(str(self.order_search_directory))
            self.logs.write(f"已恢复订单 PDF 目录：{self.order_search_directory}")
            self._start_directory_scan()

    def browse_output_directory(self):
        if self.packing_busy:
            self.logs.write("正在处理文件，请完成后再更改输出目录。", "WARNING")
            return
        directory = filedialog.askdirectory(parent=self, title="选择 Excel 输出目录",
            initialdir=self.excel_output_directory if self.excel_output_directory.is_dir() else PROJECT_DIR,
            mustexist=True)
        if not directory:
            self.logs.write("已取消选择输出目录。")
            return
        path = Path(directory)
        if not path.is_dir():
            self.logs.write(f"输出目录不存在或无法访问：{path}", "ERROR")
            return
        self.excel_output_directory = path
        self.output_directory_path.set(str(path))
        self._remember_selection('output_directory', path)
        self.logs.write(f"装箱单输出目录已设置：{path}")

    def generate_all_pdf_packing(self):
        if self.packing_busy:
            return
        if self.directory_scan_busy or self.product_loading or self.schedule_loading:
            self.logs.write("请等待目录搜索和资料表读取完成。", "WARNING")
            return
        if self.order_search_directory is None or self.product_catalog is None or self.order_schedule is None:
            self.logs.write("请先选择订单 PDF 目录，并加载 JP 产品资料表和订单排期表。", "WARNING")
            return
        self.packing_busy = True
        for control in (self.all_pdf_button, self.order_directory_button,
                        self.product_button, self.schedule_button):
            control.configure(state="disabled")
        self.all_pdf_button.configure(text="正在处理 PDF…")
        self.logs.write(f"正在扫描全部订单 PDF：{self.order_search_directory}")
        self._set_table_text("正在扫描订单 PDF 目录（包含子目录）…")
        self.table_info.configure(text="扫描 PDF 中…")
        Thread(target=run_packing_job, args=({}, self.product_catalog,
               self.order_schedule, PACKING_TEMPLATE, self.excel_output_directory, self.events),
               kwargs={"pdf_directory": self.order_search_directory}, daemon=True).start()

    def _set_table_text(self, text: str):
        self.table_text.configure(state="normal")
        self.table_text.delete("1.0", "end")
        self.table_text.insert("1.0", text)
        self.table_text.configure(state="disabled")

    def _display_pdf_files(self, files):
        self._set_table_text("订单 PDF 文件列表（包含子目录）\n\n" + ("\n".join(
            f"{index:03d} | {path.relative_to(self.order_search_directory)}"
            for index, path in enumerate(files, 1)) or "目录中没有 PDF 文件。"))
        self.table_info.configure(text=f"{len(files)} 个 PDF")

    def _restore_product_catalog(self):
        try:
            settings = read_settings(SETTINGS_FILE)
            path = get_product_path(settings, SETTINGS_FILE)
        except (ValueError, OSError) as exc:
            self.logs.write(f"资料表路径配置读取失败：{exc}，将尝试查找脚本同目录。", "WARNING")
            settings, path = {}, None
        try:
            local_files = find_local_product_files(PROJECT_DIR)
        except OSError as exc:
            self.logs.write(f"无法查找脚本同目录的资料表：{exc}", "WARNING")
            local_files = []
        if local_files:
            if len(local_files) == 1:
                local_path = local_files[0]
            elif settings.get("jp_product_confirmed") and path in local_files:
                local_path = path
            else:
                self.logs.write("脚本同目录中找到多份 JP 产品资料表，请选择需要使用的一份。", "WARNING")
                self.browse_product_catalog(initial_path=local_files[-1])
                return
            self.logs.write(f"已自动找到脚本同目录的 JP 产品资料表：{local_path.name}")
            self._load_product_catalog(local_path, remember=True)
            return
        if settings.get("jp_product_confirmed") and path and path.is_file():
            self._load_product_catalog(path, remember=False)
            return
        reason = "脚本同目录未找到 JP 产品资料表，请选择资料表。" if not settings.get("jp_product_confirmed") else "脚本同目录和记录路径均未找到资料表，请重新选择。"
        self.logs.write(reason, "WARNING")
        self.browse_product_catalog(initial_path=path or DEFAULT_JP_FILE)

    def browse_product_catalog(self, initial_path=None):
        if self.product_loading:
            return
        initial_path = initial_path or (self.product_catalog.source if self.product_catalog else DEFAULT_JP_FILE)
        filename = filedialog.askopenfilename(
            parent=self, title="选择 JP 产品净重毛重外箱尺寸表",
            initialdir=initial_path.parent if initial_path.parent.is_dir() else WORKSPACE_DIR,
            initialfile=initial_path.name, filetypes=[("Excel 产品资料表", "*.xlsx")])
        if filename:
            self._load_product_catalog(Path(filename), remember=True)
        else:
            self.logs.write("已取消选择 JP 产品资料表，可稍后点击“选择 JP 产品资料表”。")

    def _load_product_catalog(self, path, remember):
        self.product_loading = True
        self.product_button.configure(state="disabled", text="正在读取资料表…")
        self.logs.write(f"开始预读取 JP 产品资料表：{path}")
        Thread(target=load_product_catalog_job,
               args=(path, SETTINGS_FILE, remember, self.events), daemon=True).start()

    def _restore_order_schedule(self):
        settings, saved_path = {}, None
        try:
            settings = read_settings(SETTINGS_FILE)
            saved_path = get_saved_path(settings, SETTINGS_FILE, "order_schedule_file")
        except (ValueError, OSError) as exc:
            self.logs.write(f"订单排期表配置读取失败：{exc}", "WARNING")
        try:
            candidates = find_local_schedule_files(PROJECT_DIR)
        except OSError as exc:
            self.logs.write(f"订单排期表目录查找失败：{exc}", "WARNING")
            candidates = []
        if len(candidates) == 1:
            path = candidates[0]
        elif candidates and settings.get("order_schedule_confirmed") and saved_path in candidates:
            path = saved_path
        elif candidates:
            self.logs.write("找到多份订单排期表，请点击“选择订单排期表”指定需要使用的文件。", "WARNING")
            self.schedule_path.set("找到多份订单排期表，请选择")
            return
        elif settings.get("order_schedule_confirmed") and saved_path and saved_path.is_file():
            path = saved_path
        else:
            self.logs.write("未找到订单排期表，请点击“选择订单排期表”寻找文件。", "WARNING")
            self.schedule_path.set("未找到排期表，请点击右侧按钮")
            return
        self._load_order_schedule(path)

    def browse_order_schedule(self):
        if self.schedule_loading:
            return
        initial = self.order_schedule.source if self.order_schedule else None
        filename = filedialog.askopenfilename(
            parent=self, title="选择订单分类排期汇总 Excel",
            initialdir=initial.parent if initial and initial.parent.is_dir() else PROJECT_DIR,
            initialfile=initial.name if initial else "", filetypes=[("订单排期 Excel", "*.xlsx")])
        if filename:
            self._load_order_schedule(Path(filename))
        else:
            self.logs.write("已取消选择订单排期表。")

    def _load_order_schedule(self, path):
        self.schedule_loading = True
        self.schedule_button.configure(state="disabled", text="读取排期表…")
        self.logs.write(f"开始预读取订单排期表：{path}")
        Thread(target=load_order_schedule_job,
               args=(path, SETTINGS_FILE, True, self.events), daemon=True).start()


    def browse_order_directory(self):
        if self.directory_scan_busy:
            return
        directory = filedialog.askdirectory(
            parent=self, title="选择搜索订单的目录", mustexist=True,
            initialdir=self.order_search_directory or WORKSPACE_DIR)
        if not directory:
            self.logs.write("已取消选择订单搜索目录。")
            return
        path = Path(directory)
        if not path.is_dir():
            self.logs.write(f"订单搜索目录不存在：{path}", "ERROR")
            return
        self.order_search_directory = path
        self.order_directory_path.set(str(path))
        self._remember_selection('order_pdf_directory', path)
        self.logs.write(f"已设置订单搜索目录：{path}")
        self._start_directory_scan()

    def _start_directory_scan(self):
        if self.directory_scan_busy or self.order_search_directory is None:
            return
        self.order_pdf_files = []
        self.directory_scan_ready = False
        self.directory_scan_warnings = 0
        self.directory_scan_busy = True
        self.order_directory_button.configure(state="disabled", text="正在搜索 PDF…")
        self.logs.write("开始搜索订单 PDF（包含子目录）…")
        self._set_table_text("正在扫描订单 PDF 目录（包含子目录）…")
        self.table_info.configure(text="扫描 PDF 中…")
        Thread(target=scan_order_directory, args=(self.order_search_directory, self.events), daemon=True).start()

    def _poll_events(self):
        try:
            # Limit work per tick so large directories do not stall the GUI.
            for _ in range(100):
                event, payload = self.events.get_nowait()
                if event == "packing_pdf_files":
                    self._display_pdf_files(payload)
                elif event == "packing_log":
                    self.logs.write(*payload)
                elif event == "packing_choice":
                    item, candidates, packaging_hint, ready, answer = payload
                    try:
                        dialog = PackingChoiceDialog(self, item, candidates, packaging_hint)
                        self.wait_window(dialog)
                        answer.append(dialog.result)
                    finally:
                        ready.set()
                elif event == "packing_finished":
                    generated, failed, skipped = payload
                    self.logs.write(f"装箱单生成结束：成功 {len(generated)} 个，跳过 {len(skipped)} 个，失败 {len(failed)} 个。")
                    messagebox.showinfo("生成结束", f"成功：{len(generated)} 个\n跳过：{len(skipped)} 个\n失败：{len(failed)} 个（原因见日志）\n保存目录：{self.excel_output_directory}", parent=self)
                elif event == "packing_error":
                    self.logs.write(f"装箱单生成已停止：{payload}", "ERROR")
                elif event == "packing_done":
                    self.packing_busy = False
                    for control in (self.all_pdf_button, self.order_directory_button,
                                    self.product_button, self.schedule_button):
                        control.configure(state="normal")
                    self.all_pdf_button.configure(text="直接处理所有pdf订单")
                elif event == "schedule_loaded":
                    self.order_schedule = payload
                    self.schedule_path.set(str(payload.source))
                    self.logs.write(f"订单排期表已预读取：{len(payload.sheets)} 个工作表，{payload.nonempty_rows} 行非空内容（含表头）。")
                elif event == "schedule_remembered":
                    self.logs.write(f"订单排期表路径已保存到 {payload}。")
                elif event == "schedule_config_error":
                    self.logs.write(f"订单排期表已读取，但路径保存失败：{payload}", "ERROR")
                elif event == "schedule_error":
                    self.logs.write(f"订单排期表读取失败：{payload}，请点击按钮重新选择。", "ERROR")
                elif event == "schedule_done":
                    self.schedule_loading = False
                    self.schedule_button.configure(state="normal", text="选择订单排期表")
                elif event == "product_loaded":
                    self.product_catalog = payload
                    self.product_path.set(str(payload.source))
                    count = sum(len(rows) for rows in payload.items.values())
                    self.logs.write(f"JP 产品资料表已预读取：{len(payload.sheets)} 个工作表，{count} 条产品记录，{len(payload.items)} 个不同 ITEM NO。")
                elif event == "product_remembered":
                    self.logs.write(f"资料表路径已保存到 {payload}，下次启动自动读取。")
                elif event == "product_config_error":
                    self.logs.write(f"资料表已读取，但路径保存失败：{payload}", "ERROR")
                elif event == "product_error":
                    self.logs.write(f"JP 产品资料表读取失败：{payload}，请重新选择资料表。", "ERROR")
                elif event == "product_done":
                    self.product_loading = False
                    self.product_button.configure(state="normal", text="选择 JP 产品资料表")
                elif event == "pdf_scan_file":
                    number, path, relative_path = payload
                    self.order_pdf_files.append(path)
                    self.logs.write(f"PDF {number:03d} | 文件名：{path.name} | 相对路径：{relative_path}")
                elif event == "pdf_scan_warning":
                    self.logs.write(payload, "WARNING")
                elif event == "pdf_scan_summary":
                    count, warnings = payload
                    self.directory_scan_ready = True
                    self.directory_scan_warnings = warnings
                    if warnings:
                        self.logs.write(f"搜索结束：已找到 {count} 个 PDF 文件；{warnings} 处无法读取，统计可能不完整。", "WARNING")
                    else:
                        self.logs.write(f"搜索完成：共找到 {count} 个 PDF 文件（包含子目录）。")
                elif event == "pdf_scan_error":
                    self.logs.write(f"订单目录搜索失败：{payload}", "ERROR")
                    self._set_table_text(f"订单目录搜索失败：{payload}")
                    self.table_info.configure(text="扫描失败")
                elif event == "pdf_scan_done":
                    self.directory_scan_busy = False
                    self.order_directory_button.configure(state="normal", text="搜索订单目录")
                    if self.directory_scan_ready:
                        self._display_pdf_files(self.order_pdf_files)
        except Empty:
            pass
        self.after(100, self._poll_events)
