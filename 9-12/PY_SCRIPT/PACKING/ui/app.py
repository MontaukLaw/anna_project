from pathlib import Path
from queue import Empty, Queue
from threading import Thread
from tkinter import filedialog, messagebox
import customtkinter as ctk

from config import APP_NAME, IMAGE_FILE_TYPES, WORKSPACE_DIR, SETTINGS_FILE, DEFAULT_JP_FILE, PROJECT_DIR
from services.settings_service import read_settings, get_product_path, find_local_product_files
from services.product_catalog_service import load_product_catalog_job
from services.settings_service import find_local_schedule_files, get_saved_path, save_file_path
from services.schedule_service import load_order_schedule_job
from config import PACKING_OUTPUT_DIR, PACKING_TEMPLATE
from services.packing_job import run_packing_job
from services.packing_generation_service import output_paths
from ui.packing_choice_dialog import PackingChoiceDialog
from models.table import TableData
from services.excel_service import write_table
from services.text_table import format_table
from services.ocr_service import TableRecognitionService
from services.recognition_job import run_recognition
from services.order_directory_service import scan_order_directory
from services.order_matching_service import match_order_pdfs
from services.pdf_remarks_service import read_remarks_job
from ui import theme as t
from ui.components import LogPanel, Section, button, label


class PackingApp(ctk.CTk):
    def __init__(self):
        ctk.set_appearance_mode("dark")
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1180x880")
        self.minsize(960, 760)
        self.configure(fg_color=t.BG)
        self.selected_image: Path | None = None
        self.order_search_directory: Path | None = None
        self.order_pdf_files: list[Path] = []
        self.directory_scan_busy = False
        self.directory_scan_ready = False
        self.directory_scan_warnings = 0
        self.order_pdf_matches: dict[str, list[Path]] = {}
        self.table: TableData | None = None
        self.excel_output_directory = PACKING_OUTPUT_DIR
        self.busy = False
        self.events = Queue()
        self.ocr_service = TableRecognitionService()
        self.output_path: Path | None = None
        self.product_catalog = None
        self.product_loading = False
        self.order_schedule = None
        self.schedule_loading = False
        self.packing_busy = False
        self.remarks_generation = 0
        self.remarks_busy = False
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=3)
        self.grid_rowconfigure(3, weight=1)
        self._build_header()
        self._build_source()
        self._build_table()
        self.logs = LogPanel(self)
        self.logs.grid(row=3, column=0, sticky="nsew", padx=28, pady=(0, 12))
        label(self, "本地工作空间  /  图片 → 表格识别 → Excel", size=11, color=t.MUTED).grid(
            row=4, column=0, sticky="w", padx=30, pady=(0, 14))
        self.logs.write("系统已就绪。请选择箱单明细图片。")
        self.logs.write(f"选图后自动识别，Excel 输出目录：{self.excel_output_directory}")
        self.after(100, self._poll_events)
        self.after(250, self._restore_product_catalog)
        self.after(350, self._restore_order_schedule)
        self.after(450, self._restore_last_selections)

    def _build_header(self):
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=30, pady=(24, 18))
        label(header, APP_NAME, size=27).pack(anchor="w")
        label(header, "PACKING WORKSPACE    /    箱单数据工作台", size=12, color=t.MUTED).pack(anchor="w", pady=(5, 0))
        badge = label(header, "  自动识别 · 本地运行  ", size=12, color=t.ACCENT,
                      fg_color=t.PANEL, corner_radius=6)
        badge.place(relx=1, y=8, anchor="ne")

    def _build_source(self):
        panel = Section(self, "01", "选择源文件", "支持 PNG / JPG / BMP / WEBP / TIFF")
        panel.grid(row=1, column=0, sticky="ew", padx=28, pady=(0, 16))
        controls = ctk.CTkFrame(panel, fg_color="transparent")
        controls.grid(row=1, column=0, sticky="ew", padx=20, pady=(2, 10))
        for column in (0, 2, 4, 6, 8):
            controls.grid_columnconfigure(column, weight=1, uniform="source_paths")
        self.file_path = ctk.StringVar(value="尚未选择图片")
        self.path_entry = ctk.CTkEntry(controls, textvariable=self.file_path, state="readonly",
                                      height=38, fg_color=t.INSET, border_color=t.BORDER,
                                      text_color=t.MUTED, font=(t.FONT, 12))
        self.path_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.browse_button = button(controls, "选择图片", self.browse_image, width=88)
        self.browse_button.grid(row=0, column=1, padx=(0, 12))
        self.order_directory_path = ctk.StringVar(value="尚未选择订单搜索目录")
        self.order_directory_entry = ctk.CTkEntry(
            controls, textvariable=self.order_directory_path, state="readonly",
            height=38, fg_color=t.INSET, border_color=t.BORDER,
            text_color=t.MUTED, font=(t.FONT, 12))
        self.order_directory_entry.grid(row=0, column=2, sticky="ew", padx=(0, 8))
        self.order_directory_button = button(
            controls, "搜索订单目录", self.browse_order_directory, width=110)
        self.order_directory_button.grid(row=0, column=3, padx=(0, 12))
        self.product_path = ctk.StringVar(value="JP 产品资料表：尚未选择")
        self.product_entry = ctk.CTkEntry(
            controls, textvariable=self.product_path, state="readonly", height=38,
            fg_color=t.INSET, border_color=t.BORDER, text_color=t.MUTED, font=(t.FONT, 12))
        self.product_entry.grid(row=0, column=4, sticky="ew", padx=(0, 8))
        self.product_button = button(controls, "选择 JP 产品资料表", self.browse_product_catalog, width=144)
        self.product_button.grid(row=0, column=5, padx=(0, 12))
        self.schedule_path = ctk.StringVar(value="订单排期表：尚未选择")
        self.schedule_entry = ctk.CTkEntry(
            controls, textvariable=self.schedule_path, state="readonly", height=38,
            fg_color=t.INSET, border_color=t.BORDER, text_color=t.MUTED, font=(t.FONT, 12))
        self.schedule_entry.grid(row=0, column=6, sticky="ew", padx=(0, 8))
        self.schedule_button = button(controls, "选择订单排期表", self.browse_order_schedule, width=120)
        self.schedule_button.grid(row=0, column=7, padx=(0, 12))
        self.output_directory_path = ctk.StringVar(value=str(self.excel_output_directory))
        self.output_directory_entry = ctk.CTkEntry(
            controls, textvariable=self.output_directory_path, state="readonly", height=38,
            fg_color=t.INSET, border_color=t.BORDER, text_color=t.MUTED, font=(t.FONT, 12))
        self.output_directory_entry.grid(row=0, column=8, sticky="ew", padx=(0, 8))
        self.output_directory_button = button(controls, "输出目录", self.browse_output_directory, width=88)
        self.output_directory_button.grid(row=0, column=9)

    def _build_table(self):
        panel = Section(self, "02", "表格内容", "识别结果以纯文本表格显示")
        panel.grid(row=2, column=0, sticky="nsew", padx=28, pady=(0, 16))
        panel.grid_rowconfigure(2, weight=1)
        toolbar = ctk.CTkFrame(panel, fg_color="transparent")
        toolbar.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 10))
        self.recognize_button = button(toolbar, "重新识别", self.start_recognition, state="disabled", width=130)
        self.recognize_button.pack(side="left")
        self.export_button = button(toolbar, "另存为 Excel", self.export_excel, state="disabled", width=130)
        self.export_button.pack(side="left", padx=10)
        self.packing_button = button(toolbar, "生成装箱单", self.generate_packing, width=125)
        self.packing_button.pack(side="left", padx=(0, 10))
        self.table_info = label(toolbar, "0 行 / 0 列", size=12, color=t.MUTED)
        self.table_info.pack(side="right")
        self.table_text = ctk.CTkTextbox(panel, fg_color=t.INSET, text_color=t.TEXT,
                                        font=("Consolas", 13), wrap="none", height=160)
        self.table_text.grid(row=2, column=0, sticky="nsew", padx=20, pady=(0, 16))
        self._set_table_text("请选择一张表格图片。\n\n选择后自动识别，在此显示表格，并将 Excel 保存到所选输出目录。")

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
        for prefix, title in (("output_directory", "输出目录"), ("order_pdf_directory", "订单 PDF 目录"), ("selected_image", "图片")):
            try:
                path = get_saved_path(settings, SETTINGS_FILE, f"{prefix}_file")
                if path is not None:
                    valid = path.is_file() if prefix == 'selected_image' else path.is_dir()
                    if valid:
                        paths[prefix] = path
                    else:
                        self.logs.write(f"上次选择的{title}已不存在，请重新选择：{path}", "WARNING")
            except (OSError, ValueError) as exc:
                self.logs.write(f"无法恢复{title}：{exc}", "WARNING")
        # Set output and PDF paths before starting OCR so all jobs use restored selections.
        if 'output_directory' in paths:
            self.excel_output_directory = paths['output_directory']
            self.output_directory_path.set(str(self.excel_output_directory))
            self.logs.write(f"已恢复输出目录：{self.excel_output_directory}")
        if 'order_pdf_directory' in paths:
            self.order_search_directory = paths['order_pdf_directory']
            self.order_directory_path.set(str(self.order_search_directory))
            self.logs.write(f"已恢复订单 PDF 目录：{self.order_search_directory}")
            self._start_directory_scan()
        if 'selected_image' in paths:
            self.logs.write(f"已恢复上次的图片：{paths['selected_image']}")
            self.select_image(paths['selected_image'])

    def browse_output_directory(self):
        if self.busy or self.packing_busy:
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
        self.logs.write(f"Excel 输出目录已设置：{path}（识别结果和装箱单均保存到此目录）")

    def generate_packing(self):
        if self.packing_busy:
            return
        if self.busy or self.directory_scan_busy or self.product_loading or self.schedule_loading or self.remarks_busy:
            self.logs.write("请等待图片识别、目录搜索和 Excel 读取完成。", "WARNING")
            return
        if self.table is None or not self.order_pdf_matches or self.product_catalog is None or self.order_schedule is None:
            self.logs.write("请先识别图片、选择订单 PDF 目录，并加载 JP 产品资料表和订单排期表。", "WARNING")
            return
        try:
            paths = output_paths(self.order_pdf_matches, self.excel_output_directory)
        except ValueError as exc:
            self.logs.write(f"订单号错误：{exc}", "ERROR")
            return
        self.packing_busy = True
        for control in (self.packing_button, self.browse_button, self.order_directory_button,
                        self.product_button, self.schedule_button, self.recognize_button):
            control.configure(state="disabled")
        self.packing_button.configure(text="正在生成…")
        self.logs.write(f"开始生成装箱单，共 {len(paths)} 个订单。输出目录：{self.excel_output_directory}")
        Thread(target=run_packing_job, args=(dict(self.order_pdf_matches), self.product_catalog,
               self.order_schedule, PACKING_TEMPLATE, self.excel_output_directory, self.events), daemon=True).start()

    def _set_table_text(self, text: str):
        self.table_text.configure(state="normal")
        self.table_text.delete("1.0", "end")
        self.table_text.insert("1.0", text)
        self.table_text.configure(state="disabled")

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

    def browse_image(self):
        filename = filedialog.askopenfilename(parent=self, title="选择箱单明细图片",
                                              initialdir=self.selected_image.parent if self.selected_image else WORKSPACE_DIR,
                                              filetypes=IMAGE_FILE_TYPES)
        if filename:
            self.select_image(Path(filename))
        else:
            self.logs.write("已取消选择图片。")

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
        self.remarks_generation += 1
        self.remarks_busy = False
        if self.table is not None and 'REMARKS' in self.table.columns:
            index = self.table.columns.index('REMARKS')
            self.table = TableData(self.table.columns[:index], [row[:index] for row in self.table.rows])
            self._set_table_text(format_table(self.table))
        self.order_pdf_matches = {}
        self.directory_scan_ready = False
        self.directory_scan_warnings = 0
        self.directory_scan_busy = True
        self.order_directory_button.configure(state="disabled", text="正在搜索 PDF…")
        self.logs.write("开始搜索订单 PDF（包含子目录）…")
        Thread(target=scan_order_directory, args=(self.order_search_directory, self.events), daemon=True).start()

    def _match_order_pdfs(self):
        if self.table is None or not self.directory_scan_ready or self.directory_scan_busy:
            return
        self.order_pdf_matches = {}
        try:
            result = match_order_pdfs(self.table, self.order_pdf_files)
        except ValueError as exc:
            self.logs.write(str(exc), "WARNING")
            return
        self.order_pdf_matches = result.files_by_order
        self.logs.write(f"开始匹配订单 PDF | 订单列：{result.column} | 目录：{self.order_search_directory}")
        if result.blank_rows:
            self.logs.write(f"{result.blank_rows} 行订单号为空，已跳过。", "WARNING")
        if not self.order_pdf_matches:
            self.logs.write("订单号列没有有效订单号，未执行匹配。", "WARNING")
            return
        unique_files = set()
        matched_orders = 0
        for order, files in self.order_pdf_matches.items():
            if files:
                matched_orders += 1
                unique_files.update(files)
                self.logs.write(f"订单 {order}：找到 {len(files)} 个 PDF。")
                for path in files:
                    self.logs.write(f"  匹配文件：{path.relative_to(self.order_search_directory)}")
            else:
                self.logs.write(f"订单 {order}：找到 0 个 PDF，未匹配到文件。", "WARNING")
        total = len(self.order_pdf_matches)
        self.logs.write(f"订单 PDF 匹配完成：共 {total} 个不同订单号，{matched_orders} 个已匹配，"
                        f"{total - matched_orders} 个未匹配；共找到 {len(unique_files)} 个不同 PDF 文件。")
        if self.directory_scan_warnings:
            self.logs.write("部分目录无法读取，以上匹配结果可能不完整。", "WARNING")
        self.remarks_generation += 1
        self.remarks_busy = True
        self.logs.write("开始读取所有匹配订单 PDF 的 REMARKS…")
        Thread(target=read_remarks_job, args=(self.table, result.column, dict(self.order_pdf_matches),
               self.remarks_generation, self.events), daemon=True).start()

    def select_image(self, path: Path) -> bool:
        if self.busy:
            return False
        if not path.is_file():
            self.logs.write(f"图片不存在或无法访问：{path}", "ERROR")
            return False
        self.selected_image = path
        self.file_path.set(str(path))
        self._remember_selection('selected_image', path)
        self.table = None
        self.export_button.configure(state="disabled")
        self.table_info.configure(text="0 行 / 0 列")
        self.logs.write(f"已选择图片，开始读取识别：{path.name}")
        self.start_recognition()
        return True

    def start_recognition(self):
        if self.busy or self.selected_image is None:
            return
        self.busy = True
        self.remarks_generation += 1
        self.remarks_busy = False
        self.table = None
        self.order_pdf_matches = {}
        self.output_path = None
        self.browse_button.configure(state="disabled")
        self.recognize_button.configure(state="disabled", text="正在识别…")
        self.export_button.configure(state="disabled")
        self.table_info.configure(text="识别中…")
        self._set_table_text(f"正在识别：{self.selected_image.name}\n\n正在加载本地模型并识别表格，请稍候…\n完成后会自动显示结果并生成 Excel。")
        self.logs.write("开始自动识别表格…")
        Thread(target=run_recognition,
               args=(self.ocr_service, self.selected_image, self.excel_output_directory, self.events), daemon=True).start()

    def _poll_events(self):
        try:
            # Limit work per tick so large directories do not stall the GUI.
            for _ in range(100):
                event, payload = self.events.get_nowait()
                if event == "remarks_log":
                    generation, message, level = payload
                    if generation == self.remarks_generation:
                        self.logs.write(message, level)
                elif event == "remarks_ready":
                    generation, table = payload
                    if generation == self.remarks_generation:
                        self.table = table
                        self._set_table_text(format_table(table))
                        self.table_info.configure(text=f"{len(table.rows)} 行 / {len(table.columns)} 列 · REMARKS 已读取")
                        self.logs.write("包装备注已显示在 REMARKS 列。" if 'REMARKS' in table.columns else "没有匹配的包装备注，已隐藏 REMARKS 列。")
                elif event == "remarks_done":
                    if payload == self.remarks_generation:
                        self.remarks_busy = False
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
                    self.logs.write(f"装箱单生成结束：成功 {len(generated)} 个，已存在跳过 {len(skipped)} 个，失败 {len(failed)} 个。")
                    messagebox.showinfo("生成结束", f"成功：{len(generated)} 个\n已存在跳过：{len(skipped)} 个\n失败：{len(failed)} 个（原因见日志）\n保存目录：{self.excel_output_directory}", parent=self)
                elif event == "packing_error":
                    self.logs.write(f"装箱单生成已停止：{payload}", "ERROR")
                elif event == "packing_done":
                    self.packing_busy = False
                    for control in (self.packing_button, self.browse_button, self.order_directory_button,
                                    self.product_button, self.schedule_button, self.recognize_button):
                        control.configure(state="normal")
                    self.packing_button.configure(text="生成装箱单")
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
                elif event == "pdf_scan_done":
                    self.directory_scan_busy = False
                    self.order_directory_button.configure(state="normal", text="搜索订单目录")
                    self._match_order_pdfs()
                elif event == "table":
                    self.display_table(payload)
                    self.export_button.configure(state="disabled")
                elif event == "saved":
                    self.output_path = payload
                    self.logs.write(f"Excel 已自动生成：{payload}")
                    self.table_info.configure(text=f"{len(self.table.rows)} 行 / {len(self.table.columns)} 列 · Excel 已保存")
                elif event == "save_error":
                    self.logs.write(f"识别完成，但 Excel 保存失败：{payload}。可点击“另存为 Excel”重试。", "ERROR")
                    self.table_info.configure(text="识别完成 · 保存失败，可另存为")
                elif event == "error":
                    self.logs.write(f"识别失败：{payload}", "ERROR")
                    self._set_table_text(f"识别失败\n\n{payload}\n\n可重新选择图片或点击“重新识别”。")
                    self.table_info.configure(text="识别失败")
                elif event == "done":
                    self.busy = False
                    self.browse_button.configure(state="normal")
                    self.recognize_button.configure(state="normal", text="重新识别")
                    self.export_button.configure(state="normal" if self.table is not None else "disabled")
        except Empty:
            pass
        self.after(100, self._poll_events)

    def display_table(self, table: TableData):
        """在主线程显示后台 OCR 生成的表格。"""
        rendered = format_table(table)
        self.table = table
        self._set_table_text(rendered)
        self.table_info.configure(text=f"{len(table.rows)} 行 / {len(table.columns)} 列")
        self.export_button.configure(state="normal")
        self.logs.write(f"表格已更新：{len(table.rows)} 行，{len(table.columns)} 列。")
        if self.order_search_directory is None:
            self.logs.write("请选择订单 PDF 目录，选择后会自动按订单号匹配文件。")
        elif not self.directory_scan_busy:
            # Refresh the directory so newly added PDFs are included for each image.
            self._start_directory_scan()

    def export_excel(self):
        if self.table is None:
            return
        filename = filedialog.asksaveasfilename(parent=self, title="保存箱单明细",
                                               initialdir=self.excel_output_directory if self.excel_output_directory.exists() else PROJECT_DIR,
                                               initialfile=f"{self.selected_image.stem if self.selected_image else '箱单明细'}.xlsx",
                                               defaultextension=".xlsx", filetypes=[("Excel 工作簿", "*.xlsx")])
        if not filename:
            return
        try:
            destination = write_table(self.table, Path(filename))
        except Exception as exc:
            self.logs.write(f"Excel 导出失败：{exc}", "ERROR")
        else:
            self.logs.write(f"Excel 已保存：{destination}")
