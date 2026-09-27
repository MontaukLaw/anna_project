"""Order PDF selection with filename filtering independent of checked state."""
from pathlib import Path

import customtkinter as ctk

from ui import theme as t
from ui.components import button, label


class PdfSelection(ctk.CTkFrame):
    def __init__(self, parent, on_change):
        super().__init__(parent, fg_color=t.PANEL)
        self.on_change = on_change
        self.files = []
        self.entries = {}
        self.visible_files = []
        self.enabled = True
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        toolbar = ctk.CTkFrame(self, fg_color='transparent')
        toolbar.grid(row=0, column=0, padx=12, pady=(10, 4), sticky='ew')
        toolbar.grid_columnconfigure(1, weight=1)
        label(toolbar, '订单 PDF').grid(row=0, column=0, padx=(0, 12))
        self.filter_var = ctk.StringVar()
        self.filter_entry = ctk.CTkEntry(toolbar, textvariable=self.filter_var,
            placeholder_text='输入文件名或订单号片段，例如 167、759',
            height=36, fg_color=t.INSET, text_color=t.TEXT, font=t.font(13))
        self.filter_entry.grid(row=0, column=1, sticky='ew', padx=(0, 10))
        self.select_all_button = button(toolbar, '全选', lambda: self.select_visible(True), width=72)
        self.select_all_button.grid(row=0, column=2, padx=(0, 8))
        self.select_none_button = button(toolbar, '全不选', lambda: self.select_visible(False), width=80)
        self.select_none_button.grid(row=0, column=3)
        self.summary = label(self, '', color=t.ACCENT)
        self.summary.grid(row=1, column=0, padx=12, sticky='w')
        self.table = ctk.CTkScrollableFrame(self, fg_color=t.INSET, height=120)
        self.table.grid(row=2, column=0, padx=12, pady=5, sticky='nsew')
        self.table.grid_columnconfigure(1, weight=1)
        label(self.table, '选择', color=t.MUTED).grid(row=0, column=0, padx=8, pady=3)
        label(self.table, 'PDF 文件（含相对路径）', color=t.MUTED, anchor='w').grid(
            row=0, column=1, sticky='ew', padx=8)
        self.empty = label(self.table, '请选择订单 PDF 目录', color=t.MUTED)
        self.empty.grid(row=1, column=0, columnspan=2, pady=10)
        label(self, '默认不勾选；筛选保留勾选，全选 / 全不选仅作用于当前显示的文件。',
              size=12, color=t.MUTED).grid(row=3, column=0, padx=12, pady=(0, 8), sticky='w')
        self.filter_var.trace_add('write', lambda *_: self.apply_filter())
        self.table.bind('<Configure>', self.resize_labels, add='+')
        self.update_summary()

    def set_files(self, files, directory):
        for variable, checkbox, text in self.entries.values():
            checkbox.destroy()
            text.destroy()
        self.files = list(dict.fromkeys(Path(path) for path in files))
        self.entries = {}
        self.filter_var.set('')
        for path in self.files:
            variable = ctk.BooleanVar(value=False)
            checkbox = ctk.CTkCheckBox(self.table, text='', variable=variable,
                command=self.selection_changed, width=24, checkbox_width=20, checkbox_height=20)
            relative = path.relative_to(directory)
            text = label(self.table, str(relative), anchor='w', justify='left',
                         wraplength=max(220, self.table.winfo_width() - 85))
            self.entries[path] = (variable, checkbox, text)
        self.apply_filter()
        self.set_enabled(self.enabled)

    def resize_labels(self, event):
        for path in self.visible_files:
            self.entries[path][2].configure(wraplength=max(220, event.width - 85))

    def apply_filter(self):
        query = self.filter_var.get().strip().casefold()
        self.visible_files = [path for path in self.files if query in path.name.casefold()]
        for _, checkbox, text in self.entries.values():
            checkbox.grid_remove()
            text.grid_remove()
        for index, path in enumerate(self.visible_files, 1):
            if path not in self.entries:
                continue
            _, checkbox, text = self.entries[path]
            checkbox.grid(row=index, column=0, padx=8, pady=6, sticky='n')
            text.grid(row=index, column=1, padx=8, pady=6, sticky='ew')
        if self.visible_files:
            self.empty.grid_remove()
        else:
            self.empty.configure(text='没有匹配的 PDF' if self.files else '目录中暂无 PDF，请选择订单目录')
            self.empty.grid(row=1, column=0, columnspan=2, pady=10)
        self.table._parent_canvas.yview_moveto(0)
        self.update_summary()

    def selected_files(self):
        return tuple(path for path in self.files if path in self.entries and self.entries[path][0].get())

    def update_summary(self):
        selected = self.selected_files()
        visible = set(self.visible_files)
        hidden = sum(path not in visible for path in selected)
        self.summary.configure(text=f'共 {len(self.files)} 个 · 显示 {len(self.visible_files)} 个 · '
                                    f'已勾选 {len(selected)} 个（隐藏项已勾选 {hidden} 个）')

    def selection_changed(self):
        self.update_summary()
        self.on_change()

    def select_visible(self, selected):
        if not self.enabled:
            return
        for path in self.visible_files:
            self.entries[path][0].set(selected)
        self.selection_changed()

    def set_enabled(self, enabled):
        self.enabled = enabled
        state = 'normal' if enabled else 'disabled'
        for control in (self.filter_entry, self.select_all_button, self.select_none_button):
            control.configure(state=state)
        for _, checkbox, _ in self.entries.values():
            checkbox.configure(state=state)
