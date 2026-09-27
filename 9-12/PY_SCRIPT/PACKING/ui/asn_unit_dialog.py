"""Resolve a failed legal-unit lookup on the Tk main thread."""
import customtkinter as ctk

from services.hs_units_service import SUPPORTED_UNITS, LegalUnitError, manual_legal_units
from ui import theme as t
from ui.components import button, label


class AsnUnitDialog(ctk.CTkToplevel):
    NO_SECOND = '无第二法定单位'

    def __init__(self, parent, code, error, details=''):
        super().__init__(parent)
        self.result = None
        self.title('ASN · 法定单位查询失败')
        self.geometry('740x560')
        self.minsize(620, 450)
        self.configure(fg_color=t.BG)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        label(self, f'商品编码 {code} 查询失败，如何继续？',
              font=t.font(18, 'bold'), color='#FFD18A', wraplength=660).grid(
                  row=0, column=0, sticky='w', padx=24, pady=(20, 12))
        form = ctk.CTkScrollableFrame(self, fg_color=t.PANEL)
        form.grid(row=1, column=0, sticky='nsew', padx=24)
        form.grid_columnconfigure(0, weight=1)
        description = label(form, f'{details}\n\n失败原因：{error}' if details else f'失败原因：{error}',
                            justify='left', anchor='w', wraplength=620, color=t.MUTED)
        description.grid(row=0, column=0, sticky='ew', padx=14, pady=14)
        form.bind('<Configure>', lambda e: description.configure(wraplength=max(200, e.width - 50)))
        label(form, '可以重试查询，或手动确认以下单位后继续。',
              wraplength=540, justify='left').grid(row=1, column=0, sticky='w', padx=14, pady=8)
        label(form, '手动选择仅用于本次生成，不会保存为下次的默认值。',
              color=t.MUTED, wraplength=540, justify='left').grid(row=2, column=0, sticky='w', padx=14, pady=8)
        self.first = ctk.StringVar(value='请选择法1单位')
        self.second = ctk.StringVar(value='请选择法2单位')
        units = ['个', '套', '件', '千克'] + sorted(SUPPORTED_UNITS - {'个', '套', '件', '千克'})
        for row, caption, variable, values in (
                (3, '法1单位', self.first, units),
                (4, '法2单位', self.second, [self.NO_SECOND] + units)):
            line = ctk.CTkFrame(form, fg_color='transparent')
            line.grid(row=row, column=0, sticky='ew', padx=14, pady=10)
            line.grid_columnconfigure(1, weight=1)
            label(line, caption).grid(row=0, column=0, padx=(0, 20))
            ctk.CTkOptionMenu(line, variable=variable, values=values,
                              font=t.font(14), dropdown_font=t.font(14)).grid(row=0, column=1, sticky='ew')
        self.error_label = label(form, '', color='#FF8585', wraplength=540, justify='left')
        self.error_label.grid(row=5, column=0, sticky='w', padx=14, pady=10)
        actions = ctk.CTkFrame(self, fg_color='transparent')
        actions.grid(row=2, column=0, sticky='ew', padx=24, pady=20)
        actions.grid_columnconfigure((0, 1, 2), weight=1)
        button(actions, '重试查询', self.retry, width=130).grid(row=0, column=0, sticky='ew', padx=(0, 10))
        button(actions, '手动确认并继续', self.confirm, width=170).grid(row=0, column=1, sticky='ew', padx=(0, 10))
        button(actions, '取消生成', self.destroy, width=130).grid(row=0, column=2, sticky='ew')
        self.protocol('WM_DELETE_WINDOW', self.destroy)
        self.bind('<Escape>', lambda _event: self.destroy())
        self.transient(parent.winfo_toplevel())
        self.after(100, self.grab_set)

    def retry(self):
        self.result = 'retry'
        self.destroy()

    def confirm(self):
        second = '' if self.second.get() == self.NO_SECOND else self.second.get()
        try:
            self.result = manual_legal_units(self.first.get(), second)
        except LegalUnitError as exc:
            self.error_label.configure(text=str(exc))
            return
        self.destroy()
