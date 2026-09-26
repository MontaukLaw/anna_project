"""Ask only for a contract/product mapping that cannot be determined uniquely."""
from tkinter import messagebox
import customtkinter as ctk

from services.asn_generation_service import product_key, matching_products
from ui import theme as t
from ui.components import label, button


class AsnGenerationDialog(ctk.CTkToplevel):
    def __init__(self, parent, report, packing):
        super().__init__(parent)
        self.title('ASN · 选择合同与商品对应关系')
        self.geometry('880x520')
        self.result = None
        self.report = report
        self.configure(fg_color=t.BG)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        label(self, '无法唯一确定以下对应关系，请选择；每条海关装箱单明细仍独立输出。',
              color=t.ACCENT, wraplength=800).grid(row=0, column=0, padx=20, pady=14, sticky='w')
        form = ctk.CTkScrollableFrame(self, fg_color=t.PANEL)
        form.grid(row=1, column=0, sticky='nsew', padx=20)
        form.grid_columnconfigure(1, weight=1)
        self.contract = ctk.StringVar(value=next(iter(report.documents)) if len(report.documents) == 1 else '请选择')
        label(form, '合同号').grid(row=0, column=0, padx=12, pady=10, sticky='w')
        ctk.CTkOptionMenu(form, variable=self.contract, values=list(report.documents),
                         command=self.change_contract, font=t.font(13), dropdown_font=t.font(13)).grid(
                             row=0, column=1, sticky='ew', padx=12, pady=10)
        self.mapping_controls = {}
        for row in packing.rows:
            key = product_key(row)
            if key in self.mapping_controls:
                continue
            line = len(self.mapping_controls) + 1
            label(form, f'货号 {row.values.get("item")} / {row.values.get("chinese_name", "")} / USD {row.values.get("price")}',
                  wraplength=300).grid(row=line, column=0, sticky='w', padx=12, pady=10)
            var = ctk.StringVar(value='请选择')
            control = ctk.CTkOptionMenu(form, variable=var, values=['请选择'], font=t.font(13), dropdown_font=t.font(13))
            control.grid(row=line, column=1, sticky='ew', padx=12, pady=10)
            self.mapping_controls[key] = (row, var, control)
        self.change_contract()
        actions = ctk.CTkFrame(self, fg_color='transparent')
        actions.grid(row=2, column=0, sticky='e', padx=20, pady=14)
        button(actions, '取消', self.destroy, width=90).pack(side='left', padx=8)
        button(actions, '确定', self.confirm, width=120).pack(side='left')
        self.transient(parent.winfo_toplevel())
        self.after(100, self.grab_set)

    def change_contract(self, _value=None):
        docs = self.report.documents.get(self.contract.get())
        choices = [f'{i + 1}. {item.name} / {item.code} / USD {item.price}'
                   for i, item in enumerate(docs['装箱单'].items)] if docs else ['请选择']
        for row, var, control in self.mapping_controls.values():
            control.configure(values=choices)
            candidates = matching_products(docs['装箱单'], row) if docs else []
            var.set(choices[candidates[0]] if len(candidates) == 1 else '请选择')

    def confirm(self):
        mapping = {}
        for key, (_, var, _) in self.mapping_controls.items():
            try:
                mapping[key] = int(var.get().split('.', 1)[0]) - 1
            except ValueError:
                messagebox.showwarning('需要选择商品', '请明确选择每个货号对应的三单商品。', parent=self)
                return
        if self.contract.get() not in self.report.documents:
            messagebox.showwarning('需要选择合同', '请选择本份 ASN 使用的合同号。', parent=self)
            return
        self.result = {'contract': self.contract.get(), 'product_mapping': mapping}
        self.destroy()
