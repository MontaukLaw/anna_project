"""选择同一 ITEM 的具体装箱记录，必要时补充英文名称。"""
import tkinter as tk
from tkinter import ttk, messagebox
import customtkinter as ctk
from ui import theme as t
from ui.components import button, label
from services.packing_display_service import packing_with_cm, packaging_keywords
from services.packing_recommendation_service import recommend_candidate
from services.product_weight_service import product_weights


class PackingChoiceDialog(ctk.CTkToplevel):
    def __init__(self, parent, item, candidates, packaging_hint=""):
        super().__init__(parent)
        self.title(f"选择装箱资料 — {item.item_no}")
        self.geometry("1180x650")
        self.minsize(850, 550)
        self.configure(fg_color=t.BG)
        self.result = None
        self.candidates = candidates
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        label(self, f"ITEM {item.item_no}    PDF CASE PACK：{item.case_pack}    请选择正确的装箱方式", size=15).grid(row=0, column=0, padx=20, pady=16, sticky="w")
        hint_panel = ctk.CTkFrame(self, fg_color=t.PANEL)
        hint_panel.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 12))
        hint_panel.grid_columnconfigure((0, 1), weight=1, uniform="packing_hints")
        label(hint_panel, "订单排期表 · 包装要求", size=14, color=t.ACCENT).grid(row=0, column=0, sticky="w", padx=12, pady=(8, 4))
        label(hint_panel, "订单 PDF · 包装信息", size=14, color=t.ACCENT).grid(row=0, column=1, sticky="w", padx=12, pady=(8, 4))
        self.packaging_text = ctk.CTkTextbox(hint_panel, height=145, wrap="word",
            fg_color=t.INSET, text_color=t.TEXT, font=(t.FONT, 14))
        self.packaging_text.grid(row=1, column=0, sticky="ew", padx=(12, 6), pady=(0, 10))
        self.packaging_text.insert("1.0", packaging_hint or "未提供包装要求，请核对订单排期表。")
        self.packaging_text.configure(state="disabled")
        self.pdf_packaging_text = ctk.CTkTextbox(hint_panel, height=145, wrap="word",
            fg_color=t.INSET, text_color=t.TEXT, font=(t.FONT, 14))
        self.pdf_packaging_text.grid(row=1, column=1, sticky="ew", padx=(6, 12), pady=(0, 10))
        pdf_hint = f"Packaging：{item.packaging or '未读取到'}\n\nPacking：{packing_with_cm(item.packing) if item.packing else '未读取到'}"
        if item.shipping_remarks:
            pdf_hint += f"\n\nREMARKS：\n{item.shipping_remarks}"
        self.pdf_packaging_text.insert("1.0", pdf_hint)
        self.packaging_emphasis_font = ctk.CTkFont(family=t.FONT, size=18, weight="bold")
        # CTkTextbox's public tag API excludes fonts; use the underlying Tk text tag.
        self.pdf_packaging_text._textbox.tag_configure('packaging_emphasis', foreground='#FF4545',
            font=self.pdf_packaging_text._apply_font_scaling(self.packaging_emphasis_font))
        offset = len('Packaging：')
        for match in packaging_keywords(item.packaging):
            self.pdf_packaging_text.tag_add('packaging_emphasis',
                f'1.0 + {offset + match.start()} chars', f'1.0 + {offset + match.end()} chars')
        self.pdf_packaging_text.configure(state="disabled")
        style = ttk.Style(self)
        # The Windows native theme paints an unconfigurable white tree background.
        style.theme_use('clam')
        style.configure("Packing.Treeview", background=t.INSET, foreground=t.TEXT, fieldbackground=t.INSET,
            bordercolor=t.INSET, lightcolor=t.INSET, darkcolor=t.INSET,
            borderwidth=0, relief='flat', rowheight=40, font=(t.FONT, 13))
        style.configure('Packing.Treeview.Heading', background='#293548', foreground=t.TEXT,
            relief='flat', borderwidth=0, padding=(3, 9), font=(t.FONT, 13, 'bold'))
        style.map('Packing.Treeview', background=[('selected','#385879')], foreground=[('selected','#FFFFFF')])
        style.map('Packing.Treeview.Heading', background=[('active','#34465E')])
        frame = ctk.CTkFrame(self, fg_color=t.INSET, corner_radius=8, border_width=1, border_color=t.BORDER)
        frame.grid(row=2, column=0, sticky="nsew", padx=20)
        frame.grid_rowconfigure(1, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        label(frame, "装箱资料   ·   绿色：尺寸与包装均匹配   ·   蓝色：部分条件匹配   ·   请核对后选择", size=12, color=t.MUTED).grid(row=0, column=0, sticky='w', padx=12, pady=10)
        columns = ("row", "name", "pack", "size", "net", "gross", "note", "reason")
        self.tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="browse", style="Packing.Treeview")
        for key, title, width in zip(columns, ("来源行", "产品名称", "装箱数量", "长 × 宽 × 高 (CM)", "净重 KG", "毛重 KG", "备注", "推荐理由"), (75, 125, 85, 160, 80, 80, 190, 270)):
            self.tree.heading(key, text=title)
            self.tree.column(key, width=width, minwidth=20, stretch=False)
        self.tree.tag_configure('even', background=t.INSET, foreground=t.TEXT)
        self.tree.tag_configure('odd', background='#1B2533', foreground=t.TEXT)
        self.tree.tag_configure('strong', background='#183D35', foreground='#B9F3D8')
        self.tree.tag_configure('partial', background='#263953', foreground='#D1E5FF')
        for i, record in enumerate(candidates):
            v = record['values']
            net, gross, weight_source = product_weights(v)
            recommendation = recommend_candidate(item, record)
            tag = recommendation.level if recommendation.level != 'normal' else ('odd' if i % 2 else 'even')
            self.tree.insert("", "end", iid=str(i), values=(f"{record['sheet']}:{record['row']}", v.get("产品名称"), v.get("装箱数量"),
                f"{v.get('长')} × {v.get('宽')} × {v.get('高')}", net if net is not None else '缺失', gross if gross is not None else '缺失', v.get("备注"), recommendation.reason + f'；重量来源：{weight_source}'), tags=(tag,))
        self.tree.grid(row=1, column=0, sticky="nsew", padx=(8,0))
        self.row_detail = label(frame, "选择资料行可在这里查看完整备注与推荐理由", size=13, color=t.MUTED)
        self.row_detail.grid(row=2, column=0, sticky="ew", padx=8, pady=6)
        self.tree.bind('<Configure>', self._fit_columns)
        self.tree.bind('<<TreeviewSelect>>', self._show_row_detail)
        vertical = ctk.CTkScrollbar(frame, orientation="vertical", command=self.tree.yview, fg_color=t.INSET, button_color='#3A4A60', button_hover_color='#566F8E')
        vertical.grid(row=1, column=1, sticky="ns", padx=(0,4))
        self.tree.configure(yscrollcommand=vertical.set)
        self.english = ctk.StringVar(value=item.description)
        label(self, "玩具种类 · PDF DESCRIPTION 第一行", color=t.MUTED).grid(row=3, column=0, sticky="w", padx=20, pady=(12, 2))
        ctk.CTkEntry(self, textvariable=self.english, height=36, state="readonly",
            fg_color=t.INSET, text_color=t.TEXT, border_color=t.BORDER).grid(row=4, column=0, sticky="ew", padx=20)
        actions = ctk.CTkFrame(self, fg_color="transparent")
        actions.grid(row=5, column=0, sticky="e", padx=20, pady=14)
        button(actions, "跳过此订单", self.destroy).pack(side="left", padx=8)
        button(actions, "使用所选资料", self.confirm).pack(side="left")
        if len(candidates) == 1:
            self.tree.selection_set("0")
        self.transient(parent)
        self.after(100, self.grab_set)

    def _fit_columns(self, event):
        # Keep every column inside the available viewport, including on resize.
        weights = (6, 10, 8, 17, 7, 7, 20, 25)
        available = max(1, event.width - 4)
        widths = [int(available * weight / 100) for weight in weights]
        widths[-1] += available - sum(widths)
        for key, width in zip(self.tree['columns'], widths):
            self.tree.column(key, width=max(1, width))
        self.row_detail.configure(wraplength=max(100, available - 16))

    def _show_row_detail(self, _event=None):
        selected = self.tree.selection()
        if selected:
            values = self.tree.item(selected[0], 'values')
            self.row_detail.configure(text=f"备注：{values[6]}　｜　推荐理由：{values[7]}")

    def confirm(self):
        selected = self.tree.selection()
        if not selected or not self.english.get().strip():
            messagebox.showwarning("请选择资料", "请选择具体资料行，并确认 PDF DESCRIPTION 非空。", parent=self)
            return
        self.result = (self.candidates[int(selected[0])], self.english.get().strip())
        self.destroy()
