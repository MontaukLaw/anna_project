from datetime import datetime
import customtkinter as ctk
from ui import theme as t


def label(parent, text, size=13, color=t.TEXT, **kwargs):
    return ctk.CTkLabel(parent, text=text, font=(t.FONT, size), text_color=color, **kwargs)


def button(parent, text, command=None, **kwargs):
    return ctk.CTkButton(parent, text=text, command=command, height=36,
                         corner_radius=8, fg_color=t.BUTTON, hover_color=t.BUTTON_HOVER,
                         text_color="#182235", text_color_disabled="#7A899E",
                         font=(t.FONT, 13), **kwargs)


class Section(ctk.CTkFrame):
    def __init__(self, parent, number, title, subtitle):
        super().__init__(parent, fg_color=t.PANEL, corner_radius=12, border_width=1, border_color=t.BORDER)
        self.grid_columnconfigure(0, weight=1)
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=20, pady=(14, 8))
        label(header, number, color=t.ACCENT).pack(side="left", padx=(0, 12))
        label(header, title, size=16).pack(side="left")
        label(header, subtitle, size=12, color=t.MUTED).pack(side="right")


class LogPanel(Section):
    def __init__(self, parent):
        super().__init__(parent, "03", "运行日志", "文件选择 · 操作状态 · 异常信息")
        self.grid_rowconfigure(1, weight=1)
        self.textbox = ctk.CTkTextbox(self, fg_color=t.INSET, text_color=t.MUTED,
                                     font=(t.FONT, 14), height=90, state="disabled")
        self.textbox.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 12))

    def write(self, message, level="INFO"):
        self.textbox.configure(state="normal")
        self.textbox.insert("end", f"{datetime.now():%H:%M:%S}  [{level}]  {message}\n")
        self.textbox.see("end")
        self.textbox.configure(state="disabled")
