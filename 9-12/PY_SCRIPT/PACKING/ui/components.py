from datetime import datetime
import customtkinter as ctk
from ui import theme as t


def label(parent, text, size=13, color=t.TEXT, **kwargs):
    return ctk.CTkLabel(parent, text=text, font=kwargs.pop('font', None) or t.font(size), text_color=color, **kwargs)


def button(parent, text, command=None, **kwargs):
    return ctk.CTkButton(parent, text=text, command=command, height=36,
                         corner_radius=8, fg_color=t.BUTTON, hover_color=t.BUTTON_HOVER,
                         text_color="#182235", text_color_disabled="#7A899E",
                         font=t.font(13), **kwargs)


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
    def __init__(self, parent, title='运行日志', subtitle='文件选择 · 操作状态 · 异常信息', show_metadata=True):
        super().__init__(parent, "03", title, subtitle)
        self.show_metadata = show_metadata
        self.grid_rowconfigure(1, weight=1)
        self.textbox = ctk.CTkTextbox(self, fg_color=t.INSET, text_color=t.MUTED,
                                     font=t.font(14), height=90, state="disabled")
        self.textbox.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 12))
        self.emphasis_font = t.font(15, 'bold')
        self.emphasis_font.add_size_configure_callback(self._update_tags)
        self._update_tags()

    def _update_tags(self):
        for name, color in (('ERROR', '#FF8585'), ('WARNING', '#FFD18A'),
                            ('SUCCESS', '#83DDB2'), ('TITLE', t.ACCENT)):
            # CTk's tag API excludes fonts; keep the underlying Tk tag in sync.
            self.textbox._textbox.tag_configure(name, foreground=color,
                font=self.textbox._apply_font_scaling(self.emphasis_font))

    def clear(self):
        self.textbox.configure(state='normal')
        self.textbox.delete('1.0', 'end')
        self.textbox.configure(state='disabled')

    def destroy(self):
        self.emphasis_font.remove_size_configure_callback(self._update_tags)
        super().destroy()

    def write(self, message, level="INFO"):
        self.textbox.configure(state="normal")
        prefix = f"{datetime.now():%H:%M:%S}  [{level}]  " if self.show_metadata else ''
        self.textbox.insert("end", f"{prefix}{message}\n", (level,))
        self.textbox.see("end")
        self.textbox.configure(state="disabled")
