"""Persistent application-wide text size controls."""
import customtkinter as ctk

from services.settings_service import save_settings
from ui import theme as t
from ui.components import label


class FontControls(ctk.CTkFrame):
    def __init__(self, parent, settings_file, on_error):
        super().__init__(parent, fg_color=t.PANEL, corner_radius=10)
        self.settings_file = settings_file
        self.on_error = on_error
        self.caption = label(self, '', size=12, color=t.MUTED)
        self.caption.pack(side='left', padx=(12, 10))
        options = dict(width=34, height=30, corner_radius=7, font=t.font(17),
                       fg_color=t.INSET, hover_color=t.BORDER, text_color=t.TEXT)
        self.minus = ctk.CTkButton(self, text='−', command=lambda: self.change(-1), **options)
        self.minus.pack(side='left', padx=(0, 4), pady=5)
        self.plus = ctk.CTkButton(self, text='+', command=lambda: self.change(1), **options)
        self.plus.pack(side='left', padx=(0, 6), pady=5)
        self.refresh()

    def refresh(self):
        size = t.get_font_size()
        self.caption.configure(text=f'字体 {size}')
        self.minus.configure(state='normal' if size > t.MIN_FONT_SIZE else 'disabled')
        self.plus.configure(state='normal' if size < t.MAX_FONT_SIZE else 'disabled')

    def change(self, step):
        size = max(t.MIN_FONT_SIZE, min(t.MAX_FONT_SIZE, t.get_font_size() + step))
        t.set_font_size(size)
        self.refresh()
        try:
            save_settings(self.settings_file, {'ui_font_size': size})
        except (ValueError, OSError) as exc:
            self.on_error(f'字号已调整，但保存失败：{exc}', 'ERROR')
