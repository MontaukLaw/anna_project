BG = "#11151D"
PANEL = "#1B2230"
INSET = "#141B26"
BORDER = "#303B4D"
TEXT = "#EDF2FA"
MUTED = "#9CAEC5"
BUTTON = "#E0E9F5"
BUTTON_HOVER = "#C1D2E8"
ACCENT = "#9FC5FC"
FONT = "Microsoft YaHei UI"

# All application fonts share one persisted base size. CTkFont callbacks update
# existing controls; controls created later automatically use the current size.
from weakref import ref
from tkinter import TclError
import customtkinter as ctk

DEFAULT_FONT_SIZE = 14
MIN_FONT_SIZE = 10
MAX_FONT_SIZE = 24
_font_size = DEFAULT_FONT_SIZE
_fonts = []


def valid_font_size(value):
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return DEFAULT_FONT_SIZE
    if not MIN_FONT_SIZE <= value <= MAX_FONT_SIZE:
        return DEFAULT_FONT_SIZE
    return round(value)


def font(size=13, weight='normal'):
    value = ctk.CTkFont(family=FONT, size=round(size * _font_size / DEFAULT_FONT_SIZE), weight=weight)
    _fonts.append((ref(value), size))
    return value


def set_font_size(size):
    global _font_size
    _font_size = valid_font_size(size)
    _fonts[:] = [(reference, base) for reference, base in _fonts if reference() is not None]
    for reference, base in list(_fonts):
        value = reference()
        if value is not None:
            try:
                value.configure(size=round(base * _font_size / DEFAULT_FONT_SIZE))
            except TclError:
                # A closed window may still have Python references during shutdown.
                _fonts.remove((reference, base))


def get_font_size():
    return _font_size
