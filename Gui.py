import os
import math
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog
import importlib.util
import threading
import time
import requests
import subprocess
import sys
import socket
import pyAesCrypt
import psutil
import shutil
from functools import partial

from PIL import ImageTk, Image, ImageDraw, ImageChops, ImageOps
from tkinter import messagebox

try:
    import customtkinter as ctk
except ImportError:  # friendly message instead of a traceback
    _root = tk.Tk()
    _root.withdraw()
    messagebox.showerror(
        "Missing dependency",
        "AntiSociales DS4 Windows AI needs the 'customtkinter' package.\n\n"
        "Open a terminal and run:\n\n    py -m pip install customtkinter\n\n"
        "Then start AntiSociales DS4 Windows AI again.")
    sys.exit(1)

from Config import Config
from CaptureCard import CaptureCard
from CaptureDesktop import CaptureDesktop
from Controller import Controller
from ControllerMapping import Buttons

import win32com.client

currentDirectory = os.path.dirname(os.path.abspath(__file__))
ICON_DIR = os.path.join(currentDirectory, "res", "icons")

# get directories only inside relative folder to current script
def get_directories(relative_path):
    script_path = os.path.dirname(os.path.abspath(__file__))

    target_dir = os.path.join(script_path, relative_path)

    # Check if the directory exists
    if not os.path.isdir(target_dir):
        raise ValueError(f"The directory '{relative_path}' does not exist in the relative script path.")

    # Get the folder names inside the directory
    folder_names = [name for name in os.listdir(target_dir) if os.path.isdir(os.path.join(target_dir, name))]
    return folder_names

class MonitorBtnPositions:
    BTN_TOUCHPAD = [237, 120]
    BTN_GUIDE = [237, 147]
    BTN_SOUTH = [323, 160]
    BTN_EAST = [337, 145]
    BTN_WEST = [308, 145]
    BTN_NORTH = [323, 131]
    BTN_LEFT_SHOULDER = [166, 72]
    BTN_RIGHT_SHOULDER = [308, 72]
    BTN_BACK = [208, 147]
    BTN_OPTIONS = [267, 147]
    BTN_LEFT_THUMB = [151, 146]
    BTN_RIGHT_THUMB = [282, 207]
    BTN_DPAD_UP = [193, 193]
    BTN_DPAD_DOWN = [193, 222]
    BTN_DPAD_LEFT = [179, 207]
    BTN_DPAD_RIGHT = [208, 207]
    BTN_LEFT_TRIGGER = [171, 36]
    BTN_RIGHT_TRIGGER = [302, 36]

class Server:
    socket = None

    def __init__(self, ip, port, root):
        self.ip = ip
        self.port = port
        self.root = root

        self.socket = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        self.socket.bind((self.ip, self.port))
        self.conn = None

        self.thread = threading.Thread(target=self.main, daemon=True)
        self.running = False
        self.connected = False
        self.last_message = None
        self.seqNo = 0


    def main(self):
        addedLog = False

        while self.running:
            try:
                if (addedLog == False):
                    self.root.add_log("Waiting for client to connect...")
                    addedLog = True

                data, addr = self.socket.recvfrom(255)

                if (addr and addr != self.conn):
                    self.conn = addr
                    self.root.add_log("Client connected: " + str(addr), 'green')
                    self.connected = True
                    self.seqNo = 0

            except Exception as e:
                self.connected = False
                self.conn = None
                print(e)
            
            while self.connected:
                try:
                    data, addr = self.socket.recvfrom(255)

                    if (addr == self.conn and data != b''):
                        self.last_message = data.decode('utf-8')

                except Exception as e:
                    print(e)
                    self.connected = False
                    self.conn = None
                    pass

            time.sleep(0.01)

    def clear(self):
        self.last_message = None

    def start(self):
        self.running = True
        self.thread.start()

    def send(self, msg):
        if self.connected:
            try:
                msg = '<' + str(self.seqNo) + ':' + str(msg) + '>'
                self.socket.sendto(msg.encode(), self.conn)
                self.seqNo += 1
            except Exception as e:
                print(e)
                pass

    def stop(self):
        self.running = False

        time.sleep(0.1)
        self.socket.close()


class Encryptor:
    def __init__(self, file_path):
        if not os.path.exists(file_path):
            raise Exception('{} does not exists!'.format(file_path))
        
        self.file_path = file_path
        self.file_name = os.path.basename(self.file_path)
        self.file_dir = os.path.dirname(self.file_path)
        self.file_base_name, self.file_extension = os.path.splitext(self.file_name)

        if not self.file_extension == '.py':
            raise Exception('{} is not .py format!'.format(self.file_extension))
        
        # Creating .pvx file
        self.file_pyx_name = self.file_base_name + '.pyx'

        shutil.copy(self.file_path, self.file_pyx_name)
        self.setup_file()
        self.encrypt()

    def setup_file(self):
        with open('setup.py', '+w') as file:
            file.write("from distutils.core import setup\n"
                       "from Cython.Build import cythonize\n\n"
                       "setup(ext_modules=cythonize('{}'))".format(self.file_pyx_name))

    def encrypt(self):
        command = 'python setup.py build_ext --inplace'
        os.system(str(command))

        pyd_path = self.file_base_name + '.cp310-win_amd64.pyd'
        c_path = self.file_base_name + '.c'

        shutil.copy(pyd_path, self.file_dir)
        os.remove(pyd_path)
        os.remove(self.file_pyx_name)
        os.remove(c_path)
        os.remove('setup.py')

# =============================================================================
#  Interface layer  (Apple-inspired: rounded cards, soft greys, system blue)
# =============================================================================

THEMES = ["Ember", "Light", "Dark"]      # first entry is the default theme
THEME_MODE = {"Ember": "Dark", "Light": "Light", "Dark": "Dark"}


class UI:
    """Design tokens. Every colour is a (light, dark) pair. Light and Dark use
    the BASE palette; Ember is Dark with the EMBER overrides (warm black + orange)."""

    BASE = dict(
        BG=("#F2F2F7", "#141416"), CARD=("#FFFFFF", "#232326"), CARD_BORDER=("#E4E4EA", "#333337"),
        FIELD=("#F2F2F7", "#333337"), FIELD_HOVER=("#E7E7EE", "#3E3E43"), ROW_HOVER=("#F1F1F6", "#2D2D31"),
        SEP=("#E9E9EE", "#38383C"), TEXT=("#1C1C1E", "#F5F5F7"), TEXT2=("#8E8E93", "#9A9AA0"),
        BLUE=("#007AFF", "#0A84FF"), BLUE_HOVER=("#0068DB", "#3B99FF"),      # the accent colour
        RED=("#FF3B30", "#FF453A"), RED_HOVER=("#E12F25", "#FF6A61"),
        GREEN=("#34C759", "#30D158"), ORANGE=("#FF9500", "#FF9F0A"),
        ON_ACCENT=("#FFFFFF", "#FFFFFF"), ON_ACCENT_DISABLED=("#E8F1FF", "#B5D5FF"),
        SWITCH_ON=("#34C759", "#30D158"), SWITCH_OFF=("#D1D1D6", "#48484C"), SLIDER_TRACK=("#D9D9DE", "#48484C"),
        CHIP=("#E4E4EA", "#2E2E32"),
        SEG=("#E3E3E8", "#26262A"), SEG_SEL=("#FFFFFF", "#56565C"), SEG_SEL_HOVER=("#FFFFFF", "#5E5E64"),
        SEG_HOVER=("#D9D9DF", "#303035"),
        SCROLL=("#C7C7CC", "#56565B"), SCROLL_HOVER=("#AEAEB2", "#6A6A70"),
        HOT=("#FF3B30", "#FF453A"), STICK=("#007AFF", "#0A84FF"),
        DISABLED_TEXT=("#A5A5AA", "#6B6B70"),
    )
    EMBER = dict(
        BG="#130D09", CARD="#1F1610", CARD_BORDER="#35261A", FIELD="#2B1E14", FIELD_HOVER="#38291C",
        ROW_HOVER="#281C13", SEP="#35261A", TEXT="#FCF4EB", TEXT2="#AB978A",
        BLUE="#FF7A1A", BLUE_HOVER="#FF9642", RED="#F0524F", RED_HOVER="#FF6B68",
        ON_ACCENT="#1C0F06", ON_ACCENT_DISABLED="#7A4A26",
        SWITCH_ON="#FF7A1A", SWITCH_OFF="#41301F", SLIDER_TRACK="#43321F", CHIP="#2E2015",
        SEG="#241A12", SEG_SEL="#4A3524", SEG_SEL_HOVER="#563E2B", SEG_HOVER="#2F2217",
        SCROLL="#5A4433", SCROLL_HOVER="#735743", HOT="#FF7A1A", STICK="#FFD166", DISABLED_TEXT="#6E5A4B",
    )
    WHITE = "#FFFFFF"
    theme = "Ember"

    @classmethod
    def apply(cls, theme):
        """Point every colour token at the chosen theme."""
        cls.theme = theme
        for name, (light, dark) in cls.BASE.items():
            if theme == "Ember" and name in cls.EMBER:
                dark = cls.EMBER[name]
            setattr(cls, name, (light, dark))

    SANS = MONO = None
    f_title = f_head = f_body = f_bold = f_small = f_caption = f_button = f_mono = None

    @classmethod
    def init_fonts(cls, root):
        families = set(tkfont.families(root))

        def pick(options, fallback):
            for name in options:
                if name in families:
                    return name
            return fallback

        cls.SANS = pick(["SF Pro Display", "SF Pro Text", "Segoe UI Variable Display",
                         "Segoe UI Variable Text", "Segoe UI", "Helvetica Neue",
                         "Inter", "Cantarell", "Ubuntu", "DejaVu Sans"], "Arial")
        cls.MONO = pick(["SF Mono", "Cascadia Mono", "Cascadia Code", "Consolas",
                         "Menlo", "DejaVu Sans Mono", "Liberation Mono"], "Courier New")

        cls.f_title   = ctk.CTkFont(family=cls.SANS, size=24, weight="bold")
        cls.f_head    = ctk.CTkFont(family=cls.SANS, size=15, weight="bold")
        cls.f_body    = ctk.CTkFont(family=cls.SANS, size=13)
        cls.f_bold    = ctk.CTkFont(family=cls.SANS, size=13, weight="bold")
        cls.f_button  = ctk.CTkFont(family=cls.SANS, size=14, weight="bold")
        cls.f_small   = ctk.CTkFont(family=cls.SANS, size=12)
        cls.f_caption = ctk.CTkFont(family=cls.SANS, size=11, weight="bold")
        cls.f_mono    = ctk.CTkFont(family=cls.MONO, size=12)


UI.apply("Ember")


def mode_index():
    """0 = light, 1 = dark."""
    return 1 if ctk.get_appearance_mode() == "Dark" else 0


def resolve(pair):
    """Pick the colour for the current appearance from a (light, dark) pair."""
    return pair[mode_index()] if isinstance(pair, (tuple, list)) else pair


# ---------------------------------------------------------------- icon badges
def _white_glyph_from_ico(name):
    """The bundled .ico files are black silhouettes - turn them white."""
    im = Image.open(os.path.join(ICON_DIR, name))
    try:
        im.size = max(im.info["sizes"])
    except Exception:
        pass
    im = im.convert("RGBA")
    white = Image.new("RGBA", im.size, (255, 255, 255, 255))
    white.putalpha(im.split()[3])
    return white


def _draw_glyph(kind, size=256):
    im = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(im)
    w = (255, 255, 255, 255)
    if kind == "shield":
        d.polygon([(128, 18), (214, 48), (214, 122), (196, 172), (128, 238),
                   (60, 172), (42, 122), (42, 48)], fill=w)
        d.line([(92, 126), (120, 156), (170, 96)], fill=(0, 0, 0, 0), width=20, joint="curve")
    elif kind == "lock":
        d.rounded_rectangle((52, 112, 204, 226), radius=24, fill=w)
        d.arc((80, 30, 176, 134), 180, 360, fill=w, width=22)
        d.rectangle((80, 82, 102, 118), fill=w)
        d.rectangle((154, 82, 176, 118), fill=w)
        d.ellipse((116, 152, 140, 176), fill=(0, 0, 0, 0))
    elif kind == "eye":
        pts = [(128 + 108 * t / 20, 128 - 62 * (1 - (t / 20) ** 2)) for t in range(-20, 21)]
        pts += [(128 + 108 * t / 20, 128 + 62 * (1 - (t / 20) ** 2)) for t in range(20, -21, -1)]
        d.polygon(pts, fill=w)
        d.ellipse((92, 92, 164, 164), fill=(0, 0, 0, 0))
        d.ellipse((110, 110, 146, 146), fill=w)
        d.line([(58, 210), (198, 46)], fill=(0, 0, 0, 0), width=40)
        d.line([(58, 210), (198, 46)], fill=w, width=18)
    elif kind == "contrast":
        d.ellipse((34, 34, 222, 222), outline=w, width=20)
        d.pieslice((34, 34, 222, 222), 90, 270, fill=w)
    return im


def make_badge(color, glyph, size=30):
    """iOS-Settings style coloured rounded square with a white glyph."""
    scale = 4
    S = size * scale
    base = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    ImageDraw.Draw(base).rounded_rectangle((0, 0, S - 1, S - 1), radius=int(S * 0.27), fill=color)
    if glyph is not None:
        g = glyph.resize((int(S * 0.56), int(S * 0.56)), Image.LANCZOS)
        base.alpha_composite(g, ((S - g.width) // 2, (S - g.height) // 2))
    base = base.resize((size, size), Image.LANCZOS)
    return ctk.CTkImage(light_image=base, dark_image=base, size=(size, size))


def round_image(img, radius):
    """Round the corners of a PIL image."""
    img = img.convert("RGBA")
    scale = 3
    mask = Image.new("L", (img.width * scale, img.height * scale), 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, mask.width - 1, mask.height - 1), radius=radius * scale, fill=255)
    mask = mask.resize(img.size, Image.LANCZOS)
    img.putalpha(mask)
    return img


def cover_crop(img, width, height):
    """Scale + centre-crop to exactly width x height without distortion."""
    ratio = max(width / img.width, height / img.height)
    resized = img.resize((max(1, round(img.width * ratio)), max(1, round(img.height * ratio))), Image.LANCZOS)
    left = (resized.width - width) // 2
    top = (resized.height - height) // 2
    return resized.crop((left, top, left + width, top + height))


def set_window_icon(window, icon_name="app.ico"):
    """Windows resets custom-tkinter's icon after ~200ms, so set it twice."""
    path = os.path.join(ICON_DIR, icon_name)

    def apply():
        try:
            window.iconbitmap(path)
        except Exception:
            pass

    apply()
    try:
        window.after(250, apply)
    except Exception:
        pass


def place_near(window, parent, side="center", size=None):
    """Position a top-level relative to the main window (kept on-screen)."""
    window.update_idletasks()
    w, h = size if size else (window.winfo_reqwidth(), window.winfo_reqheight())
    sw, sh = window.winfo_screenwidth(), window.winfo_screenheight()
    px, py, pw, ph = parent.winfo_x(), parent.winfo_y(), parent.winfo_width(), parent.winfo_height()
    gap = 16

    if side == "right" and px + pw + gap + w <= sw:
        x, y = px + pw + gap, py
    elif side == "right" and px - gap - w >= 0:
        x, y = px - gap - w, py
    else:
        x, y = px + (pw - w) // 2, py + (ph - h) // 2

    x = max(0, min(x, sw - w))
    y = max(0, min(y, sh - h - 40))
    window.geometry(f"+{int(x)}+{int(y)}")


# ---------------------------------------------------------------- widgets
def make_card(master, **kw):
    kw.setdefault("corner_radius", 18)
    return ctk.CTkFrame(master, fg_color=UI.CARD, border_width=1, border_color=UI.CARD_BORDER, **kw)


def make_button(master, text, command=None, kind="primary", width=None, height=36, **kw):
    styles = {
        "primary":   dict(fg_color=UI.BLUE, hover_color=UI.BLUE_HOVER, text_color=UI.ON_ACCENT),
        "danger":    dict(fg_color=UI.RED, hover_color=UI.RED_HOVER, text_color=UI.WHITE),
        "secondary": dict(fg_color=UI.FIELD, hover_color=UI.FIELD_HOVER, text_color=UI.BLUE),
    }
    opts = dict(styles[kind])
    opts.setdefault("font", UI.f_bold)
    opts.update(kw)
    if width is not None:
        opts["width"] = width
    return ctk.CTkButton(master, text=text, command=command, height=height, corner_radius=12,
                         text_color_disabled=UI.DISABLED_TEXT, **opts)


def make_option(master, values, current=None, command=None, width=190):
    menu = ctk.CTkOptionMenu(
        master, values=[str(v) for v in values], command=command, width=width, height=32, corner_radius=10,
        dynamic_resizing=False, font=UI.f_body, dropdown_font=UI.f_body,
        fg_color=UI.FIELD, button_color=UI.FIELD, button_hover_color=UI.FIELD_HOVER,
        text_color=UI.TEXT, dropdown_fg_color=UI.CARD, dropdown_text_color=UI.TEXT,
        dropdown_hover_color=UI.FIELD_HOVER)
    if current is not None:
        menu.set(str(current))
    return menu


def make_entry(master, text="", width=110, **kw):
    entry = ctk.CTkEntry(master, width=width, height=32, corner_radius=10, border_width=1,
                         fg_color=UI.FIELD, border_color=UI.CARD_BORDER, text_color=UI.TEXT,
                         font=UI.f_body, **kw)
    if text != "":
        entry.insert(0, str(text))
    return entry


def make_switch(master, command=None, value=False):
    """iOS-style switch. The ring around the knob follows the track colour."""
    off = UI.SWITCH_OFF

    def sync():
        sw.configure(border_color=UI.SWITCH_ON if sw.get() else off)

    def on_toggle():
        sync()
        if command:
            command()

    sw = ctk.CTkSwitch(master, text="", width=52, height=28, switch_width=52, switch_height=28,
                       border_width=3, command=on_toggle, progress_color=UI.SWITCH_ON, button_color="#FFFFFF",
                       button_hover_color="#FFFFFF", fg_color=off, border_color=off)
    if value:
        sw.select()
    sync()
    return sw


def make_slider(master, from_, to, steps=None, command=None):
    return ctk.CTkSlider(master, from_=from_, to=to, number_of_steps=steps, command=command, height=18,
                         progress_color=UI.BLUE, button_color="#FFFFFF", button_hover_color="#F2F2F2",
                         fg_color=UI.SLIDER_TRACK, border_width=0)


class ListRow(ctk.CTkFrame):
    """A settings row: icon badge, title/subtitle, optional trailing widget
    or chevron. Clickable rows get a soft hover highlight."""

    def __init__(self, master, title, subtitle=None, icon=None, command=None, chevron=False, trailing=None):
        super().__init__(master, fg_color="transparent", corner_radius=12)
        self.command = command
        self.grid_columnconfigure(1, weight=1)

        if icon is not None:
            ctk.CTkLabel(self, text="", image=icon).grid(row=0, column=0, padx=(10, 12), pady=9)

        text = ctk.CTkFrame(self, fg_color="transparent")
        text.grid(row=0, column=1, sticky="w", pady=8, padx=(0 if icon else 12, 0))
        self.title_label = ctk.CTkLabel(text, text=title, font=UI.f_body, text_color=UI.TEXT, anchor="w", height=18)
        self.title_label.pack(anchor="w")
        self.subtitle_label = None
        if subtitle is not None:
            self.subtitle_label = ctk.CTkLabel(text, text=subtitle, font=UI.f_small, text_color=UI.TEXT2, anchor="w", height=16)
            self.subtitle_label.pack(anchor="w")

        self.trailing = None
        if chevron:
            self.trailing = ctk.CTkLabel(self, text="\u203A", font=ctk.CTkFont(family=UI.SANS, size=22),
                                         text_color=UI.TEXT2, width=20)
            self.trailing.grid(row=0, column=2, padx=(6, 12))
        elif trailing is not None:
            self.trailing = trailing(self)
            self.trailing.grid(row=0, column=2, padx=(8, 12))

        if command is not None:
            self._bind_tree(self)

    def set_subtitle(self, text):
        if self.subtitle_label is not None:
            self.subtitle_label.configure(text=text)

    def _bind_tree(self, widget):
        widget.bind("<Enter>", self._enter, add="+")
        widget.bind("<Leave>", self._leave, add="+")
        widget.bind("<Button-1>", self._click, add="+")
        for child in widget.winfo_children():
            self._bind_tree(child)

    def _enter(self, _event):
        self.configure(fg_color=UI.ROW_HOVER)

    def _leave(self, event):
        try:
            under = self.winfo_containing(event.x_root, event.y_root)
        except Exception:
            under = None
        if under is not None and str(under).startswith(str(self)):
            return
        self.configure(fg_color="transparent")

    def _click(self, _event):
        if self.command:
            self.command()


class Group(ctk.CTkFrame):
    """Rounded card holding several ListRow objects separated by hairlines."""

    def __init__(self, master):
        super().__init__(master, fg_color=UI.CARD, border_width=1, border_color=UI.CARD_BORDER, corner_radius=16)
        self.grid_columnconfigure(0, weight=1)
        self._n = 0

    def add(self, title, subtitle=None, icon=None, command=None, chevron=False, trailing=None):
        if self._n:
            ctk.CTkFrame(self, height=1, fg_color=UI.SEP, corner_radius=0).grid(
                row=self._n * 2 - 1, column=0, sticky="ew", padx=(58 if icon else 16, 16))
        row = ListRow(self, title, subtitle, icon, command, chevron, trailing)
        row.grid(row=self._n * 2, column=0, sticky="ew", padx=6, pady=3)
        self._n += 1
        return row


def section_label(master, text):
    return ctk.CTkLabel(master, text=text.upper(), font=UI.f_caption, text_color=UI.TEXT2, anchor="w")


# ---------------------------------------------------------------- dialogs
class Modal(ctk.CTkToplevel):
    """Blocking, rounded replacement for messagebox / simpledialog."""

    def __init__(self, gui, title, message, ok_text="OK", cancel_text=None, prompt=None,
                 secret=False, badge=None):
        super().__init__(gui.root, fg_color=UI.BG)
        self.withdraw()
        self.result = None
        self._prompt = None

        self.title(title)
        self.resizable(False, False)
        set_window_icon(self)
        self.transient(gui.root)
        self.protocol("WM_DELETE_WINDOW", self._cancel)

        card = make_card(self, corner_radius=20)
        card.pack(fill="both", expand=True, padx=16, pady=16)

        head = ctk.CTkFrame(card, fg_color="transparent")
        head.pack(fill="x", padx=20, pady=(20, 6))
        if badge is not None:
            ctk.CTkLabel(head, text="", image=badge).pack(side="left", padx=(0, 12), anchor="n")
        ctk.CTkLabel(head, text=title, font=UI.f_head, text_color=UI.TEXT, anchor="w",
                     justify="left").pack(side="left", anchor="n")

        ctk.CTkLabel(card, text=message, font=UI.f_body, text_color=UI.TEXT2, wraplength=340,
                     justify="left", anchor="w").pack(fill="x", padx=20, pady=(2, 12))

        if prompt is not None:
            self._prompt = make_entry(card, width=340, show="\u2022" if secret else "")
            self._prompt.pack(padx=20, pady=(0, 12), fill="x")

        buttons = ctk.CTkFrame(card, fg_color="transparent")
        buttons.pack(fill="x", padx=20, pady=(4, 20))
        make_button(buttons, ok_text, self._ok, "primary", width=110, height=36).pack(side="right")
        if cancel_text:
            make_button(buttons, cancel_text, self._cancel, "secondary", width=100, height=36).pack(side="right", padx=(0, 8))

        self.bind("<Return>", lambda _e: self._ok())
        self.bind("<Escape>", lambda _e: self._cancel())

        place_near(self, gui.root)
        self.deiconify()
        self.attributes("-topmost", True)
        self.after(120, self._grab)

    def _grab(self):
        try:
            self.grab_set()
            self.lift()
            (self._prompt or self).focus_force()
        except Exception:
            pass

    def _ok(self):
        self.result = self._prompt.get() if self._prompt is not None else True
        self.destroy()

    def _cancel(self):
        self.result = None if self._prompt is not None else False
        self.destroy()


class Sheet(ctk.CTkToplevel):
    """A settings window: title, a card of form rows, and a Done button."""

    def __init__(self, gui, key, title, icon_name):
        super().__init__(gui.root, fg_color=UI.BG)
        self.withdraw()
        self.gui = gui
        self.key = key
        self._rows = 0

        self.title(title)
        self.resizable(False, False)
        set_window_icon(self, icon_name)
        self.transient(gui.root)
        self.protocol("WM_DELETE_WINDOW", self.close)

        outer = ctk.CTkFrame(self, fg_color="transparent")
        outer.pack(fill="both", expand=True, padx=18, pady=18)
        ctk.CTkLabel(outer, text=title, font=UI.f_head, text_color=UI.TEXT, anchor="w").pack(fill="x", pady=(0, 10))

        self.form = make_card(outer, corner_radius=16)
        self.form.pack(fill="both", expand=True)
        self.form.grid_columnconfigure(0, weight=1)

        make_button(outer, "Done", self.close, "primary", height=38).pack(fill="x", pady=(14, 0))
        self.bind("<Escape>", lambda _e: self.close())

    def add_field(self, label, factory):
        """Label on the left, control on the right. Returns (row, label, control)."""
        row = ctk.CTkFrame(self.form, fg_color="transparent")
        row.grid(row=self._rows, column=0, sticky="ew", padx=16, pady=(12, 0))
        row.grid_columnconfigure(0, weight=1)
        lbl = ctk.CTkLabel(row, text=label, font=UI.f_body, text_color=UI.TEXT, anchor="w")
        lbl.grid(row=0, column=0, sticky="w", padx=(0, 18))
        control = factory(row)
        control.grid(row=0, column=1, sticky="e")
        self._rows += 1
        return row, lbl, control

    def add_slider(self, label, from_, to, resolution, initial, on_release):
        """Label + live value, slider underneath. Saves when released."""
        row = ctk.CTkFrame(self.form, fg_color="transparent")
        row.grid(row=self._rows, column=0, sticky="ew", padx=16, pady=(12, 0))
        row.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(row, text=label, font=UI.f_body, text_color=UI.TEXT, anchor="w").grid(row=0, column=0, sticky="w")
        value = ctk.CTkLabel(row, text="", font=UI.f_body, text_color=UI.TEXT2, anchor="e")
        value.grid(row=0, column=1, sticky="e")

        decimals = len(("%.4f" % resolution).rstrip("0").split(".")[1])
        steps = max(1, round((to - from_) / resolution))

        def fmt(v):
            return f"{v:.{decimals}f}"

        slider = make_slider(row, from_, to, steps, command=lambda v: value.configure(text=fmt(v)))
        slider.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        slider.set(min(max(initial, from_), to))
        value.configure(text=fmt(slider.get()))
        slider.bind("<ButtonRelease-1>", lambda _e: on_release(round(slider.get(), decimals)), add="+")
        self._rows += 1
        return row

    def show(self, side="center"):
        ctk.CTkFrame(self.form, height=14, fg_color="transparent").grid(row=self._rows, column=0)
        place_near(self, self.gui.root, side)
        self.deiconify()
        self.attributes("-topmost", True)
        self.lift()

    def close(self):
        self.gui._sheets.pop(self.key, None)
        try:
            self.destroy()
        finally:
            self.gui.refresh_summaries()


class Gui:
    version = "0.0.1"
    root = None
    config = Config("res/config.ini")
    output_log = None

    script_thread = None
    capture_thread = None

    capture_method = None
    script_running = False

    server_socket = None

    hiddenDevices = []

    controller_reader = {}
    controllerInstance = 0

    moduleInstance = None
    script_settings = None
    script_settings_window = None

    # controller artwork: crop the dead white margin, then scale
    MONITOR_CROP = (50, 10, 410, 330)
    MONITOR_SCALE = 0.8

    def hid_hider_command(self, command):
        if (self.config.get_setting('auto_hide', 'enabled', True)):
            if (os.path.exists(self.hidhide_cli_path)):
                self.kill_process_by_name("hidhideclient.exe")
                subprocess.Popen(command, shell=False, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    def __init__(self):
        appearance = str(self.config.get_setting('ui', 'appearance', THEMES[0]))
        if appearance not in THEMES:  # e.g. the old "System" option
            appearance = THEMES[0]
            self.config.set_setting('ui', 'appearance', appearance)
        self.theme = appearance
        UI.apply(appearance)
        ctk.set_appearance_mode(THEME_MODE[appearance])
        ctk.set_default_color_theme("blue")

        self.root = ctk.CTk(fg_color=UI.BG)
        UI.init_fonts(self.root)

        self._sheets = {}
        self._status_state = None
        self._last_status_time = 0.0
        self._summary_rows = {}
        self._main_widgets = []
        self._log_history = []
        self._status_text = "Waiting to run..."
        self._button_mode = "start"

        program_files = os.getenv('PROGRAMFILES')
        self.hidhide_cli_path = program_files + "\\Nefarius Software Solutions\\HidHide\\x64\\HidHideCLI.exe"
        self.hidhide_client_path = program_files + "\\Nefarius Software Solutions\\HidHide\\x64\\HidHideClient.exe"

        # check if file exists
        if (not os.path.exists(self.hidhide_cli_path) and not os.path.exists(self.hidhide_client_path)):
            messagebox.showerror("Error", "HidHideCLI.exe not found on this machine, please install HidHide.")
            script_path = os.path.dirname(os.path.abspath(__file__))
            subprocess.Popen([script_path + '/res/installers/HidHide_1.2.98_x64.exe'])
            os._exit(0)

        # run an .exe file with parameters
        drive = os.path.splitdrive(os.getcwd())[0]
        gpuEnvPath = drive.upper() + "\\tools\\miniconda3\\envs\\gpu\\python.exe"

        # Add python path to hidhider
        command = '"' + self.hidhide_cli_path + '" --app-reg "' + gpuEnvPath + '"'
        self.hid_hider_command(command)

        # enable hiding
        command = '"' + self.hidhide_cli_path + '" --cloak-on'
        self.hid_hider_command(command)

        self.root.title("AntiSociales DS4 Windows AI v" + str(self.version))
        set_window_icon(self.root)

        scale = self._window_scale()
        screen_h = self.root.winfo_screenheight() / scale
        height = int(min(730, max(620, screen_h - 110)))
        self.root.geometry(f"540x{height}")
        self.root.minsize(540, 620)

        # width is fixed, the log area grows with the window
        self.root.resizable(False, True)
        self.root.protocol("WM_DELETE_WINDOW", self.close_app)

        app_data = os.getenv('APPDATA')
        if app_data:
            os.makedirs(os.path.join(app_data, 'AntiSocialesDS4'), exist_ok=True)

        def report(exc, val, tb):
            import traceback
            text = "".join(traceback.format_exception(exc, val, tb))
            print(text)
            try:
                with open(os.path.join(currentDirectory, "antisociales_error.log"), "a") as f:
                    f.write(text + "\n")
            except Exception:
                pass
        self.root.report_callback_exception = report

        self.create_widgets()

    # ------------------------------------------------------------ helpers
    def _window_scale(self):
        try:
            return float(self.root._get_window_scaling())
        except Exception:
            return 1.0

    def peek(self, section, key, default=None):
        """Read a setting without writing the default back to config.ini."""
        return self.config.config.get(section, key, fallback=default)

    def notice(self, title, message, ok_text="OK", badge=None):
        """Blocking rounded dialog (replaces messagebox.showinfo)."""
        dialog = Modal(self, title, message, ok_text=ok_text, badge=badge or self.icons.get("input"))
        self.root.wait_window(dialog)
        return dialog.result

    def ask_password(self, title, message):
        dialog = Modal(self, title, message, ok_text="Continue", cancel_text="Cancel", prompt=True,
                       secret=True, badge=self.icons.get("lock"))
        self.root.wait_window(dialog)
        return dialog.result

    def open_sheet(self, key, title, icon_name):
        """Create a settings window, or focus the one that is already open."""
        existing = self._sheets.get(key)
        if existing is not None:
            try:
                if existing.winfo_exists():
                    existing.deiconify()
                    existing.lift()
                    existing.focus_force()
                    return None
            except Exception:
                pass
        sheet = Sheet(self, key, title, icon_name)
        self._sheets[key] = sheet
        return sheet

    def close_app(self):
        if (self.script_running):
            self.stop()

        self.root.destroy()

    def handle_server_connection(self):
        if (self.server_socket == None):
            return

        self.add_log("Waiting for client to connect...")

    # ------------------------------------------------------------ script loop
    def run_script(self):
        videoMode = self.config.get_setting("capture", "type", "None")

        try:
            if (self.moduleInstance == None or self.script_running == False):
                return 
            
            if (self.controller_reader[self.controllerInstance].state == 0):
                return
            
            start_time = time.time()

            # action decoder
            if (self.server_socket != None):
                if (self.server_socket.connected):
                    if (self.server_socket.last_message != None and self.server_socket.last_message != ""):
                        messages = self.server_socket.last_message.split(';')
                        messages = [msg for msg in messages if msg]
                        last_message = messages[-1]

                        split_message = last_message.split(' ')
                        action = split_message[0]
                        
                        # action contains ;
                        actionCheck = action.split(';')

                        if (action != '' and len(actionCheck) == 1):
                            self.moduleInstance.lastMessage = last_message
                            self.moduleInstance.lastMessageTime = time.time()
                            
                            actions = split_message[1:]
                            cleanActions = []

                            # correct type of actions
                            for act in actions:
                                split_action = act.split(':')
                                type = split_action[0]
                                value = split_action[1]

                                if (type == 'int'):
                                    value = int(value)

                                elif (type == 'float'):
                                    value = float(value)

                                elif (type == 'bool'):
                                    value = bool(value)

                                elif (type == 'str'):
                                    value = str(value)

                                cleanActions.append(value)

                            try:
                                func = getattr(self.moduleInstance, action)
                                func(*cleanActions)
                                #implodeActions = ", ".join(map(str, cleanActions))
                                #self.add_log("Client: " + action + " (" + str(implodeActions) + ")", "green")
                            except Exception as e:
                                print(e)
                                self.add_log("Failed to call action: " + action, "red")

                        self.server_socket.clear()

            # run script loop
            if(videoMode != "None"):
                try:
                    cleanFrame = self.capture_method.clean_frame.copy()
                except:
                    cleanFrame = None

                frame = self.moduleInstance.run(cleanFrame)
                self.capture_method.output_frame = frame
            else:
                self.moduleInstance.run(None)
            

            if (self.controller_reader[self.controllerInstance].emulated_controller != None):
                self.controller_reader[self.controllerInstance].emulated_controller.update()

            # print log messages
            if (len(self.moduleInstance.printQueue) > 0):
                for message in self.moduleInstance.printQueue:
                    self.add_log(str(message[0]), message[1], message[2])

                self.moduleInstance.printQueue = []

            # Calculate FPS
            elapsed_time = time.time() - start_time

            try:
                fps = 1.0 / elapsed_time
            except ZeroDivisionError:
                fps = 999

            if (fps > 999):
                fps = 999
            
            if(videoMode != "None"):
                self.set_status("Status: Running (" + str(round(fps)) + " Script FPS / " + str(round(self.capture_method.output_fps)) + " Capture FPS)")
            else:
                self.set_status("Status: Running (" + str(round(fps)) + " Script FPS)")

            if (self.moduleInstance.reloadScript):
                self.add_log("Reloading script...", "green")
                self.stop(True)
                self.start(None, True)

            if (self.script_running):
                self.root.after(1, self.run_script)

                if (self.controller_reader[self.controllerInstance].emulated_controller != None):
                    self.root.after(5, self.update_monitor)
            else:
                self.add_log("Script stopped.", "red")

        except Exception as e:
            # get line of code
            line = sys.exc_info()[-1].tb_lineno
            self.add_log("Script crashed: " + str(e) + ' Line:' + str(line), "red")
            self.stop(True)

    # ------------------------------------------------------------ controller monitor
    # (kind, Buttons attribute / MonitorBtnPositions attribute)
    MONITOR_BUTTONS = [
        ("special", "BTN_GUIDE"), ("special", "BTN_TOUCHPAD"),
        ("button", "BTN_SOUTH"), ("button", "BTN_EAST"), ("button", "BTN_WEST"), ("button", "BTN_NORTH"),
        ("button", "BTN_LEFT_SHOULDER"), ("button", "BTN_RIGHT_SHOULDER"),
        ("button", "BTN_BACK"), ("button", "BTN_OPTIONS"),
        ("button", "BTN_LEFT_THUMB"), ("button", "BTN_RIGHT_THUMB"),
        ("button", "BTN_DPAD_UP"), ("button", "BTN_DPAD_DOWN"),
        ("button", "BTN_DPAD_LEFT"), ("button", "BTN_DPAD_RIGHT"),
    ]

    def _monitor_colors(self):
        if UI.theme == "Ember":
            return dict(hot="#FF7A1A", glow="#5A2F12", stick="#FFD166", text_hot="#FF9A4D", text_stick="#FFDB8A")
        if UI.theme == "Light":
            return dict(hot="#FF3B30", glow="#FFD3D0", stick="#007AFF", text_hot="#D70015", text_stick="#0A60D0")
        return dict(hot="#FF453A", glow="#5B2B28", stick="#0A84FF", text_hot="#FF6961", text_stick="#64A8FF")

    def _build_monitor_image(self):
        """Controller artwork for the current appearance (recoloured for dark mode)."""
        base = Image.open(os.path.join(currentDirectory, "res", "images", "monitor.png")).convert("RGB")
        base = base.crop(self.MONITOR_CROP)
        size = (round(base.width * self._mon_scale), round(base.height * self._mon_scale))

        if mode_index() == 0:
            image = base.resize(size, Image.LANCZOS)
        else:
            darkest = ImageChops.darker(ImageChops.darker(base.getchannel("R"), base.getchannel("G")), base.getchannel("B"))
            alpha = ImageOps.invert(darkest).point(lambda v: min(255, int(v * 1.7)))
            card = resolve(UI.CARD).lstrip("#")
            bg_rgb = tuple(int(card[i:i + 2], 16) for i in (0, 2, 4))
            stroke = (245, 226, 208) if UI.theme == "Ember" else (222, 228, 240)
            composite = Image.composite(Image.new("RGB", base.size, stroke),
                                        Image.new("RGB", base.size, bg_rgb), alpha)
            image = composite.resize(size, Image.LANCZOS)

        self.device_monitor_template = ImageTk.PhotoImage(image)
        self.device_monitor_canvas.configure(bg=resolve(UI.CARD))

    def _mon_pos(self, pos):
        return ((pos[0] - self.MONITOR_CROP[0]) * self._mon_scale, (pos[1] - self.MONITOR_CROP[1]) * self._mon_scale)

    def update_monitor(self):
        canvas = self.device_monitor_canvas

        # reset image in canvas and add monitor image
        canvas.delete("all")
        canvas.create_image(0, 0, image=self.device_monitor_template, anchor="nw")

        if (self.moduleInstance == None or self.script_running == False):
            return 

        module = self.moduleInstance
        colors = self._monitor_colors()
        s = self._mon_scale
        r = 5 * s
        font = self._mon_font
        pressed = []

        def dot(pos):
            x, y = self._mon_pos(pos)
            canvas.create_oval(x - r - 3 * s, y - r - 3 * s, x + r + 3 * s, y + r + 3 * s, fill=colors["glow"], outline="")
            canvas.create_oval(x - r, y - r, x + r, y + r, fill=colors["hot"], outline="")

        for kind, name in self.MONITOR_BUTTONS:
            check = module.is_emulated_special_button_pressed if kind == "special" else module.is_emulated_button_pressed
            if (check(getattr(Buttons, name))):
                dot(getattr(MonitorBtnPositions, name))

        leftTrigger = module.get_emulated_left_trigger()
        if (leftTrigger > 0):
            dot(MonitorBtnPositions.BTN_LEFT_TRIGGER)
        x, y = self._mon_pos(MonitorBtnPositions.BTN_LEFT_TRIGGER)
        canvas.create_text(x + 20 * s, y, text=round(leftTrigger, 2), fill=colors["text_hot"], anchor="w", font=font)

        rightTrigger = module.get_emulated_right_trigger()
        if (rightTrigger > 0):
            dot(MonitorBtnPositions.BTN_RIGHT_TRIGGER)
        x, y = self._mon_pos(MonitorBtnPositions.BTN_RIGHT_TRIGGER)
        canvas.create_text(x + 20 * s, y, text=round(rightTrigger, 2), fill=colors["text_hot"], anchor="w", font=font)

        def stick(pos, sx, sy, text_dx, text_dy):
            x, y = self._mon_pos(pos)
            if (abs(sx) >= 0.01 or abs(sy) >= 0.01):
                ex, ey = x + sx * 20 * s, y + sy * 20 * s
                canvas.create_line(x, y, ex, ey, fill=colors["stick"], width=max(2, round(2 * s)), capstyle="round")
                canvas.create_oval(ex - 2.5 * s, ey - 2.5 * s, ex + 2.5 * s, ey + 2.5 * s, fill=colors["stick"], outline="")
            canvas.create_text(x + text_dx * s, y + text_dy * s, text="x: " + str(round(sx, 2)), fill=colors["text_stick"], anchor="w", font=font)
            canvas.create_text(x + text_dx * s, y + (text_dy + 15) * s, text="y: " + str(round(sy, 2)), fill=colors["text_stick"], anchor="w", font=font)

        stick(MonitorBtnPositions.BTN_LEFT_THUMB, module.get_emulated_left_stick_x(), module.get_emulated_left_stick_y(), -45, 40)
        stick(MonitorBtnPositions.BTN_RIGHT_THUMB, module.get_emulated_right_stick_x(), module.get_emulated_right_stick_y(), 35, 20)

        self.root.update_idletasks()

    # ------------------------------------------------------------ script settings window
    @staticmethod
    def _paint(name):
        """Map plain colour names from settings.ini onto the Apple palette."""
        named = {"red": UI.RED, "green": UI.GREEN, "blue": UI.BLUE, "orange": UI.ORANGE,
                 "white": UI.WHITE, "black": "#000000"}
        return named.get(str(name).strip().lower(), name)

    def load_settings_window(self):
        moduleName = self.config.get_setting("script", "name", "None")
        scriptDirectory = os.path.join(currentDirectory, "scripts", moduleName)

        if (self.moduleInstance == None):
            return

        # check if settings.ini file exists inside script folder
        settingsPath = os.path.join(scriptDirectory, "settings.ini")
        if (not os.path.isfile(settingsPath)):
            return

        self.script_settings = Config(settingsPath)
        self.moduleInstance._set_settings(self.script_settings)

        # floating panel that stays on top of the game and can only be closed by stopping the script
        window = ctk.CTkToplevel(self.root, fg_color=UI.BG)
        window.withdraw()
        self.script_settings_window = window

        window.title(moduleName + " Settings")
        set_window_icon(window)
        try:
            window.wm_attributes("-toolwindow", True)
        except Exception:
            pass
        window.attributes("-topmost", 1)
        window.protocol("WM_DELETE_WINDOW", lambda: None)

        scale = self._window_scale()
        width = 400
        height = int(min(660, max(360, self.root.winfo_screenheight() / scale - 120)))
        window.geometry(f"{width}x{height}")
        window.minsize(width, 300)
        window.resizable(False, True)

        scroll = ctk.CTkScrollableFrame(window, fg_color="transparent", corner_radius=0,
                                        scrollbar_button_color=UI.SCROLL,
                                        scrollbar_button_hover_color=UI.SCROLL_HOVER)
        scroll.pack(fill="both", expand=True, padx=(10, 4), pady=10)

        def card():
            frame = make_card(scroll, corner_radius=16)
            frame.pack(fill="x", padx=(2, 8), pady=5)
            frame.grid_columnconfigure(0, weight=1)
            return frame

        for section in self.script_settings.get_sections():
            field_info = self.script_settings.config[section]

            try:
                kind = field_info['type']

                if (kind == 'image'):
                    image = Image.open(os.path.join(scriptDirectory, field_info['path'].lstrip('/\\')))
                    banner = round_image(cover_crop(image.convert("RGBA"), 720, 440), 36)
                    photo = ctk.CTkImage(light_image=banner, dark_image=banner, size=(340, 208))
                    label = ctk.CTkLabel(scroll, text="", image=photo)
                    label.image = photo
                    label.pack(padx=(2, 8), pady=(2, 6))

                elif (kind == 'label'):
                    text = field_info['label']
                    color = field_info.get('color')
                    bg = field_info.get('bg')
                    try:
                        if bg:
                            fill, ink = self._paint(bg), (self._paint(color) if color else UI.WHITE)
                            if UI.theme == "Ember" and str(bg).strip().lower() == "red":
                                fill, ink = UI.BLUE, UI.ON_ACCENT   # red banners become the orange accent
                            ctk.CTkLabel(scroll, text=text, font=UI.f_bold, corner_radius=12, height=34,
                                         fg_color=fill, text_color=ink).pack(fill="x", padx=(2, 8), pady=(12, 4))
                        else:
                            ctk.CTkLabel(scroll, text=text, font=UI.f_bold, wraplength=330, justify="left", anchor="w",
                                         text_color=self._paint(color) if color else UI.TEXT).pack(fill="x", padx=(8, 8), pady=(10, 2))
                    except Exception:
                        ctk.CTkLabel(scroll, text=text, font=UI.f_bold, text_color=UI.TEXT).pack(fill="x", padx=8, pady=6)

                elif (kind == 'onoff'):
                    frame = card()
                    ctk.CTkLabel(frame, text=field_info['label'], font=UI.f_body, text_color=UI.TEXT, anchor="w",
                                 wraplength=250, justify="left").grid(row=0, column=0, sticky="w", padx=16, pady=14)

                    current = self.script_settings.get_setting(section, 'value', field_info['default'])
                    enabled = str(current).strip().lower() == "true"

                    switch = make_switch(frame, value=enabled)
                    switch.configure(command=partial(self._on_script_switch, switch, section))
                    switch.grid(row=0, column=1, padx=16)

                elif (kind == 'slider'):
                    self._build_script_slider(card(), section, field_info)

                elif (kind == 'dropdown' or kind == 'filelister'):
                    frame = card()
                    ctk.CTkLabel(frame, text=field_info['label'], font=UI.f_body, text_color=UI.TEXT, anchor="w").grid(
                        row=0, column=0, sticky="w", padx=16, pady=(14, 6))

                    if (kind == 'dropdown'):
                        options = [o for o in field_info['options'].split(', ') if o != '']
                    else:
                        folderPath = field_info['path'].replace('\\', '/').strip('/').strip('"').strip("'")
                        fullpath = os.path.join(scriptDirectory, *folderPath.split('/'))
                        extensions = tuple(field_info['ext'].split(', '))

                        try:
                            options = [f for f in os.listdir(fullpath) if f.endswith(extensions)]
                        except OSError:
                            options = []

                    menu = make_option(frame, options or [" "], None, width=300)
                    menu.configure(command=partial(self.update_script_setting, section))
                    menu.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 14))
                    menu.set(str(self.script_settings.get_setting(section, 'value', field_info['default'])))

            except Exception as e:
                print("Could not build setting '" + str(section) + "': " + str(e))

        place_near(window, self.root, "right", size=(width * scale, height * scale))
        window.deiconify()
        window.attributes("-topmost", 1)

    def _build_script_slider(self, frame, section, field_info):
        low, high = float(field_info['min']), float(field_info['max'])
        step = float(field_info['step']) if field_info.get('step') else 1.0
        decimals = len(("%.6f" % step).rstrip("0").split(".")[1])

        def fmt(value):
            return f"{value:.{decimals}f}"

        def store(value):
            # whole-number sliders are saved as ints, like the old tk.Scale did
            return int(round(value)) if decimals == 0 else round(value, decimals)

        try:
            initial = float(self.script_settings.get_setting(section, 'value', field_info['default']))
        except (TypeError, ValueError):
            initial = float(field_info['default'])
        initial = min(max(initial, low), high)

        ctk.CTkLabel(frame, text=field_info['label'], font=UI.f_body, text_color=UI.TEXT, anchor="w").grid(
            row=0, column=0, sticky="w", padx=(16, 8), pady=(14, 0))

        entry = make_entry(frame, fmt(initial), width=76, justify="center")
        entry.grid(row=0, column=1, padx=(0, 16), pady=(12, 0), sticky="e")

        slider = make_slider(frame, low, high, max(1, round((high - low) / step)),
                             command=lambda v: (entry.delete(0, tk.END), entry.insert(0, fmt(v))))
        slider.grid(row=1, column=0, columnspan=2, sticky="ew", padx=16, pady=(10, 16))
        slider.set(initial)

        def released(_event):
            self.update_script_setting(section, store(slider.get()))
            entry.delete(0, tk.END)
            entry.insert(0, fmt(slider.get()))

        def typed(_event):
            try:
                val = float(entry.get())
            except ValueError:
                val = 0.0
            val = min(max(val, low), high)
            slider.set(val)
            self.update_script_setting(section, store(val))

        slider.bind("<ButtonRelease-1>", released, add="+")
        entry.bind("<KeyRelease>", typed)

    def _on_script_switch(self, switch, section):
        self.update_script_setting(section, bool(switch.get()))

    def update_script_setting(self, section, value, key='value'):
        self.script_settings.set_setting(section, key, value)

        if (self.moduleInstance != None):
            self.moduleInstance._set_settings(self.script_settings)

    # ------------------------------------------------------------ start / stop
    def _start_button_mode(self, mode):
        self._button_mode = mode
        if mode == "start":
            self.start_button.configure(text="Start", state="normal", fg_color=UI.BLUE, hover_color=UI.BLUE_HOVER,
                                        text_color=UI.ON_ACCENT)
        elif mode == "stop":
            self.start_button.configure(text="Stop", state="normal", fg_color=UI.RED, hover_color=UI.RED_HOVER,
                                        text_color=UI.WHITE)
        else:  # busy
            self.start_button.configure(text="Stopping\u2026", state="disabled")

    def start(self, event=None, force=False):
        self.start_button.configure(state="disabled")

        if (force == False):
            if (self.start_button.cget("text") == "Stop"):
                self.stop()        
                return
        
        videoMode = self.config.get_setting("capture", "type", "None")
        controllerType = self.config.get_setting("input", "device_type", "Xbox Controller")

        strComputer = "."
        objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
        objSWbemServices = objWMIService.ConnectServer(strComputer, "root\\cimv2")
        colItems = objSWbemServices.ExecQuery("SELECT * FROM Win32_PnPEntity")
        self.hiddenDevices = []

        currentDirectory = os.path.dirname(os.path.abspath(__file__))

        if (self.config.get_setting('auto_hide', 'enabled', True)):
            controllerNames = [
                "Xbox Controller",
                "Xbox One Controller",
                "Xbox Wireless Controller",
                "Xbox Gaming Device",
                "Wireless Controller",
                "Game Controller",
                "Xbox",
                "Duelsense"
            ]

            if (os.path.exists(currentDirectory + "/res/devices.txt")):
                self.add_log("Resetting previous hide states...")

                with open(currentDirectory + "/res/devices.txt", "r") as file:
                    for line in file:
                        command = '"' + self.hidhide_cli_path + '" --dev-unhide "' + line.strip() + '"'
                        self.hid_hider_command(command)

            for objItem in colItems:
                for controllerName in controllerNames:
                    if objItem.Name and controllerName.lower() in objItem.Name.lower():
                        command = '"' + self.hidhide_cli_path + '" --dev-hide "' + str(objItem.DeviceID) + '"'
                        self.hid_hider_command(command)
                        self.hiddenDevices.append(objItem.DeviceID)
                        self.add_log("Hiding device: " + str(objItem.DeviceID))

            with open(currentDirectory + "/res/devices.txt", "w") as file:
                for device in self.hiddenDevices:
                    file.write(device + "\n")

        if (controllerType == 'Xbox Controller'):
            # Unplug 
            self.notice('Unplug your controller', 'Please unplug your controller from the PC and press OK.')

        # Start emulation
        try:
            self.controller_reader[self.controllerInstance] = Controller()
            self.controller_reader[self.controllerInstance].emulate()
        except:
           print("Error: Failed to start controller emulation.")
           self.start_button.configure(state="normal")
           return 
        
        if (controllerType == 'Xbox Controller'):
            # Wait for device to be replugged in
            self.notice('Plug it back in', 'Please plug your controller back into the PC, wait a few seconds and then press OK.')
            
        self.controller_reader[self.controllerInstance].waitPeriod = False

        # Start the capture card
        try: 
            if (videoMode == "Capture Card"):
                self.capture_method = CaptureCard()
                capture_thread = threading.Thread(target=self.capture_method.start, daemon=True)
                capture_thread.start()
            elif(videoMode == "Desktop"):
                self.capture_method = CaptureDesktop()
                capture_thread = threading.Thread(target=self.capture_method.start, daemon=True)
                capture_thread.start()
        except:
            print("Error: Failed to start capture card.")
            self.start_button.configure(state="normal")
            return 
        
        # Start server if enabled
        if (self.config.get_setting('server', 'enabled', 'Disabled') == 'Enabled'):
            if (self.server_socket == None):
                #try:
                ip_address = self.config.get_setting('server', 'ip_address', 'localhost')
                port = self.config.get_setting('server', 'port', '65432')

                self.server_socket = Server(ip_address, int(port), self)
                self.server_socket.start()

                self.add_log("Server started on " + ip_address + ":" + str(port), "green")
                # except Exception as e:
                #     print(e)
                #     self.add_log("Failed to start server, please check your settings.", "red")

        # Start the script
        try:
            moduleName = self.config.get_setting("script", "name", "None")

            if (moduleName != "None"):
                modulePath = os.path.join(os.path.dirname(os.path.abspath(__file__)), "scripts", moduleName, moduleName + ".py")

                if (not os.path.exists(modulePath)):
                    modulePath = os.path.join(os.path.dirname(os.path.abspath(__file__)), "scripts", moduleName, moduleName + ".cp310-win_amd64.pyd")
                    print("Loading compiled module: " + str(modulePath))

                spec = importlib.util.spec_from_file_location(moduleName, modulePath)
                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)

                scriptClass = getattr(module, 'Script')
                self.moduleInstance = scriptClass(self.controller_reader[self.controllerInstance].emulated_controller, self.controller_reader[self.controllerInstance].actual_report)
                self.moduleInstance.server = self.server_socket

                self.script_settings_window = None
                self.load_settings_window()
            else:
                modulePath = os.path.join(os.path.dirname(os.path.abspath(__file__)), "NoneScript.py")

                spec = importlib.util.spec_from_file_location(moduleName, modulePath)
                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)

                scriptClass = getattr(module, 'Script')
                self.moduleInstance = scriptClass(self.controller_reader[self.controllerInstance].emulated_controller, self.controller_reader[self.controllerInstance].actual_report)
                self.moduleInstance.server = self.server_socket

            self.script_running = True
            self.run_script()

            self.set_status("Status: Running")
            self._start_button_mode("stop")
            self.update_monitor()

            return True
        except Exception as e:
            print(e)
            self.stop(True)

    def kill_process_by_name(self, process_name):
        for proc in psutil.process_iter(attrs=['pid', 'name']):
            if proc.info['name'].lower() == process_name.lower():
                try:
                    psutil.Process(proc.info['pid']).terminate()
                    self.add_log("Closing process: " + str(proc.info['name']), "green")
                except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                    pass

    def stop(self, crashed = False):
        self.script_running = False

        if (crashed == False):
            self.run_script()

        self.add_log("Stopping script...", "red")

        self.set_status("Status: Stopping...")
        self._start_button_mode("busy")

        self.root.update_idletasks()

        if (len(self.hiddenDevices) > 0):
            for device in self.hiddenDevices:
                command = '"' + self.hidhide_cli_path + '" --dev-unhide "' + str(device) + '"'
                self.hid_hider_command(command)
                self.add_log("Unhiding device: " + str(device))
                self.root.update_idletasks()
        
        try:
            if (self.script_settings_window != None):
                self.script_settings_window.destroy()
        except:
            pass

        try:
            if (self.moduleInstance != None):
                self.moduleInstance.__del__()
                time.sleep(1)
                self.moduleInstance = None
        except:
            pass

        try:
            if (self.capture_method != None):
                self.capture_method.stop()
        except:
            pass

        try:
            if (self.controller_reader[self.controllerInstance] != None):
                self.controller_reader[self.controllerInstance].stop()
                self.controllerInstance += 1
        except:
            pass

        if (self.server_socket != None):
            self.add_log("Stopping server...", "red")

            try:
                self.server_socket.stop()
                self.server_socket = None
            except:
                pass

        self.set_status("Status: Stopped")
        self._start_button_mode("start")
        self.update_monitor()

    # ------------------------------------------------------------ status + log
    def _status_looks(self):
        return {
            "idle":     ("Idle",     UI.TEXT2),
            "running":  ("Running",  UI.GREEN),
            "stopping": ("Stopping", UI.ORANGE),
            "stopped":  ("Stopped",  UI.TEXT2),
        }

    def set_status(self, text):
        # the FPS readout changes every frame - don't repaint faster than 4x per second
        now = time.time()
        self._status_text = text
        if text.startswith("Status: Running (") and now - self._last_status_time < 0.25:
            return
        self._last_status_time = now

        self.status_bar_left_label.configure(text=text)

        lowered = text.lower()
        state = "running" if "running" in lowered else "stopping" if "stopping" in lowered else \
                "stopped" if "stopped" in lowered else "idle"
        if state != self._status_state:
            self._status_state = state
            word, color = self._status_looks()[state]
            self.status_pill.configure(text="\u25CF  " + word, text_color=color)

    def _apply_log_colors(self):
        log = self.log_text
        palette = {
            "Light": {"red": "#D70015", "green": "#248A3D", "blue": "#0A60D0"},
            "Dark":  {"red": "#FF6961", "green": "#32D74B", "blue": "#64A8FF"},
            "Ember": {"red": "#FF6B6B", "green": "#8FD14F", "blue": "#FFB454"},
        }[UI.theme]
        text, dim = resolve(UI.TEXT), resolve(UI.TEXT2)

        log.tag_configure("black", foreground=text)
        log.tag_configure("bold-black", foreground=text, font=self._log_bold)
        log.tag_configure("time", foreground=dim)
        for name, color in palette.items():
            log.tag_configure(name, foreground=color)
            log.tag_configure("bold-" + name, foreground=color, font=self._log_bold)

    def _log_tag(self, color):
        if str(color) in self.log_text.tag_names():
            return str(color)
        # unknown colour name from a script: create it on the fly, fall back to plain text
        try:
            self.log_text.tag_configure("custom-" + str(color), foreground=str(color))
            return "custom-" + str(color)
        except tk.TclError:
            return "black"

    def _write_log(self, stamp, text, color, newline):
        log = self.log_text
        log.configure(state="normal")
        log.insert("end", "[" + stamp + "] ", "time")
        log.insert("end", text + ("\n" if newline else ""), self._log_tag(color))
        log.configure(state="disabled")
        log.see("end")

    def add_log(self, text, color="black", printNewLine=True):
        if (self.output_log == None):
            return

        timeOfDay = time.strftime("%H:%M:%S", time.localtime())
        self._log_history.append((timeOfDay, str(text), color, printNewLine))
        del self._log_history[:-1500]

        self._write_log(timeOfDay, str(text), color, printNewLine)

    def clear_log(self):
        self._log_history.clear()
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    # ------------------------------------------------------------ appearance
    def set_appearance(self, theme):
        if theme not in THEMES or theme == self.theme:
            return
        self.config.set_setting('ui', 'appearance', theme)
        # rebuild after the click handler returns (the clicked widget is about to be replaced)
        self.root.after(30, lambda: self._switch_theme(theme))

    def _switch_theme(self, theme):
        """Rebuild the interface in another theme, keeping log, status and running state."""
        tab = self.tabs.get()

        for sheet in list(self._sheets.values()):
            try:
                sheet.destroy()
            except Exception:
                pass
        self._sheets.clear()

        self.output_log = None
        for widget in self._main_widgets:
            widget.destroy()

        self.theme = theme
        UI.apply(theme)
        ctk.set_appearance_mode(THEME_MODE[theme])
        self.root.configure(fg_color=UI.BG)

        self.create_widgets()
        self.tabs.set(tab)
        self.show_tab(tab)

        # the floating script panel is recreated so it picks up the new colours too
        try:
            if self.script_running and self.script_settings_window is not None and self.script_settings_window.winfo_exists():
                self.script_settings_window.destroy()
                self.load_settings_window()
        except Exception as e:
            print(e)

        self.update_monitor()

    def center_window(self, window):
        place_near(window, self.root)

    # ------------------------------------------------------------ settings dialogs
    def updateSetting(self, section, key, value):
        self.config.set_setting(section, key, value)

    def show_input_device_settings(self):
        sheet = self.open_sheet("input", "Input Device", "joystick.ico")
        if sheet is None:
            return

        save = lambda key: (lambda value: self.updateSetting("input", key, value))

        sheet.add_field("Device Index", lambda p: make_option(
            p, range(6), self.config.get_setting("input", "device_index", 0), save("device_index"), 200))
        sheet.add_field("Device Type", lambda p: make_option(
            p, ["None", "Xbox Controller", "Playstation Controller"],
            self.config.get_setting("input", "device_type", "None"), save("device_type"), 200))
        sheet.show()

    def show_server_settings(self):
        sheet = self.open_sheet("server", "Server", "server.ico")
        if sheet is None:
            return

        sheet.add_field("Enabled", lambda p: make_option(
            p, ["Enabled", "Disabled"], self.config.get_setting("server", "enabled", "Disabled"),
            lambda value: self.updateSetting("server", "enabled", value), 200))

        def text_field(key, default):
            def build(parent):
                entry = make_entry(parent, self.config.get_setting("server", key, default), width=200)
                entry.bind("<KeyRelease>", lambda _e: self.updateSetting("server", key, entry.get()))
                return entry
            return build

        sheet.add_field("IP Address", text_field("ip_address", "localhost"))
        sheet.add_field("Port", text_field("port", "5647"))
        sheet.show()

    def show_capture_settings(self):
        sheet = self.open_sheet("capture", "Capture Method", "zoom.ico")
        if sheet is None:
            return

        cfg = self.config
        save = lambda key: (lambda value: self.updateSetting("capture", key, value))

        digits_only = sheet.register(lambda text: text.isdigit() or text == "")

        def number_field(key, default=0):
            def build(parent):
                entry = make_entry(parent, cfg.get_setting("capture", key, default), width=200,
                                   validate="key", validatecommand=(digits_only, "%P"))
                # an empty box is stored as 0 so the capture code never reads ""
                entry.bind("<KeyRelease>", lambda _e: self.updateSetting("capture", key, entry.get() or 0))
                return entry
            return build

        rows = {}

        def refresh(_value=None):
            kind = rows["type"][2].get()
            crop = rows["crop_mode"][2].get()
            rows["index"][1].configure(text="Monitor Index" if kind == "Desktop" else "Video Index")

            visible = {
                "method": kind == "Capture Card", "resolution": kind == "Capture Card",
                "crop_mode": kind == "Desktop",
                "left": kind == "Desktop" and crop == "Exact", "right": kind == "Desktop" and crop == "Exact",
                "top": kind == "Desktop" and crop == "Exact", "bottom": kind == "Desktop" and crop == "Exact",
                "width": kind == "Desktop" and crop == "From Center", "height": kind == "Desktop" and crop == "From Center",
            }
            for name, show in visible.items():
                rows[name][0].grid() if show else rows[name][0].grid_remove()

        def on_type(value):
            self.updateSetting("capture", "type", value)
            refresh()

        def on_crop_mode(value):
            self.updateSetting("capture", "crop_mode", value)
            refresh()

        kind = cfg.get_setting("capture", "type", "None")
        rows["type"] = sheet.add_field("Capture Type", lambda p: make_option(p, ["None", "Desktop", "Capture Card"], kind, on_type, 200))
        rows["index"] = sheet.add_field("Video Index", lambda p: make_option(
            p, range(6), cfg.get_setting("capture", "index", 0), save("index"), 200))
        rows["fps"] = sheet.add_field("FPS Limit", lambda p: make_option(
            p, ["15", "30", "60", "120", "144", "165", "240", "999"], cfg.get_setting("capture", "fps_limit", 144), save("fps_limit"), 200))
        rows["method"] = sheet.add_field("Video Method", lambda p: make_option(
            p, ["Microsoft Foundation", "Direct Show"], cfg.get_setting("capture", "video_capture_method", "Microsoft Foundation"),
            save("video_capture_method"), 200))
        rows["resolution"] = sheet.add_field("Video Resolution", lambda p: make_option(
            p, ["3840x2160", "2560x1440", "1920x1080", "1280x720"], cfg.get_setting("capture", "video_capture_resolution", "1920x1080"),
            save("video_capture_resolution"), 200))
        rows["crop_mode"] = sheet.add_field("Crop Mode", lambda p: make_option(
            p, ["Exact", "From Center"], cfg.get_setting("capture", "crop_mode", "Exact"), on_crop_mode, 200))
        rows["left"] = sheet.add_field("Left Crop", number_field("left_crop"))
        rows["right"] = sheet.add_field("Right Crop", number_field("right_crop"))
        rows["top"] = sheet.add_field("Top Crop", number_field("top_crop"))
        rows["bottom"] = sheet.add_field("Bottom Crop", number_field("bottom_crop"))
        rows["width"] = sheet.add_field("Width", number_field("center_crop_width"))
        rows["height"] = sheet.add_field("Height", number_field("center_crop_height"))

        refresh()
        sheet.show()

    def show_mouse_calibration_settings(self):
        sheet = self.open_sheet("mouse", "Mouse Calibration", "computer-mouse.ico")
        if sheet is None:
            return

        sliders = [
            ("Horizontal Sensitivity", "horizontal_sensitivity", 0.01, 20, 0.05, 10),
            ("Vertical Sensitivity", "vertical_sensitivity", 0.01, 20, 0.05, 10),
            ("Smoothness", "smoothness", 1, 20, 1, 10),
            ("Deadzone", "deadzone", 0.01, 1.00, 0.01, 0.05),
            ("Deadzone Amplification", "deadzone_amplification", 1, 2, 0.01, 1.5),
            ("Acceleration Dampener", "acceleration_dampener", 0.01, 1, 0.01, 0.7),
        ]

        for label, key, low, high, step, default in sliders:
            try:
                initial = float(self.config.get_setting("mouse", key, default))
            except (TypeError, ValueError):
                initial = float(default)
            sheet.add_slider(label, low, high, step, initial,
                             lambda value, key=key: self.updateSetting("mouse", key, value))
        sheet.show()

    # ------------------------------------------------------------ misc actions
    def on_mouse_move(self, event):
        # print mouse position
        self.set_status(f"Mouse Position: {event.x}, {event.y}")

    def updateSelf(self):
        self.notice("Update", "AntiSociales DS4 Windows AI will now close to update. Re-launch AntiSociales DS4 Windows AI once the update is complete.")

        try:
            os.startfile('update.bat')
            self.close_app()
        except:
            print("Failed to update!")

    def refresh_scripts(self, event=None):
        scripts = ["None"]
        scripts.extend(get_directories('scripts'))

        self.script_combo.configure(values=scripts)

        currentScript = self.script_combo.get()
        
        if (currentScript in scripts):
            self.script_combo.set(currentScript)
        else:
            self.script_combo.set("None")

    def show_hid_hider(self):
        subprocess.Popen(self.hidhide_client_path)

    def compile_script_files(self):
        files = filedialog.askopenfilenames(title="Select Files to Compile", filetypes=[('Python Files', '*.py')])

        if (len(files) > 0):
            self.add_log("Compiling files...", "blue")

            for file in files:
                self.add_log("Compiling file: " + str(file), "blue")
                self.root.update_idletasks()

                try:
                    Encryptor(file)
                except Exception as e:
                    self.add_log("Failed to compile file: " + str(file) + " (" + str(e) + ")", "red")

                self.root.update_idletasks()

            self.add_log("Finished compiling files!", "green")

            self.notice("Compile", "Finished compiling files!", badge=self.icons.get("compile"))

    def select_files_to_encrypt(self):
        files = filedialog.askopenfilenames(title="Select Files to Encrypt", filetypes=[('All Files', '*.*')])

        if (len(files) == 0):
            return

        password = self.ask_password("Encrypt files", "Enter a password to encrypt the selected files with.")

        if (password is None):
            self.add_log("Encryption cancelled.", "red")
            return

        self.add_log("Encrypting files...", "blue")

        for file in files:
            self.add_log("Encrypting file: " + str(file), "blue")

            try:
                pyAesCrypt.encryptFile(file, file + '.encrypted', password)
            except:
                self.add_log("Failed to encrypt file: " + str(file), "red")

        self.add_log("Finished encrypting files!", "green")

        self.notice("Encrypt", "Finished encrypting files!", badge=self.icons.get("lock"))

    def select_files_to_decrypt(self):
        files = filedialog.askopenfilenames(title="Select Files to Decrypt", filetypes=[('All Files', '*.*')])

        if (len(files) == 0):
            return

        password = self.ask_password("Decrypt files", "Enter the password these files were encrypted with.")

        if (password is None):
            self.add_log("Decryption cancelled.", "red")
            return

        self.add_log("Decrypting files...", "blue")

        for file in files:
            self.add_log("Decrypting file: " + str(file), "blue")

            try:
                pyAesCrypt.decryptFile(file, file.replace('.encrypted', '') + '.decrypted', password)
            except:
                self.add_log("Failed to decrypt file: " + str(file), "red")

        self.add_log("Finished decrypting files!", "green")

        self.notice("Decrypt", "Finished decrypting files!", badge=self.icons.get("lock"))

    def toggleAutoHide(self):
        if (self.config.get_setting('auto_hide', 'enabled', True)):
            self.config.set_setting('auto_hide', 'enabled', False)
            self.add_log("Auto Hide Devices Disabled", "red")
        else:
            self.config.set_setting('auto_hide', 'enabled', True)
            self.add_log("Auto Hide Devices Enabled", "green")

        # keep the switch truthful even if the setting could not be saved
        if (self.config.get_setting('auto_hide', 'enabled', True)):
            self.auto_hide_switch.select()
            self.auto_hide_switch.configure(border_color=UI.SWITCH_ON)
        else:
            self.auto_hide_switch.deselect()
            self.auto_hide_switch.configure(border_color=UI.SWITCH_OFF)

    def refresh_summaries(self):
        """Update the grey subtitles in the Settings list."""
        if not self._summary_rows:
            return
        try:
            self._summary_rows["input"].set_subtitle(str(self.peek("input", "device_type", "None")))
            self._summary_rows["capture"].set_subtitle(str(self.peek("capture", "type", "None")))

            if self.peek("server", "enabled", "Disabled") == "Enabled":
                server = "Enabled  \u00B7  " + str(self.peek("server", "ip_address", "localhost")) + ":" + str(self.peek("server", "port", "5647"))
            else:
                server = "Disabled"
            self._summary_rows["server"].set_subtitle(server)
        except Exception:
            pass

    # ------------------------------------------------------------ layout
    def show_tab(self, name):
        for tab, frame in self.tab_frames.items():
            if tab == name:
                frame.grid()
            else:
                frame.grid_remove()

    def _build_icons(self):
        badge = make_badge
        if UI.theme == "Ember":
            c = dict(input="#FF7A1A", capture="#F2542D", server="#F5A623", hid="#8C7163",
                     compile="#D9480F", shield="#FF9F43", lock="#7A5C49", theme="#E8590C")
        else:
            c = dict(input="#007AFF", capture="#AF52DE", server="#FF9500", hid="#8E8E93",
                     compile="#5856D6", shield="#34C759", lock="#636366", theme="#5E5CE6")
        self.icons = {
            "input":   badge(c["input"], _white_glyph_from_ico("joystick.ico")),
            "capture": badge(c["capture"], _white_glyph_from_ico("zoom.ico")),
            "server":  badge(c["server"], _white_glyph_from_ico("server.ico")),
            "hid":     badge(c["hid"], _draw_glyph("eye")),
            "compile": badge(c["compile"], _white_glyph_from_ico("python.ico")),
            "shield":  badge(c["shield"], _draw_glyph("shield")),
            "lock":    badge(c["lock"], _draw_glyph("lock")),
            "theme":   badge(c["theme"], _draw_glyph("contrast")),
        }

    def create_widgets(self):
        root = self.root
        self._build_icons()

        root.grid_columnconfigure(0, weight=1)
        root.grid_rowconfigure(2, weight=1)

        # ---- header
        header = ctk.CTkFrame(root, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=22, pady=(18, 10))
        header.grid_columnconfigure(2, weight=1)

        ctk.CTkLabel(header, text="DS4 Windows AI", font=UI.f_title, text_color=UI.TEXT).grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(header, text="v" + str(self.version), font=UI.f_small, text_color=UI.TEXT2,
                     fg_color=UI.CHIP, corner_radius=9, width=54, height=22).grid(
            row=0, column=1, sticky="w", padx=(10, 0), pady=(4, 0))
        self.status_pill = ctk.CTkLabel(header, text="\u25CF  Idle", font=UI.f_small, text_color=UI.TEXT2,
                                        fg_color=UI.CARD, corner_radius=13, width=104, height=26)
        self.status_pill.grid(row=0, column=3, sticky="e")

        # ---- tab switcher
        self.tabs = ctk.CTkSegmentedButton(
            root, values=["Dashboard", "Settings", "Tools"], command=self.show_tab, height=34, corner_radius=10,
            border_width=3, font=UI.f_bold, text_color=UI.TEXT,
            fg_color=UI.SEG,
            selected_color=UI.SEG_SEL, selected_hover_color=UI.SEG_SEL_HOVER,
            unselected_color=UI.SEG, unselected_hover_color=UI.SEG_HOVER)
        self.tabs.grid(row=1, column=0, sticky="ew", padx=22, pady=(0, 12))
        self._summary_rows = {}

        # ---- tab content
        content = ctk.CTkFrame(root, fg_color="transparent")
        content.grid(row=2, column=0, sticky="nsew", padx=22)
        content.grid_columnconfigure(0, weight=1)
        content.grid_rowconfigure(0, weight=1)

        self.tab_frames = {}
        for name in ("Dashboard", "Settings", "Tools"):
            frame = ctk.CTkFrame(content, fg_color="transparent")
            frame.grid(row=0, column=0, sticky="nsew")
            frame.grid_columnconfigure(0, weight=1)
            self.tab_frames[name] = frame

        self._build_dashboard(self.tab_frames["Dashboard"])
        self._build_settings_tab(self.tab_frames["Settings"])
        self._build_tools_tab(self.tab_frames["Tools"])

        # ---- status bar
        status_bar = ctk.CTkFrame(root, fg_color="transparent")
        status_bar.grid(row=3, column=0, sticky="ew", padx=24, pady=(8, 12))
        status_bar.grid_columnconfigure(0, weight=1)

        # create a label for the left column of the status bar
        self.status_bar_left_label = ctk.CTkLabel(status_bar, text=self._status_text, font=UI.f_small,
                                                  text_color=UI.TEXT2, anchor="w")
        self.status_bar_left_label.grid(row=0, column=0, sticky="w")

        # create a label for the right column of the status bar
        status_bar_right_label = ctk.CTkLabel(status_bar, text="Made with \U0001F496 by AntiSociales", font=UI.f_small,
                                              text_color=UI.TEXT2, anchor="e")
        status_bar_right_label.grid(row=0, column=1, sticky="e")

        self._main_widgets = [header, self.tabs, content, status_bar]

        self.tabs.set("Dashboard")
        self.show_tab("Dashboard")
        self.refresh_summaries()

        # restore status pill / button / log (matters when the theme was just switched)
        self._status_state = None
        self._last_status_time = 0.0
        self.set_status(self._status_text)
        self._start_button_mode(self._button_mode)

        if self._log_history:
            for entry in self._log_history:
                self._write_log(*entry)
        else:
            self.add_log("Welcome to AntiSociales DS4 Windows AI v" + str(self.version), "green")

    def _build_dashboard(self, parent):
        parent.grid_rowconfigure(2, weight=1)
        scale = self._window_scale()

        # ---- live controller card
        monitor = make_card(parent)
        monitor.grid(row=0, column=0, sticky="ew")

        top = ctk.CTkFrame(monitor, fg_color="transparent")
        top.pack(fill="x", padx=18, pady=(14, 0))
        ctk.CTkLabel(top, text="Controller", font=UI.f_head, text_color=UI.TEXT).pack(side="left")
        ctk.CTkLabel(top, text="\u25CF Sticks", font=UI.f_small, text_color=UI.STICK).pack(side="right")
        ctk.CTkLabel(top, text="\u25CF Buttons   ", font=UI.f_small, text_color=UI.HOT).pack(side="right")

        self._mon_scale = self.MONITOR_SCALE * scale
        self._mon_font = (UI.SANS, -max(9, round(11 * scale)), "bold")
        crop = self.MONITOR_CROP
        canvas_w = round((crop[2] - crop[0]) * self._mon_scale)
        canvas_h = round((crop[3] - crop[1]) * self._mon_scale)

        # add a canvas that fills the monitor card, and add the controller image to it
        self.device_monitor_canvas = tk.Canvas(monitor, width=canvas_w, height=canvas_h,
                                               highlightthickness=0, bd=0, bg=resolve(UI.CARD))
        self.device_monitor_canvas.pack(pady=(2, 12))
        self._build_monitor_image()
        self.device_monitor_canvas.create_image(0, 0, image=self.device_monitor_template, anchor="nw")

        # when mouse moves over the canvas print out the position of the mouse on the canvas
        #self.device_monitor_canvas.bind("<Motion>", self.on_mouse_move)

        # ---- controls card
        controls = make_card(parent)
        controls.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        controls.grid_columnconfigure(1, weight=1)

        self.start_button = ctk.CTkButton(
            controls, text="Start", command=self.start, width=112, height=42, corner_radius=14, font=UI.f_button,
            fg_color=UI.BLUE, hover_color=UI.BLUE_HOVER, text_color=UI.ON_ACCENT, text_color_disabled=UI.ON_ACCENT_DISABLED)
        self.start_button.grid(row=0, column=0, padx=(14, 10), pady=14)

        # add a combo box to the controls
        scripts = ["None"]
        scripts.extend(get_directories('scripts'))
        currentScript = self.config.get_setting("script", "name", scripts[0])

        self.script_combo = make_option(controls, scripts, None,
                                        lambda value: self.updateSetting("script", "name", value), 200)
        self.script_combo.configure(height=42, corner_radius=12)
        self.script_combo.grid(row=0, column=1, sticky="ew", pady=14)

        if (currentScript in scripts):
            self.script_combo.set(currentScript)
        else:
            self.script_combo.set("None")

        # add refresh button that updates the script combo list
        refresh_button = make_button(controls, "Refresh", self.refresh_scripts, "secondary", width=84, height=42)
        refresh_button.grid(row=0, column=2, padx=(10, 14), pady=14)

        # ---- output log card
        log_card = make_card(parent)
        log_card.grid(row=2, column=0, sticky="nsew", pady=(12, 0))
        log_card.grid_columnconfigure(0, weight=1)
        log_card.grid_rowconfigure(1, weight=1)

        log_head = ctk.CTkFrame(log_card, fg_color="transparent")
        log_head.grid(row=0, column=0, sticky="ew", padx=18, pady=(12, 8))
        ctk.CTkLabel(log_head, text="Output Log", font=UI.f_head, text_color=UI.TEXT).pack(side="left")
        make_button(log_head, "Clear", self.clear_log, "secondary", width=58, height=26, font=UI.f_small).pack(side="right")

        self.output_log = ctk.CTkTextbox(
            log_card, height=120, corner_radius=12, border_width=0, wrap="word", font=UI.f_mono,
            fg_color=UI.FIELD, text_color=UI.TEXT,
            scrollbar_button_color=UI.SCROLL, scrollbar_button_hover_color=UI.SCROLL_HOVER)
        self.output_log.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 12))
        self.output_log.configure(state="disabled")

        # the wrapped tk.Text lets us use per-tag fonts and colours
        self.log_text = getattr(self.output_log, "_textbox", self.output_log)
        self._log_bold = tkfont.Font(family=UI.MONO, size=-round(12 * scale), weight="bold")
        self.log_text.configure(padx=8, pady=6, spacing1=1, spacing3=1)
        self._apply_log_colors()

    def _build_settings_tab(self, parent):
        section_label(parent, "General").grid(row=0, column=0, sticky="w", padx=8, pady=(2, 6))

        general = Group(parent)
        general.grid(row=1, column=0, sticky="ew")

        auto_hide = general.add(
            "Auto Hide Devices", "Hide physical controllers from games while running", self.icons["shield"],
            trailing=lambda p: make_switch(p, command=self.toggleAutoHide,
                                           value=bool(self.config.get_setting('auto_hide', 'enabled', True))))
        self.auto_hide_switch = auto_hide.trailing

        def appearance_control(p):
            seg = ctk.CTkSegmentedButton(
                p, values=THEMES, command=self.set_appearance, width=200, height=28,
                corner_radius=8, border_width=2, font=UI.f_small, text_color=UI.TEXT,
                fg_color=UI.SEG,
                selected_color=UI.SEG_SEL, selected_hover_color=UI.SEG_SEL_HOVER,
                unselected_color=UI.SEG, unselected_hover_color=UI.SEG_HOVER)
            seg.set(self.theme)
            return seg

        general.add("Appearance", "Ember (orange), Light or Dark", self.icons["theme"], trailing=appearance_control)

        section_label(parent, "Configuration").grid(row=2, column=0, sticky="w", padx=8, pady=(18, 6))

        configuration = Group(parent)
        configuration.grid(row=3, column=0, sticky="ew")
        self._summary_rows = {
            "input":   configuration.add("Input Device", "", self.icons["input"], self.show_input_device_settings, chevron=True),
            "capture": configuration.add("Capture Method", "", self.icons["capture"], self.show_capture_settings, chevron=True),
            "server":  configuration.add("Server", "", self.icons["server"], self.show_server_settings, chevron=True),
        }
        configuration.add("HidHider", "Open the HidHide client", self.icons["hid"], self.show_hid_hider, chevron=True)

    def _build_tools_tab(self, parent):
        section_label(parent, "Scripts").grid(row=0, column=0, sticky="w", padx=8, pady=(2, 6))
        scripts = Group(parent)
        scripts.grid(row=1, column=0, sticky="ew")
        scripts.add("Compile Script", "Turn .py files into protected modules", self.icons["compile"],
                    self.compile_script_files, chevron=True)

        section_label(parent, "Files").grid(row=2, column=0, sticky="w", padx=8, pady=(18, 6))
        files = Group(parent)
        files.grid(row=3, column=0, sticky="ew")
        files.add("Encrypt Weights / Files", "Protect files with a password (AES)", self.icons["lock"],
                  self.select_files_to_encrypt, chevron=True)
        #files.add("Decrypt Weights / Files", "Restore files encrypted with AntiSociales DS4 Windows AI", self.icons["lock"],
        #          self.select_files_to_decrypt, chevron=True)

    def _ensure_visible(self):
        try:
            if self.root.state() in ("withdrawn", "iconic"):
                self.root.deiconify()
            self.root.lift()
        except Exception:
            pass

    def run(self):
        try:
            self.root.deiconify()
            self.root.lift()
            self.root.after(300, self._ensure_visible)
            self.root.after(1200, self._ensure_visible)
            self.root.mainloop()
        except Exception:
            import traceback
            with open(os.path.join(currentDirectory, "antisociales_error.log"), "a") as f:
                f.write(traceback.format_exc() + "\n")
            raise