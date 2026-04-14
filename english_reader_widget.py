import base64
import ctypes
import json
import re
import subprocess
import threading
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from typing import Callable

import tkinter as tk
from tkinter import ttk
from ctypes import wintypes


YOUDAO_URL = "https://dict.youdao.com/suggest"
MYMEMORY_URL = "https://api.mymemory.translated.net/get"

WM_HOTKEY = 0x0312
WM_USER = 0x0400
WM_TRAYICON = WM_USER + 20
WM_RBUTTONUP = 0x0205
WM_LBUTTONUP = 0x0202
PM_REMOVE = 0x0001
HOTKEY_ID = 9281

MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004

HOTKEY_MODIFIERS = {
    "Ctrl+Alt": MOD_CONTROL | MOD_ALT,
    "Ctrl+Shift": MOD_CONTROL | MOD_SHIFT,
    "Alt+Shift": MOD_ALT | MOD_SHIFT,
    "Ctrl+Alt+Shift": MOD_CONTROL | MOD_ALT | MOD_SHIFT,
}
HOTKEY_KEYS = list("ABCDEFGHIJKLMNOPQRSTUVWXYZ") + [f"F{i}" for i in range(1, 13)]
VK_KEYS = {chr(code): code for code in range(ord("A"), ord("Z") + 1)}
VK_KEYS.update({f"F{i}": 0x70 + i - 1 for i in range(1, 13)})

FALLBACK_MEANINGS = {
    "hello": "你好；喂",
    "world": "世界",
    "book": "书；书籍",
    "read": "读；阅读",
    "write": "写；书写",
    "good": "好的；优秀的",
    "bad": "坏的；糟糕的",
    "happy": "快乐的；幸福的",
    "study": "学习；研究",
    "computer": "电脑；计算机",
    "word": "单词；词语",
    "apple": "苹果",
    "water": "水",
    "time": "时间；次数",
    "love": "爱；喜爱",
}


class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]


class MSG(ctypes.Structure):
    _fields_ = [
        ("hwnd", ctypes.c_void_p),
        ("message", ctypes.c_uint),
        ("wParam", ctypes.c_size_t),
        ("lParam", ctypes.c_size_t),
        ("time", ctypes.c_ulong),
        ("pt", POINT),
    ]


class NOTIFYICONDATAW(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("hWnd", wintypes.HWND),
        ("uID", wintypes.UINT),
        ("uFlags", wintypes.UINT),
        ("uCallbackMessage", wintypes.UINT),
        ("hIcon", wintypes.HICON),
        ("szTip", wintypes.WCHAR * 128),
    ]


@dataclass
class LookupResult:
    word: str
    meaning: str
    source: str


class TrayIcon:
    NIM_ADD = 0
    NIM_DELETE = 2
    NIF_MESSAGE = 1
    NIF_ICON = 2
    NIF_TIP = 4
    IDI_APPLICATION = 32512

    MF_STRING = 0x0000
    MF_SEPARATOR = 0x0800
    MF_CHECKED = 0x0008
    MF_POPUP = 0x0010
    TPM_RIGHTBUTTON = 0x0002
    TPM_RETURNCMD = 0x0100

    CMD_SHOW_HIDE = 1001
    CMD_AUTO_READ = 1002
    CMD_AUTO_LOOKUP = 1003
    CMD_TOPMOST = 1004
    CMD_HOTKEY = 1005
    CMD_SPEED_SLOW = 1006
    CMD_SPEED_NORMAL = 1007
    CMD_SPEED_FAST = 1008
    CMD_EXIT = 1099
    CMD_VOICE_BASE = 1200

    def __init__(self, app: "EnglishReaderWidget") -> None:
        self.app = app
        self.root = app.root
        self.user32 = ctypes.windll.user32
        self.shell32 = ctypes.windll.shell32
        self.hwnd = wintypes.HWND(self.root.winfo_id())
        self._callback = None
        self._old_wndproc = None
        self._active = False
        self._setup_win_api()

    def _setup_win_api(self) -> None:
        self.WNDPROC = ctypes.WINFUNCTYPE(ctypes.c_longlong, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)
        self.user32.CallWindowProcW.argtypes = [ctypes.c_void_p, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
        self.user32.CallWindowProcW.restype = ctypes.c_longlong
        self.user32.LoadIconW.argtypes = [wintypes.HINSTANCE, wintypes.LPCWSTR]
        self.user32.LoadIconW.restype = wintypes.HICON
        self.user32.TrackPopupMenu.argtypes = [
            wintypes.HMENU,
            wintypes.UINT,
            ctypes.c_int,
            ctypes.c_int,
            ctypes.c_int,
            wintypes.HWND,
            ctypes.c_void_p,
        ]
        self.user32.TrackPopupMenu.restype = wintypes.UINT

    def add(self) -> None:
        if self._active:
            return
        self._callback = self.WNDPROC(self._wndproc)
        if hasattr(self.user32, "SetWindowLongPtrW"):
            self.user32.SetWindowLongPtrW.restype = ctypes.c_void_p
            self.user32.SetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_void_p]
            self._old_wndproc = self.user32.SetWindowLongPtrW(self.hwnd, -4, self._callback)
        else:
            self.user32.SetWindowLongW.restype = ctypes.c_void_p
            self.user32.SetWindowLongW.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_void_p]
            self._old_wndproc = self.user32.SetWindowLongW(self.hwnd, -4, self._callback)

        nid = NOTIFYICONDATAW()
        nid.cbSize = ctypes.sizeof(NOTIFYICONDATAW)
        nid.hWnd = self.hwnd
        nid.uID = 1
        nid.uFlags = self.NIF_MESSAGE | self.NIF_ICON | self.NIF_TIP
        nid.uCallbackMessage = WM_TRAYICON
        nid.hIcon = self.user32.LoadIconW(None, ctypes.c_wchar_p(self.IDI_APPLICATION))
        nid.szTip = "EnglishSpiker"
        self.shell32.Shell_NotifyIconW(self.NIM_ADD, ctypes.byref(nid))
        self._active = True

    def remove(self) -> None:
        if not self._active:
            return
        nid = NOTIFYICONDATAW()
        nid.cbSize = ctypes.sizeof(NOTIFYICONDATAW)
        nid.hWnd = self.hwnd
        nid.uID = 1
        self.shell32.Shell_NotifyIconW(self.NIM_DELETE, ctypes.byref(nid))
        if self._old_wndproc:
            if hasattr(self.user32, "SetWindowLongPtrW"):
                self.user32.SetWindowLongPtrW(self.hwnd, -4, self._old_wndproc)
            else:
                self.user32.SetWindowLongW(self.hwnd, -4, self._old_wndproc)
        self._active = False

    def _wndproc(self, hwnd: wintypes.HWND, msg: int, wparam: int, lparam: int) -> int:
        if msg == WM_TRAYICON:
            if lparam == WM_RBUTTONUP:
                self.root.after(0, self.show_menu)
                return 0
            if lparam == WM_LBUTTONUP:
                self.root.after(0, self.app.toggle_window)
                return 0
        return self.user32.CallWindowProcW(self._old_wndproc, hwnd, msg, wparam, lparam)

    def show_menu(self) -> None:
        menu = self.user32.CreatePopupMenu()
        voice_menu = self.user32.CreatePopupMenu()

        self._append(menu, self.CMD_SHOW_HIDE, "隐藏窗口" if self.app.is_visible else "显示窗口")
        self._append(menu, self.CMD_AUTO_READ, "复制后自动朗读", self.app.auto_clipboard.get())
        self._append(menu, self.CMD_AUTO_LOOKUP, "自动显示中文释义", self.app.auto_lookup.get())
        self._append(menu, self.CMD_TOPMOST, "窗口始终置顶", self.app.always_on_top.get())
        self.user32.AppendMenuW(menu, self.MF_SEPARATOR, 0, None)

        for index, voice in enumerate(self.app.voices):
            is_checked = voice == self.app.voice.get()
            self._append(voice_menu, self.CMD_VOICE_BASE + index, voice, is_checked)
        self.user32.AppendMenuW(menu, self.MF_POPUP, voice_menu, "选择声音")

        self.user32.AppendMenuW(menu, self.MF_SEPARATOR, 0, None)
        self._append(menu, self.CMD_SPEED_SLOW, "语速：慢", self.app.rate.get() == -2)
        self._append(menu, self.CMD_SPEED_NORMAL, "语速：正常", self.app.rate.get() == 0)
        self._append(menu, self.CMD_SPEED_FAST, "语速：快", self.app.rate.get() == 2)
        self.user32.AppendMenuW(menu, self.MF_SEPARATOR, 0, None)
        self._append(menu, self.CMD_HOTKEY, f"设置快捷键：{self.app.hotkey_text()}")
        self._append(menu, self.CMD_EXIT, "退出程序")

        point = POINT()
        self.user32.GetCursorPos(ctypes.byref(point))
        self.user32.SetForegroundWindow(self.hwnd)
        command = self.user32.TrackPopupMenu(
            menu,
            self.TPM_RIGHTBUTTON | self.TPM_RETURNCMD,
            point.x,
            point.y,
            0,
            self.hwnd,
            None,
        )
        self.user32.DestroyMenu(menu)
        if command:
            self.app.handle_tray_command(command)

    def _append(self, menu: int, command: int, text: str, checked: bool = False) -> None:
        flags = self.MF_STRING | (self.MF_CHECKED if checked else 0)
        self.user32.AppendMenuW(menu, flags, command, text)


class EnglishReaderWidget:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Lexicon")
        self.root.overrideredirect(True)
        self.root.geometry("320x120+1200+740")
        self.root.minsize(320, 120)
        self.root.maxsize(380, 160)
        self.root.resizable(False, False)

        self.word = tk.StringVar(value="Lexicon")
        self.meaning = tk.StringVar(value="复制英语单词后自动朗读")
        self.status = tk.StringVar(value="Listening...")
        self.voice = tk.StringVar(value="系统默认")
        self.rate = tk.IntVar(value=0)
        self.auto_clipboard = tk.BooleanVar(value=True)
        self.auto_lookup = tk.BooleanVar(value=True)
        self.always_on_top = tk.BooleanVar(value=True)
        self.hotkey_modifier = tk.StringVar(value="Ctrl+Alt")
        self.hotkey_key = tk.StringVar(value="R")

        self.voices = ["系统默认"]
        self.last_clipboard_text = ""
        self.last_lookup_word = ""
        self.hotkey_registered = False
        self.hotkey_after_id: str | None = None
        self.is_visible = True
        self.tray: TrayIcon | None = None

        self._build_ui()
        self._build_context_menu()
        self._enable_drag()
        self.root.update_idletasks()
        self.root.attributes("-topmost", True)
        self.root.protocol("WM_DELETE_WINDOW", self.hide_to_tray)

        self.tray = TrayIcon(self)
        self.tray.add()
        self.register_hotkey()
        self.poll_clipboard()
        self.poll_hotkey()
        self._load_voices()

    def _build_context_menu(self) -> None:
        self.context_menu = tk.Menu(self.root, tearoff=False)
        self.voice_menu = tk.Menu(self.context_menu, tearoff=False)
        self.speed_menu = tk.Menu(self.context_menu, tearoff=False)

    def show_context_menu(self, event: tk.Event | None = None) -> None:
        self.context_menu.delete(0, "end")
        self.voice_menu.delete(0, "end")
        self.speed_menu.delete(0, "end")

        self.context_menu.add_command(label="隐藏窗口" if self.is_visible else "显示窗口", command=self.toggle_window)
        self.context_menu.add_checkbutton(
            label="复制后自动朗读",
            variable=self.auto_clipboard,
            onvalue=True,
            offvalue=False,
        )
        self.context_menu.add_checkbutton(
            label="自动显示中文释义",
            variable=self.auto_lookup,
            onvalue=True,
            offvalue=False,
        )
        self.context_menu.add_checkbutton(
            label="窗口始终置顶",
            variable=self.always_on_top,
            onvalue=True,
            offvalue=False,
            command=lambda: self.root.attributes("-topmost", self.always_on_top.get()),
        )
        self.context_menu.add_separator()

        for voice in self.voices:
            self.voice_menu.add_radiobutton(label=voice, variable=self.voice, value=voice)
        self.context_menu.add_cascade(label="选择声音", menu=self.voice_menu)

        self.speed_menu.add_radiobutton(label="慢", variable=self.rate, value=-2)
        self.speed_menu.add_radiobutton(label="正常", variable=self.rate, value=0)
        self.speed_menu.add_radiobutton(label="快", variable=self.rate, value=2)
        self.context_menu.add_cascade(label="语速", menu=self.speed_menu)
        self.context_menu.add_separator()
        self.context_menu.add_command(label=f"设置快捷键：{self.hotkey_text()}", command=self.show_hotkey_dialog)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="退出程序", command=self.exit_app)

        if event:
            x, y = event.x_root, event.y_root
        else:
            x = self.root.winfo_pointerx()
            y = self.root.winfo_pointery()
        self.context_menu.tk_popup(x, y)
        self.context_menu.grab_release()

    def _build_ui(self) -> None:
        self.root.configure(bg="#f8f9fa")
        wrapper = tk.Frame(self.root, bg="#f8f9fa", highlightthickness=1, highlightbackground="#dbe4e7")
        wrapper.pack(fill="both", expand=True)

        header = tk.Frame(wrapper, bg="#f8f9fa", height=30)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(
            header,
            text="Lexicon",
            bg="#f8f9fa",
            fg="#2b3437",
            font=("Inter", 8, "bold"),
        ).pack(side="left", padx=(12, 0))
        close = tk.Label(header, text="×", bg="#f8f9fa", fg="#586064", font=("Segoe UI", 12, "bold"), cursor="hand2")
        close.pack(side="right", padx=(0, 10))
        close.bind("<Button-1>", lambda _event: self.hide_to_tray())

        body = tk.Frame(wrapper, bg="#f8f9fa")
        body.pack(fill="both", expand=True, padx=14, pady=(2, 8))
        self.word_entry = tk.Entry(
            body,
            textvariable=self.word,
            state="readonly",
            readonlybackground="#f8f9fa",
            relief="flat",
            bd=0,
            bg="#f8f9fa",
            fg="#2b3437",
            insertbackground="#2b3437",
            font=("Inter", 18, "bold"),
        )
        self.word_entry.pack(fill="x", anchor="w")
        tk.Label(
            body,
            textvariable=self.meaning,
            bg="#f8f9fa",
            fg="#586064",
            anchor="w",
            font=("Microsoft YaHei UI", 10),
            wraplength=288,
        ).pack(fill="x", anchor="w", pady=(2, 0))

        footer = tk.Frame(body, bg="#f8f9fa")
        footer.pack(fill="x", side="bottom")
        tk.Label(footer, text="●", bg="#f8f9fa", fg="#005ac2", font=("Segoe UI", 8)).pack(side="right")
        tk.Label(
            footer,
            textvariable=self.status,
            bg="#f8f9fa",
            fg="#586064",
            font=("Inter", 7, "bold"),
        ).pack(side="right", padx=(0, 6))

    def _enable_drag(self) -> None:
        def start(event: tk.Event) -> None:
            self._drag_x = event.x
            self._drag_y = event.y

        def move(event: tk.Event) -> None:
            x = self.root.winfo_x() + event.x - self._drag_x
            y = self.root.winfo_y() + event.y - self._drag_y
            self.root.geometry(f"+{x}+{y}")

        self.root.bind("<Button-1>", start)
        self.root.bind("<B1-Motion>", move)
        self.root.bind("<Button-3>", self.show_context_menu)
        self.root.bind("<Control-Button-1>", self.show_context_menu)

    def _load_voices(self) -> None:
        threading.Thread(target=self._load_voices_worker, daemon=True).start()

    def _load_voices_worker(self) -> None:
        voices = get_installed_voices()
        if voices:
            self.root.after(0, lambda: self.voices.extend([voice for voice in voices if voice not in self.voices]))

    def poll_clipboard(self) -> None:
        if self.auto_clipboard.get():
            text = self.get_clipboard_text()
            word = normalize_word(text)
            if word and text != self.last_clipboard_text:
                self.last_clipboard_text = text
                self.process_word(word, speak=True, lookup=self.auto_lookup.get())
        self.root.after(180, self.poll_clipboard)

    def poll_hotkey(self) -> None:
        msg = MSG()
        while ctypes.windll.user32.PeekMessageW(ctypes.byref(msg), None, WM_HOTKEY, WM_HOTKEY, PM_REMOVE):
            if msg.message == WM_HOTKEY and msg.wParam == HOTKEY_ID:
                self.read_from_clipboard()
        self.hotkey_after_id = self.root.after(80, self.poll_hotkey)

    def register_hotkey(self) -> None:
        if self.hotkey_registered:
            ctypes.windll.user32.UnregisterHotKey(None, HOTKEY_ID)
            self.hotkey_registered = False
        modifiers = HOTKEY_MODIFIERS.get(self.hotkey_modifier.get(), HOTKEY_MODIFIERS["Ctrl+Alt"])
        vk = VK_KEYS.get(self.hotkey_key.get(), VK_KEYS["R"])
        self.hotkey_registered = bool(ctypes.windll.user32.RegisterHotKey(None, HOTKEY_ID, modifiers, vk))
        self.status.set("Listening..." if self.hotkey_registered else "Hotkey used")

    def hotkey_text(self) -> str:
        return f"{self.hotkey_modifier.get()}+{self.hotkey_key.get()}"

    def handle_tray_command(self, command: int) -> None:
        if command == TrayIcon.CMD_SHOW_HIDE:
            self.toggle_window()
        elif command == TrayIcon.CMD_AUTO_READ:
            self.auto_clipboard.set(not self.auto_clipboard.get())
        elif command == TrayIcon.CMD_AUTO_LOOKUP:
            self.auto_lookup.set(not self.auto_lookup.get())
        elif command == TrayIcon.CMD_TOPMOST:
            self.always_on_top.set(not self.always_on_top.get())
            self.root.attributes("-topmost", self.always_on_top.get())
        elif command == TrayIcon.CMD_SPEED_SLOW:
            self.rate.set(-2)
        elif command == TrayIcon.CMD_SPEED_NORMAL:
            self.rate.set(0)
        elif command == TrayIcon.CMD_SPEED_FAST:
            self.rate.set(2)
        elif command == TrayIcon.CMD_HOTKEY:
            self.show_hotkey_dialog()
        elif command == TrayIcon.CMD_EXIT:
            self.exit_app()
        elif TrayIcon.CMD_VOICE_BASE <= command < TrayIcon.CMD_VOICE_BASE + len(self.voices):
            self.voice.set(self.voices[command - TrayIcon.CMD_VOICE_BASE])

    def show_hotkey_dialog(self) -> None:
        dialog = tk.Toplevel(self.root)
        dialog.title("设置快捷键")
        dialog.geometry("260x120")
        dialog.resizable(False, False)
        dialog.attributes("-topmost", True)
        dialog.configure(bg="#f8f9fa")
        tk.Label(dialog, text="重读剪贴板快捷键", bg="#f8f9fa", fg="#2b3437", font=("Microsoft YaHei UI", 10, "bold")).pack(
            anchor="w", padx=14, pady=(14, 8)
        )
        row = tk.Frame(dialog, bg="#f8f9fa")
        row.pack(fill="x", padx=14)
        ttk.Combobox(row, textvariable=self.hotkey_modifier, values=list(HOTKEY_MODIFIERS), state="readonly", width=13).pack(
            side="left"
        )
        ttk.Combobox(row, textvariable=self.hotkey_key, values=HOTKEY_KEYS, state="readonly", width=7).pack(
            side="left", padx=(8, 0)
        )
        ttk.Button(dialog, text="应用", command=lambda: (self.register_hotkey(), dialog.destroy())).pack(anchor="e", padx=14, pady=12)

    def process_word(self, text: str, *, speak: bool, lookup: bool) -> None:
        word = normalize_word(text)
        if not word:
            return
        self.word.set(word[:28])
        self.status.set("Reading...")
        if speak:
            self.speak_word(word)
        if lookup:
            self.lookup_word(word)

    def read_from_clipboard(self) -> None:
        word = normalize_word(self.get_clipboard_text())
        if word:
            self.process_word(word, speak=True, lookup=self.auto_lookup.get())

    def speak_word(self, word: str) -> None:
        voice = "" if self.voice.get() == "系统默认" else self.voice.get()
        rate = int(self.rate.get())
        threading.Thread(target=speak_word, args=(word, voice, rate, self._set_status_from_thread), daemon=True).start()

    def lookup_word(self, word: str) -> None:
        if word == self.last_lookup_word:
            return
        self.last_lookup_word = word
        self.meaning.set("查询中...")
        threading.Thread(target=self._lookup_worker, args=(word,), daemon=True).start()

    def _lookup_worker(self, word: str) -> None:
        try:
            result = lookup_meaning(word)
        except Exception:
            result = LookupResult(word=word, meaning="暂时没有查到中文释义", source="提示")
        self.root.after(0, lambda: self.meaning.set(result.meaning.splitlines()[0][:34]))

    def _set_status_from_thread(self, text: str) -> None:
        self.root.after(0, lambda: self.status.set(text))

    def get_clipboard_text(self) -> str:
        try:
            return self.root.clipboard_get()
        except tk.TclError:
            return ""

    def hide_to_tray(self) -> None:
        self.root.withdraw()
        self.is_visible = False

    def show_window(self) -> None:
        self.root.deiconify()
        self.root.lift()
        self.root.attributes("-topmost", self.always_on_top.get())
        self.is_visible = True

    def toggle_window(self) -> None:
        if self.is_visible:
            self.hide_to_tray()
        else:
            self.show_window()

    def exit_app(self) -> None:
        if self.hotkey_registered:
            ctypes.windll.user32.UnregisterHotKey(None, HOTKEY_ID)
        if self.tray:
            self.tray.remove()
        self.root.destroy()


def normalize_word(text: str) -> str:
    match = re.search(r"[A-Za-z][A-Za-z' -]*", text.strip())
    if not match:
        return ""
    word = match.group(0).strip(" '-").lower()
    return re.sub(r"\s+", " ", word)


def lookup_meaning(word: str) -> LookupResult:
    errors: list[str] = []
    meaning = try_lookup(lookup_youdao, word, errors)
    if meaning:
        return LookupResult(word=word, meaning=meaning, source="有道词典")
    meaning = try_lookup(lookup_mymemory, word, errors)
    if meaning:
        return LookupResult(word=word, meaning=meaning, source="MyMemory 翻译")
    fallback = FALLBACK_MEANINGS.get(word.lower())
    if fallback:
        return LookupResult(word=word, meaning=fallback, source="本地小词库")
    detail = "；".join(errors) if errors else "网络词典没有返回可用结果。"
    raise RuntimeError(f"{detail} 请检查网络后重试。")


def try_lookup(lookup: Callable[[str], str], word: str, errors: list[str]) -> str:
    try:
        return lookup(word)
    except Exception as exc:
        errors.append(str(exc))
        return ""


def lookup_youdao(word: str) -> str:
    query = urllib.parse.urlencode({"num": 5, "ver": "3.0", "doctype": "json", "cache": "false", "le": "en", "q": word})
    data = fetch_json(f"{YOUDAO_URL}?{query}")
    entries = data.get("data", {}).get("entries", [])
    lines = [entry.get("explain", "") for entry in entries if entry.get("explain")]
    return "\n".join(dict.fromkeys(lines))


def lookup_mymemory(word: str) -> str:
    query = urllib.parse.urlencode({"q": word, "langpair": "en|zh-CN"})
    data = fetch_json(f"{MYMEMORY_URL}?{query}")
    translated = data.get("responseData", {}).get("translatedText", "")
    if translated and translated.lower() != word.lower():
        return translated
    return ""


def fetch_json(url: str) -> dict:
    request = urllib.request.Request(
        url,
        headers={"User-Agent": "Mozilla/5.0 EnglishReader/1.0", "Accept": "application/json,text/plain,*/*"},
    )
    try:
        with urllib.request.urlopen(request, timeout=8) as response:
            payload = response.read().decode("utf-8", errors="replace")
    except urllib.error.URLError as exc:
        raise RuntimeError(f"网络请求失败：{exc.reason}") from exc
    return json.loads(payload)


def get_installed_voices() -> list[str]:
    script = """
Add-Type -AssemblyName System.Speech
$speaker = New-Object System.Speech.Synthesis.SpeechSynthesizer
$speaker.GetInstalledVoices() | ForEach-Object { $_.VoiceInfo.Name }
"""
    result = run_powershell(script)
    if result.returncode != 0:
        return []
    return [line.strip() for line in result.stdout.splitlines() if line.strip()]


def speak_word(word: str, voice: str, rate: int, set_status: Callable[[str], None]) -> None:
    escaped_word = word.replace("'", "''")
    escaped_voice = voice.replace("'", "''")
    select_voice = f"$speaker.SelectVoice('{escaped_voice}')" if escaped_voice else ""
    script = f"""
Add-Type -AssemblyName System.Speech
$speaker = New-Object System.Speech.Synthesis.SpeechSynthesizer
$speaker.Rate = {rate}
$speaker.Volume = 100
{select_voice}
$speaker.Speak('{escaped_word}')
"""
    result = run_powershell(script)
    set_status("Listening..." if result.returncode == 0 else "Voice error")


def run_powershell(script: str) -> subprocess.CompletedProcess[str]:
    encoded = base64.b64encode(script.encode("utf-16le")).decode("ascii")
    return subprocess.run(
        ["powershell", "-NoProfile", "-EncodedCommand", encoded],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
    )


def main() -> None:
    root = tk.Tk()
    EnglishReaderWidget(root)
    root.mainloop()


if __name__ == "__main__":
    main()
