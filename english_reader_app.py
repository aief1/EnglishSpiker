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


YOUDAO_URL = "https://dict.youdao.com/suggest"
MYMEMORY_URL = "https://api.mymemory.translated.net/get"

WM_HOTKEY = 0x0312
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

HOTKEY_KEYS = list("ABCDEFGHIJKLMNOPQRSTUVWXYZ") + ["F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12"]
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


@dataclass
class LookupResult:
    word: str
    meaning: str
    source: str


class EnglishReaderApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("复制即读")
        self.root.geometry("820x560")
        self.root.minsize(680, 480)

        self.word = tk.StringVar()
        self.status = tk.StringVar(value="复制一个英语单词，我会自动朗读。")
        self.hotkey_status = tk.StringVar(value="快捷键：Ctrl+Alt+R")
        self.voice = tk.StringVar(value="系统默认")
        self.rate = tk.IntVar(value=0)
        self.auto_clipboard = tk.BooleanVar(value=True)
        self.auto_lookup = tk.BooleanVar(value=True)
        self.hotkey_modifier = tk.StringVar(value="Ctrl+Alt")
        self.hotkey_key = tk.StringVar(value="R")

        self.lookup_after_id: str | None = None
        self.last_clipboard_text = ""
        self.last_lookup_word = ""
        self.current_meaning_word = ""
        self.hotkey_registered = False
        self.is_setting_word = False

        self._build_ui()
        self._bind_events()
        self._load_voices()
        self.register_hotkey()
        self.poll_clipboard()
        self.poll_hotkey()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _build_ui(self) -> None:
        self.root.configure(bg="#eef3f5")
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TButton", padding=(14, 9), font=("Microsoft YaHei UI", 10))
        style.configure("TCheckbutton", background="#eef3f5", font=("Microsoft YaHei UI", 10))
        style.configure("TCombobox", padding=(8, 5), font=("Microsoft YaHei UI", 10))
        style.configure("Horizontal.TScale", background="#eef3f5")

        shell = tk.Frame(self.root, bg="#eef3f5", padx=28, pady=24)
        shell.pack(fill="both", expand=True)

        title_row = tk.Frame(shell, bg="#eef3f5")
        title_row.pack(fill="x")
        tk.Label(
            title_row,
            text="复制即读",
            bg="#eef3f5",
            fg="#102027",
            font=("Microsoft YaHei UI", 24, "bold"),
        ).pack(side="left")
        tk.Label(
            title_row,
            textvariable=self.hotkey_status,
            bg="#dce8ed",
            fg="#23404a",
            padx=12,
            pady=6,
            font=("Microsoft YaHei UI", 10, "bold"),
        ).pack(side="right")

        tk.Label(
            shell,
            text="复制英语单词后自动朗读；需要重读时按快捷键。",
            bg="#eef3f5",
            fg="#52636d",
            font=("Microsoft YaHei UI", 11),
        ).pack(anchor="w", pady=(8, 18))

        self.entry = tk.Entry(
            shell,
            textvariable=self.word,
            font=("Segoe UI", 30, "bold"),
            relief="flat",
            bg="#ffffff",
            fg="#111827",
            insertbackground="#111827",
            highlightthickness=2,
            highlightbackground="#b8c7cf",
            highlightcolor="#0f766e",
        )
        self.entry.pack(fill="x", ipady=14)
        self.entry.focus_set()

        button_row = tk.Frame(shell, bg="#eef3f5")
        button_row.pack(fill="x", pady=(14, 14))
        ttk.Button(button_row, text="读剪贴板", command=self.read_from_clipboard).pack(side="left")
        ttk.Button(button_row, text="重读一次", command=self.speak_current_word).pack(side="left", padx=(10, 0))
        ttk.Button(button_row, text="查中文", command=self.lookup_current_word).pack(side="left", padx=(10, 0))
        ttk.Checkbutton(button_row, text="复制后自动读", variable=self.auto_clipboard).pack(side="left", padx=(18, 0))
        ttk.Checkbutton(button_row, text="自动查中文", variable=self.auto_lookup).pack(side="left", padx=(10, 0))

        settings = tk.Frame(shell, bg="#eef3f5")
        settings.pack(fill="x", pady=(0, 14))
        self._label(settings, "声音").grid(row=0, column=0, sticky="w")
        self.voice_box = ttk.Combobox(settings, textvariable=self.voice, values=["系统默认"], state="readonly", width=34)
        self.voice_box.grid(row=0, column=1, sticky="w", padx=(8, 24))

        self._label(settings, "语速").grid(row=0, column=2, sticky="w")
        ttk.Scale(settings, from_=-5, to=5, orient="horizontal", variable=self.rate, length=150).grid(
            row=0, column=3, sticky="w", padx=(8, 24)
        )

        self._label(settings, "快捷键").grid(row=1, column=0, sticky="w", pady=(12, 0))
        ttk.Combobox(
            settings,
            textvariable=self.hotkey_modifier,
            values=list(HOTKEY_MODIFIERS.keys()),
            state="readonly",
            width=15,
        ).grid(row=1, column=1, sticky="w", padx=(8, 8), pady=(12, 0))
        ttk.Combobox(settings, textvariable=self.hotkey_key, values=HOTKEY_KEYS, state="readonly", width=8).grid(
            row=1, column=1, sticky="w", padx=(144, 0), pady=(12, 0)
        )
        ttk.Button(settings, text="应用快捷键", command=self.register_hotkey).grid(
            row=1, column=2, sticky="w", pady=(12, 0), padx=(0, 10)
        )

        tk.Label(
            shell,
            text="中文意思",
            bg="#eef3f5",
            fg="#102027",
            font=("Microsoft YaHei UI", 12, "bold"),
        ).pack(anchor="w", pady=(6, 8))

        self.meaning_box = tk.Text(
            shell,
            height=8,
            wrap="word",
            font=("Microsoft YaHei UI", 15),
            relief="flat",
            padx=16,
            pady=14,
            fg="#17212b",
            bg="#ffffff",
            state="disabled",
        )
        self.meaning_box.pack(fill="both", expand=True)

        tk.Label(
            shell,
            textvariable=self.status,
            bg="#eef3f5",
            fg="#52636d",
            font=("Microsoft YaHei UI", 10),
        ).pack(anchor="w", pady=(12, 0))

    def _label(self, parent: tk.Widget, text: str) -> tk.Label:
        return tk.Label(parent, text=text, bg="#eef3f5", fg="#33444d", font=("Microsoft YaHei UI", 10, "bold"))

    def _bind_events(self) -> None:
        self.word.trace_add("write", self._on_word_changed)
        self.entry.bind("<Return>", lambda _event: self.process_word(self.word.get(), speak=True, lookup=True))

    def _load_voices(self) -> None:
        threading.Thread(target=self._load_voices_worker, daemon=True).start()

    def _load_voices_worker(self) -> None:
        voices = get_installed_voices()
        if voices:
            self.root.after(0, lambda: self.voice_box.configure(values=["系统默认", *voices]))

    def _on_word_changed(self, *_args: object) -> None:
        if self.is_setting_word:
            return
        if self.lookup_after_id:
            self.root.after_cancel(self.lookup_after_id)
        if self.auto_lookup.get():
            self.lookup_after_id = self.root.after(450, self.lookup_current_word)

    def poll_clipboard(self) -> None:
        if self.auto_clipboard.get():
            text = self.get_clipboard_text()
            word = normalize_word(text)
            if word and text != self.last_clipboard_text:
                self.last_clipboard_text = text
                self.process_word(word, speak=True, lookup=self.auto_lookup.get())
        self.root.after(250, self.poll_clipboard)

    def poll_hotkey(self) -> None:
        msg = MSG()
        while ctypes.windll.user32.PeekMessageW(ctypes.byref(msg), None, WM_HOTKEY, WM_HOTKEY, PM_REMOVE):
            if msg.message == WM_HOTKEY and msg.wParam == HOTKEY_ID:
                self.read_from_clipboard()
        self.root.after(80, self.poll_hotkey)

    def register_hotkey(self) -> None:
        if self.hotkey_registered:
            ctypes.windll.user32.UnregisterHotKey(None, HOTKEY_ID)
            self.hotkey_registered = False

        modifier_name = self.hotkey_modifier.get()
        key_name = self.hotkey_key.get()
        modifiers = HOTKEY_MODIFIERS.get(modifier_name, HOTKEY_MODIFIERS["Ctrl+Alt"])
        vk = VK_KEYS.get(key_name, VK_KEYS["R"])

        if ctypes.windll.user32.RegisterHotKey(None, HOTKEY_ID, modifiers, vk):
            self.hotkey_registered = True
            self.hotkey_status.set(f"快捷键：{modifier_name}+{key_name}")
            self.status.set(f"已启用快捷键：{modifier_name}+{key_name}")
        else:
            self.hotkey_status.set("快捷键注册失败")
            self.status.set("这个快捷键可能被其他软件占用了，换一个试试。")

    def read_from_clipboard(self) -> None:
        word = normalize_word(self.get_clipboard_text())
        if not word:
            self.status.set("剪贴板里没有识别到英语单词。")
            return
        self.process_word(word, speak=True, lookup=self.auto_lookup.get())

    def process_word(self, text: str, *, speak: bool, lookup: bool) -> None:
        word = normalize_word(text)
        if not word:
            self.status.set("请输入或复制一个英语单词。")
            return

        self.is_setting_word = True
        self.word.set(word)
        self.entry.icursor("end")
        self.is_setting_word = False

        if speak:
            self.speak_word(word)
        if lookup:
            self.lookup_word(word)

    def speak_current_word(self) -> None:
        word = normalize_word(self.word.get())
        if not word:
            self.status.set("请先输入或复制一个英语单词。")
            return
        self.speak_word(word)

    def speak_word(self, word: str) -> None:
        self.status.set(f"正在朗读：{word}")
        voice = "" if self.voice.get() == "系统默认" else self.voice.get()
        rate = int(float(self.rate.get()))
        threading.Thread(target=speak_word, args=(word, voice, rate, self._set_status_from_thread), daemon=True).start()

    def lookup_current_word(self) -> None:
        word = normalize_word(self.word.get())
        if not word:
            self._set_meaning("")
            self.status.set("请输入或复制一个英语单词。")
            return
        self.lookup_word(word)

    def lookup_word(self, word: str) -> None:
        if word == self.last_lookup_word and word == self.current_meaning_word:
            return
        self.last_lookup_word = word
        self.status.set(f"正在查询：{word}")
        self._set_meaning("查询中...")
        threading.Thread(target=self._lookup_worker, args=(word,), daemon=True).start()

    def _lookup_worker(self, word: str) -> None:
        try:
            result = lookup_meaning(word)
        except Exception as exc:
            result = LookupResult(word=word, meaning=f"暂时没有查到中文释义。\n原因：{exc}", source="提示")
        self.root.after(0, lambda: self._show_lookup_result(result))

    def _show_lookup_result(self, result: LookupResult) -> None:
        self.current_meaning_word = result.word
        self._set_meaning(result.meaning)
        self.status.set(f"释义来源：{result.source}")

    def _set_meaning(self, text: str) -> None:
        self.meaning_box.configure(state="normal")
        self.meaning_box.delete("1.0", "end")
        self.meaning_box.insert("1.0", text)
        self.meaning_box.configure(state="disabled")

    def _set_status_from_thread(self, text: str) -> None:
        self.root.after(0, lambda: self.status.set(text))

    def get_clipboard_text(self) -> str:
        try:
            return self.root.clipboard_get()
        except tk.TclError:
            return ""

    def on_close(self) -> None:
        if self.hotkey_registered:
            ctypes.windll.user32.UnregisterHotKey(None, HOTKEY_ID)
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
    query = urllib.parse.urlencode(
        {
            "num": 5,
            "ver": "3.0",
            "doctype": "json",
            "cache": "false",
            "le": "en",
            "q": word,
        }
    )
    data = fetch_json(f"{YOUDAO_URL}?{query}")
    entries = data.get("data", {}).get("entries", [])
    lines: list[str] = []

    for entry in entries:
        explain = entry.get("explain")
        if explain and explain not in lines:
            lines.append(explain)

    return "\n".join(lines[:5])


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
        headers={
            "User-Agent": "Mozilla/5.0 EnglishReader/1.0",
            "Accept": "application/json,text/plain,*/*",
        },
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
    escaped_word = escape_powershell_single_quote(word)
    escaped_voice = escape_powershell_single_quote(voice)
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
    if result.returncode == 0:
        set_status(f"已朗读：{word}")
    else:
        set_status("朗读失败：请换一个声音或检查 Windows 语音设置。")


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


def escape_powershell_single_quote(text: str) -> str:
    return text.replace("'", "''")


def main() -> None:
    root = tk.Tk()
    EnglishReaderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
