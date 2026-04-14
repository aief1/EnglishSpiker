from english_reader_app import main as run_new_app

if __name__ == "__main__":
    run_new_app()
    raise SystemExit

import base64
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
from tkinter import messagebox, ttk


YOUDAO_URL = "https://dict.youdao.com/suggest"
MYMEMORY_URL = "https://api.mymemory.translated.net/get"

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


@dataclass
class LookupResult:
    word: str
    meaning: str
    source: str


class EnglishReaderApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("英语单词朗读和中文释义")
        self.root.geometry("760x520")
        self.root.minsize(560, 420)

        self.lookup_after_id: str | None = None
        self.last_lookup_word = ""
        self.current_meaning_word = ""
        self.is_auto_read = tk.BooleanVar(value=True)
        self.status = tk.StringVar(value="输入或粘贴一个英语单词。")
        self.word = tk.StringVar()

        self._build_ui()
        self._bind_events()

    def _build_ui(self) -> None:
        self.root.configure(bg="#f5f7f8")
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TButton", padding=(14, 9), font=("Microsoft YaHei UI", 10))
        style.configure("TCheckbutton", background="#f5f7f8", font=("Microsoft YaHei UI", 10))

        shell = tk.Frame(self.root, bg="#f5f7f8", padx=28, pady=24)
        shell.pack(fill="both", expand=True)

        title = tk.Label(
            shell,
            text="英语单词朗读",
            bg="#f5f7f8",
            fg="#182027",
            font=("Microsoft YaHei UI", 22, "bold"),
        )
        title.pack(anchor="w")

        helper = tk.Label(
            shell,
            text="把英语单词复制到下面的框里，程序会读出来并显示中文意思。",
            bg="#f5f7f8",
            fg="#4b5863",
            font=("Microsoft YaHei UI", 11),
        )
        helper.pack(anchor="w", pady=(8, 18))

        input_row = tk.Frame(shell, bg="#f5f7f8")
        input_row.pack(fill="x")

        self.entry = tk.Entry(
            input_row,
            textvariable=self.word,
            font=("Segoe UI", 24),
            relief="solid",
            bd=1,
            highlightthickness=2,
            highlightbackground="#d4dce2",
            highlightcolor="#2d7ff9",
        )
        self.entry.pack(side="left", fill="x", expand=True, ipady=10)
        self.entry.focus_set()

        paste_button = ttk.Button(input_row, text="读剪贴板", command=self.read_from_clipboard)
        paste_button.pack(side="left", padx=(12, 0))

        actions = tk.Frame(shell, bg="#f5f7f8")
        actions.pack(fill="x", pady=(14, 18))

        ttk.Button(actions, text="朗读", command=self.speak_current_word).pack(side="left")
        ttk.Button(actions, text="查中文意思", command=self.lookup_current_word).pack(side="left", padx=(10, 0))
        ttk.Checkbutton(actions, text="输入后自动朗读和查询", variable=self.is_auto_read).pack(side="left", padx=(18, 0))

        meaning_label = tk.Label(
            shell,
            text="中文意思",
            bg="#f5f7f8",
            fg="#182027",
            font=("Microsoft YaHei UI", 12, "bold"),
        )
        meaning_label.pack(anchor="w", pady=(6, 8))

        self.meaning_box = tk.Text(
            shell,
            height=8,
            wrap="word",
            font=("Microsoft YaHei UI", 15),
            relief="solid",
            bd=1,
            padx=14,
            pady=12,
            fg="#17212b",
            bg="#ffffff",
            state="disabled",
        )
        self.meaning_box.pack(fill="both", expand=True)

        status_bar = tk.Label(
            shell,
            textvariable=self.status,
            bg="#f5f7f8",
            fg="#5e6b75",
            font=("Microsoft YaHei UI", 10),
        )
        status_bar.pack(anchor="w", pady=(12, 0))

    def _bind_events(self) -> None:
        self.word.trace_add("write", self._schedule_auto_lookup)
        self.entry.bind("<Return>", lambda _event: self.lookup_current_word())
        self.entry.bind("<Control-v>", self._handle_paste)

    def _handle_paste(self, _event: tk.Event) -> None:
        self.root.after(50, self._schedule_auto_lookup)

    def _schedule_auto_lookup(self, *_args: object) -> None:
        if self.lookup_after_id:
            self.root.after_cancel(self.lookup_after_id)
        if self.is_auto_read.get():
            self.lookup_after_id = self.root.after(650, self.lookup_current_word)

    def read_from_clipboard(self) -> None:
        try:
            text = self.root.clipboard_get()
        except tk.TclError:
            messagebox.showinfo("剪贴板为空", "剪贴板里没有可读取的文字。")
            return

        word = normalize_word(text)
        if not word:
            messagebox.showinfo("没有识别到单词", "剪贴板内容里没有英语单词。")
            return

        self.word.set(word)
        self.entry.icursor("end")
        self.lookup_current_word()

    def speak_current_word(self) -> None:
        word = normalize_word(self.word.get())
        if not word:
            self.status.set("请先输入一个英语单词。")
            return

        self.status.set(f"正在朗读：{word}")
        threading.Thread(target=speak_word, args=(word, self._set_status_from_thread), daemon=True).start()

    def lookup_current_word(self) -> None:
        word = normalize_word(self.word.get())
        if not word:
            self._set_meaning("")
            self.status.set("输入或粘贴一个英语单词。")
            return

        if word == self.last_lookup_word and word == self.current_meaning_word:
            self.speak_current_word()
            return

        self.last_lookup_word = word
        if self.word.get() != word:
            self.word.set(word)
        self.status.set(f"正在查询：{word}")
        self._set_meaning("查询中...")

        thread = threading.Thread(target=self._lookup_worker, args=(word,), daemon=True)
        thread.start()
        self.speak_current_word()

    def _lookup_worker(self, word: str) -> None:
        try:
            result = lookup_meaning(word)
        except Exception as exc:
            result = LookupResult(word=word, meaning=f"暂时没有查到中文释义。\n原因：{exc}", source="本地提示")

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


def speak_word(word: str, set_status: Callable[[str], None]) -> None:
    escaped_word = word.replace("'", "''")
    script = f"""
Add-Type -AssemblyName System.Speech
$speaker = New-Object System.Speech.Synthesis.SpeechSynthesizer
$speaker.Rate = -1
$speaker.Volume = 100
$speaker.Speak('{escaped_word}')
"""
    encoded = base64.b64encode(script.encode("utf-16le")).decode("ascii")
    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-EncodedCommand", encoded],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        set_status(f"已朗读：{word}")
    except (subprocess.CalledProcessError, FileNotFoundError) as exc:
        set_status(f"朗读失败：{exc}")


def main() -> None:
    root = tk.Tk()
    app = EnglishReaderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
