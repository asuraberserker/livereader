"""
Live Reader - 主程序
依赖：pip install paddlepaddle==3.2.0 paddleocr pynput pywin32 psutil opencv-python pillow
"""
import ctypes
import logging
import re
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
from dataclasses import dataclass
from multiprocessing.connection import Listener
from pathlib import Path
from tkinter import messagebox, ttk
from typing import Dict, List, Optional, Tuple

import cv2
import numpy as np
import psutil
import win32con
import win32gui
import win32process
from PIL import Image, ImageGrab, ImageTk
from paddleocr import PaddleOCR
from pynput import mouse as pynput_mouse

logging.getLogger("ppocr").setLevel(logging.WARNING)
logging.getLogger("paddle").setLevel(logging.WARNING)

# 子程序通信端口（本机 localhost）
IPC_PORT = 6123


# ─────────────────────────── OCR ───────────────────────────

class PaddleOcrRecognizer:
    LANGUAGE_MAP: Dict[str, str] = {
        "简体中文": "ch",
        "繁体中文": "chinese_cht",
        "英文": "en",
        "日文": "japan",
        "韩文": "korean",
    }
    DET_MODEL = "PP-OCRv5_mobile_det"
    REC_MODEL = "PP-OCRv5_mobile_rec"

    def __init__(self) -> None:
        self._instances: Dict[str, PaddleOCR] = {}
        self._tmp = str(Path(tempfile.mkdtemp(prefix="lr_")) / "frame.png")

    def recognize(self, image: Image.Image, lang_label: str) -> str:
        lang = self.LANGUAGE_MAP.get(lang_label, "ch")
        if lang not in self._instances:
            self._instances[lang] = PaddleOCR(
                text_detection_model_name=self.DET_MODEL,
                text_recognition_model_name=self.REC_MODEL,
                device="cpu",
            )
        bgr = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
        h, w = bgr.shape[:2]
        if max(h, w) > 640:
            s = 640 / max(h, w)
            bgr = cv2.resize(bgr, None, fx=s, fy=s, interpolation=cv2.INTER_AREA)
        cv2.imwrite(self._tmp, bgr)
        result = self._instances[lang].predict(self._tmp)
        if not result:
            return ""
        return "".join(t for t in result[0].get("rec_texts", []) if t.strip())


# ─────────────────────────── TTS ───────────────────────────

class SpeechWorker:
    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._voice = None

    def speak(self, text: str, rate: int = 200) -> None:
        self.stop()
        threading.Thread(target=self._run, args=(text,), daemon=True).start()

    def stop(self) -> None:
        with self._lock:
            v, self._voice = self._voice, None
        if v:
            try:
                v.Speak("", 3)
            except Exception:
                pass

    def _run(self, text: str) -> None:
        import win32com.client as wc
        try:
            v = wc.Dispatch("SAPI.SpVoice")
            v.Rate = 0
            with self._lock:
                self._voice = v
            v.Speak(text)
        except Exception:
            pass


# ─────────────────────────── 数据类 ───────────────────────────

@dataclass
class WindowInfo:
    hwnd: int
    title: str
    pid: int
    process_name: str

    @property
    def label(self) -> str:
        return f"{self.title}  ({self.process_name}, PID={self.pid})"


# ─────────────────────────── 主应用 ───────────────────────────

class LiveReaderApp:
    CLICK_DELAY_MS = 1000

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Live Reader")
        self.root.geometry("560x200")

        self.BG = "#1e1e1e"
        self.FG = "#ffffff"
        self.ACCENT = "#3a3f44"
        self.IBGC = "#2d2d2d"

        self._apply_theme()

        self.base_dir = Path(__file__).parent
        self.config_path = self.base_dir / "last_process.txt"
        self.range_path = self.base_dir / "capture_range.txt"
        self.words_path = self.base_dir / "words.txt"

        self.saved_process = self._load_text(self.config_path)
        self.capture_range: Optional[Tuple[int, int, int, int]] = self._load_range()
        self.word_mappings: List[Tuple[str, str]] = self._load_words()

        self.selected_window: Optional[WindowInfo] = None
        self.capture_running = False
        self._pending = False

        self.ocr = PaddleOcrRecognizer()
        self.speaker = SpeechWorker()
        self._mouse_listener: Optional[pynput_mouse.Listener] = None

        self._build_ui()
        self._refresh_windows()

        # IPC 服务线程：接收子程序回报
        threading.Thread(target=self._ipc_server, daemon=True).start()

    # ── 主题 ──

    def _apply_theme(self) -> None:
        self.root.configure(bg=self.BG)
        s = ttk.Style(self.root)
        s.theme_use("clam")
        s.configure(".", background=self.BG, foreground=self.FG)
        s.configure("TFrame", background=self.BG)
        s.configure("TLabel", background=self.BG, foreground=self.FG)
        s.configure("TButton", background=self.ACCENT, foreground=self.FG,
                    borderwidth=0, focusthickness=0, padding=(10, 6))
        s.map("TButton",
              background=[("active", "#4a4d52"), ("disabled", "#2c2c2c")],
              foreground=[("disabled", "#8b8b8b")])
        s.configure("TCombobox", fieldbackground=self.IBGC, background=self.IBGC,
                    foreground=self.FG, arrowcolor=self.FG,
                    selectbackground="#3c3c3c", selectforeground=self.FG)
        s.map("TCombobox", fieldbackground=[("readonly", self.IBGC)])

    # ── UI ──

    def _build_ui(self) -> None:
        f = ttk.Frame(self.root, padding=12)
        f.pack(fill="both", expand=True)

        row1 = ttk.Frame(f)
        row1.pack(fill="x", pady=(0, 6))
        ttk.Label(row1, text="识别语言：").pack(side="left")
        self.lang_var = tk.StringVar(value="简体中文")
        ttk.Combobox(row1, state="readonly", textvariable=self.lang_var,
                     values=["简体中文", "繁体中文", "英文", "日文", "韩文"],
                     width=12).pack(side="left", padx=(6, 0))
        ttk.Label(row1, text="读速：").pack(side="left", padx=(16, 0))
        self.rate_var = tk.IntVar(value=200)
        self.rate_lbl = tk.StringVar(value="200")
        ttk.Scale(row1, from_=50, to=400, orient="horizontal", variable=self.rate_var,
                  length=120,
                  command=lambda v: self.rate_lbl.set(str(int(float(v))))
                  ).pack(side="left", padx=(4, 0))
        ttk.Label(row1, textvariable=self.rate_lbl, width=4).pack(side="left", padx=(4, 0))

        row2 = ttk.Frame(f)
        row2.pack(fill="x", pady=6)
        self.win_combo = ttk.Combobox(row2, state="readonly")
        self.win_combo.pack(side="left", fill="x", expand=True)

        row3 = ttk.Frame(f)
        row3.pack(fill="x", pady=(8, 4))
        ttk.Button(row3, text="刷新列表", command=self._refresh_windows).pack(side="left")
        ttk.Button(row3, text="划定范围", command=self._open_range_tool).pack(side="left", padx=(8, 0))
        self.start_btn = ttk.Button(row3, text="开始捕获", command=self._start)
        self.start_btn.pack(side="left", padx=(8, 0))
        self.stop_btn = ttk.Button(row3, text="停止捕获", command=self._stop, state="disabled")
        self.stop_btn.pack(side="left", padx=(8, 0))
        ttk.Button(row3, text="配置读音", command=self._open_words_tool).pack(side="left", padx=(8, 0))

        self.status_var = tk.StringVar(value=self._range_hint())
        ttk.Label(f, textvariable=self.status_var).pack(anchor="w")

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _range_hint(self) -> str:
        return f"状态：已就绪  范围={'未设置' if not self.capture_range else self.capture_range}"

    def _set_status(self, text: str) -> None:
        self.status_var.set(text)
        print(f"[LiveReader] {text}", flush=True)

    # ── 窗口列表 ──

    def _refresh_windows(self) -> None:
        wins = self._enum_windows()
        self._win_map = {w.label: w for w in wins}
        self.win_combo["values"] = list(self._win_map)
        if wins:
            idx = next((i for i, w in enumerate(wins) if w.process_name == self.saved_process), 0)
            self.win_combo.current(idx)
        self._set_status(f"状态：已加载 {len(wins)} 个窗口")

    def _selected_window(self) -> Optional[WindowInfo]:
        label = self.win_combo.get().strip()
        if not label:
            messagebox.showwarning("提示", "请先选择窗口")
            return None
        w = self._win_map.get(label)
        if not w or not win32gui.IsWindow(w.hwnd):
            messagebox.showerror("错误", "窗口无效，请刷新后重试")
            return None
        return w

    # ── 子程序启动 ──

    def _open_range_tool(self) -> None:
        target = self._selected_window()
        if not target:
            return
        self.selected_window = target
        # 激活目标窗口再截图
        if win32gui.IsIconic(target.hwnd):
            win32gui.ShowWindow(target.hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(target.hwnd)
        self.root.after(300, lambda: subprocess.Popen(
            [sys.executable, str(self.base_dir / "range_selector.pyw")]
        ))

    def _open_words_tool(self) -> None:
        subprocess.Popen(
            [sys.executable, str(self.base_dir / "words_config.pyw")]
        )

    # ── 捕获 ──

    def _start(self) -> None:
        target = self._selected_window()
        if not target:
            return
        if not self.capture_range:
            messagebox.showwarning("提示", "请先划定识别范围")
            return
        self.selected_window = target
        if target.process_name != self.saved_process:
            self.saved_process = target.process_name
            self.config_path.write_text(self.saved_process, encoding="utf-8")
        self.capture_running = True
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self._set_status(f"状态：监听中 -> {target.title}")
        # 激活目标窗口
        if win32gui.IsIconic(target.hwnd):
            win32gui.ShowWindow(target.hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(target.hwnd)
        self._start_mouse_listener()

    def _stop(self) -> None:
        self.capture_running = False
        self._stop_mouse_listener()
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        self.speaker.stop()
        self._set_status("状态：已停止捕获")

    # ── 鼠标监听（pynput） ──

    def _start_mouse_listener(self) -> None:
        self._stop_mouse_listener()

        def on_click(x, y, button, pressed):
            if not pressed or button != pynput_mouse.Button.left:
                return
            if not self.capture_running or self._pending:
                return
            target = self.selected_window
            if not target:
                return
            fg = win32gui.GetForegroundWindow()
            _, fg_pid = win32process.GetWindowThreadProcessId(fg)
            if fg_pid != target.pid:
                return
            self._pending = True
            self.root.after(self.CLICK_DELAY_MS, self._do_capture)

        self._mouse_listener = pynput_mouse.Listener(on_click=on_click)
        self._mouse_listener.start()

    def _stop_mouse_listener(self) -> None:
        if self._mouse_listener:
            self._mouse_listener.stop()
            self._mouse_listener = None

    def _do_capture(self) -> None:
        self._pending = False
        if not self.capture_running or not self.capture_range:
            return
        x, y, w, h = self.capture_range
        try:
            img = ImageGrab.grab(bbox=(x, y, x + w, y + h), all_screens=True)
        except Exception:
            self._set_status("状态：截图失败")
            return
        self._set_status("状态：识别中…")
        try:
            text = self.ocr.recognize(img, self.lang_var.get())
        except Exception as e:
            self._set_status("状态：OCR 失败")
            print(f"[LiveReader] OCR error: {e}", flush=True)
            return
        normalized = self._normalize(text)
        if not normalized:
            self._set_status("状态：未识别到文字")
            return
        mapped = self._apply_words(normalized)
        self.speaker.stop()
        self.speaker.speak(mapped, rate=self.rate_var.get())
        preview = mapped[:30] + ("…" if len(mapped) > 30 else "")
        self._set_status(f"状态：朗读中 -> {preview}")
        print(f"[LiveReader] speak: {mapped}", flush=True)

    # ── IPC 服务（接收子程序回报） ──

    def _ipc_server(self) -> None:
        """在后台线程监听子程序发来的消息。"""
        try:
            with Listener(("127.0.0.1", IPC_PORT), authkey=b"livereader") as srv:
                while True:
                    try:
                        conn = srv.accept()
                        threading.Thread(
                            target=self._handle_ipc, args=(conn,), daemon=True
                        ).start()
                    except Exception:
                        pass
        except Exception as e:
            print(f"[LiveReader] IPC server error: {e}", flush=True)

    def _handle_ipc(self, conn) -> None:
        try:
            msg = conn.recv()  # {"type": "range"|"words", "data": ...}
            conn.close()
            msg_type = msg.get("type")
            if msg_type == "range":
                self.root.after(0, self._on_range_update, msg["data"])
            elif msg_type == "words":
                self.root.after(0, self._on_words_update, msg["data"])
        except Exception as e:
            print(f"[LiveReader] IPC handle error: {e}", flush=True)

    def _on_range_update(self, data: Tuple[int, int, int, int]) -> None:
        self.capture_range = tuple(data)
        self.range_path.write_text(
            f"abs:{data[0]},{data[1]},{data[2]},{data[3]}", encoding="utf-8"
        )
        self._set_status(f"状态：范围已更新 {self.capture_range}")

    def _on_words_update(self, data: List[Tuple[str, str]]) -> None:
        self.word_mappings = [tuple(item) for item in data]
        self._set_status("状态：读音配置已更新")

    # ── 文本处理 ──

    @staticmethod
    def _normalize(text: str) -> str:
        s = text.strip()
        if not s:
            return ""
        s = re.sub(r"\s+", "", s)
        s = s.replace("\u201c", "").replace("\u201d", "")
        for ch in "…~、。?？!！,.":
            s = s.replace(ch, "，")
        s = re.sub("，+", "，", s)
        s = re.sub(r"[^\w\u4e00-\u9fff\u3040-\u30ff\uac00-\ud7a3\uff0c\uff1a\uff1b]+$", "", s)
        return s.strip()

    def _apply_words(self, text: str) -> str:
        for old, new in self.word_mappings:
            if old:
                text = text.replace(old, new)
        return text

    # ── 持久化 ──

    def _load_range(self) -> Optional[Tuple[int, int, int, int]]:
        if not self.range_path.exists():
            return None
        raw = self.range_path.read_text(encoding="utf-8").strip()
        try:
            raw = raw.removeprefix("abs:")
            x, y, w, h = map(int, raw.split(","))
            return x, y, w, h
        except ValueError:
            return None

    def _load_words(self) -> List[Tuple[str, str]]:
        if not self.words_path.exists():
            return []
        result = []
        for line in self.words_path.read_text(encoding="utf-8").splitlines():
            if "=>" in line:
                a, b = line.split("=>", 1)
                result.append((a, b))
        return result

    @staticmethod
    def _load_text(path: Path) -> str:
        return path.read_text(encoding="utf-8").strip() if path.exists() else ""

    # ── 关闭 ──

    def _on_close(self) -> None:
        self.capture_running = False
        self._stop_mouse_listener()
        self.speaker.stop()
        self.root.destroy()

    # ── 枚举窗口 ──

    @staticmethod
    def _enum_windows() -> List[WindowInfo]:
        wins = []

        def cb(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return True
            title = win32gui.GetWindowText(hwnd).strip()
            if not title:
                return True
            if win32gui.GetWindow(hwnd, win32con.GW_OWNER) != 0:
                return True
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid == 0:
                return True
            try:
                pname = psutil.Process(pid).name()
            except Exception:
                pname = "unknown"
            wins.append(WindowInfo(hwnd, title, pid, pname))
            return True

        win32gui.EnumWindows(cb, 0)
        wins.sort(key=lambda w: w.title.lower())
        return wins


# ─────────────────────────── 入口 ───────────────────────────

def main() -> None:
    if sys.platform != "win32":
        raise RuntimeError("仅支持 Windows")
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = tk.Tk()
    LiveReaderApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
