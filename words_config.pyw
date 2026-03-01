"""
words_config.pyw - 配置读音替换，保存后通过 IPC 回报主程序
"""
import ctypes
import sys
import tkinter as tk
from multiprocessing.connection import Client
from pathlib import Path
from tkinter import messagebox, ttk
from typing import List, Optional, Tuple

IPC_PORT = 6123


class WordsConfigApp:
    BG    = "#1f1f1f"
    FG    = "#f3f3f3"
    IBG   = "#2d2d30"
    IFG   = "#f3f3f3"
    ACCENT = "#3a3d41"

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("配置读音")
        self.root.geometry("550x320")

        self.base_dir = Path(__file__).resolve().parent
        self.words_path = self.base_dir / "words.txt"
        self.items: List[Tuple[str, str]] = []

        self._apply_theme()
        self._build_ui()
        self._load()

    def _apply_theme(self) -> None:
        self.root.configure(bg=self.BG)
        s = ttk.Style(self.root)
        s.theme_use("clam")
        s.configure(".", background=self.BG, foreground=self.FG)
        s.configure("TFrame", background=self.BG)
        s.configure("TLabel", background=self.BG, foreground=self.FG)
        s.configure("TButton", background=self.ACCENT, foreground=self.FG,
                    borderwidth=0, focusthickness=0, padding=(10, 6))
        s.configure("TEntry", fieldbackground=self.IBG, foreground=self.IFG)

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=12)
        frame.pack(fill="both", expand=True)

        self.listbox = tk.Listbox(
            frame, bg=self.IBG, fg=self.IFG,
            selectbackground="#3c3c3c", highlightthickness=0,
        )
        self.listbox.pack(fill="both", expand=True)
        self.listbox.bind("<<ListboxSelect>>", self._on_select)

        entry_row = ttk.Frame(frame)
        entry_row.pack(fill="x", pady=(10, 0))
        ttk.Label(entry_row, text="原词：").pack(side="left")
        self.src_var = tk.StringVar()
        ttk.Entry(entry_row, textvariable=self.src_var, width=16).pack(side="left", padx=(6, 12))
        ttk.Label(entry_row, text="替换为：").pack(side="left")
        self.dst_var = tk.StringVar()
        ttk.Entry(entry_row, textvariable=self.dst_var, width=16).pack(side="left", padx=(6, 0))

        btn_row = ttk.Frame(frame)
        btn_row.pack(fill="x", pady=(10, 0))
        ttk.Button(btn_row, text="新增", command=self._add).pack(side="left")
        ttk.Button(btn_row, text="修改", command=self._update).pack(side="left", padx=(8, 0))
        ttk.Button(btn_row, text="删除", command=self._delete).pack(side="left", padx=(8, 0))
        ttk.Button(btn_row, text="保存", command=self._save).pack(side="left", padx=(8, 0))

    def _load(self) -> None:
        self.items.clear()
        if self.words_path.exists():
            for line in self.words_path.read_text(encoding="utf-8").splitlines():
                if "=>" in line:
                    a, b = line.split("=>", 1)
                    self.items.append((a, b))
        self._refresh()

    def _refresh(self) -> None:
        self.listbox.delete(0, tk.END)
        for a, b in self.items:
            self.listbox.insert(tk.END, f"{a} => {b}")

    def _on_select(self, _=None) -> None:
        idx = self._sel()
        if idx is not None:
            self.src_var.set(self.items[idx][0])
            self.dst_var.set(self.items[idx][1])

    def _sel(self) -> Optional[int]:
        s = self.listbox.curselection()
        return int(s[0]) if s else None

    def _add(self) -> None:
        src = self.src_var.get().strip()
        if not src:
            messagebox.showwarning("提示", "原词不能为空")
            return
        self.items.append((src, self.dst_var.get().strip()))
        self._refresh()

    def _update(self) -> None:
        idx = self._sel()
        if idx is None:
            messagebox.showwarning("提示", "请先选择要修改的项目")
            return
        src = self.src_var.get().strip()
        if not src:
            messagebox.showwarning("提示", "原词不能为空")
            return
        self.items[idx] = (src, self.dst_var.get().strip())
        self._refresh()

    def _delete(self) -> None:
        idx = self._sel()
        if idx is None:
            messagebox.showwarning("提示", "请先选择要删除的项目")
            return
        self.items.pop(idx)
        self._refresh()

    def _save(self) -> None:
        # 写入文件
        content = "\n".join(f"{a}=>{b}" for a, b in self.items)
        self.words_path.write_text(content, encoding="utf-8")

        # 通过 IPC 通知主程序；若主程序未启动则静默忽略
        try:
            conn = Client(("127.0.0.1", IPC_PORT), authkey=b"livereader")
            conn.send({"type": "words", "data": self.items})
            conn.close()
        except Exception:
            pass

        self.root.destroy()


def main() -> None:
    if sys.platform != "win32":
        raise RuntimeError("仅支持 Windows")
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = tk.Tk()
    WordsConfigApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
