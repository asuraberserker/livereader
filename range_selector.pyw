"""
range_selector.pyw - 划定截图范围，完成后通过 IPC 回报主程序
"""
import ctypes
import sys
import tkinter as tk
from multiprocessing.connection import Client
from pathlib import Path
from tkinter import messagebox, ttk

from PIL import ImageGrab, ImageTk

IPC_PORT = 6123


class RangeSelectorApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.withdraw()  # 先隐藏，截图后再显示 overlay

        # 获取所有屏幕的整体区域（支持多显示器）
        screen = ImageGrab.grab(all_screens=True)
        # bbox 左上角偏移
        import ctypes
        monitors = self._get_all_monitors()
        left   = min(m[0] for m in monitors)
        top    = min(m[1] for m in monitors)
        right  = max(m[2] for m in monitors)
        bottom = max(m[3] for m in monitors)
        self._offset_x = left
        self._offset_y = top
        w = right - left
        h = bottom - top

        screenshot = ImageGrab.grab(bbox=(left, top, right, bottom), all_screens=True)

        overlay = tk.Toplevel(self.root)
        overlay.attributes("-topmost", True)
        overlay.overrideredirect(True)
        overlay.geometry(f"{w}x{h}+{left}+{top}")

        canvas = tk.Canvas(overlay, highlightthickness=0, cursor="cross")
        canvas.pack(fill="both", expand=True)

        tk_img = ImageTk.PhotoImage(screenshot)
        canvas.image = tk_img
        canvas.create_image(0, 0, image=tk_img, anchor="nw")

        # 半透明蒙版提示
        canvas.create_rectangle(0, 0, w, h, fill="black", stipple="gray25", outline="")
        canvas.create_text(w // 2, 40, text="拖动鼠标选择字幕区域  |  右键或 Esc 取消",
                           fill="white", font=("微软雅黑", 14))

        state = {"sx": None, "sy": None, "rect": None}

        def on_press(e):
            state["sx"], state["sy"] = e.x, e.y
            if state["rect"]:
                canvas.delete(state["rect"])
            state["rect"] = canvas.create_rectangle(
                e.x, e.y, e.x, e.y, outline="#00ff66", width=2
            )

        def on_drag(e):
            if state["sx"] is not None and state["rect"]:
                canvas.coords(state["rect"], state["sx"], state["sy"], e.x, e.y)

        def on_release(e):
            if state["sx"] is None or state["rect"] is None:
                return
            x1, y1, x2, y2 = canvas.coords(state["rect"])
            lft = int(min(x1, x2))
            tp  = int(min(y1, y2))
            rgt = int(max(x1, x2))
            btm = int(max(y1, y2))
            ww  = rgt - lft
            hh  = btm - tp
            if ww < 5 or hh < 5:
                return
            # 转换为绝对屏幕坐标
            abs_x = self._offset_x + lft
            abs_y = self._offset_y + tp
            overlay.destroy()
            self._send_and_quit((abs_x, abs_y, ww, hh))

        def cancel(_=None):
            overlay.destroy()
            self.root.destroy()

        canvas.bind("<ButtonPress-1>", on_press)
        canvas.bind("<B1-Motion>", on_drag)
        canvas.bind("<ButtonRelease-1>", on_release)
        canvas.bind("<ButtonPress-3>", cancel)
        overlay.bind("<Escape>", cancel)
        overlay.protocol("WM_DELETE_WINDOW", cancel)

    def _send_and_quit(self, data: tuple) -> None:
        """把范围通过 IPC 发回主程序，然后退出。"""
        try:
            conn = Client(("127.0.0.1", IPC_PORT), authkey=b"livereader")
            conn.send({"type": "range", "data": data})
            conn.close()
        except Exception as e:
            messagebox.showerror("错误", f"无法回报主程序：{e}\n请确保主程序已启动。")
        finally:
            self.root.destroy()

    @staticmethod
    def _get_all_monitors():
        """返回所有显示器的 (left, top, right, bottom) 列表。"""
        import ctypes
        monitors = []

        def cb(hmon, hdc, rect, data):
            monitors.append((rect.left, rect.top, rect.right, rect.bottom))
            return True

        MONITORENUMPROC = ctypes.WINFUNCTYPE(
            ctypes.c_bool,
            ctypes.c_ulong, ctypes.c_ulong,
            ctypes.POINTER(ctypes.wintypes.RECT),
            ctypes.c_double,
        )
        ctypes.windll.user32.EnumDisplayMonitors(None, None, MONITORENUMPROC(cb), 0)
        return monitors or [(0, 0, 1920, 1080)]


def main() -> None:
    if sys.platform != "win32":
        raise RuntimeError("仅支持 Windows")
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = tk.Tk()
    RangeSelectorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
