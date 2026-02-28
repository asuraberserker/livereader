# Live Reader

A Windows desktop tool that listens for mouse clicks in a target game or application window, captures a designated screen region, recognizes text with OCR, and reads it aloud using the system text-to-speech engine.

Useful for visual novels, games with subtitles, or any situation where you want audio narration of on-screen text without modifying the application.

---

## Features

- **Click-triggered capture** — no polling; OCR runs only when you click in the target window
- **PP-OCRv5 recognition** — powered by PaddleOCR 3.x with PP-OCRv5 mobile models for low CPU usage
- **Multi-language support** — Simplified Chinese, Traditional Chinese, English, Japanese, Korean
- **Flexible region selection** — drag to select the subtitle/text area once; the selection is saved across sessions
- **Word replacement** — configure pronunciation corrections or custom substitutions via a companion tool
- **Dark UI** — clean dark-themed interface

---

## Requirements

- Windows 10 / 11 (64-bit)
- Python 3.10 – 3.12
- PaddlePaddle **3.2.0** (CPU version — other versions have known bugs on Windows CPU)

---

## Installation

```bash
# 1. Install PaddlePaddle CPU (must be exactly 3.2.0)
pip install paddlepaddle==3.2.0

# 2. Install remaining dependencies
pip install -r requirements.txt
```

> **Note:** Do not upgrade `paddlepaddle` beyond 3.2.0. Version 3.3.0 has a confirmed oneDNN bug on Windows CPU that causes inference to fail.

---

## File Structure

```
livereader.py         Main program
range_selector.pyw    Region selection tool (launched by main program)
words_config.pyw      Word replacement configuration tool
requirements.txt      Python dependencies
words.txt             Word replacement rules (auto-created on first save)
capture_range.txt     Saved capture region (auto-created after region selection)
last_process.txt      Last selected process name (auto-created)
```

---

## Usage

### 1. Start the main program

```bash
python livereader.py
```

### 2. Select target window

Choose the game or application window from the dropdown list. Click **刷新列表 (Refresh)** if it does not appear.

### 3. Define capture region

Click **划定范围 (Set Region)**. The screen dims and a crosshair cursor appears. Drag to select the area containing subtitles or dialogue text. Release the mouse to confirm — the region is sent back to the main program automatically.

### 4. Start capture

Click **开始捕获 (Start)**. The main program installs a global mouse listener.

From this point, every left click inside the target window triggers the following sequence:

```
click → wait 1 second → screenshot → OCR → read aloud
```

The 1-second delay allows the game to finish rendering the next line of text after you click.

### 5. Stop capture

Click **停止捕获 (Stop)** to remove the mouse listener and halt speech.

---

## Word Replacement

Click **配置读音 (Configure Pronunciation)** to open the word replacement editor.

Each rule maps an original string to a replacement string (useful for correcting misread characters or adjusting TTS pronunciation). Click **保存并通知主程序 (Save & Notify)** — changes take effect in the main program immediately without a restart.

Rules are stored in `words.txt` in the format:

```
original=>replacement
```

---

## IPC Communication

The three programs communicate over a local TCP connection on `127.0.0.1:6123` using `multiprocessing.connection` with a shared key. The main program acts as the server; the two companion tools connect, send one message, and exit.

| Message type | Sent by | Payload |
|---|---|---|
| `range` | range_selector.pyw | `(x, y, width, height)` absolute screen coordinates |
| `words` | words_config.pyw | list of `(original, replacement)` tuples |

---

## Troubleshooting

**OCR recognizes 0 lines of text**
Make sure the capture region actually contains text. The PP-OCRv5 models expect natural-color images — avoid regions that are entirely white or have unusual contrast.

**Speech does not play**
The program uses Windows SAPI (`SAPI.SpVoice`). Make sure at least one TTS voice is installed in Windows Settings → Time & Language → Speech.

**Mouse listener fails to install**
If you see `pynput` errors, ensure `pynput` is installed and that no other application is blocking global input hooks (some antivirus tools do this).

**Region selector does not show the correct screen**
On multi-monitor setups, the overlay covers all monitors. The coordinates saved are absolute screen coordinates and should work correctly regardless of which monitor the game runs on.

---

## License

MIT
