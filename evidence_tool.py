"""エビデンス取得ツール（Windows / タスクトレイ常駐）

ホットキーまたはタスクトレイのメニューからスクリーンショットを撮り、
撮影日時付きで改ページ済みの Excel にまとめる。
"""
import ctypes
import datetime
import hashlib
import json
import os
import queue
import sys
import threading
import tkinter as tk
from ctypes import wintypes
from pathlib import Path
from tkinter import filedialog, messagebox

import pystray
from PIL import Image, ImageDraw, ImageEnhance, ImageFont, ImageGrab, ImageTk

from excel_report import build_report, export_session

APP_NAME = "エビデンス取得ツール"

# DPI 対応（高DPI環境で座標・画像サイズがずれないように、tk より先に設定する）
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

user32 = ctypes.windll.user32
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
dwmapi = ctypes.windll.dwmapi

DEFAULT_CONFIG = {
    "output_dir": "",                 # 空なら ドキュメント\エビデンス
    "hotkeys": {
        "fullscreen": "Ctrl+Alt+1",   # カーソルのあるモニター全体
        "window": "Ctrl+Alt+2",       # アクティブウィンドウ
        "region": "Ctrl+Alt+3",       # 範囲指定
        "export": "Ctrl+Alt+E",       # Excel 出力
    },
    "fullscreen_scope": "monitor",    # monitor = カーソルのあるモニター / all = 全モニター
    "ask_comment": False,             # 撮影ごとにメモを入力する
    "notify": True,                   # 撮影時に通知を出す
    "jpeg_quality": 90,               # 保存する JPG の画質 (1〜100)
    "watch_clipboard": True,          # PrtSc / Alt+PrtSc / Win+Shift+S の画像も取り込む
    "excel": {"paper": "A4", "orientation": "portrait", "per_page": 2},
}

MODE_LABELS = {"fullscreen": "全画面", "window": "ウィンドウ", "region": "範囲指定",
               "clipboard": "標準スクショ"}


def app_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def load_config():
    path = app_dir() / "config.json"
    cfg = json.loads(json.dumps(DEFAULT_CONFIG))
    if path.exists():
        try:
            user = json.loads(path.read_text(encoding="utf-8"))
            for k, v in user.items():
                if isinstance(v, dict) and isinstance(cfg.get(k), dict):
                    cfg[k].update(v)
                else:
                    cfg[k] = v
        except Exception as ex:
            messagebox.showwarning(APP_NAME, f"config.json を読み込めませんでした。既定値で起動します。\n{ex}")
    else:
        path.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    return cfg


# ---------------------------------------------------------------- Win32 ヘルパー

def window_title(hwnd):
    n = user32.GetWindowTextLengthW(hwnd)
    buf = ctypes.create_unicode_buffer(n + 1)
    user32.GetWindowTextW(hwnd, buf, n + 1)
    return buf.value


def window_class(hwnd):
    buf = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buf, 256)
    return buf.value


def window_rect(hwnd):
    """影を含まない見た目どおりのウィンドウ矩形。"""
    rect = wintypes.RECT()
    DWMWA_EXTENDED_FRAME_BOUNDS = 9
    if dwmapi.DwmGetWindowAttribute(hwnd, DWMWA_EXTENDED_FRAME_BOUNDS,
                                    ctypes.byref(rect), ctypes.sizeof(rect)) != 0:
        user32.GetWindowRect(hwnd, ctypes.byref(rect))
    return rect.left, rect.top, rect.right, rect.bottom


def virtual_screen():
    x, y = user32.GetSystemMetrics(76), user32.GetSystemMetrics(77)
    w, h = user32.GetSystemMetrics(78), user32.GetSystemMetrics(79)
    return x, y, x + w, y + h


class MONITORINFO(ctypes.Structure):
    _fields_ = [("cbSize", wintypes.DWORD), ("rcMonitor", wintypes.RECT),
                ("rcWork", wintypes.RECT), ("dwFlags", wintypes.DWORD)]


def cursor_monitor_rect():
    pt = wintypes.POINT()
    user32.GetCursorPos(ctypes.byref(pt))
    user32.MonitorFromPoint.restype = wintypes.HMONITOR
    hmon = user32.MonitorFromPoint(pt, 2)  # MONITOR_DEFAULTTONEAREST
    info = MONITORINFO()
    info.cbSize = ctypes.sizeof(MONITORINFO)
    user32.GetMonitorInfoW(hmon, ctypes.byref(info))
    r = info.rcMonitor
    return r.left, r.top, r.right, r.bottom


def clamp_to_screen(bbox):
    vl, vt, vr, vb = virtual_screen()
    l, t, r, b = bbox
    return max(l, vl), max(t, vt), min(r, vr), min(b, vb)


# ---------------------------------------------------------------- グローバルホットキー

MOD_KEYS = {"alt": 0x1, "ctrl": 0x2, "control": 0x2, "shift": 0x4, "win": 0x8}
SPECIAL_VK = {"printscreen": 0x2C, "prtsc": 0x2C, "pause": 0x13, "insert": 0x2D,
              "home": 0x24, "end": 0x23, "pageup": 0x21, "pagedown": 0x22,
              "space": 0x20, "scrolllock": 0x91}


def parse_hotkey(text):
    mods, vk = 0, None
    for part in [p.strip().lower() for p in text.split("+") if p.strip()]:
        if part in MOD_KEYS:
            mods |= MOD_KEYS[part]
        elif part in SPECIAL_VK:
            vk = SPECIAL_VK[part]
        elif len(part) >= 2 and part[0] == "f" and part[1:].isdigit():
            vk = 0x70 + int(part[1:]) - 1
        elif len(part) == 1 and part.isalnum():
            vk = ord(part.upper())
        else:
            raise ValueError(f"不明なキー: {part}")
    if vk is None:
        raise ValueError("キーが指定されていません")
    return mods, vk


class HotkeyThread(threading.Thread):
    """RegisterHotKey で登録し、押されたら callback(name) を呼ぶ。"""

    def __init__(self, hotkeys, callback):
        super().__init__(daemon=True)
        self.hotkeys = hotkeys
        self.callback = callback
        self.failed = []
        self.thread_id = None
        self.ready = threading.Event()

    def run(self):
        self.thread_id = kernel32.GetCurrentThreadId()
        names = {}
        for i, (name, text) in enumerate(self.hotkeys.items(), start=1):
            if not text:
                continue
            try:
                mods, vk = parse_hotkey(text)
            except ValueError:
                self.failed.append(text)
                continue
            if user32.RegisterHotKey(None, i, mods | 0x4000, vk):  # MOD_NOREPEAT
                names[i] = name
            else:
                self.failed.append(text)
        self.ready.set()
        msg = wintypes.MSG()
        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
            if msg.message == 0x0312 and msg.wParam in names:  # WM_HOTKEY
                self.callback(names[msg.wParam])
        for i in names:
            user32.UnregisterHotKey(None, i)

    def stop(self):
        if self.thread_id:
            user32.PostThreadMessageW(self.thread_id, 0x0012, 0, 0)  # WM_QUIT


# ---------------------------------------------------------------- アイコン

def make_icon(count=0, active=False):
    """監視中は緑、停止中は灰色。右上の赤丸は監視中の撮影枚数。"""
    color = (33, 115, 70) if active else (128, 128, 128)
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([4, 16, 60, 56], radius=8, fill=color)     # 本体
    d.rectangle([20, 9, 40, 18], fill=color)                       # ファインダー
    d.ellipse([20, 22, 44, 46], fill="white")
    d.ellipse([26, 28, 38, 40], fill=color)
    if active and count:
        d.ellipse([40, 0, 63, 23], fill=(220, 50, 50))
        txt = str(count) if count < 100 else "99"
        d.text((51.5, 11.5), txt, fill="white", anchor="mm", font=ImageFont.load_default(size=16))
    return img


# ---------------------------------------------------------------- 本体

class EvidenceApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self.cfg = load_config()
        out = self.cfg.get("output_dir") or str(Path.home() / "Documents" / "エビデンス")
        self.output_root = Path(os.path.expandvars(out))
        self.queue = queue.Queue()
        self.busy = False
        self.monitoring = False
        self.new_session()

        self.icon = pystray.Icon("evidence_tool", make_icon(), APP_NAME, self.build_menu())
        threading.Thread(target=self.icon.run, daemon=True).start()

        self.hotkeys = HotkeyThread(self.cfg["hotkeys"], lambda name: self.queue.put(name))
        self.hotkeys.start()
        self.hotkeys.ready.wait(3)
        if self.hotkeys.failed:
            self.root.after(1500, lambda: self.notify(
                "次のホットキーは他のアプリが使用中のため登録できませんでした：\n"
                + "、".join(self.hotkeys.failed), force=True))

        self.clip_seq = user32.GetClipboardSequenceNumber()
        self.clip_pending = None   # (シーケンス番号, 変化を検知した時刻)
        self.clip_hash = None
        self.root.after(100, self.poll)

    # ---- メニュー
    def build_menu(self):
        hk = self.cfg["hotkeys"]
        put = lambda name: (lambda icon, item: self.queue.put(name))
        on = lambda item: self.monitoring
        return pystray.Menu(
            pystray.MenuItem(lambda item: f"保存先を設定…（{self.short_path(self.output_root)}）",
                             put("choose_folder"), enabled=lambda item: not self.monitoring),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(lambda item: "■ 監視を終了" if self.monitoring else "● 監視を開始",
                             put("toggle_monitor"), default=True),
            pystray.MenuItem(f"Excel 出力\t{hk.get('export', '')}", put("export"),
                             enabled=lambda item: not self.monitoring and bool(self.entries)),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(f"全画面を撮影\t{hk.get('fullscreen', '')}", put("menu:fullscreen"), enabled=on),
            pystray.MenuItem(f"範囲を指定して撮影\t{hk.get('region', '')}", put("menu:region"), enabled=on),
            pystray.MenuItem("直前の1枚を取り消す", put("undo"),
                             enabled=lambda item: self.monitoring and bool(self.entries)),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("PrtSc・Win+Shift+S の画像も取り込む", put("toggle_clipboard"),
                             checked=lambda item: bool(self.cfg.get("watch_clipboard", True))),
            pystray.MenuItem("撮影時にメモを入力する", put("toggle_comment"),
                             checked=lambda item: bool(self.cfg.get("ask_comment"))),
            pystray.MenuItem("過去の撮影フォルダから Excel 出力…", put("export_past"),
                             enabled=lambda item: not self.monitoring),
            pystray.MenuItem("保存フォルダを開く", put("open_folder")),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("ツールを終了", put("quit")),
        )

    def poll(self):
        try:
            while True:
                cmd = self.queue.get_nowait()
                if self.busy:
                    continue
                self.dispatch(cmd)
        except queue.Empty:
            pass
        if not self.busy and self.monitoring:
            self.check_clipboard()
        self.root.after(80, self.poll)

    def check_clipboard(self):
        """Windows 標準のスクショ操作でクリップボードに入った画像をエビデンスとして取り込む。"""
        seq = user32.GetClipboardSequenceNumber()
        if seq == self.clip_seq:
            return
        if not self.cfg.get("watch_clipboard", True):
            self.clip_seq = seq
            return
        # 書き込みが続けて起きることがあるので、変化が落ち着いてから読む
        now = datetime.datetime.now()
        if not self.clip_pending or self.clip_pending[0] != seq:
            self.clip_pending = (seq, now)
            return
        if (now - self.clip_pending[1]).total_seconds() < 0.25:
            return
        try:
            img = ImageGrab.grabclipboard()
        except Exception:
            return   # 他のアプリがクリップボードを使用中。次の周期で再試行
        self.clip_seq = seq
        now = self.clip_pending[1]
        self.clip_pending = None
        if not isinstance(img, Image.Image):
            return   # 文字やファイルのコピーは無視
        digest = hashlib.md5(img.tobytes()).hexdigest()
        if digest == self.clip_hash:
            return   # 同じ画像が再度コピーされただけ
        self.clip_hash = digest
        hwnd = user32.GetForegroundWindow()
        title = window_title(hwnd) if hwnd else ""
        if window_class(hwnd) in ("Shell_TrayWnd", "Shell_SecondaryTrayWnd", "Progman", "WorkerW"):
            title = ""
        self.busy = True
        try:
            self.add_entry(img.convert("RGB"), now, "clipboard", title)
        except Exception as ex:
            self.notify(f"取り込みに失敗しました：{ex}", force=True)
        finally:
            self.busy = False

    def dispatch(self, cmd):
        if cmd.startswith("menu:") or cmd in ("fullscreen", "window", "region"):
            if not self.monitoring:
                self.notify("監視を開始していません。トレイメニューの「監視を開始」を押してください。", force=True)
                return
        if cmd == "export" and self.monitoring:
            self.notify("監視中です。「監視を終了」してから Excel 出力してください。", force=True)
            return
        if cmd.startswith("menu:"):
            # メニューが閉じて画面から消えるのを待ってから撮る
            mode = cmd.split(":", 1)[1]
            self.busy = True
            self.root.after(350, lambda: self.capture(mode))
            return
        handler = {
            "fullscreen": lambda: self.capture("fullscreen"),
            "window": lambda: self.capture("window"),
            "region": lambda: self.capture("region"),
            "export": self.export,
            "undo": self.undo,
            "toggle_comment": self.toggle_comment,
            "toggle_clipboard": self.toggle_clipboard,
            "choose_folder": self.choose_folder,
            "open_folder": self.open_folder,
            "export_past": self.export_past,
            "toggle_monitor": lambda: (self.stop_monitoring() if self.monitoring
                                       else self.start_monitoring()),
            "quit": self.quit,
        }.get(cmd)
        if handler:
            handler()

    # ---- セッション
    def new_session(self):
        self.session_id = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        self.session_dir = self.output_root / self.session_id
        self.entries = []
        self.exported_count = 0

    def save_meta(self):
        meta = {"title": f"エビデンス {self.session_id}", "entries": self.entries}
        (self.session_dir / "session.json").write_text(
            json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")

    def refresh_icon(self):
        n = len(self.entries)
        self.icon.icon = make_icon(n, self.monitoring)
        state = f"監視中 {n} 枚" if self.monitoring else "停止中"
        self.icon.title = f"{APP_NAME}（{state}）"
        self.icon.update_menu()

    def start_monitoring(self):
        if not self.ask_unexported("新しく監視を始める"):
            return
        self.new_session()
        self.session_dir.mkdir(parents=True, exist_ok=True)
        self.monitoring = True
        # 開始前からクリップボードにあった画像は取り込まない
        self.clip_seq = user32.GetClipboardSequenceNumber()
        self.clip_pending = None
        self.clip_hash = None
        self.refresh_icon()
        self.notify(f"監視を開始しました。保存先：{self.session_dir}", force=True)

    def stop_monitoring(self):
        self.monitoring = False
        self.refresh_icon()
        if self.entries:
            self.notify(f"監視を終了しました（{len(self.entries)} 枚）。"
                        "トレイメニューの「Excel 出力」で Excel にまとめられます。", force=True)
        else:
            self.notify("監視を終了しました（撮影なし）。", force=True)

    def notify(self, msg, force=False):
        if force or self.cfg.get("notify", True):
            try:
                self.icon.notify(msg, APP_NAME)
            except Exception:
                pass

    # ---- 撮影
    def capture(self, mode):
        self.busy = True
        try:
            self._capture(mode)
        except Exception as ex:
            self.notify(f"撮影に失敗しました：{ex}", force=True)
        finally:
            self.busy = False

    def _capture(self, mode):
        now = datetime.datetime.now()
        hwnd = user32.GetForegroundWindow()
        title = window_title(hwnd) if hwnd else ""
        if window_class(hwnd) in ("Shell_TrayWnd", "Shell_SecondaryTrayWnd", "NotifyIconOverflowWindow"):
            title = ""

        if mode == "fullscreen":
            if self.cfg.get("fullscreen_scope") == "all":
                img = ImageGrab.grab(all_screens=True)
            else:
                img = ImageGrab.grab(bbox=cursor_monitor_rect(), all_screens=True)
        elif mode == "window":
            if not hwnd or window_class(hwnd) in ("Shell_TrayWnd", "Progman", "WorkerW"):
                self.notify("撮影するウィンドウがありません。対象をクリックしてから押してください。", force=True)
                return
            l, t, r, b = clamp_to_screen(window_rect(hwnd))
            if r - l < 2 or b - t < 2:
                self.notify("ウィンドウが画面外にあるため撮影できません。", force=True)
                return
            img = ImageGrab.grab(bbox=(l, t, r, b), all_screens=True)
        elif mode == "region":
            vl, vt, vr, vb = virtual_screen()
            full = ImageGrab.grab(all_screens=True)
            box = self.select_region(full, vl, vt)
            if not box:
                return
            img = full.crop(box)
        else:
            return
        self.add_entry(img, now, mode, title)

    def add_entry(self, img, now, mode, title):
        comment = ""
        if self.cfg.get("ask_comment"):
            comment = self.ask_comment(len(self.entries) + 1)
            if comment is None:   # キャンセル = 撮影を破棄
                return

        self.session_dir.mkdir(parents=True, exist_ok=True)
        no = (self.entries[-1]["no"] + 1) if self.entries else 1
        fname = f"{no:04d}_{now:%Y%m%d_%H%M%S}.jpg"
        img.convert("RGB").save(self.session_dir / fname, "JPEG",
                                quality=int(self.cfg.get("jpeg_quality", 90)))
        self.entries.append({"no": no, "file": fname, "timestamp": now.isoformat(timespec="seconds"),
                             "mode": mode, "window": title, "comment": comment})
        self.save_meta()
        self.refresh_icon()
        verb = "取り込みました" if mode == "clipboard" else "撮影しました"
        self.notify(f"No.{no} を{verb}（{MODE_LABELS[mode]}）{now:%H:%M:%S}")

    def select_region(self, full, vx, vy):
        """画面を固めた状態で表示し、ドラッグで範囲を選ばせる。Esc / 右クリックで中止。"""
        w, h = full.size
        top = tk.Toplevel(self.root)
        top.overrideredirect(True)
        top.attributes("-topmost", True)
        top.geometry(f"{w}x{h}{vx:+d}{vy:+d}")
        dim = ImageTk.PhotoImage(ImageEnhance.Brightness(full).enhance(0.6))
        canvas = tk.Canvas(top, width=w, height=h, highlightthickness=0, cursor="crosshair")
        canvas.pack()
        canvas.create_image(0, 0, anchor="nw", image=dim)
        ml, mt, mr, _ = cursor_monitor_rect()
        canvas.create_text((ml + mr) // 2 - vx, mt - vy + 40,
                           text="ドラッグで範囲を選択（Esc / 右クリックで中止）",
                           fill="white", font=("Meiryo UI", 16, "bold"))
        state = {"start": None, "rect": None, "bright": None, "photo": None, "label": None, "box": None}

        def press(ev):
            state["start"] = (ev.x, ev.y)

        def drag(ev):
            if not state["start"]:
                return
            x0, y0 = state["start"]
            l, t, r, b = min(x0, ev.x), min(y0, ev.y), max(x0, ev.x), max(y0, ev.y)
            for key in ("rect", "bright", "label"):
                if state[key]:
                    canvas.delete(state[key])
            if r - l > 1 and b - t > 1:
                state["photo"] = ImageTk.PhotoImage(full.crop((l, t, r, b)))
                state["bright"] = canvas.create_image(l, t, anchor="nw", image=state["photo"])
            state["rect"] = canvas.create_rectangle(l, t, r, b, outline="#ff3030", width=2)
            state["label"] = canvas.create_text(l + 4, max(t - 14, 10), anchor="w",
                                                text=f"{r - l} × {b - t}", fill="white",
                                                font=("Meiryo UI", 10, "bold"))

        def release(ev):
            if not state["start"]:
                return
            x0, y0 = state["start"]
            l, t, r, b = min(x0, ev.x), min(y0, ev.y), max(x0, ev.x), max(y0, ev.y)
            if r - l >= 5 and b - t >= 5:
                state["box"] = (max(l, 0), max(t, 0), min(r, w), min(b, h))
            top.destroy()

        canvas.bind("<ButtonPress-1>", press)
        canvas.bind("<B1-Motion>", drag)
        canvas.bind("<ButtonRelease-1>", release)
        canvas.bind("<Button-3>", lambda ev: top.destroy())
        top.bind("<Escape>", lambda ev: top.destroy())
        top.after(50, lambda: (top.focus_force(), canvas.focus_set()))
        top.grab_set()
        self.root.wait_window(top)
        return state["box"]

    def ask_comment(self, no):
        """メモ入力ダイアログ。OK/スキップ = 文字列、キャンセル = None（撮影破棄）。"""
        dlg = tk.Toplevel(self.root)
        dlg.title(f"{APP_NAME} - No.{no} のメモ")
        dlg.attributes("-topmost", True)
        dlg.resizable(False, False)
        tk.Label(dlg, text="操作内容などのメモ（空欄可）", font=("Meiryo UI", 10)).pack(padx=12, pady=(12, 4), anchor="w")
        var = tk.StringVar()
        entry = tk.Entry(dlg, textvariable=var, width=60, font=("Meiryo UI", 11))
        entry.pack(padx=12, pady=4)
        result = {"value": None}

        def ok(ev=None):
            result["value"] = var.get().strip()
            dlg.destroy()

        btns = tk.Frame(dlg)
        btns.pack(pady=(4, 12))
        tk.Button(btns, text="OK (Enter)", width=12, command=ok).pack(side="left", padx=4)
        tk.Button(btns, text="撮影を取り消す (Esc)", width=18, command=dlg.destroy).pack(side="left", padx=4)
        dlg.bind("<Return>", ok)
        dlg.bind("<Escape>", lambda ev: dlg.destroy())
        dlg.update_idletasks()
        sw, sh = dlg.winfo_screenwidth(), dlg.winfo_screenheight()
        dlg.geometry(f"+{(sw - dlg.winfo_width()) // 2}+{(sh - dlg.winfo_height()) // 3}")
        dlg.after(50, lambda: (dlg.focus_force(), entry.focus_set()))
        dlg.grab_set()
        self.root.wait_window(dlg)
        return result["value"]

    # ---- その他の操作
    def export(self):
        if not self.entries:
            self.notify("まだ撮影していません。", force=True)
            return
        ex = self.cfg.get("excel", {})
        base = self.session_dir / f"エビデンス_{self.session_id}.xlsx"
        out = base
        for i in range(1, 100):
            try:
                with open(out, "ab"):
                    pass
                break
            except PermissionError:   # Excel で開いたままのとき
                out = base.with_name(f"{base.stem}_{i}{base.suffix}")
        try:
            path = build_report(self.entries, self.session_dir, out,
                                title=f"エビデンス {self.session_id}",
                                paper=ex.get("paper", "A4"),
                                orientation=ex.get("orientation", "portrait"),
                                per_page=ex.get("per_page", 2))
        except Exception as e:
            self.notify(f"Excel の出力に失敗しました：{e}", force=True)
            return
        self.exported_count = len(self.entries)
        os.startfile(path)

    def undo(self):
        if not self.entries:
            return
        e = self.entries.pop()
        try:
            (self.session_dir / e["file"]).unlink()
        except FileNotFoundError:
            pass
        self.save_meta()
        self.refresh_icon()
        self.notify(f"No.{e['no']} を取り消しました。", force=True)

    @staticmethod
    def short_path(path, limit=40):
        text = str(path)
        return text if len(text) <= limit else "…" + text[-(limit - 1):]

    def choose_folder(self):
        """保存先フォルダを選び、config.json に覚えて次回以降の既定にする。"""
        parent = tk.Toplevel(self.root)
        parent.withdraw()
        parent.attributes("-topmost", True)
        initial = self.output_root if self.output_root.exists() else Path.home()
        chosen = filedialog.askdirectory(parent=parent, title="エビデンスの保存先を選択",
                                         initialdir=str(initial), mustexist=False)
        parent.destroy()
        if not chosen:
            return
        self.output_root = Path(chosen)
        self.cfg["output_dir"] = str(self.output_root)
        self.save_config()
        self.notify(f"保存先を設定しました：{self.output_root}", force=True)

    def toggle_clipboard(self):
        self.cfg["watch_clipboard"] = not self.cfg.get("watch_clipboard", True)
        self.clip_seq = user32.GetClipboardSequenceNumber()
        self.save_config()

    def toggle_comment(self):
        self.cfg["ask_comment"] = not self.cfg.get("ask_comment")
        self.save_config()

    def save_config(self):
        try:
            (app_dir() / "config.json").write_text(
                json.dumps(self.cfg, ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError as ex:
            self.notify(f"設定を保存できませんでした（次回は元に戻ります）：{ex}", force=True)
        self.icon.update_menu()

    def export_past(self):
        """以前に撮影したフォルダ（session.json があるフォルダ）を選んで Excel を作り直す。"""
        parent = tk.Toplevel(self.root)
        parent.withdraw()
        parent.attributes("-topmost", True)
        initial = self.output_root if self.output_root.exists() else Path.home()
        chosen = filedialog.askdirectory(parent=parent, title="Excel にする撮影フォルダを選択",
                                         initialdir=str(initial), mustexist=True)
        parent.destroy()
        if not chosen:
            return
        folder = Path(chosen)
        if not (folder / "session.json").exists():
            self.notify("撮影フォルダ（監視開始日時の名前のフォルダ）を選んでください。", force=True)
            return
        out = folder / f"エビデンス_{folder.name}.xlsx"
        for i in range(1, 100):
            try:
                with open(out, "ab"):
                    pass
                break
            except PermissionError:   # Excel で開いたままのとき
                out = out.with_name(f"エビデンス_{folder.name}_{i}.xlsx")
        try:
            path = export_session(folder, self.cfg.get("excel", {}), out)
        except Exception as e:
            self.notify(f"Excel の出力に失敗しました：{e}", force=True)
            return
        os.startfile(path)

    def open_folder(self):
        target = self.session_dir if self.session_dir.exists() else self.output_root
        target.mkdir(parents=True, exist_ok=True)
        os.startfile(target)

    def ask_unexported(self, action):
        """未出力の撮影があれば確認する。続行してよければ True。"""
        if len(self.entries) <= self.exported_count:
            return True
        parent = tk.Toplevel(self.root)
        parent.withdraw()
        parent.attributes("-topmost", True)
        ans = messagebox.askyesnocancel(
            APP_NAME, f"Excel に出力していない撮影があります（{len(self.entries)} 枚）。\n"
                      f"{action}前に Excel を出力しますか？\n（画像は保存フォルダに残ります）", parent=parent)
        parent.destroy()
        if ans is None:
            return False
        if ans:
            self.export()
        return True

    def quit(self):
        if not self.ask_unexported("終了する"):
            return
        self.monitoring = False
        self.hotkeys.stop()
        self.icon.stop()
        self.root.quit()

    def run(self):
        self.refresh_icon()
        self.notify("常駐しました。トレイアイコンを右クリック →「監視を開始」で撮影を始めます。", force=True)
        self.root.mainloop()


def main():
    # 二重起動防止
    kernel32.CreateMutexW(None, False, "EvidenceTool_SingleInstance")
    if ctypes.get_last_error() == 183:  # ERROR_ALREADY_EXISTS
        messagebox.showinfo(APP_NAME, "すでに起動しています。タスクトレイを確認してください。")
        return
    EvidenceApp().run()


if __name__ == "__main__":
    main()
