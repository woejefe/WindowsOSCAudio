#!/usr/bin/env python3
# OSC → Windows Audio Controller with GUI + Tray + Run-at-startup + Console toggle
# pip install python-osc pycaw comtypes psutil pystray pillow

import os
import sys
import threading
import traceback
import ctypes
from typing import Any, Optional

# --- OSC ---
from pythonosc import dispatcher
from pythonosc.osc_server import ThreadingOSCUDPServer

# --- Windows CoreAudio / COM ---
from comtypes import CoInitialize, CoUninitialize, CLSCTX_ALL, GUID
from ctypes import POINTER, cast, byref, c_void_p

from pycaw.pycaw import (
    AudioUtilities,
    ISimpleAudioVolume,
    IAudioEndpointVolume,
    IMMDeviceEnumerator,
    EDataFlow,
    ERole,
)

# --- GUI / Tray ---
import tkinter as tk
from tkinter import ttk, messagebox
import pystray
from PIL import Image, ImageDraw

# --- Misc ---
import psutil
import winreg

APP_NAME = "OSC Audio Controller"
REG_RUN_PATH = r"Software\Microsoft\Windows\CurrentVersion\Run"
RUN_VALUE_NAME = "OSCAudioController"

# ================= Console Manager =================

class ConsoleManager:
    """Create/hide/show a Windows console on demand (even under pythonw)."""
    SW_HIDE, SW_SHOW = 0, 5

    def __init__(self):
        self.kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        self.user32 = ctypes.WinDLL("user32", use_last_error=True)
        self.console_visible = False
        # Hide any existing console at startup (no console by default)
        hwnd = self.get_console_hwnd()
        if hwnd:
            self.user32.ShowWindow(hwnd, self.SW_HIDE)

    def get_console_hwnd(self):
        return self.kernel32.GetConsoleWindow()

    def _open_streams(self):
        try:
            sys.stdout = open("CONOUT$", "w", encoding="utf-8", buffering=1)
            sys.stderr = open("CONOUT$", "w", encoding="utf-8", buffering=1)
        except Exception:
            pass

    def _close_streams(self):
        for stream_name in ("stdout", "stderr"):
            try:
                s = getattr(sys, stream_name)
                if getattr(s, "name", "") == "CONOUT$":
                    s.close()
            except Exception:
                pass

    def show(self):
        hwnd = self.get_console_hwnd()
        if hwnd:
            self.user32.ShowWindow(hwnd, self.SW_SHOW)
            self.console_visible = True
            return
        # No console attached: allocate one
        if self.kernel32.AllocConsole():
            self._open_streams()
            self.console_visible = True
            print(f"{APP_NAME}: Console opened.")
        else:
            # Try attach to parent
            self.kernel32.AttachConsole(-1)  # ATTACH_PARENT_PROCESS
            self._open_streams()
            hwnd2 = self.get_console_hwnd()
            if hwnd2:
                self.user32.ShowWindow(hwnd2, self.SW_SHOW)
                self.console_visible = True

    def hide(self):
        hwnd = self.get_console_hwnd()
        if hwnd:
            self.user32.ShowWindow(hwnd, self.SW_HIDE)
        self.console_visible = False

    def toggle(self, want_visible: bool):
        if want_visible:
            self.show()
        else:
            self.hide()

console_mgr = ConsoleManager()

# =============== Helpers =================

def clamp01(x: float) -> float:
    return max(0.0, min(1.0, x))

def to_float(v: Any) -> Optional[float]:
    try:
        return float(v)
    except Exception:
        return None

def to_bool(v: Any) -> bool:
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return abs(float(v)) > 1e-9
    s = str(v).strip().lower()
    return s in ("1", "true", "on", "yes")

def normalize_volume(v: Any) -> Optional[float]:
    """Accepts 0..1 or 0..100; returns 0..1 clamped float."""
    f = to_float(v)
    if f is None:
        return None
    if f > 1.0:
        f = f / 100.0
    return clamp01(f)

def as_int(x, default):
    try:
        return int(x)
    except Exception:
        return default

# =============== Core Audio (robust) =================

# Official GUIDs
CLSID_MMDeviceEnumerator = GUID("{BCDE0395-E52F-467C-8E3D-C4579291692E}")
IID_IMMDeviceEnumerator  = GUID("{A95664D2-9614-4F35-A746-DE8DB63617E6}")

def _get_endpoint_volume() -> IAudioEndpointVolume:
    """
    Returns IAudioEndpointVolume for the default render/multimedia device.
    1) Try pycaw helper.
    2) Fallback to raw CoCreateInstance with GUIDs (works on all comtypes builds).
    """
    # Attempt 1: pycaw helper
    try:
        ev = AudioUtilities.GetAudioEndpointVolume()
        if ev:
            return ev
    except Exception:
        pass

    # Attempt 2: raw COM (avoid CreateObject/CoCreateInstance imports)
    p_enum = c_void_p()
    hr = ctypes.oledll.ole32.CoCreateInstance(
        byref(CLSID_MMDeviceEnumerator),
        None,
        CLSCTX_ALL,
        byref(IID_IMMDeviceEnumerator),
        byref(p_enum)
    )
    if hr != 0 or not p_enum:
        raise OSError(f"CoCreateInstance IMMDeviceEnumerator failed, hr=0x{hr & 0xFFFFFFFF:08X}")

    enumerator = cast(p_enum, POINTER(IMMDeviceEnumerator))
    df_render = as_int(getattr(EDataFlow, "eRender", 0), 0)           # 0 = Render
    role_multimedia = as_int(getattr(ERole, "eMultimedia", 1), 1)     # 1 = Multimedia

    device = enumerator.GetDefaultAudioEndpoint(df_render, role_multimedia)
    iface = device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    return cast(iface, POINTER(IAudioEndpointVolume))

def set_master_volume(vol01: float):
    ev = _get_endpoint_volume()
    ev.SetMasterVolumeLevelScalar(vol01, None)

def set_master_mute(mute: bool):
    ev = _get_endpoint_volume()
    ev.SetMute(bool(mute), None)

def set_app_volume(proc_name_no_ext: str, vol01: float) -> bool:
    """
    Sets per-app session volume. Returns True if at least one active session matched.
    The target process must be playing (or have recently played) audio to appear.
    """
    target = proc_name_no_ext.lower().removesuffix(".exe")
    changed = False
    for s in AudioUtilities.GetAllSessions():
        try:
            if s.Process is None:
                continue
            pname = s.Process.name().lower().removesuffix(".exe")
            if pname == target:
                s._ctl.QueryInterface(ISimpleAudioVolume).SetMasterVolume(vol01, None)
                changed = True
        except Exception:
            # Ignore zombie/permission issues
            pass
    return changed

# =============== OSC Server =================

class OscServer:
    def __init__(self):
        self.server: Optional[ThreadingOSCUDPServer] = None
        self.thread: Optional[threading.Thread] = None
        self.port = 9001
        self.running = False

    def _safe(self, fn):
        def wrapper(addr, *args):
            try:
                CoInitialize()
                fn(addr, *args)
            except Exception:
                print(f"[ERR] {addr} {args}")
                traceback.print_exc()
            finally:
                try:
                    CoUninitialize()
                except Exception:
                    pass
        return wrapper

    def _build_dispatcher(self):
        disp = dispatcher.Dispatcher()

        @self._safe
        def handle_ping(a, *x):
            print(f"[/ping] pong {x}")

        @self._safe
        def handle_master_volume(a, *x):
            if not x:
                print("[/master/volume] missing arg"); return
            v = normalize_volume(x[0])
            if v is None:
                print(f"[/master/volume] bad value: {x[0]}"); return
            set_master_volume(v)
            print(f"[/master/volume] set {v:.2f}")

        @self._safe
        def handle_master_mute(a, *x):
            if not x:
                print("[/master/mute] missing arg"); return
            b = to_bool(x[0])
            set_master_mute(b)
            print(f"[/master/mute] {'ON' if b else 'OFF'}")

        @self._safe
        def handle_app_volume(a, *x):
            """
            Supports both:
              /app/volume <processName> <value>
              /app/volume/<processName> <value>
            """
            proc = None
            vol = None
            if len(x) >= 2:
                proc = str(x[0]).strip().replace(".exe", "")
                vol = normalize_volume(x[1])
            elif len(x) >= 1 and a.startswith("/app/volume/"):
                proc = a.split("/app/volume/")[-1].strip()
                vol = normalize_volume(x[0])
            else:
                print(f"[/app/volume] usage: /app/volume <proc> <val> or /app/volume/<proc> <val>")
                return

            if not proc or vol is None:
                print(f"[/app/volume] invalid: {proc} {x}")
                return

            ok = set_app_volume(proc, vol)
            print(f"[/app/volume] {proc} -> {vol:.2f} {'OK' if ok else 'no session'}")

        disp.map("/ping", handle_ping)
        disp.map("/master/volume", handle_master_volume)
        disp.map("/master/mute", handle_master_mute)
        disp.map("/app/volume", handle_app_volume)
        disp.map("/app/volume/*", handle_app_volume)
        return disp

    def start(self, port: int):
        if self.running:
            return
        self.port = port
        disp = self._build_dispatcher()
        try:
            self.server = ThreadingOSCUDPServer(("0.0.0.0", self.port), disp)
        except OSError as e:
            raise RuntimeError(f"Could not bind UDP {self.port}: {e}")

        self.running = True
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()
        print(f"[READY] Listening on UDP {self.port}")
        print("Routes: /ping, /master/volume, /master/mute, /app/volume, /app/volume/<proc>")

    def stop(self):
        if not self.running:
            return
        try:
            self.server.shutdown()
            self.server.server_close()
        except Exception:
            pass
        self.server = None
        self.thread = None
        self.running = False
        print("[STOPPED] OSC server stopped.")

osc_server = OscServer()

# =============== Run at startup =================

def is_run_at_startup_enabled() -> bool:
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_RUN_PATH, 0, winreg.KEY_READ) as key:
            val, _ = winreg.QueryValueEx(key, RUN_VALUE_NAME)
            return bool(val)
    except FileNotFoundError:
        return False
    except OSError:
        return False

def set_run_at_startup(enabled: bool, port: int):
    pythonw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    exe = pythonw if os.path.exists(pythonw) else sys.executable
    script = os.path.abspath(sys.argv[0])
    cmd = f'"{exe}" "{script}" {port}'
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_RUN_PATH, 0, winreg.KEY_SET_VALUE) as key:
            if enabled:
                winreg.SetValueEx(key, RUN_VALUE_NAME, 0, winreg.REG_SZ, cmd)
            else:
                try:
                    winreg.DeleteValue(key, RUN_VALUE_NAME)
                except FileNotFoundError:
                    pass
    except FileNotFoundError:
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, REG_RUN_PATH)
        with key:
            if enabled:
                winreg.SetValueEx(key, RUN_VALUE_NAME, 0, winreg.REG_SZ, cmd)
            else:
                try:
                    winreg.DeleteValue(key, RUN_VALUE_NAME)
                except FileNotFoundError:
                    pass

# =============== Tray icon =================

def make_tray_icon_image():
    img = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rectangle((8, 24, 26, 40), fill=(255, 255, 255, 255))
    d.polygon([(26, 24), (42, 16), (42, 48), (26, 40)], fill=(255,255,255,255))
    d.arc((44, 18, 62, 46), start=300, end=60, fill=(255,255,255,220), width=3)
    return img

class TrayController:
    def __init__(self, app):
        self.app = app
        self.icon = pystray.Icon(APP_NAME, make_tray_icon_image(), APP_NAME, self._menu())

    def _menu(self):
        return pystray.Menu(
            pystray.MenuItem(lambda item: f"Status: {'Running' if osc_server.running else 'Stopped'}", lambda: None, enabled=False),
            pystray.MenuItem("Start server", self.on_start),
            pystray.MenuItem("Stop server", self.on_stop),
            pystray.MenuItem(lambda item: f"Run at startup: {'On' if is_run_at_startup_enabled() else 'Off'}", self.on_toggle_startup),
            pystray.MenuItem(lambda item: f"{'Hide' if console_mgr.console_visible else 'Show'} log console", self.on_toggle_logs),
            pystray.MenuItem("Show window", self.on_show),
            pystray.MenuItem("Quit", self.on_quit),
        )

    def on_start(self, *_):
        self.app.start_server_from_gui()
        self.icon.menu = self._menu()

    def on_stop(self, *_):
        self.app.stop_server_from_gui()
        self.icon.menu = self._menu()

    def on_toggle_startup(self, *_):
        enabled = is_run_at_startup_enabled()
        set_run_at_startup(not enabled, self.app.get_port())
        self.icon.menu = self._menu()

    def on_toggle_logs(self, *_):
        want = not console_mgr.console_visible
        console_mgr.toggle(want)
        self.icon.menu = self._menu()

    def on_show(self, *_):
        self.app.show_window()

    def on_quit(self, *_):
        try:
            osc_server.stop()
        except Exception:
            pass
        self.icon.stop()
        self.app.force_quit()

    def run(self):
        self.icon.run_detached()

# =============== GUI (Tk) =================

class App:
    def __init__(self, port_from_argv: Optional[int] = None):
        self.root = tk.Tk()
        self.root.title(APP_NAME)
        self.root.geometry("380x230")

        # Minimize → tray; Close → tray
        self.root.bind("<Unmap>", self._on_unmap)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        self.port_var = tk.StringVar(value=str(port_from_argv or 9001))
        self.status_var = tk.StringVar(value="Stopped")
        self.startup_var = tk.BooleanVar(value=is_run_at_startup_enabled())
        self.console_var = tk.BooleanVar(value=False)  # start hidden

        frm = ttk.Frame(self.root, padding=12)
        frm.pack(fill="both", expand=True)

        row = 0
        ttk.Label(frm, text="OSC UDP Port:").grid(row=row, column=0, sticky="w")
        self.port_entry = ttk.Entry(frm, textvariable=self.port_var, width=8)
        self.port_entry.grid(row=row, column=1, sticky="w")
        ttk.Button(frm, text="Start", command=self.start_server_from_gui).grid(row=row, column=2, padx=8)
        ttk.Button(frm, text="Stop", command=self.stop_server_from_gui).grid(row=row, column=3)

        row += 1
        ttk.Label(frm, text="Status:").grid(row=row, column=0, sticky="w", pady=(10, 0))
        self.status_label = ttk.Label(frm, textvariable=self.status_var, foreground="#0a84ff")
        self.status_label.grid(row=row, column=1, columnspan=3, sticky="w", pady=(10, 0))

        row += 1
        self.cb_start = ttk.Checkbutton(frm, text="Run at Windows startup", variable=self.startup_var, command=self.on_toggle_startup)
        self.cb_start.grid(row=row, column=0, columnspan=4, sticky="w", pady=(12, 0))

        row += 1
        self.cb_console = ttk.Checkbutton(frm, text="Show log console", variable=self.console_var, command=self.on_toggle_console)
        self.cb_console.grid(row=row, column=0, columnspan=4, sticky="w", pady=(6, 0))

        row += 1
        ttk.Label(frm, text="Per-app route format: /app/volume/<process>", foreground="#666").grid(row=row, column=0, columnspan=4, sticky="w", pady=(12, 0))

        row += 1
        demo = (
            "Examples:\n"
            "  /master/volume 50\n"
            "  /master/mute 1\n"
            "  /app/volume/firefox 20\n"
            "  /app/volume/chrome 75"
        )
        txt = tk.Text(frm, height=5, width=46)
        txt.insert("1.0", demo)
        txt.configure(state="disabled")
        txt.grid(row=row, column=0, columnspan=4, sticky="nsew", pady=(6, 0))

        frm.grid_rowconfigure(row, weight=1)
        frm.grid_columnconfigure(3, weight=1)

        # Tray
        self.tray = TrayController(self)
        self.tray.run()

        # Auto-start server if argv provided, then minimize to tray
        if port_from_argv is not None:
            try:
                self.start_server_from_gui()
                self.root.after(500, self.minimize_to_tray)
            except Exception as e:
                messagebox.showerror(APP_NAME, f"Failed to start server on port {port_from_argv}:\n{e}")

    def _on_unmap(self, event):
        # If user clicked minimize, Tk goes 'iconic' → send to tray
        try:
            if self.root.state() == 'iconic':
                self.minimize_to_tray()
        except Exception:
            pass

    def get_port(self) -> int:
        try:
            return int(self.port_var.get())
        except Exception:
            return 9001

    def start_server_from_gui(self):
        port = self.get_port()
        try:
            if not osc_server.running:
                osc_server.start(port)
            self.status_var.set(f"Running on UDP {port}")
            self.status_label.configure(foreground="#2e7d32")
        except Exception as e:
            self.status_var.set("Error")
            self.status_label.configure(foreground="#b00020")
            messagebox.showerror(APP_NAME, str(e))

    def stop_server_from_gui(self):
        osc_server.stop()
        self.status_var.set("Stopped")
        self.status_label.configure(foreground="#0a84ff")

    def on_toggle_startup(self):
        enabled = self.startup_var.get()
        set_run_at_startup(enabled, self.get_port())
        # Refresh tray menu label
        self.tray.icon.menu = self.tray._menu()

    def on_toggle_console(self):
        want = self.console_var.get()
        console_mgr.toggle(want)
        # Refresh tray menu label
        self.tray.icon.menu = self.tray._menu()

    def on_close(self):
        # Close button → tray
        self.minimize_to_tray()

    def minimize_to_tray(self):
        self.root.withdraw()

    def show_window(self):
        self.root.deiconify()
        self.root.after(0, self.root.lift)

    def force_quit(self):
        try:
            osc_server.stop()
        except Exception:
            pass
        self.root.quit()

    def run(self):
        self.root.mainloop()

# =============== Entry =================

def main():
    # Optional port from argv
    port_arg = None
    for a in sys.argv[1:]:
        try:
            port_arg = int(a)
            break
        except Exception:
            pass

    app = App(port_arg)
    app.run()

if __name__ == "__main__":
    main()
