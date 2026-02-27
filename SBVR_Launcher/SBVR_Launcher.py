import os
import sys
import time
import ctypes
import threading
import subprocess
import psutil
import webview
import json
from pyinjector import inject

JUEGO_EJECUTABLE = "fnaf9-Win64-Shipping.exe"
RUTA_UEVR_BACKEND = os.path.join("UEVR", "UEVRBackend.dll")
RUTA_UEVR_LOADER  = os.path.join("UEVR", "openxr_loader.dll")

def obtener_pid(nombre_ejecutable):
    for proceso in psutil.process_iter(['pid', 'name']):
        try:
            if proceso.info['name'].lower() == nombre_ejecutable.lower():
                return proceso.info['pid']
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return None

def cargar_configuracion(directorio):
    config_path = os.path.join(directorio, "launcher_config.json")
    if os.path.exists(config_path):
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data.get("lenguaje", "es").lower()
        except Exception as e:
            print(f"Error leyendo config: {e}")
    return "es"

class LauncherAPI:
    def __init__(self, lang="es"):
        self._hwnd = None
        self.lang = lang

    def get_language(self):
        return self.lang

    class _POINT(ctypes.Structure):
        _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

    class _RECT(ctypes.Structure):
        _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                    ("right", ctypes.c_long), ("bottom", ctypes.c_long)]

    def _get_hwnd(self):
        if self._hwnd:
            return self._hwnd
        hwnd = ctypes.windll.user32.FindWindowW(None, "FNAF9 VR Launcher")
        if hwnd:
            self._hwnd = hwnd
        return hwnd

    def start_drag(self):
        hwnd = self._get_hwnd()
        if not hwnd:
            return

        user32 = ctypes.windll.user32
        pt = self._POINT()
        user32.GetCursorPos(ctypes.byref(pt))
        start_mouse_x, start_mouse_y = pt.x, pt.y

        rc = self._RECT()
        user32.GetWindowRect(hwnd, ctypes.byref(rc))
        start_win_x = rc.left
        start_win_y = rc.top
        win_w = rc.right - rc.left
        win_h = rc.bottom - rc.top

        def _drag_loop():
            while user32.GetAsyncKeyState(0x01) & 0x8000:
                user32.GetCursorPos(ctypes.byref(pt))
                new_x = start_win_x + (pt.x - start_mouse_x)
                new_y = start_win_y + (pt.y - start_mouse_y)
                user32.MoveWindow(hwnd, new_x, new_y, win_w, win_h, False)
                time.sleep(0.008)

        threading.Thread(target=_drag_loop, daemon=True).start()

    def _wait_and_close(self, process):
        process.wait()
        self.close_app()

    def launch_vr(self):
        # Traducciones para errores en Python
        is_en = self.lang == 'en'
        msg_err_open = f"Error opening game: {{}}" if is_en else f"Error al abrir el juego: {{}}"
        hint_err_open = f"Verify that '{JUEGO_EJECUTABLE}' is next to the launcher." if is_en else f"Verifica que el ejecutable '{JUEGO_EJECUTABLE}' esté junto al launcher."
        msg_no_uevr = "UEVRBackend.dll not found" if is_en else "No se encontró UEVRBackend.dll"
        hint_no_uevr = "Make sure the 'UEVR' folder exists next to the launcher." if is_en else "Asegúrate de que la carpeta 'UEVR' exista junto al launcher."
        msg_inject = "Injection error: {}" if is_en else "Error de inyección: {}"
        hint_inject = "Verify that your antivirus is not blocking the injection." if is_en else "Verifica que tu antivirus no esté bloqueando la inyección."
        msg_no_pid = "Could not detect game process." if is_en else "No se pudo detectar el proceso del juego."
        hint_no_pid = "The game took too long to start or closed." if is_en else "El juego tardó demasiado en iniciar o se cerró."

        try:
            game_process = subprocess.Popen([JUEGO_EJECUTABLE])
        except Exception as e:
            return {"status": "error", "message": msg_err_open.format(e), "hint": hint_err_open}

        time.sleep(2)
        pid = obtener_pid(JUEGO_EJECUTABLE)
        if not pid:
            time.sleep(3)
            pid = obtener_pid(JUEGO_EJECUTABLE)

        if pid:
            try:
                if os.path.exists(RUTA_UEVR_LOADER):
                    inject(pid, RUTA_UEVR_LOADER)
                    time.sleep(0.5)
                if os.path.exists(RUTA_UEVR_BACKEND):
                    inject(pid, RUTA_UEVR_BACKEND)
                    threading.Thread(target=self._wait_and_close, args=(game_process,), daemon=True).start()
                    return {"status": "success"}
                else:
                    return {"status": "error", "message": msg_no_uevr, "hint": hint_no_uevr}
            except Exception as e:
                return {"status": "error", "message": msg_inject.format(e), "hint": hint_inject}
        else:
            return {"status": "error", "message": msg_no_pid, "hint": hint_no_pid}

    def launch_flat(self):
        is_en = self.lang == 'en'
        msg_err_open = f"Error opening game: {{}}" if is_en else f"Error al abrir el juego: {{}}"
        hint_err_open = f"Verify that '{JUEGO_EJECUTABLE}' is next to the launcher." if is_en else f"Verifica que el ejecutable '{JUEGO_EJECUTABLE}' esté junto al launcher."

        try:
            game_process = subprocess.Popen([JUEGO_EJECUTABLE])
            threading.Thread(target=self._wait_and_close, args=(game_process,), daemon=True).start()
            return {"status": "success"}
        except Exception as e:
            return {"status": "error", "message": msg_err_open.format(e), "hint": hint_err_open}

    def close_app(self):
        if webview.windows:
            webview.windows[0].destroy()
        sys.exit(0)

if __name__ == "__main__":
    if getattr(sys, 'frozen', False):
        base_dir = sys._MEIPASS
        exe_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        exe_dir = base_dir

    os.chdir(exe_dir)
    
    # Cargar lenguaje desde la configuración
    lenguaje_config = cargar_configuracion(exe_dir)

    api = LauncherAPI(lang=lenguaje_config)
    html_path = os.path.join(base_dir, "launcherUI.html")

    window = webview.create_window(
        title='FNAF9 VR Launcher',
        url=html_path,
        js_api=api,
        width=400,
        height=600,
        frameless=True,
        easy_drag=False,
        transparent=False
    )

    def on_shown():
        import time as _t
        _t.sleep(0.3)
        api._get_hwnd()

    window.events.shown += on_shown
    webview.start()