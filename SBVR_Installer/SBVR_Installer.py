import webview
import os
import sys
import shutil
import zipfile
import subprocess
import glob
import json
from pyshortcuts import make_shortcut  # Reemplazamos win32com.client por pyshortcuts

def get_base_path():
    if hasattr(sys, '_MEIPASS'):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))

class InstallApi:
    def __init__(self):
        self.local_app_data = os.environ.get('LOCALAPPDATA', '')
        self.app_data = os.environ.get('APPDATA', '')
        self.desktop_path = os.path.join(os.environ.get('USERPROFILE', ''), 'Desktop')
        self.base_dir = get_base_path()

    def check_saves(self):
        saves_path = os.path.join(self.local_app_data, "fnaf9", "Saved", "SaveGames")
        return os.path.exists(saves_path)

    def select_folder(self):
        try:
            window = webview.windows[0]
            result = window.create_file_dialog(webview.FileDialog.FOLDER, directory="")
            if result and len(result) > 0:
                return result[0]
        except Exception:
            pass
        return None

    def verify_exe(self, folder_path):
        folder_path = os.path.normpath(folder_path)

        # Chequeo 1: Es la carpeta base correcta que ya contiene fnaf9/Binaries...
        exe_path_A = os.path.join(folder_path, "fnaf9", "Binaries", "Win64", "fnaf9-Win64-Shipping.exe")
        if os.path.exists(exe_path_A):
            return folder_path

        # Chequeo 2: El usuario seleccionó la carpeta Win64 directamente
        exe_path_B = os.path.join(folder_path, "fnaf9-Win64-Shipping.exe")
        if os.path.exists(exe_path_B):
            base_path = os.path.dirname(os.path.dirname(os.path.dirname(folder_path)))
            return base_path

        # Chequeo 3: El usuario seleccionó la carpeta Binaries
        exe_path_C = os.path.join(folder_path, "Win64", "fnaf9-Win64-Shipping.exe")
        if os.path.exists(exe_path_C):
            base_path = os.path.dirname(os.path.dirname(folder_path))
            return base_path

        # Chequeo 4: El usuario seleccionó la carpeta fnaf9
        exe_path_D = os.path.join(folder_path, "Binaries", "Win64", "fnaf9-Win64-Shipping.exe")
        if os.path.exists(exe_path_D):
            base_path = os.path.dirname(folder_path)
            return base_path

        return False

    def get_common_steam_paths(self):
        # Rutas comunes en diferentes unidades posibles (Buscando siempre la carpeta Quarters)
        paths = [
            r"C:\Program Files (x86)\Steam\steamapps\common\Quarters",
            r"C:\Program Files\Steam\steamapps\common\Quarters",
            r"D:\SteamLibrary\steamapps\common\Quarters",
            r"E:\SteamLibrary\steamapps\common\Quarters",
            r"F:\SteamLibrary\steamapps\common\Quarters",
        ]
        
        # Intentar obtener la ruta principal de Steam mediante el registro de Windows
        try:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Valve\Steam")
            steam_path = winreg.QueryValueEx(key, "InstallPath")[0]
            paths.insert(0, os.path.join(steam_path, "steamapps", "common", "Quarters"))
        except Exception:
            pass
            
        return paths

    def auto_find_dir(self):
        for p in self.get_common_steam_paths():
            base = self.verify_exe(p)
            if base:
                return base
        return None

    def install(self, folder_path, lenguaje="en"):
        try:
            if not self.check_saves():
                return "error"

            win64_path = os.path.join(folder_path, "fnaf9", "Binaries", "Win64")
            paks_path = os.path.join(folder_path, "fnaf9", "Content", "Paks")
            uevr_profile_dir = os.path.join(self.app_data, "UnrealVRMod")
            saves_dest_path = os.path.join(self.local_app_data, "fnaf9", "Saved", "SaveGames")

            os.makedirs(uevr_profile_dir, exist_ok=True)
            target_zip = os.path.join(self.base_dir, "fnaf9-Win64-Shipping.zip")
            if os.path.exists(target_zip):
                with zipfile.ZipFile(target_zip, 'r') as zip_ref:
                    zip_ref.extractall(uevr_profile_dir)

            sav_files = glob.glob(os.path.join(self.base_dir, "*.sav"))
            for sav_file in sav_files:
                shutil.copy2(sav_file, saves_dest_path)

            os.makedirs(paks_path, exist_ok=True)
            pak_files = glob.glob(os.path.join(self.base_dir, "*.pak"))
            for pak_file in pak_files:
                shutil.copy2(pak_file, paks_path)

            os.makedirs(win64_path, exist_ok=True)
            launcher_src = os.path.join(self.base_dir, "SBVR_Launcher.exe")
            launcher_dest = os.path.join(win64_path, "SBVR_Launcher.exe")
            if os.path.exists(launcher_src):
                shutil.copy2(launcher_src, launcher_dest)

            config_data = {
                "lenguaje": lenguaje
            }
            config_path = os.path.join(win64_path, "launcher_config.json")
            with open(config_path, "w", encoding="utf-8") as json_file:
                json.dump(config_data, json_file, indent=4)

            try:
                icon_path = os.path.join(self.base_dir, "logo.ico")
                if not os.path.exists(icon_path):
                    icon_path = launcher_dest

                make_shortcut(
                    script=launcher_dest,         # El archivo base a ejecutar
                    executable=launcher_dest,     # Forzamos el ejecutable en el acceso directo
                    name='FNAF: Security Breach VR',
                    description='FNAF: Security Breach VR',
                    icon=icon_path,               # Ruta al ícono
                    terminal=False,    
                    desktop=True,      
                    startmenu=True     
                )
                print("Acceso directo creado exitosamente con pyshortcuts.")
            except Exception as e:
                print(f"No se pudo crear el acceso directo: {e}")

            return "success"
            
        except Exception as e:
            print(f"Error durante la instalación: {e}")
            return "error"

    def close_app(self):
        webview.windows[0].destroy()

def main():
    api = InstallApi()
    
    webview.create_window(
        title='FNAF: SBVR - Setup',
        url='SBVR_Installer.html',
        js_api=api,
        width=1080,
        height=720,
        frameless=True,
        easy_drag=False,
        background_color='#EAE8E3',
        transparent=False
    )
    
    webview.start(debug=False)

if __name__ == '__main__':
    main()