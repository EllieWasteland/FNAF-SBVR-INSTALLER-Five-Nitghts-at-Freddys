import webview
import os
import sys
import shutil
import zipfile
import subprocess
import glob
import json

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
            # Corregido el warning: usando webview.FileDialog.FOLDER en lugar de FOLDER_DIALOG
            result = window.create_file_dialog(webview.FileDialog.FOLDER, directory="")
            if result and len(result) > 0:
                return result[0]
        except Exception:
            pass
        return None

    def verify_exe(self, folder_path):
        exe_path = os.path.join(folder_path, "fnaf9", "Binaries", "Win64", "fnaf9-Win64-Shipping.exe")
        return os.path.exists(exe_path)

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

            # Generar el archivo launcher_config.json
            config_data = {
                "lenguaje": lenguaje
            }
            config_path = os.path.join(win64_path, "launcher_config.json")
            with open(config_path, "w", encoding="utf-8") as json_file:
                json.dump(config_data, json_file, indent=4)

            shortcut_path = os.path.join(self.desktop_path, "FNAF SBVR.lnk")
            
            vbs_content = f"""Set ws = CreateObject("WScript.Shell")
Set shortcut = ws.CreateShortcut("{shortcut_path}")
shortcut.TargetPath = "{launcher_dest}"
shortcut.WorkingDirectory = "{win64_path}"
shortcut.Description = "FNAF: Security Breach VR"
shortcut.Save"""
            vbs_path = os.path.join(self.base_dir, "create_shortcut.vbs")
            with open(vbs_path, "w", encoding="utf-8") as f:
                f.write(vbs_content)
            
            subprocess.run(["cscript.exe", "//Nologo", vbs_path], creationflags=subprocess.CREATE_NO_WINDOW)
            if os.path.exists(vbs_path):
                os.remove(vbs_path)

            return "success"
            
        except Exception:
            return "error"

    def close_app(self):
        webview.windows[0].destroy()

def main():
    api = InstallApi()
    
    webview.create_window(
        title='FNAF: SBVR - Setup',
        url='index.html',
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