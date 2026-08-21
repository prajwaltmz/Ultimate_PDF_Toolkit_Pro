"""
Ultimate PDF Toolkit Pro - Graphical Setup Wizard
A standalone Windows Installer GUI that installs the application, creates shortcuts,
and registers in Windows Add/Remove Programs.
"""

import os
import sys
import shutil
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import winreg

# DPI awareness
try:
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

APP_NAME = "Ultimate PDF Toolkit Pro"
REG_KEY = "UltimatePDFToolkitPro"
EXE_NAME = "Ultimate_PDF_Toolkit_Pro.exe"
DEFAULT_INSTALL_DIR = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "Programs", APP_NAME)

def get_bundle_dir():
    """Get directory where script or PyInstaller bundle resides."""
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS if hasattr(sys, '_MEIPASS') else os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def get_source_exe():
    """Find the source executable to install."""
    bundle_dir = get_bundle_dir()
    exe_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
    
    candidates = [
        os.path.join(bundle_dir, EXE_NAME),
        os.path.join(exe_dir, EXE_NAME),
        os.path.join(exe_dir, "Package", EXE_NAME),
        os.path.join(exe_dir, "..", "Package", EXE_NAME),
        os.path.join(exe_dir, "dist", EXE_NAME),
        os.path.join(exe_dir, "..", "dist", EXE_NAME)
    ]
    for c in candidates:
        if os.path.exists(c):
            return os.path.abspath(c)
    return None

def create_shortcut(target, shortcut_path, description="", work_dir=""):
    """Create a Windows .lnk shortcut using WScript.Shell COM object."""
    try:
        import win32com.client
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(shortcut_path)
        shortcut.TargetPath = target
        shortcut.WorkingDirectory = work_dir or os.path.dirname(target)
        shortcut.Description = description
        shortcut.Save()
        return True
    except Exception:
        # Fallback to powershell shortcut creation
        ps_cmd = f'''$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut("{shortcut_path}"); $s.TargetPath = "{target}"; $s.WorkingDirectory = "{work_dir or os.path.dirname(target)}"; $s.Description = "{description}"; $s.Save()'''
        try:
            subprocess.run(["powershell.exe", "-NoProfile", "-Command", ps_cmd], check=True, capture_output=True)
            return True
        except Exception:
            return False

class InstallerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} Setup")
        self.geometry("620x460")
        self.resizable(False, False)
        self.configure(bg="#1e1e1e")
        
        # Center window
        self.eval('tk::PlaceWindow . center')
        
        self.install_path_var = tk.StringVar(value=DEFAULT_INSTALL_DIR)
        self.create_desktop_var = tk.BooleanVar(value=True)
        self.create_startmenu_var = tk.BooleanVar(value=True)
        self.register_windows_var = tk.BooleanVar(value=True)
        self.launch_after_var = tk.BooleanVar(value=True)
        
        self.installed_exe_path = ""
        
        self.init_styles()
        self.create_header()
        
        self.container = tk.Frame(self, bg="#1e1e1e")
        self.container.pack(fill=tk.BOTH, expand=True, padx=25, pady=15)
        
        self.show_welcome_screen()

    def init_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TProgressbar", thickness=18, troughcolor="#2d2d2d", background="#007acc")
        style.configure("Install.TButton", background="#007acc", foreground="white", font=("Segoe UI", 11, "bold"), padding=8)
        style.map("Install.TButton", background=[("active", "#005999")])
        style.configure("Browse.TButton", background="#3c3f41", foreground="white", font=("Segoe UI", 9), padding=4)
        style.map("Browse.TButton", background=[("active", "#4c5052")])

    def create_header(self):
        header_frame = tk.Frame(self, bg="#252526", height=70)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        lbl_title = tk.Label(header_frame, text=APP_NAME, font=("Segoe UI", 16, "bold"), bg="#252526", fg="white")
        lbl_title.pack(anchor="w", padx=25, pady=(12, 2))
        
        lbl_sub = tk.Label(header_frame, text="Windows Setup & Installation Wizard", font=("Segoe UI", 9), bg="#252526", fg="#9cdcfe")
        lbl_sub.pack(anchor="w", padx=25)

    def show_welcome_screen(self):
        for widget in self.container.winfo_children():
            widget.destroy()
            
        lbl_desc = tk.Label(self.container, text="This wizard will install Ultimate PDF Toolkit Pro on your computer.\nChoose your installation preferences below:", font=("Segoe UI", 10), bg="#1e1e1e", fg="#cccccc", justify="left")
        lbl_desc.pack(anchor="w", pady=(0, 15))
        
        # Directory Selection
        dir_frame = tk.LabelFrame(self.container, text=" Installation Folder ", font=("Segoe UI", 9, "bold"), bg="#1e1e1e", fg="#007acc", padx=10, pady=10)
        dir_frame.pack(fill=tk.X, pady=(0, 15))
        
        entry_dir = tk.Entry(dir_frame, textvariable=self.install_path_var, font=("Segoe UI", 10), bg="#2d2d2d", fg="white", insertbackground="white", borderwidth=1, relief="solid")
        entry_dir.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10), ipady=3)
        
        btn_browse = ttk.Button(dir_frame, text="Browse...", style="Browse.TButton", command=self.browse_folder)
        btn_browse.pack(side=tk.RIGHT)
        
        # Options Frame
        opt_frame = tk.LabelFrame(self.container, text=" Shortcut & System Options ", font=("Segoe UI", 9, "bold"), bg="#1e1e1e", fg="#007acc", padx=10, pady=10)
        opt_frame.pack(fill=tk.X, pady=(0, 20))
        
        chk1 = tk.Checkbutton(opt_frame, text="Create Desktop Shortcut", variable=self.create_desktop_var, bg="#1e1e1e", fg="white", selectcolor="#2d2d2d", activebackground="#1e1e1e", activeforeground="white", font=("Segoe UI", 9))
        chk1.pack(anchor="w", pady=2)
        
        chk2 = tk.Checkbutton(opt_frame, text="Create Start Menu Entry (Searchable in Windows Search)", variable=self.create_startmenu_var, bg="#1e1e1e", fg="white", selectcolor="#2d2d2d", activebackground="#1e1e1e", activeforeground="white", font=("Segoe UI", 9))
        chk2.pack(anchor="w", pady=2)
        
        chk3 = tk.Checkbutton(opt_frame, text="Register in Windows 'Installed Apps' / Add or Remove Programs", variable=self.register_windows_var, bg="#1e1e1e", fg="white", selectcolor="#2d2d2d", activebackground="#1e1e1e", activeforeground="white", font=("Segoe UI", 9))
        chk3.pack(anchor="w", pady=2)
        
        # Bottom Buttons
        btn_frame = tk.Frame(self.container, bg="#1e1e1e")
        btn_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        btn_cancel = tk.Button(btn_frame, text="Cancel", font=("Segoe UI", 10), bg="#3c3f41", fg="white", activebackground="#505050", activeforeground="white", borderwidth=0, padx=15, pady=6, command=self.destroy)
        btn_cancel.pack(side=tk.RIGHT, padx=(10, 0))
        
        btn_install = tk.Button(btn_frame, text="  Install Now  ", font=("Segoe UI", 10, "bold"), bg="#007acc", fg="white", activebackground="#005999", activeforeground="white", borderwidth=0, padx=20, pady=6, command=self.start_installation)
        btn_install.pack(side=tk.RIGHT)

    def browse_folder(self):
        folder = filedialog.askdirectory(initialdir=self.install_path_var.get())
        if folder:
            self.install_path_var.set(os.path.join(folder, APP_NAME))

    def start_installation(self):
        target_dir = self.install_path_var.get().strip()
        if not target_dir:
            messagebox.showerror("Error", "Please specify a valid installation folder.")
            return
            
        src_exe = get_source_exe()
        if not src_exe or not os.path.exists(src_exe):
            messagebox.showerror("Error", f"Could not locate '{EXE_NAME}'. Please make sure the executable is present.")
            return

        for widget in self.container.winfo_children():
            widget.destroy()
            
        self.lbl_status = tk.Label(self.container, text="Preparing installation...", font=("Segoe UI", 11), bg="#1e1e1e", fg="white")
        self.lbl_status.pack(anchor="w", pady=(20, 10))
        
        self.progress = ttk.Progressbar(self.container, orient="horizontal", mode="determinate", style="TProgressbar", length=570)
        self.progress.pack(pady=10)
        
        self.lbl_details = tk.Label(self.container, text="", font=("Segoe UI", 9), bg="#1e1e1e", fg="#888888")
        self.lbl_details.pack(anchor="w", pady=(5, 20))
        
        threading.Thread(target=self.run_install_worker, args=(src_exe, target_dir), daemon=True).start()

    def update_progress(self, percent, status_text, detail_text=""):
        self.progress['value'] = percent
        self.lbl_status.config(text=status_text)
        self.lbl_details.config(text=detail_text)
        self.update_idletasks()

    def run_install_worker(self, src_exe, target_dir):
        try:
            # Step 1: Create Directory
            self.update_progress(15, "Creating application directory...", target_dir)
            os.makedirs(target_dir, exist_ok=True)
            
            # Step 2: Copy executable
            self.update_progress(35, "Copying application files...", f"Installing {EXE_NAME}")
            dest_exe = os.path.join(target_dir, EXE_NAME)
            
            # If destination exists and might be running, try to close it or overwrite
            try:
                subprocess.run(["taskkill", "/F", "/IM", EXE_NAME], capture_output=True)
            except Exception:
                pass
                
            shutil.copy2(src_exe, dest_exe)
            self.installed_exe_path = dest_exe
            
            # Step 3: Create uninstaller script
            self.update_progress(55, "Generating uninstaller...", "Creating clean removal scripts")
            self.create_uninstaller_scripts(target_dir)
            
            # Step 4: Shortcuts
            if self.create_desktop_var.get():
                self.update_progress(70, "Creating Desktop shortcut...", "Adding shortcut to your Desktop")
                desktop_dir = os.path.join(os.environ.get("USERPROFILE", ""), "Desktop")
                if os.path.exists(desktop_dir):
                    create_shortcut(dest_exe, os.path.join(desktop_dir, f"{APP_NAME}.lnk"), description=APP_NAME, work_dir=target_dir)
                    
            if self.create_startmenu_var.get():
                self.update_progress(85, "Creating Start Menu entries...", "Registering with Windows Start Menu & Search")
                appdata = os.environ.get("APPDATA", "")
                if appdata:
                    start_folder = os.path.join(appdata, "Microsoft", "Windows", "Start Menu", "Programs", APP_NAME)
                    os.makedirs(start_folder, exist_ok=True)
                    create_shortcut(dest_exe, os.path.join(start_folder, f"{APP_NAME}.lnk"), description=APP_NAME, work_dir=target_dir)
                    uninst_bat = os.path.join(target_dir, "Uninstall_App.bat")
                    create_shortcut(uninst_bat, os.path.join(start_folder, f"Uninstall {APP_NAME}.lnk"), description=f"Uninstall {APP_NAME}", work_dir=target_dir)
            
            # Step 5: Windows Registry
            if self.register_windows_var.get():
                self.update_progress(95, "Registering in Windows Settings...", "Adding to Add/Remove Programs")
                self.register_in_windows(target_dir, dest_exe)
                
            self.update_progress(100, "Installation Complete!", "Ready to launch")
            self.after(500, self.show_completed_screen)
            
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Installation Error", f"An error occurred during installation:\n{e}"))
            self.after(0, self.show_welcome_screen)

    def create_uninstaller_scripts(self, target_dir):
        lines = [
            '$ErrorActionPreference = "SilentlyContinue"',
            'Write-Host "Closing running instances..." -ForegroundColor Yellow',
            'Stop-Process -Name "Ultimate_PDF_Toolkit_Pro" -Force -ErrorAction SilentlyContinue',
            'Start-Sleep -Seconds 1',
            '',
            '# Desktop shortcut',
            f'$d = Join-Path ([Environment]::GetFolderPath("Desktop")) "{APP_NAME}.lnk"',
            'if (Test-Path -LiteralPath $d) { Remove-Item -LiteralPath $d -Force }',
            '',
            '# Start Menu',
            f'$s = Join-Path "$env:APPDATA\\Microsoft\\Windows\\Start Menu\\Programs" "{APP_NAME}"',
            'if (Test-Path -LiteralPath $s) { Remove-Item -LiteralPath $s -Recurse -Force }',
            '',
            '# Registry',
            f'Remove-Item -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{REG_KEY}" -Recurse -Force -ErrorAction SilentlyContinue',
            '',
            '# Target EXE',
            f'$exe = Join-Path "{target_dir.replace("\\", "/")}" "{EXE_NAME}"',
            'if (Test-Path -LiteralPath $exe) { Remove-Item -LiteralPath $exe -Force -ErrorAction SilentlyContinue }',
            '',
            f'Write-Host "{APP_NAME} was successfully uninstalled." -ForegroundColor Green',
            ''
        ]
        uninst_ps1 = "\r\n".join(lines)
        with open(os.path.join(target_dir, "uninstall.ps1"), "w", encoding="utf-8") as f:
            f.write(uninst_ps1)
            
        bat_lines = [
            "@echo off",
            f"title {APP_NAME} - Uninstaller",
            'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall.ps1"',
            "pause",
            ""
        ]
        uninst_bat = "\r\n".join(bat_lines)
        with open(os.path.join(target_dir, "Uninstall_App.bat"), "w", encoding="ascii") as f:
            f.write(uninst_bat)

    def register_in_windows(self, target_dir, dest_exe):
        try:
            key_path = rf"Software\Microsoft\Windows\CurrentVersion\Uninstall\{REG_KEY}"
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                winreg.SetValueEx(key, "DisplayName", 0, winreg.REG_SZ, APP_NAME)
                winreg.SetValueEx(key, "DisplayVersion", 0, winreg.REG_SZ, "1.0.0")
                winreg.SetValueEx(key, "Publisher", 0, winreg.REG_SZ, "Ultimate PDF Toolkit")
                winreg.SetValueEx(key, "DisplayIcon", 0, winreg.REG_SZ, f'"{dest_exe}",0')
                winreg.SetValueEx(key, "InstallLocation", 0, winreg.REG_SZ, target_dir)
                winreg.SetValueEx(key, "UninstallString", 0, winreg.REG_SZ, f'"{os.path.join(target_dir, "Uninstall_App.bat")}"')
                winreg.SetValueEx(key, "NoModify", 0, winreg.REG_DWORD, 1)
                winreg.SetValueEx(key, "NoRepair", 0, winreg.REG_DWORD, 1)
                try:
                    size_kb = int(os.path.getsize(dest_exe) / 1024)
                    winreg.SetValueEx(key, "EstimatedSize", 0, winreg.REG_DWORD, size_kb)
                except Exception:
                    pass
        except Exception as e:
            print(f"Registry registration error: {e}")

    def show_completed_screen(self):
        for widget in self.container.winfo_children():
            widget.destroy()
            
        lbl_success = tk.Label(self.container, text="🎉 Installation Complete!", font=("Segoe UI", 16, "bold"), bg="#1e1e1e", fg="#4EC9B0")
        lbl_success.pack(anchor="w", pady=(10, 10))
        
        lbl_info = tk.Label(self.container, text=f"{APP_NAME} has been successfully installed on your computer.\n\nYou can launch it anytime from your Desktop, Windows Start Menu, or search bar.", font=("Segoe UI", 10), bg="#1e1e1e", fg="#dcdcdc", justify="left")
        lbl_info.pack(anchor="w", pady=(0, 20))
        
        chk_launch = tk.Checkbutton(self.container, text=f"Launch {APP_NAME} now", variable=self.launch_after_var, bg="#1e1e1e", fg="white", selectcolor="#2d2d2d", activebackground="#1e1e1e", activeforeground="white", font=("Segoe UI", 10, "bold"))
        chk_launch.pack(anchor="w", pady=(0, 25))
        
        btn_finish = tk.Button(self.container, text="   Finish   ", font=("Segoe UI", 11, "bold"), bg="#007acc", fg="white", activebackground="#005999", activeforeground="white", borderwidth=0, padx=25, pady=7, command=self.on_finish)
        btn_finish.pack(side=tk.RIGHT, pady=(10, 0))

    def on_finish(self):
        if self.launch_after_var.get() and self.installed_exe_path and os.path.exists(self.installed_exe_path):
            try:
                subprocess.Popen([self.installed_exe_path], cwd=os.path.dirname(self.installed_exe_path))
            except Exception as e:
                messagebox.showerror("Error", f"Failed to launch app: {e}")
        self.destroy()

if __name__ == "__main__":
    app = InstallerApp()
    app.mainloop()
