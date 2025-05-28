import subprocess 
import customtkinter as CTk 
import threading 
import pystray 
from PIL import Image 
import os 
import win32com.client 
import sys

with open(r"resource\data.qlhm", "rt") as f:
    adapter_name = f.readline().strip()
    cur_theme = f.readline().strip()
    start_value = f.readline()

    print("Contents loaded.")

def format_commands(adapter_name):
    global start_cmd
    start_cmd = """
    $WirelessAdapterName = \\"{}\\"
    $connectionProfile = [Windows.Networking.Connectivity.NetworkInformation,Windows.Networking.Connectivity,ContentType=WindowsRuntime]::GetInternetConnectionProfile()
    $tetheringManager = [Windows.Networking.NetworkOperators.NetworkOperatorTetheringManager,Windows.Networking.NetworkOperators,ContentType=WindowsRuntime]::CreateFromConnectionProfile($connectionProfile)

    Start-Sleep -Seconds 2
    netsh wlan set autoconfig enabled=yes interface=$WirelessAdapterName

    $tetheringManager.StartTetheringAsync()
    Start-Sleep -Seconds 2

    netsh wlan set autoconfig enabled=no interface=$WirelessAdapterName
    """.format(adapter_name)

    global end_cmd
    end_cmd = """
    $WirelessAdapterName = \\"{}\\"
    $connectionProfile = [Windows.Networking.Connectivity.NetworkInformation,Windows.Networking.Connectivity,ContentType=WindowsRuntime]::GetInternetConnectionProfile()
    $tetheringManager = [Windows.Networking.NetworkOperators.NetworkOperatorTetheringManager,Windows.Networking.NetworkOperators,ContentType=WindowsRuntime]::CreateFromConnectionProfile($connectionProfile)

    netsh wlan set autoconfig enabled=yes interface=$WirelessAdapterName
    $tetheringManager.StopTetheringAsync()
    """.format(adapter_name)

    global scan_cmd
    scan_cmd = """
    $WirelessAdapterName = \\"{}\\"
    netsh interface show interface
    netsh interface show interface $WirelessAdapterName
    PAUSE
    """.format(adapter_name)


def run_command(cmd):
    subprocess.run(["powershell", "Start-Process", "powershell", "-ArgumentList", f"'-ExecutionPolicy Bypass -Command {cmd}'", "-Verb", "RunAs"], capture_output=True, text=True)

class Settings(CTk.CTkToplevel):
    def __init__(self, *args, **kwargs):
        # Main constructor
        super().__init__(*args, **kwargs)
        self.geometry("240x300+460+100")
        self.minsize(width=200, height=300)
        self.title("Settings")
        self.grid_columnconfigure(0, weight=1)

        self.header = CTk.CTkLabel(self, text="Settings", font=("Arial", 24, "bold")) # Creates and styles text element
        self.header.grid(row=0, column=0, padx=10, pady=10) # Positions said element

        self.label = CTk.CTkLabel(self, text="Adapter Name", fg_color=("white", "gray30"), corner_radius=6)
        self.label.grid(row=1, column=0, padx=10, pady=(10, 0))

        self.text_input = CTk.CTkEntry(self, width=64, height=32) # Creates an object for text entry
        self.text_input.insert(0, f"{adapter_name}")
        self.text_input.grid(row=2, column=0, pady=10)

        self.option_menu = CTk.CTkOptionMenu(self, values=["System", "Dark", "Light"], command=self.option_menu_callback)
        self.option_menu.grid(row=3, column=0, padx=(10, 0), pady=20)

        self.check_var = CTk.StringVar(value=start_value)

        self.start_checkbox = CTk.CTkCheckBox(self, text="Launch on start-up",
                                               variable=self.check_var, onvalue="1", offvalue="0")
        self.start_checkbox.grid(row=4, column=0)

        self.button = CTk.CTkButton(self, text="Apply & Save", command=self.save)
        self.button.grid(row=5, column=0, padx=(10, 0), pady=30)
    
    def add_to_startup(self, shortcut_name="QLHM", exe_name="qlhm_main.exe"):
        # Shortcut creation and assignment
        startup_path = os.path.join(os.getenv('APPDATA'), r'Microsoft\Windows\Start Menu\Programs\Startup')
        shortcut_path = os.path.join(startup_path, f"{shortcut_name}.lnk")

        exe_dir = os.path.dirname(sys.executable)
        exe_full_path = os.path.join(exe_dir, exe_name)

        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.TargetPath = exe_full_path
        shortcut.WorkingDirectory = exe_dir
        shortcut.IconLocation = exe_full_path
        shortcut.save()

    def remove_startup(self, shortcut_name="QLHM"):
        # Checks for and removes existing shortcuts
        startup_path = os.path.join(os.getenv('APPDATA'), r'Microsoft\Windows\Start Menu\Programs\Startup')
        shortcut_path = os.path.join(startup_path, f"{shortcut_name}.lnk")
        if os.path.exists(shortcut_path):
            os.remove(shortcut_path)
    
    def option_menu_callback(self, choice):
        CTk.set_appearance_mode(choice.lower())
        global cur_theme
        cur_theme = choice.lower()
    
    def save(self):
        print("Saving...")
        match self.check_var.get():
            case "1":
                self.add_to_startup()
            case _:
                self.remove_startup()
        global adapter_name
        adapter_name = self.text_input.get()
        with open(r"resource\data.qlhm", 'w') as f:
            f.write(f"{adapter_name}\n{cur_theme}\n{self.check_var.get()}")
        
        format_commands(self.text_input.get()) # Formats commands to avoid restart

class App(CTk.CTk):
    def __init__(self):
        super().__init__()
        CTk.set_appearance_mode(cur_theme)
        self.settings_window = None
        # Window definition
        self.title("QL Hotspot Manager") 
        self.geometry("360x420+100+100")
        self.minsize(width=240, height=400)
        self.grid_columnconfigure(0, weight=1) # Spans column 0 across window via weight

        self.protocol("WM_DELETE_WINDOW", self.to_tray)

        # Header
        self.header = CTk.CTkLabel(self, text="Quest Link\nHotspot Manager", font=("Arial", 24, "bold")) # Creates and styles text element
        self.header.grid(row=0, column=0, padx=10, pady=20) # Positions said element

        # Adapter Label
        self.label = CTk.CTkLabel(self, text="Adapter Name", fg_color=("white", "gray30"), corner_radius=6)
        self.label.grid(row=1, column=0, padx=10, pady=(10, 0))

        self.label = CTk.CTkLabel(self, text=f"{adapter_name}", fg_color=("white", "gray30"), corner_radius=6)
        self.label.grid(row=2, column=0, padx=10, pady=(10, 25))

        # Settings Button
        self.settings_button = CTk.CTkButton(self, text="Settings", fg_color="#226", hover_color="#115", command=self.open_settings_window)
        self.settings_button.grid(row=3, column=0, pady=1)

        self.device_scan_button = CTk.CTkButton(self, text="Scan for Devices", fg_color="#226", hover_color="#115", command=self.scan_devices)
        self.device_scan_button.grid(row=4, column=0, pady=10)
        # Main functions
        self.status = CTk.CTkLabel(self, text="", font=("Arial", 11))
        self.status.grid(row=5, column=0)

        self.button = CTk.CTkButton(self, text="Start Hotspot", font=("Arial", 14, "bold"), command=self.start_hotspot)
        self.button.grid(column=0, padx=50, pady=10, sticky="ew") # Expands button east+west to borders

        self.button = CTk.CTkButton(self, text="End Hotspot", font=("Arial", 14, "bold"), command=self.end_hotspot)
        self.button.grid(column=0, padx=50, pady=10, sticky="ew")
    
    def open_settings_window(self):
        if self.settings_window is None or not self.settings_window.winfo_exists():
            self.settings_window = Settings(self)
            self.settings_window.focus()
        else:
            self.settings_window.focus()

    def scan_devices(self):
        self.status.configure(text=f"Check the Terminal Window")
        run_command(scan_cmd)
        print("Check the Terminal Window")
        self.refresh_name()
    
    def start_hotspot(self):
        self.status.configure(text=f"Starting Hotspot...")
        run_command(start_cmd)
        print("Hotspot starting...")
        self.refresh_name()

    def end_hotspot(self):
        self.status.configure(text=f"Ending Hotspot...")
        run_command(end_cmd)
        print("Ending hotspot...")
        self.refresh_name()
    
    def refresh_name(self):
        self.label = CTk.CTkLabel(self, text=f"{adapter_name}", fg_color=("white", "gray30"), corner_radius=6)
        self.label.grid(row=2, column=0, padx=10, pady=(10, 25))

    # system tray functions

    def to_tray(self):
        self.withdraw()
        self.start_tray()

    def quit_app(self, icon):
        icon.stop()
        self.quit() # Fully closes all related processes

    def restore_window(self, icon):
        icon.stop()
        self.deiconify()

    def start_tray(self):
        def tray_target():
            # Contains tray building resources
            if cur_theme == "light":
                tray_img = Image.open(r"resource\light_ico.png")
            else:
                tray_img = Image.open(r"resource\dark_ico.png")

            menu = pystray.Menu(
                pystray.MenuItem("Maximise", self.restore_window),
                pystray.MenuItem("Start Hotspot", self.start_hotspot),
                pystray.MenuItem("End Hotspot", self.end_hotspot),
                pystray.MenuItem("Exit", self.quit_app)
            )
            self.tray_icon = pystray.Icon("QLHM", tray_img, "Quest Link Hotspot Manager", menu)
            self.tray_icon.run()

        self.tray_thread = threading.Thread(target=tray_target, daemon=True) # Creates a new process thread containing final build from tray_target
        self.tray_thread.start() # Runs thread


format_commands(adapter_name)

app = App()
app.mainloop()