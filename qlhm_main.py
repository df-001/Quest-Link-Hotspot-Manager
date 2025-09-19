import subprocess
import re
import threading
import os 
import sys

import customtkinter as CTk
import pystray
from PIL import Image
import win32com.client
import darkdetect
from windows_toasts import Toast, ToastDisplayImage, WindowsToaster


# reads data in from data.qlhm
with open(r"resource\data.qlhm", "rt") as f:
    adapter_name = f.readline().strip()
    cur_theme = f.readline().strip()
    start_value = f.readline()


def format_commands(adapter_name):
    """ Defines each command with the current variable for adapter name. """
    global start_cmd
    start_cmd = """
    $WirelessAdapterName = \\"{}\\"
    $connectionProfile = [Windows.Networking.Connectivity.NetworkInformation,Windows.Networking.Connectivity,ContentType=WindowsRuntime]::GetInternetConnectionProfile()
    $tetheringManager = [Windows.Networking.NetworkOperators.NetworkOperatorTetheringManager,Windows.Networking.NetworkOperators,ContentType=WindowsRuntime]::CreateFromConnectionProfile($connectionProfile)

    Start-Sleep -Seconds 1
    netsh wlan set autoconfig enabled=yes interface=$WirelessAdapterName

    $tetheringManager.StartTetheringAsync()
    Start-Sleep -Seconds 2

    netsh wlan set autoconfig enabled=no interface=$WirelessAdapterName
    exit
    """.format(adapter_name)

    global end_cmd
    end_cmd = """
    $WirelessAdapterName = \\"{}\\"
    $connectionProfile = [Windows.Networking.Connectivity.NetworkInformation,Windows.Networking.Connectivity,ContentType=WindowsRuntime]::GetInternetConnectionProfile()
    $tetheringManager = [Windows.Networking.NetworkOperators.NetworkOperatorTetheringManager,Windows.Networking.NetworkOperators,ContentType=WindowsRuntime]::CreateFromConnectionProfile($connectionProfile)

    netsh wlan set autoconfig enabled=yes interface=$WirelessAdapterName
    $tetheringManager.StopTetheringAsync()
    exit
    """.format(adapter_name)

    global scan_cmd
    scan_cmd = """
    $WirelessAdapterName = \"{}\"
    netsh interface show interface
    netsh interface show interface $WirelessAdapterName
    """.format(adapter_name)


def run_command(cmd):
    """ Runs the inputted command in a background PS process and returns the output. """
    result = subprocess.run(
    ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", cmd],
    capture_output=True,
    text=True
    )
    return result.stdout


class App(CTk.CTk):
    def __init__(self):
        """ 
        Runs on class instantiation:
        Defines layout and widget attributes
        If ran with "-min" argument, launches in tray instead
        """
        super().__init__()
        CTk.set_appearance_mode(cur_theme)
        self.settings_window = None
        # window definition
        self.title("QL Hotspot Manager") 
        self.geometry("360x420+100+100")
        self.minsize(240, 400)
        self.maxsize(380, 420)
        self.grid_columnconfigure(0, weight=1)

        self.protocol("WM_DELETE_WINDOW", self.to_tray) # minimizes app to tray on closure


        # header
        self.header = CTk.CTkLabel(self, text="Quest Link\nHotspot Manager", font=("Arial", 24, "bold")) # creates and styles text label
        self.header.grid(row=0, column=0, padx=10, pady=20) # positions label on grid


        # adapter label
        self.label = CTk.CTkLabel(self, text="Adapter Name", fg_color=("white", "gray30"), corner_radius=6)
        self.label.grid(row=1, column=0, padx=10, pady=(10, 0))

        self.label = CTk.CTkLabel(self, text=f"{adapter_name}", fg_color=("white", "gray30"), corner_radius=6)
        self.label.grid(row=2, column=0, padx=10, pady=(10, 25))


        # settings button
        self.settings_button = CTk.CTkButton(self, text="Settings", fg_color="#226", hover_color="#115", command=self.open_settings_window) # creates button to open settings window
        self.settings_button.grid(row=3, column=0, pady=1)


        # command buttons
        self.device_scan_button = CTk.CTkButton(self, text="Scan for Devices", fg_color="#226", hover_color="#115", command=self.scan_devices)
        self.device_scan_button.grid(row=4, column=0, pady=10)

        self.status = CTk.CTkLabel(self, text="", font=("Arial", 12)) # defines adaptive status label
        self.status.grid(row=5, column=0)

        self.button = CTk.CTkButton(self, text="Start Hotspot", font=("Arial", 14, "bold"), command=self.start_hotspot)
        self.button.grid(column=0, padx=50, pady=10, sticky="ew") # expands button east+west to borders

        self.button = CTk.CTkButton(self, text="End Hotspot", font=("Arial", 14, "bold"), command=self.end_hotspot)
        self.button.grid(column=0, padx=50, pady=10, sticky="ew")

        # minimizes program if stated within argument
        try:
            match sys.argv[1]:
                case "-min":
                    self.to_tray()
        except: # normal run condition
            pass
        
        # creates shortcut in working directory if not already present
        cwd = os.getcwd()
        shortcut_path = os.path.join(cwd, f"Quest Link Hotspot Manager.lnk")

        exe_dir = os.path.dirname(sys.executable)
        exe_full_path = os.path.join(exe_dir, "qlhm_main.exe")

        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.TargetPath = exe_full_path
        shortcut.WorkingDirectory = exe_dir
        shortcut.IconLocation = exe_full_path
        shortcut.save()

    # methods for main buttons
    def open_settings_window(self):
        """ Opens the settings Top Level Window if not already open. """
        if self.settings_window is None or not self.settings_window.winfo_exists():
            self.settings_window = Settings(self)
            self.settings_window.focus()
        else:
            self.settings_window.focus()

    def scan_devices(self):
        """ Parses the values from PS and outputs the available network adapters. """
        self.status.configure(text=f"Remember to apply any changes.")
        output = run_command(scan_cmd)
        # parses command output and creates readable list of available adapters

        if output.endswith(".\n\n") == True: # triggers if adapter isn't found
            formatted_str = "Please pick the device from the list:\n"
            parsed_output = re.split(" {2}|\n", output)
            while "" in parsed_output:
                parsed_output.remove("")
                adapter_count = int((len(parsed_output) - 6) / 4)
                adapter_list = []
            for i in range(adapter_count):
                adapter_list.append(parsed_output[(i*4) + 8])
                formatted_str += adapter_list[i] + "\n"

            DialogueBox(self, formatted_str, title="Couldn't find Adapter")

        else: # triggers if adapter is found
            formatted_str = "List of adapters:\n"
            parsed_output = re.split(" {2}|\n", output)
            while "" in parsed_output:
                parsed_output.remove("")
                adapter_count = int((len(parsed_output) - 5) / 4) - 2
                adapter_list = []
            for i in range(adapter_count):
                adapter_list.append(parsed_output[(i*4) + 8])
                formatted_str += adapter_list[i] + "\n"

            DialogueBox(self, formatted_str, title="Found Adapter!")
   
        self.refresh_name()
    
    def start_hotspot(self):
        """ Runs the PS command to start the hotspot. """
        self.status.configure(text=f"Hotspot Running.")
        run_command(start_cmd)
        self.create_toast("Hotspot Started.")
        self.refresh_name()

    def end_hotspot(self):
        """ Runs the PS command to end the hotspot. """
        self.status.configure(text=f"Hotspot Ended.")
        run_command(end_cmd)
        self.create_toast("Hotspot Ended.")
        self.refresh_name()
    
    # other methods

    def refresh_name(self):
        """ Updates the Adapter Name value in the main app. """
        self.label = CTk.CTkLabel(self, text=f" "*64) # clears old name off screen
        self.label.grid(row=2, column=0, padx=10, pady=(10, 25))
        self.label = CTk.CTkLabel(self, text=f"{adapter_name}", fg_color=("white", "gray30"), corner_radius=6) # redefines new name
        self.label.grid(row=2, column=0, padx=10, pady=(10, 25))

    def create_toast(self, text):
        """ Creates a Windows notification with custom text. """
        toaster = WindowsToaster('Quest Link Hotspot Manager')

        newToast = Toast() # creates Toast object
        newToast.text_fields = [text] # sets inner text
        newToast.AddImage(ToastDisplayImage.fromPath('C:/Windows/System32/@WLOGO_96x96.png')) # sets logo to display on left

        toaster.show_toast(newToast) # sends toast to user

    # system tray methods

    def to_tray(self):
        """ Minimizes main application and starts the tray app. """
        self.withdraw()
        self.start_tray()

    def quit_app(self, icon):
        """ Closes tray app and all related processes. """
        icon.stop()
        self.quit()

    def restore_window(self, icon):
        """ Re-opens full size window. """
        icon.stop()
        self.deiconify()

    def start_tray(self):
        """ Defines and starts the tray application. """
        def tray_target():
            # sets tray building resources
            if darkdetect.theme().lower() == "dark":
                tray_img = Image.open(r"resource\dark_ico.png")
            else:
                tray_img = Image.open(r"resource\light_ico.png")

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


class Settings(CTk.CTkToplevel):
    def __init__(self, *args, **kwargs):
        """
        Runs on class instantiation:
        Defines layout and widget attributes
        """
        super().__init__(*args, **kwargs)
        # retrieves size and pos of parent window
        parent_x, parent_y = app.winfo_x(), app.winfo_y()
        parent_w, parent_h = app.winfo_width(), app.winfo_height()
        x, y = 240, 420
        self.geometry(f"{x}x{y}+{int(((parent_w-x)/2)+parent_x)}+{int(((parent_h-y)/2)+parent_y)}") # centers widget inside parent
        self.resizable(False, False)
        self.title("Settings")
        self.grab_set()
        self.grid_columnconfigure(0, weight=1)


        # title
        self.header = CTk.CTkLabel(self, text="Settings", font=("Arial", 24, "bold")) # Creates and styles text element
        self.header.grid(row=0, column=0, padx=10, pady=10) # Positions said element


        # all inputs
        self.label = CTk.CTkLabel(self, text="Adapter Name", fg_color=("white", "gray30"), corner_radius=6)
        self.label.grid(row=1, column=0, padx=10, pady=(10, 0))

        self.text_input = CTk.CTkEntry(self, width=128, height=32) # Creates an object for text entry
        self.text_input.insert(0, f"{adapter_name}")
        self.text_input.grid(row=2, column=0, pady=10)

        self.option_menu = CTk.CTkOptionMenu(self, values=["System", "Dark", "Light"], command=self.theme_updater)
        self.option_menu.grid(row=3, column=0, padx=(10, 0), pady=10)

        self.button = CTk.CTkButton(self, text="Create Desktop Shortcut", command=self.desktop_shortcut)
        self.button.grid(row=4, column=0, padx=10, pady=15)

        self.check_var = CTk.StringVar(value=start_value)

        self.start_checkbox = CTk.CTkCheckBox(self, text="Launch on start-up",
                                               variable=self.check_var, onvalue="1", offvalue="0")
        self.start_checkbox.grid(row=5, column=0)

        self.startup_menu = CTk.CTkOptionMenu(self, values=["Window", "Tray"], command=self.apply_args)
        self.startup_menu.grid(row=6, column=0, padx=(10, 0), pady=20)


        # save button
        self.button = CTk.CTkButton(self, text="Apply & Save", command=self.save)
        self.button.grid(row=7, column=0, padx=(10, 0), pady=30)

    
    def add_to_startup(self, args, shortcut_name="QLHM", exe_name="qlhm_main.exe"):
        """ Creates a shortcut with arguments to the executable and adds it to Startup. """
        startup_path = os.path.join(os.getenv('APPDATA'), r'Microsoft\Windows\Start Menu\Programs\Startup')
        shortcut_path = os.path.join(startup_path, f"{shortcut_name}.lnk")

        exe_dir = os.path.dirname(sys.executable)
        exe_full_path = os.path.join(exe_dir, exe_name)

        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.TargetPath = exe_full_path
        shortcut.Arguments = args # applys arguments to shortcut
        shortcut.WorkingDirectory = exe_dir
        shortcut.IconLocation = exe_full_path
        shortcut.save()

    def remove_startup(self, shortcut_name="QLHM"):
        """ Checks for and removes existing shortcuts. """
        startup_path = os.path.join(os.getenv('APPDATA'), r'Microsoft\Windows\Start Menu\Programs\Startup')
        shortcut_path = os.path.join(startup_path, f"{shortcut_name}.lnk")
        if os.path.exists(shortcut_path):
            os.remove(shortcut_path)
    
    def desktop_shortcut(self, shortcut_name="Quest Link Hotspot Manager", exe_name="qlhm_main.exe"):
        """ Creates a shortcut and adds it to the user's desktop. """
        desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        shortcut_path = os.path.join(desktop_path, f"{shortcut_name}.lnk")

        exe_dir = os.path.dirname(sys.executable)
        exe_full_path = os.path.join(exe_dir, exe_name)

        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.TargetPath = exe_full_path
        shortcut.WorkingDirectory = exe_dir
        shortcut.IconLocation = exe_full_path
        shortcut.save()
        DialogueBox(self, "Added shortcut to Desktop.", "Success")
    
    def theme_updater(self, choice):
        """ Applys currently selected theme and sets the variable for use during saving. """
        CTk.set_appearance_mode(choice.lower())
        global cur_theme
        cur_theme = choice.lower()
    
    def apply_args(self, choice):
        """ Assigns arg a value based on the option selected in the dropdown. """
        global args
        match choice:
            case "Window":
                args = "-max"
            case "Tray":
                args = "-min"
    
    def save(self):
        """ Writes current settings to data.qlhm then applys them to the running instance. """
        global args, adapter_name
        try:
            args
        except:
            args = "-max"

        match self.check_var.get(): # applys startup setting
            case "1":
                self.add_to_startup(args)
            case _:
                self.remove_startup()
        adapter_name = self.text_input.get() # recieves adapter name from input box
        with open(r"resource\data.qlhm", 'w') as f:
            f.write(f"{adapter_name}\n{cur_theme}\n{self.check_var.get()}") # writes settings to data.qlhm
        
        format_commands(self.text_input.get())


class DialogueBox(CTk.CTkToplevel):
    def __init__(self, parent, message, title, *args, **kwargs):
        """
        Runs on class instantiation:
        Defines layout and widget attributes
        """
        super().__init__(parent, *args, **kwargs)

        parent_x, parent_y = app.winfo_x(), app.winfo_y()
        parent_w, parent_h = app.winfo_width(), app.winfo_height()
        x, y = 300, 200
        self.geometry(f"{x}x{y}+{int(((parent_w-x)/2)+parent_x)}+{int(((parent_h-y)/2)+parent_y)}")
        self.resizable(False, False)

        self.title("Notice")
        self.grab_set() # blocks main window until closed
        self.grid_columnconfigure(0, weight=1)
        
        self.header = CTk.CTkLabel(self, text=title, font=("Arial", 18, "bold"))
        self.header.grid(row=0, column=0, padx=10, pady=10)

        self.label = CTk.CTkLabel(self, text=message, font=("Arial", 12, "bold"))
        self.label.grid(row=1, column=0, padx=10)

        self.button = CTk.CTkButton(self, text="OK", command=self.destroy) # closes window on click
        self.button.grid(row=2, column=0, padx=(10, 0), pady=10)

# assigns command variables
format_commands(adapter_name)
# starts app
app = App()
app.mainloop()