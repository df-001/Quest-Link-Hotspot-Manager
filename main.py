import subprocess
import customtkinter as CTk

with open(r"data.qlhm", "rt") as f:
    adapter_name = f.read()
    print("Contents loaded.")

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

end_cmd = """
$WirelessAdapterName = \\"{}\\"
$connectionProfile = [Windows.Networking.Connectivity.NetworkInformation,Windows.Networking.Connectivity,ContentType=WindowsRuntime]::GetInternetConnectionProfile()
$tetheringManager = [Windows.Networking.NetworkOperators.NetworkOperatorTetheringManager,Windows.Networking.NetworkOperators,ContentType=WindowsRuntime]::CreateFromConnectionProfile($connectionProfile)

netsh wlan set autoconfig enabled=yes interface=$WirelessAdapterName
$tetheringManager.StopTetheringAsync()
""".format(adapter_name)

test_command = f"""
# Variables
$WirelessAdapterName = \"{adapter_name}\"

# Pause
Write-Host \"Press ENTER to restore $WirelessAdapterName adapter scan and disable HotSpot\"
pause
"""


def run_command(cmd):
    subprocess.run(["powershell", "Start-Process", "powershell", "-ArgumentList", f"'-ExecutionPolicy Bypass -Command {cmd}'", "-Verb", "RunAs"], capture_output=True, text=True)

class App(CTk.CTk):
    def __init__(self):
        super().__init__()
        # Window definition
        self.title("QL Hotspot Manager") 
        self.geometry("360x420")
        self.grid_columnconfigure(0, weight=1) # Spans column 0 across window via weight
        # Header
        self.header = CTk.CTkLabel(self, text="Quest Link\nHotspot Manager", font=("Arial", 24, "bold")) # Creates and styles text element
        self.header.grid(row=0, column=0, padx=10, pady=20) # Positions said element
        # Adapter Input
        self.label = CTk.CTkLabel(self, text="Adapter Name", fg_color="gray30", corner_radius=6)
        self.label.grid(row=1, column=0, padx=10, pady=(10, 0))

        self.text_input = CTk.CTkEntry(self, width=64, height=32) # Creates an object for text entry
        self.text_input.insert(0, f"{adapter_name}")
        self.text_input.grid(row=2, column=0, pady=10)

        self.label2 = CTk.CTkLabel(self, text="Don't change unless necessary!", font=("Arial", 10))
        self.label2.grid(row=3, column=0, pady=1)

        self.apply_change_button = CTk.CTkButton(self, text="Apply Changes", fg_color="#226", command=self.submit_adapter_name)
        self.apply_change_button.grid(row=4, column=0, pady=10)
        # Main functions
        self.status = CTk.CTkLabel(self, text="", font=("Arial", 11))
        self.status.grid(row=5, column=0)

        self.button = CTk.CTkButton(self, text="Start Hotspot", font=("Arial", 14, "bold"), command=self.start_hotspot)
        self.button.grid(column=0, padx=50, pady=10, sticky="ew") # Expands button east+west to borders

        self.button = CTk.CTkButton(self, text="End Hotspot", font=("Arial", 14, "bold"), command=self.end_hotspot)
        self.button.grid(column=0, padx=50, pady=10, sticky="ew")
        
    def submit_adapter_name(self):
        self.adapter_name = self.text_input.get()
        with open(r"data.qlhm", "wt") as f:
            f.write(self.adapter_name)
        self.label2.configure(text=f"Changed to: {self.adapter_name}\nRestart program to take effect.")
        return self.adapter_name
    
    def start_hotspot(self):
        self.status.configure(text=f"Starting Hotspot...")
        run_command(start_cmd)
        print("Hotspot starting...")

    def end_hotspot(self):
        self.status.configure(text=f"Ending Hotspot...")
        run_command(end_cmd)
        print("Ending hotspot...")
    

app = App()
app.mainloop()