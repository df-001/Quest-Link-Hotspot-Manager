import webview
import subprocess
import re
from pathlib import Path
import json
import os
import urllib.request
import json
import webbrowser
import threading

CURRENT_VERSION = "v2.0.1"
CONFIG_FILE = Path(os.getenv("APPDATA")) / "QLHotspot" / "config.json"
CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)

def check_for_updates():
    try:
        url = f"https://api.github.com/repos/df-001/Quest-Link-Hotspot-Manager/releases/latest"
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=3) as response:
            data = json.loads(response.read().decode())
            latest_version = data.get("tag_name")
            
            if latest_version and latest_version != CURRENT_VERSION:
                print(f"Update available: {latest_version}")
                webbrowser.open(data.get("html_url"))
    except Exception:
        pass

def save_adapter_name(adapter_name: str):
    CONFIG_FILE.write_text(json.dumps({"adapter_name": adapter_name}))

def load_adapter_name():
    if CONFIG_FILE.exists():
        return json.loads(CONFIG_FILE.read_text()).get("adapter_name")
    return None

run_cmd = """
    $WirelessAdapterName = "{}"
    $connectionProfile = [Windows.Networking.Connectivity.NetworkInformation,Windows.Networking.Connectivity,ContentType=WindowsRuntime]::GetInternetConnectionProfile()
    $tetheringManager = [Windows.Networking.NetworkOperators.NetworkOperatorTetheringManager,Windows.Networking.NetworkOperators,ContentType=WindowsRuntime]::CreateFromConnectionProfile($connectionProfile)

    Start-Sleep -Seconds 1
    netsh wlan set autoconfig enabled=yes interface="$WirelessAdapterName"

    $tetheringManager.StartTetheringAsync()
    Start-Sleep -Seconds 2

    netsh wlan set autoconfig enabled=no interface="$WirelessAdapterName"
    exit
"""
end_cmd = """
    $WirelessAdapterName = "{}"
    $connectionProfile = [Windows.Networking.Connectivity.NetworkInformation,Windows.Networking.Connectivity,ContentType=WindowsRuntime]::GetInternetConnectionProfile()
    $tetheringManager = [Windows.Networking.NetworkOperators.NetworkOperatorTetheringManager,Windows.Networking.NetworkOperators,ContentType=WindowsRuntime]::CreateFromConnectionProfile($connectionProfile)

    netsh wlan set autoconfig enabled=yes interface="$WirelessAdapterName"
    $tetheringManager.StopTetheringAsync()
    exit
"""
scan_cmd = "netsh interface show interface"

def format_cmd(script: str, adapterName: str):
    return script.format(adapterName)

def run_powershell(script: str):
    """
    Run multiline PowerShell commands and return stdout.
    
    :param script: Multiline PowerShell script as a string
    :return: stdout output (str)
    :raises RuntimeError: if PowerShell returns an error
    """
    result = subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-Command", script
        ],
        capture_output=True,
        text=True
    )

    if result.returncode != 0:
        raise RuntimeError(f"PowerShell error:\n{result.stderr.strip()}")

    return result.stdout


class Api():    
    def refresh(self):
        """
        Parses Powershell return values to update network adapters.

        :return: Payload containing all found adapters and the user's saved adapter.
        """
        
        output = (run_powershell(scan_cmd))
        
        parsed_output = re.split(" {2}|\n", output)
        parsed_output = [i for i in parsed_output if i.strip()]
        adapter_list = parsed_output[8::4]

        stored_name = load_adapter_name()

        return [adapter_list, stored_name]

    def start_hotspot(self, adapter_name: str):
        run_powershell(format_cmd(run_cmd, adapter_name))
        return True

    def end_hotspot(self, adapter_name: str):
        run_powershell(format_cmd(end_cmd, adapter_name))
        return True

    def save_adapter(self, adapter_name: str):
        save_adapter_name(adapter_name)
        return "Saved."


if __name__ == "__main__":
    threading.Thread(target=check_for_updates, daemon=True).start()
    
    api = Api()
    webview.create_window(
        "Quest Link Hotspot Manager",
        "main.html",
        js_api=api,
        width=480,
        height=700,
        resizable=False,
    )
    webview.start(debug=False)