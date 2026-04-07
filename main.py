import webview
import subprocess
import re

run_cmd = """
    $WirelessAdapterName = "{}"
    $connectionProfile = [Windows.Networking.Connectivity.NetworkInformation,Windows.Networking.Connectivity,ContentType=WindowsRuntime]::GetInternetConnectionProfile()
    $tetheringManager = [Windows.Networking.NetworkOperators.NetworkOperatorTetheringManager,Windows.Networking.NetworkOperators,ContentType=WindowsRuntime]::CreateFromConnectionProfile($connectionProfile)

    Start-Sleep -Seconds 1
    netsh wlan set autoconfig enabled=yes interface=$WirelessAdapterName

    $tetheringManager.StartTetheringAsync()
    Start-Sleep -Seconds 2

    netsh wlan set autoconfig enabled=no interface=$WirelessAdapterName
    exit
"""

end_cmd = """
    $WirelessAdapterName = "{}"
    $connectionProfile = [Windows.Networking.Connectivity.NetworkInformation,Windows.Networking.Connectivity,ContentType=WindowsRuntime]::GetInternetConnectionProfile()
    $tetheringManager = [Windows.Networking.NetworkOperators.NetworkOperatorTetheringManager,Windows.Networking.NetworkOperators,ContentType=WindowsRuntime]::CreateFromConnectionProfile($connectionProfile)

    netsh wlan set autoconfig enabled=yes interface=$WirelessAdapterName
    $tetheringManager.StopTetheringAsync()
    exit
"""

scan_cmd = "netsh interface show interface"

def format_cmd(script, adapterName):
    return script.format(adapterName)

def run_powershell(script: str) -> str:
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
        """
        
        output = (run_powershell(scan_cmd))
        
        parsed_output = re.split(" {2}|\n", output)
        parsed_output = [i for i in parsed_output if i.strip()]
        adapter_list = parsed_output[8::4]

        return adapter_list

    def start_hotspot(self, adapter_name):
        run_powershell(format_cmd(run_cmd, adapter_name))
        return True

    def end_hotspot(self, adapter_name):
        run_powershell(format_cmd(end_cmd, adapter_name))
        return True


if __name__ == "__main__":
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