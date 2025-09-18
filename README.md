# Quest Link Hotspot Manager

A simple tool to fix *Windows* hotspot for use with Quest Link, drastically improving connection stability.

## Usage
> [!NOTE]
> - Ensure your network adapter supports **5GHz** and has up to date drivers.
1. **Set up a Windows Hotspot:**
   - Go to **Settings > Network & Internet > Mobile hotspot**
   - Set the **Network band** to **5GHz**
   - Set the SSID (Name) and password used to connect to the AP

2. **Using the app:**
   - Download and run Quest Link Hotspot Manager.
   - Click *Scan for Devices* and check the output.
       - Note the output: `Couldn't find Adapter`
       - If you see this error, copy the Interface Name of your Wi-Fi card and save it to adapter name in settings.

3. **Optional Tweaks for Stability:**
   - Disable Wi-Fi power saving mode:
     - Open **Control Panel** → **Network and Internet** → **Network Connections**
     - Find your adapter, right-click, and choose **Properties**
     - Click **Configure** → **Power Management**
     - Uncheck **Allow the computer to turn off this device to save power** or similar.
> [!WARNING]
> If hotspot doesn't start, check **Scan for Devices**
> If the adapter is not found, you must add it manually in settings.

## Finding Your Adapter Name

To identify your network adapter:

- Run **Scan for Devices** in the application
- Choose the Wireless adapter you want to use and save it in the settings.
> [!TIP]
> The adapter name should look something like WiFi/Wi-Fi 2. Don't pick any non-wireless options.

## Using the system tray

By default a system tray icon will open with the programs closure, you can drag it out to
the task bar for easier interaction. The program can be set to launch in this mode as one
of the start-up arguments within the settings. (Change option in Window/Tray drop down)

## What does it do

This program fixes an issue in Windows 10/11 which causes latency spikes every few seconds 
for devices connected to it's hotspot caused by automatically searching for new networks.
The script now changes `ExecutionPolicy` in PowerShell, so no UAC prompt is necessary.

## Source code

Use `pip install -r requirements.txt` and run *qlhm_main.py*.