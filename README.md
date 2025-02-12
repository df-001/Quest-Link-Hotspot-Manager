# Quest Link Hotspot Manager

A simple tool to fix Windows hotspot for use with Quest Link, improving connection stability.

## Usage
> [!NOTE]
> - Ensure your network adapter supports **5GHz** and has up to date drivers.
1. **Set up a Windows Hotspot:**
   - Go to **Settings > Network & Internet > Mobile hotspot**
   - Set the **Network band** to **5GHz**
   - Enable the **Mobile hotspot**

2. **Run the Executable:**
   - Download and run the provided file.

3. **Optional Tweaks for Stability:**
   - Disable Wi-Fi power saving mode:
     - Open **Control Panel** → **Network and Internet** → **Network Connections**
     - Find your adapter, right-click, and choose **Properties**
     - Click **Configure** → **Power Management**
     - Uncheck **Allow the computer to turn off this device to save power** or similar.
> [!WARNING]
> If the script cannot find the adapter, write the desired adapter name in the program
> then restart it and try again.

## Finding Your Adapter Name

To identify your network adapter:

- Navigate to **Control Panel\Network and Internet\Network Connections**
- Choose the Wireless adapter you want to use and apply it in the program by name.
> [!TIP]
> The adapter name will be the first line of the device e.g. WiFi/Wi-Fi 2

## What does it do

This program fixes an issue in Windows 10/11 which causes latency spikes every few seconds 
for devices connected to it's hotspot caused by automatically searching for new networks.
The script does not change global `ExecutionPolicy` in PowerShell, rather allowing just the
necessary scripts to run to avoid security compromises. This means that script execution
does require Administrator permissions.