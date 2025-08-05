# SmartDeployTool

**SmartDeployTool** is a PowerShell-based GUI utility designed to automate and streamline the deployment and configuration of Windows workstations. It provides a user-friendly interface for IT administrators to perform common setup tasks, install selected applications, and apply system tweaks efficiently.

## Features

- **Graphical User Interface:** Built with Windows Forms for easy interaction.
- **Configurable Options:** Enable or disable deployment steps such as disabling hibernation, importing Wi-Fi profiles, uninstalling Microsoft 365, installing TeamViewer, antivirus, and more.
- **Application Selection:** Choose which applications to install from a configurable list (`config.json`).
- **Automated Installations:** Downloads and installs selected applications with silent arguments.
- **System Tweaks:** Apply registry and system settings changes as defined in the configuration.
- **Domain Join & Local Admin Creation:** Automate joining to a domain and creating local administrator accounts.
- **Logging:** Actions and errors are logged to `C:\deploy-log.txt` and `C:\deploy-errors.txt`.
- **Progress Tracking:** Visual progress bar and real-time log output.

## Requirements

- **Windows 10/11**
- **PowerShell 5.1+**
- **Administrator Privileges**
- **.NET Framework** (for Windows Forms)
- A properly configured `config.json` file in the script directory.

## Usage

1. **Prepare `config.json`:**  
   Define your applications, sources, and settings in the `config.json` file. See the example below.

2. **Run the Script as Administrator:**  
   Right-click `SmartDeployTool.ps1` and select "Run with PowerShell" (ensure you have admin rights).

3. **Select Deployment Options:**  
   Use the checkboxes to enable or disable deployment steps.

4. **Choose Applications:**  
   Click "Wybierz aplikacje" to select which applications to install.

5. **Start Deployment:**  
   Click "START" to begin the deployment process.

## Example `config.json`

```json
{
  "DefaultInstallSource": "Web",
  "InstallSourcePaths": {
    "Network": "\\server\share\apps\\",
    "Web": "https://example.com/apps/"
  },
  "Programs": {
    "7zip": {
      "FileName": "7z1900-x64.exe",
      "SilentArgs": "/S",
      "Enabled": true
    },
    "Notepad++": {
      "FileName": "npp.8.1.9.3.Installer.x64.exe",
      "SilentArgs": "/S",
      "Enabled": false
    }
  },
  "TeamViewer": {
    "FileName": "TeamViewer_Host.msi",
    "Arguments": "/qn"
  },
  "AntyVirus": {
    "FileName": "antivirus_installer.exe"
  },
  "WiFiProfile": {
    "FileName": "wifi-profile.xml"
  },
  "LocalAdmin": {
    "Username": "localadmin"
  },
  "DomainJoin": {
    "DomainName": "yourdomain.local",
    "Username": "domainuser"
  },
  "SystemSettings": {
    "DisableDeliveryOptimization": true,
    "EnableWin10StartMenu": false,
    "DisableTelemetry": true,
    "DisableCortana": true,
    "DisableFastStartup": true,
    "DisableNewsAndInterests": true
  }
}