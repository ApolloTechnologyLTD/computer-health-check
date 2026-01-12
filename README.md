<div align="center">

## Windows Computer Health Check Tool

**Quickly assess the health of a Windows computer in one command**

## 📖 About

This PowerShell script performs a comprehensive **computer health check** and reports on key system components such as hardware status, operating system health, disk space, performance indicators, and security configuration.

The tool is designed for **IT administrators, MSPs, and support engineers** to quickly identify potential issues on Windows devices.

> ⚠️ **Disclaimer** – This is an independent project and is not affiliated with Microsoft.

---

## ✨ Features

- 🖥️ **System Information** – OS version, uptime, device name, manufacturer
- 💾 **Disk Health Checks** – Free space, disk usage, and drive status
- 🧠 **Memory & CPU** – RAM usage and processor information
- 🔐 **Security Status** – Antivirus and Windows Defender status
- 🌐 **Network Overview** – Active network adapters and IP configuration
- 🔄 **Windows Update Status** – Update service and patching visibility
- 🧾 **Event Log Review** – Highlights recent critical and error events
- 📊 **Clear Output** – Easy-to-read console output for fast diagnostics
- ⚡ **Single-Command Execution** – No installation required

---

## 📋 Requirements

| Requirement | Details |
|-------------|---------|
| **OS** | Windows 10 / 11 or Windows Server |
| **PowerShell** | PowerShell 5.1 or later |
| **Permissions** | Administrator privileges recommended |
| **Internet** | Required only to download the script |

---

## 🚀 Quick Start

Run the health check without cloning the repository:

```powershell
iwr https://short.apollotechnology.co.uk/health_check -OutFile heathcheck.ps1; powershell -ExecutionPolicy Bypass .\heathcheck.ps1
```

⚠️ **Important:**  
Run PowerShell **as Administrator** to ensure all system and security checks can complete successfully.  
Limited permissions may result in incomplete data.

---

## 📚 What the Script Checks

### 🖥️ System Overview
- Computer name
- Operating system and build
- System uptime

### 💾 Storage
- All fixed drives
- Free space thresholds
- Disk health indicators

### 🧠 Performance
- Installed memory
- Memory usage
- CPU details

### 🔐 Security
- Installed antivirus
- Windows Defender status
- Real-time protection state

### 🔄 Windows Updates
- Update service status
- Pending reboot detection (if applicable)

### 🌐 Networking
- Active adapters
- IP address information

### 🧾 Event Logs
- Recent critical and error events
- Helps identify underlying system issues

---

## 🛠️ Intended Use Cases

- First-line IT diagnostics
- Pre-deployment device checks
- Remote support troubleshooting
- Routine system health audits
- MSP onboarding checks

---

## ⚠️ Disclaimer

This script is provided **as-is**, without warranty of any kind.

While it **does not make system changes**, the output should be reviewed by a qualified technician before taking action.

---</div>

<div align="center">

_Last updated: 2026-01-12_

</div>
