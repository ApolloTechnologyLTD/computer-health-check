<div align="center">

# Apollo Technology Computer Health Check

**Comprehensive system diagnostics and maintenance with automated reporting!**

</div>

---

## 📖 About

This PowerShell script performs a full system health check on Windows 10/11 devices. It automates common maintenance tasks, validates system integrity, checks hardware health, updates software, and generates a **professional Apollo Technology–branded report** for engineers or end users.

The script is designed for **IT engineers, MSPs, and support teams** and features an interactive **Engineer Mode** allowing steps to be skipped, delayed, or executed immediately. A **Demo Mode** is also included for training and testing report output.

> ⚠️ **Disclaimer** – This is an independent project developed by **Apollo Technology** and is not affiliated with Microsoft.

---

## ✨ Features

### 🛡️ System Integrity Checks
- Runs DISM image repair
- Runs SFC system file validation
- Extracts repair details into the final report

### 🚀 Software Update Management
- Detects outdated applications via Winget
- Automatically upgrades supported applications
- Logs update success or failure per app

### 🧹 System Cleanup
- Clears temporary files
- Empties Recycle Bin
- Removes obsolete Windows Update cache

### 🔋 Hardware & Device Health
- Battery health and charge status
- CPU, RAM, and system information
- Disk usage and drive status

### 💾 Disk Optimisation
- Optional disk optimisation with engineer confirmation
- Safe defaults to prevent long unintended runs

### 📝 Branded Reporting
- Generates Apollo Technology–branded HTML report
- Automatically converts report to PDF
- Saves report to the Public Desktop

### 📧 Email Integration
- Optional automatic emailing of the final report
- SMTP configuration supported

### ⏱️ Interactive Engineer Mode
- Timed prompts for each major task
- Start immediately, skip, or auto-run
- All decisions logged into the report

---

## 📋 Requirements

| Requirement | Details |
|------------|---------|
| **Operating System** | Windows 10 or Windows 11 |
| **PowerShell** | PowerShell 5.1 or later |
| **Permissions** | Administrator privileges required (auto-elevation supported) |
| **Internet Access** | Required for Winget, Windows Updates, and email |
| **Browser** | Microsoft Edge (used for PDF generation) |

---

## 🚀 Quick Start

```powershell
iwr https://short.apollotechnology.co.uk/health_check -OutFile heathcheck.ps1; powershell -ExecutionPolicy Bypass .\heathcheck.ps1
```

> **Important:** Run PowerShell as **Administrator**.

---

## ⚙️ Configuration

The script is configured by editing variables at the top of the `heathcheck.ps1` file.

| Variable | Default | Description |
|--------|---------|-------------|
| `$DemoMode` | `$false` | Simulates all checks with dummy data for testing/demo purposes |
| `$EmailEnabled` | `$false` | Enables automatic emailing of the final report |
| `$ToAddress` | `support@...` | Email address to send the report to |
| `$SmtpServer` | `smtp.office365.com` | SMTP server used for sending email |
| `$LogoUrl` | `(URL)` | Apollo Technology logo URL used in the report |

---

## 📚 Usage Guide

### Interactive Execution

On launch, the script prompts for an **Engineer Name**.  
This name is stamped into the final report.

The script proceeds through structured maintenance stages. For longer-running tasks, an interactive timer is displayed:

```plaintext
[NEXT STEP] DISM & SFC Scans
[S] Skip  |  [Enter] Start Now  |  Auto-start in 5 seconds
```

**Controls:**
- **Enter** – Start the step immediately
- **S** – Skip the step (marked as *Skipped by Engineer* in the report)
- **Wait** – Step auto-starts after the countdown

---

## 📝 Report Output

Once completed, the script generates:

- **Apollo_Health_Report_YYYY-MM-DD.pdf**
- **Location:** Public Desktop

### Additional Report Details
- HTML version is created if PDF conversion fails
- SFC repair entries are extracted from `CBS.log`
- Skipped or failed steps are clearly documented
- Engineer name and execution timestamps are included

---

## 🔍 Key Features Explained

### 🛡️ System Integrity (SFC & DISM)
- `DISM /RestoreHealth` repairs the Windows system image
- `SFC /Scannow` validates protected system files
- Relevant log entries are embedded into the report if repairs are made

### 📦 Winget Integration

Detects installed applications with available updates and automatically attempts upgrades using:

```powershell
winget upgrade --all --accept-source-agreements
```

Upgrade results are captured and written to the report.

### 🔄 Windows Updates
- Queries Windows Update Agent for pending updates
- Updates are reported only, not force-installed
- Prevents unexpected long reboots during quick health checks

### 💾 Disk Optimisation

Includes a confirmation timer (default **60 seconds**).  
Defaults to **Skip** to prevent unintended long operations.

If approved, executes:

```powershell
Optimize-Volume -DriveLetter C
```

---

## ⚠️ Disclaimer

This script is provided **as-is**, without warranty of any kind.  
It performs administrative actions including:

- 🔥 Deleting temporary files
- 🔥 Modifying system files (SFC / DISM)
- 🔥 Installing application updates

Always ensure:
- A system restore point exists (the script attempts to create one)
- Output is reviewed by a qualified engineer before remediation

---

<div align="center">

**Last updated:** 2026-01-12

</div>
