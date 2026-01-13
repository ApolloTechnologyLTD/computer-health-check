<#
.SYNOPSIS
    Apollo Technology Full Health Check Script v17.1
.DESCRIPTION
    Full system health check.
    - UPDATED: Disk Optimization now auto-skips on SSD/NVMe drives with specific report message.
    - NEW: NVMe/SSD/HDD Detection with SMART Status.
    - NEW: Detailed Device Information Section.
    - NEW: Customer/Ticket Input Validation Loop.
    - RETAINED: Prevents Windows from Sleeping/Locking while running.
    - Report saves to C:\temp\Apollo_Reports.
    - Storage Analysis with Pie Chart & Top Folder usage.
    - Captures and embeds SFC [SR] logs into the report.
    - PATCH: Added check for cleanmgr.exe to prevent crash on Server 2008.
.NOTES
    Author: Apollo Technology (Lewis Wiltshire)
#>

# --- CONFIGURATION ---
$LogoUrl = "https://raw.githubusercontent.com/ApolloTechnologyLTD/computer-health-check/main/Apollo%20Cropped.png"
$DemoMode = $false           # Set to $true for fast simulation with DUMMY DATA
$EmailEnabled = $false       # Set to $true to enable email

# Email Settings
$SmtpServer   = "smtp.office365.com"
$SmtpPort     = 587
$FromAddress  = "reports@yourdomain.com"
$ToAddress    = "support@yourdomain.com"
$UseSSL       = $true

# --- 1. ELEVATE TO ADMIN (Skipped if DemoMode is True) ---
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if ((-not $DemoMode) -and (-not $isAdmin)) {
    Write-Host "Requesting Administrator privileges..." -ForegroundColor Yellow
    try {
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        exit
    } catch {
        Write-Error "Failed to elevate. Please right-click and 'Run as Administrator'."
        exit
    }
}

# --- 2. DISABLE QUICK-EDIT (PREVENTS FREEZING) ---
$consoleFuncs = @"
using System;
using System.Runtime.InteropServices;
public class ConsoleUtils {
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern IntPtr GetStdHandle(int nStdHandle);
    [DllImport("kernel32.dll")]
    public static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);
    [DllImport("kernel32.dll")]
    public static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);
    public static void DisableQuickEdit() {
        IntPtr hConsole = GetStdHandle(-10); // STD_INPUT_HANDLE
        uint mode;
        GetConsoleMode(hConsole, out mode);
        mode &= ~0x0040u; // ENABLE_QUICK_EDIT_MODE = 0x0040
        SetConsoleMode(hConsole, mode);
    }
}
"@
try {
    Add-Type -TypeDefinition $consoleFuncs -Language CSharp
    [ConsoleUtils]::DisableQuickEdit()
} catch { }

# --- 2.5 PREVENT SLEEP ---
$sleepBlocker = @"
using System;
using System.Runtime.InteropServices;
public class SleepUtils {
    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern uint SetThreadExecutionState(uint esFlags);
}
"@
try {
    Add-Type -TypeDefinition $sleepBlocker -Language CSharp
    $null = [SleepUtils]::SetThreadExecutionState(0x80000003) # ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
} catch { }

# --- HELPER FUNCTION: TIMER ---
function Test-UserSkip {
    param([string]$StepName, [int]$Seconds = 5)
    
    if ($psISE) { return $false } 

    Write-Host "`n   [NEXT STEP] $StepName" -ForegroundColor Cyan
    Write-Host "   [S] Skip  |  [Enter] Start Now  |  Wait for Auto-Start" -ForegroundColor Gray
    
    # Flush Input Buffer
    while ([Console]::KeyAvailable) { $null = [Console]::ReadKey($true) }

    $EndTime = (Get-Date).AddSeconds($Seconds)
    
    while ((Get-Date) -lt $EndTime) {
        $Remaining = [math]::Ceiling(($EndTime - (Get-Date)).TotalSeconds)
        Write-Host "`r   > Auto-starting in $Remaining seconds...    " -NoNewline -ForegroundColor Yellow
        
        if ([Console]::KeyAvailable) {
            $Key = [Console]::ReadKey($true)
            if ($Key.Key -eq 'S') {
                Write-Host "`r   > SKIPPED BY ENGINEER               " -ForegroundColor Red
                return $true
            }
            if ($Key.Key -eq 'Enter') {
                Write-Host "`r   > STARTING IMMEDIATELY...           " -ForegroundColor Green
                return $false
            }
        }
        Start-Sleep -Milliseconds 100
    }
    
    Write-Host "`r   > AUTO-STARTING...                  " -ForegroundColor Green
    return $false
}

# --- 3. INTRO & SETUP ---
Clear-Host
$ApolloASCII = @"
    ___    ____  ____  __    __    ____     ____________________  ___   ______  __    ____  ________  __
   /   |  / __ \/ __ \/ /   / /   / __ \   /_  __/ ____/ ____/ / / / | / / __ \/ /   / __ \/ ____/\ \/ /
  / /| | / /_/ / / / / /   / /   / / / /    / / / __/ / /   / /_/ /  |/ / / / / /   / / / / / __   \  / 
 / ___ |/ ____/ /_/ / /___/ /___/ /_/ /    / / / /___/ /___/ __  / /|  / /_/ / /___/ /_/ / /_/ /   / /  
/_/  |_/_/    \____/_____/_____/\____/    /_/ /_____/\____/_/ /_/_/ |_/\____/_____/\____/\____/   /_/   
                                                                                                        
"@

Write-Host $ApolloASCII -ForegroundColor Cyan
if ($DemoMode) { Write-Host "      *** DEMO MODE ACTIVE - GENERATING DUMMY DATA ***" -ForegroundColor Magenta }
if (-not $isAdmin -and $DemoMode) { Write-Host "      [NOTICE] Running as Standard User" -ForegroundColor Yellow }
Write-Host "      [POWER] Sleep Mode & Screen Timeout Blocked." -ForegroundColor DarkGray

# --- CHANGED PATH HERE ---
$BaseDir = "C:\temp\Apollo_Reports"

try {
    if (-not (Test-Path $BaseDir)) { New-Item -Path $BaseDir -ItemType Directory -Force -ErrorAction Stop | Out-Null }
} catch {
    $BaseDir = "$env:USERPROFILE\Desktop\Apollo_Reports"
    if (-not (Test-Path $BaseDir)) { New-Item -Path $BaseDir -ItemType Directory -Force | Out-Null }
    Write-Warning "Could not create temp path. Using Desktop path: $BaseDir"
}

# --- INPUT & VALIDATION LOOP ---
do {
    Write-Host "`n--- REPORT DETAILS SETUP ---" -ForegroundColor Cyan
    $EngineerName = Read-Host "Enter Engineer Name"
    $CustomerName = Read-Host "Enter Customer Full Name"
    $TicketNumber = Read-Host "Enter Ticket Number"

    # Check if # is present, if not add it
    if ($TicketNumber -notmatch "^#") {
        $TicketNumber = "#$TicketNumber"
    }

    Write-Host "`nPlease confirm the details below:" -ForegroundColor Yellow
    Write-Host "Engineer: $EngineerName"
    Write-Host "Customer: $CustomerName"
    Write-Host "Ticket:   $TicketNumber"
    
    $Confirmation = Read-Host "`nAre these details correct? (Press [Enter] or [Y] for Yes, [N] for No)"
    if ($Confirmation -eq "") { $Confirmation = "Y" }

    if ($Confirmation -match "N") {
        Clear-Host
        Write-Host $ApolloASCII -ForegroundColor Cyan
        Write-Host "Restarting input..." -ForegroundColor Red
    }

} while ($Confirmation -match "N")


# --- INTERNET CHECK (STARTUP) ---
Write-Host "`n   [CHECK] Verifying Internet Connection..." -ForegroundColor Yellow
while ($true) {
    if ($DemoMode) { 
        Write-Host "`r   > Simulating connection check...       " -NoNewline -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        Write-Host "`r   > Connection Verified (DEMO).          " -ForegroundColor Green
        break 
    }
    if (Test-Connection -ComputerName 8.8.8.8 -Count 3 -Quiet -ErrorAction SilentlyContinue) {
        Write-Host "`r   > Connection Verified.                 " -ForegroundColor Green
        break
    }
    Write-Host "`r   > Waiting for internet connection...   " -NoNewline -ForegroundColor Red
    Start-Sleep -Seconds 3
}

$CurrentDate = Get-Date -Format "yyyy-MM-dd HH:mm"
$CurrentYear = (Get-Date).Year
$FileDate = Get-Date -Format "yyyy-MM-dd_HH-mm"
$ReportFilename = "Apollo_Health_Report_$FileDate"
$ReportData = @{}
$StorageReportHTML = "" # Initialize Storage HTML
$SfcLogContent = "" # Initialize Log Content

# Email Creds
$EmailCreds = $null
if ($EmailEnabled) {
    Write-Host "`nEmail sending is enabled." -ForegroundColor Cyan
    $EmailPass = Read-Host "Please enter the Password for $FromAddress" -AsSecureString
    $EmailCreds = New-Object System.Management.Automation.PSCredential ($FromAddress, $EmailPass)
}

Write-Host "`nInitializing checks for Engineer: $EngineerName..." -ForegroundColor Cyan
Start-Sleep -Seconds 2

# --- 4. GATHER DEVICE INFORMATION ---
Write-Host "`n[INFO] Gathering Hardware Information..." -ForegroundColor Yellow
$ComputerInfo = Get-CimInstance Win32_ComputerSystem
$OSInfo = Get-CimInstance Win32_OperatingSystem
$CPUInfo = Get-CimInstance Win32_Processor
$FormattedInstallDate = "Unknown"
if ($OSInfo.InstallDate) {
    try { $FormattedInstallDate = $OSInfo.InstallDate.ToString("yyyy-MM-dd HH:mm") } catch {}
}

# DRIVE DETECTION (NVMe/SSD/HDD + SMART)
$PhysicalDisks = Get-PhysicalDisk | Sort-Object DeviceId
$DriveListHTML = ""

foreach ($disk in $PhysicalDisks) {
    # Type Logic
    $Type = ""
    if ($disk.MediaType -eq "HDD") { $Type = "Hard Disk Drive (HDD)" }
    elseif ($disk.MediaType -eq "SSD") {
        if ($disk.BusType -eq "NVMe") { $Type = "NVMe Solid State Drive" } 
        else { $Type = "SATA Solid State Drive (SSD)" }
    }
    else { $Type = "Unknown ($($disk.MediaType))" }
    
    # Size Logic
    $SizeGB = [math]::Round($disk.Size / 1GB, 2)
    $SizeStr = if ($SizeGB -gt 1000) { "$([math]::Round($SizeGB / 1024, 2)) TB" } else { "$SizeGB GB" }

    # Health Logic
    $HealthColor = if ($disk.HealthStatus -eq "Healthy") { "green" } else { "red" }
    
    $DriveListHTML += "<li><strong>Disk $($disk.DeviceId):</strong> $Type <br> Size: $SizeStr | Health: <span style='color:$HealthColor'>$($disk.HealthStatus)</span></li>"
}

# --- 5. SYSTEM RESTORE POINT ---
if (Test-UserSkip -StepName "Restore Point Creation") {
    $ReportData.RestorePoint = "Skipped by Engineer."
} else {
    if ($DemoMode) {
        Start-Sleep -Seconds 1
        $ReportData.RestorePoint = "Success: Restore point created."
    } else {
        try {
            Write-Host "   > Creating Restore Point..." -ForegroundColor Yellow
            Enable-ComputerRestore -Drive "C:\" -ErrorAction SilentlyContinue
            Checkpoint-Computer -Description "ApolloHealthCheck" -RestorePointType "MODIFY_SETTINGS" -ErrorAction Stop
            $ReportData.RestorePoint = "Success: Restore point created successfully."
        } catch {
            $ReportData.RestorePoint = "Note: Could not create restore point (System Protection disabled or Access Denied)."
        }
    }
    Write-Host "   > Done." -ForegroundColor Green
}

# --- 6. BATTERY STATUS (Instant) ---
Write-Host "`n[INFO] Checking Battery Health (Instant)..." -ForegroundColor Yellow
if ($DemoMode) {
     $ReportData.Battery = "Battery Detected. Status: OK - Charge: 94% (Healthy)"
} else {
    $Battery = Get-CimInstance -ClassName Win32_Battery -ErrorAction SilentlyContinue
    if ($Battery) {
        $ReportData.Battery = "Battery Detected. Status: $($Battery.Status) - Charge: $($Battery.EstimatedChargeRemaining)%"
    } else {
        $ReportData.Battery = "No Battery Detected (Desktop)."
    }
}
Write-Host "   > Done." -ForegroundColor Green

# --- 7. DISK CLEANUP ---
if (Test-UserSkip -StepName "Disk Cleanup") {
    $ReportData.DiskCleanup = "Skipped by Engineer."
} else {
    Write-Host "   > Cleaning System Files..." -ForegroundColor Yellow
    if ($DemoMode) {
        Start-Sleep -Seconds 1
        $ReportData.DiskCleanup = "Maintenance Complete: Removed 1.2GB of temporary files."
    } else {
        # PATCH: Check if cleanmgr.exe exists before running
        if (Get-Command "cleanmgr.exe" -ErrorAction SilentlyContinue) {
            $StateFlags = 1337
            $RegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches"
            $Handlers = @("Temporary Files", "Recycle Bin", "Active Setup Temp Folders", "Downloaded Program Files", "Temporary Setup Files", "Old ChkDsk Files", "Update Cleanup")
            foreach ($Handler in $Handlers) {
                $Key = "$RegPath\$Handler"
                if (Test-Path $Key) { Set-ItemProperty -Path $Key -Name "StateFlags$StateFlags" -Value 2 -Type DWord -ErrorAction SilentlyContinue }
            }
            try {
                Start-Process -FilePath "cleanmgr.exe" -ArgumentList "/sagerun:$StateFlags" -Wait -WindowStyle Hidden -ErrorAction Stop
                $ReportData.DiskCleanup = "Maintenance Complete: Temporary files and Recycle Bin cleaned."
            } catch {
                $ReportData.DiskCleanup = "Error: Failed to run Disk Cleanup (Process Error)."
            }
        } else {
            Write-Warning "Disk Cleanup (cleanmgr.exe) not found on this system."
            $ReportData.DiskCleanup = "Skipped: Disk Cleanup utility not installed."
        }
    }
    Write-Host "   > Done." -ForegroundColor Green
}

# --- 8. STORAGE ANALYSIS ---
Write-Host "`n[INFO] Analyzing Storage Usage..." -ForegroundColor Yellow

$StorageUsedGB = 0
$StorageFreeGB = 0
$StorageTotalGB = 0
$StoragePercent = 0
$StorageTableRows = ""

if ($DemoMode) {
    Start-Sleep -Seconds 2
    # DUMMY DATA
    $StorageUsedGB = 340.5
    $StorageFreeGB = 135.5
    $StorageTotalGB = 476.0
    $StoragePercent = 71
    $StorageTableRows = @"
    <tr><td>C:\Users\Apollo</td><td>120.4 GB</td></tr>
    <tr><td>C:\Windows</td><td>45.2 GB</td></tr>
    <tr><td>C:\Program Files</td><td>32.1 GB</td></tr>
    <tr><td>C:\Program Files (x86)</td><td>18.5 GB</td></tr>
    <tr><td>C:\hiberfil.sys</td><td>12.0 GB</td></tr>
"@
} else {
    try {
        # 1. Get Drive Info
        $Drive = Get-PSDrive C -ErrorAction SilentlyContinue
        if ($Drive) {
            $StorageUsedGB  = [math]::Round($Drive.Used / 1GB, 2)
            $StorageFreeGB  = [math]::Round($Drive.Free / 1GB, 2)
            $StorageTotalGB = [math]::Round(($Drive.Used + $Drive.Free) / 1GB, 2)
            
            if ($StorageTotalGB -gt 0) {
                $StoragePercent = [math]::Round(($StorageUsedGB / $StorageTotalGB) * 100)
            }
        }

        # 2. Get Top Heavy Folders/Files (Root Level)
        Write-Host "   > Scanning largest folders in C:\ (This may take a moment)..." -ForegroundColor Yellow
        $RootItems = Get-ChildItem "C:\" -Force -ErrorAction SilentlyContinue
        $FolderStats = @()
        
        foreach ($Item in $RootItems) {
            try {
                $Size = 0
                if ($Item.PSIsContainer) {
                    $Size = (Get-ChildItem $Item.FullName -Recurse -Force -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                } else {
                    $Size = $Item.Length
                }
                
                if ($Size -gt 100MB) {
                    $Stats = [PSCustomObject]@{
                        Name = $Item.FullName
                        SizeGB = [math]::Round($Size / 1GB, 2)
                    }
                    $FolderStats += $Stats
                }
            } catch { }
        }
        
        # Sort and build HTML Rows
        $TopFolders = $FolderStats | Sort-Object SizeGB -Descending | Select-Object -First 5
        foreach ($f in $TopFolders) {
            $StorageTableRows += "<tr><td>$($f.Name)</td><td>$($f.SizeGB) GB</td></tr>"
        }
        
    } catch {
        $StorageTableRows = "<tr><td colspan='2'>Error calculating storage details.</td></tr>"
    }
}

# 3. Build HTML Section (With CSS Pie Chart)
$StorageReportHTML = @"
<h2>Storage Analysis</h2>
<div class="section">
    <div style="display: flex; align-items: center; justify-content: space-around; flex-wrap: wrap;">
        <div style="
            width: 160px; 
            height: 160px; 
            border-radius: 50%; 
            background: conic-gradient(#d9534f 0% $($StoragePercent)%, #5cb85c $($StoragePercent)% 100%);
            border: 4px solid #fff;
            box-shadow: 0 0 10px rgba(0,0,0,0.1);
            position: relative;
        ">
            <div style="
                position: absolute; top: 50%; left: 50%; 
                transform: translate(-50%, -50%); 
                font-weight: bold; font-size: 1.2em; color: #fff; text-shadow: 0px 0px 3px #000;">
                $($StoragePercent)% Used
            </div>
        </div>
        
        <div style="padding: 10px;">
            <ul style="list-style: none; padding: 0;">
                <li style="margin-bottom:5px;"><span style="display:inline-block; width:15px; height:15px; background:#d9534f; margin-right:5px;"></span><strong>Used Space:</strong> $StorageUsedGB GB</li>
                <li style="margin-bottom:5px;"><span style="display:inline-block; width:15px; height:15px; background:#5cb85c; margin-right:5px;"></span><strong>Free Space:</strong> $StorageFreeGB GB</li>
                <li style="border-top:1px solid #ccc; padding-top:5px; margin-top:5px;"><strong>Total Capacity:</strong> $StorageTotalGB GB</li>
            </ul>
        </div>
    </div>

    <h3 style="margin-top: 20px; font-size: 1em; color: #444; border-bottom: 1px solid #ddd; padding-bottom: 5px;">Largest Items (Root Directory)</h3>
    <table style="width: 100%; border-collapse: collapse; font-size: 0.9em;">
        <thead>
            <tr style="background: #eee; text-align: left;">
                <th style="padding: 8px; border-bottom: 1px solid #ddd;">Directory / File</th>
                <th style="padding: 8px; border-bottom: 1px solid #ddd;">Size</th>
            </tr>
        </thead>
        <tbody>
            $StorageTableRows
        </tbody>
    </table>
</div>
"@

Write-Host "   > Done." -ForegroundColor Green

# --- 9. INSTALLED APPLICATIONS ---
Write-Host "`n[INFO] Listing Installed Applications (Instant)..." -ForegroundColor Yellow
if ($DemoMode) {
    $AppListHTML = "<li>Microsoft Office 365</li><li>Google Chrome</li><li>Adobe Acrobat Reader</li><li>Zoom Workplace</li><li>7-Zip 23.01</li><li>VLC Media Player</li><li>Microsoft Teams</li>"
} else {
    $UninstallKeys = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    $Apps = foreach ($key in $UninstallKeys) {
        if (Test-Path $key) {
            Get-ChildItem $key | Get-ItemProperty -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -and $_.UninstallString } | Select-Object -ExpandProperty DisplayName
        }
    }
    $AppListHTML = ($Apps | Sort-Object -Unique | ForEach-Object { "<li>$_</li>" }) -join ""
}
Write-Host "   > Done." -ForegroundColor Green

# --- 10. DISM & SFC SCANS ---
if (Test-UserSkip -StepName "DISM & SFC Scans" -Seconds 5) {
    $ReportData.SystemHealth = "Skipped by Engineer."
} else {
    if ($DemoMode) {
        Write-Host "      [55%] DISM CheckHealth..." -NoNewline; Start-Sleep -Seconds 1; Write-Host " OK" -ForegroundColor Green
        Write-Host "      [75%] SFC /Scannow..." -NoNewline; Start-Sleep -Seconds 2; Write-Host " OK" -ForegroundColor Green
        $ReportData.SystemHealth = "Issues Found & Repaired (See logs below)."
        $SfcLogContent = "2024-03-15 [SR] Repairing corrupted file \SystemRoot\Win32\webio.dll"
    } else {
        # DISM
        Write-Host "   > Running DISM RestoreHealth..." -ForegroundColor Yellow
        $DismOut = Dism /Online /Cleanup-Image /RestoreHealth 2>&1 | Out-String
        
        # SFC
        Write-Host "   > Running SFC /Scannow..." -ForegroundColor Yellow
        $SfcOut = sfc /scannow 2>&1 | Out-String
        
        # Analyze Results
        $HealthStatus = "Unknown"
        if ($SfcOut -match "found no integrity violations") {
            $HealthStatus = "Healthy (No integrity violations found)."
        }
        elseif ($SfcOut -match "found corrupt files and successfully repaired") {
            $HealthStatus = "<span style='color:green'><strong>FIXED:</strong> Corrupt files found and successfully repaired.</span>"
            try {
                $CBSLog = "$env:windir\Logs\CBS\CBS.log"
                if (Test-Path $CBSLog) {
                    $SfcLogContent = Get-Content $CBSLog | Select-String "\[SR\]" | Select-Object -Last 20 | Out-String
                }
            } catch { $SfcLogContent = "Error reading CBS.log: $_" }
            
        } elseif ($SfcOut -match "found corrupt files but was unable to fix") {
            $HealthStatus = "<span style='color:red'><strong>CRITICAL:</strong> Corrupt files found but Windows could NOT fix them. Manual intervention required.</span>"
            try {
                $CBSLog = "$env:windir\Logs\CBS\CBS.log"
                if (Test-Path $CBSLog) {
                    $SfcLogContent = Get-Content $CBSLog | Select-String "\[SR\]" | Select-Object -Last 20 | Out-String
                }
            } catch { $SfcLogContent = "Error reading CBS.log: $_" }
        } else {
            $HealthStatus = "Scan Completed (Check logs below for details)."
        }
        
        if ($DismOut -match "The restore operation completed successfully") {
            $HealthStatus += "<br>(DISM RestoreHealth completed successfully)"
        }
        
        $ReportData.SystemHealth = $HealthStatus
    }
    Write-Host "   > Done." -ForegroundColor Green
}

# --- 11. WINGET SOFTWARE UPDATES ---
if (Test-UserSkip -StepName "Software Updates (Winget)" -Seconds 5) {
    $ReportData.WingetStatus = "Skipped by Engineer."
} else {
    Write-Host "   > Detecting available updates..." -ForegroundColor Yellow
    if ($DemoMode) {
        Start-Sleep -Seconds 2
        $ReportData.WingetStatus = "Success: Packages updated."
    } else {
        try {
            if (Get-Command "winget" -ErrorAction SilentlyContinue) {
                $UpgradeListRaw = winget upgrade | Out-String
                $PackagesToUpdate = @()
                
                if ($UpgradeListRaw -match "No installed package found matching input criteria") {
                    $ReportData.WingetStatus = "Status OK: No software updates found."
                } else {
                    $Lines = $UpgradeListRaw -split "`n" | Where-Object { $_ -notmatch "Name|---" -and $_.Trim() -ne "" }
                    foreach ($Line in $Lines) {
                        $PackagesToUpdate += "<li>$($Line.Trim().Substring(0, [math]::Min($Line.Length, 40)).Trim())...</li>"
                    }
                    if ($PackagesToUpdate.Count -gt 0) {
                        $PackageHTML = "<ul style='margin-top:5px; margin-bottom:5px; font-size:0.9em;'>$($PackagesToUpdate -join '')</ul>"
                        Write-Host "   > Installing updates..." -ForegroundColor Yellow
                        $WingetProcess = Start-Process -FilePath "winget" -ArgumentList "upgrade --all --accept-source-agreements --accept-package-agreements" -Wait -PassThru -NoNewWindow
                        if ($WingetProcess.ExitCode -eq 0) {
                             $ReportData.WingetStatus = "Success: The following packages were updated:<br>$PackageHTML"
                        } else {
                             $ReportData.WingetStatus = "Warning: Winget returned code $($WingetProcess.ExitCode).<br>Targeted packages:<br>$PackageHTML"
                        }
                    } else {
                        $ReportData.WingetStatus = "Status OK: No software updates required."
                    }
                }
            } else {
                 $ReportData.WingetStatus = "Failed: Winget command not found."
            }
        } catch {
            $ReportData.WingetStatus = "Error: Winget execution failed."
        }
    }
    Write-Host "   > Done." -ForegroundColor Green
}

# --- 12. WINDOWS UPDATES (PENDING LIST) ---
if (Test-UserSkip -StepName "Windows Update Check") {
    $ReportData.Updates = "Skipped by Engineer."
} else {
    Write-Host "   > Checking for updates..." -ForegroundColor Yellow
    if ($DemoMode) {
        $ReportData.Updates = "Action Required: 2 updates pending."
    } else {
        try {
            $UpdateSession = New-Object -ComObject Microsoft.Update.Session
            $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
            $SearchResult = $UpdateSearcher.Search("IsInstalled=0")
            $Count = $SearchResult.Updates.Count
            
            if ($Count -gt 0) { 
                $PendingUpdates = @()
                for ($i = 0; $i -lt $Count; $i++) {
                    $UpdateItem = $SearchResult.Updates.Item($i)
                    $PendingUpdates += "<li>$($UpdateItem.Title)</li>"
                }
                $UpdateListHTML = "<ul style='margin-top:5px; margin-bottom:5px; font-size:0.9em; color:#d9534f;'>$($PendingUpdates -join '')</ul>"
                $ReportData.Updates = "Action Required: $Count updates pending.<br>$UpdateListHTML" 
            } 
            else { 
                $ReportData.Updates = "Status OK: System is up to date." 
            }
        } catch {
            $ReportData.Updates = "Manual Check Required (Error accessing COM Object)."
        }
    }
    Write-Host "   > Done." -ForegroundColor Green
}

# --- 13. DISK OPTIMIZATION (UPDATED LOGIC) ---
Write-Host "`n   [NEXT STEP] Disk Defragmentation" -ForegroundColor Cyan

# Check if C: drive is SSD/NVMe or HDD
if ($DemoMode) { $IsSSD = $false }
else {
    try {
        # Determine media type of System Drive (C:)
        $SystemDisk = Get-Partition -DriveLetter C -ErrorAction SilentlyContinue | Get-Disk | Get-PhysicalDisk
        # NVMe drives report as SSD MediaType in Powershell (BusType is NVMe)
        $IsSSD = ($SystemDisk.MediaType -eq 'SSD')
    } catch { $IsSSD = $false }
}

if ($IsSSD) {
    Write-Host "   > Drive Type Detected: SSD/NVMe (Solid State Drive)" -ForegroundColor Green
    Write-Host "   > SKIPPING DEFRAGMENTATION (Not required for this drive type)." -ForegroundColor Green
    $ReportData.Defrag = "Scan not required for this drive type"
} else {
    # It is a HDD (or unknown), so run the existing prompt/timer logic
    Write-Host "   > Drive Type Detected: HDD (Hard Disk Drive)" -ForegroundColor Yellow
    Write-Host "   [Y] Yes  |  [N] No (Default)  |  [Enter] Default" -ForegroundColor Gray

    # Flush buffer
    while ([Console]::KeyAvailable) { $null = [Console]::ReadKey($true) }

    $OptimizeChoice = $null
    $Timeout = 60
    $EndTime = (Get-Date).AddSeconds($Timeout)

    while ((Get-Date) -lt $EndTime) {
        $Remaining = [math]::Ceiling(($EndTime - (Get-Date)).TotalSeconds)
        Write-Host "`r   > Waiting for input (Default: No) in $Remaining seconds...   " -NoNewline -ForegroundColor Yellow
        
        if ([Console]::KeyAvailable) {
            $KeyInfo = [Console]::ReadKey($true)
            $Key = $KeyInfo.Key.ToString().ToUpper()
            
            if ($Key -eq "Y") { $OptimizeChoice = "Y"; break }
            if ($Key -eq "N") { $OptimizeChoice = "N"; break }
            if ($Key -eq "ENTER") { $OptimizeChoice = "ENTER"; break }
        }
        Start-Sleep -Milliseconds 100
    }

    # Default to 'N' if timeout occurred
    if ([string]::IsNullOrWhiteSpace($OptimizeChoice)) { 
        $OptimizeChoice = 'TIMEOUT' 
        Write-Host "`r   > Timeout reached. Selecting Default: No.                  " -ForegroundColor Yellow
    } else {
        Write-Host "`r                                                              " -NoNewline
    }

    if ($OptimizeChoice -eq 'Y') {
        Write-Host "`r   > Option Selected: YES                                     " -ForegroundColor Green
        if ($DemoMode) {
            Start-Sleep -Seconds 1
            $ReportData.Defrag = "Optimization Complete."
        } else {
            Write-Host "   > Optimizing Drive C:..." -ForegroundColor Yellow
            Optimize-Volume -DriveLetter C -Analyze -Verbose | Out-Null
            Optimize-Volume -DriveLetter C -Defrag -Verbose | Out-Null
            $ReportData.Defrag = "Optimization Complete: Drive C: optimized."
        }
        Write-Host "   > Done." -ForegroundColor Green

    } elseif ($OptimizeChoice -eq 'ENTER') {
        Write-Host "`r   > Option Selected: DEFAULT (No Input)                      " -ForegroundColor Yellow
        $ReportData.Defrag = "Skipped: No Input"

    } else {
        Write-Host "`r   > SKIPPED: Disk Optimization.                              " -ForegroundColor Yellow
        $ReportData.Defrag = "Skipped by Engineer."
    }
}

# --- 14. GENERATE REPORT ---
Write-Host "`n[REPORT] Generating Report..." -ForegroundColor Yellow

# CHECK INTERNET FOR REPORT
if ($DemoMode) {
    $ReportData.Internet = "Active (Verified)"
} elseif (Test-Connection -ComputerName 8.8.8.8 -Count 1 -Quiet -ErrorAction SilentlyContinue) {
    $ReportData.Internet = "Active (Verified)"
} else {
    $ReportData.Internet = "Disconnected"
}

# Build Log HTML Section if content exists
$LogSectionHTML = ""
if (-not [string]::IsNullOrWhiteSpace($SfcLogContent)) {
    $LogSectionHTML = @"
    <h2>Diagnostic Logs (SFC Details)</h2>
    <div class="section">
        <div style="background:#f0f0f0; padding:10px; border:1px solid #ccc; font-family:Consolas, monospace; font-size:0.8em; white-space:pre-wrap;">$SfcLogContent</div>
    </div>
"@
}

# SAVE PATHS
$HtmlFile = "$BaseDir\$ReportFilename.html"
$PdfFile  = "$BaseDir\$ReportFilename.pdf"
$PublicDesktopPdf = "$env:PUBLIC\Desktop\$ReportFilename.pdf"
$ModeLabel = if ($DemoMode) { "(DEMO MODE)" } else { "" }

$HtmlContent = @"
<!DOCTYPE html>
<html>
<head>
<style>
    body { font-family: 'Segoe UI', sans-serif; color: #333; padding: 20px; }
    .header { text-align: center; margin-bottom: 20px; }
    .header img { max-height: 100px; }
    h1 { color: #0056b3; margin-bottom: 5px; }
    .meta { font-size: 0.9em; color: #666; text-align: center; margin-bottom: 30px; }
    .section { background: #f9f9f9; padding: 15px; border-left: 6px solid #0056b3; margin-bottom: 20px; }
    .item { margin-bottom: 12px; border-bottom: 1px solid #e0e0e0; padding-bottom: 8px; }
    .item:last-child { border-bottom: none; }
    .label { font-weight: bold; color: #444; display: block; margin-bottom: 2px; }
    ul { padding-left: 20px; margin: 0; }
    .app-list ul { font-size: 0.8em; columns: 2; -webkit-columns: 2; -moz-columns: 2; }
</style>
</head>
<body>
<div class="header">
    <img src="$LogoUrl" alt="Apollo Technology" onerror="this.style.display='none'">
    <h1>Health Check Report $ModeLabel</h1>
    <p>This report has been generated by <strong>$EngineerName</strong> for ticket (<strong>$TicketNumber</strong>) for <strong>$CustomerName</strong></p>
    <div class="meta">
        <strong>Report Date:</strong> $CurrentDate
    </div>
</div>

<h2>Device Information</h2>
<div class="section">
    <div class="item"><span class="label">Device Name/Hostname:</span> $($ComputerInfo.Name)</div>
    <div class="item"><span class="label">OS Version:</span> $($OSInfo.Caption)</div>
    <div class="item"><span class="label">OS:</span> $($OSInfo.OSArchitecture) ($($OSInfo.Version))</div>
    <div class="item"><span class="label">RAM Size:</span> $([math]::Round($ComputerInfo.TotalPhysicalMemory / 1GB, 0)) GB</div>
    <div class="item"><span class="label">CPU Type/Model:</span> $($CPUInfo.Name)</div>
    <div class="item"><span class="label">Number of Processors:</span> $($CPUInfo.NumberOfLogicalProcessors) Threads / $($CPUInfo.NumberOfCores) Cores</div>
    <div class="item"><span class="label">Date on Windows Install:</span> $FormattedInstallDate</div>
    <div class="item">
        <span class="label">Disk Drives:</span>
        <ul>$DriveListHTML</ul>
    </div>
</div>

<h2>System Status</h2>
<div class="section">
    <div class="item"><span class="label">Internet Connection:</span> $($ReportData.Internet)</div>
    <div class="item"><span class="label">Battery Status:</span> $($ReportData.Battery)</div>
    <div class="item"><span class="label">Restore Point:</span> $($ReportData.RestorePoint)</div>
</div>

$StorageReportHTML

<h2>Maintenance & Updates</h2>
<div class="section">
    <div class="item"><span class="label">System Integrity (SFC/DISM):</span> $($ReportData.SystemHealth)</div>
    <div class="item"><span class="label">Software Upgrades (Winget):</span> $($ReportData.WingetStatus)</div>
    <div class="item"><span class="label">Windows Updates:</span> $($ReportData.Updates)</div>
    <div class="item"><span class="label">Storage Cleanup:</span> $($ReportData.DiskCleanup)</div>
    <div class="item"><span class="label">Disk Optimization:</span> $($ReportData.Defrag)</div>
</div>

<h2>Installed Software Inventory</h2>
<div class="section app-list"><ul>$AppListHTML</ul></div>

$LogSectionHTML

<p style="text-align:center; font-size:0.8em; color:#888; margin-top:50px;">&copy; $CurrentYear by Apollo Technology. All rights reserved | This tool has been created by Lewis Wiltshire (Apollo Technology)</p>
</body>
</html>
"@

# 15. Save HTML
$HtmlContent | Out-File -FilePath $HtmlFile -Encoding UTF8
Start-Sleep -Seconds 1

# 16. Convert to PDF using Edge
$EdgeLoc1 = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
$EdgeLoc2 = "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
$EdgeExe = if (Test-Path $EdgeLoc1) { $EdgeLoc1 } elseif (Test-Path $EdgeLoc2) { $EdgeLoc2 } else { $null }

if ($EdgeExe) {
    Write-Host "   > Converting to PDF..." -ForegroundColor Cyan
    $EdgeUserData = "$BaseDir\EdgeTemp"
    if (-not (Test-Path $EdgeUserData)) { New-Item -Path $EdgeUserData -ItemType Directory -Force | Out-Null }
    try {
        $Process = Start-Process -FilePath $EdgeExe -ArgumentList "--headless", "--disable-gpu", "--print-to-pdf=`"$PdfFile`"", "--no-pdf-header-footer", "--user-data-dir=`"$EdgeUserData`"", "`"$HtmlFile`"" -PassThru -Wait
        Start-Sleep -Seconds 2 
        if (Test-Path $PdfFile) {
            Write-Host "   > Success! Report Generated." -ForegroundColor Green
            Copy-Item -Path $PdfFile -Destination $PublicDesktopPdf -Force
            Remove-Item $HtmlFile -ErrorAction SilentlyContinue
            Remove-Item $EdgeUserData -Recurse -Force -ErrorAction SilentlyContinue
        } else { throw "PDF not found" }
    } catch {
        Write-Warning "PDF Conversion failed. Opening HTML instead."
        Copy-Item -Path $HtmlFile -Destination "$env:PUBLIC\Desktop\$ReportFilename.html" -Force
        Start-Process "$env:PUBLIC\Desktop\$ReportFilename.html"
        $PdfFile = $null
    }
} else {
    Write-Warning "Edge not found. Saving HTML report."
    Copy-Item -Path $HtmlFile -Destination "$env:PUBLIC\Desktop\$ReportFilename.html" -Force
    Start-Process "$env:PUBLIC\Desktop\$ReportFilename.html"
    $PdfFile = $null
}

# --- 17. EMAIL REPORT ---
if ($EmailEnabled -and $PdfFile -and (Test-Path $PdfFile)) {
    Write-Host "`nSending Email to $ToAddress..." -ForegroundColor Yellow
    try {
        Send-MailMessage -From $FromAddress -To $ToAddress -Subject "Health Check Report: $env:COMPUTERNAME" -Body "Report attached." -SmtpServer $SmtpServer -Port $SmtpPort -UseSsl $UseSSL -Credential $EmailCreds -Attachments $PdfFile -ErrorAction Stop
        Write-Host "   > Email Sent Successfully!" -ForegroundColor Green
    } catch {
        Write-Error "   > Failed to send email. Error: $_"
    }
}

# --- ALLOW SLEEP AGAIN ---
try { [SleepUtils]::SetThreadExecutionState(0x80000000) | Out-Null } catch { }

Write-Host "`n------------------------------------------------------------"
Write-Host "          PROCESS COMPLETED" -ForegroundColor Green
if (Test-Path $PublicDesktopPdf) { Start-Process $PublicDesktopPdf; Start-Sleep -Seconds 1 }
Write-Host "------------------------------------------------------------"
exit