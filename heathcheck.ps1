<#
.SYNOPSIS
    Apollo Technology Full Health Check Script v16.2
.DESCRIPTION
    Full system health check.
    - NEW: Prevents Windows from Sleeping/Locking while running.
    - Report saves to C:\temp\Apollo_Reports.
    - Storage Analysis with Pie Chart & Top Folder usage.
    - Visual Progress Bars for every step (0-100% in steps of 5).
    - Captures and embeds SFC [SR] logs into the report.
    - Internet Check visualizes properly in Demo Mode.
    - Disk Defragmentation has 60s auto-skip.
    - Demo Mode generates rich "dummy" data including logs.
    - PATCH: Added check for cleanmgr.exe to prevent crash on Server 2008.
.NOTES
    Author: Apollo Technology
    Additional Notes: Requires PowerShell 5.1+, Windows 10/11 (Server Compatible)
    Author Name: Lewis Wiltshire
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

# --- 2.5 PREVENT SLEEP (NEW) ---
# Uses Windows API to prevent the system from idling to sleep or turning off display
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
    # ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
    # 0x80000000 | 0x00000001 | 0x00000002 = 0x80000003
    $null = [SleepUtils]::SetThreadExecutionState(0x80000003)
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

# --- HELPER FUNCTION: PROGRESS BAR ---
function Show-StepProgress {
    param(
        [string]$Activity, 
        [int]$StepDelay = 30 # Milliseconds between increments
    )
    
    # Loop from 0 to 100 in steps of 5
    for ($i = 0; $i -le 100; $i += 5) {
        Write-Progress -Activity $Activity -Status "Processing... $i%" -PercentComplete $i
        Start-Sleep -Milliseconds $StepDelay
    }
    # Clean up bar
    Write-Progress -Activity $Activity -Completed
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

$EngineerName = Read-Host "Please enter the Engineer Name"

# --- INTERNET CHECK (STARTUP) ---
Write-Host "`n   [CHECK] Verifying Internet Connection..." -ForegroundColor Yellow
Show-StepProgress -Activity "Internet Connection" -StepDelay 20

while ($true) {
    if ($DemoMode) { 
        Write-Host "`r   > Simulating connection check...       " -NoNewline -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        Write-Host "`r   > Connection Verified (DEMO).          " -ForegroundColor Green
        break 
    }
    if (Test-Connection -ComputerName 8.8.8.8 -Count 10 -Quiet -ErrorAction SilentlyContinue) {
        Write-Host "`r   > Connection Verified.                 " -ForegroundColor Green
        break
    }
    Write-Host "`r   > Waiting for internet connection...   " -NoNewline -ForegroundColor Red
    Start-Sleep -Seconds 3
}

$CurrentDate = Get-Date -Format "yyyy-MM-dd HH:mm"
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

# --- 4. SYSTEM RESTORE POINT ---
if (Test-UserSkip -StepName "Restore Point Creation") {
    $ReportData.RestorePoint = "Skipped by Engineer."
} else {
    Show-StepProgress -Activity "Creating System Restore Point" -StepDelay 50
    
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

# --- 5. BATTERY STATUS (Instant) ---
Write-Host "`n[INFO] Checking Battery Health (Instant)..." -ForegroundColor Yellow
Show-StepProgress -Activity "Checking Battery Health" -StepDelay 10

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

# --- 6. DISK CLEANUP (PATCHED FOR SERVER 2008) ---
if (Test-UserSkip -StepName "Disk Cleanup") {
    $ReportData.DiskCleanup = "Skipped by Engineer."
} else {
    Write-Host "   > Cleaning System Files..." -ForegroundColor Yellow
    Show-StepProgress -Activity "Disk Cleanup (Cleanmgr)" -StepDelay 40
    
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
            # Tool missing (common on Server 2008 without Desktop Experience)
            Write-Warning "Disk Cleanup (cleanmgr.exe) not found on this system."
            $ReportData.DiskCleanup = "Skipped: Disk Cleanup utility not installed."
        }
    }
    Write-Host "   > Done." -ForegroundColor Green
}

# --- 6.5 STORAGE ANALYSIS (NEW) ---
Write-Host "`n[INFO] Analyzing Storage Usage..." -ForegroundColor Yellow
Show-StepProgress -Activity "Storage Analysis" -StepDelay 20

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
                    # Measure directory size (recurse)
                    # Limit depth or errors to avoid hanging forever on protected folders
                    $Size = (Get-ChildItem $Item.FullName -Recurse -Force -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                } else {
                    # It is a file (e.g., pagefile.sys)
                    $Size = $Item.Length
                }
                
                if ($Size -gt 100MB) { # Only care about things > 100MB
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

# --- 7. INSTALLED APPLICATIONS (Instant) ---
Write-Host "`n[INFO] Listing Installed Applications (Instant)..." -ForegroundColor Yellow
Show-StepProgress -Activity "Software Inventory" -StepDelay 10

if ($DemoMode) {
    # Dummy list for report
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

# --- 8. DISM & SFC SCANS (WITH LOG EXTRACTION) ---
if (Test-UserSkip -StepName "DISM & SFC Scans" -Seconds 5) {
    $ReportData.SystemHealth = "Skipped by Engineer."
} else {
    Show-StepProgress -Activity "Initializing System Scans" -StepDelay 30

    if ($DemoMode) {
        Write-Host "      [55%] DISM CheckHealth..." -NoNewline; Start-Sleep -Seconds 1; Write-Host " OK" -ForegroundColor Green
        Write-Host "      [75%] SFC /Scannow..." -NoNewline; Start-Sleep -Seconds 2; Write-Host " OK" -ForegroundColor Green
        
        $ReportData.SystemHealth = "Issues Found & Repaired (See logs below)."
        # Generate Dummy Logs
        $SfcLogContent = @"
2024-03-15 10:14:22, Info                  CSI    0000021a [SR] Repairing 2 components
2024-03-15 10:14:22, Info                  CSI    0000021b [SR] Beginning Verify and Repair transaction
2024-03-15 10:14:23, Info                  CSI    0000021c [SR] Repairing corrupted file \SystemRoot\Win32\webio.dll from store
2024-03-15 10:14:23, Info                  CSI    0000021d [SR] Repairing corrupted file \SystemRoot\Win32\tcpip.sys from store
2024-03-15 10:14:24, Info                  CSI    0000021e [SR] Repair complete
"@
    } else {
        # DISM
        Write-Host "   > Running DISM RestoreHealth..." -ForegroundColor Yellow
        $DismOut = Dism /Online /Cleanup-Image /RestoreHealth 2>&1 | Out-String
        
        # SFC
        Write-Host "   > Running SFC /Scannow..." -ForegroundColor Yellow
        $SfcOut = sfc /scannow 2>&1 | Out-String
        
        # Analyze Results
        $HealthStatus = "Unknown"
        
        # Check SFC Output
        if ($SfcOut -match "found corrupt files and successfully repaired") {
            $HealthStatus = "Issues Found & Repaired (Corrupt files fixed)."
            # Extract Logs
            try {
                $CBSLog = "$env:windir\Logs\CBS\CBS.log"
                if (Test-Path $CBSLog) {
                    $SfcLogContent = Get-Content $CBSLog | Select-String "\[SR\]" | Select-Object -Last 20 | Out-String
                }
            } catch { $SfcLogContent = "Error reading CBS.log: $_" }
            
        } elseif ($SfcOut -match "found corrupt files but was unable to fix") {
            $HealthStatus = "WARNING: Corrupt files found but repair FAILED."
            try {
                $CBSLog = "$env:windir\Logs\CBS\CBS.log"
                if (Test-Path $CBSLog) {
                    $SfcLogContent = Get-Content $CBSLog | Select-String "\[SR\]" | Select-Object -Last 20 | Out-String
                }
            } catch { $SfcLogContent = "Error reading CBS.log: $_" }
            
        } elseif ($SfcOut -match "did not find any integrity violations") {
            $HealthStatus = "Healthy: No integrity violations found."
        } else {
            $HealthStatus = "Scan Completed (See logs for details)."
        }
        
        if ($DismOut -match "The restore operation completed successfully") {
            $HealthStatus += "<br>(DISM RestoreHealth completed successfully)"
        }
        
        $ReportData.SystemHealth = $HealthStatus
    }
    Write-Host "   > Done." -ForegroundColor Green
}

# --- 9. WINGET SOFTWARE UPDATES ---
if (Test-UserSkip -StepName "Software Updates (Winget)" -Seconds 5) {
    $ReportData.WingetStatus = "Skipped by Engineer."
} else {
    Write-Host "   > Detecting available updates..." -ForegroundColor Yellow
    Show-StepProgress -Activity "Checking Winget Repositories" -StepDelay 30

    if ($DemoMode) {
        Start-Sleep -Seconds 2
        # Dummy Winget Data
        $ReportData.WingetStatus = "Success: The following packages were updated:<br><ul style='margin-top:5px; margin-bottom:5px; font-size:0.9em;'><li>Google Chrome (v118 -> v119)</li><li>Zoom Workplace (v5.16 -> v5.17)</li><li>VLC Media Player (v3.0.18 -> v3.0.20)</li></ul>"
    } else {
        try {
            if (Get-Command "winget" -ErrorAction SilentlyContinue) {
                # Get List
                $UpgradeListRaw = winget upgrade | Out-String
                $PackagesToUpdate = @()
                
                if ($UpgradeListRaw -match "No installed package found matching input criteria") {
                    $ReportData.WingetStatus = "Status OK: No software updates found."
                } else {
                    # Parse List
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

# --- 10. WINDOWS UPDATES (PENDING LIST) ---
if (Test-UserSkip -StepName "Windows Update Check") {
    $ReportData.Updates = "Skipped by Engineer."
} else {
    Write-Host "   > Checking for updates..." -ForegroundColor Yellow
    Show-StepProgress -Activity "Windows Update Check" -StepDelay 30

    if ($DemoMode) {
        # Dummy Windows Update Data
        $ReportData.Updates = "Action Required: 2 updates pending.<br><ul style='margin-top:5px; margin-bottom:5px; font-size:0.9em; color:#d9534f;'><li>2024-02 Cumulative Update for Windows 11 (KB5034765)</li><li>Windows Defender Antivirus Security Intelligence (KB2267602)</li></ul>"
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

# --- 11. DISK OPTIMIZATION (WITH 60s TIMER) ---
Write-Host "`n   [NEXT STEP] Disk Defragmentation" -ForegroundColor Cyan
Write-Host "   [Y] Yes  |  [N] No (Default)" -ForegroundColor Gray

# Flush buffer
while ([Console]::KeyAvailable) { $null = [Console]::ReadKey($true) }

$OptimizeChoice = $null
$Timeout = 60
$EndTime = (Get-Date).AddSeconds($Timeout)

while ((Get-Date) -lt $EndTime) {
    $Remaining = [math]::Ceiling(($EndTime - (Get-Date)).TotalSeconds)
    Write-Host "`r   > Waiting for input (Default: No) in $Remaining seconds...   " -NoNewline -ForegroundColor Yellow
    
    if ([Console]::KeyAvailable) {
        $Key = [Console]::ReadKey($true).Key.ToString().ToUpper()
        if ($Key -eq "Y") { $OptimizeChoice = "Y"; break }
        if ($Key -eq "N") { $OptimizeChoice = "N"; break }
    }
    Start-Sleep -Milliseconds 100
}

# Default to 'N' if timeout occurred
if ([string]::IsNullOrWhiteSpace($OptimizeChoice)) { 
    $OptimizeChoice = 'N' 
    Write-Host "`r   > Timeout reached. Selecting Default: No.                  " -ForegroundColor Yellow
} else {
    Write-Host "`r   > Option Selected: $OptimizeChoice                         " -ForegroundColor Yellow
}

if ($OptimizeChoice -eq 'Y') {
    Show-StepProgress -Activity "Optimizing Drive" -StepDelay 40

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
} else {
    Write-Host "   > SKIPPED: Disk Optimization." -ForegroundColor Yellow
    $ReportData.Defrag = "Skipped by Engineer."
}

# --- 12. GENERATE REPORT ---
Write-Host "`n[REPORT] Generating Report..." -ForegroundColor Yellow
Show-StepProgress -Activity "Generating Report" -StepDelay 20

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
    <div class="meta">
        <strong>Engineer:</strong> $EngineerName &nbsp;|&nbsp; 
        <strong>Date:</strong> $CurrentDate &nbsp;|&nbsp; 
        <strong>Device:</strong> $env:COMPUTERNAME
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

<p style="text-align:center; font-size:0.8em; color:#888; margin-top:50px;">This report has been generated by $EngineerName at Apollo Technology using the Health Check Tool</p>
</body>
</html>
"@

# 1. Save HTML
$HtmlContent | Out-File -FilePath $HtmlFile -Encoding UTF8
Start-Sleep -Seconds 1

# 2. Convert to PDF using Edge
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

# --- 13. EMAIL REPORT ---
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
# Revert execution state (Good practice, though exit also clears it)
try { [SleepUtils]::SetThreadExecutionState(0x80000000) | Out-Null } catch { }

Write-Host "`n------------------------------------------------------------"
Write-Host "          PROCESS COMPLETED" -ForegroundColor Green
if (Test-Path $PublicDesktopPdf) { Start-Process $PublicDesktopPdf; Start-Sleep -Seconds 1 }
Write-Host "------------------------------------------------------------"
exit