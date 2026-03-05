<#
.SYNOPSIS
    Apollo Technology Full Health Check Script v17.9.2 (Parser-Safe Edition)
.DESCRIPTION
    Full system health check.
    - REMOVED: Temperature sensors.
    - RETAINED: GPU Detection in Device Details.
    - NEW: Red "[NOTICE] Running in Elevated Permissions" warning on startup.
    - UPDATED: Disk Optimization now auto-skips on SSD/NVMe drives.
    - NEW: NVMe/SSD/HDD Detection with SMART Status.
    - NEW: Detailed Device Information Section.
    - NEW: Customer/Ticket Input Validation Loop.
    - NEW: Added Network Adapter Information to Device Info.
    - UPDATED: Split DISM and SFC into separate logic blocks.
    - RETAINED: Prevents Windows from Sleeping/Locking while running.
    - UPDATED: Report strictly saves to C:\temp\Apollo_Reports.
    - FIXED: Winget execution logic with prompt bypasses.
    - NEW UI: Report HTML/PDF generation completely refactored to match Hardware Diagnostics style.
    - PATCHED: Winget Mojibake/Progress bar artifacts strictly filtered.
    - PATCHED: Removed fragile subexpressions and backticks to prevent copy-paste AST parser crashes.
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

# --- ADDED: ELEVATED PERMISSIONS NOTICE ---
if ($isAdmin) { 
    Write-Host "        [NOTICE] Running in Elevated Permissions" -ForegroundColor Red 
} elseif ($DemoMode) {
    Write-Host "      [NOTICE] Running as Standard User" -ForegroundColor Yellow 
}

Write-Host "        Created by Lewis Wiltshire, Version 17.9.2" -ForegroundColor Yellow
Write-Host "      [POWER] Sleep Mode & Screen Timeout Blocked." -ForegroundColor DarkGray

# --- STRICT PATH SETUP ---
$BaseDir = "C:\temp\Apollo_Reports"

try {
    if (-not (Test-Path $BaseDir)) { New-Item -Path $BaseDir -ItemType Directory -Force -ErrorAction Stop | Out-Null }
} catch {
    Write-Error "Could not create path: $BaseDir. Please ensure you have correct permissions."
    exit
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

# --- FILE NAMING LOGIC ---
$TicketForFile = $TicketNumber -replace "#",""
$CustomerForFile = $CustomerName -replace '[\\/:*?"<>|]', '' 
$ReportFilename = "Health_Check_$($CustomerForFile)_$($TicketForFile)"

$ReportData = @{}
$SfcLogContent = "" 

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

# --- GPU DETECTION ---
try {
    $GPUInfo = Get-CimInstance Win32_VideoController | Select-Object -ExpandProperty Name -Unique
    $GPUString = $GPUInfo -join ", "
} catch {
    $GPUString = "Unknown / Standard VGA Adapter"
}

# --- NETWORK ADAPTER DETECTION ---
$NetworkAdaptersHTML = ""
try {
    $Adapters = Get-NetAdapter -Physical | Where-Object { $_.Status -eq 'Up' }
    foreach ($nic in $Adapters) {
        $ipInfo = Get-NetIPAddress -InterfaceAlias $nic.Name -AddressFamily IPv4 -ErrorAction SilentlyContinue
        $ip = if ($ipInfo) { $ipInfo.IPAddress } else { "No IPv4" }
        $NetworkAdaptersHTML += "<li><strong>$($nic.Name)</strong>: $($nic.InterfaceDescription) (IP: $ip)</li>"
    }
    if ([string]::IsNullOrWhiteSpace($NetworkAdaptersHTML)) {
        $NetworkAdaptersHTML = "<li>No active physical network adapters found.</li>"
    }
} catch {
    $NetworkAdaptersHTML = "<li>Unable to retrieve network adapter info.</li>"
}

# --- DRIVE DETECTION ---
$PhysicalDisks = Get-PhysicalDisk | Sort-Object DeviceId
$DriveListHTML = ""

foreach ($disk in $PhysicalDisks) {
    $Type = ""
    if ($disk.MediaType -eq "HDD") { $Type = "Hard Disk Drive (HDD)" }
    elseif ($disk.MediaType -eq "SSD") {
        if ($disk.BusType -eq "NVMe") { $Type = "NVMe Solid State Drive" } 
        else { $Type = "SATA Solid State Drive (SSD)" }
    }
    else { $Type = "Unknown ($($disk.MediaType))" }
    
    $SizeGB = [math]::Round($disk.Size / 1GB, 2)
    $SizeStr = if ($SizeGB -gt 1000) { "$([math]::Round($SizeGB / 1024, 2)) TB" } else { "$SizeGB GB" }
    $HealthColor = if ($disk.HealthStatus -eq "Healthy") { "green" } else { "red" }
    
    $DriveListHTML += "<li><strong>Disk $($disk.DeviceId):</strong> $Type ($SizeStr) | Health: <span style='color:$HealthColor'>$($disk.HealthStatus)</span></li>"
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
$StorageUsedGB = 0; $StorageFreeGB = 0; $StorageTotalGB = 0; $StoragePercent = 0; $StorageTableRows = ""

if ($DemoMode) {
    Start-Sleep -Seconds 2
    $StorageUsedGB = 340.5; $StorageFreeGB = 135.5; $StorageTotalGB = 476.0; $StoragePercent = 71
    $StorageTableRows = "<tr><td>C:\Users\Apollo</td><td>120.4 GB</td></tr><tr><td>C:\Windows</td><td>45.2 GB</td></tr><tr><td>C:\Program Files</td><td>32.1 GB</td></tr>"
} else {
    try {
        $Drive = Get-PSDrive C -ErrorAction SilentlyContinue
        if ($Drive) {
            $StorageUsedGB  = [math]::Round($Drive.Used / 1GB, 2)
            $StorageFreeGB  = [math]::Round($Drive.Free / 1GB, 2)
            $StorageTotalGB = [math]::Round(($Drive.Used + $Drive.Free) / 1GB, 2)
            if ($StorageTotalGB -gt 0) { $StoragePercent = [math]::Round(($StorageUsedGB / $StorageTotalGB) * 100) }
        }
        Write-Host "   > Scanning largest folders in C:\ (This may take a moment)..." -ForegroundColor Yellow
        $RootItems = Get-ChildItem "C:\" -Force -ErrorAction SilentlyContinue
        $FolderStats = @()
        foreach ($Item in $RootItems) {
            try {
                $Size = if ($Item.PSIsContainer) { (Get-ChildItem $Item.FullName -Recurse -Force -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum } else { $Item.Length }
                if ($Size -gt 100MB) { $FolderStats += [PSCustomObject]@{ Name = $Item.FullName; SizeGB = [math]::Round($Size / 1GB, 2) } }
            } catch { }
        }
        $TopFolders = $FolderStats | Sort-Object SizeGB -Descending | Select-Object -First 4
        foreach ($f in $TopFolders) { $StorageTableRows += "<tr><td>$($f.Name)</td><td>$($f.SizeGB) GB</td></tr>" }
    } catch {
        $StorageTableRows = "<tr><td colspan='2'>Error calculating storage details.</td></tr>"
    }
}
Write-Host "   > Done." -ForegroundColor Green

# --- 9. INSTALLED APPLICATIONS ---
Write-Host "`n[INFO] Listing Installed Applications (Instant)..." -ForegroundColor Yellow
if ($DemoMode) {
    $AppListHTML = "<li>Microsoft Office 365</li><li>Google Chrome</li><li>Adobe Acrobat Reader</li><li>Zoom Workplace</li><li>7-Zip 23.01</li><li>VLC Media Player</li><li>Microsoft Teams</li>"
} else {
    $UninstallKeys = @("HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall", "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall", "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall")
    $Apps = foreach ($key in $UninstallKeys) { if (Test-Path $key) { Get-ChildItem $key | Get-ItemProperty -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -and $_.UninstallString } | Select-Object -ExpandProperty DisplayName } }
    $AppListHTML = ($Apps | Sort-Object -Unique | ForEach-Object { "<li>$_</li>" }) -join ""
}
Write-Host "   > Done." -ForegroundColor Green

# --- 10a. DISM SCANS ---
if (Test-UserSkip -StepName "DISM Health Checks" -Seconds 5) {
    $ReportData.DismStatus = "Skipped by Engineer."
} else {
    if ($DemoMode) {
        Write-Host "      [33%] DISM CheckHealth..." -NoNewline; Start-Sleep -Seconds 1; Write-Host " OK" -ForegroundColor Green
        Write-Host "      [66%] DISM ScanHealth..." -NoNewline; Start-Sleep -Seconds 1; Write-Host " OK" -ForegroundColor Green
        Write-Host "      [100%] DISM RestoreHealth..." -NoNewline; Start-Sleep -Seconds 1; Write-Host " OK" -ForegroundColor Green
        $ReportData.DismStatus = "Success: No corruption found."
    } else {
        Write-Host "   > Running DISM /CheckHealth..." -ForegroundColor Yellow
        $DismCheck = Dism /Online /Cleanup-Image /CheckHealth 2>&1 | Out-String
        Write-Host "   > Running DISM /ScanHealth..." -ForegroundColor Yellow
        $DismScan = Dism /Online /Cleanup-Image /ScanHealth 2>&1 | Out-String
        Write-Host "   > Running DISM /RestoreHealth..." -ForegroundColor Yellow
        $DismRestore = Dism /Online /Cleanup-Image /RestoreHealth 2>&1 | Out-String
        
        $DismSummary = "<strong>CheckHealth:</strong> "
        if ($DismCheck -match "No component store corruption") { $DismSummary += "No corruption.<br>" } else { $DismSummary += "Corruption detected.<br>" }
        
        $DismSummary += "<strong>ScanHealth:</strong> "
        if ($DismScan -match "No component store corruption") { $DismSummary += "No corruption.<br>" } else { $DismSummary += "Issues found.<br>" }

        $DismSummary += "<strong>RestoreHealth:</strong> "
        if ($DismRestore -match "completed successfully") { $DismSummary += "Completed successfully." } else { $DismSummary += "Completed (Check logs)." }
        
        $ReportData.DismStatus = $DismSummary
    }
    Write-Host "   > DISM Checks Completed." -ForegroundColor Green
}

# --- 10b. SFC SCAN ---
if (Test-UserSkip -StepName "SFC (System File Checker)" -Seconds 5) {
    $ReportData.SfcStatus = "Skipped by Engineer."
} else {
    if ($DemoMode) {
        Write-Host "      [75%] SFC /Scannow..." -NoNewline; Start-Sleep -Seconds 2; Write-Host " OK" -ForegroundColor Green
        $ReportData.SfcStatus = "Issues Found & Repaired."
    } else {
        Write-Host "   > Running SFC /Scannow..." -ForegroundColor Yellow
        $SfcOut = sfc /scannow 2>&1 | Out-String
        if ($SfcOut -match "found no integrity violations") { $SfcStatus = "Healthy (No integrity violations found)." }
        elseif ($SfcOut -match "found corrupt files and successfully repaired") { $SfcStatus = "<span style='color:green'><strong>FIXED:</strong> Corrupt files successfully repaired.</span>" } 
        elseif ($SfcOut -match "found corrupt files but was unable to fix") { $SfcStatus = "<span style='color:red'><strong>CRITICAL:</strong> Corrupt files found, Windows could NOT fix them.</span>" } 
        else { $SfcStatus = "Scan Completed." }
        
        try {
            $CBSLog = "$env:windir\Logs\CBS\CBS.log"
            if (Test-Path $CBSLog -and ($SfcStatus -match "FIXED|CRITICAL")) {
                $SfcLogContent = Get-Content $CBSLog | Select-String '\[SR\]' | Select-Object -Last 10 | Out-String
            }
        } catch { }
        $ReportData.SfcStatus = $SfcStatus
    }
    Write-Host "   > SFC Scan Completed." -ForegroundColor Green
}

# --- 11. WINGET SOFTWARE UPDATES ---
if (Test-UserSkip -StepName 'Software Updates (Winget)' -Seconds 5) {
    $ReportData.WingetStatus = 'Skipped by Engineer.'
} else {
    Write-Host '   > Detecting available updates...' -ForegroundColor Yellow
    if ($DemoMode) {
        Start-Sleep -Seconds 2
        $ReportData.WingetStatus = 'Success: Packages updated.'
    } else {
        try {
            $WingetExe = Get-Command 'winget' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source
            if (-not $WingetExe) { 
                $LocalWinget = "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe"
                if (Test-Path $LocalWinget) { $WingetExe = $LocalWinget } 
            }
            
            if ($WingetExe) {
                $UpgradeListRaw = cmd /c 'winget upgrade --accept-source-agreements --disable-interactivity' 2>&1 | Out-String
                $PackagesToUpdate = @()
                
                if ($UpgradeListRaw -match 'No installed package found matching input criteria') {
                    $ReportData.WingetStatus = 'Status OK: No software updates found.'
                } else {
                    # Pure ASCII regex, single quotes to prevent text editor corruption
                    $Lines = $UpgradeListRaw -split '\r?\n' | Where-Object { 
                        $_ -notmatch '^Name|^---|^\s*$' -and 
                        $_.Length -gt 5 -and 
                        $_ -notmatch '[^\x20-\x7E]{3,}' -and 
                        $_ -match '[a-zA-Z]' 
                    }
                    
                    foreach ($Line in $Lines) {
                        $CleanName = $Line -replace '[^\x20-\x7E]', ''
                        $CleanName = $CleanName.Trim()
                        
                        if ($CleanName.Length -gt 40) { $CleanName = $CleanName.Substring(0, 40) + '...' }
                        if ($CleanName.Length -gt 2) { $PackagesToUpdate += "<li>$CleanName</li>" }
                    }
                    
                    if ($PackagesToUpdate.Count -gt 0) {
                        # Removed subexpressions entirely from HTML generation
                        $JoinedPkgs = $PackagesToUpdate -join ''
                        $PackageHTML = "<ul style='margin-top:5px; margin-bottom:5px; font-size:0.9em;'>$JoinedPkgs</ul>"
                        
                        Write-Host '   > Installing updates...' -ForegroundColor Yellow
                        $ProcArgs = 'upgrade --all --accept-source-agreements --accept-package-agreements --include-unknown --disable-interactivity'
                        $WingetProcess = Start-Process -FilePath $WingetExe -ArgumentList $ProcArgs -Wait -PassThru -NoNewWindow
                        
                        $ExitCode = $WingetProcess.ExitCode
                        if ($ExitCode -eq 0) { 
                            $ReportData.WingetStatus = "Success: The following packages were updated:<br>$PackageHTML" 
                        } else { 
                            $ReportData.WingetStatus = "Warning: Winget returned code ${ExitCode}.<br>Targeted packages:<br>$PackageHTML" 
                        }
                    } else {
                        $ReportData.WingetStatus = 'Status OK: No software updates required.'
                    }
                }
            } else { $ReportData.WingetStatus = 'Failed: Winget command not found.' }
        } catch { $ReportData.WingetStatus = 'Error: Winget execution failed.' }
    }
    Write-Host '   > Done.' -ForegroundColor Green
}

# --- 12. WINDOWS UPDATES ---
if (Test-UserSkip -StepName 'Windows Update Check') {
    $ReportData.Updates = 'Skipped by Engineer.'
} else {
    Write-Host '   > Checking for updates...' -ForegroundColor Yellow
    if ($DemoMode) {
        $ReportData.Updates = 'Action Required: 2 updates pending.'
    } else {
        try {
            $UpdateSession = New-Object -ComObject Microsoft.Update.Session
            $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
            $SearchResult = $UpdateSearcher.Search("IsInstalled=0")
            $Count = $SearchResult.Updates.Count
            if ($Count -gt 0) { 
                $PendingUpdates = @()
                for ($i = 0; $i -lt $Count; $i++) { $UpdateItem = $SearchResult.Updates.Item($i); $PendingUpdates += "<li>$($UpdateItem.Title)</li>" }
                $UpdateListHTML = "<ul style='margin-top:5px; margin-bottom:5px; font-size:0.9em; color:#d9534f;'>$($PendingUpdates -join '')</ul>"
                $ReportData.Updates = "Action Required: $Count updates pending.<br>$UpdateListHTML" 
            } else { $ReportData.Updates = 'Status OK: System is up to date.' }
        } catch { $ReportData.Updates = 'Manual Check Required.' }
    }
    Write-Host '   > Done.' -ForegroundColor Green
}

# --- 13. DISK OPTIMIZATION ---
Write-Host "`n   [NEXT STEP] Disk Defragmentation" -ForegroundColor Cyan
if ($DemoMode) { $IsSSD = $false } else {
    try { $SystemDisk = Get-Partition -DriveLetter C -ErrorAction SilentlyContinue | Get-Disk | Get-PhysicalDisk; $IsSSD = ($SystemDisk.MediaType -eq 'SSD') } catch { $IsSSD = $false }
}

if ($IsSSD) {
    Write-Host '   > Drive Type Detected: SSD/NVMe (Solid State Drive)' -ForegroundColor Green
    Write-Host '   > SKIPPING DEFRAGMENTATION (Not required for this drive type).' -ForegroundColor Green
    $ReportData.Defrag = 'Scan not required for SSD/NVMe'
} else {
    Write-Host '   > Drive Type Detected: HDD (Hard Disk Drive)' -ForegroundColor Yellow
    Write-Host '   [Y] Yes  |  [N] No (Default)  |  [Enter] Default' -ForegroundColor Gray
    while ([Console]::KeyAvailable) { $null = [Console]::ReadKey($true) }

    $OptimizeChoice = $null
    $EndTime = (Get-Date).AddSeconds(60)

    while ((Get-Date) -lt $EndTime) {
        $Remaining = [math]::Ceiling(($EndTime - (Get-Date)).TotalSeconds)
        Write-Host "`r   > Waiting for input (Default: No) in $Remaining seconds...   " -NoNewline -ForegroundColor Yellow
        if ([Console]::KeyAvailable) {
            $Key = [Console]::ReadKey($true).Key.ToString().ToUpper()
            if ($Key -in "Y","N","ENTER") { $OptimizeChoice = $Key; break }
        }
        Start-Sleep -Milliseconds 100
    }

    if ([string]::IsNullOrWhiteSpace($OptimizeChoice)) { $OptimizeChoice = 'TIMEOUT' }
    Write-Host "`r                                                              " -NoNewline

    if ($OptimizeChoice -eq 'Y') {
        Write-Host '   > Option Selected: YES                                     ' -ForegroundColor Green
        if ($DemoMode) { Start-Sleep -Seconds 1; $ReportData.Defrag = 'Optimization Complete.' } 
        else { Write-Host '   > Optimizing Drive C:...' -ForegroundColor Yellow; Optimize-Volume -DriveLetter C -Defrag -Verbose | Out-Null; $ReportData.Defrag = 'Optimization Complete: Drive C: optimized.' }
        Write-Host '   > Done.' -ForegroundColor Green
    } elseif ($OptimizeChoice -in 'ENTER','TIMEOUT') {
        Write-Host '   > Option Selected: DEFAULT (No Input)                      ' -ForegroundColor Yellow
        $ReportData.Defrag = 'Skipped: No Input'
    } else {
        Write-Host '   > SKIPPED: Disk Optimization.                              ' -ForegroundColor Yellow
        $ReportData.Defrag = 'Skipped by Engineer.'
    }
}

# --- 14. NEW REPORT GENERATION ARRAY BUILDER ---
Write-Host "`n[REPORT] Generating Diagnostic Style Report..." -ForegroundColor Yellow

# CHECK INTERNET FOR REPORT
if ($DemoMode -or (Test-Connection -ComputerName 8.8.8.8 -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
    $ReportData.Internet = "Active (Verified)"
} else { $ReportData.Internet = "Disconnected" }

$ReportItems = @()
function Add-Result($Category, $Component, $Details, $Status) {
    $script:ReportItems += [PSCustomObject]@{ Category = $Category; Component = $Component; Details = $Details; Status = $Status }
}

# 1. Device Information
Add-Result "Device Information" "Hostname & OS" "<b>Name:</b> $($ComputerInfo.Name)<br><b>OS:</b> $($OSInfo.Caption) $($OSInfo.OSArchitecture) ($($OSInfo.Version))<br><b>Install Date:</b> $FormattedInstallDate" "OK"
Add-Result "Device Information" "Compute & Graphics" "<b>CPU:</b> $($CPUInfo.Name)<br><b>Cores/Threads:</b> $($CPUInfo.NumberOfCores)C / $($CPUInfo.NumberOfLogicalProcessors)T<br><b>RAM Size:</b> $([math]::Round($ComputerInfo.TotalPhysicalMemory / 1GB, 0)) GB<br><b>GPU:</b> $GPUString" "OK"
Add-Result "Device Information" "Network Adapters" "<ul style='margin:0; padding-left:20px;'>$NetworkAdaptersHTML</ul>" "Info"
Add-Result "Device Information" "Physical Disks" "<ul style='margin:0; padding-left:20px;'>$DriveListHTML</ul>" "Info"

# 2. System Status
$IntStatus = if ($ReportData.Internet -match "Active") {"Connected"} else {"FAIL"}
Add-Result "System Status" "Internet Connection" "<b>Status:</b> $($ReportData.Internet)" $IntStatus

$BatStatus = if ($ReportData.Battery -match "OK|No Battery") {"Pass"} else {"Warning"}
Add-Result "System Status" "Battery Health" "<b>Telemetry:</b> $($ReportData.Battery)" $BatStatus

$RPStatus = if ($ReportData.RestorePoint -match "Success") {"Pass"} elseif ($ReportData.RestorePoint -match "Skipped") {"Skipped"} else {"Warning"}
Add-Result "System Status" "System Restore Point" "<b>Creation Status:</b> $($ReportData.RestorePoint)" $RPStatus

# 3. Storage Analysis (Embedded Graphic)
$StorageDetails = @"
<div style='display:flex; align-items:center; gap:20px; margin-bottom: 10px;'>
    <div style='width:60px; height:60px; border-radius:50%; background:conic-gradient(#d9534f 0% $($StoragePercent)%, #5cb85c $($StoragePercent)% 100%); border:2px solid #ddd; flex-shrink: 0;'></div>
    <div style='font-size: 0.9em;'>
        <b>Total Capacity:</b> $StorageTotalGB GB<br>
        <span style='color:#5cb85c;'>&#9632;</span> <b>Free:</b> $StorageFreeGB GB<br>
        <span style='color:#d9534f;'>&#9632;</span> <b>Used:</b> $StorageUsedGB GB ($($StoragePercent)%)
    </div>
</div>
<b>Largest Root Items:</b>
<table style='width:100%; font-size:0.85em; margin-top:5px; border:none; box-shadow:none;'><tbody style='border:none;'>$StorageTableRows</tbody></table>
"@
$StorStatus = if ($StoragePercent -lt 85) {"Pass"} else {"Warning"}
Add-Result "Storage Analysis" "Drive C: Capacity" $StorageDetails $StorStatus

# 4. Maintenance & Updates
$DismStatColor = if ($ReportData.DismStatus -match "Corruption detected|Issues") {"FAIL"} else {"Pass"}
Add-Result "Maintenance & Updates" "DISM Health Check" $($ReportData.DismStatus) $DismStatColor

$SfcStatColor = if ($ReportData.SfcStatus -match "CRITICAL") {"FAIL"} elseif ($ReportData.SfcStatus -match "FIXED") {"Warning"} else {"Pass"}
$SfcDetails = $($ReportData.SfcStatus)
if ($SfcLogContent) { $SfcDetails += "<br><br><b>Logs:</b><pre style='background:#f4f4f4; padding:5px; font-size:0.8em; overflow-x:auto;'>$SfcLogContent</pre>" }
Add-Result "Maintenance & Updates" "SFC System Integrity" $SfcDetails $SfcStatColor

$WinGetColor = if ($ReportData.WingetStatus -match "Error|Failed") {"FAIL"} elseif ($ReportData.WingetStatus -match "Skipped") {"Skipped"} else {"Pass"}
Add-Result "Maintenance & Updates" "Software Upgrades (Winget)" $($ReportData.WingetStatus) $WinGetColor

$WinUpColor = if ($ReportData.Updates -match "Action Required") {"Warning"} elseif ($ReportData.Updates -match "Skipped") {"Skipped"} else {"Pass"}
Add-Result "Maintenance & Updates" "Windows Updates" $($ReportData.Updates) $WinUpColor

Add-Result "Maintenance & Updates" "Storage Cleanup" $($ReportData.DiskCleanup) "Pass"

$DefragColor = if ($ReportData.Defrag -match "Skipped") {"Skipped"} else {"Pass"}
Add-Result "Maintenance & Updates" "Disk Optimization" $($ReportData.Defrag) $DefragColor

# 5. Installed Software
Add-Result "Installed Software" "Application Inventory" "<div style='column-count:2; font-size:0.85em;'><ul style='margin:0; padding-left:15px;'>$AppListHTML</ul></div>" "Info"

# --- HTML/PDF ASSEMBLY (MATCHING HARDWARECHECK.PS1) ---
$TableRowsHTML = ""
$CurrentCategory = ""

foreach ($Row in $ReportItems) {
    if ($Row.Category -ne $CurrentCategory) {
        $TableRowsHTML += "<tr class='category-row'><td colspan='3'>$($Row.Category)</td></tr>"
        $CurrentCategory = $Row.Category
    }

    $StatusColor = switch -Regex ($Row.Status) {
        "Pass|OK|Connected|Activated|Encrypted|Enabled" { "#28a745" }
        "Warning|Low Memory|Low Space|Degraded" { "#fd7e14" }
        "FAIL|Error|Not Activated|Disabled" { "#dc3545" }
        "Skipped" { "#6c757d" }
        Default { "#17a2b8" } # Info color
    }
    $TableRowsHTML += "<tr><td style='width: 25%;'><strong>$($Row.Component)</strong></td><td style='width: 55%;'>$($Row.Details)</td><td style='width: 20%; text-align: center;'><strong style='color:$StatusColor'>$($Row.Status)</strong></td></tr>"
}

$HtmlFile = "$BaseDir\$ReportFilename.html"
$PdfFile  = "$BaseDir\$ReportFilename.pdf"
$FinalReportPath = ""
$ModeLabel = if ($DemoMode) { "(DEMO MODE)" } else { "" }

$HtmlContent = @"
<!DOCTYPE html>
<html>
<head>
<style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; padding: 30px; line-height: 1.4; }
    .header { text-align: center; margin-bottom: 30px; border-bottom: 2px solid #0056b3; padding-bottom: 10px; }
    h1 { color: #0056b3; margin-bottom: 5px; font-size: 24px; }
    .meta { font-size: 14px; color: #555; display: flex; justify-content: space-between; margin-top: 15px; }
    table { width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 13px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
    th { text-align: left; background: #0056b3; color: white; padding: 10px; }
    td { padding: 10px; border-bottom: 1px solid #e0e0e0; vertical-align: top; }
    .category-row td { background: #f0f4f8; font-weight: bold; font-size: 14px; color: #0056b3; text-transform: uppercase; border-top: 2px solid #d0dce8; }
    b { color: #222; }
    .footer { text-align: center; font-size: 11px; color: #888; margin-top: 40px; border-top: 1px solid #ddd; padding-top: 10px; }
</style>
</head>
<body>
<div class="header">
    <img src="$LogoUrl" alt="Apollo Technology" style="max-height:80px;" onerror="this.style.display='none'">
    <h1>Health Check Report $ModeLabel</h1>
    <div class="meta">
        <div><strong>Ticket:</strong> $TicketNumber<br><strong>Customer:</strong> $CustomerName</div>
        <div style="text-align: right;"><strong>Date:</strong> $CurrentDate<br><strong>Engineer:</strong> $EngineerName</div>
    </div>
</div>

<table>
    <thead><tr><th>Component / Check</th><th>Detailed Specifications & Diagnostics</th><th style='text-align: center;'>Health / Status</th></tr></thead>
    <tbody>$TableRowsHTML</tbody>
</table>

<div class="footer">
    &copy; $CurrentYear by Apollo Technology LTD. Created by Lewis Wiltshire (Apollo Technology).
</div>
</body>
</html>
"@

$HtmlContent | Out-File -FilePath $HtmlFile -Encoding UTF8
Start-Sleep -Seconds 1

# Convert to PDF using Edge
$EdgeLoc1 = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
$EdgeLoc2 = "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
$EdgeExe = if (Test-Path $EdgeLoc1) { $EdgeLoc1 } elseif (Test-Path $EdgeLoc2) { $EdgeLoc2 } else { $null }

if ($EdgeExe) {
    Write-Host "   > Converting to PDF..." -ForegroundColor Cyan
    $EdgeUserData = "$BaseDir\EdgeTemp"
    if (-not (Test-Path $EdgeUserData)) { New-Item -Path $EdgeUserData -ItemType Directory -Force | Out-Null }
    try {
        Start-Process -FilePath $EdgeExe -ArgumentList "--headless", "--disable-gpu", "--print-to-pdf=`"$PdfFile`"", "--no-pdf-header-footer", "--user-data-dir=`"$EdgeUserData`"", "`"$HtmlFile`"" -PassThru -Wait
        Start-Sleep -Seconds 2 
        if (Test-Path $PdfFile) {
            Write-Host "   > Success! Report Generated at $PdfFile" -ForegroundColor Green
            $FinalReportPath = $PdfFile
            Remove-Item $HtmlFile -ErrorAction SilentlyContinue
            Remove-Item $EdgeUserData -Recurse -Force -ErrorAction SilentlyContinue
        } else { throw "PDF not found" }
    } catch {
        Write-Warning "PDF Conversion failed. Opening HTML instead."
        $FinalReportPath = $HtmlFile
    }
} else {
    Write-Warning "Edge not found. Saving HTML report."
    $FinalReportPath = $HtmlFile
}

# --- 17. EMAIL REPORT ---
if ($EmailEnabled -and $PdfFile -and (Test-Path $PdfFile)) {
    Write-Host "`nSending Email to $ToAddress..." -ForegroundColor Yellow
    try {
        Send-MailMessage -From $FromAddress -To $ToAddress -Subject "Health Check Report: $env:COMPUTERNAME ($TicketNumber)" -Body "Report attached for $CustomerName." -SmtpServer $SmtpServer -Port $SmtpPort -UseSsl $UseSSL -Credential $EmailCreds -Attachments $PdfFile -ErrorAction Stop
        Write-Host "   > Email Sent Successfully!" -ForegroundColor Green
    } catch {
        Write-Error "   > Failed to send email. Error: $_"
    }
}

try { [SleepUtils]::SetThreadExecutionState(0x80000000) | Out-Null } catch { }

Write-Host "`n------------------------------------------------------------"
Write-Host "          PROCESS COMPLETED" -ForegroundColor Green
if (Test-Path $FinalReportPath) { Start-Process $FinalReportPath; Start-Sleep -Seconds 1 }
Write-Host "------------------------------------------------------------"
exit