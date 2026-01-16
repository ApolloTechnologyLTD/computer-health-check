<#
.SYNOPSIS
    Apollo Technology Manual Health Check Report Generator
.DESCRIPTION
    Generates a PDF Health Check report based on MANUAL user input.
    - No automatic scanning.
    - User types in all details.
    - Auto-calculates Storage Pie Chart based on user numbers.
    - Converts to PDF using Edge.
.NOTES
    Modified for Manual Entry
#>

# --- CONFIGURATION ---
$LogoUrl = "https://raw.githubusercontent.com/ApolloTechnologyLTD/computer-health-check/main/Apollo%20Cropped.png"
$EmailEnabled = $false       # Set to $true to enable email prompt at end

# Email Settings (Only used if enabled above)
$SmtpServer   = "smtp.office365.com"
$SmtpPort     = 587
$FromAddress  = "reports@yourdomain.com"
$ToAddress    = "support@yourdomain.com"
$UseSSL       = $true

# --- 1. ELEVATE TO ADMIN (Required for file saving/PDF conversion) ---
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "Requesting Administrator privileges..." -ForegroundColor Yellow
    try {
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        exit
    } catch {
        Write-Error "Failed to elevate. Please right-click and 'Run as Administrator'."
        exit
    }
}

# --- 2. SETUP PATHS ---
Clear-Host
$ApolloASCII = @"
    ___    ____  ____  __    __    ____     ____________________  ___   ______  __    ____  ________  __
   /   |  / __ \/ __ \/ /   / /   / __ \   /_  __/ ____/ ____/ / / / | / / __ \/ /   / __ \/ ____/\ \/ /
  / /| | / /_/ / / / / /   / /   / / / /    / / / __/ / /   / /_/ /  |/ / / / / /   / / / / / __   \  / 
 / ___ |/ ____/ /_/ / /___/ /_/ /    / / / /___/ /___/ __  / /|  / /_/ / /___/ /_/ / /_/ /   / /  
/_/  |_/_/    \____/_____/_____/\____/    /_/ /_____/\____/_/ /_/_/ |_/\____/_____/\____/\____/   /_/   
                                                                                                        
"@
Write-Host $ApolloASCII -ForegroundColor Cyan
Write-Host "      MANUAL REPORT GENERATOR MODE" -ForegroundColor Yellow
Write-Host "      ----------------------------" -ForegroundColor Gray

$BaseDir = "C:\temp\Apollo_Reports"
try {
    if (-not (Test-Path $BaseDir)) { New-Item -Path $BaseDir -ItemType Directory -Force -ErrorAction Stop | Out-Null }
} catch {
    $BaseDir = "$env:USERPROFILE\Desktop\Apollo_Reports"
    if (-not (Test-Path $BaseDir)) { New-Item -Path $BaseDir -ItemType Directory -Force | Out-Null }
}

# --- 3. INPUT: TICKET DETAILS ---
do {
    Write-Host "`n--- 1. TICKET DETAILS ---" -ForegroundColor Cyan
    $EngineerName = Read-Host "Engineer Name"
    $CustomerName = Read-Host "Customer Name"
    $TicketNumber = Read-Host "Ticket Number"

    if ($TicketNumber -notmatch "^#") { $TicketNumber = "#$TicketNumber" }

    Write-Host "`n   Engineer: $EngineerName" -ForegroundColor Gray
    Write-Host "   Customer: $CustomerName" -ForegroundColor Gray
    Write-Host "   Ticket:   $TicketNumber" -ForegroundColor Gray
    
    $Confirmation = Read-Host "`n   Correct? (Y/N) [Default: Y]"
    if ($Confirmation -eq "") { $Confirmation = "Y" }
} while ($Confirmation -match "N")

$CurrentDate = Get-Date -Format "yyyy-MM-dd HH:mm"
$CurrentYear = (Get-Date).Year
$ReportFilename = "Manual_Report_$(Get-Date -Format 'yyyy-MM-dd_HH-mm')"

# --- 4. INPUT: HARDWARE SPECS ---
Write-Host "`n--- 2. HARDWARE SPECIFICATIONS ---" -ForegroundColor Cyan
Write-Host "(Press Enter to leave blank or accept defaults)" -ForegroundColor DarkGray

$In_DevName = Read-Host "Device Name (Hostname)"
$In_OS      = Read-Host "Operating System (e.g. Windows 10 Pro)"
$In_CPU     = Read-Host "Processor (e.g. Intel i5-10500)"
$In_RAM     = Read-Host "RAM Amount (e.g. 16GB)"
$In_GPU     = Read-Host "Graphics Card (e.g. Intel UHD / Nvidia RTX 3060)"
$In_Age     = Read-Host "Approx Age/Install Date (e.g. 2021)"
$In_Drive   = Read-Host "Drive Details (e.g. 500GB NVMe SSD - Healthy)"

# --- 5. INPUT: STORAGE DATA (For Chart) ---
Write-Host "`n--- 3. STORAGE CALCULATION ---" -ForegroundColor Cyan
$StorageTotalGB = Read-Host "Total Drive Capacity (GB) [Just the number, e.g. 500]"
$StorageFreeGB  = Read-Host "Free Space Remaining (GB) [Just the number, e.g. 120]"

# Calculate Logic
$StorageUsedGB = 0
$StoragePercent = 0
try {
    if ($StorageTotalGB -and $StorageFreeGB) {
        $StorageTotalGB = [double]$StorageTotalGB
        $StorageFreeGB = [double]$StorageFreeGB
        $StorageUsedGB = $StorageTotalGB - $StorageFreeGB
        $StoragePercent = [math]::Round(($StorageUsedGB / $StorageTotalGB) * 100)
    }
} catch {
    $StoragePercent = 0
    Write-Warning "Could not calculate chart data. Check numbers."
}

# --- 6. INPUT: HEALTH CHECKS ---
Write-Host "`n--- 4. HEALTH CHECKS (Status Updates) ---" -ForegroundColor Cyan
Write-Host "Suggested Inputs: 'Pass', 'Fail', 'Updated', 'Cleaned', 'Skipped'" -ForegroundColor DarkGray

$Status_Internet = Read-Host "Internet Connection Status [Default: Active]"
if ($Status_Internet -eq "") { $Status_Internet = "Active (Verified)" }

$Status_Battery  = Read-Host "Battery Health Status    [Default: OK / Desktop]"
if ($Status_Battery -eq "") { $Status_Battery = "Battery Status OK (Healthy)" }

$Status_Restore  = Read-Host "Restore Point Created?   [Default: Success]"
if ($Status_Restore -eq "") { $Status_Restore = "Success: Restore point created." }

$Status_Updates  = Read-Host "Windows Updates Status   [Default: Up to Date]"
if ($Status_Updates -eq "") { $Status_Updates = "System is fully up to date." }

$Status_Clean    = Read-Host "Disk Cleanup Status      [Default: Cleaned]"
if ($Status_Clean -eq "") { $Status_Clean = "Maintenance Complete: Temp files removed." }

$Status_Defrag   = Read-Host "Defrag/Optimize Status   [Default: Optimized/Skipped SSD]"
if ($Status_Defrag -eq "") { $Status_Defrag = "Optimization Complete." }

$Status_SFC      = Read-Host "SFC/DISM Scan Results    [Default: No Integrity Violations]"
if ($Status_SFC -eq "") { $Status_SFC = "Healthy (No integrity violations found)." }

# --- 7. INPUT: SOFTWARE / NOTES ---
Write-Host "`n--- 5. SOFTWARE & NOTES ---" -ForegroundColor Cyan
$SoftwareList = Read-Host "Key Software Installed (Comma separated or short note)"
$EngineerNotes = Read-Host "Additional Engineer Notes (Optional)"

if ($SoftwareList -eq "") { $SoftwareList = "Standard Business Apps, Office 365, Chrome." }

# Split software into list items for HTML
$SoftwareHTML = ""
$SoftwareList -split "," | ForEach-Object { $SoftwareHTML += "<li>$($_)</li>" }

# --- 8. GENERATE HTML ---
Write-Host "`n[REPORT] Generating Report..." -ForegroundColor Yellow

# Build Storage Chart HTML
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
</div>
"@

$NotesHTML = ""
if ($EngineerNotes) {
    $NotesHTML = @"
    <h2>Engineer Notes</h2>
    <div class="section">
        <p>$EngineerNotes</p>
    </div>
"@
}

$HtmlFile = "$BaseDir\$ReportFilename.html"
$PdfFile  = "$BaseDir\$ReportFilename.pdf"
$PublicDesktopPdf = "$env:PUBLIC\Desktop\$ReportFilename.pdf"

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
</style>
</head>
<body>
<div class="header">
    <img src="$LogoUrl" alt="Apollo Technology" onerror="this.style.display='none'">
    <h1>Health Check Report</h1>
    <p>This report has been generated by <strong>$EngineerName</strong> for ticket (<strong>$TicketNumber</strong>) for <strong>$CustomerName</strong></p>
    <div class="meta">
        <strong>Report Date:</strong> $CurrentDate
    </div>
</div>

<h2>Device Information</h2>
<div class="section">
    <div class="item"><span class="label">Device Name:</span> $In_DevName</div>
    <div class="item"><span class="label">OS Version:</span> $In_OS</div>
    <div class="item"><span class="label">CPU:</span> $In_CPU</div>
    <div class="item"><span class="label">RAM:</span> $In_RAM</div>
    <div class="item"><span class="label">GPU:</span> $In_GPU</div>
    <div class="item"><span class="label">Approx Age:</span> $In_Age</div>
    <div class="item"><span class="label">Storage Drive:</span> $In_Drive</div>
</div>

<h2>System Status</h2>
<div class="section">
    <div class="item"><span class="label">Internet Connection:</span> $Status_Internet</div>
    <div class="item"><span class="label">Battery Status:</span> $Status_Battery</div>
    <div class="item"><span class="label">Restore Point:</span> $Status_Restore</div>
</div>

$StorageReportHTML

<h2>Maintenance & Updates</h2>
<div class="section">
    <div class="item"><span class="label">System Integrity (SFC/DISM):</span> $Status_SFC</div>
    <div class="item"><span class="label">Windows Updates:</span> $Status_Updates</div>
    <div class="item"><span class="label">Storage Cleanup:</span> $Status_Clean</div>
    <div class="item"><span class="label">Disk Optimization:</span> $Status_Defrag</div>
</div>

<h2>Installed Software Inventory</h2>
<div class="section"><ul>$SoftwareHTML</ul></div>

$NotesHTML

<p style="text-align:center; font-size:0.8em; color:#888; margin-top:50px;">&copy; $CurrentYear by Apollo Technology. All rights reserved.</p>
</body>
</html>
"@

# --- 9. SAVE & CONVERT ---
$HtmlContent | Out-File -FilePath $HtmlFile -Encoding UTF8

# Convert to PDF using Edge
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
        Start-Process $HtmlFile
    }
} else {
    Write-Warning "Edge not found. Opening HTML report."
    Start-Process $HtmlFile
}

# --- 10. OPTIONAL EMAIL ---
if ($EmailEnabled) {
    $SendEmail = Read-Host "`nSend Email to $ToAddress? (Y/N)"
    if ($SendEmail -eq "Y") {
        $EmailPass = Read-Host "Enter Password for $FromAddress" -AsSecureString
        $EmailCreds = New-Object System.Management.Automation.PSCredential ($FromAddress, $EmailPass)
        try {
            Send-MailMessage -From $FromAddress -To $ToAddress -Subject "Health Check Report: $CustomerName" -Body "Manual Report attached." -SmtpServer $SmtpServer -Port $SmtpPort -UseSsl $UseSSL -Credential $EmailCreds -Attachments $PublicDesktopPdf -ErrorAction Stop
            Write-Host "   > Email Sent!" -ForegroundColor Green
        } catch {
            Write-Error "   > Email Failed: $_"
        }
    }
}

Write-Host "`n------------------------------------------------------------"
Write-Host "          REPORT COMPLETED" -ForegroundColor Green
if (Test-Path $PublicDesktopPdf) { Start-Process $PublicDesktopPdf }
Write-Host "------------------------------------------------------------"
Pause