# =====================================================================
# VMware ESXi 6.7 Multi-VM Snapshot Automation via SSH + Excel Report
# Description : Connects via SSH to ESXi 6.7, creates VM snapshots
#               using vim-cmd, and exports results to an Excel report.
# =====================================================================

# =====================================================================
# BACKEND CONFIGURATION
# =====================================================================

$ESXiHost     = "192.168.5.2"
$ESXiUsername = "root"

$VMList = @(
    "SRV_01",
    "SRV_02",
    "SRV_03",
    "SRV_04",
    "SRV_05",
    "SRV_06",
    "SRV_07",
    "Tally_Data",
    "Controller_SRV"
)

# =====================================================================
# START
# =====================================================================

Clear-Host

Write-Host ""
Write-Host "==============================================================" -ForegroundColor Cyan
Write-Host "     VMware ESXi 6.7 SSH SNAPSHOT AUTOMATION" -ForegroundColor White
Write-Host "                  + EXCEL REPORT" -ForegroundColor White
Write-Host "==============================================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "ESXi Host : $ESXiHost"
Write-Host "Username  : $ESXiUsername"
Write-Host "VM Count  : $($VMList.Count)"
Write-Host ""

Write-Host "VMs configured for snapshot:" -ForegroundColor Cyan
foreach ($VMName in $VMList) {
    Write-Host " - $VMName"
}
Write-Host ""

# =====================================================================
# DESKTOP LOCATION
# =====================================================================

$Desktop = [Environment]::GetFolderPath("Desktop")
if ([string]::IsNullOrWhiteSpace($Desktop)) {
    Write-Host "[ERROR] Could not determine Desktop location." -ForegroundColor Red
    Read-Host "Press Enter to close"
    return
}

# =====================================================================
# EXCEL MODULE
# =====================================================================

Write-Host "==============================================================" -ForegroundColor Yellow
Write-Host " Preparing Excel Reporting" -ForegroundColor Yellow
Write-Host "==============================================================" -ForegroundColor Yellow
Write-Host ""

try {
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "[INFO] Installing ImportExcel module..." -ForegroundColor Yellow
        Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }
    Import-Module ImportExcel -Force -ErrorAction Stop
    Write-Host "[ OK ] Excel reporting module ready." -ForegroundColor Green
    Write-Host ""
}
catch {
    Write-Host "[ERROR] Failed to load ImportExcel: $($_.Exception.Message)" -ForegroundColor Red
    Read-Host "Press Enter to close"
    return
}

# =====================================================================
# SSH HELPER SETUP (Using SSH.NET via PowerShell or native SSH / Plink)
# =====================================================================

Write-Host "==============================================================" -ForegroundColor Yellow
Write-Host " ESXi Authentication (SSH)" -ForegroundColor Yellow
Write-Host "==============================================================" -ForegroundColor Yellow
Write-Host ""

# Check for SSH.NET module (Posh-SSH) or fallback to plink / OpenSSH
if (-not (Get-Module -ListAvailable -Name Posh-SSH)) {
    Write-Host "[INFO] Posh-SSH module not found. Installing Posh-SSH for clean SSH execution..." -ForegroundColor Yellow
    try {
        Install-Module -Name Posh-SSH -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    } catch {
        Write-Host "[WARN] Could not install Posh-SSH automatically." -ForegroundColor Yellow
    }
}

Import-Module Posh-SSH -ErrorAction SilentlyContinue

$SecurePassword = Read-Host "Enter ESXi root password" -AsSecureString
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
$PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
[System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)

$Credential = [System.Management.Automation.PSCredential]::new($ESXiUsername, $SecurePassword)

$SSHSession = $null
try {
    Write-Host "Connecting to $ESXiHost via SSH..." -ForegroundColor Cyan
    $SSHSession = New-SSHSession -ComputerName $ESXiHost -Credential $Credential -AcceptKey -ErrorAction Stop
    Write-Host "Connected successfully via SSH." -ForegroundColor Green
    Write-Host ""
}
catch {
    Write-Host "[ERROR] Could not connect to ESXi via SSH: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Ensure SSH service is running on ESXi 6.7 and port 22 is reachable." -ForegroundColor Yellow
    Read-Host "Press Enter to close"
    return
}

# Function to execute remote SSH command cleanly
function Invoke-ESXiCommand ($Command) {
    $result = Invoke-SSHCommand -SSHSession $SSHSession -Command $Command
    return $result.Output
}

# =====================================================================
# PREPARE EXCEL FILE & ARRAYS
# =====================================================================

$DateStamp = Get-Date -Format "MMM-dd-yyyy_HH-mm-ss"
$ExcelFile = Join-Path $Desktop "ESXi_${ESXiHost}_Snapshot_Report_$DateStamp.xlsx"

$ESXiReportRows     = @()
$VMReportRows       = @()
$SnapshotReportRows = @()

# =====================================================================
# ESXi HOST INFO
# =====================================================================

Write-Host "==============================================================" -ForegroundColor Cyan
Write-Host " Collecting ESXi Information" -ForegroundColor Cyan
Write-Host "==============================================================" -ForegroundColor Cyan
Write-Host ""

$ESXiVersionRaw = (Invoke-ESXiCommand "vmware -v") -join " "
$ESXiHostName   = ((Invoke-ESXiCommand "hostname") -join "").Trim()

Write-Host "Host Name : $ESXiHostName"
Write-Host "Version   : $ESXiVersionRaw"
Write-Host ""

# =====================================================================
# RETRIEVE VM INVENTORY (VMID Mapping)
# =====================================================================

Write-Host "Fetching VM catalog from ESXi..." -ForegroundColor Cyan
$VMCatalogRaw = Invoke-ESXiCommand "vim-cmd vmsvc/getallvms"

# Parse vim-cmd table output: Vmid, Name, File, Guest OS, Version, Annotation
$VMInventory = @{}
foreach ($line in $VMCatalogRaw) {
    if ($line -match '^\s*(\d+)\s+([^\s]+)\s+\[([^\]]+)\]\s+([^ ]+)\s+([^\s]+)') {
        $vmid   = $matches[1]
        $vmname = $matches[2]
        $ds     = $matches[3]
        $guest  = $matches[5]
        $VMInventory[$vmname] = @{
            Vmid      = $vmid
            Datastore = $ds
            GuestOS   = $guest
        }
    }
}

# =====================================================================
# PROCESS EACH VM
# =====================================================================

$SuccessfulSnapshots = 0
$FailedSnapshots     = 0
$VMsFound            = 0
$VMsNotFound         = 0

$CurrentVM = 0
$TotalVMs  = $VMList.Count

foreach ($VMName in $VMList) {
    $CurrentVM++
    $Percent = [math]::Round(($CurrentVM / $TotalVMs) * 100)
    Write-Progress -Activity "VMware Snapshot Process (SSH)" -Status "Processing $VMName [$CurrentVM of $TotalVMs]" -PercentComplete $Percent

    $TimeStamp    = Get-Date -Format "yyyy-MM-dd-HH-mm-ss"
    $SnapshotName = "AUTO-$VMName-$TimeStamp"
    $SnapshotDesc = "Automated snapshot taken on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

    Write-Host ""
    Write-Host "==============================================================" -ForegroundColor Cyan
    Write-Host " Processing VM [$CurrentVM/$TotalVMs]: $VMName" -ForegroundColor Cyan
    Write-Host "==============================================================" -ForegroundColor Cyan

    if (-not $VMInventory.ContainsKey($VMName)) {
        Write-Host "VM NOT FOUND in ESXi inventory!" -ForegroundColor Red
        $VMsNotFound++
        $FailedSnapshots++
        $ErrMsg = "VM '$VMName' not present in vim-cmd inventory."

        $VMReportRows += [PSCustomObject]@{
            "VM Name"         = $VMName
            "Power State"     = "Unavailable"
            "CPU"             = "Unavailable"
            "Memory GB"       = "Unavailable"
            "Guest OS"        = "Unavailable"
            "VM Version"      = "Unavailable"
            "Datastore"       = "Unavailable"
            "VM Host"         = $ESXiHostName
            "VM IP Address"   = "Unavailable"
            "Snapshot Result" = "FAILED"
            "Error"           = $ErrMsg
        }

        $SnapshotReportRows += [PSCustomObject]@{
            "VM Name"          = $VMName
            "Snapshot Name"    = $SnapshotName
            "Snapshot Created" = ""
            "Snapshot Size GB" = ""
            "Memory Included"  = $false
            "Quiesced"         = $false
            "Description"      = $SnapshotDesc
            "Snapshot State"   = "FAILED"
            "Result"           = "FAILED"
            "Error"            = $ErrMsg
        }
        continue
    }

    $VMsFound++
    $VMInfo = $VMInventory[$VMName]
    $VMID   = $VMInfo.Vmid

    # Get runtime state
    $PowerStateRaw = (Invoke-ESXiCommand "vim-cmd vmsvc/power.getstate $VMID") -join " "
    $PowerState    = if ($PowerStateRaw -match "Powered on") { "PoweredOn" } elseif ($PowerStateRaw -match "Powered off") { "PoweredOff" } else { "Suspended" }

    # Take Snapshot using vim-cmd: vim-cmd vmsvc/snapshot.create <vmid> <name> <desc> <memory> <quiesce>
    Write-Host "Triggering snapshot for VMID $VMID ($VMName)..." -ForegroundColor Cyan
    $CreateCmd = "vim-cmd vmsvc/snapshot.create $VMID `"$SnapshotName`" `"$SnapshotDesc`" 0 0"
    $SnapOut   = (Invoke-ESXiCommand $CreateCmd) -join "`n"

    # vim-cmd returns empty or "Create Snapshot:" when successful, or an explicit exception/fault
    if ($SnapOut -match "(?i)fault|failed|error") {
        $FailedSnapshots++
        $ErrMsg = $SnapOut.Trim()
        Write-Host "SNAPSHOT FAILED: $ErrMsg" -ForegroundColor Red

        $VMReportRows += [PSCustomObject]@{
            "VM Name"         = $VMName
            "Power State"     = $PowerState
            "CPU"             = "Unavailable"
            "Memory GB"       = "Unavailable"
            "Guest OS"        = $VMInfo.GuestOS
            "VM Version"      = "ESXi 6.7"
            "Datastore"       = $VMInfo.Datastore
            "VM Host"         = $ESXiHostName
            "VM IP Address"   = "Unavailable"
            "Snapshot Result" = "FAILED"
            "Error"           = $ErrMsg
        }

        $SnapshotReportRows += [PSCustomObject]@{
            "VM Name"          = $VMName
            "Snapshot Name"    = $SnapshotName
            "Snapshot Created" = ""
            "Snapshot Size GB" = ""
            "Memory Included"  = $false
            "Quiesced"         = $false
            "Description"      = $SnapshotDesc
            "Snapshot State"   = "FAILED"
            "Result"           = "FAILED"
            "Error"            = $ErrMsg
        }
    }
    else {
        $SuccessfulSnapshots++
        $CreatedTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Write-Host "SNAPSHOT CREATED SUCCESSFULLY" -ForegroundColor Green

        $VMReportRows += [PSCustomObject]@{
            "VM Name"         = $VMName
            "Power State"     = $PowerState
            "CPU"             = "Configured"
            "Memory GB"       = "Configured"
            "Guest OS"        = $VMInfo.GuestOS
            "VM Version"      = "ESXi 6.7"
            "Datastore"       = $VMInfo.Datastore
            "VM Host"         = $ESXiHostName
            "VM IP Address"   = "Unavailable"
            "Snapshot Result" = "SUCCESS"
            "Error"           = ""
        }

        $SnapshotReportRows += [PSCustomObject]@{
            "VM Name"          = $VMName
            "Snapshot Name"    = $SnapshotName
            "Snapshot Created" = $CreatedTime
            "Snapshot Size GB" = "0 (Initial delta)"
            "Memory Included"  = $false
            "Quiesced"         = $false
            "Description"      = $SnapshotDesc
            "Snapshot State"   = "Created"
            "Result"           = "SUCCESS"
            "Error"            = ""
        }
    }
}

Write-Progress -Activity "VMware Snapshot Process (SSH)" -Completed

# =====================================================================
# ESXi SUMMARY ROWS
# =====================================================================

$ESXiReportRows += [PSCustomObject]@{ "Property" = "ESXi Host Name";       "Value" = $ESXiHostName }
$ESXiReportRows += [PSCustomObject]@{ "Property" = "ESXi IP Address";       "Value" = $ESXiHost }
$ESXiReportRows += [PSCustomObject]@{ "Property" = "ESXi Product Version";  "Value" = $ESXiVersionRaw }
$ESXiReportRows += [PSCustomObject]@{ "Property" = "Configured VM Count";   "Value" = $VMList.Count }
$ESXiReportRows += [PSCustomObject]@{ "Property" = "VMs Found";             "Value" = $VMsFound }
$ESXiReportRows += [PSCustomObject]@{ "Property" = "VMs Not Found";         "Value" = $VMsNotFound }
$ESXiReportRows += [PSCustomObject]@{ "Property" = "Successful Snapshots";  "Value" = $SuccessfulSnapshots }
$ESXiReportRows += [PSCustomObject]@{ "Property" = "Failed Snapshots";      "Value" = $FailedSnapshots }
$ESXiReportRows += [PSCustomObject]@{ "Property" = "Report Generated";      "Value" = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') }

# Disconnect SSH
if ($SSHSession) {
    Remove-SSHSession -SSHSession $SSHSession -ErrorAction SilentlyContinue | Out-Null
    Write-Host "`nDisconnected SSH session." -ForegroundColor Green
}

# =====================================================================
# EXPORT EXCEL REPORT
# =====================================================================

Write-Host "`nGenerating Excel Report: $ExcelFile" -ForegroundColor Cyan

if (Test-Path $ExcelFile) {
    Remove-Item $ExcelFile -Force -ErrorAction SilentlyContinue
}

$ESXiReportRows | Export-Excel -Path $ExcelFile -WorksheetName "ESXi Summary" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
$VMReportRows | Export-Excel -Path $ExcelFile -WorksheetName "VM Details" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
$SnapshotReportRows | Export-Excel -Path $ExcelFile -WorksheetName "Snapshot Details" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow

# Formatting widths & styles
$ExcelPackage = Open-ExcelPackage -Path $ExcelFile

$WS1 = $ExcelPackage.Workbook.Worksheets["ESXi Summary"]
if ($WS1) {
    $WS1.Cells["A1:B1"].Style.Font.Bold = $true
    $WS1.Column(1).Width = 30; $WS1.Column(2).Width = 60
}

$WS2 = $ExcelPackage.Workbook.Worksheets["VM Details"]
if ($WS2) {
    $WS2.Cells["A1:K1"].Style.Font.Bold = $true
    $WS2.Column(1).Width = 35; $WS2.Column(2).Width = 15; $WS2.Column(5).Width = 30
    $WS2.Column(7).Width = 25; $WS2.Column(10).Width = 18; $WS2.Column(11).Width = 50
}

$WS3 = $ExcelPackage.Workbook.Worksheets["Snapshot Details"]
if ($WS3) {
    $WS3.Cells["A1:J1"].Style.Font.Bold = $true
    $WS3.Column(1).Width = 35; $WS3.Column(2).Width = 50; $WS3.Column(3).Width = 22
    $WS3.Column(7).Width = 40; $WS3.Column(9).Width = 15; $WS3.Column(10).Width = 50
}

Close-ExcelPackage $ExcelPackage

Write-Host "[ OK ] Excel report ready." -ForegroundColor Green

# Open Excel File
if (Test-Path $ExcelFile) {
    Invoke-Item $ExcelFile
}

Write-Host "`nCompleted: $SuccessfulSnapshots Successful, $FailedSnapshots Failed." -ForegroundColor Cyan
Read-Host "Press Enter to exit"