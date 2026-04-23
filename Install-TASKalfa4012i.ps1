# Install-TASKalfa4012i.ps1
# Downloads the TASKalfa 4012i driver files from a GitHub release zip,
# installs the Kyocera KX driver, creates a TCP/IP port, and adds the printer.
#
# Requires Administrator privileges.
#
# Example:
#   .\Install-TASKalfa4012i.ps1 -IpAddress 10.1.2.50 -PrinterName "Copier - Floor 3"

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$IpAddress,

    [Parameter(Mandatory = $true)]
    [string]$PrinterName,

    # Full URL to the release zip on GitHub. Use the "Release asset" URL, e.g.
    # https://github.com/<owner>/<repo>/releases/download/v1.0.0/TASKalfa4012i_minimal.zip
    # Or the "raw" archive URL if you prefer.
    [string]$ZipUrl = 'https://github.com/ryansuzuki/kyocera4012/releases/latest/download/TASKalfa4012i_minimal.zip',

    # Driver model name as it appears in OEMSETUP.INF.
    # Leave as-is for TASKalfa 4012i; change if you repackage for another model.
    [string]$Model = 'Kyocera TASKalfa 4012i KX',

    [string]$PortName,
    [switch]$DisableSnmp,
    [switch]$KeepFiles   # If set, the extracted driver folder is not deleted after install.
)

function Assert-Admin {
    $isAdmin = ([Security.Principal.WindowsPrincipal] `
            [Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if (-not $isAdmin) {
        throw "This script must be run as Administrator."
    }
}

Assert-Admin

if (-not $PortName) {
    $PortName = "IP_$IpAddress"
}

# Ensure Print Spooler is running
$spooler = Get-Service Spooler -ErrorAction Stop
if ($spooler.Status -ne 'Running') {
    Start-Service Spooler
}

# --- 1. Download the release zip -----------------------------------------
$workRoot  = Join-Path $env:TEMP ("TASKalfa4012i_" + [guid]::NewGuid().ToString('N'))
$zipPath   = Join-Path $workRoot 'driver.zip'
$extractTo = Join-Path $workRoot 'driver'
New-Item -ItemType Directory -Force -Path $workRoot | Out-Null

Write-Host "Downloading driver from $ZipUrl"
# Force TLS 1.2 for older PowerShell / Windows defaults.
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
$ProgressPreference = 'SilentlyContinue'   # faster download on older PS
Invoke-WebRequest -Uri $ZipUrl -OutFile $zipPath -UseBasicParsing

Write-Host "Extracting driver package"
Expand-Archive -Path $zipPath -DestinationPath $extractTo -Force

# The zip contains a top-level folder "TASKalfa4012i_minimal" with OEMSETUP.INF inside.
$infFile = Get-ChildItem -Path $extractTo -Filter 'OEMSETUP.INF' -Recurse -File | Select-Object -First 1
if (-not $infFile) {
    throw "OEMSETUP.INF not found in extracted package at $extractTo"
}
$InfPath = $infFile.FullName
Write-Host "Using INF: $InfPath"

# --- 2. Install the driver ----------------------------------------------
if (-not (Get-PrinterDriver -Name $Model -ErrorAction SilentlyContinue)) {
    # Step 2a: stage the INF into the Windows driver store with pnputil.
    # This is much more reliable than `rundll32 printui.dll,PrintUIEntry /ia`
    # on Windows 10/11 and gives us real error messages on failure.
    Write-Host "Staging driver into Windows driver store (pnputil)..."
    $pnputilOutput = & pnputil.exe /add-driver $InfPath 2>&1
    $pnputilOutput | ForEach-Object { Write-Host "  $_" }
    if ($LASTEXITCODE -ne 0) {
        throw "pnputil failed with exit code $LASTEXITCODE. Driver not staged."
    }

    # Step 2b: register the printer driver under its friendly name.
    Write-Host "Registering printer driver '$Model'..."
    try {
        Add-PrinterDriver -Name $Model -InfPath $InfPath -ErrorAction Stop
    } catch {
        # Fallback: after pnputil staging, the driver is in the Windows driver
        # store (FileRepository) and Add-PrinterDriver can find it by name alone.
        Write-Host "  InfPath form failed ($($_.Exception.Message)); retrying by name..."
        Add-PrinterDriver -Name $Model -ErrorAction Stop
    }

    if (-not (Get-PrinterDriver -Name $Model -ErrorAction SilentlyContinue)) {
        throw "Driver install failed. Check that '$Model' exactly matches an entry in OEMSETUP.INF."
    }
} else {
    Write-Host "Driver '$Model' already installed."
}

# --- 3. Create TCP/IP port ----------------------------------------------
if (-not (Get-PrinterPort -Name $PortName -ErrorAction SilentlyContinue)) {
    Write-Host "Creating TCP/IP port $PortName ($IpAddress)"
    $portParams = @{
        Name               = $PortName
        PrinterHostAddress = $IpAddress
        PortNumber         = 9100
        Protocol           = 1   # 1 = RAW, 2 = LPR
    }
    if ($DisableSnmp) { $portParams.SNMPEnabled = $false }
    Add-PrinterPort @portParams
} else {
    Write-Host "Port $PortName already exists."
}

# --- 4. Create the printer ----------------------------------------------
if (-not (Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue)) {
    Write-Host "Creating printer '$PrinterName'"
    Add-Printer -Name $PrinterName -DriverName $Model -PortName $PortName
} else {
    Write-Host "Printer '$PrinterName' already exists."
}

# --- 5. Clean up --------------------------------------------------------
if (-not $KeepFiles) {
    try { Remove-Item -Path $workRoot -Recurse -Force -ErrorAction Stop }
    catch { Write-Warning "Could not remove temp folder $workRoot : $_" }
} else {
    Write-Host "Driver files kept at: $extractTo"
}

Write-Host "Printer installation complete." -ForegroundColor Green
