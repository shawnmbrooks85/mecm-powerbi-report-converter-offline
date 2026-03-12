<#
.SYNOPSIS
    Converts .pbit templates to .pbix reports using Power BI Desktop RS.

.DESCRIPTION
    Opens each .pbit file in Power BI Desktop (RS), waits for the data model
    to load, then saves as .pbix. Uses Win32 API for reliable window focus
    and keyboard automation.

    Requires: Power BI Desktop (Report Server) installed.
    Must run in a visible desktop session (not headless).

.PARAMETER SourcePath
    Directory containing .pbit files.

.PARAMETER OutputPath
    Directory to save converted .pbix files. Defaults to SourcePath.

.PARAMETER PBIDesktopPath
    Path to PBIDesktop.exe. Auto-detected if not specified.

.PARAMETER TimeoutSeconds
    Max seconds to wait for PBI Desktop to load each report. Default: 180.
#>

[CmdletBinding()]
param(
    [string]$SourcePath     = "",
    [string]$OutputPath     = "",
    [string]$PBIDesktopPath = "",
    [int]$TimeoutSeconds    = 180,
    [switch]$Force
)

if ([string]::IsNullOrWhiteSpace($SourcePath)) {
    $repoReportsPath = Join-Path (Split-Path $PSScriptRoot) '..\reports'
    $repoReportsPath = [System.IO.Path]::GetFullPath($repoReportsPath)
    if (Test-Path $repoReportsPath) {
        $SourcePath = $repoReportsPath
    }
    else {
        $SourcePath = "C:\tmp\MECM PowerBI Dashboards\FixedOffline"
    }
}

$ErrorActionPreference = 'Stop'

# --- Pre-flight: Disable native query prompt (critical for template loading) ---
$regPath = "HKCU:\SOFTWARE\Microsoft\Microsoft Power BI Desktop SSRS"
if (-not (Test-Path $regPath)) { New-Item $regPath -Force | Out-Null }
Set-ItemProperty $regPath -Name "AllowNativeQueries" -Value 1 -Type DWord -ErrorAction SilentlyContinue
Set-ItemProperty $regPath -Name "EnableNativeQueryPrompt" -Value 0 -Type DWord -ErrorAction SilentlyContinue
Write-Host "  [PRE] Native query prompt disabled via registry" -ForegroundColor DarkGray

# --- Pre-flight: Interactive session check ---
if ([Environment]::UserInteractive -eq $false) {
    Write-Host "  [WARN] This script must run in an interactive desktop session (RDP/console)." -ForegroundColor Yellow
    Write-Host "         PBI Desktop RS requires a visible desktop for Save automation." -ForegroundColor Yellow
}

# --- Win32 helpers for window focus (fallback) ---
Add-Type @'
using System;
using System.Runtime.InteropServices;

public class Win32Focus {
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    public const int SW_RESTORE = 9;
    public const int SW_SHOW = 5;

    public static void FocusWindow(IntPtr hWnd) {
        ShowWindow(hWnd, SW_RESTORE);
        SetForegroundWindow(hWnd);
    }
}
'@

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

# --- WScript.Shell COM for reliable SendKeys from background processes ---
$script:wshShell = New-Object -ComObject WScript.Shell

# --- Find PBI Desktop RS ---
if ([string]::IsNullOrEmpty($PBIDesktopPath)) {
    $searchPaths = @(
        "C:\Program Files\Microsoft Power BI Desktop RS\bin\PBIDesktop.exe",
        "C:\Program Files (x86)\Microsoft Power BI Desktop RS\bin\PBIDesktop.exe"
    )
    foreach ($p in $searchPaths) {
        if (Test-Path $p) { $PBIDesktopPath = $p; break }
    }
}

if ([string]::IsNullOrEmpty($PBIDesktopPath) -or -not (Test-Path $PBIDesktopPath)) {
    Write-Host "  [ERROR] Power BI Desktop RS not found. Specify -PBIDesktopPath." -ForegroundColor Red
    exit 1
}

if ([string]::IsNullOrEmpty($OutputPath)) { $OutputPath = $SourcePath }
if (-not (Test-Path $OutputPath)) { New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null }

Write-Host ""
Write-Host "  +------------------------------------------------+" -ForegroundColor DarkCyan
Write-Host "  |  Convert .pbit Templates to .pbix Reports      |" -ForegroundColor DarkCyan
Write-Host "  +------------------------------------------------+" -ForegroundColor DarkCyan
Write-Host "  PBI Desktop: $PBIDesktopPath" -ForegroundColor DarkGray
Write-Host "  Source:       $SourcePath" -ForegroundColor DarkGray
Write-Host "  Output:       $OutputPath" -ForegroundColor DarkGray
Write-Host ""

function Test-PbixPackage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path $Path)) {
        return $false
    }

    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
        try {
            return (($archive.GetEntry("Report/Layout") -ne $null) -and ($archive.GetEntry("DataModel") -ne $null))
        }
        finally {
            $archive.Dispose()
        }
    }
    catch {
        return $false
    }
}

function Get-PendingNativeQueryResourcePaths {
    $userZip = Join-Path $env:LOCALAPPDATA "Microsoft\Power BI Desktop SSRS\User.zip"
    if (-not (Test-Path $userZip)) {
        return @()
    }

    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $archive = [System.IO.Compression.ZipFile]::OpenRead($userZip)
        try {
            $entry = $archive.GetEntry("NativeQueries/NativeQueries.xml")
            if (-not $entry) {
                return @()
            }

            $reader = New-Object System.IO.StreamReader($entry.Open())
            try {
                [xml]$xml = $reader.ReadToEnd()
            }
            finally {
                $reader.Dispose()
            }
        }
        finally {
            $archive.Dispose()
        }

        return @(
            $xml.NativeQueriesList.NativeQueries.NativeQuery |
                Where-Object { $_.QueryPermissionType -eq 'EvaluateNativeQueryUnpermitted' } |
                ForEach-Object { $_.NonNormalizedResourcePath, $_.ResourcePath } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Select-Object -Unique
        )
    }
    catch {
        return @()
    }
}

function Test-TemplateLoadPrompt {
    param(
        [Parameter(Mandatory = $true)]
        [IntPtr]$WindowHandle
    )

    if ($WindowHandle -eq [IntPtr]::Zero) {
        return $false
    }

    try {
        $window = [System.Windows.Automation.AutomationElement]::FromHandle($WindowHandle)
        if (-not $window) {
            return $false
        }

        $elements = $window.FindAll(
            [System.Windows.Automation.TreeScope]::Descendants,
            [System.Windows.Automation.Condition]::TrueCondition
        )

        return @(
            $elements |
                Where-Object { $_.Current.Name -eq 'Once loaded, your data will appear in the Data pane.' }
        ).Count -gt 0
    }
    catch {
        return $false
    }
}

# --- Find .pbit files ---
$pbitFiles = @(Get-ChildItem -Path $SourcePath -Filter "*.pbit" -File)
if ($pbitFiles.Count -eq 0) {
    Write-Host "  [INFO] No .pbit files found in: $SourcePath" -ForegroundColor DarkYellow
    exit 0
}

# Check if .pbix already exist for all of them
$needConversion = @()
foreach ($f in $pbitFiles) {
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
    $pbixTarget = Join-Path $OutputPath "$baseName.pbix"
    $mustConvert = $Force.IsPresent -or (-not (Test-Path $pbixTarget))

    if (-not $mustConvert -and (Get-Item $pbixTarget).LastWriteTimeUtc -lt $f.LastWriteTimeUtc) {
        $mustConvert = $true
    }

    if (-not $mustConvert -and -not (Test-PbixPackage -Path $pbixTarget)) {
        $mustConvert = $true
    }

    if ($mustConvert) {
        if (Test-Path $pbixTarget) {
            Write-Host "  [REPLACE] $baseName.pbix will be regenerated." -ForegroundColor DarkYellow
            Remove-Item $pbixTarget -Force -ErrorAction SilentlyContinue
        }

        $needConversion += $f
    }
    else {
        Write-Host "  [SKIP] $baseName.pbix already exists and passed validation." -ForegroundColor DarkYellow
    }
}

if ($needConversion.Count -eq 0) {
    Write-Host "  All templates already converted." -ForegroundColor Green
    exit 0
}

Write-Host "  Converting $($needConversion.Count) template(s):" -ForegroundColor Cyan
foreach ($f in $needConversion) { Write-Host "    - $($f.Name)" -ForegroundColor DarkGray }
Write-Host ""

$converted = 0
$idx = 0

foreach ($pbit in $needConversion) {
    $idx++
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($pbit.Name)
    $pbixPath = Join-Path $OutputPath "$baseName.pbix"

    Write-Host "  [$idx/$($needConversion.Count)] Converting: $baseName" -ForegroundColor Cyan

    # Kill any existing PBI Desktop
    Get-Process -Name "PBIDesktop" -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Seconds 3

    # Open the .pbit file
    Write-Host "    Opening in PBI Desktop..." -ForegroundColor DarkGray
    Start-Process -FilePath $PBIDesktopPath -ArgumentList "`"$($pbit.FullName)`""

    # Wait for PBI Desktop to fully load
    $elapsed = 0
    $windowTitle = ""
    $pbiHwnd = [IntPtr]::Zero

    while ($elapsed -lt $TimeoutSeconds) {
        Start-Sleep -Seconds 5
        $elapsed += 5

        $proc = Get-Process -Name "PBIDesktop" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($proc -and $proc.MainWindowHandle -ne [IntPtr]::Zero -and $proc.MainWindowTitle -ne "") {
            $windowTitle = $proc.MainWindowTitle
            $pbiHwnd = $proc.MainWindowHandle

            # When the title contains the file name or shows "Untitled", it has loaded
            if ($windowTitle -match "Power BI Desktop") {
                Write-Host "    Window: $windowTitle (${elapsed}s)" -ForegroundColor DarkGray

                # Check if there's a parameter prompt dialog — wait for it to close
                # Give PBI Desktop time to fully initialize and process the template
                # Look for the title to stabilize (not "Loading...")
                Start-Sleep -Seconds 10

                # Re-check title after waiting
                $proc = Get-Process -Name "PBIDesktop" -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($proc) {
                    $windowTitle = $proc.MainWindowTitle
                    $pbiHwnd = $proc.MainWindowHandle
                    Write-Host "    Ready: $windowTitle" -ForegroundColor DarkGray
                }
                break
            }
        }

        if ($elapsed % 30 -eq 0) {
            Write-Host "    Waiting... (${elapsed}s)" -ForegroundColor DarkGray
        }
    }

    if ($pbiHwnd -eq [IntPtr]::Zero) {
        Write-Host "    [ERROR] PBI Desktop did not load within ${TimeoutSeconds}s." -ForegroundColor Red
        Get-Process -Name "PBIDesktop" -ErrorAction SilentlyContinue | Stop-Process -Force
        Start-Sleep -Seconds 3
        continue
    }

    # Wait additional time for data model initialization
    if (Test-TemplateLoadPrompt -WindowHandle $pbiHwnd) {
        Write-Host "    [INFO] Template parameter prompt detected. Auto-clicking Load..." -ForegroundColor Cyan

        # Focus the PBI Desktop window and send Enter to click the Load button (default button)
        Start-Sleep -Seconds 2
        [Win32Focus]::FocusWindow($pbiHwnd)
        Start-Sleep -Milliseconds 500
        [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
        Write-Host "    [INFO] Sent Enter to activate Load button." -ForegroundColor DarkGray
        Start-Sleep -Seconds 3

        # If prompt is still showing, try Tab to reach Load button then Enter
        if (Test-TemplateLoadPrompt -WindowHandle $pbiHwnd) {
            Write-Host "    [INFO] Prompt still visible, trying Tab+Enter fallback..." -ForegroundColor DarkYellow
            [Win32Focus]::FocusWindow($pbiHwnd)
            Start-Sleep -Milliseconds 300
            # Tab through the parameter fields to reach the Load button
            [System.Windows.Forms.SendKeys]::SendWait("{TAB}{TAB}{TAB}{TAB}{TAB}")
            Start-Sleep -Milliseconds 300
            [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
            Start-Sleep -Seconds 3
        }

        # If still showing, try Alt+L (potential access key for Load)
        if (Test-TemplateLoadPrompt -WindowHandle $pbiHwnd) {
            Write-Host "    [INFO] Trying Alt+L fallback..." -ForegroundColor DarkYellow
            [Win32Focus]::FocusWindow($pbiHwnd)
            Start-Sleep -Milliseconds 300
            [System.Windows.Forms.SendKeys]::SendWait("%l")
            Start-Sleep -Seconds 3
        }

        # Wait for data to load (prompt should disappear)
        $promptWait = 0
        while ($promptWait -lt $TimeoutSeconds -and (Test-TemplateLoadPrompt -WindowHandle $pbiHwnd)) {
            Start-Sleep -Seconds 5
            $promptWait += 5

            if ($promptWait % 30 -eq 0) {
                Write-Host "    Waiting for data to load... (${promptWait}s)" -ForegroundColor DarkGray
            }
        }

        if (Test-TemplateLoadPrompt -WindowHandle $pbiHwnd) {
            Write-Host "    [ERROR] Template parameter prompt did not complete after auto-click attempts." -ForegroundColor Red
            Write-Host "             Please click 'Load' manually in the PBI Desktop RS window." -ForegroundColor DarkYellow
            # Give the operator 60s to click manually
            $manualWait = 0
            while ($manualWait -lt 60 -and (Test-TemplateLoadPrompt -WindowHandle $pbiHwnd)) {
                Start-Sleep -Seconds 5
                $manualWait += 5
            }
            if (Test-TemplateLoadPrompt -WindowHandle $pbiHwnd) {
                Write-Host "    [ERROR] Template prompt still not dismissed. Skipping." -ForegroundColor Red
                Get-Process -Name "PBIDesktop" -ErrorAction SilentlyContinue | Stop-Process -Force
                Start-Sleep -Seconds 3
                continue
            }
        }

        Write-Host "    [OK] Template parameter prompt completed." -ForegroundColor Green
        Start-Sleep -Seconds 10
    }

    Write-Host "    Waiting for data model..." -ForegroundColor DarkGray
    Start-Sleep -Seconds 20

    # --- Save As .pbix using WScript.Shell (PID-based activation) ---
    # PBI Desktop RS (Jan 2026+) uses Chromium-based UI; WScript.Shell.AppActivate(PID)
    # is the most reliable method for bringing it to focus and sending keystrokes.
    Write-Host "    Saving as: $pbixPath" -ForegroundColor DarkGray

    try {
        $proc = Get-Process -Name "PBIDesktop" -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $proc) {
            Write-Host "    [ERROR] PBI Desktop process not found." -ForegroundColor Red
            continue
        }

        # Activate via PID (works from background processes, unlike title-based activation)
        $activated = $script:wshShell.AppActivate($proc.Id)
        Write-Host "    AppActivate(PID=$($proc.Id)): $activated" -ForegroundColor DarkGray

        if (-not $activated) {
            # Fallback: try Win32 API
            Write-Host "    [WARN] PID activation failed, trying Win32 fallback..." -ForegroundColor Yellow
            [Win32Focus]::FocusWindow($pbiHwnd)
            Start-Sleep -Milliseconds 500
            $activated = $script:wshShell.AppActivate($proc.Id)
        }

        if ($activated) {
            Start-Sleep -Seconds 1

            # Build the output path (without file extension — PBI adds .pbix)
            $outDir = [System.IO.Path]::GetDirectoryName($pbixPath)
            $outBase = [System.IO.Path]::GetFileNameWithoutExtension($pbixPath)
            $typePath = Join-Path $outDir $outBase

            # Escape special SendKeys characters in the path: + ^ % ~ ( ) { } [ ]
            $escapedPath = $typePath -replace '([\+\^\%\~\(\)\{\}\[\]])', '{$1}'

            # Send Ctrl+S (opens Save dialog in PBI Desktop RS)
            $script:wshShell.SendKeys("^s")
            Start-Sleep -Seconds 5

            # Re-activate after Save dialog opens (it may steal focus from the shell)
            $script:wshShell.AppActivate($proc.Id)
            Start-Sleep -Seconds 1

            # Focus the File name field using Alt+N (standard Windows Save dialog accelerator)
            $script:wshShell.SendKeys("%n")
            Start-Sleep -Milliseconds 500

            # Select all existing text: Home then Shift+End (more reliable than Ctrl+A with Chromium)
            $script:wshShell.SendKeys("{HOME}")
            Start-Sleep -Milliseconds 200
            $script:wshShell.SendKeys("+{END}")
            Start-Sleep -Milliseconds 300

            # Type the output path (replaces selection)
            $script:wshShell.SendKeys($escapedPath)
            Start-Sleep -Seconds 1

            # Press Enter to confirm save
            $script:wshShell.SendKeys("{ENTER}")
            Start-Sleep -Seconds 5

            # Handle any overwrite confirmation dialog
            $script:wshShell.SendKeys("{ENTER}")
            Start-Sleep -Seconds 5
        }

        # Verify the file was created with DataModel
        if ((Test-Path $pbixPath) -and (Test-PbixPackage -Path $pbixPath)) {
            $sizeMB = [math]::Round((Get-Item $pbixPath).Length / 1MB, 2)
            Write-Host "    [OK] Saved: $baseName.pbix ($sizeMB MB)" -ForegroundColor Green
            $converted++
        }
        else {
            # Second attempt: Ctrl+Shift+S (Save A Copy)
            Write-Host "    [WARN] First save attempt failed. Trying Ctrl+Shift+S..." -ForegroundColor Yellow
            $script:wshShell.AppActivate($proc.Id)
            Start-Sleep -Seconds 1
            $script:wshShell.SendKeys("^+s")
            Start-Sleep -Seconds 5

            # Re-activate and focus filename field
            $script:wshShell.AppActivate($proc.Id)
            Start-Sleep -Seconds 1
            $script:wshShell.SendKeys("%n")
            Start-Sleep -Milliseconds 500
            $script:wshShell.SendKeys("{HOME}")
            Start-Sleep -Milliseconds 200
            $script:wshShell.SendKeys("+{END}")
            Start-Sleep -Milliseconds 300
            $script:wshShell.SendKeys($escapedPath)
            Start-Sleep -Seconds 1
            $script:wshShell.SendKeys("{ENTER}")
            Start-Sleep -Seconds 5
            $script:wshShell.SendKeys("{ENTER}")
            Start-Sleep -Seconds 5

            if ((Test-Path $pbixPath) -and (Test-PbixPackage -Path $pbixPath)) {
                $sizeMB = [math]::Round((Get-Item $pbixPath).Length / 1MB, 2)
                Write-Host "    [OK] Saved: $baseName.pbix ($sizeMB MB)" -ForegroundColor Green
                $converted++
            }
            else {
                if (Test-Path $pbixPath) {
                    Remove-Item $pbixPath -Force -ErrorAction SilentlyContinue
                }
                Write-Host "    [ERROR] Could not save $baseName.pbix" -ForegroundColor Red
                Write-Host "    [HINT] Try running this script from an interactive RDP session" -ForegroundColor DarkYellow
                Write-Host "           (not via remote PowerShell or scheduled task)." -ForegroundColor DarkYellow
            }
        }
    }
    catch {
        Write-Host "    [ERROR] Automation failed: $($_.Exception.Message)" -ForegroundColor Red
    }

    # Close PBI Desktop
    Write-Host "    Closing PBI Desktop..." -ForegroundColor DarkGray
    Get-Process -Name "PBIDesktop" -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Seconds 5
}

# --- Summary ---
Write-Host ""
Write-Host "  +------------------------------------------------+" -ForegroundColor Green
Write-Host "  |  CONVERSION COMPLETE                           |" -ForegroundColor Green
Write-Host "  +------------------------------------------------+" -ForegroundColor Green
Write-Host "  Converted: $converted/$($needConversion.Count) templates" -ForegroundColor White
Write-Host "  Output:    $OutputPath" -ForegroundColor DarkGray
Write-Host ""

$pbixFiles = @(Get-ChildItem -Path $OutputPath -Filter "*.pbix" -File)
if ($pbixFiles.Count -gt 0) {
    Write-Host "  Ready for publishing:" -ForegroundColor Cyan
    foreach ($f in $pbixFiles) {
        $sizeMB = [math]::Round($f.Length / 1MB, 2)
        Write-Host "    - $($f.Name) ($sizeMB MB)" -ForegroundColor DarkGray
    }
    Write-Host ""
}
