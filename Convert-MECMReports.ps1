<#
.SYNOPSIS
    One-click conversion and publishing of MECM Power BI templates to PBIRS.

.DESCRIPTION
    Wrapper script that:
      1. Locates MECM .pbit templates (or uses provided path)
      2. Converts .pbit → .pbix using Power BI Desktop RS
      3. Publishes .pbix to Power BI Report Server via SOAP API
      4. Optionally configures data source credentials

.PARAMETER SqlServer
    MECM SQL Server hostname (e.g. "mecmdb.corp.local" or "mecmdb\INST1").

.PARAMETER DatabaseName
    MECM database name (e.g. "CM_PS1" or "ConfigMgr_CHQ").

.PARAMETER ReportServerUrl
    PBIRS web portal URL (e.g. "https://pbirs:8443/Reports").

.PARAMETER SourcePath
    Path to .pbit templates. Auto-detected if not specified.

.PARAMETER OutputPath
    Directory for converted .pbix files. Defaults to .\output.

.PARAMETER TargetFolder
    PBIRS folder to publish into. Defaults to "/".

.PARAMETER SkipConversion
    Skip the .pbit → .pbix conversion (use existing .pbix files).

.PARAMETER SkipPublish
    Skip publishing to PBIRS (convert only).

.PARAMETER ConfigureDataSource
    Configure data source credentials after publishing.

.EXAMPLE
    .\Convert-MECMReports.ps1 -SqlServer "cm1" -DatabaseName "ConfigMgr_CHQ" -ReportServerUrl "https://cm1:8443/Reports"

.EXAMPLE
    .\Convert-MECMReports.ps1 -SourcePath "C:\templates" -OutputPath "C:\output" -SkipPublish
#>

[CmdletBinding()]
param(
    [string]$SqlServer       = "",
    [string]$DatabaseName    = "",
    [string]$ReportServerUrl = "",
    [string]$SourcePath      = "",
    [string]$OutputPath      = (Join-Path $PSScriptRoot "output"),
    [string]$TargetFolder    = "/",
    [switch]$SkipConversion,
    [switch]$SkipPublish,
    [switch]$ConfigureDataSource,
    [switch]$PromptForDataSourceCredential
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------
# Banner
# ---------------------------------------------------------------
Clear-Host
Write-Host ""
Write-Host "  +================================================================+" -ForegroundColor DarkCyan
Write-Host "  |  MECM Power BI Report Converter (Offline / Air-Gapped)         |" -ForegroundColor DarkCyan
Write-Host "  +================================================================+" -ForegroundColor DarkCyan
Write-Host ""

# ---------------------------------------------------------------
# 1. Gather Configuration
# ---------------------------------------------------------------
if ([string]::IsNullOrWhiteSpace($SqlServer)) {
    $SqlServer = Read-Host '  MECM SQL Server [e.g. mecmdb.corp.local or mecmdb\INST1]'
    while ([string]::IsNullOrWhiteSpace($SqlServer)) {
        Write-Host "  [!] SQL Server is required." -ForegroundColor Red
        $SqlServer = Read-Host "  MECM SQL Server"
    }
}

if ([string]::IsNullOrWhiteSpace($DatabaseName)) {
    $DatabaseName = Read-Host '  MECM Database name [e.g. CM_PS1 or ConfigMgr_CHQ]'
    while ([string]::IsNullOrWhiteSpace($DatabaseName)) {
        Write-Host "  [!] Database name is required." -ForegroundColor Red
        $DatabaseName = Read-Host "  MECM Database name"
    }
}

if (-not $SkipPublish -and [string]::IsNullOrWhiteSpace($ReportServerUrl)) {
    $defaultUrl = "https://$($env:COMPUTERNAME):8443/Reports"
    $ReportServerUrl = Read-Host "  PBIRS Web Portal URL [$defaultUrl]"
    if ([string]::IsNullOrWhiteSpace($ReportServerUrl)) { $ReportServerUrl = $defaultUrl }
}

# ---------------------------------------------------------------
# 2. Locate Templates
# ---------------------------------------------------------------
if ([string]::IsNullOrWhiteSpace($SourcePath)) {
    # Auto-detect MECM built-in templates
    $candidates = @(
        "$env:ProgramFiles\Microsoft Configuration Manager\Reporting\PowerBITemplates",
        "$env:ProgramFiles\Microsoft Endpoint Configuration Manager\Reporting\PowerBITemplates",
        (Join-Path $PSScriptRoot "templates")
    )
    foreach ($c in $candidates) {
        if (Test-Path $c) {
            $count = @(Get-ChildItem $c -Filter '*.pbit' -File -ErrorAction SilentlyContinue).Count
            if ($count -gt 0) {
                $SourcePath = $c
                Write-Host "  [AUTO] Found $count templates in: $c" -ForegroundColor Green
                break
            }
        }
    }
    if ([string]::IsNullOrWhiteSpace($SourcePath)) {
        $SourcePath = Read-Host "  Path to .pbit template files"
    }
}

if (-not (Test-Path $SourcePath)) {
    Write-Host "  [ERROR] Template path not found: $SourcePath" -ForegroundColor Red
    exit 1
}

$pbitFiles = @(Get-ChildItem $SourcePath -Filter '*.pbit' -File -ErrorAction SilentlyContinue)
$pbixFiles = @(Get-ChildItem $SourcePath -Filter '*.pbix' -File -ErrorAction SilentlyContinue)

Write-Host ""
Write-Host "  --------------------------------------------------------" -ForegroundColor DarkGray
Write-Host "  SQL Server:       $SqlServer" -ForegroundColor DarkGray
Write-Host "  Database:         $DatabaseName" -ForegroundColor DarkGray
Write-Host "  Templates:        $SourcePath ($($pbitFiles.Count) .pbit, $($pbixFiles.Count) .pbix)" -ForegroundColor DarkGray
Write-Host "  Output:           $OutputPath" -ForegroundColor DarkGray
if (-not $SkipPublish) {
    Write-Host "  PBIRS:            $ReportServerUrl" -ForegroundColor DarkGray
}
Write-Host "  --------------------------------------------------------" -ForegroundColor DarkGray
Write-Host ""

$confirm = Read-Host "  Proceed? (Y/n)"
if ($confirm -match '^[Nn]') { exit 0 }

# ---------------------------------------------------------------
# 3. Convert .pbit → .pbix
# ---------------------------------------------------------------
if (-not $SkipConversion -and $pbitFiles.Count -gt 0) {
    Write-Host ""
    Write-Host "  [STEP 1] CONVERTING .PBIT TO .PBIX" -ForegroundColor Cyan
    Write-Host "  This requires a visible desktop session (PBI Desktop RS)." -ForegroundColor DarkGray
    Write-Host ""

    $convertScript = Join-Path $PSScriptRoot "scripts\Convert-PbitToPbix.ps1"
    if (-not (Test-Path $convertScript)) {
        Write-Host "  [ERROR] Convert-PbitToPbix.ps1 not found at: $convertScript" -ForegroundColor Red
        exit 1
    }

    & $convertScript -SourcePath $SourcePath -OutputPath $OutputPath -Force

    $convertedFiles = @(Get-ChildItem $OutputPath -Filter '*.pbix' -File -ErrorAction SilentlyContinue)
    Write-Host ""
    Write-Host "  Converted: $($convertedFiles.Count) .pbix files" -ForegroundColor Green
} elseif ($SkipConversion) {
    Write-Host ""
    Write-Host "  [STEP 1] SKIPPED: Conversion (.pbit → .pbix)" -ForegroundColor DarkYellow
    # Use existing .pbix files from source or output
    if ($pbixFiles.Count -gt 0) {
        $OutputPath = $SourcePath
    }
} else {
    Write-Host "  [INFO] No .pbit files found. Looking for .pbix files..." -ForegroundColor DarkGray
    if ($pbixFiles.Count -gt 0) {
        $OutputPath = $SourcePath
        Write-Host "  Found $($pbixFiles.Count) .pbix files to publish." -ForegroundColor Green
    } else {
        Write-Host "  [ERROR] No .pbit or .pbix files found in: $SourcePath" -ForegroundColor Red
        exit 1
    }
}

# ---------------------------------------------------------------
# 4. Publish to PBIRS
# ---------------------------------------------------------------
if (-not $SkipPublish) {
    Write-Host ""
    Write-Host "  [STEP 2] PUBLISHING TO PBIRS" -ForegroundColor Cyan
    Write-Host ""

    $publishScript = Join-Path $PSScriptRoot "scripts\Publish-ToPBIRS.ps1"
    if (-not (Test-Path $publishScript)) {
        Write-Host "  [ERROR] Publish-ToPBIRS.ps1 not found at: $publishScript" -ForegroundColor Red
        exit 1
    }

    $publishArgs = @{
        ReportServerUrl = $ReportServerUrl
        SourcePath      = $OutputPath
        TargetFolder    = $TargetFolder
        SqlServer       = $SqlServer
        DatabaseName    = $DatabaseName
        Overwrite       = $true
    }

    if ($ConfigureDataSource) { $publishArgs.ConfigureDataSource = $true }
    if ($PromptForDataSourceCredential) { $publishArgs.PromptForDataSourceCredential = $true }

    & $publishScript @publishArgs
} else {
    Write-Host ""
    Write-Host "  [STEP 2] SKIPPED: Publishing to PBIRS" -ForegroundColor DarkYellow
    Write-Host "  Your .pbix files are in: $OutputPath" -ForegroundColor White
    Write-Host ""
    Write-Host "  To publish manually:" -ForegroundColor White
    Write-Host "    1. Open each .pbix in Power BI Desktop (RS edition)" -ForegroundColor DarkGray
    Write-Host "    2. File → Publish → Publish to Power BI Report Server" -ForegroundColor DarkGray
    Write-Host "    3. Enter your PBIRS ReportServer URL" -ForegroundColor DarkGray
}

# ---------------------------------------------------------------
# Done
# ---------------------------------------------------------------
Write-Host ""
Write-Host "  +================================================================+" -ForegroundColor Green
Write-Host "  |  COMPLETE                                                      |" -ForegroundColor Green
Write-Host "  +================================================================+" -ForegroundColor Green
Write-Host ""
