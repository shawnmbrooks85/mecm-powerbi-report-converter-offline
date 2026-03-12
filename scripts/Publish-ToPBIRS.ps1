<#
.SYNOPSIS
    Publishes Power BI reports (.pbix) to Power BI Report Server (PBIRS).

.DESCRIPTION
    Uploads .pbix files to PBIRS using the SOAP API via Write-RsCatalogItem
    from the ReportingServicesTools module. This method properly initializes
    the Analysis Services data model, unlike the REST API which only stores
    raw bytes and causes "An error has occurred" when viewing reports.

    Optionally configures data source credentials after upload using the
    PBIRS REST API.

.PARAMETER ReportServerUrl
    The PBIRS web portal URL, e.g. "https://cm1:8443/Reports"

.PARAMETER SourcePath
    Directory containing .pbix files to publish.

.PARAMETER TargetFolder
    PBIRS folder path to publish into. Use "/" for root.

.PARAMETER Overwrite
    If set, overwrites existing reports with the same name.

.PARAMETER ConfigureDataSource
    If set, configures each report's data source with Windows Auth after upload.

.PARAMETER PromptForDataSourceCredential
    Prompts for the Windows credential to store for data sources.

.PARAMETER SqlServer
    SQL Server name for data source configuration.

.PARAMETER DatabaseName
    Database name for data source configuration.

.EXAMPLE
    .\Publish-ToPBIRS.ps1 -ReportServerUrl "https://cm1:8443/Reports" -SourcePath ".\output"

.EXAMPLE
    .\Publish-ToPBIRS.ps1 -ReportServerUrl "https://cm1:8443/Reports" -SourcePath ".\output" -ConfigureDataSource -PromptForDataSourceCredential
#>

[CmdletBinding()]
param(
    [string]$ReportServerUrl = "https://$($env:COMPUTERNAME):8443/Reports",
    [string]$SourcePath      = "",
    [string]$TargetFolder    = "/",
    [switch]$Overwrite,
    [switch]$ConfigureDataSource,
    [switch]$PromptForDataSourceCredential,
    [string]$DataSourceUsername = ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name),
    [string]$SqlServer       = $env:COMPUTERNAME,
    [string]$DatabaseName    = ""
)

$ErrorActionPreference = 'Stop'

# --- Name mapping: MECM template names -> dashboard-expected names ---
$nameMap = @{
    'Client Status'                      = 'Microsoft Client Status'
    'Content Status'                     = 'Microsoft Content Status'
    'Software Update Compliance Status'  = 'Microsoft Software Update Compliance Status'
    'Software Update Deployment Status'  = 'Microsoft Software Update Deployment Status'
}

if ([string]::IsNullOrWhiteSpace($SourcePath)) {
    $SourcePath = Join-Path (Split-Path $PSScriptRoot) "output"
    if (-not (Test-Path $SourcePath)) {
        $SourcePath = Read-Host "  Path to .pbix files"
    }
}

if (-not (Test-Path $SourcePath)) {
    Write-Host "  [ERROR] Source path not found: $SourcePath" -ForegroundColor Red
    exit 1
}

# --- Install/import ReportingServicesTools ---
function Initialize-ReportingServicesTools {
    if (Get-Module -Name ReportingServicesTools) { return $true }
    if (Get-Module -ListAvailable -Name ReportingServicesTools) {
        Import-Module ReportingServicesTools -ErrorAction Stop
        return $true
    }
    Write-Host "    [INFO] ReportingServicesTools not found. Installing..." -ForegroundColor DarkYellow
    try {
        if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
            Install-PackageProvider -Name NuGet -MinimumVersion '2.8.5.201' -Scope CurrentUser -Force | Out-Null
        }
        Install-Module -Name ReportingServicesTools -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop | Out-Null
        Import-Module ReportingServicesTools -ErrorAction Stop
        return $true
    }
    catch {
        Write-Host "    [ERROR] Unable to install ReportingServicesTools: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# --- Trust self-signed certs ---
if ($PSVersionTable.PSVersion.Major -ge 6) {
    $splatWeb = @{ SkipCertificateCheck = $true }
} else {
    try {
        if (-not ([System.Management.Automation.PSTypeName]'TrustAllCertsPolicy').Type) {
            Add-Type -TypeDefinition 'using System.Net; using System.Security.Cryptography.X509Certificates; public class TrustAllCertsPolicy : ICertificatePolicy { public bool CheckValidationResult(ServicePoint sp, X509Certificate cert, WebRequest req, int problem) { return true; } }'
        }
        [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
    } catch { }
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    $splatWeb = @{}
}

# --- Banner ---
Write-Host ""
Write-Host "  +------------------------------------------------+" -ForegroundColor DarkCyan
Write-Host "  |  Publish Power BI Reports to PBIRS (SOAP API)  |" -ForegroundColor DarkCyan
Write-Host "  +------------------------------------------------+" -ForegroundColor DarkCyan
Write-Host "  Server:  $ReportServerUrl" -ForegroundColor DarkGray
Write-Host "  Source:  $SourcePath" -ForegroundColor DarkGray
Write-Host "  Target:  $TargetFolder" -ForegroundColor DarkGray
Write-Host ""

# --- Find report files ---
$reportFiles = @(Get-ChildItem $SourcePath -Filter '*.pbix' -File)
if ($reportFiles.Count -eq 0) {
    Write-Host "  [ERROR] No .pbix files found in: $SourcePath" -ForegroundColor Red
    exit 1
}
Write-Host "  [INFO] Found $($reportFiles.Count) .pbix file(s) to publish." -ForegroundColor Cyan
Write-Host ""

# --- Initialize ReportingServicesTools ---
if (-not (Initialize-ReportingServicesTools)) {
    Write-Host "  [ERROR] ReportingServicesTools module is required." -ForegroundColor Red
    Write-Host "  Install: Install-Module ReportingServicesTools -Scope CurrentUser" -ForegroundColor Yellow
    exit 1
}

# Derive the SOAP endpoint from the portal URL
$rsWebServiceUrl = $ReportServerUrl -replace '/Reports\s*$', '/ReportServer'
Write-Host "  SOAP endpoint: $rsWebServiceUrl" -ForegroundColor DarkGray
Write-Host ""

# --- Create target folder if needed ---
if ($TargetFolder -ne "/") {
    try {
        $baseApi = $ReportServerUrl.TrimEnd('/') + "/api/v2.0"
        $folders = Invoke-RestMethod -Uri "$baseApi/Folders" -UseDefaultCredentials @splatWeb
        $exists = $folders.value | Where-Object { $_.Path -eq $TargetFolder }
        if (-not $exists) {
            $folderName = $TargetFolder.Split('/')[-1]
            $folderBody = @{ Name = $folderName; Path = "/"; '@odata.type' = '#Model.Folder' } | ConvertTo-Json
            Invoke-RestMethod -Uri "$baseApi/Folders" -UseDefaultCredentials -Method POST -ContentType "application/json" -Body $folderBody @splatWeb | Out-Null
            Write-Host "  [OK] Created folder: $TargetFolder" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "  [WARN] Could not verify/create folder: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# --- Upload each report via Write-RsCatalogItem (SOAP API) ---
$uploaded = @()
$step = 1

foreach ($file in $reportFiles) {
    $reportName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)

    # Apply name mapping
    $needsRename = $false
    if ($nameMap.ContainsKey($reportName)) {
        $mappedName = $nameMap[$reportName]
        Write-Host "    [MAP] '$reportName' -> '$mappedName'" -ForegroundColor DarkGray
        $needsRename = $true
    } else {
        $mappedName = $reportName
    }

    Write-Host ""
    Write-Host "  [$step/$($reportFiles.Count)] PUBLISHING: $mappedName" -ForegroundColor Cyan
    $sizeMB = [math]::Round($file.Length / 1MB, 2)
    Write-Host "    File size: $sizeMB MB" -ForegroundColor DarkGray

    try {
        if ($needsRename) {
            $tmpDir = Join-Path $env:TEMP "pbirs_publish"
            if (-not (Test-Path $tmpDir)) { New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null }
            $tmpFile = Join-Path $tmpDir "$mappedName.pbix"
            Copy-Item -Path $file.FullName -Destination $tmpFile -Force
            $publishPath = $tmpFile
        } else {
            $publishPath = $file.FullName
        }

        Write-RsCatalogItem -Path $publishPath -RsFolder $TargetFolder -ReportServerUri $rsWebServiceUrl -Overwrite -ErrorAction Stop

        Write-Host "    [OK] Published: $mappedName" -ForegroundColor Green
        $itemPath = if ($TargetFolder -eq "/") { "/$mappedName" } else { "$($TargetFolder.TrimEnd('/'))/$mappedName" }
        $uploaded += [PSCustomObject]@{ Name = $mappedName; Path = $itemPath }

        if ($needsRename -and (Test-Path $tmpFile)) {
            Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue
        }
    }
    catch {
        Write-Host "    [ERROR] Failed to publish '$mappedName': $($_.Exception.Message)" -ForegroundColor Red
    }

    $step++
}

# --- Configure data sources ---
if ($ConfigureDataSource -and $uploaded.Count -gt 0) {
    Write-Host ""
    Write-Host "  CONFIGURING DATA SOURCES" -ForegroundColor Cyan
    Write-Host "    SQL Server: $SqlServer" -ForegroundColor DarkGray
    Write-Host "    Database:   $DatabaseName" -ForegroundColor DarkGray
    Write-Host ""

    $cred = $null
    if ($PromptForDataSourceCredential) {
        $cred = Get-Credential -Message "Enter Windows credential for PBIRS data sources" -UserName $DataSourceUsername
    }

    if ($cred) {
        $plainPwd = [Runtime.InteropServices.Marshal]::PtrToStringBSTR(
            [Runtime.InteropServices.Marshal]::SecureStringToBSTR($cred.Password)
        )

        try {
            foreach ($report in $uploaded) {
                try {
                    $dataSources = @(Get-RsRestItemDataSource -ReportPortalUri $ReportServerUrl -RsItem $report.Path -ErrorAction Stop)
                    if ($dataSources.Count -eq 0) {
                        Write-Host "    [WARN] $($report.Name): no data sources found." -ForegroundColor Yellow
                        continue
                    }

                    foreach ($ds in $dataSources) {
                        if ($ds.DataSourceSubType -ne 'DataModel' -or -not $ds.PSObject.Properties['DataModelDataSource']) { continue }
                        $ds.DataModelDataSource.AuthType = 'Windows'
                        $ds.DataModelDataSource.Username = $cred.UserName
                        $ds.DataModelDataSource.Secret = $plainPwd
                    }

                    Set-RsRestItemDataSource -ReportPortalUri $ReportServerUrl -RsItem $report.Path -RsItemType PowerBIReport -DataSources $dataSources -Confirm:$false -ErrorAction Stop

                    $testResults = @(Test-RsRestItemDataSource -ReportPortalUri $ReportServerUrl -RsReport $report.Path -ErrorAction Stop)
                    $failed = @($testResults | Where-Object { -not $_.IsSuccessful })

                    if ($failed.Count -eq 0) {
                        Write-Host "    [OK] $($report.Name): data source configured" -ForegroundColor Green
                    } else {
                        Write-Host "    [WARN] $($report.Name): credentials stored, connection test failed" -ForegroundColor Yellow
                    }
                }
                catch {
                    Write-Host "    [WARN] $($report.Name): $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
        }
        finally {
            $plainPwd = $null
        }
    } else {
        Write-Host "    [INFO] No credential provided. Configure data sources manually in PBIRS portal." -ForegroundColor DarkYellow
    }
}

# --- Summary ---
Write-Host ""
Write-Host "  +------------------------------------------------+" -ForegroundColor Green
Write-Host "  |  PUBLISH COMPLETE                              |" -ForegroundColor Green
Write-Host "  +------------------------------------------------+" -ForegroundColor Green
Write-Host "  Published: $($uploaded.Count)/$($reportFiles.Count) reports" -ForegroundColor White
Write-Host ""

if ($uploaded.Count -gt 0) {
    Write-Host "  Report URLs:" -ForegroundColor DarkGray
    foreach ($r in $uploaded) {
        $encodedName = [Uri]::EscapeDataString($r.Name)
        $folder = ""
        if ($TargetFolder -ne "/") { $folder = $TargetFolder.TrimStart('/') + "/" }
        Write-Host "    $ReportServerUrl/powerbi/$folder$encodedName" -ForegroundColor White
    }
    Write-Host ""
}
