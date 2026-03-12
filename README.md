# MECM Power BI Report Converter (Offline / Air-Gapped)

Scripts and guidance to modernize legacy Microsoft Configuration Manager (MECM/SCCM) sample Power BI reports for current environments, including **disconnected and offline networks**.

## The Problem

Microsoft's sample MECM Power BI templates (`.pbit`) were designed for Power BI Service (cloud). In air-gapped, SIPR, or disconnected environments you must use **Power BI Report Server (PBIRS)** — but the templates don't just work out of the box:

- Templates prompt for **native query approval** which blocks automation
- `.pbit` files must be converted to `.pbix` before publishing to PBIRS
- PBIRS REST API uploads produce reports that **fail to render** ("An error has occurred")
- Connection strings and data sources need reconfiguration for each environment
- SSL certificates must be configured for PBIRS on non-standard ports (e.g., 8443)

## What This Repo Does

| Script | Purpose |
|--------|---------|
| `Convert-MECMReports.ps1` | **One-click wrapper** — converts `.pbit` → `.pbix` and publishes to PBIRS |
| `scripts/Convert-PbitToPbix.ps1` | Converts `.pbit` templates to `.pbix` using PBI Desktop RS automation |
| `scripts/Publish-ToPBIRS.ps1` | Publishes `.pbix` reports to PBIRS via SOAP API (`Write-RsCatalogItem`) |

## Prerequisites

| Requirement | Notes |
|-------------|-------|
| **Power BI Desktop (Report Server)** | Must be the RS edition, not the Store version |
| **Power BI Report Server** | Installed and configured with HTTPS |
| **MECM .pbit Templates** | Usually at `C:\Program Files\Microsoft Configuration Manager\Reporting\PowerBITemplates` |
| **PowerShell 5.1+** | Windows PowerShell (not PowerShell Core) |
| **ReportingServicesTools** | Auto-installed if missing (`Install-Module ReportingServicesTools`) |

## Quick Start

### Option 1: Full Automation (Convert + Publish)

```powershell
# Run from an interactive desktop session (RDP or console)
.\Convert-MECMReports.ps1 -SqlServer "mecmdb.corp.local" -DatabaseName "CM_PS1" -ReportServerUrl "https://pbirs.corp.local:8443/Reports"
```

### Option 2: Convert Only (Manual Publish Later)

```powershell
# Convert .pbit templates to .pbix
.\scripts\Convert-PbitToPbix.ps1 -SourcePath "C:\path\to\templates" -OutputPath "C:\path\to\output"
```

### Option 3: Publish Only (Already Have .pbix Files)

```powershell
# Publish existing .pbix files to PBIRS
.\scripts\Publish-ToPBIRS.ps1 -ReportServerUrl "https://pbirs:8443/Reports" -SourcePath "C:\path\to\pbix"
```

## How It Works

### 1. Template Conversion (`.pbit` → `.pbix`)

The conversion script:
1. Disables the native query prompt via registry (`AllowNativeQueries = 1`)
2. Opens each `.pbit` in Power BI Desktop RS
3. Waits for the model to load (monitors window title for "Loading model")
4. Sends `Ctrl+Shift+S` (Save As) with Win32 API focus management
5. Saves as `.pbix` to the output directory

> **Important**: This requires a **visible desktop session**. PBI Desktop RS cannot run headless.

### 2. Publishing to PBIRS

The publish script uses `Write-RsCatalogItem` from the **ReportingServicesTools** module, which communicates via the **SOAP API**. This is critical because:

| Upload Method | Initializes Model? | Reports Render? |
|---------------|:-------------------:|:---------------:|
| REST API (`/api/v2.0/CatalogItems`) | ❌ | ❌ "An error has occurred" |
| `Write-RsCatalogItem` (SOAP) | ✅ | ✅ |
| PBI Desktop → Publish (XMLA) | ✅ | ✅ |

The REST API only stores raw bytes — it doesn't tell PBIRS's Analysis Services engine to initialize the data model. The SOAP API handles this properly.

### 3. Data Source Configuration

After publishing, each report's data source must be configured with credentials. The script can do this automatically:

```powershell
.\scripts\Publish-ToPBIRS.ps1 -ReportServerUrl "https://pbirs:8443/Reports" -SourcePath ".\output" -ConfigureDataSource -PromptForDataSourceCredential -SqlServer "mecmdb" -DatabaseName "CM_PS1"
```

## PBIRS Port Configuration

In MECM environments, PBIRS cannot use port 443 because the MECM Management Point already occupies it. Common alternative: **port 8443**.

### Steps to Configure PBIRS on Port 8443

1. **Power BI Report Server Configuration Manager** → connect to `SERVERNAME\PBIRS`
2. **Web Service URL** → Advanced → Remove port 443 entries → Add port 8443 with your SSL cert
3. **Web Portal URL** → Advanced → Remove port 443 entries → Add port 8443 with your SSL cert
4. **URL ACL Reservation** (run as admin):
   ```powershell
   netsh http add urlacl url=https://+:8443/PowerBI/ sddl="D:(A;;GX;;;S-1-5-80-1730998386-2757299892-37364343-1607169425-3512908663)"
   ```
5. **SSL Certificate Binding**:
   ```powershell
   $cert = Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.Subject -match "YourCertName" }
   netsh http add sslcert ipport=0.0.0.0:8443 certhash=$($cert.Thumbprint) appid='{00000000-0000-0000-0000-000000000000}'
   ```

## Known Issues & Troubleshooting

### "An error has occurred" when viewing reports on PBIRS
**Cause**: Reports uploaded via REST API. Re-publish using this repo's SOAP-based script or PBI Desktop → Publish.

### Native query prompt blocks conversion
**Cause**: Registry key not set. The script sets this automatically, but you can also set it manually:
```powershell
$regPath = "HKCU:\SOFTWARE\Microsoft\Microsoft Power BI Desktop SSRS"
New-Item $regPath -Force | Out-Null
Set-ItemProperty $regPath -Name "AllowNativeQueries" -Value 1 -Type DWord
Set-ItemProperty $regPath -Name "EnableNativeQueryPrompt" -Value 0 -Type DWord
```

### SQL Server 2022 encrypted connection errors
**Cause**: SQL Server 2022 requires encrypted connections by default. Set the trusted servers environment variable:
```powershell
[Environment]::SetEnvironmentVariable("PBI_SQL_TRUSTED_SERVERS", "mecmdb.corp.local", "Machine")
```

### Reports fail to connect after publishing
**Cause**: Data sources need credential configuration. Use the `-ConfigureDataSource` and `-PromptForDataSourceCredential` flags, or configure manually in the PBIRS web portal under each report's data source settings.

## Included MECM Templates

These are the standard Microsoft MECM sample templates:

| Template | Description |
|----------|-------------|
| Microsoft Client Status | Client health, activity, and compliance |
| Microsoft Content Status | Content distribution and status |
| Microsoft Edge Management | Edge browser deployment and usage |
| Microsoft Software Update Compliance Status | Patch compliance overview |
| Microsoft Software Update Deployment Status | SUG deployment tracking |

## License

MIT

## Contributing

Issues and PRs welcome. This was born out of necessity in air-gapped DoD/SIPR environments where Power BI Service is not available.
