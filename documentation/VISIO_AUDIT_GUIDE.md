# Enterprise Visio Installation & Usage Audit Scripts
## Complete Implementation Guide

---

## Overview

These PowerShell scripts provide comprehensive Visio installation tracking and usage analytics across your enterprise domain, with support for Office 365, 2019, 2016, and 2013.

### Scripts Included

1. **Visio-Enterprise-Audit.ps1** - Main installation scanning script
2. **Visio-Usage-Analytics.ps1** - Detailed usage tracking and analytics

---

## Prerequisites

### System Requirements

- **Windows Server 2012 R2** or later
- **PowerShell 4.0+** (5.1+ recommended)
- **Active Directory Module** for PowerShell installed
- **Administrator privileges** on the scanning machine
- **Network connectivity** to domain computers

### Installation Steps

#### 1. Install Active Directory Module (if not present)

```powershell
# On Windows Server
Import-Module ActiveDirectory

# On Windows 10/11 (if needed)
Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"
```

#### 2. Verify PowerShell Version

```powershell
$PSVersionTable.PSVersion
# Should be 5.1 or higher
```

#### 3. Enable Script Execution Policy

```powershell
# Set execution policy for current user
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force

# Or for all users (requires admin)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine -Force
```

---

## Quick Start Guide

### Basic Usage - Default Scan

```powershell
# Run the audit script
.\Visio-Enterprise-Audit.ps1

# This will:
# - Query all domain computers
# - Scan each computer for Visio installations
# - Generate CSV and HTML reports
# - Save to C:\Temp\VisioAudit\
```

### Advanced Usage - Custom Parameters

```powershell
# Specify custom output directory
.\Visio-Enterprise-Audit.ps1 -OutputPath "C:\Reports\VisioAudit"

# Scan only computers matching a filter (e.g., starting with "WS-")
.\Visio-Enterprise-Audit.ps1 -ComputerFilter "WS-*"

# Increase parallel threads for faster scanning (use with caution on network load)
.\Visio-Enterprise-Audit.ps1 -ThreadCount 20

# Combine parameters
.\Visio-Enterprise-Audit.ps1 `
    -OutputPath "C:\Reports\VisioAudit" `
    -ComputerFilter "DEPT-*" `
    -ThreadCount 15
```

### Run Usage Analytics

```powershell
# Analyze specific computers
.\Visio-Usage-Analytics.ps1 -ComputerNames @("COMPUTER1", "COMPUTER2", "COMPUTER3")

# Or scan all computers with Visio installed (from CSV report)
$visioComputers = Import-Csv "C:\Temp\VisioAudit\VisioAudit_20240115_143022.csv" | 
    Where-Object { $_.VisioInstalled -eq 'Yes' } | 
    Select-Object -ExpandProperty ComputerName

.\Visio-Usage-Analytics.ps1 -ComputerNames $visioComputers
```

---

## Output & Reports

### CSV Report Format

```
ComputerName,IsOnline,VisioInstalled,VisioVersion,Office365,LastUsedDate,InstallPath,Error
WS-001,Yes,Yes,16.0.14931.20354,Yes,2024-01-15 14:30:22,C:\Program Files\Microsoft Office\root\Office16\VISIO.EXE,None
WS-002,Yes,No,N/A,No,N/A,N/A,None
WS-003,No,Unknown,N/A,N/A,N/A,N/A,Computer offline
```

### HTML Report Features

- **Dashboard Metrics**: Total scans, installations, offline computers, install rate
- **Installation Summary**: Detailed table of all Visio installations with versions
- **Usage Data**: Last accessed dates and installation paths
- **Office 365 Detection**: Identifies cloud-based vs. desktop installations
- **Error Tracking**: Documents any scanning failures or access issues
- **Responsive Design**: Mobile-friendly report viewer

### Report Locations

Default: `C:\Temp\VisioAudit\`

Files generated:
- `VisioAudit_YYYYMMDD_HHMMSS.csv`
- `VisioAudit_YYYYMMDD_HHMMSS.html`
- `VisioUsageAnalytics_YYYYMMDD_HHMMSS.html`

---

## What the Scripts Check

### Visio-Enterprise-Audit.ps1

#### Installation Detection
- ✓ Office 365 installations
- ✓ Office 2019/2016 installations
- ✓ Office 2013 installations
- ✓ 32-bit and 64-bit variants
- ✓ Standard and custom installation paths

#### Usage Information
- ✓ Last file access time (from VISIO.EXE)
- ✓ Registry-based last open tracking
- ✓ Recent documents registry entries
- ✓ Installation date and version

#### Computer Status
- ✓ Online/offline status
- ✓ Operating system detection
- ✓ Network connectivity verification
- ✓ WMI/Registry access status

### Visio-Usage-Analytics.ps1

#### Active Usage
- ✓ Currently running Visio processes
- ✓ Active user information
- ✓ Multi-user detection

#### Document Analysis
- ✓ Recent .vsd and .vsdx files
- ✓ File modification dates
- ✓ File size tracking
- ✓ Inactivity periods

#### Configuration
- ✓ Auto-recovery settings
- ✓ Installed add-ins
- ✓ Default file format preferences
- ✓ License status (Office 365)

---

## Scheduling Automated Scans

### Create Scheduled Task (Windows Task Scheduler)

```powershell
# Define task parameters
$taskName = "VisioEnterpriseAudit"
$scriptPath = "C:\Scripts\Visio-Enterprise-Audit.ps1"
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At 2am

# Create action
$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""

# Create task
Register-ScheduledTask `
    -TaskName $taskName `
    -Trigger $trigger `
    -Action $action `
    -RunLevel Highest `
    -Description "Weekly Visio installation audit"
```

### Weekly Scan Template

```powershell
# Schedule for every Sunday at 2:00 AM
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At "02:00"
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument '-File "C:\Scripts\Visio-Enterprise-Audit.ps1"'
Register-ScheduledTask -TaskName "VisioAudit-Weekly" -Trigger $trigger -Action $action -RunLevel Highest
```

---

## Troubleshooting

### Issue: "Access Denied" Errors

**Solution**: Run PowerShell as Administrator
```powershell
# Verify admin privileges
[System.Security.Principal.WindowsIdentity]::GetCurrent().Owner
# Should show S-1-5-21-xxx-xxx-xxx-512 (Administrators)
```

### Issue: "The Active Directory Module is not loaded"

**Solution**: Import the module
```powershell
Import-Module ActiveDirectory
```

### Issue: Slow Scanning Performance

**Solution**: Reduce thread count or scan smaller groups
```powershell
# Scan 5 computers at a time instead of 10
.\Visio-Enterprise-Audit.ps1 -ThreadCount 5

# Filter to specific department
.\Visio-Enterprise-Audit.ps1 -ComputerFilter "SALES-*"
```

### Issue: WMI/Registry Access Timeouts

**Solution**: Increase timeout or manually target specific computers
```powershell
# Create a text file with computer names (one per line)
$computers = Get-Content "C:\computers.txt"
$computers | ForEach-Object {
    Write-Host "Scanning $_"
    Get-VisioInstallationInfo -ComputerName $_
}
```

### Issue: "Computer offline" for many devices

**Reason**: Devices may legitimately be off. Filter results:
```powershell
# Load CSV and show only online computers with Visio
$report = Import-Csv "C:\Temp\VisioAudit\VisioAudit_20240115_143022.csv"
$report | Where-Object { $_.IsOnline -eq "Yes" -and $_.VisioInstalled -eq "Yes" }
```

---

## Best Practices

### 1. Scan Frequency
- **Weekly scans** for tracking installations and changes
- **Monthly scans** for compliance auditing
- **Ad-hoc scans** after major Office updates

### 2. Network Considerations
- Run during off-hours to minimize network impact
- Use `ThreadCount 10-15` for enterprise (1000+ computers)
- Reduce threads on congested networks

### 3. Report Management
- Archive reports for compliance/audit trails
- Compare consecutive reports to track changes
- Export to SIEM for monitoring

### 4. Security
- Run script with minimal required privileges
- Secure report files (they contain inventory data)
- Use network shares with restricted access
- Enable audit logging on script execution

### 5. Error Handling
- Review "Error" column in CSV reports
- Verify network access to failed computers
- Check WMI service status on problematic machines
- Whitelist corporate firewalls for PowerShell remoting

---

## Advanced Scenarios

### Scenario 1: Find All Unused Visio Installations

```powershell
$cutoffDate = (Get-Date).AddMonths(-6)

Import-Csv "C:\Temp\VisioAudit\VisioAudit_20240115_143022.csv" |
    Where-Object { 
        $_.VisioInstalled -eq "Yes" -and 
        [datetime]$_.LastUsedDate -lt $cutoffDate 
    } |
    Select-Object ComputerName, LastUsedDate
```

### Scenario 2: Generate Cost Analysis

```powershell
$report = Import-Csv "C:\Temp\VisioAudit\VisioAudit_20240115_143022.csv"

$analysis = @{
    TotalVisioInstalls = ($report | Where-Object { $_.VisioInstalled -eq "Yes" }).Count
    Office365Licenses  = ($report | Where-Object { $_.Office365 -eq "Yes" }).Count
    DesktopVersions    = ($report | Where-Object { $_.Office365 -eq "No" -and $_.VisioInstalled -eq "Yes" }).Count
}

Write-Host "Visio Installation Summary:"
Write-Host "Total Visio Installations: $($analysis.TotalVisioInstalls)"
Write-Host "Office 365 Subscriptions: $($analysis.Office365Licenses)"
Write-Host "Desktop Licenses Needed: $($analysis.DesktopVersions)"
```

### Scenario 3: Monitor Specific Department

```powershell
# Filter for specific department and export
.\Visio-Enterprise-Audit.ps1 -ComputerFilter "DESIGN-*" -OutputPath "C:\Reports\Design-Dept"

# Then analyze
$deptReport = Import-Csv "C:\Reports\Design-Dept\VisioAudit_*.csv" | Sort-Object -Property LastUsedDate -Descending
$deptReport | Out-GridView
```

---

## Integration with Other Tools

### Export to Excel

```powershell
$report = Import-Csv "C:\Temp\VisioAudit\VisioAudit_20240115_143022.csv"

# Install ImportExcel module if needed
# Install-Module ImportExcel -Scope CurrentUser

$report | Export-Excel -Path "C:\Reports\VisioAudit.xlsx" `
    -WorksheetName "Installations" `
    -AutoFilter `
    -FreezeTopRow `
    -TableStyle Light10
```

### Send Email Report

```powershell
$emailParams = @{
    To          = "it-admin@company.com"
    From        = "visio-audit@company.com"
    Subject     = "Weekly Visio Audit Report"
    Body        = Get-Content "C:\Temp\VisioAudit\VisioAudit_20240115_143022.html" -Raw
    BodyAsHtml  = $true
    SmtpServer  = "smtp.company.com"
}

Send-MailMessage @emailParams -Attachments "C:\Temp\VisioAudit\VisioAudit_20240115_143022.csv"
```

---

## Performance Benchmarks

| Scenario | Computers | Thread Count | Time |
|----------|-----------|--------------|------|
| Small Business | 50 | 5 | ~5-10 min |
| Medium Enterprise | 200 | 10 | ~15-25 min |
| Large Enterprise | 500 | 15 | ~30-45 min |
| Very Large | 1000+ | 20 | ~60-90 min |

---

## Support & Troubleshooting Resources

### Common Errors

1. **"The term 'Get-ADComputer' is not recognized"**
   - Solution: `Import-Module ActiveDirectory`

2. **"Access denied" on remote registries**
   - Solution: Ensure script runs as Administrator with domain credentials

3. **"CimSession timeout"**
   - Solution: Reduce ThreadCount or add WinRM firewall rules

4. **HTML report shows no data**
   - Solution: Verify CSV was generated successfully, check output path

---

## Version History

- **v1.0** (2024-01-15): Initial release with Office 365, 2019, 2016, 2013 support
- **v1.1** (2024-01-20): Added detailed usage analytics
- **v1.2** (2024-02-01): Enhanced HTML reporting and parallel processing

---

## License & Support

These scripts are provided for enterprise IT administration purposes.
For issues or enhancements, document:
- PowerShell version
- Active Directory structure
- Number of computers scanned
- Error messages and stack traces

---

## Quick Reference Cheat Sheet

```powershell
# Basic scan
.\Visio-Enterprise-Audit.ps1

# Fast scan with custom output
.\Visio-Enterprise-Audit.ps1 -OutputPath "C:\Reports" -ThreadCount 20

# Filter department
.\Visio-Enterprise-Audit.ps1 -ComputerFilter "DEPT-*"

# Analyze usage
.\Visio-Usage-Analytics.ps1 -ComputerNames @("PC1", "PC2")

# View results
Import-Csv "C:\Temp\VisioAudit\VisioAudit_*.csv" | Format-Table

# Find Visio computers
Import-Csv "C:\Temp\VisioAudit\VisioAudit_*.csv" | Where-Object { $_.VisioInstalled -eq "Yes" }
```

---

Generated: 2024-01-15
Last Updated: 2024-02-01
