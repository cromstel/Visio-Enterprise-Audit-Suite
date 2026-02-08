# 🚀 Deployment Guide
## Visio Enterprise Audit Suite

---

## 📦 Deployment Overview

This guide walks you through deploying the Visio Enterprise Audit Suite on your network.

---

## ✅ Pre-Deployment Checklist

- [ ] Extract ZIP file to desired location
- [ ] User account has domain admin or delegated AD permissions
- [ ] PowerShell 5.1+ installed
- [ ] RSAT tools installed (Windows 11 only)
- [ ] Network access to domain computers
- [ ] WMI service running on target computers
- [ ] Output directory exists or can be created

---

## 🔧 Installation Steps

### **Step 1: Extract the Package**

```
C:\Scripts\
└── Visio-Enterprise-Audit-Suite/
    ├── README.md
    ├── VISIO_AUDIT_GUIDE.md
    ├── Visio-Enterprise-Audit.ps1
    ├── Visio-Usage-Analytics.ps1
    ├── Visio-Helper-Utils.ps1
    ├── DEPLOYMENT.md
    ├── TROUBLESHOOTING.md
    └── CHANGELOG.md
```

### **Step 2: Windows 11 Prerequisites (Skip if Windows Server)**

Run PowerShell as Administrator:

```powershell
# Install RSAT Active Directory tools
Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"

# Verify installation
Import-Module ActiveDirectory
Get-ADComputer -Filter * -ResultSize 1
```

### **Step 3: Set Execution Policy**

Run PowerShell as Administrator:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
```

Verify:
```powershell
Get-ExecutionPolicy
# Should return: RemoteSigned
```

### **Step 4: Test Domain Connectivity**

```powershell
# Verify domain membership
whoami /fqdn
# Should show: DOMAIN\username

# Test AD access
Get-ADComputer -Filter * -ResultSize 5
# Should list 5 domain computers
```

### **Step 5: Verify WMI/CIM Access**

```powershell
# Test CIM session to a computer
$session = New-CimSession -ComputerName "COMPUTER1"
Get-CimInstance -CimSession $session -ClassName Win32_Product | Select-Object -First 3
Remove-CimSession $session
```

If this fails, check:
- Computer is online
- Network firewall allows WMI (ports 135, 445)
- Your user has admin rights on target computer

---

## 🎯 Deployment Scenarios

### **Scenario 1: Single Audit (IT Admin Workstation)**

```powershell
# Navigate to script directory
cd "C:\Scripts\Visio-Enterprise-Audit-Suite"

# Run audit
.\Visio-Enterprise-Audit.ps1

# View results
Start-Process "C:\Temp\VisioAudit"
```

**Time Required:** 10-30 minutes (depending on domain size)

---

### **Scenario 2: Scheduled Weekly Audits (Windows Server)**

#### **Option A: Using Task Scheduler UI**

1. Open **Task Scheduler**
2. Right-click **Task Scheduler Library** → **Create Basic Task**
3. **Name:** "Weekly Visio Audit"
4. **Trigger:** Weekly, Sunday 2:00 AM
5. **Action:** Start a program
   - Program: `powershell.exe`
   - Arguments: `-NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\Visio-Enterprise-Audit-Suite\Visio-Enterprise-Audit.ps1"`
6. **Conditions:** Check "Wake the computer to run this task"
7. **Settings:** Check "Run with highest privileges"
8. Click **OK**

#### **Option B: Using PowerShell (Automated)**

Run as Administrator:

```powershell
$scriptPath = "C:\Scripts\Visio-Enterprise-Audit-Suite\Visio-Enterprise-Audit.ps1"
$taskName = "VisioAudit-Weekly"

$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At "02:00"

$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""

Register-ScheduledTask `
    -TaskName $taskName `
    -Trigger $trigger `
    -Action $action `
    -RunLevel Highest `
    -Description "Weekly Visio installation audit" `
    -Force

Write-Host "Task created: $taskName"
```

Verify:
```powershell
Get-ScheduledTask -TaskName "VisioAudit-Weekly" | Select-Object TaskName, State, LastTaskResult
```

---

### **Scenario 3: Department-Specific Scans**

```powershell
# Scan only SALES department computers
.\Visio-Enterprise-Audit.ps1 -ComputerFilter "SALES-*" -OutputPath "C:\Reports\Sales"

# Scan only DESIGN department
.\Visio-Enterprise-Audit.ps1 -ComputerFilter "DESIGN-*" -OutputPath "C:\Reports\Design"
```

---

### **Scenario 4: Large Enterprise (1000+ Computers)**

```powershell
# Use 20 threads for faster scanning
.\Visio-Enterprise-Audit.ps1 `
    -ThreadCount 20 `
    -OutputPath "C:\Reports\VisioAudit"

# Or split by department if timeout issues occur
.\Visio-Enterprise-Audit.ps1 -ComputerFilter "DEPT1-*" -ThreadCount 10
.\Visio-Enterprise-Audit.ps1 -ComputerFilter "DEPT2-*" -ThreadCount 10
```

---

## 📊 Network Requirements

### **Required Ports (Inbound to Target Computers)**

| Port | Protocol | Service | Purpose |
|------|----------|---------|---------|
| 135 | TCP/UDP | RPC | Remote Procedure Call |
| 445 | TCP | SMB | Network file sharing |
| 5985 | TCP | WinRM | Windows Remote Management |
| 5986 | TCP | WinRM SSL | WinRM encrypted |

### **Firewall Configuration**

If computers are behind strict firewall:

```powershell
# Allow WMI on scanning machine
netsh advfirewall firewall add rule name="WMI-Allow-In" `
    dir=in action=allow protocol=tcp localport=135
```

---

## 📁 Output & Reports

### **Default Output Location**
```
C:\Temp\VisioAudit\
├── VisioAudit_20240215_143022.csv
├── VisioAudit_20240215_143022.html
├── VisioUsageAnalytics_20240215_143022.html
└── [Previous reports...]
```

### **Custom Output Location**

```powershell
# Store on network share
.\Visio-Enterprise-Audit.ps1 -OutputPath "\\FileServer\VisioAudit\Reports"

# Or local secure directory
.\Visio-Enterprise-Audit.ps1 -OutputPath "D:\Audits\Visio"
```

### **Network Share Configuration**

```powershell
# Create shared folder with restricted access
New-Item -ItemType Directory -Path "D:\VisioAudit" -Force

$acl = Get-Acl "D:\VisioAudit"
$ace = New-Object System.Security.AccessControl.FileSystemAccessRule(
    "DOMAIN\IT-Group",
    "FullControl",
    "ContainerInherit,ObjectInherit",
    "None",
    "Allow"
)
$acl.SetAccessRule($ace)
Set-Acl -Path "D:\VisioAudit" -AclObject $acl

# Create SMB share
New-SmbShare -Name "VisioAudit" `
    -Path "D:\VisioAudit" `
    -FullAccess "DOMAIN\IT-Group"
```

---

## 🔄 Backup & Retention

### **Automatic Archiving Script**

```powershell
$reportPath = "C:\Temp\VisioAudit"
$archivePath = "C:\Temp\VisioAudit\Archive"
$daysToKeep = 90

# Create archive directory
if (!(Test-Path $archivePath)) {
    New-Item -ItemType Directory -Path $archivePath -Force | Out-Null
}

# Archive old reports
Get-ChildItem -Path $reportPath -Filter "VisioAudit_*.csv" |
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$daysToKeep) } |
    Move-Item -Destination $archivePath

Write-Host "Archived old reports to: $archivePath"
```

Add to scheduled task to run monthly.

---

## 🔐 Security Best Practices

### **1. Run with Minimal Privileges**

```powershell
# Use delegated AD permissions instead of full admin
# In Active Directory Users and Computers:
# 1. Create a custom role with only "Query Computers" permission
# 2. Assign to IT Audit group
# 3. Run script with that account
```

### **2. Secure Report Files**

```powershell
# Restrict access to reports
$reportPath = "C:\Temp\VisioAudit"
$acl = Get-Acl $reportPath

# Remove Everyone access
$aceToRemove = $acl.Access | Where-Object { $_.IdentityReference -like "*Everyone*" }
if ($aceToRemove) {
    $acl.RemoveAccessRule($aceToRemove) | Out-Null
}

# Add only IT admins
$ace = New-Object System.Security.AccessControl.FileSystemAccessRule(
    "DOMAIN\IT-Admins",
    "FullControl",
    "ContainerInherit,ObjectInherit",
    "None",
    "Allow"
)
$acl.SetAccessRule($ace)
Set-Acl -Path $reportPath -AclObject $acl
```

### **3. Enable Audit Logging**

```powershell
# Log script execution
$logPath = "C:\Logs\VisioAudit.log"

# Add to script execution:
# "$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) - Audit started" | Out-File -Append $logPath
```

### **4. Encrypt Sensitive Reports**

```powershell
# Optional: Encrypt report files
$reportFile = "C:\Temp\VisioAudit\VisioAudit_20240215_143022.csv"
Cipher /e /s:C:\Temp\VisioAudit
```

---

## 🧪 Testing Deployment

### **Test 1: Verify Prerequisites**

```powershell
# Create test script
$testScript = @"
Write-Host "✓ PowerShell: $(`$PSVersionTable.PSVersion)"

try {
    Import-Module ActiveDirectory
    Write-Host "✓ AD Module: Available"
}
catch {
    Write-Host "✗ AD Module: Not available"
}

try {
    Get-ADComputer -Filter * -ResultSize 1 | Out-Null
    Write-Host "✓ Domain Access: OK"
}
catch {
    Write-Host "✗ Domain Access: Failed"
}

try {
    `$session = New-CimSession -ComputerName "localhost"
    Write-Host "✓ WMI/CIM: OK"
    Remove-CimSession `$session
}
catch {
    Write-Host "✗ WMI/CIM: Failed"
}
"@

$testScript | Out-File "C:\test-prerequisites.ps1"
& "C:\test-prerequisites.ps1"
```

### **Test 2: Scan Sample Computers**

```powershell
# Test with just 5 computers
$computers = Get-ADComputer -Filter * -ResultSize 5
Write-Host "Testing with $($computers.Count) computers..."

# Run manual scan on first computer
$computer = $computers[0]
.\Visio-Enterprise-Audit.ps1 -ComputerFilter $computer.Name
```

### **Test 3: Verify Reports**

```powershell
# Check if reports were generated
Get-ChildItem "C:\Temp\VisioAudit\VisioAudit_*.csv" -Latest 1 | 
    Invoke-Item  # Opens in Excel

# Count records
Import-Csv "C:\Temp\VisioAudit\VisioAudit_*.csv" | Measure-Object
```

---

## 📈 Scaling Considerations

### **For 100-500 Computers**
- Use default settings (10 threads)
- Run once daily during off-hours
- Standard 8GB RAM machine sufficient

### **For 500-2000 Computers**
- Increase threads to 15-20
- Run on dedicated server
- 16GB RAM recommended
- Consider splitting by OU/department

### **For 2000+ Computers**
- Split scans by department
- Use multiple scanning machines
- Run in parallel (non-overlapping time windows)
- 32GB RAM recommended
- Monitor disk I/O during scans

---

## 📞 Post-Deployment Support

### **Monitor First Run**

```powershell
# Watch progress
Get-ChildItem "C:\Temp\VisioAudit\*.csv" -Recurse | 
    Sort-Object LastWriteTime -Descending | 
    Select-Object Name, LastWriteTime, Length
```

### **Troubleshoot Issues**

See `TROUBLESHOOTING.md` for common issues.

### **Regular Maintenance**

```powershell
# Monthly: Archive old reports
# Quarterly: Review and update computer filters
# Annually: Update documentation and procedures
```

---

## ✨ Next Steps

1. **Extract the ZIP file**
2. **Follow installation steps above**
3. **Run initial test scan**
4. **Review generated reports**
5. **Schedule automated audits** (optional)
6. **Set up email notifications** (optional)
7. **Configure backup/archival** (optional)

---

**Deployment complete!** For detailed usage, see README.md and VISIO_AUDIT_GUIDE.md
