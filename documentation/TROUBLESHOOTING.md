# 🆘 Troubleshooting Guide
## Visio Enterprise Audit Suite

---

## 🔴 Critical Issues

### **1. Execution Policy Error**

**Error Message:**
```
The file ... cannot be loaded. The file is not digitally signed.
```

**Cause:** PowerShell execution policy blocks unsigned scripts

**Solutions:**

**Option A: Bypass (Session Only)**
```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
.\Visio-Enterprise-Audit.ps1
```

**Option B: RemoteSigned (Permanent)**
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
.\Visio-Enterprise-Audit.ps1
```

**Option C: Command Line Bypass**
```powershell
powershell.exe -ExecutionPolicy Bypass -File ".\Visio-Enterprise-Audit.ps1"
```

**Option D: Unblock File**
- Right-click script → Properties
- Check "Unblock" at bottom
- Click OK
- Run script normally

---

### **2. Active Directory Module Not Found**

**Error Message:**
```
Import-Module : The specified module 'ActiveDirectory' was not loaded because no valid module file was found in any module directory.
```

**Cause:** RSAT tools not installed (especially Windows 11)

**Solution:**

**Windows 11:**
```powershell
# Install RSAT
Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"

# Verify
Import-Module ActiveDirectory
```

**Windows Server:**
```powershell
# Already installed, try reloading
Remove-Module ActiveDirectory
Import-Module ActiveDirectory
```

**If still failing:**
```powershell
# Restart PowerShell and try again
# Or check Windows features
Get-WindowsCapability -Online | Where-Object Name -like "Rsat*"
```

---

### **3. Domain Not Accessible**

**Error Message:**
```
Cannot bind argument to parameter 'Filter' because it is null.
Get-ADComputer : The server is not operational
```

**Cause:** Computer not domain-joined or domain unreachable

**Solutions:**

**Check Domain Membership:**
```powershell
# Should show DOMAIN\username (not just username)
whoami /fqdn

# If not domain-joined, it shows: COMPUTERNAME\username
# Solution: Join computer to domain
```

**Test Domain Connectivity:**
```powershell
# Verify you can reach domain
nslookup your-domain.com
ping your-domain.com

# Test AD connectivity
Get-ADComputer -Filter * -ResultSize 1
```

**VPN/Network Issues:**
```powershell
# If on VPN, ensure it's connected
# Check DNS settings point to domain DNS
# Restart network adapter if needed
```

---

### **4. WMI/CIM Session Failures**

**Error Message:**
```
New-CimSession : WS Management service cannot process the request
```

**Cause:** WMI service issues, firewall, or timeout

**Solutions:**

**Test WMI Service:**
```powershell
# Check if WMI service is running on target computer
Get-Service -Name winmgmt -ComputerName "COMPUTER-NAME"

# If stopped, start it
Start-Service -Name winmgmt -ComputerName "COMPUTER-NAME"
```

**Reduce Thread Count:**
```powershell
# Too many parallel sessions can cause timeouts
.\Visio-Enterprise-Audit.ps1 -ThreadCount 5
```

**Check Firewall:**
```powershell
# WMI uses ports 135, 445
# Check Windows Firewall
Get-NetFirewallRule -DisplayName "*WMI*" -Enabled True

# If disabled, enable
Set-NetFirewallRule -DisplayName "*Windows Management Instrumentation*" -Enabled True
```

**Network Connectivity:**
```powershell
# Test if target computer is reachable
Test-Connection -ComputerName "COMPUTER-NAME" -Count 3
```

---

## 🟠 Performance Issues

### **5. Script Running Slowly**

**Symptoms:** Scanning seems hung, not progressing

**Solutions:**

**Reduce Thread Count:**
```powershell
# Default is 10, reduce to 5 for slow networks
.\Visio-Enterprise-Audit.ps1 -ThreadCount 5
```

**Filter Computers:**
```powershell
# Scan only specific department
.\Visio-Enterprise-Audit.ps1 -ComputerFilter "DEPT-*"

# Or create manual list
$computers = @("WS-001", "WS-002", "WS-003")
# Then modify script to use this list
```

**Increase Timeout:**

Edit script, find this line:
```powershell
$session = New-CimSession -ComputerName $ComputerName -ErrorAction Stop
```

Change to:
```powershell
$timeout = New-CimSessionOption -OperationTimeoutSec 30
$session = New-CimSession -ComputerName $ComputerName -SessionOption $timeout -ErrorAction Stop
```

**Check System Resources:**
```powershell
# Monitor CPU and RAM usage
Get-Process powershell | Select-Object ProcessName, CPU, WorkingSet | Format-Table -AutoSize

# If maxed out, close other applications
```

---

### **6. Memory Issues / Out of Memory**

**Error Message:**
```
Out of memory
PSArgumentException: Cannot allocate memory
```

**Solutions:**

**Reduce Thread Count:**
```powershell
.\Visio-Enterprise-Audit.ps1 -ThreadCount 5
```

**Increase Available RAM:**
```powershell
# Close unnecessary applications
Stop-Process -Name "chrome", "firefox" -ErrorAction SilentlyContinue

# Or split scan into smaller batches
```

**Scan by Department:**
```powershell
# Instead of scanning all at once
.\Visio-Enterprise-Audit.ps1 -ComputerFilter "SALES-*"
.\Visio-Enterprise-Audit.ps1 -ComputerFilter "DESIGN-*"
# Then merge reports
```

---

## 🟡 Data & Reporting Issues

### **7. No Results or Empty CSV**

**Symptoms:** Script completes but CSV is empty or shows no Visio installations

**Causes:** 
- No computers found
- All computers offline
- Filter too restrictive
- No Visio installed

**Solutions:**

**Verify Computers Found:**
```powershell
# Check if computers exist
Get-ADComputer -Filter "Name -like '*'" | Measure-Object

# If 0, domain has no computers
# Check with admin who created computer objects
```

**Check Specific Filter:**
```powershell
# If using filter, verify it matches computers
Get-ADComputer -Filter "Name -like 'DEPT-*'" | Select-Object -First 10

# If no results, adjust filter
.\Visio-Enterprise-Audit.ps1 -ComputerFilter "*"
```

**Verify Computers Are Online:**
```powershell
# Manually check if computers are reachable
$computers = Get-ADComputer -Filter * | Select-Object -First 10
$computers | ForEach-Object {
    $online = Test-Connection -ComputerName $_.Name -Count 1 -Quiet
    Write-Host "$($_.Name): $(if($online){'Online'} else {'Offline'})"
}
```

**Check WMI Access:**
```powershell
# Test WMI on a specific computer
$computer = "COMPUTER-NAME"
Get-CimInstance -ComputerName $computer -ClassName Win32_Product | Measure-Object

# If fails, WMI might be blocked
```

---

### **8. HTML Report Not Generated**

**Error Message:** 
```
HTML report saved: ... [but file doesn't exist]
```

**Cause:** Permissions issue or CSV generation failed

**Solutions:**

**Check Output Directory:**
```powershell
# Verify directory exists and is writable
Test-Path "C:\Temp\VisioAudit"

# If doesn't exist, create it
New-Item -ItemType Directory -Path "C:\Temp\VisioAudit" -Force
```

**Check Permissions:**
```powershell
# Make sure you have write access
$acl = Get-Acl "C:\Temp\VisioAudit"
$acl.Access | Format-Table

# If needed, grant yourself access
$ace = New-Object System.Security.AccessControl.FileSystemAccessRule(
    "$env:USERDOMAIN\$env:USERNAME",
    "FullControl",
    "ContainerInherit,ObjectInherit",
    "None",
    "Allow"
)
$acl.SetAccessRule($ace)
Set-Acl -Path "C:\Temp\VisioAudit" -AclObject $acl
```

**Check CSV Generation:**
```powershell
# Verify CSV was actually created
Get-ChildItem "C:\Temp\VisioAudit\*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

# If no CSV, script failed earlier
```

---

### **9. Report Shows All Computers Offline**

**Symptoms:** CSV shows IsOnline = "No" for all computers

**Cause:** Network connectivity issue or Test-Connection blocked

**Solutions:**

**Test Connectivity Manually:**
```powershell
# Try pinging computers directly
Test-Connection -ComputerName "WS-001" -Count 1 -Verbose

# If fails, issue is network-related
```

**Check Firewall:**
```powershell
# Some firewalls block ping
# Try with WMI instead
Get-CimInstance -ComputerName "WS-001" -ClassName Win32_ComputerSystem

# If this works but ping doesn't, ping is just blocked (OK)
```

**Try WinRM:**
```powershell
# Some networks have WMI blocked but WinRM open
Test-WSMan -ComputerName "WS-001"

# If works, remove Test-Connection from script and rely on WMI
```

---

### **10. Last Used Date is NULL for All Entries**

**Symptom:** LastUsedDate column empty for all computers

**Cause:** File access times not available or disabled

**Solutions:**

**Check File Timestamp Access:**
```powershell
# Windows might not track file access times
# Check registry setting
Get-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\FileSystem" -Name NtfsDisableLastAccessUpdate

# 0 = enabled (recommended for performance)
# 1 = disabled (we need this for last access tracking)
```

**If Disabled, Enable It:**
```powershell
# WARNING: Enabling impacts performance
# Only on audit machine, not all computers
Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\FileSystem" `
    -Name "NtfsDisableLastAccessUpdate" -Value 0 -Force
```

**Alternative: Use Registry Data:**
```powershell
# Edit script to check:
# HKCU:\Software\Microsoft\Office\16.0\Common\Open Find
# This might have usage history
```

---

## 🟢 Verification Checklist

After troubleshooting, verify with this checklist:

```powershell
# 1. PowerShell Version
$PSVersionTable.PSVersion  # Should be 5.1+

# 2. Execution Policy
Get-ExecutionPolicy  # Should be RemoteSigned or Bypass

# 3. AD Module
Import-Module ActiveDirectory  # Should not error

# 4. Domain Access
Get-ADComputer -Filter * -ResultSize 1  # Should return computer

# 5. Network Access
Test-Connection -ComputerName "COMPUTER-NAME" -Count 1  # Should succeed

# 6. WMI Access
Get-CimInstance -ComputerName "COMPUTER-NAME" -ClassName Win32_ComputerSystem  # Should return data

# 7. Output Directory
Test-Path "C:\Temp\VisioAudit"  # Should be True
```

If all return successfully, you're ready to run!

---

## 📋 Common Error Messages & Solutions

| Error | Cause | Solution |
|-------|-------|----------|
| "File cannot be loaded" | Execution policy | `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned` |
| "ActiveDirectory module not found" | RSAT not installed | Install RSAT on Windows 11 |
| "Get-ADComputer: The server is not operational" | Domain unreachable | Check domain connectivity, VPN, DNS |
| "Access denied" | No admin privileges | Run PowerShell as Administrator |
| "New-CimSession: WS Management service cannot process" | WMI timeout/blocked | Reduce thread count, check firewall |
| "Cannot allocate memory" | Too many threads | Reduce ThreadCount parameter |
| "No CSV generated" | All computers offline | Check network connectivity |
| "Empty results" | Wrong filter | Verify computer naming convention |

---

## 🔧 Advanced Troubleshooting

### **Enable Debug Logging**

Add to script for detailed output:

```powershell
$DebugPreference = "Continue"
$VerbosePreference = "Continue"

# Then run
.\Visio-Enterprise-Audit.ps1 -Verbose -Debug
```

### **Test Individual Functions**

```powershell
# Source the script to load functions
. .\Visio-Enterprise-Audit.ps1

# Test on single computer
Get-VisioInstallationInfo -ComputerName "WS-001" -Paths $VisioPaths -RegPaths $RegistryPaths
```

### **Check Event Logs**

```powershell
# WMI errors appear in Event Viewer
Get-WinEvent -LogName System -FilterXPath "*[System[EventID=19]]" -MaxEvents 10
```

---

## 📞 When to Escalate

If you've tried all solutions above, contact:

- **Network team**: If WMI/SMB connectivity fails
- **Domain admin**: If AD permissions needed
- **System admins**: If firewall/WMI service issues
- **Office 365 team**: If license data incorrect

Provide:
- Error messages (exact text)
- PowerShell version (`$PSVersionTable`)
- Number of computers scanned
- Network topology diagram
- Firewall rules in place

---

## 📚 Additional Resources

- `README.md` - Quick start guide
- `VISIO_AUDIT_GUIDE.md` - Complete documentation
- `DEPLOYMENT.md` - Deployment procedures
- PowerShell Help: `Get-Help Get-ADComputer`
- Microsoft Docs: about_Execution_Policies

---

**Still stuck?** Check the main documentation or review the error message carefully for clues!
