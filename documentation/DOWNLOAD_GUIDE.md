# 📥 Download & Installation Guide
## Visio Enterprise Audit Suite v1.0

---

## 📦 Package Information

**File Name:** `Visio-Enterprise-Audit-Suite.zip`  
**File Size:** 39 KB (compressed)  
**Uncompressed Size:** 118 KB  
**Download Time:** <1 second on modern connections  
**Checksum:** Available upon request  

---

## 🔽 Download Instructions

### **Option 1: Direct Download**
1. Click the download link above
2. Save to desired location (e.g., E:\automation package\)
3. Proceed to extraction

### **Option 2: Browser Download**
1. Right-click download link
2. Select "Save As..."
3. Choose destination folder
4. Click Save

### **Option 3: Command Line (PowerShell)**
```powershell
# If you have a direct URL
Invoke-WebRequest -Uri "https://github.com/cromstel/visio-enterprise-audit-suite.zip" `
    -OutFile "$env:USERPROFILE\Downloads\visio-enterprise-audit-suite.zip"
```

---

## 📂 Extraction Instructions

### **Windows Explorer (Easiest)**

1. **Locate the ZIP file**
   - Find `visio-enterprise-audit-suite.zip` in Downloads or saved location

2. **Right-Click the ZIP**
   - Select "Extract All..."

3. **Choose Destination**
   - Recommended: `C:\Scripts\` or `E:\automation package\`
   - Click "Extract"

4. **Folder Created**
   - `visio-enterprise-audit-suite/` folder appears with 10 files

### **PowerShell Method**

```powershell
# Navigate to download location
cd $env:USERPROFILE\Downloads

# Extract ZIP
Expand-Archive -Path "Visio-Enterprise-Audit-Suite.zip" -DestinationPath "C:\Scripts\"

# Verify extraction
Get-ChildItem "C:\Scripts\visio-enterprise-audit-suite"
```

### **Command Prompt Method**

```cmd
# Navigate to download folder
cd %USERPROFILE%\Downloads

# Extract using built-in tar (Windows 10+)
tar -xf visio-enterprise-audit-suite.zip -C C:\Scripts

# Or use 7-Zip if installed
7z x visio-enterprise-audit-suite.zip -oC:\Scripts
```

---

## ✅ Verify Extraction

After extraction, you should see this structure:

```
visio-enterprise-audit-suite/
├── README.md
├── PACKAGE_INFO.txt
├── VISIO_AUDIT_GUIDE.md
├── DEPLOYMENT.md
├── TROUBLESHOOTING.md
├── CHANGELOG.md
├── Visio-Enterprise-Audit.ps1
├── Visio-Usage-Analytics.ps1
└── Visio-Helper-Utils.ps1
```

**Verify in PowerShell:**
```powershell
cd "C:\Scripts\Visio-Enterprise-Audit-Suite"
Get-ChildItem | Select-Object Name, Length

# Should show 10 items with sizes:
# README.md                 ~8 KB
# PACKAGE_INFO.txt          ~10 KB
# VISIO_AUDIT_GUIDE.md      ~12 KB
# DEPLOYMENT.md             ~11 KB
# TROUBLESHOOTING.md        ~12 KB
# CHANGELOG.md              ~8 KB
# Visio-Enterprise-Audit.ps1   ~24 KB
# Visio-Usage-Analytics.ps1    ~12 KB
# Visio-Helper-Utils.ps1       ~20 KB
```

---

## 🚀 Quick Start (After Extraction)

### **Step 1: Open PowerShell as Administrator**

- **Windows 11:** 
  - Press `Win + X`
  - Select "Windows Terminal (Admin)" or "PowerShell (Admin)"

- **Windows Server:**
  - Start Menu → Windows PowerShell → Right-click → Run as Administrator

### **Step 2: Navigate to Folder**

```powershell
cd "C:\Scripts\Visio-Enterprise-Audit-Suite"
```

### **Step 3: Read Quick Start**

```powershell
# Open README in Notepad
notepad README.md
```

### **Step 4: Set Execution Policy (If Needed)**

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
```

### **Step 5: (Windows 11 Only) Install RSAT**

```powershell
Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"
```

### **Step 6: Run the Audit**

```powershell
.\Visio-Enterprise-Audit.ps1
```

### **Step 7: View Results**

Reports automatically saved to: `C:\Temp\VisioAudit\`

Open HTML report in browser:
```powershell
Start-Process "C:\Temp\VisioAudit"
```

---

## 🔧 System Requirements Check

Before running, verify you have everything:

```powershell
# Check PowerShell version
$PSVersionTable.PSVersion
# Should be 5.1 or higher

# Check if domain-joined
[System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
# Should return your domain name

# Check AD module availability (Windows 11)
Get-Module -ListAvailable ActiveDirectory
# Should list ActiveDirectory module

# Test AD access
Get-ADComputer -Filter * -ResultSize 1
# Should return a computer object
```

If any of these fail, see TROUBLESHOOTING.md after extraction.

---

## 📖 Documentation Guide

**Start here based on your needs:**

### **I just want to run a scan**
→ Read: **README.md** (5 minutes)

### **I want detailed information**
→ Read: **VISIO_AUDIT_GUIDE.md** (20 minutes)

### **I'm setting up automation**
→ Read: **DEPLOYMENT.md** (15 minutes)

### **Something went wrong**
→ Read: **TROUBLESHOOTING.md** (varies)

### **What's new in this version**
→ Read: **CHANGELOG.md** (5 minutes)

### **Package contents overview**
→ Read: **PACKAGE_INFO.txt** (5 minutes)

---

## 🛡️ Safety Checks

**Before running the scripts:**

- [ ] ZIP file extracted successfully
- [ ] All 10 files present in folder
- [ ] Running as Administrator
- [ ] PowerShell 5.1+ installed
- [ ] Domain membership verified
- [ ] Network connectivity confirmed
- [ ] Read README.md

---

## 💾 Storage Location Options

### **Recommended Locations**

**Home Users/Labs:**
```
C:\Scripts\Visio-Enterprise-Audit-Suite
```

**Small Business:**
```
C:\Audits\Visio\
D:\Tools\VisioAudit\
```

**Enterprise:**
```
\\FileServer\IT-Tools\VisioAudit\
C:\Program Files (x86)\VisioAudit\
```

---

## 🔐 Antivirus / Security Considerations

**The scripts contain no malware, but:**

1. **Antivirus May Flag Scripts**
   - PowerShell scripts are sometimes flagged
   - Safe to whitelist in your antivirus
   - Common false positives only

2. **Execution Policy**
   - Set to RemoteSigned for production
   - Allows unsigned local scripts

3. **Network Scanning**
   - Scripts query network computers
   - May trigger IDS/IPS alerts
   - Contact security team beforehand

4. **Data Access**
   - Scripts need admin permissions
   - Can be run with delegated permissions
   - See DEPLOYMENT.md for security setup

---

## 🆘 Common Download Issues

### **Issue: File shows as .zip.zip**
**Solution:** Check file extension, rename if needed:
```powershell
Rename-Item "Visio-Enterprise-Audit-Suite.zip.zip" "Visio-Enterprise-Audit-Suite.zip"
```

### **Issue: Extraction fails with error**
**Solution:** Try different extraction method:
```powershell
# Method 1: PowerShell
Expand-Archive -Path "path\Visio-Enterprise-Audit-Suite.zip" -DestinationPath "C:\Scripts"

# Method 2: Windows Explorer context menu
# Right-click → Extract All

# Method 3: 7-Zip (if installed)
7z x "Visio-Enterprise-Audit-Suite.zip" -oC:\Scripts
```

### **Issue: Missing files after extraction**
**Solution:** Verify ZIP integrity:
```powershell
# Check contents before extraction
unzip -l "Visio-Enterprise-Audit-Suite.zip"

# If missing files, redownload
```

### **Issue: Slow download**
**Solution:** Download might be resumable:
```powershell
# Check file was fully downloaded
(Get-Item "Visio-Enterprise-Audit-Suite.zip").Length
# Should be ~39 KB (39936 bytes)
```

---

## 📱 Multi-Device Setup

### **Scenario: Running from Network Share**

```powershell
# Copy to network location
Copy-Item "Visio-Enterprise-Audit-Suite" "\\FileServer\Scripts\"

# Run from network
cd "\\FileServer\Scripts\Visio-Enterprise-Audit-Suite"
.\Visio-Enterprise-Audit.ps1
```

### **Scenario: Deploying to Multiple Machines**

```powershell
# Create deployment package
Compress-Archive -Path "Visio-Enterprise-Audit-Suite" -DestinationPath "Visio-Enterprise-Audit-Suite.zip"

# Copy to multiple machines
$machines = "PC1", "PC2", "PC3"
foreach ($machine in $machines) {
    Copy-Item "Visio-Enterprise-Audit-Suite" "\\$machine\C$\Scripts\"
}
```

---

## ✨ Post-Installation Steps

### **1. Initial Testing (10 minutes)**

```powershell
# Navigate to folder
cd "C:\Scripts\Visio-Enterprise-Audit-Suite"

# Test with minimal parameters
.\Visio-Enterprise-Audit.ps1 -ComputerFilter "WS-*" | Select-Object -First 10
```

### **2. Configuration (Optional)**

```powershell
# Set custom output directory
mkdir "C:\Reports\VisioAudit"

# Test with custom output
.\Visio-Enterprise-Audit.ps1 -OutputPath "C:\Reports\VisioAudit"
```

### **3. Scheduling (Optional)**

```powershell
# Create weekly automated scan
.\Visio-Helper-Utils.ps1
# Select option 9: Create Scheduled Task
```

### **4. First Full Scan**

```powershell
# Run complete audit
.\Visio-Enterprise-Audit.ps1

# View results
Start-Process "C:\Temp\VisioAudit"
```

---

## 🎯 What's Next?

1. ✅ **Download** the ZIP file
2. ✅ **Extract** to your desired location
3. ✅ **Read** README.md
4. ✅ **Run** Visio-Enterprise-Audit.ps1
5. ✅ **View** reports in C:\Temp\VisioAudit\
6. ✅ **Explore** other scripts and utilities
7. ✅ **Reference** documentation as needed

---

## 📞 Support Resources

After extraction, all documentation is included:

| Document | Purpose |
|----------|---------|
| README.md | Quick start (read first!) |
| VISIO_AUDIT_GUIDE.md | Complete reference guide |
| DEPLOYMENT.md | Setup & deployment |
| TROUBLESHOOTING.md | Problem solving |
| CHANGELOG.md | Version history |
| PACKAGE_INFO.txt | Package details |

---

## ✅ Validation Checklist

After downloading and extracting:

- [ ] ZIP file downloaded successfully (~39 KB)
- [ ] File extracted without errors
- [ ] All 10 files present in folder
- [ ] All PS1 scripts are readable
- [ ] All MD files are readable
- [ ] Running PowerShell as Administrator
- [ ] Ready to proceed to setup

**If all checked, you're ready to begin auditing!** 🚀

---

**Version:** 1.0  
**Last Updated:** February 8, 2024  
**Status:** ✅ Ready for Download
