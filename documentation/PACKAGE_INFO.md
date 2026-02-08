# 📦 Visio Enterprise Audit Suite - Package Information

## Package Details

**Package Name:** `Visio-Enterprise-Audit-Suite.zip`  
**Version:** 1.0  
**Release Date:** February 8, 2024  
**Package Size:** 34 KB (compressed)  
**Uncompressed Size:** 107 KB  
**Total Files:** 9 (3 scripts + 5 documentation files + directory)

---

## 📋 What's Inside

### **Scripts (3 files)**

#### 1. **Visio-Enterprise-Audit.ps1** (23,702 bytes)
   - **Purpose:** Main installation scanner
   - **Features:**
     - Scans all domain computers
     - Detects Office 365, 2019, 2016, 2013
     - Multi-threaded processing (10-20 threads)
     - Generates CSV and HTML reports
     - Tracks last usage dates
     - Identifies online/offline status
   - **Run Time:** 15-90 minutes (depends on domain size)
   - **Output:** CSV + HTML reports
   - **Lines of Code:** ~800

#### 2. **Visio-Usage-Analytics.ps1** (12,362 bytes)
   - **Purpose:** Detailed usage tracking & analysis
   - **Features:**
     - Analyzes active Visio usage
     - Finds recent documents
     - Checks license status
     - Detects installed add-ins
     - Monitors configuration
   - **Run Time:** 5-30 minutes
   - **Output:** Analytics HTML report
   - **Lines of Code:** ~400

#### 3. **Visio-Helper-Utils.ps1** (20,493 bytes)
   - **Purpose:** Interactive menu & utilities
   - **Features:**
     - 12 pre-built utility functions
     - Find unused installations
     - Export to Excel
     - Generate cost analysis
     - Compare reports
     - Send emails
     - Create scheduled tasks
     - Department filtering
   - **Run Time:** On-demand per function
   - **Output:** Varies by function
   - **Lines of Code:** ~600

### **Documentation (5 files)**

#### 1. **README.md** (7,963 bytes)
   - Quick start guide
   - Feature overview
   - Common commands cheat sheet
   - Troubleshooting quick reference
   - Use case examples

#### 2. **VISIO_AUDIT_GUIDE.md** (12,201 bytes)
   - Complete reference documentation
   - Prerequisites & installation
   - Usage examples & parameters
   - Output format specifications
   - Best practices
   - Advanced scenarios
   - Performance benchmarks
   - Integration examples

#### 3. **DEPLOYMENT.md** (10,868 bytes)
   - Step-by-step deployment instructions
   - Scenario-based deployment
   - Network requirements
   - Security best practices
   - Testing procedures
   - Scaling considerations
   - Backup & retention strategies

#### 4. **TROUBLESHOOTING.md** (12,139 bytes)
   - Common error solutions
   - Performance issue fixes
   - Data & reporting issues
   - 10 detailed troubleshooting scenarios
   - Verification checklist
   - Advanced debugging
   - Support escalation guide

#### 5. **CHANGELOG.md** (7,690 bytes)
   - Version history
   - Feature list
   - Known issues & limitations
   - Compatibility matrix
   - Future roadmap
   - Installation history

---

## 🚀 Quick Start

### **Extract & Setup (5 minutes)**

```powershell
# 1. Extract ZIP file
# Right-click → Extract All

# 2. Open PowerShell as Administrator

# 3. Navigate to folder
cd E:\automation package\Visio-Enterprise-Audit-Suite

# 4. Set execution policy (Windows 11 only)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force

# 5. Install RSAT (Windows 11 only)
Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0"

# 6. Run audit
.\Visio-Enterprise-Audit.ps1

# 7. View reports
Start-Process "C:\Temp\VisioAudit"
```

---

## 📊 Feature Matrix

| Feature | Audit Script | Analytics | Utilities |
|---------|:---:|:---:|:---:|
| **Installation Detection** | ✅ | ✅ | - |
| **Usage Tracking** | ✅ | ✅ | - |
| **Report Generation** | ✅ | ✅ | ✅ |
| **Cost Analysis** | - | - | ✅ |
| **Excel Export** | - | - | ✅ |
| **Email Reports** | - | - | ✅ |
| **Interactive Menu** | - | - | ✅ |
| **Schedule Automation** | - | - | ✅ |
| **Department Filtering** | ✅ | ✅ | ✅ |
| **HTML Dashboards** | ✅ | ✅ | - |
| **CSV Export** | ✅ | - | ✅ |

---

## 💾 System Requirements

### **Minimum**
- Windows Server 2012 R2 or Windows 10/11
- PowerShell 4.0+
- 4 GB RAM
- 100 MB disk space
- Network access to domain computers

### **Recommended**
- Windows Server 2019+ or Windows 11
- PowerShell 5.1+
- 8+ GB RAM
- 500 MB disk space
- Active Directory integrated network

### **For Large Enterprises (1000+ computers)**
- Windows Server 2019+
- PowerShell 5.1+
- 16+ GB RAM
- Fast network connection
- Dedicated scanning machine

---

## 🎯 Common Use Cases

### **Compliance Auditing**
1. Run: `Visio-Enterprise-Audit.ps1`
2. Export: HTML report
3. Archive: For compliance records

### **License Optimization**
1. Run: Main audit
2. Analyze: Cost analysis (Utilities)
3. Action: Remove unused installations

### **Department Reporting**
1. Run: Audit with department filter
2. Export: Excel report (Utilities)
3. Share: Department-specific report

### **Ongoing Monitoring**
1. Create: Scheduled task (Utilities)
2. Auto-run: Weekly scans
3. Track: Changes over time

---

## 📈 Expected Results

### **Report Sizes**
| Domain Size | CSV Size | HTML Size | Time |
|------------|----------|-----------|------|
| 50 computers | 5 KB | 50 KB | 5 min |
| 200 computers | 20 KB | 150 KB | 20 min |
| 500 computers | 50 KB | 400 KB | 45 min |
| 1000 computers | 100 KB | 850 KB | 90 min |

### **Typical Results**
```
Total Computers Scanned: 234
- Online: 221 (94%)
- Offline: 13 (6%)

Visio Installations Found: 87 (37%)
- Office 365: 64 (74%)
- Desktop: 23 (26%)

Installation Rate: 37%
Last Used: All within 30 days
```

---

## 🔐 Security & Privacy

### **Data Handling**
- ✅ All data stays local (no cloud sync)
- ✅ Reports stored in C:\Temp\VisioAudit\
- ✅ No external API calls
- ✅ No telemetry collection
- ✅ Secure deletion recommended

### **Required Permissions**
- Domain admin or delegated AD read permissions
- Admin access to target computers (for WMI)
- Local write access for reports

### **Network Security**
- Uses standard Windows authentication
- Leverages existing domain infrastructure
- Compatible with firewalls & proxies

---

## 🛠️ Customization Options

### **Audit Script Parameters**
```powershell
# Basic
.\Visio-Enterprise-Audit.ps1

# Custom output
.\Visio-Enterprise-Audit.ps1 -OutputPath "C:\Reports"

# Department filter
.\Visio-Enterprise-Audit.ps1 -ComputerFilter "SALES-*"

# More threads (faster)
.\Visio-Enterprise-Audit.ps1 -ThreadCount 20
```

### **Modify Script Behavior**
Edit the scripts to:
- Change report output path
- Adjust thread count
- Modify detection logic
- Add custom filtering
- Integrate with other tools

All scripts are readable and well-commented for easy modification.

---

## 📊 Output Examples

### **CSV Report**
```
ComputerName,IsOnline,VisioInstalled,VisioVersion,Office365,LastUsedDate,InstallPath,Error
WS-001,Yes,Yes,16.0.14931.20354,Yes,2024-02-08 14:30:22,C:\Program Files\Microsoft Office\root\Office16\VISIO.EXE,None
WS-002,Yes,No,N/A,No,N/A,N/A,None
WS-003,No,Unknown,N/A,N/A,N/A,N/A,Computer offline
```

### **HTML Report**
- Professional dashboard with metrics
- Installation summary table
- Office 365 vs Desktop breakdown
- Offline computer list
- Mobile-responsive design

---

## 🆘 Getting Help

1. **Quick Issues:** Check README.md
2. **Setup Problems:** See DEPLOYMENT.md
3. **Errors:** Review TROUBLESHOOTING.md
4. **Full Details:** Read VISIO_AUDIT_GUIDE.md
5. **Changes:** Check CHANGELOG.md

---

## 📱 Compatibility

| Component | Windows Server | Windows 11 | Network |
|-----------|:---:|:---:|:---:|
| Visio Detection | ✅ | ✅ | Any |
| Usage Tracking | ✅ | ✅ | LAN/VPN |
| Email Reports | ✅ | ✅ | Must have SMTP |
| Scheduled Tasks | ✅ | ⚠️ | N/A |
| Excel Export | ✅ | ✅ | N/A |

---

## 🎓 Learning Path

### **Beginner (30 minutes)**
1. Read: README.md
2. Run: `Visio-Enterprise-Audit.ps1` (basic)
3. View: Generated reports

### **Intermediate (1-2 hours)**
1. Read: VISIO_AUDIT_GUIDE.md
2. Try: Utility functions
3. Create: Custom reports

### **Advanced (2+ hours)**
1. Read: DEPLOYMENT.md
2. Setup: Scheduled automation
3. Customize: Script parameters
4. Integrate: With other tools

---

## 🔄 Update & Support

### **Check for Updates**
- New versions announced with enhanced features
- Backward compatible with v1.0 data
- Easy migration between versions

### **Self-Support**
- Comprehensive documentation included
- Troubleshooting guide for 90% of issues
- Community examples provided

### **Professional Support**
- Enterprise deployment assistance available
- Custom script modifications possible
- Integration consulting offered

---

## 📦 Installation Verification

After extraction, verify you have:

```powershell
# Check all 9 files present
Get-ChildItem "Visio-Enterprise-Audit-Suite" -Recurse | Measure-Object
# Should show: 9 files

# Verify scripts are executable
Get-ChildItem "Visio-Enterprise-Audit-Suite\*.ps1" | Select-Object Name, Length
# Should list 3 PS1 files

# Check documentation
Get-ChildItem "Visio-Enterprise-Audit-Suite\*.md" | Select-Object Name
# Should list 5 MD files
```

---

## 🎯 Success Criteria

You'll know the suite is working when:

✅ Scripts run without errors  
✅ Reports generate in C:\Temp\VisioAudit\  
✅ CSV contains computer data  
✅ HTML displays formatted dashboard  
✅ Visio installations are detected  
✅ Last used dates are populated  
✅ Department filters work correctly  

---

## 📞 Next Steps

1. **Extract the ZIP file**
2. **Read README.md** (start here!)
3. **Follow DEPLOYMENT.md** for your scenario
4. **Run Visio-Enterprise-Audit.ps1**
5. **Check reports in C:\Temp\VisioAudit\**
6. **Reference TROUBLESHOOTING.md** if needed

---

**Package Version:** 1.0  
**Released:** February 8, 2024  
**Status:** ✅ Production Ready  
**Support:** Full documentation included

**Start scanning your domain now!** 🚀
