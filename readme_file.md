# 📊 Productivity Activity Monitor

A comprehensive Windows productivity tracking tool that monitors your computer usage and automatically generates detailed reports via email. Perfect for freelancers, remote workers, and anyone who wants to understand their digital productivity patterns.

## ✨ Features

### 📈 **Comprehensive Activity Tracking**
- **Foreground App Monitoring**: Tracks all applications and websites you actively use
- **Background Video Detection**: Monitors streaming services and videos playing in background tabs
- **Audio Verification**: Uses audio detection to verify when videos are actually playing vs paused
- **Smart Categorization**: Automatically categorizes activities as Productive, Unproductive, or Uncategorized

### 📧 **Flexible Email Reporting**
- **Multiple Timing Modes**: 
  - Daily reports at specific times (e.g., 6:00 PM)
  - Interval-based reports (e.g., every 30 minutes)
  - Friday-only mode for weekly summaries
- **Professional Reports**: Detailed HTML-formatted reports with productivity scores
- **Automatic Email**: Integrates with Microsoft Outlook for seamless report delivery

### 🖥️ **Advanced System Integration**
- **Login/Logout Tracking**: Monitors work sessions and break times
- **System Information**: Includes location, IP address, and system details in reports
- **Data Persistence**: Automatically saves and restores tracking data across sessions
- **Hybrid Outlook Support**: Works with both Microsoft Store and Desktop Outlook versions

### 🔍 **Intelligent Analysis**
- **Productivity Scoring**: Calculates daily productivity percentages
- **Session Chaining**: Groups long work sessions for better analysis
- **Website Recognition**: Identifies popular sites like YouTube, Amazon, GitHub, etc.
- **Background Activity Impact**: Assesses how background videos affect productivity

## 🚀 Quick Start

### Prerequisites
- Windows 10/11
- Microsoft Outlook (Store or Desktop version)
- Python 3.8+ (if running from source) OR use the compiled executable

### ⚠️ **IMPORTANT: Configuration File Location**

**The `config.txt` file MUST be placed in the same directory as:**
- The Python script (`activity_monitor.py`) if running from source
- The compiled executable (`activity_monitor.exe`) if using the standalone version

**Incorrect placement will cause the script to fail!**

```
✅ Correct structure:
my-folder/
├── activity_monitor.exe    (or activity_monitor.py)
└── config.txt

❌ Incorrect - config.txt in wrong location:
my-folder/
├── activity_monitor.exe
└── configs/
    └── config.txt          (This won't work!)
```

## ⚙️ Configuration

Edit `config.txt` to customize your monitoring preferences:

### Basic Setup
```ini
# Your email address for reports
to_email=your.email@company.com

# Choose your schedule
friday_only=false              # Set to 'true' for weekly reports only
email_timing_mode=time_of_day   # or 'interval' for regular updates
daily_email_time=18:00          # 6:00 PM daily report
```

### Timing Modes

#### Daily Reports (Recommended)
```ini
email_timing_mode=time_of_day
daily_email_time=17:00          # 5:00 PM
friday_only=false               # Run every day
```

#### Weekly Reports (Friday Only)
```ini
email_timing_mode=time_of_day
daily_email_time=17:00          # 5:00 PM on Fridays
friday_only=true                # Only run on Fridays
```

#### Active Monitoring (Interval-based)
```ini
email_timing_mode=interval
email_interval=1800             # Every 30 minutes
friday_only=false               # Run every day
```

## 📊 Sample Report Features

Your productivity reports include:

- **📈 Productivity Score**: Overall daily productivity percentage
- **⏱️ Time Allocation**: Breakdown of productive vs unproductive time
- **🏆 Top Activities**: Most-used productive applications
- **⚠️ Time Drains**: Biggest unproductive distractions
- **📺 Background Activity**: Videos/music playing while working
- **📋 Session Analysis**: Login/logout times and work sessions
- **👤 System Information**: Location, IP address, computer details
- **🌐 Website Tracking**: Detailed breakdown of website usage

## 🛠️ Technical Details

### System Requirements
- **OS**: Windows 10/11 (64-bit recommended)
- **RAM**: 100MB+ available memory
- **Storage**: 50MB+ free space for logs and reports
- **Network**: Internet connection for email and location services

### Dependencies (For Source Installation)
```
psutil>=5.9.0
win32gui
win32process
win32com.client
win32api
wmi
pycaw
requests
```

### Data Storage
- **Logs**: Stored in `logs/` subdirectory
- **Reports**: Saved as dated text files
- **Persistence**: Automatic data backup and recovery
- **Cleanup**: Old files automatically removed after 7 days

## 🔒 Privacy & Security

- **Local Processing**: All analysis happens on your computer
- **No Cloud Storage**: Data stays on your machine
- **Email Only**: Reports only sent to your configured email address
- **Geolocation**: Uses public IP services (can be disabled in code)
- **System Access**: Requires standard Windows API permissions

## 🐛 Troubleshooting

### Common Issues

#### Email Not Sending
- ✅ Verify Outlook is installed and configured
- ✅ Check `to_email` address in config.txt
- ✅ Ensure Windows allows script to access Outlook
- ✅ Try running as Administrator

#### Script Won't Start
- ✅ Verify `config.txt` is in the same folder as executable/script
- ✅ Check Windows Defender isn't blocking the executable
- ✅ Ensure all dependencies are installed (if running from source)

#### No Activity Detected
- ✅ Try running as Administrator for better system access
- ✅ Check Windows privacy settings for app permissions
- ✅ Verify script is running during active computer use

#### Friday-Only Mode Not Working
- ✅ Check `friday_only=true` in config.txt (case sensitive)
- ✅ Verify it's actually Friday in your timezone
- ✅ Look for "Friday detected!" message in console

### Debug Mode
Run with verbose output to troubleshoot:
```bash
# From source
python activity_monitor.py --debug

# Compiled executable (redirect output)
activity_monitor.exe > debug.log 2>&1
```

## ❓ FAQ

**Q: Will this slow down my computer?**
A: No, the monitor uses minimal CPU and memory (~10-20MB RAM).

**Q: Can I run this on multiple computers?**
A: Yes! Just install and configure on each machine with the same email address.

**Q: Is my data private?**
A: Absolutely. Everything processes locally and only your email receives reports.

**Q: Can I customize what gets tracked?**
A: Yes! Edit the categorization rules in the source code to match your workflow.

**Q: What if I forget to start the script?**
A: Set it up as a Windows startup program or scheduled task for automatic launching.

---
