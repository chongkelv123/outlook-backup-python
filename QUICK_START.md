# Quick Start Guide

Get up and running with the Outlook Email Backup Tool in 3 easy steps!

## Prerequisites

Before you begin, make sure you have:
- ✅ Windows Operating System
- ✅ Microsoft Outlook installed and configured
- ✅ Python 3.7+ installed ([Download Python](https://www.python.org/downloads/))

## Step 1: Install Dependencies

**Option A: Using the install script (Recommended)**
1. Double-click `install_dependencies.bat`
2. Wait for installation to complete
3. Press any key when done

**Option B: Manual installation**
1. Open Command Prompt in this folder
2. Run: `pip install -r requirements.txt`

## Step 2: Run the Application

**Option A: Using the launcher (Recommended)**
1. Double-click `run_backup_tool.bat`
2. The application window will open

**Option B: Manual start**
1. Open Command Prompt in this folder
2. Run: `python main.py`

## Step 3: Backup Your Emails

### First-Time Setup
1. **Select Backup Location**
   - Click "Browse..." under "Backup Location"
   - Choose where you want to save your emails
   - Example: `D:\Email Backups\`

2. **Configure Filters** (Optional but recommended for first test)
   - Enable "Date Range" filter
   - Set a short range (e.g., last 7 days)
   - This makes your first backup quick

3. **Configure Output Settings**
   - Check "Include Attachments" if you want to save attachments
   - Check "Organize by Date" for organized folder structure

### Running Your First Backup

1. **Preview Count**
   - Click "Preview Count" button
   - See how many emails will be backed up
   - This is a good test to verify everything works

2. **Start Backup**
   - Click "Start Backup" button
   - Confirm the settings
   - Watch the progress bar and status log
   - Wait for completion message

3. **Verify Results**
   - Navigate to your backup location
   - Open a few backed up email files to verify
   - They should open in Outlook or your default email client

## Common First-Time Use Cases

### Backup Last 3 Months of Inbox
```
✅ Date Range: Last 3 months
✅ Folder: Inbox
✅ Include Attachments: Yes
✅ Organize by Date: Yes
```

### Backup All Sent Emails
```
✅ Folder: Sent Items
✅ Include Attachments: Yes
✅ Organize by Date: Yes
```

### Backup Emails from Specific Sender
```
✅ Sender Filter: boss@company.com
✅ Include Attachments: Yes
```

### Backup Project-Related Emails
```
✅ Subject Filter: Project Alpha
✅ Include Attachments: Yes
```

## Troubleshooting First Run

### "Failed to connect to Outlook"
**Fix:**
1. Open Microsoft Outlook
2. Make sure you're logged in
3. Try running the backup tool again

### "Invalid Location"
**Fix:**
1. Click "Browse..." button
2. Select a valid folder (not a system folder)
3. Make sure you have write permissions

### Nothing happens when I click "Start Backup"
**Fix:**
1. Make sure you selected a backup location
2. Check if Outlook is running
3. Look at the status log for error messages

## Tips for Best Experience

1. **Start Small**: For your first backup, use a short date range to test
2. **Check Preview**: Always click "Preview Count" before backing up
3. **Keep Outlook Open**: Don't close Outlook during backup
4. **Be Patient**: Large backups take time - watch the progress bar
5. **Organize by Date**: Makes finding emails later much easier

## What's Next?

After your first successful backup:
- Set up regular backup schedule (e.g., monthly)
- Experiment with different filters
- Backup different folders (Sent Items, specific projects)
- Store backups on external drive for safety

## Need Help?

- Check the main [README.md](README.md) for detailed documentation
- Review the Troubleshooting section in README.md
- Check the status log in the application for specific errors

## Quick Reference

### File Locations
- **Application**: `main.py`
- **Configuration**: `config.json` (auto-created)
- **Dependencies**: `requirements.txt`

### Keyboard Shortcuts (in app)
- ESC: Cancel operation
- Enter: Start backup (when button focused)

### Backup File Format
```
email_YYYYMMDD_HHMMSS_[subject].msg
```
Example: `email_20260129_103045_Meeting_notes.msg`

---

**Ready to go?** Double-click `run_backup_tool.bat` to start!
