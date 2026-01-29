# Outlook Email Backup Tool

A Python GUI application for backing up Microsoft Outlook emails to local drive as .msg files (Outlook native format) with advanced filtering and flexible folder organization options.

## Features

- **Multiple Filter Options**
  - Date range filtering
  - Sender email filtering
  - Subject keyword search
  - Folder selection (Inbox, Sent Items, custom folders)

- **Flexible Output Settings**
  - Choose backup location
  - Include/exclude attachments (Note: .msg format embeds attachments automatically)
  - Organize by date (YYYY/MM folder structure)
  - Organize by sender email (when sender filter is enabled)
  - Organize by subject (when subject filter is enabled)
  - Priority-based organization: Sender > Subject > Date
  - Automatic filename collision handling

- **User-Friendly Interface**
  - Preview email count before backup
  - Real-time progress tracking
  - Detailed status logging
  - Responsive GUI with background processing

## Requirements

- Windows OS (required for Outlook COM automation)
- Python 3.7 or higher
- **Microsoft Outlook Classic (Desktop Version) - REQUIRED**
- Active Outlook profile (must be logged in)

### ⚠️ Important: Classic Outlook Only

**This tool ONLY works with Classic Outlook (Desktop version with COM automation).**

**NOT Compatible with:**
- ❌ New Outlook (web-based version introduced in 2023)
- ❌ Outlook Web App (OWA)
- ❌ Outlook.com web interface

**Why Classic Outlook Only?**
The new Outlook is a web-based application that doesn't expose COM/MAPI interfaces required for automation. Classic Outlook provides full COM automation support, allowing this tool to access and export your emails locally.

### How to Switch to Classic Outlook

If you're using the new Outlook, you can easily switch back:

1. **Open Outlook**
2. **Look for the toggle switch** in the top-right corner of the window
3. **Find the option labeled:** "Try the new Outlook" or similar
4. **Turn OFF the toggle** (slide it to the left/off position)
5. **Outlook will restart** in Classic mode
6. **Run this backup tool** after Classic Outlook has fully loaded

**Note:** Classic Outlook is fully supported by Microsoft and provides all features plus COM automation capabilities.

## Installation

### 1. Clone or Download

Download this project to your local machine.

### 2. Install Python Dependencies

Open Command Prompt or PowerShell in the project directory and run:

```bash
pip install -r requirements.txt
```

This will install:
- `pywin32` - For Outlook COM automation
- `tkcalendar` - For date picker widgets

**Note:** tkinter is included with Python standard library on Windows.

### 3. Verify Outlook Installation

Make sure Microsoft Outlook is:
- Installed on your computer
- Running (at least once to set up profile)
- Configured with an email account

## Usage

### Starting the Application

Run the application from command line:

```bash
python main.py
```

Or double-click `main.py` if Python is associated with .py files.

### Using the Application

1. **Set Up Filters (Optional)**
   - **Date Range**: Enable to filter emails by date (default: last 3 months)
   - **Sender**: Enable to filter by sender email address
   - **Subject**: Enable to filter by keywords in subject
   - **Folder**: Select which Outlook folder to backup (default: Inbox)

2. **Configure Output Settings**
   - Click "Browse..." to select backup location
   - Check "Include Attachments" to save email attachments (Note: .msg format embeds attachments by default)
   - Check "Organize by Date" to create YYYY/MM folder structure
   - **Note**: When Sender filter is enabled, emails will be organized by sender email address
   - **Note**: When Subject filter is enabled, emails will be organized by subject name
   - Priority: Sender organization > Subject organization > Date organization

3. **Preview Count** (Recommended)
   - Click "Preview Count" to see how many emails match your filters
   - This helps verify your filters before starting backup

4. **Start Backup**
   - Click "Start Backup" to begin the export process
   - Progress bar and status log will show real-time updates
   - Do not close Outlook during backup

5. **After Backup**
   - A summary report will show:
     - Number of emails exported
     - Total size
     - Number of errors (if any)
   - Find your backed up emails in the selected backup location

### Output File Structure

**Basic (No organization):**
```
BackupFolder/
├── email_20260129_103045_Meeting_reminder.msg
├── email_20260128_150230_Project_update.msg
└── ...
```

**With "Organize by Date":**
```
BackupFolder/
├── 2026/
│   ├── 01/
│   │   ├── email_20260129_103045_Meeting_reminder.msg
│   │   └── email_20260128_150230_Project_update.msg
│   └── 02/
│       └── email_20260201_093015_Status_report.msg
└── 2025/
    └── 12/
        └── email_20251231_170000_Happy_New_Year.msg
```

**With Sender Filter Enabled (Organize by Sender):**
```
BackupFolder/
├── john.doe@company.com/
│   ├── email_20260129_103045_Meeting_reminder.msg
│   └── email_20260128_150230_Project_update.msg
├── jane.smith@company.com/
│   └── email_20260127_093015_Status_report.msg
└── team@company.com/
    └── email_20260126_140000_Weekly_update.msg
```

**With Sender Filter + Date Organization:**
```
BackupFolder/
├── john.doe@company.com/
│   └── 2026/
│       └── 01/
│           ├── email_20260129_103045_Meeting_reminder.msg
│           └── email_20260128_150230_Project_update.msg
└── jane.smith@company.com/
    └── 2026/
        └── 01/
            └── email_20260127_093015_Status_report.msg
```

**With Subject Filter Enabled (Organize by Subject):**
```
BackupFolder/
├── Project_Alpha_Updates/
│   ├── email_20260129_103045_Re_Project_Alpha_.msg
│   └── email_20260128_150230_Fwd_Project_Alpha.msg
├── Weekly_Report/
│   └── email_20260127_093015_Weekly_Report.msg
└── No_Subject/
    └── email_20260126_140000_no_subject.msg
```

**Note:** .MSG format automatically embeds attachments within the email file, preserving all Outlook metadata including categories, flags, importance, read receipts, and RTF formatting.

### Filename Format

Emails are saved with the following naming convention:
```
email_YYYYMMDD_HHMMSS_[first20chars_of_subject].msg
```

Examples:
- `email_20260129_103045_Meeting_reminder.msg`
- `email_20260128_150230_Project_update.msg`
- `email_20260127_143000_no_subject.msg`

**Note:** .MSG is Microsoft Outlook's native message format that preserves all email properties and embedded attachments.

### Configuration File

The application automatically saves your last backup location in `config.json`. This file is created automatically and doesn't need manual editing.

## Troubleshooting

### "Failed to connect to Outlook" or "CLASSIC Outlook Required"

**Most Common Cause: You're using new Outlook**

**Solution:**
1. Check if you're using the new Outlook (web-based version)
2. Look for a toggle switch in the top-right corner of Outlook
3. Turn OFF "Try the new Outlook"
4. Outlook will restart in Classic mode
5. Run this backup tool again

**Other Solutions:**
- Make sure Classic Outlook is installed and running
- Open Outlook at least once to set up your profile
- Ensure you're logged into your Outlook account
- Try closing and reopening Classic Outlook
- Run this application as Administrator

### "Permission denied" or "Access denied"

**Solution:**
- Make sure the backup location is writable
- Don't backup to system folders (C:\Windows, C:\Program Files)
- Try selecting a different backup location

### Progress bar stuck or application freezes

**Solution:**
- The application is processing large emails or many attachments
- Wait for the operation to complete
- Check the status log for progress updates
- If truly frozen, close and restart the application

### Why .MSG format instead of .EML?

**Benefits of .MSG format:**
- Microsoft Outlook's native format
- Preserves ALL email metadata: categories, flags, importance, read receipts, voting buttons, etc.
- Automatically embeds attachments within the file
- Maintains RTF and HTML formatting perfectly
- Can be opened directly in Outlook with full functionality
- More reliable than .EML for Outlook-specific features

### Some emails are missing

**Possible causes:**
- Filters are too restrictive - check your filter settings
- Emails were corrupted - check the error log
- Folder selection - verify you selected the correct folder

### "Could not process email date/sender/subject"

**Explanation:**
- Some emails may have corrupted or missing metadata
- These warnings are logged but don't stop the backup
- The email may still be exported if possible

### How to Tell if You're Using New Outlook vs Classic Outlook

**New Outlook (NOT Compatible):**
- Modern, colorful interface similar to web interface
- Toggle switch in top-right corner that says "Try the new Outlook"
- Simplified ribbon with icons
- Process name: `olk.exe` or `HxOutlook.exe`

**Classic Outlook (Compatible):**
- Traditional Microsoft Office interface
- Full ribbon with File, Home, Send/Receive tabs
- May have a toggle that says "Try the new Outlook" (turn it OFF)
- Process name: `OUTLOOK.EXE`

**Quick Check:**
- Open Task Manager (Ctrl+Shift+Esc)
- Look for processes:
  - `OUTLOOK.EXE` = Classic Outlook ✓
  - `olk.exe` or `HxOutlook.exe` = New Outlook ✗

## Technical Details

### File Structure

```
outlook_backup_tool/
├── main.py                 # GUI application and main logic
├── outlook_connector.py    # Outlook COM automation
├── email_exporter.py       # Email export functionality
├── filter_engine.py        # Email filtering logic
├── requirements.txt        # Python dependencies
├── config.json            # User preferences (auto-generated)
└── README.md              # This file
```

### How It Works

1. **Connection**: Uses `win32com.client` to connect to the local running Outlook application
2. **Retrieval**: Accesses the specified mail folder and retrieves email items
3. **Filtering**: Applies user-specified filters (date, sender, subject)
4. **Organization**: Determines folder structure based on priority logic:
   - Priority 1: Sender-based organization (if sender filter enabled)
   - Priority 2: Subject-based organization (if subject filter enabled)
   - Priority 3: Date-based organization (if organize by date enabled)
5. **Export**: Saves each email using Outlook's SaveAs method in .msg format (olMSG = 3)
6. **Sender Extraction**: Handles both SMTP and Exchange email addresses properly
7. **Threading**: Backup runs in background thread to keep GUI responsive

### Limitations

- Windows-only (requires Outlook COM interface)
- Requires Outlook to be installed and configured
- Large mailboxes may take significant time to backup
- .msg files are Windows/Outlook-specific (though other email clients may support them)

## Security & Privacy

- **No data transmission**: All operations are local, no data is sent anywhere
- **No authentication needed**: Uses your existing Outlook session
- **Read-only**: Application only reads emails, doesn't modify or delete
- **Safe filters**: All filters are applied in memory before export

## Tips for Best Results

1. **Start small**: Test with a small date range first
2. **Use filters**: Reduce backup time by filtering what you need
3. **Preview count**: Always preview before large backups
4. **Organize by date**: Makes it easier to find specific emails later
5. **Regular backups**: Schedule periodic backups (e.g., monthly)
6. **Close other programs**: For better performance during large backups

## Support

For issues, questions, or feature requests:
- Check the troubleshooting section above
- Review the status log for specific error messages
- Ensure your Outlook installation is up to date

## License

This tool is provided as-is for personal and business use.

## Version History

- **v1.1 (refinement-v1)** (2026-01-29)
  - **BREAKING CHANGE**: Switched from .EML to .MSG format (Outlook native format)
  - Added sender-based folder organization (when sender filter is enabled)
  - Added subject-based folder organization (when subject filter is enabled)
  - Implemented priority-based organization logic (Sender > Subject > Date)
  - Fixed sender filter bug for Exchange emails (now properly retrieves SMTP addresses)
  - Added diagnose_sender.py utility for troubleshooting
  - Improved folder name sanitization
  - .MSG format now preserves ALL Outlook metadata and embeds attachments automatically
  - **Added Classic Outlook detection and warnings**
  - Detects new Outlook (not supported) and shows helpful migration instructions
  - Added compatibility notice in GUI
  - Enhanced error messages to guide users to switch to Classic Outlook
  - Added comprehensive compatibility documentation in README

- **v1.0** (2026-01-29)
  - Initial release
  - Core backup functionality
  - Multiple filter options
  - Progress tracking
  - Attachment support
  - Date-based organization
