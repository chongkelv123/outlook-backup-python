# Outlook Email Backup Tool

A Python GUI application for backing up Microsoft Outlook emails to local drive as .eml/.msg files with advanced filtering options.

## Features

- **Multiple Filter Options**
  - Date range filtering
  - Sender email filtering
  - Subject keyword search
  - Folder selection (Inbox, Sent Items, custom folders)

- **Flexible Output Settings**
  - Choose backup location
  - Include/exclude attachments
  - Organize by date (YYYY/MM folder structure)
  - Automatic filename collision handling

- **User-Friendly Interface**
  - Preview email count before backup
  - Real-time progress tracking
  - Detailed status logging
  - Responsive GUI with background processing

## Requirements

- Windows OS (required for Outlook COM automation)
- Python 3.7 or higher
- Microsoft Outlook installed and configured
- Active Outlook profile (must be logged in)

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
   - Check "Include Attachments" to save email attachments
   - Check "Organize by Date" to create YYYY/MM folder structure

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

**Without "Organize by Date":**
```
BackupFolder/
├── email_20260129_103045_Meeting_reminder.msg
├── email_20260129_103045_Meeting_reminder_attachments/
│   └── document.pdf
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

### Filename Format

Emails are saved with the following naming convention:
```
email_YYYYMMDD_HHMMSS_[first20chars_of_subject].[msg/eml]
```

Examples:
- `email_20260129_103045_Meeting_reminder.msg`
- `email_20260128_150230_Project_update.msg`

### Configuration File

The application automatically saves your last backup location in `config.json`. This file is created automatically and doesn't need manual editing.

## Troubleshooting

### "Failed to connect to Outlook"

**Solution:**
- Make sure Outlook is installed
- Open Outlook at least once to set up your profile
- Ensure you're logged into your Outlook account
- Try closing and reopening Outlook

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

### Emails saved as .msg instead of .eml

**Explanation:**
- Outlook COM API doesn't directly support .eml export in all versions
- .msg format is Microsoft's native format and fully preserves email content
- Both formats can be opened in Outlook and most email clients

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
4. **Export**: Saves each email using Outlook's SaveAs method in .msg format
5. **Attachments**: If enabled, saves attachments to separate subfolders
6. **Threading**: Backup runs in background thread to keep GUI responsive

### Limitations

- Windows-only (requires Outlook COM interface)
- Requires Outlook to be installed and configured
- Large mailboxes may take significant time to backup
- .msg format used instead of .eml (Outlook API limitation)

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

- **v1.0** (2026-01-29)
  - Initial release
  - Core backup functionality
  - Multiple filter options
  - Progress tracking
  - Attachment support
  - Date-based organization
