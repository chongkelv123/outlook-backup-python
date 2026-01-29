# Troubleshooting Guide

## Common Error: "Server execution failed" (-2146959355)

This is the most common error when connecting to Outlook. Here are proven solutions:

### ✅ Solution 1: Start Outlook First (Fixes 80% of cases)

**Steps:**
1. **Close the backup tool completely**
2. **Open Microsoft Outlook manually**
3. **Wait until Outlook is fully loaded** (you should see your inbox)
4. **Keep Outlook open** (minimize it, don't close)
5. **Now run the backup tool again**

**Why this works:** The COM interface requires Outlook to be running and fully initialized before other applications can connect to it.

---

### ✅ Solution 2: Run as Administrator

**Steps:**
1. Close the backup tool
2. Right-click on `run_backup_tool.bat`
3. Select **"Run as administrator"**
4. Allow the User Account Control prompt
5. Try connecting to Outlook again

**Why this works:** COM automation sometimes requires elevated privileges, especially on corporate or secured systems.

---

### ✅ Solution 3: Register COM Libraries

**Steps:**
1. Open **Command Prompt as Administrator**
2. Run this command:
   ```
   python -m win32com.client.makepy Outlook
   ```
3. Wait for completion (might take 30 seconds)
4. Close Command Prompt
5. Try the backup tool again

**Why this works:** This creates early binding for Outlook COM objects, which improves reliability.

---

### ✅ Solution 4: Check Outlook Profile

**Steps:**
1. Open Outlook manually
2. Go to **File → Account Settings → Account Settings**
3. Verify you have **at least one email account** configured
4. Try sending a test email to yourself
5. If the email works, try the backup tool again

**Why this works:** Outlook needs a configured profile with an active email account for the MAPI namespace to work properly.

---

### ✅ Solution 5: Reinstall pywin32

**Steps:**
1. Open **Command Prompt as Administrator**
2. Run these commands:
   ```
   pip uninstall pywin32
   pip install pywin32
   python -m pywin32_postinstall -install
   ```
3. Restart your computer
4. Start Outlook, then try the backup tool

**Why this works:** Reinstalling pywin32 with post-install ensures COM registration is correct.

---

### ✅ Solution 6: Run Diagnostic Tool

**Steps:**
1. Run `diagnose.bat` or `python diagnose_outlook.py`
2. Review the diagnostic results
3. Follow the specific recommendations provided
4. Fix any failed checks

**What it checks:**
- Python installation
- pywin32 installation
- Outlook installation
- Outlook process running
- Administrator rights
- Actual COM connection test

---

## Other Common Errors

### Error: "Outlook is not running"

**Solution:**
- Simply start Microsoft Outlook
- Wait until you can see your mailbox
- Keep Outlook open while using the backup tool

---

### Error: "Cannot access folders"

**Possible causes:**
1. Outlook profile not configured
2. No email account set up
3. Outlook running in safe mode

**Solution:**
1. Open Outlook normally (not safe mode)
2. Configure at least one email account
3. Verify you can access your Inbox manually
4. Try the backup tool again

---

### Error: "Invalid class string" (-2147221005)

**Cause:** Outlook is not properly installed or registered

**Solution:**
1. Repair Microsoft Office installation:
   - Control Panel → Programs → Microsoft Office
   - Click "Change" → "Quick Repair" or "Online Repair"
2. Restart computer
3. Try again

---

### Error: "Access denied" or Permission errors

**Solution:**
1. Run the backup tool as Administrator
2. Check antivirus settings (might block COM access)
3. Check Windows Defender:
   - Settings → Update & Security → Windows Security
   - Virus & threat protection → Manage settings
   - Add Python to exclusions if needed

---

### Error: "The application called an interface that was marshalled for a different thread" (-2147417842)

**Cause:** COM threading issue - trying to use Outlook objects across different threads

**This has been FIXED in the current version!**

If you see this error with the latest code, try:
1. Make sure you have the latest version of main.py
2. Restart the application
3. If problem persists, reinstall: `pip install --upgrade pywin32`

**Technical explanation:** See `THREADING_FIX_APPLIED.md` for full details.

---

## Testing Your Connection

### Quick Test Steps:

1. **Open Outlook** and verify you can see your inbox
2. **Run the diagnostic tool**: `python diagnose_outlook.py`
3. **Check all results** - all tests should pass
4. **If tests pass**, try the backup tool with a small date range (last 7 days)

---

## Performance Issues

### Backup is slow or hanging

**Causes:**
- Large mailbox (thousands of emails)
- Large attachments
- Slow Outlook connection (network folders)

**Solutions:**
1. **Use smaller date ranges** - backup in chunks (e.g., one month at a time)
2. **Disable attachments** temporarily to test
3. **Close other programs** to free up memory
4. **Check Outlook is not synchronizing** - wait for sync to complete
5. **Use local Outlook data files** (.ost/.pst) instead of online mode

---

## System Requirements

### Minimum Requirements:
- Windows 7/10/11
- Python 3.7 or higher
- Microsoft Outlook 2010 or later
- 4GB RAM (8GB recommended for large mailboxes)
- Outlook configured with active email account

### Verified Configurations:
✓ Windows 10 + Outlook 2016 + Python 3.9
✓ Windows 11 + Outlook 2021 + Python 3.11
✓ Windows 10 + Office 365 + Python 3.10

---

## Advanced Troubleshooting

### Check if pywin32 is properly installed:

Open Python and run:
```python
import win32com.client
print("pywin32 is installed!")
```

If you get an error, reinstall pywin32:
```bash
pip uninstall pywin32
pip install pywin32
```

---

### Check if Outlook COM object is accessible:

Open Python and run:
```python
import win32com.client
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
inbox = namespace.GetDefaultFolder(6)
print(f"Inbox: {inbox.Name}, Items: {inbox.Items.Count}")
```

If this works, the backup tool should work too!

---

### Check Windows Event Viewer:

1. Open **Event Viewer** (search in Start menu)
2. Go to **Windows Logs → Application**
3. Look for errors related to:
   - Outlook
   - MAPI
   - COM/OLE
4. Google any specific error codes you find

---

## Still Having Issues?

### Diagnostic Checklist:

- [ ] Outlook is installed and opens successfully
- [ ] Outlook has at least one email account configured
- [ ] You can manually access your inbox in Outlook
- [ ] Outlook is running (not closed)
- [ ] Python 3.7+ is installed
- [ ] pywin32 is installed (`pip list | findstr pywin32`)
- [ ] You tried running as Administrator
- [ ] You ran the diagnostic tool (`diagnose.bat`)
- [ ] You tried restarting Outlook
- [ ] You tried restarting your computer

### If All Else Fails:

1. **Restart your computer** (fixes ~90% of persistent COM issues)
2. **Repair Office installation**
3. **Update Windows** (some COM issues are fixed in Windows updates)
4. **Check corporate IT policies** (some organizations block COM automation)
5. **Try a different computer** (to isolate if it's system-specific)

---

## Getting Help

### Information to collect:

When asking for help, provide:
1. **Error message** (exact text or screenshot)
2. **Windows version** (run: `winver`)
3. **Python version** (run: `python --version`)
4. **Outlook version** (Outlook → File → Office Account)
5. **Diagnostic tool output** (run `diagnose.bat` and save the output)
6. **When the error occurs** (at startup, during backup, etc.)

### Useful Commands:

Check Python version:
```bash
python --version
```

Check installed packages:
```bash
pip list
```

Check if Outlook is running:
```bash
tasklist | findstr OUTLOOK
```

Test pywin32:
```bash
python -c "import win32com.client; print('OK')"
```

---

## Prevention Tips

### To avoid future issues:

1. **Always start Outlook first** before running the backup tool
2. **Keep Outlook updated** (latest updates have better COM support)
3. **Run backups during low-activity times** (not during email sync)
4. **Use Task Scheduler** for automated backups (with Outlook always running)
5. **Keep pywin32 updated**: `pip install --upgrade pywin32`
6. **Regular system maintenance** (Windows updates, disk cleanup)

---

## Quick Reference: Error Codes

| Error Code | Meaning | Solution |
|------------|---------|----------|
| -2146959355 | Server execution failed | Start Outlook first, run as admin |
| -2147417842 | Wrong thread (marshalling) | Fixed in code - update to latest version |
| -2147221005 | Invalid class string | Reinstall/repair Office |
| -2147221021 | Operation unavailable | Outlook not running |
| -2147024891 | Access denied | Run as administrator |
| -2147352567 | Exception occurred | Check Outlook profile, restart Outlook |

---

**Last Updated:** 2026-01-29
