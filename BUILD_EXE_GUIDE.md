# Building Executable (.exe) Guide

This guide will help you create a standalone executable file that users can run without installing Python.

## Method 1: Automated Build (Recommended)

### Quick Steps:

1. **Run the build script:**
   ```bash
   # Option A: Double-click
   build_exe.bat

   # Option B: Command line
   python build_exe.py
   ```

2. **Wait for completion** (takes 3-5 minutes)

3. **Find your executable:**
   - Single file: `dist/OutlookBackupTool.exe`
   - Complete package: `OutlookBackupTool_Portable/`

That's it! The automated script handles everything.

---

## Method 2: Manual Build

If you prefer to build manually or the automated script doesn't work:

### Step 1: Install PyInstaller

```bash
pip install pyinstaller
```

### Step 2: Build the Executable

```bash
pyinstaller --name "OutlookBackupTool" ^
    --onefile ^
    --windowed ^
    --clean ^
    --add-data "config.json;." ^
    --hidden-import win32com ^
    --hidden-import win32com.client ^
    --hidden-import pythoncom ^
    --hidden-import pywintypes ^
    --hidden-import win32timezone ^
    --hidden-import tkcalendar ^
    --hidden-import babel.numbers ^
    main.py
```

### Step 3: Find Your Executable

The executable will be in: `dist/OutlookBackupTool.exe`

---

## Understanding the Build Options

### PyInstaller Options Explained:

- `--onefile` - Create a single .exe file (not a folder)
- `--windowed` - No console window (GUI only)
- `--clean` - Clean previous builds
- `--name "OutlookBackupTool"` - Name of the executable
- `--add-data "config.json;."` - Include config file
- `--hidden-import` - Include modules that PyInstaller might miss

### Why Hidden Imports?

PyInstaller sometimes misses dynamically imported modules:
- `win32com.client` - Outlook COM automation
- `pythoncom` - COM threading support
- `tkcalendar` - Date picker widget
- `babel.numbers` - Required by tkcalendar

---

## Build Output

After building, you'll have:

```
project/
├── dist/
│   └── OutlookBackupTool.exe    ← Your standalone executable!
├── build/                        ← Temporary files (can delete)
└── OutlookBackupTool.spec       ← Build configuration (can delete)
```

### Executable Size:
- Expected size: 25-40 MB
- Includes Python runtime and all dependencies
- Single file, no installation needed

---

## Distribution Package

The automated build creates a complete package:

```
OutlookBackupTool_Portable/
├── OutlookBackupTool.exe        ← Main executable
├── START_HERE.txt               ← Quick instructions
├── README.md                    ← Full documentation
├── QUICK_START.md              ← Quick start guide
├── TROUBLESHOOTING.md          ← Troubleshooting guide
├── FIX_YOUR_ERROR.txt          ← Error solutions
└── Diagnostic_Tools/            ← Diagnostic utilities
    ├── test_connection.py
    ├── test_outlook.bat
    ├── diagnose_outlook.py
    └── diagnose.bat
```

### To Distribute:

1. **Option A: Share the folder**
   - Zip the entire `OutlookBackupTool_Portable` folder
   - Share the ZIP file

2. **Option B: Share just the .exe**
   - Copy `dist/OutlookBackupTool.exe`
   - Share the single file

---

## Testing the Executable

### Before Distribution:

1. **Test locally:**
   ```bash
   dist\OutlookBackupTool.exe
   ```

2. **Test on another computer** (without Python installed)
   - Copy the .exe to another PC
   - Make sure that PC has Outlook installed
   - Run the .exe

3. **Test all features:**
   - Connection to Outlook ✓
   - Preview Count ✓
   - Backup emails ✓
   - Save attachments ✓

---

## Troubleshooting Build Issues

### Issue: "PyInstaller not found"

**Solution:**
```bash
pip install pyinstaller
```

---

### Issue: "Module not found" when running .exe

**Solution:** Add the missing module as a hidden import:
```bash
pyinstaller ... --hidden-import module_name ...
```

Common missing modules already included:
- win32com
- pythoncom
- tkcalendar

---

### Issue: .exe file is too large (>50 MB)

**This is normal!** The executable includes:
- Python runtime (~15 MB)
- PyWin32 libraries (~10 MB)
- Tkinter GUI (~5 MB)
- Other dependencies

**Ways to reduce size:**
1. Use `--onefile` (already used)
2. Use UPX compression:
   ```bash
   pyinstaller ... --upx-dir=path/to/upx ...
   ```

---

### Issue: Antivirus flags the .exe

**This is common with PyInstaller executables!**

**Why it happens:**
- PyInstaller bundles everything into one .exe
- Some antivirus software flags this as suspicious
- It's a false positive

**Solutions:**
1. **Submit to antivirus vendors** as false positive
2. **Code sign the executable** (requires certificate)
3. **Use --onedir** instead of --onefile (less suspicious)
4. **Build on a clean system** (antivirus scans build environment)

**For users:**
- Add executable to antivirus exclusions
- Right-click → Properties → Unblock

---

### Issue: Build fails with errors

**Check these:**
1. All dependencies installed?
   ```bash
   pip install -r requirements.txt
   pip install pyinstaller
   ```

2. No syntax errors in Python code?
   ```bash
   python main.py
   ```

3. All import statements work?
   ```bash
   python -c "import win32com.client; import pythoncom; import tkcalendar"
   ```

4. Try cleaning first:
   ```bash
   # Delete these folders if they exist
   rmdir /s /q build dist __pycache__
   del OutlookBackupTool.spec
   ```

---

## Advanced: Customizing the Build

### Adding an Icon

1. **Get an .ico file** (256x256 recommended)

2. **Add to PyInstaller command:**
   ```bash
   pyinstaller ... --icon=outlook_backup.ico ...
   ```

3. **Or in spec file:**
   ```python
   exe = EXE(
       ...
       icon='outlook_backup.ico',
       ...
   )
   ```

---

### Creating an Installer

Want to create a proper installer (like Setup.exe)?

**Option 1: Inno Setup** (Recommended)
1. Download: https://jrsoftware.org/isinfo.php
2. Create setup script
3. Compile to Setup.exe

**Option 2: NSIS**
1. Download: https://nsis.sourceforge.io/
2. Create NSIS script
3. Compile installer

**Option 3: WiX Toolset**
- Creates .msi installers
- More complex but professional

---

## Build Best Practices

### 1. Clean Environment
```bash
# Create virtual environment
python -m venv venv
venv\Scripts\activate

# Install only required packages
pip install -r requirements.txt
pip install pyinstaller

# Build
python build_exe.py
```

### 2. Test Thoroughly
- Test on Windows 7, 10, 11
- Test with and without Python installed
- Test with different Outlook versions
- Test on clean VM

### 3. Version Control
- Tag releases in git
- Keep build logs
- Document any build issues

### 4. Distribution
- Create SHA256 checksum
- Sign the executable (if possible)
- Provide virus scan results
- Include documentation

---

## Alternative: Portable Python Distribution

Instead of creating an .exe, you can create a portable Python bundle:

```
OutlookBackupTool_Portable/
├── python/              ← Portable Python
├── app/                 ← Your application
│   ├── main.py
│   ├── outlook_connector.py
│   └── ...
└── run.bat             ← Launcher
```

**Advantages:**
- No build process
- Easy to update
- No antivirus issues
- Users can see source code

**Disadvantages:**
- Larger size
- More files
- Less "professional" look

---

## FAQ

### Q: Do users need Python installed?
**A:** No! The .exe includes everything.

### Q: Do users need Outlook installed?
**A:** Yes, Microsoft Outlook must be installed and configured.

### Q: Will it work on Mac or Linux?
**A:** No, this is Windows-only (requires Outlook COM interface).

### Q: Can I distribute this commercially?
**A:** Yes, but check licenses of included libraries.

### Q: How do I update the executable?
**A:** Rebuild with updated code:
```bash
python build_exe.py
```

### Q: Why is the .exe so large?
**A:** It includes Python runtime and all dependencies. This is normal for PyInstaller.

### Q: Can I make it smaller?
**A:** Some options exist (UPX compression), but 25-40 MB is typical and acceptable.

---

## Quick Reference

### Build Commands:

**Automated:**
```bash
python build_exe.py
```

**Manual:**
```bash
pip install pyinstaller
pyinstaller OutlookBackupTool.spec
```

**Clean build:**
```bash
pyinstaller --clean OutlookBackupTool.spec
```

### Output Locations:

- Executable: `dist/OutlookBackupTool.exe`
- Complete package: `OutlookBackupTool_Portable/`
- Build files: `build/` (can delete)

### File Sizes:

- Source code: ~100 KB
- Executable: ~25-40 MB
- Complete package: ~26-41 MB

---

## Support

If you have build issues:

1. Check this guide's troubleshooting section
2. Verify all dependencies are installed
3. Try the automated build script
4. Check PyInstaller documentation: https://pyinstaller.org/

---

**Ready to build?**

```bash
python build_exe.py
```

**That's it! Your executable will be ready in a few minutes.**

---

**Last Updated:** 2026-01-29
**PyInstaller Version:** 5.0+
**Tested On:** Windows 10, Windows 11
