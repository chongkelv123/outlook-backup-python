"""
Build Script for Creating Executable
Prepares and builds the Outlook Backup Tool as a standalone .exe
"""

import os
import sys
import shutil
import subprocess


def print_header(text):
    """Print formatted header"""
    print("\n" + "="*60)
    print(f"  {text}")
    print("="*60 + "\n")


def check_pyinstaller():
    """Check if PyInstaller is installed"""
    try:
        import PyInstaller
        print("✓ PyInstaller is installed")
        return True
    except ImportError:
        print("✗ PyInstaller is not installed")
        print("\nInstalling PyInstaller...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
            print("✓ PyInstaller installed successfully")
            return True
        except Exception as e:
            print(f"✗ Failed to install PyInstaller: {e}")
            return False


def clean_build_folders():
    """Clean previous build folders"""
    print("Cleaning previous build folders...")
    folders_to_clean = ['build', 'dist', '__pycache__']

    for folder in folders_to_clean:
        if os.path.exists(folder):
            try:
                shutil.rmtree(folder)
                print(f"  ✓ Removed {folder}/")
            except Exception as e:
                print(f"  ⚠ Could not remove {folder}/: {e}")

    # Remove .spec file if exists
    spec_file = "OutlookBackupTool.spec"
    if os.path.exists(spec_file):
        try:
            os.remove(spec_file)
            print(f"  ✓ Removed {spec_file}")
        except Exception as e:
            print(f"  ⚠ Could not remove {spec_file}: {e}")

    print()


def create_spec_file():
    """Create PyInstaller spec file with proper configuration"""
    print("Creating PyInstaller spec file...")

    spec_content = '''# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('config.json', '.'),
    ],
    hiddenimports=[
        'win32com',
        'win32com.client',
        'pythoncom',
        'pywintypes',
        'win32timezone',
        'tkcalendar',
        'babel.numbers',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='OutlookBackupTool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Add icon='icon.ico' if you have an icon file
    version_file=None,
)
'''

    with open('OutlookBackupTool.spec', 'w') as f:
        f.write(spec_content)

    print("✓ Spec file created\n")


def build_executable():
    """Build the executable using PyInstaller"""
    print_header("Building Executable")

    print("This may take several minutes...\n")

    try:
        # Run PyInstaller with the spec file
        cmd = [sys.executable, "-m", "PyInstaller", "--clean", "OutlookBackupTool.spec"]

        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode == 0:
            print("✓ Build completed successfully!")
            return True
        else:
            print("✗ Build failed!")
            print("\nError output:")
            print(result.stderr)
            return False

    except Exception as e:
        print(f"✗ Build failed with exception: {e}")
        return False


def create_distribution_package():
    """Create a complete distribution package"""
    print_header("Creating Distribution Package")

    # Create distribution folder
    dist_folder = "OutlookBackupTool_Portable"

    if os.path.exists(dist_folder):
        shutil.rmtree(dist_folder)

    os.makedirs(dist_folder)
    print(f"Created folder: {dist_folder}/")

    # Copy executable
    exe_source = os.path.join("dist", "OutlookBackupTool.exe")
    if os.path.exists(exe_source):
        shutil.copy2(exe_source, dist_folder)
        print(f"  ✓ Copied OutlookBackupTool.exe")
    else:
        print(f"  ✗ Executable not found at {exe_source}")
        return False

    # Copy documentation files
    docs_to_copy = [
        'README.md',
        'QUICK_START.md',
        'TROUBLESHOOTING.md',
        'FIX_YOUR_ERROR.txt',
    ]

    for doc in docs_to_copy:
        if os.path.exists(doc):
            shutil.copy2(doc, dist_folder)
            print(f"  ✓ Copied {doc}")

    # Copy diagnostic tools
    tools_to_copy = [
        'test_connection.py',
        'test_outlook.bat',
        'diagnose_outlook.py',
        'diagnose.bat',
    ]

    tools_folder = os.path.join(dist_folder, "Diagnostic_Tools")
    os.makedirs(tools_folder)

    for tool in tools_to_copy:
        if os.path.exists(tool):
            shutil.copy2(tool, tools_folder)
            print(f"  ✓ Copied {tool} to Diagnostic_Tools/")

    # Create README for distribution
    create_distribution_readme(dist_folder)

    print(f"\n✓ Distribution package created: {dist_folder}/")
    return True


def create_distribution_readme(dist_folder):
    """Create a README for the distribution package"""
    readme_content = """# Outlook Email Backup Tool - Portable Edition

This is a standalone version that doesn't require Python installation.

## Quick Start

1. Make sure Microsoft Outlook is installed and running
2. Double-click **OutlookBackupTool.exe** to start
3. Select your backup location
4. Configure filters (optional)
5. Click "Start Backup"

## Requirements

- Windows 7/10/11
- Microsoft Outlook installed (part of Microsoft Office)
- Outlook must be running with an active profile

## First Time Setup

1. **Start Outlook** manually first
2. Wait until Outlook is fully loaded
3. Keep Outlook running
4. Run OutlookBackupTool.exe

## Troubleshooting

If you encounter errors:

1. Make sure Outlook is running
2. Try running as Administrator (right-click → Run as administrator)
3. Check the TROUBLESHOOTING.md file
4. Run diagnostic tools in the Diagnostic_Tools folder

## Files Included

- **OutlookBackupTool.exe** - Main application
- **README.md** - Complete documentation
- **QUICK_START.md** - Quick start guide
- **TROUBLESHOOTING.md** - Troubleshooting guide
- **FIX_YOUR_ERROR.txt** - Common error solutions
- **Diagnostic_Tools/** - Connection test and diagnostic tools

## Support

For detailed documentation, see README.md
For quick start, see QUICK_START.md
For problems, see TROUBLESHOOTING.md

## Version

Version: 1.1
Build Date: 2026-01-29
"""

    with open(os.path.join(dist_folder, "START_HERE.txt"), 'w') as f:
        f.write(readme_content)

    print("  ✓ Created START_HERE.txt")


def get_exe_size():
    """Get the size of the generated executable"""
    exe_path = os.path.join("dist", "OutlookBackupTool.exe")
    if os.path.exists(exe_path):
        size_bytes = os.path.getsize(exe_path)
        size_mb = size_bytes / (1024 * 1024)
        return f"{size_mb:.2f} MB"
    return "Unknown"


def main():
    """Main build process"""
    print("\n" + "="*60)
    print("  OUTLOOK BACKUP TOOL - EXECUTABLE BUILDER")
    print("="*60 + "\n")

    print("This script will create a standalone .exe file")
    print("that can be distributed without Python.\n")

    # Step 1: Check PyInstaller
    print_header("Step 1: Checking PyInstaller")
    if not check_pyinstaller():
        print("\nBuild aborted.")
        input("Press Enter to exit...")
        return

    # Step 2: Clean previous builds
    print_header("Step 2: Cleaning Previous Builds")
    clean_build_folders()

    # Step 3: Create spec file
    print_header("Step 3: Creating Configuration")
    create_spec_file()

    # Step 4: Build executable
    if not build_executable():
        print("\nBuild failed. Please check the errors above.")
        input("Press Enter to exit...")
        return

    # Step 5: Create distribution package
    if not create_distribution_package():
        print("\nFailed to create distribution package.")
        input("Press Enter to exit...")
        return

    # Success!
    print_header("BUILD SUCCESSFUL!")

    exe_size = get_exe_size()

    print(f"""
✓ Executable created successfully!

File: dist/OutlookBackupTool.exe
Size: {exe_size}

Distribution package: OutlookBackupTool_Portable/

You can now:
1. Test the executable: dist\\OutlookBackupTool.exe
2. Distribute the entire folder: OutlookBackupTool_Portable\\
3. Create a ZIP file for easy sharing

The executable includes everything needed to run the application.
Users do NOT need Python installed!

Notes:
- Users still need Microsoft Outlook installed
- Outlook must be running to use the tool
- Works on Windows 7/10/11
""")

    input("\nPress Enter to exit...")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nBuild cancelled by user.")
        sys.exit(0)
    except Exception as e:
        print(f"\n\nUnexpected error: {e}")
        import traceback
        traceback.print_exc()
        input("\nPress Enter to exit...")
        sys.exit(1)
