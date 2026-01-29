"""
Simple Outlook Connection Test
Quick test to verify Outlook connection before running the full backup tool
"""

import sys


def test_connection():
    """Test connection to Outlook"""
    print("="*60)
    print("  OUTLOOK CONNECTION TEST")
    print("="*60)
    print()

    # Step 1: Check imports
    print("Step 1: Checking imports...")
    try:
        import win32com.client
        import pywintypes
        print("✓ pywin32 is installed")
    except ImportError as e:
        print(f"✗ Error: {e}")
        print("\nFix: Run 'pip install pywin32'")
        input("\nPress Enter to exit...")
        return False
    print()

    # Step 2: Check if Outlook is running
    print("Step 2: Checking if Outlook is running...")
    try:
        import subprocess
        result = subprocess.run(
            ['tasklist', '/FI', 'IMAGENAME eq OUTLOOK.EXE'],
            capture_output=True,
            text=True,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        if 'OUTLOOK.EXE' in result.stdout:
            print("✓ Outlook is running")
        else:
            print("✗ Outlook is NOT running")
            print("\nFix: Please start Microsoft Outlook first")
            input("\nPress Enter to exit...")
            return False
    except Exception as e:
        print(f"⚠ Warning: Cannot verify Outlook process: {e}")
    print()

    # Step 3: Try to connect
    print("Step 3: Attempting to connect to Outlook...")
    try:
        outlook = win32com.client.GetActiveObject("Outlook.Application")
        print("✓ Connected to Outlook (GetActiveObject)")
    except:
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            print("✓ Connected to Outlook (Dispatch)")
        except Exception as e:
            print(f"✗ Failed to connect: {e}")
            print("\nCommon fixes:")
            print("1. Start Outlook and wait until fully loaded")
            print("2. Run this script as Administrator")
            print("3. Run: python -m win32com.client.makepy Outlook")
            input("\nPress Enter to exit...")
            return False
    print()

    # Step 4: Test MAPI namespace
    print("Step 4: Testing MAPI namespace...")
    try:
        namespace = outlook.GetNamespace("MAPI")
        print("✓ MAPI namespace accessible")
    except Exception as e:
        print(f"✗ Cannot access MAPI: {e}")
        input("\nPress Enter to exit...")
        return False
    print()

    # Step 5: Test folder access
    print("Step 5: Testing folder access...")
    try:
        inbox = namespace.GetDefaultFolder(6)  # 6 = Inbox
        print(f"✓ Accessed folder: {inbox.Name}")
        print(f"  Items in folder: {inbox.Items.Count}")
    except Exception as e:
        print(f"✗ Cannot access folders: {e}")
        print("\nFix: Make sure Outlook has an email account configured")
        input("\nPress Enter to exit...")
        return False
    print()

    # Step 6: List available folders
    print("Step 6: Listing available folders...")
    try:
        folders = {
            "Inbox": namespace.GetDefaultFolder(6),
            "Sent Items": namespace.GetDefaultFolder(5),
            "Drafts": namespace.GetDefaultFolder(16),
        }
        for name, folder in folders.items():
            try:
                print(f"  ✓ {name}: {folder.Items.Count} items")
            except:
                print(f"  ⚠ {name}: Cannot access")
    except Exception as e:
        print(f"⚠ Warning: {e}")
    print()

    # Success
    print("="*60)
    print("  ✓✓✓ ALL TESTS PASSED! ✓✓✓")
    print("="*60)
    print()
    print("Your system is ready to use the Outlook Backup Tool!")
    print()
    print("Next steps:")
    print("1. Close this window")
    print("2. Run 'run_backup_tool.bat' or 'python main.py'")
    print("3. Select your backup location and filters")
    print("4. Click 'Start Backup'")
    print()

    input("Press Enter to exit...")
    return True


if __name__ == "__main__":
    try:
        test_connection()
    except KeyboardInterrupt:
        print("\n\nTest cancelled by user.")
        sys.exit(0)
    except Exception as e:
        print(f"\n\nUnexpected error: {e}")
        import traceback
        traceback.print_exc()
        input("\nPress Enter to exit...")
        sys.exit(1)
