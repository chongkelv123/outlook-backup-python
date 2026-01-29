"""
Outlook Connection Diagnostic Tool
Helps diagnose and fix Outlook connection issues
"""

import sys
import subprocess
import platform


def print_header(text):
    """Print formatted header"""
    print("\n" + "="*60)
    print(f"  {text}")
    print("="*60)


def print_result(test_name, passed, message=""):
    """Print test result"""
    status = "✓ PASS" if passed else "✗ FAIL"
    print(f"\n{status}: {test_name}")
    if message:
        print(f"  → {message}")


def check_python_version():
    """Check Python version"""
    print_header("Checking Python Installation")
    version = sys.version_info
    print(f"Python Version: {version.major}.{version.minor}.{version.micro}")

    if version.major >= 3 and version.minor >= 7:
        print_result("Python Version", True, "Python 3.7+ detected")
        return True
    else:
        print_result("Python Version", False, "Need Python 3.7 or higher")
        return False


def check_pywin32():
    """Check if pywin32 is installed"""
    print_header("Checking pywin32 Installation")

    try:
        import win32com.client
        import pywintypes
        print_result("pywin32 Module", True, "pywin32 is installed")
        return True
    except ImportError as e:
        print_result("pywin32 Module", False, f"pywin32 not found: {str(e)}")
        print("\n  Fix: Run 'pip install pywin32'")
        return False


def check_outlook_process():
    """Check if Outlook process is running"""
    print_header("Checking Outlook Process")

    try:
        result = subprocess.run(
            ['tasklist', '/FI', 'IMAGENAME eq OUTLOOK.EXE'],
            capture_output=True,
            text=True,
            creationflags=subprocess.CREATE_NO_WINDOW
        )

        if 'OUTLOOK.EXE' in result.stdout:
            print_result("Outlook Process", True, "OUTLOOK.EXE is running")
            return True
        else:
            print_result("Outlook Process", False, "OUTLOOK.EXE not found")
            print("\n  Fix: Start Microsoft Outlook manually")
            return False
    except Exception as e:
        print_result("Outlook Process", False, f"Cannot check: {str(e)}")
        return False


def check_outlook_installation():
    """Check if Outlook is installed"""
    print_header("Checking Outlook Installation")

    try:
        import winreg

        # Check registry for Outlook installation
        paths_to_check = [
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Office\16.0\Outlook"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Office\15.0\Outlook"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Office\14.0\Outlook"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Wow6432Node\Microsoft\Office\16.0\Outlook"),
        ]

        found = False
        for hkey, path in paths_to_check:
            try:
                key = winreg.OpenKey(hkey, path)
                winreg.CloseKey(key)
                print_result("Outlook Installation", True, f"Found at: {path}")
                found = True
                break
            except WindowsError:
                continue

        if not found:
            print_result("Outlook Installation", False, "Cannot find Outlook in registry")
            print("\n  Fix: Install Microsoft Office/Outlook")
            return False

        return True
    except Exception as e:
        print_result("Outlook Installation", False, f"Cannot verify: {str(e)}")
        return False


def test_outlook_connection():
    """Test actual connection to Outlook"""
    print_header("Testing Outlook Connection")

    try:
        import win32com.client
        import pywintypes

        # Method 1: Try GetActiveObject
        try:
            outlook = win32com.client.GetActiveObject("Outlook.Application")
            print_result("GetActiveObject", True, "Connected to existing Outlook instance")

            # Try to access namespace
            namespace = outlook.GetNamespace("MAPI")
            print_result("MAPI Namespace", True, "Successfully accessed MAPI")

            # Try to access Inbox
            inbox = namespace.GetDefaultFolder(6)
            folder_name = inbox.Name
            item_count = inbox.Items.Count
            print_result("Access Inbox", True, f"Folder: {folder_name}, Items: {item_count}")

            return True

        except pywintypes.com_error as e:
            error_code = e.args[0] if e.args else "Unknown"
            print_result("GetActiveObject", False, f"Error code: {error_code}")

            # Explain common error codes
            if error_code == -2146959355:
                print("\n  Error -2146959355: Server execution failed")
                print("  Common causes:")
                print("    1. Outlook is not fully loaded yet")
                print("    2. Outlook is running with different permissions")
                print("    3. Windows security settings blocking COM")
                print("\n  Try:")
                print("    1. Close Outlook completely")
                print("    2. Start Outlook and wait until fully loaded")
                print("    3. Run this script as Administrator")
                print("    4. Run: python -m win32com.client.makepy Outlook")

            elif error_code == -2147221021:
                print("\n  Error -2147221021: Operation unavailable")
                print("  Outlook might not be running")

            elif error_code == -2147221005:
                print("\n  Error -2147221005: Invalid class string")
                print("  Outlook might not be properly installed")

            return False

    except ImportError:
        print_result("Import Error", False, "Cannot import required modules")
        return False
    except Exception as e:
        print_result("Connection Test", False, str(e))
        return False


def check_admin_rights():
    """Check if running as administrator"""
    print_header("Checking Administrator Rights")

    try:
        import ctypes
        is_admin = ctypes.windll.shell32.IsUserAnAdmin() != 0

        if is_admin:
            print_result("Admin Rights", True, "Running as Administrator")
        else:
            print_result("Admin Rights", False, "Not running as Administrator")
            print("\n  Note: Some COM operations may require Administrator rights")
            print("  Try: Right-click and 'Run as Administrator'")

        return is_admin
    except:
        print_result("Admin Rights", False, "Cannot determine admin status")
        return False


def provide_solutions():
    """Provide solution steps"""
    print_header("Recommended Solutions")

    print("""
Step-by-step troubleshooting:

1. START OUTLOOK FIRST
   - Open Microsoft Outlook manually
   - Wait until it's fully loaded (you can see your inbox)
   - Keep Outlook open while running the backup tool

2. RUN AS ADMINISTRATOR
   - Right-click on 'run_backup_tool.bat'
   - Select 'Run as administrator'
   - Try the backup tool again

3. REGISTER COM LIBRARIES
   - Open Command Prompt as Administrator
   - Run: python -m win32com.client.makepy Outlook
   - This creates early binding for Outlook

4. CHECK OUTLOOK PROFILE
   - Open Outlook
   - Go to: File → Account Settings → Account Settings
   - Make sure you have at least one email account configured
   - Try sending a test email to verify it works

5. REINSTALL PYWIN32
   - Open Command Prompt as Administrator
   - Run: pip uninstall pywin32
   - Run: pip install pywin32
   - Run: python Scripts/pywin32_postinstall.py -install

6. RESTART COMPUTER
   - Sometimes a simple restart resolves COM issues
   - After restart, start Outlook first, then the backup tool

7. CHECK WINDOWS SECURITY
   - Windows Defender or antivirus might block COM access
   - Temporarily disable and test
   - Add Python to exclusions if needed
""")


def main():
    """Run all diagnostic checks"""
    print("\n" + "="*60)
    print("  OUTLOOK CONNECTION DIAGNOSTIC TOOL")
    print("="*60)
    print("\nThis tool will help diagnose Outlook connection issues...")

    results = {}

    # Run all checks
    results['python'] = check_python_version()
    results['pywin32'] = check_pywin32()
    results['outlook_installed'] = check_outlook_installation()
    results['outlook_process'] = check_outlook_process()
    results['admin'] = check_admin_rights()

    if results['pywin32'] and results['outlook_process']:
        results['connection'] = test_outlook_connection()
    else:
        results['connection'] = False

    # Summary
    print_header("Diagnostic Summary")

    passed = sum(1 for v in results.values() if v)
    total = len(results)

    print(f"\nTests Passed: {passed}/{total}")

    if all(results.values()):
        print("\n✓ All checks passed!")
        print("  Your system should be able to connect to Outlook.")
        print("  If you still have issues, try restarting Outlook.")
    else:
        print("\n✗ Some checks failed.")
        print("  Please follow the recommended solutions below.")

    # Provide solutions
    provide_solutions()

    print("\n" + "="*60)
    print("  Diagnostic Complete")
    print("="*60 + "\n")

    input("Press Enter to exit...")


if __name__ == "__main__":
    main()
