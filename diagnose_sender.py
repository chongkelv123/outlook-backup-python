"""
Diagnostic script to check sender properties in Outlook emails
"""

import win32com.client
import pythoncom

def diagnose_sender_properties():
    """Check what sender properties are available in emails"""

    pythoncom.CoInitialize()

    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        # Get Inbox
        inbox = namespace.GetDefaultFolder(6)

        # Get first 5 emails
        emails = []
        count = 0
        for item in inbox.Items:
            if item.Class == 43:  # olMail
                emails.append(item)
                count += 1
                if count >= 5:
                    break

        print(f"Found {len(emails)} emails to analyze\n")
        print("=" * 80)

        for i, email in enumerate(emails, 1):
            print(f"\nEmail #{i}")
            print("-" * 80)

            # Check Subject
            try:
                print(f"Subject: {email.Subject[:50]}...")
            except:
                print("Subject: N/A")

            # Check SenderEmailAddress
            try:
                print(f"SenderEmailAddress: {email.SenderEmailAddress}")
            except Exception as e:
                print(f"SenderEmailAddress: ERROR - {e}")

            # Check SenderName
            try:
                print(f"SenderName: {email.SenderName}")
            except Exception as e:
                print(f"SenderName: ERROR - {e}")

            # Check Sender object
            try:
                if email.Sender:
                    print(f"Sender.Name: {email.Sender.Name}")
                    print(f"Sender.Address: {email.Sender.Address}")

                    # Try to get SMTP address from Exchange User
                    try:
                        exchange_user = email.Sender.GetExchangeUser()
                        if exchange_user:
                            print(f"Sender.GetExchangeUser().PrimarySmtpAddress: {exchange_user.PrimarySmtpAddress}")
                    except Exception as e:
                        print(f"Sender.GetExchangeUser(): ERROR - {e}")

                    # Try to get SMTP address from Exchange Distribution List
                    try:
                        exchange_dl = email.Sender.GetExchangeDistributionList()
                        if exchange_dl:
                            print(f"Sender.GetExchangeDistributionList().PrimarySmtpAddress: {exchange_dl.PrimarySmtpAddress}")
                    except Exception as e:
                        print(f"Sender.GetExchangeDistributionList(): ERROR - {e}")
            except Exception as e:
                print(f"Sender: ERROR - {e}")

            # Check SenderEmailType
            try:
                print(f"SenderEmailType: {email.SenderEmailType}")
            except Exception as e:
                print(f"SenderEmailType: ERROR - {e}")

            print("-" * 80)

        print("\n" + "=" * 80)
        print("Diagnosis complete!")

    except Exception as e:
        print(f"Error: {e}")

    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    diagnose_sender_properties()
