"""
Outlook Email Backup Tool - Main GUI Application
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry
from datetime import datetime, timedelta
import threading
import json
import os
import pythoncom

from outlook_connector import OutlookConnector
from filter_engine import FilterEngine
from email_exporter import EmailExporter


class OutlookBackupApp:
    """Main application class for Outlook Email Backup Tool"""

    def __init__(self, root):
        self.root = root
        self.root.title("Outlook Email Backup Tool")
        self.root.geometry("800x700")
        self.root.minsize(700, 600)

        # Application state
        self.outlook = OutlookConnector()
        self.config_file = "config.json"
        self.backup_location = ""
        self.is_backing_up = False
        self.current_thread = None

        # Load configuration
        self.load_config()

        # Initialize GUI
        self.create_widgets()

        # Try to connect to Outlook
        self.initialize_outlook()

    def create_widgets(self):
        """Create all GUI widgets"""

        # Main container with padding
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)

        # Title
        title_label = ttk.Label(main_frame, text="Outlook Email Backup Tool",
                                font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, pady=(0, 5))

        # copyright
        copyright_label = ttk.Label(main_frame, text="© Kelvin Chong 2026",
                                      font=('Arial', 10))
        copyright_label.grid(row=0, column=2, pady=(0, 5), sticky=tk.E)

        # Compatibility Notice
        compat_label = ttk.Label(main_frame,
                                 text="⚠️ Requires Classic Outlook (Desktop Version) - New Outlook Not Supported",
                                 font=('Arial', 8),
                                 foreground='#d35400')
        compat_label.grid(row=1, column=0, pady=(0, 15))

        # Filter Options Section
        self.create_filter_section(main_frame)

        # Output Settings Section
        self.create_output_section(main_frame)

        # Action Buttons
        self.create_action_buttons(main_frame)

        # Progress Section
        self.create_progress_section(main_frame)

    def create_filter_section(self, parent):
        """Create filter options section"""
        filter_frame = ttk.LabelFrame(parent, text="Filter Options", padding="10")
        filter_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        filter_frame.columnconfigure(1, weight=1)

        row = 0

        # Date Range Filter
        self.date_filter_var = tk.BooleanVar(value=True)
        date_check = ttk.Checkbutton(filter_frame, text="Date Range:",
                                      variable=self.date_filter_var,
                                      command=self.toggle_date_filter)
        date_check.grid(row=row, column=0, sticky=tk.W, pady=5)

        date_frame = ttk.Frame(filter_frame)
        date_frame.grid(row=row, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        ttk.Label(date_frame, text="From:").pack(side=tk.LEFT, padx=(0, 5))

        # Default: 3 months ago
        default_from = datetime.now() - timedelta(days=90)
        self.date_from = DateEntry(date_frame, width=12, background='darkblue',
                                    foreground='white', borderwidth=2,
                                    date_pattern='yyyy-mm-dd',
                                    year=default_from.year,
                                    month=default_from.month,
                                    day=default_from.day)
        self.date_from.pack(side=tk.LEFT, padx=(0, 15))

        ttk.Label(date_frame, text="To:").pack(side=tk.LEFT, padx=(0, 5))

        self.date_to = DateEntry(date_frame, width=12, background='darkblue',
                                  foreground='white', borderwidth=2,
                                  date_pattern='yyyy-mm-dd')
        self.date_to.pack(side=tk.LEFT)

        row += 1

        # Sender Filter
        self.sender_filter_var = tk.BooleanVar(value=False)
        sender_check = ttk.Checkbutton(filter_frame, text="Sender:",
                                        variable=self.sender_filter_var,
                                        command=self.toggle_sender_filter)
        sender_check.grid(row=row, column=0, sticky=tk.W, pady=5)

        self.sender_entry = ttk.Entry(filter_frame, width=40, state='disabled')
        self.sender_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5, padx=(0, 5))

        ttk.Label(filter_frame, text="(email address)").grid(row=row, column=2, sticky=tk.W)

        row += 1

        # Subject Filter
        self.subject_filter_var = tk.BooleanVar(value=False)
        subject_check = ttk.Checkbutton(filter_frame, text="Subject:",
                                         variable=self.subject_filter_var,
                                         command=self.toggle_subject_filter)
        subject_check.grid(row=row, column=0, sticky=tk.W, pady=5)

        self.subject_entry = ttk.Entry(filter_frame, width=40, state='disabled')
        self.subject_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5, padx=(0, 5))

        ttk.Label(filter_frame, text="(keywords)").grid(row=row, column=2, sticky=tk.W)

        row += 1

        # Folder Filter
        ttk.Label(filter_frame, text="Folder:").grid(row=row, column=0, sticky=tk.W, pady=5)

        folder_frame = ttk.Frame(filter_frame)
        folder_frame.grid(row=row, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        self.folder_var = tk.StringVar(value="Inbox")
        self.folder_combo = ttk.Combobox(folder_frame, textvariable=self.folder_var,
                                          width=35, state='readonly')
        self.folder_combo['values'] = ["Inbox", "Sent Items"]
        self.folder_combo.pack(side=tk.LEFT, padx=(0, 5))

        self.browse_folder_btn = ttk.Button(folder_frame, text="Browse Folders...",
                                             command=self.browse_folders)
        self.browse_folder_btn.pack(side=tk.LEFT)

    def create_output_section(self, parent):
        """Create output settings section"""
        output_frame = ttk.LabelFrame(parent, text="Output Settings", padding="10")
        output_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        output_frame.columnconfigure(1, weight=1)

        # Backup Location
        ttk.Label(output_frame, text="Backup Location:").grid(row=0, column=0,
                                                                sticky=tk.W, pady=5)

        location_frame = ttk.Frame(output_frame)
        location_frame.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5)
        location_frame.columnconfigure(0, weight=1)

        self.location_var = tk.StringVar(value=self.backup_location)
        self.location_entry = ttk.Entry(location_frame, textvariable=self.location_var,
                                         state='readonly')
        self.location_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))

        self.browse_btn = ttk.Button(location_frame, text="Browse...",
                                      command=self.browse_location)
        self.browse_btn.grid(row=0, column=1)

        # Checkboxes
        self.include_attachments_var = tk.BooleanVar(value=True)
        attach_check = ttk.Checkbutton(output_frame, text="Include Attachments",
                                        variable=self.include_attachments_var)
        attach_check.grid(row=1, column=1, sticky=tk.W, pady=5)

        self.organize_by_date_var = tk.BooleanVar(value=True)
        organize_check = ttk.Checkbutton(output_frame,
                                          text="Organize by Date (YYYY/MM folder structure)",
                                          variable=self.organize_by_date_var)
        organize_check.grid(row=2, column=1, sticky=tk.W, pady=5)

    def create_action_buttons(self, parent):
        """Create action buttons"""
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=4, column=0, pady=10)

        self.preview_btn = ttk.Button(button_frame, text="Preview Count",
                                       command=self.preview_count, width=15)
        self.preview_btn.pack(side=tk.LEFT, padx=5)

        self.backup_btn = ttk.Button(button_frame, text="Start Backup",
                                      command=self.start_backup, width=15)
        self.backup_btn.pack(side=tk.LEFT, padx=5)

        self.cancel_btn = ttk.Button(button_frame, text="Cancel",
                                      command=self.cancel_operation, width=15)
        self.cancel_btn.pack(side=tk.LEFT, padx=5)

    def create_progress_section(self, parent):
        """Create progress display section"""
        progress_frame = ttk.LabelFrame(parent, text="Progress", padding="10")
        progress_frame.grid(row=5, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        parent.rowconfigure(5, weight=1)
        progress_frame.columnconfigure(0, weight=1)
        progress_frame.rowconfigure(1, weight=1)

        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var,
                                             maximum=100, length=300)
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

        # Status text area
        self.status_text = tk.Text(progress_frame, height=10, wrap=tk.WORD,
                                    state='disabled', bg='#f0f0f0')
        self.status_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Scrollbar for status text
        scrollbar = ttk.Scrollbar(progress_frame, command=self.status_text.yview)
        scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
        self.status_text['yscrollcommand'] = scrollbar.set

    def initialize_outlook(self):
        """Initialize connection to Outlook"""
        # Check Outlook version compatibility first
        version_info = self.outlook.get_outlook_version_info()

        # Show warning if new Outlook is detected
        if version_info['warning_message']:
            self.log_status("⚠️ Compatibility Warning")
            self.log_status(version_info['warning_message'])

            # Show dialog with warning
            result = messagebox.showwarning(
                "Classic Outlook Required",
                version_info['warning_message'] + "\n\n"
                "This application will attempt to connect anyway.\n"
                "If connection fails, please switch to Classic Outlook.",
                icon='warning'
            )

        try:
            self.log_status("Connecting to Outlook...")
            self.outlook.connect()
            self.log_status("✓ Connected to Classic Outlook successfully!")

            # Load folder list
            self.load_folder_list()

        except Exception as e:
            error_msg = str(e)
            self.log_status(f"✗ Error connecting to Outlook: {error_msg}")

            # Show detailed error dialog
            if "new Outlook" in error_msg or "CLASSIC" in error_msg:
                # Show special dialog for new Outlook issue
                messagebox.showerror(
                    "Classic Outlook Required",
                    f"{error_msg}\n\n"
                    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
                    "HOW TO SWITCH TO CLASSIC OUTLOOK:\n\n"
                    "1. Open Outlook\n"
                    "2. Look for a toggle switch labeled\n"
                    "   'Try the new Outlook' in the top-right\n"
                    "3. Turn OFF the toggle\n"
                    "4. Outlook will restart in Classic mode\n"
                    "5. Run this tool again\n\n"
                    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
                )
            else:
                # Show standard error dialog
                messagebox.showerror(
                    "Connection Error",
                    f"Failed to connect to Outlook:\n\n{error_msg}\n\n"
                    "Please make sure Classic Outlook is running."
                )

    def load_folder_list(self):
        """Load list of Outlook folders"""
        try:
            default_folders = self.outlook.get_default_folders()
            folder_names = list(default_folders.keys())
            self.folder_combo['values'] = folder_names
        except Exception as e:
            self.log_status(f"Warning: Could not load folder list: {str(e)}")

    def browse_folders(self):
        """Open dialog to browse all Outlook folders"""
        try:
            folders = self.outlook.get_all_folder_names()
            if folders:
                # Create a simple selection dialog
                dialog = tk.Toplevel(self.root)
                dialog.title("Select Outlook Folder")
                dialog.geometry("400x500")

                ttk.Label(dialog, text="Select a folder:").pack(pady=10)

                # Listbox with scrollbar
                list_frame = ttk.Frame(dialog)
                list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

                scrollbar = ttk.Scrollbar(list_frame)
                scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

                folder_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set)
                folder_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
                scrollbar.config(command=folder_listbox.yview)

                for folder in folders:
                    folder_listbox.insert(tk.END, folder)

                def on_select():
                    selection = folder_listbox.curselection()
                    if selection:
                        selected_folder = folder_listbox.get(selection[0])
                        self.folder_var.set(selected_folder)
                        dialog.destroy()

                ttk.Button(dialog, text="Select", command=on_select).pack(pady=10)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to browse folders:\n{str(e)}")

    def browse_location(self):
        """Browse for backup location"""
        folder = filedialog.askdirectory(title="Select Backup Location",
                                          initialdir=self.backup_location)
        if folder:
            self.backup_location = folder
            self.location_var.set(folder)
            self.save_config()

    def toggle_date_filter(self):
        """Toggle date filter controls"""
        state = 'normal' if self.date_filter_var.get() else 'disabled'
        self.date_from.config(state=state)
        self.date_to.config(state=state)

    def toggle_sender_filter(self):
        """Toggle sender filter controls"""
        state = 'normal' if self.sender_filter_var.get() else 'disabled'
        self.sender_entry.config(state=state)

    def toggle_subject_filter(self):
        """Toggle subject filter controls"""
        state = 'normal' if self.subject_filter_var.get() else 'disabled'
        self.subject_entry.config(state=state)

    def get_active_filters(self):
        """Build filter dictionary from GUI"""
        filters = {}

        if self.date_filter_var.get():
            date_from = self.date_from.get_date()
            date_to = self.date_to.get_date()

            # Set time to start/end of day
            filters['date_from'] = datetime.combine(date_from, datetime.min.time())
            filters['date_to'] = datetime.combine(date_to, datetime.max.time())

        if self.sender_filter_var.get() and self.sender_entry.get().strip():
            filters['sender'] = self.sender_entry.get().strip()

        if self.subject_filter_var.get() and self.subject_entry.get().strip():
            filters['subject'] = self.subject_entry.get().strip()

        return filters

    def preview_count(self):
        """Preview how many emails match the filters"""
        if self.is_backing_up:
            messagebox.showwarning("Busy", "Backup is already in progress.")
            return

        self.log_status("\n--- Preview Count ---")
        self.log_status("Counting emails...")

        def count_thread():
            # Initialize COM for this thread
            pythoncom.CoInitialize()

            try:
                # Create new Outlook connection for this thread
                outlook = OutlookConnector()
                outlook.connect()

                # Get selected folder
                folder_name = self.folder_var.get()
                default_folders = outlook.get_default_folders()

                if folder_name in default_folders:
                    folder = default_folders[folder_name]
                else:
                    folder = outlook.get_folder_by_path(folder_name)

                # Get all emails
                self.log_status(f"Retrieving emails from '{folder_name}'...")
                emails = outlook.get_emails_from_folder(folder)

                # Apply filters
                filters = self.get_active_filters()
                if filters:
                    filter_summary = FilterEngine.get_filter_summary(filters)
                    self.log_status(f"Applying filters: {filter_summary}")
                    filtered_emails = FilterEngine.apply_filters(emails, filters)
                else:
                    self.log_status("No filters applied")
                    filtered_emails = emails

                # Display count
                self.log_status(f"\nTotal emails in folder: {len(emails)}")
                self.log_status(f"Emails matching filters: {len(filtered_emails)}")

                messagebox.showinfo("Preview Count",
                                    f"Found {len(filtered_emails)} emails matching your criteria.\n\n"
                                    f"Total in folder: {len(emails)}")

            except Exception as e:
                self.log_status(f"Error during preview: {str(e)}")
                messagebox.showerror("Preview Error", f"Failed to count emails:\n{str(e)}")

            finally:
                # Uninitialize COM for this thread
                pythoncom.CoUninitialize()

        # Run in background thread
        thread = threading.Thread(target=count_thread, daemon=True)
        thread.start()

    def start_backup(self):
        """Start the backup process"""
        if self.is_backing_up:
            messagebox.showwarning("Busy", "Backup is already in progress.")
            return

        # Validate backup location
        if not self.backup_location or not os.path.exists(self.backup_location):
            messagebox.showerror("Invalid Location",
                                 "Please select a valid backup location.")
            self.browse_location()
            return

        # Confirm with user
        filters = self.get_active_filters()
        filter_summary = FilterEngine.get_filter_summary(filters) if filters else "No filters"

        message = (f"Ready to start backup:\n\n"
                   f"Folder: {self.folder_var.get()}\n"
                   f"Filters: {filter_summary}\n"
                   f"Location: {self.backup_location}\n"
                   f"Include Attachments: {'Yes' if self.include_attachments_var.get() else 'No'}\n"
                   f"Organize by Date: {'Yes' if self.organize_by_date_var.get() else 'No'}\n\n"
                   f"Continue?")

        if not messagebox.askyesno("Confirm Backup", message):
            return

        # Start backup in background thread
        self.is_backing_up = True
        self.backup_btn.config(state='disabled')
        self.preview_btn.config(state='disabled')
        self.progress_var.set(0)

        self.log_status("\n--- Starting Backup ---")

        self.current_thread = threading.Thread(target=self.backup_thread, daemon=True)
        self.current_thread.start()

    def backup_thread(self):
        """Background thread for backup process"""
        # Initialize COM for this thread
        pythoncom.CoInitialize()

        try:
            # Create new Outlook connection for this thread
            outlook = OutlookConnector()
            outlook.connect()

            # Get selected folder
            folder_name = self.folder_var.get()
            default_folders = outlook.get_default_folders()

            if folder_name in default_folders:
                folder = default_folders[folder_name]
            else:
                folder = outlook.get_folder_by_path(folder_name)

            # Get all emails
            self.log_status(f"Retrieving emails from '{folder_name}'...")
            emails = outlook.get_emails_from_folder(folder)
            self.log_status(f"Retrieved {len(emails)} emails")

            # Apply filters
            filters = self.get_active_filters()
            if filters:
                filter_summary = FilterEngine.get_filter_summary(filters)
                self.log_status(f"Applying filters: {filter_summary}")
                filtered_emails = FilterEngine.apply_filters(emails, filters)
                self.log_status(f"Filtered to {len(filtered_emails)} emails")
            else:
                filtered_emails = emails

            if len(filtered_emails) == 0:
                self.log_status("No emails match the criteria. Nothing to backup.")
                messagebox.showinfo("No Emails", "No emails match the specified criteria.")
                return

            # Initialize exporter
            exporter = EmailExporter(
                self.backup_location,
                organize_by_date=self.organize_by_date_var.get(),
                include_attachments=self.include_attachments_var.get(),
                sender_filter_enabled=self.sender_filter_var.get(),
                subject_filter_enabled=self.subject_filter_var.get()
            )

            # Export emails
            total = len(filtered_emails)
            self.log_status(f"\nStarting export of {total} emails...")

            for i, email in enumerate(filtered_emails):
                if not self.is_backing_up:
                    self.log_status("\nBackup cancelled by user.")
                    break

                # Update progress
                progress = ((i + 1) / total) * 100
                self.progress_var.set(progress)

                # Export email
                success, message = exporter.export_email(
                    email,
                    progress_callback=self.log_status
                )

                # Update status
                self.log_status(f"Progress: {i + 1}/{total} emails")

            # Display summary
            summary = exporter.get_summary()
            self.log_status("\n--- Backup Complete ---")
            self.log_status(f"Emails exported: {summary['exported_count']}")
            self.log_status(f"Total size: {summary['total_size_mb']} MB")
            self.log_status(f"Location: {summary['backup_location']}")

            if summary['error_count'] > 0:
                self.log_status(f"\nErrors encountered: {summary['error_count']}")
                for error in summary['errors'][:10]:  # Show first 10 errors
                    self.log_status(f"  - {error}")

            messagebox.showinfo("Backup Complete",
                                f"Backup completed successfully!\n\n"
                                f"Emails exported: {summary['exported_count']}\n"
                                f"Total size: {summary['total_size_mb']} MB\n"
                                f"Location: {summary['backup_location']}\n"
                                f"Errors: {summary['error_count']}")

        except Exception as e:
            self.log_status(f"\nBackup failed: {str(e)}")
            messagebox.showerror("Backup Error", f"Backup failed:\n{str(e)}")

        finally:
            # Uninitialize COM for this thread
            pythoncom.CoUninitialize()

            self.is_backing_up = False
            self.backup_btn.config(state='normal')
            self.preview_btn.config(state='normal')
            self.progress_var.set(0)

    def cancel_operation(self):
        """Cancel ongoing operation or close application"""
        if self.is_backing_up:
            if messagebox.askyesno("Cancel Backup", "Are you sure you want to cancel the backup?"):
                self.is_backing_up = False
                self.log_status("\nCancelling backup...")
        else:
            self.root.quit()

    def log_status(self, message):
        """Add message to status text area"""
        self.status_text.config(state='normal')
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.status_text.config(state='disabled')
        self.root.update_idletasks()

    def load_config(self):
        """Load configuration from file"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    self.backup_location = config.get('backup_location', '')
        except Exception as e:
            print(f"Could not load config: {str(e)}")

    def save_config(self):
        """Save configuration to file"""
        try:
            config = {
                'backup_location': self.backup_location
            }
            with open(self.config_file, 'w') as f:
                json.dump(config, f, indent=2)
        except Exception as e:
            print(f"Could not save config: {str(e)}")


def main():
    """Main entry point"""
    root = tk.Tk()
    app = OutlookBackupApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
