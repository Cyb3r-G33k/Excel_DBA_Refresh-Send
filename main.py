import os
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from tkinter import Tk, filedialog, StringVar, Entry, Button, Label, OptionMenu
from watchdog.observers import Observer
from threading import Thread, Event
import time
import win32com.client  # For Excel refresh automation on Windows


class FileMonitorHandler:
    def __init__(self, monitored_files, email_list, save_folder, frequency, stop_event):
        self.monitored_files = monitored_files
        self.email_list = email_list
        self.save_folder = save_folder
        self.frequency = frequency
        self.stop_event = stop_event
        self.file_hashes = {}

    def refresh_and_check_changes(self):
        while not self.stop_event.is_set():
            for file_path in self.monitored_files:
                try:
                    # Automatically refresh the Excel file and save
                    self.refresh_excel_file(file_path)

                    current_hash = self.hash_file(file_path)
                    if file_path not in self.file_hashes or self.file_hashes[file_path] != current_hash:
                        print(f"Changes detected in {file_path}, processing...")
                        self.process_file(file_path)
                        self.file_hashes[file_path] = current_hash
                    else:
                        print(f"No changes detected in {file_path}.")
                except Exception as e:
                    print(f"Error processing file {file_path}: {e}")

            time.sleep(self.frequency * 60)

    def refresh_excel_file(self, file_path):
        """Refresh Excel file content automatically using win32com (Windows-only)."""
        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.Visible = False
        workbook = excel_app.Workbooks.Open(file_path)
        workbook.RefreshAll()  # Refresh any data connections or pivot tables
        workbook.Save()  # Save the refreshed data back to the file
        workbook.Close(SaveChanges=True)
        excel_app.Quit()

    def process_file(self, file_path):
        try:
            # Load Excel file and convert to CSV
            df = pd.read_excel(file_path)
            output_csv = os.path.join(self.save_folder, os.path.basename(file_path).replace('.xlsx', '.csv'))
            df.to_csv(output_csv, index=False)
            print(f"Converted {file_path} to {output_csv}")
            self.send_email(output_csv)
        except Exception as e:
            print(f"Error processing file {file_path}: {e}")

    def send_email(self, file_path):
        try:
            fromaddr = 'your_email@gmail.com'
            password = 'your_password'

            for email in self.email_list:
                msg = MIMEMultipart()
                msg['From'] = fromaddr
                msg['To'] = email.strip()
                msg['Subject'] = 'Monitored CSV File'

                body = 'Attached is the updated CSV file from the monitored Excel file.'
                msg.attach(MIMEText(body, 'plain'))

                # Attach the CSV file
                attachment = open(file_path, 'rb')
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(file_path)}")
                msg.attach(part)

                # Send email
                server = smtplib.SMTP('smtp.gmail.com', 587)
                server.starttls()
                server.login(fromaddr, password)
                text = msg.as_string()
                server.sendmail(fromaddr, email.strip(), text)
                server.quit()

                print(f"CSV file sent to {email}")
        except Exception as e:
            print(f"Failed to send email: {e}")

    def hash_file(self, file_path):
        """Generate a hash of the file to detect changes."""
        return hash(open(file_path, 'rb').read())


class ExcelMonitorApp:
    def __init__(self):
        self.root = Tk()
        self.root.title("Excel Monitor to CSV")

        # Labels
        self.label_folder = Label(self.root, text="Select Folder to Monitor:")
        self.label_folder.pack()
        self.label_files = Label(self.root, text="Select Excel Files to Monitor:")
        self.label_files.pack()

        self.email_label = Label(self.root, text="Enter Email Addresses (comma separated):")
        self.email_label.pack()

        # Entry for email input
        self.email_var = StringVar()
        self.email_entry = Entry(self.root, textvariable=self.email_var, width=50)
        self.email_entry.pack()

        self.save_folder_label = Label(self.root, text="Select Folder to Save CSV Files:")
        self.save_folder_label.pack()

        self.frequency_label = Label(self.root, text="Select Frequency (in minutes):")
        self.frequency_label.pack()

        # Dropdown for frequency selection
        self.frequency_var = StringVar(value="5")
        self.frequency_dropdown = OptionMenu(self.root, self.frequency_var, "1", "5", "10", "30", "60")
        self.frequency_dropdown.pack()

        # Buttons for selecting folders and files
        self.button_select_folder = Button(self.root, text="Select Folder to Monitor", command=self.select_folder)
        self.button_select_folder.pack()
        self.button_select_files = Button(self.root, text="Select Files to Monitor", command=self.select_files)
        self.button_select_files.pack()
        self.button_select_save_folder = Button(self.root, text="Select Folder to Save CSV",
                                                command=self.select_save_folder)
        self.button_select_save_folder.pack()
        self.button_start_monitoring = Button(self.root, text="Start Monitoring", command=self.start_monitoring)
        self.button_start_monitoring.pack()

        # Variables
        self.folder_path = ""
        self.selected_files = []
        self.save_folder = ""
        self.monitor_thread = None
        self.stop_event = Event()

        self.root.mainloop()

    def select_folder(self):
        self.folder_path = filedialog.askdirectory()
        print(f"Selected folder to monitor: {self.folder_path}")

    def select_files(self):
        self.selected_files = filedialog.askopenfilenames(title="Select Excel Files",
                                                          filetypes=[("Excel files", "*.xlsx")])
        print(f"Selected files: {self.selected_files}")

    def select_save_folder(self):
        self.save_folder = filedialog.askdirectory()
        print(f"Selected folder to save CSV files: {self.save_folder}")

    def start_monitoring(self):
        email_list = self.email_var.get().split(',')
        frequency = int(self.frequency_var.get())

        if not self.folder_path or not self.selected_files or not email_list or not self.save_folder:
            print("Please select a folder to monitor, files, save folder, and enter email addresses!")
            return

        # Set up the event handler
        event_handler = FileMonitorHandler(self.selected_files, email_list, self.save_folder, frequency,
                                           self.stop_event)

        # Start monitoring in a separate thread
        self.monitor_thread = Thread(target=event_handler.refresh_and_check_changes)
        self.monitor_thread.start()

        print("Monitoring started...")

        # Gracefully stop monitoring on closing
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        self.stop_event.set()
        if self.monitor_thread:
            self.monitor_thread.join()
        self.root.destroy()


# Run the application
if __name__ == '__main__':
    ExcelMonitorApp()
