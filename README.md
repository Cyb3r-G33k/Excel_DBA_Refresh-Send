
# ExcelCourier  
*Created by Alex Losev*

## Introduction

**ExcelCourier** is a Python-based application designed to monitor and refresh Excel files automatically. It converts the refreshed data into CSV format and sends the updated files via email to specified recipients.
## Features

- **Excel File Monitoring:** Continuously monitor selected Excel files for updates.  
- **Automatic Refresh:** Refresh Excel content, including data connections or pivot tables.  
- **CSV Conversion:** Save the refreshed data as CSV files in a designated folder.  
- **Email Notification:** Send updated CSV files via email to configured recipients.  
- **Configurable Frequency:** Set how often files are monitored and emails are sent.  
- **Multi-file Support:** Monitor multiple Excel files simultaneously for changes.

---

## Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/YourGitHubAccount/ExcelCourier.git
   cd ExcelCourier
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application**
   ```bash
   python ExcelCourier.py
   ```

---

## Usage

1. **Launch the application:**  
   Run `ExcelCourier.py` to open the tool.

2. **Select Excel Files to Monitor:**  
   Use the interface to select the Excel files you want to monitor.

3. **Set Monitoring Frequency:**  
   Configure the time interval for checking updates (e.g., every 5, 10, or 30 minutes).

4. **Edit Email Credentials:**  
   Open the code and update the following variables with your email and password:
   ```python
   fromaddr = 'your_email@gmail.com'
   password = 'your_password'
   ```
   **Note:** Ensure Gmail's "Less secure apps" access is enabled or generate an App Password.

5. **Start Monitoring:**  
   Click **Start** to begin monitoring. If changes are detected, the data will be saved as a CSV file.

6. **Receive Email Notifications:**  
   Updated CSV files are automatically sent to recipients via email when changes are detected.

---

## Configuration

1. **Settings File:**  
   A `config.ini` file is generated on the first run. Use it to:
   - Configure email settings (SMTP server, sender, and recipients).
   - Set the frequency for monitoring updates.
   - Define where CSV files should be saved.

2. **Folder Selection:**  
   Choose the output folder where CSV files will be stored.

---

## Logs and Reports

- **Logs:**  
  Monitoring events and errors are logged in the `/logs` directory.

- **Reports:**  
  Updated CSV files are saved in the `/reports` folder and emailed to recipients.

---

## Development Roadmap

- [x] Monitor and refresh Excel files  
- [x] Convert data to CSV and send via email  
- [ ] Support for Excel macros and advanced formatting  
- [ ] Cloud integration for distributed file management  
- [ ] Add visual dashboards for analytics

---

## Contributing

We welcome contributions from the community! Follow these steps to contribute:

1. Fork the repository.  
2. Create a new branch:
   ```bash
   git checkout -b feature-branch
   ```
3. Commit your changes:
   ```bash
   git commit -m "Added new feature"
   ```
4. Push to the branch:
   ```bash
   git push origin feature-branch
   ```
5. Create a pull request.

---

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
