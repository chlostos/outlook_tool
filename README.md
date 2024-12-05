# Outlook Tool

A Python-based script to manage Outlook emails efficiently by removing duplicate emails from a specified folder. This tool automates the cleanup process, ensuring better organization and saving time.

## Features

- **Duplicate Email Detection and Removal**: Scans through a specified Outlook folder and removes duplicate emails based on their subject, sender email address, and received time.
- **Logging**: Creates a log file (`email_cleanup.log`) to keep track of the cleanup process and any errors encountered.
- **Customizable Configuration**: Uses a `config.ini` file to specify the Outlook account and folder for email cleanup.
- **Dependency Management**: Simplifies dependency installation and project initialization using a batch script.

---

## Prerequisites

1. **Python**:
   - Requires Python 3.8 or higher.
2. **Dependencies**:
   - `win32com.client`: For interacting with the Outlook application.
   - `tqdm`: For progress bars during email processing.
   - `configparser`: For reading configuration files.
3. **Microsoft Outlook**:
   - Must be installed and configured on your system.

---

## Installation

1. **Clone the Repository**:
   ```bash
   git clone https://github.com/yourusername/outlook_tool.git
   cd outlook_tool
   ```

2. **Install Dependencies**:
   - Ensure the `requirements.txt` file includes the following:
     ```
     tqdm
     pypiwin32
     configparser
     ```
   - Run the batch script `outlook_tool.bat`:
     ```cmd
     outlook_tool.bat
     ```

---

## Configuration

Update the `config.ini` file with your Outlook account details and the target folder:

```ini
[Outlook]
account_name = your_account@example.com
folder_name = your_folder
```

- **account_name**: Your Outlook account email address.
- **folder_name**: The folder you want to clean up (default is "Inbox").

---

## Usage

1. Run the batch script:
   ```cmd
   outlook_tool.bat
   ```
   This will:
   - Check/activate the virtual environment.
   - Install necessary dependencies.
   - Start the `main.py` script for email cleanup.

2. Monitor progress via the terminal and log file (`email_cleanup.log`).

---

## Logs

All operations are logged to `email_cleanup.log`, providing detailed information on:
- Successfully deleted duplicates.
- Errors encountered during processing.

---

## Troubleshooting

- **UV Installation Failure**: If `uv` fails to install, ensure PowerShell is installed and accessible. Manually install `uv` using:
  ```powershell
  irm https://astral.sh/uv/install.ps1 | iex
  ```
- **Outlook Connection Issues**:
  - Ensure Outlook is installed and configured.
  - Verify the `account_name` in `config.ini` matches your account.

---

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.

---

## Contributing

Contributions are welcome! Please submit a pull request or report issues via the [GitHub repository](https://github.com/chlostos/outlook_tool).
