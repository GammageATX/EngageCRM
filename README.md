# EngageCRM

A desktop application for tracking engagements, organizations, personnel, projects, and reviews.

## Prerequisites

1. Windows 10 or later
2. Python 3.12 or later (SQLite is included with Python)

## Installation Steps

1. Install Python 3.12:
   - Download from: [Python.org](https://www.python.org/downloads/)
   - During installation, make sure to check "Add Python to PATH"
   - Complete the installation

2. Copy these files to your work computer:
   - `tracker.py`
   - `requirements.txt`

3. Open Command Prompt or PowerShell:
   - Press Win + R
   - Type `cmd` or `powershell` and press Enter
   - Navigate to the folder containing the files:

     ```powershell
     cd path\to\your\folder
     ```

4. Create and activate a virtual environment (recommended):

   ```powershell
   python -m venv venv
   .\venv\Scripts\activate
   ```

5. Install required packages:

   ```powershell
   python -m pip install --upgrade pip
   pip install -r requirements.txt
   ```

## Running the Application

1. Make sure you're in the correct directory and the virtual environment is activated
2. Run the application:

   ```powershell
   python tracker.py
   ```

The application will create a new SQLite database file (`engagement_tracker.db`) in the same directory if it doesn't exist.

## Features

- Track organizations and their details
- Manage personnel information
- Create and monitor projects
- Log engagements and meetings
- Create periodic reviews
- Generate reports and export to Excel

## Troubleshooting

1. If you get an error about missing DLLs:
   - Verify that Python is added to your PATH
   - Try restarting your computer after Python installation

2. If you get permission errors:
   - Run Command Prompt or PowerShell as Administrator
   - Check that you have write permissions in the application directory

## Support

For any issues or questions, please contact your system administrator or the development team.
