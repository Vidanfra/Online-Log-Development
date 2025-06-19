
# Python Online Log

## 📝 Overview

This application is a **GUI-based event logger** built with **Tkinter**, used for logging operational events onboard vessels. It supports **logging events to both an Excel workbook and an SQLite database**, with customizable buttons and real-time file monitoring.

It allows an operator to log events with a single button click, capturing dynamic data from navigation systems (via TXT files) and static data from a project-specific Excel workbook.

Logged events are simultaneously written to a daily Excel log and inserted into a robust SQLite database. The application features a synchronization engine to reconcile the Excel log with the database, ensuring data integrity even when unique identifiers are missing or have been duplicated.

---

## ⚙️ Features

- Monitor `.txt` files and folders in real-time.
- Log events to Excel and SQLite in real-time.
- Customizable user-defined event buttons organized in tabs.
- Color-coded entries in Excel.
- Sync button to update SQLite based on Excel data.
- GUI configuration for log paths and settings.
- Capture of static data from Log workbook directily from cells using Excel commands (e.g., ='Settings'!B2)

---

## 🚀 How to Use

1. **Run the Script**  
   Launch the script using Python:
   ```bash
   python Online_Log_16.py
   ```

2. **Set Up Configuration**  
   Click the `Settings` button to configure:
   - Excel log and navigation file paths.
   - SQLite database path.
   - Monitored folders.
   - Custom button names and associated actions.
   - Custom event codes

3. **Log Events**  
   Use the provided buttons:
   - Default layout of custom buttons based on the FLA Layout of the project 600013 Rheinmetall.
   - Right-click tabs to **add/edit/remove** custom buttons or edit in the Settings menu.

4. **Sync Data**  
   Use `Sync Excel->DB` to update the SQLite database with the latest Excel data.

5. **Status Feedback**  
   The bottom bar and labels display current monitoring and SQLite status.

---

## 📦 Requirements

- Python 3.13 (or newer)
- Required packages:
```bash
 pip install pandas openpyxl pyxlsb xlwings watchdog
```

- Microsoft Excel installed (for `xlwings`and `pyxlsb` to work properly)

---

## 🛟 Notes

- All button presses log data from the latest `.txt` file and add it to the Excel log and/or SQLite.
- Uses a JSON config file (`logger_settings.json`) to store settings persistently.
- Designed for stability on field operations with auto-recovery for most common errors.

## ✍🏻 Authors
- Program developed by Pierre Lowe and Vicente Danvila