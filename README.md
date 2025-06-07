
# Python Online Log

## 📝 Overview

This application is a **GUI-based event logger** built with **Tkinter**, used for logging operational events onboard vessels. It supports **logging events to both an Excel workbook and an SQLite database**, with customizable buttons and real-time file monitoring.

It is especially useful for survey/data acquisition operations that require precise logging of onboard actions, vessel states, or file timestamps.

---

## ⚙️ Features

- Monitor `.txt` files and folders in real-time.
- Log events to Excel and SQLite.
- Two customizable tabs for user-defined event buttons.
- Color-coded entries in Excel.
- Sync button to update SQLite based on Excel data.
- GUI configuration for log paths and settings.

---

## 🚀 How to Use

1. **Run the Script**  
   Launch the script using Python:
   ```bash
   python Online_Log_Rev014_2Tab.py
   ```

2. **Set Up Configuration**  
   Click the `Settings` button to configure:
   - Excel log file path.
   - SQLite database path.
   - Monitored folders.
   - Custom button names and associated actions.

3. **Log Events**  
   Use the provided buttons:
   - `Log on`, `Log off`, `Event`, `New Day`, `SVP`, and your own custom buttons.
   - Right-click tabs to **add/edit/remove** custom buttons.

4. **Sync Data**  
   Use `Sync Excel->DB` to update the SQLite database with the latest Excel data.

5. **Status Feedback**  
   The bottom bar and labels display current monitoring and SQLite status.

---

## 📦 Requirements

- Python 3.x
- Required packages:
  ```
  pip install xlwings pandas watchdog
  ```

- Microsoft Excel installed (for `xlwings` to work properly)

---

## 🛟 Notes

- All button presses log data from the latest `.txt` file and add it to the Excel log and/or SQLite.
- Uses a JSON config file (`logger_settings.json`) to store settings persistently.
- Designed for stability on field operations with auto-recovery for most common errors.

## ✍🏻 Authors
- Program developed by Pierre Lowe with contributions of Vicente Danvila