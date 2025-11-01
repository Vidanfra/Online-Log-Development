import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser, Toplevel, Label, simpledialog
import os
import xlwings as xw # Keep xlwings for Excel interaction
import threading
import time
from watchdog.observers.polling import PollingObserver
from watchdog.events import FileSystemEventHandler
import datetime
import json
import traceback
import pandas as pd
import openpyxl
import re
import asyncio
import sys
from pathlib import Path
import sqlite3
import uuid


if sys.platform == "win32":
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

#DEBUG
timings = {}
start_time = time.perf_counter()

# --- DEFINED CONSTANTS ---
# APPLICATION VERSION
APP_VERSION = "2.1"

# PATHS
# Stores the last-used project path across sessions
PROJECT_STATE_FILE = "settings/config/last_project.json"
# Path to the blank project template used when creating a new project
PROJECT_TEMPLATE_FILE = "settings/config/blank_project.json"
EVENT_CODES_FILE = "settings/event_codes.json"

# DICCTIONARY KEYS #NEEDS TO BE REVIEWED
EXCEL_LOG_REQUIRED_COLS = {'runline', 'kp', 'event'} 
DEFAULT_DATA_FIELDS = {"Date-Time", "KP", "DCC", "Line name", "Latitude", "Longitude", "Easting", "Northing", "Event", "Code", "KP Ref.", "UUID"} 
TXT_FILES_KEYS = ["None", "Main TXT", "TXT Source 2", "TXT Source 3", "TXT Source 4", "TXT Source 5"]
DEFAULT_MONITORED_FOLDERS = ["Qinsy DB", "Naviscan", "SIS", "SSS", "SBP", "Mag", "Grad", "SVP", "SpintINS", "Video", "Cathx", "Hypack RAW", "Eiva NaviPac"]


# NUMERICAL CONSTANTS
MAX_HEADER_SEARCH_ROW = 30
LAYOUT_BUTTON_COLUMNS = 5

# Global cache
folder_cache = {}

# --- Tooltip Class ---
class ToolTip:
    """
    Create a tooltip for a given widget with hover delays.
    """
    def __init__(self, widget, text, show_delay=500, hide_delay=500): # Default delays in ms
        self.widget = widget
        self.text = text
        self.show_delay = show_delay
        self.hide_delay = hide_delay
        self.tooltip_window = None
        self.show_id = None # ID for the scheduled 'after' call to show
        self.hide_id = None # ID for the scheduled 'after' call to hide
        self.last_log_on_kp = None 
        self.log_on_time = None

        # Bind events to intermediate handlers
        self.widget.bind("<Enter>", self.on_enter, add='+') # Use add='+ to coexist with button bindings
        self.widget.bind("<Leave>", self.on_leave, add='+')
        # self.widget.bind("<Destroy>", self.on_leave, add='+') # Might cause issues if triggered too often

    def on_enter(self, event=None):
        # When mouse enters, cancel any scheduled hide and schedule a show
        self.cancel_scheduled_hide()
        self.schedule_show()

    def on_leave(self, event=None):
        self.cancel_scheduled_show()
        self.schedule_hide()

    def schedule_show(self):
        self.cancel_scheduled_show()
        self.show_id = self.widget.after(self.show_delay, self.show_tooltip)

    def schedule_hide(self):
        # When mouse leaves, cancel any scheduled show and schedule a hide
        self.cancel_scheduled_show()
        # Schedule the tooltip to disappear after delay
        # Hide relatively quickly after mouse leaves
        self.hide_id = self.widget.after(max(100, self.hide_delay // 5) , self.hide_tooltip)

    def cancel_scheduled_show(self):
        if self.show_id:
            try:
                self.widget.after_cancel(self.show_id)
            except ValueError: # Ignore error if ID already invalid
                pass
            self.show_id = None

    def cancel_scheduled_hide(self):
        if self.hide_id:
            try:
                self.widget.after_cancel(self.hide_id)
            except ValueError: # Ignore error if ID already invalid
                pass
            self.hide_id = None

    def show_tooltip(self):
        # Guard against widget destruction or if it's not mapped
        if not self.widget.winfo_exists() or not self.widget.winfo_ismapped():
            self.hide_tooltip() # Ensure cleanup if widget gone
            return

        # Hide existing tooltip if somehow still visible
        self.hide_tooltip() # Call internal hide first

        # Calculate position
        try:
            x, y, _, _ = self.widget.bbox("insert")
            if x is None or y is None: x = y = 0 # Fallback
        except tk.TclError: # Handle cases where bbox fails
            x = y = 0
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20

        try:
            self.tooltip_window = Toplevel(self.widget)
            self.tooltip_window.wm_overrideredirect(True)
            self.tooltip_window.wm_attributes("-topmost", True)
            self.tooltip_window.wm_geometry(f"+{x}+{y}")

            label = Label(self.tooltip_window, text=self.text, justify='left',
                          background="#ffffe0", relief='solid', borderwidth=1,
                          font=("Arial", "9", "normal"), padx=4, pady=2)
            label.pack(ipadx=1)
            # Automatically hide after a few seconds if mouse doesn't move out
            self.hide_id = self.widget.after(5000, self.hide_tooltip)

        except tk.TclError: # Catch errors if widget destroyed during creation
            self.tooltip_window = None

    def hide_tooltip(self):
        self.cancel_scheduled_hide()
        tw = self.tooltip_window
        self.tooltip_window = None
        if tw:
            try:
                tw.destroy()
            except tk.TclError:
                pass

# --- SQLite Database Mirror ---
class SQLiteManager:
    """
    Manages SQLite database mirroring of an Excel 'LogBook' sheet.
    
    Features:
    - Full synchronization: Makes SQL DB identical to Excel sheet
    - UUID management: Validates, fixes duplicates, and generates missing UUIDs
    - Incremental updates: Adds single rows without full sync for fast logging
    - Header normalization: Replaces spaces with underscores for SQL compatibility
    """
    
    def __init__(self, db_path):
        """
        Initialize the SQLite manager with a database file path.
        
        Args:
            db_path: Path to the SQLite database file
        """
        self.db_path = db_path
        self.conn = None
        self.table_name = None  # Will be set from Excel filename
        
        try:
            self.conn = sqlite3.connect(
                db_path,
                check_same_thread=False,
                timeout=30.0
            )
            # Enable WAL mode for better concurrent access
            self.conn.execute("PRAGMA journal_mode=WAL")
            self.conn.execute("PRAGMA busy_timeout=30000")
            print(f"SQLite: Connected to database at {db_path}")
        except sqlite3.Error as e:
            print(f"SQLite: Connection error - {e}")
            raise e
    
    def close(self):
        """Close the database connection."""
        if self.conn:
            try:
                self.conn.close()
                print("SQLite: Database connection closed.")
            except:
                pass
            self.conn = None
    
    def _sanitize_column_name(self, name):
        """
        Convert Excel header to SQL-compatible column name.
        Replaces spaces and hyphens with underscores, removes invalid chars.
        
        Args:
            name: Original column name from Excel
            
        Returns:
            Sanitized column name safe for SQL
        """
        if not isinstance(name, str):
            name = str(name)
        # Replace spaces and hyphens with underscores
        name = name.replace(' ', '_').replace('-', '_')
        # Remove any characters that aren't alphanumeric or underscore
        name = re.sub(r'[^A-Za-z0-9_]', '', name)
        return name
    
    def _read_excel_data(self, excel_path, header_finder_func):
        """
        Read Excel data into a pandas DataFrame with proper header detection.
        
        Args:
            excel_path: Path to the Excel file
            header_finder_func: Function to find the header row index
            
        Returns:
            tuple: (DataFrame, header_row_index) or (None, -1) on error
        """
        try:
            # Find header row
            header_row_index = header_finder_func(excel_path)
            if header_row_index == -1:
                print("SQLite: Could not find header row in Excel file.")
                return None, -1
            
            # Read Excel data
            print(f"SQLite: Reading Excel data from row {header_row_index + 1}...")
            df = pd.read_excel(
                excel_path,
                sheet_name=0,
                header=header_row_index,
                skiprows=header_row_index
            )
            
            # Clean up
            df.dropna(how='all', inplace=True)
            df.reset_index(drop=True, inplace=True)
                        
            # Convert Excel date/time columns to readable strings
            self._convert_excel_dates_to_strings(df)
            
            print(f"SQLite: Read {len(df)} rows from Excel.")
            return df, header_row_index
            
        except Exception as e:
            print(f"SQLite: Error reading Excel - {e}")
            traceback.print_exc()
            return None, -1
            
    def _convert_excel_dates_to_strings(self, df):
        """
        Detect and convert Excel date/time columns to readable string format.
        Excel stores dates as floats (e.g., 45961.5033680556 = 2025-10-31 12:04:51).
        
        Args:
            df: DataFrame with potential date columns
        """
        # Common date/time column name patterns (case-insensitive)
        date_keywords = ['date', 'time', 'datetime', 'timestamp', 'utc', 'local']
        
        for col in df.columns:
            col_lower = str(col).lower()
            
            # Check if column name suggests it's a date/time column
            is_date_col = any(keyword in col_lower for keyword in date_keywords)
            
            if is_date_col:
                # Check if the column contains numeric values (Excel date format)
                # Excel dates are stored as floats
                if pd.api.types.is_numeric_dtype(df[col]):
                    try:
                        # Check if values look like Excel dates (typically between 1 and 50000+)
                        # Also check that not all values are NaN
                        sample = df[col].dropna()
                        if len(sample) > 0:
                            # Convert Excel date numbers to datetime objects, then to strings
                            # Excel dates are days since 1899-12-30
                            print(f"SQLite: Converting date column '{col}' from Excel format to strings...")
                            
                            # Use pandas to convert Excel serial dates to datetime
                            # Note: Excel uses 1899-12-30 as day 0 (not 1900-01-01)
                            df[col] = pd.to_datetime(df[col], unit='D', origin='1899-12-30', errors='coerce')
                            
                            # Convert datetime objects to string format
                            df[col] = df[col].dt.strftime('%Y-%m-%d %H:%M:%S')
                            
                            # Replace NaT (Not a Time) with empty string
                            df[col] = df[col].fillna('')
                            
                            print(f"SQLite: Converted '{col}' successfully.")
                    except Exception as e:
                        print(f"SQLite: Could not convert column '{col}': {e}")
                        # Leave the column as-is if conversion fails

    def _validate_and_fix_excel_rows(self, df, header_row_index):
            """
            Validate and fix excel_row values in the DataFrame.
            - Ensures excel_row column exists as first column
            - Populates with actual Excel row numbers (data row index + header row + 2)
            - Fixes empty, duplicate, or malformatted values
            
            Args:
                df: DataFrame to validate
                header_row_index: The index of the header row in the Excel file (0-based)
                
            Returns:
                DataFrame with validated excel_row as first column
            """
            print("SQLite: Validating excel_row values...")
            
            # Check if excel_row column already exists
            excel_row_col = None
            for col in df.columns:
                if str(col).strip().lower() == 'excel_row':
                    excel_row_col = col
                    break
            
            # Calculate actual Excel row numbers
            # Formula: DataFrame index (0-based) + header_row_index + 2
            # +1 for Excel's 1-based indexing, +1 more to skip the header row itself
            excel_row_numbers = [i + header_row_index + 2 for i in range(len(df))]
            
            if excel_row_col is None:
                # Create new excel_row column
                print("SQLite: Creating 'excel_row' column...")
                df.insert(0, 'excel_row', excel_row_numbers)
            else:
                # Validate and fix existing excel_row column
                print("SQLite: Validating existing 'excel_row' column...")
                
                # Helper function to check if excel_row is valid
                def is_valid_row_number(val, expected):
                    if pd.isna(val) or val == '':
                        return False
                    try:
                        val_int = int(val)
                        # Valid if it's a positive integer
                        return val_int > 0
                    except:
                        return False
                
                # Find rows needing fixes
                needs_fix = []
                for idx, (actual, expected) in enumerate(zip(df[excel_row_col], excel_row_numbers)):
                    if not is_valid_row_number(actual, expected):
                        needs_fix.append(idx)
                
                if needs_fix:
                    print(f"SQLite: Fixing {len(needs_fix)} excel_row values...")
                    for idx in needs_fix:
                        df.at[idx, excel_row_col] = excel_row_numbers[idx]
                
                # Check for duplicates
                duplicates = df[excel_row_col].duplicated(keep=False)
                if duplicates.any():
                    num_duplicates = duplicates.sum()
                    print(f"SQLite: Found {num_duplicates} duplicate excel_row values. Regenerating...")
                    for idx in df[duplicates].index:
                        df.at[idx, excel_row_col] = excel_row_numbers[idx]
                
                # Move excel_row to first position if not already
                cols = df.columns.tolist()
                if cols[0] != excel_row_col:
                    cols.remove(excel_row_col)
                    cols.insert(0, excel_row_col)
                    df = df[cols]
            
            print(f"SQLite: excel_row column validated. Range: {excel_row_numbers[0]} to {excel_row_numbers[-1]}")
            return df
    
    def _count_uuid_issues(self, df):
        """
        Quickly count how many UUIDs need fixing without modifying the DataFrame.
        
        Args:
            df: DataFrame with potential UUID column
            
        Returns:
            int: Total number of UUIDs that need fixing (empty/malformed + duplicates)
        """
        # Check if UUID column exists
        uuid_col = None
        for col in df.columns:
            if str(col).strip().upper() == 'UUID':
                uuid_col = col
                break
        
        if uuid_col is None:
            return 0
        
        # Helper function to check if UUID is valid
        def is_valid_uuid(val):
            if pd.isna(val) or val == '':
                return False
            val_str = str(val).strip()
            if len(val_str) < 32:  # UUID should be at least 32 chars (without hyphens)
                return False
            try:
                uuid.UUID(val_str)
                return True
            except:
                return False
        
        # Count rows needing UUID fixes
        needs_fix_mask = ~df[uuid_col].apply(is_valid_uuid)
        num_needs_fix = needs_fix_mask.sum()
        
        # Count duplicates (keep='first' to only count extras)
        duplicate_mask = df[uuid_col].duplicated(keep='first')
        num_duplicates = duplicate_mask.sum()
        
        return num_needs_fix + num_duplicates
    
    def _validate_and_fix_uuids(self, df):
        """
        Validate and fix UUIDs in the DataFrame if UUID column exists.
        - Generates UUIDs for empty cells
        - Fixes malformatted UUIDs
        - Regenerates duplicate UUIDs
        
        Args:
            df: DataFrame with potential UUID column
            
        Returns:
            DataFrame with validated UUIDs
        """
        # Check if UUID column exists
        uuid_col = None
        for col in df.columns:
            if str(col).strip().upper() == 'UUID':
                uuid_col = col
                break
        
        if uuid_col is None:
            print("SQLite: No UUID column found in Excel. Skipping UUID validation.")
            return df, {}
        
        print("SQLite: Validating UUIDs...")
        
        # Track which rows were fixed
        uuid_fixes = {}
        # Helper function to check if UUID is valid
        def is_valid_uuid(val):
            if pd.isna(val) or val == '':
                return False
            val_str = str(val).strip()
            if len(val_str) < 32:  # UUID should be at least 32 chars (without hyphens)
                return False
            try:
                # Try to parse as UUID
                uuid.UUID(val_str)
                return True
            except:
                return False
        
        # Find rows needing UUID fixes
        needs_fix_mask = ~df[uuid_col].apply(is_valid_uuid)
        num_needs_fix = needs_fix_mask.sum()
        
        if num_needs_fix > 0:
            print(f"SQLite: Generating/fixing {num_needs_fix} UUIDs...")
            for idx in df[needs_fix_mask].index:
                new_uuid = str(uuid.uuid4())
                df.at[idx, uuid_col] = new_uuid
                uuid_fixes[idx] = new_uuid
        
        # Check for duplicates
        duplicates = df[uuid_col].duplicated(keep=False)
        if duplicates.any():
            num_duplicates = duplicates.sum()
            print(f"SQLite: Found {num_duplicates} duplicate UUIDs. Regenerating...")
            
            # Keep first occurrence, regenerate the rest
            duplicate_mask = df[uuid_col].duplicated(keep='first')
            for idx in df[duplicate_mask].index:
                new_uuid = str(uuid.uuid4())
                df.at[idx, uuid_col] = new_uuid
                uuid_fixes[idx] = new_uuid
        
        if num_needs_fix == 0 and not duplicates.any():
            print("SQLite: All UUIDs are valid and unique.")
        
        return df, uuid_fixes
    
    def _write_uuid_fixes_to_excel(self, excel_path, header_row_index, uuid_fixes, uuid_col_name):
        """
        Write UUID fixes back to the Excel file.
        
        Args:
            excel_path: Path to the Excel file
            header_row_index: The index of the header row (0-based)
            uuid_fixes: Dictionary of {dataframe_index: new_uuid}
            uuid_col_name: Name of the UUID column in Excel
            
        Returns:
            bool: True if successful, False otherwise
        """
        if not uuid_fixes:
            return True  # Nothing to fix
        
        print(f"SQLite: Writing {len(uuid_fixes)} UUID fixes back to Excel...")
        
        app, book, sheet = None, None, None
        opened_new_app = False
        
        try:
            # Try to find already open Excel instance
            target_norm_path = os.path.normcase(os.path.abspath(excel_path))
            for running_app in xw.apps:
                for wb in running_app.books:
                    try:
                        if os.path.normcase(os.path.abspath(wb.fullname)) == target_norm_path:
                            book, app = wb, running_app
                            break
                    except:
                        continue
                if book:
                    break
            
            # If not found, open new instance
            if book is None:
                app = xw.App(visible=False)
                opened_new_app = True
                book = app.books.open(excel_path)
            
            sheet = book.sheets[0]
            
            # Find UUID column position (1-based for Excel)
            header_row_excel = header_row_index + 1
            header_range = sheet.range(f'A{header_row_excel}').expand('right')
            headers = header_range.value
            
            uuid_col_index = None
            for i, header in enumerate(headers, start=1):
                if header and str(header).strip().upper() == uuid_col_name.strip().upper():
                    uuid_col_index = i
                    break
            
            if uuid_col_index is None:
                print(f"SQLite: WARNING - Could not find UUID column '{uuid_col_name}' in Excel header.")
                return False
            
            # Get column letter from index
            col_letter = xw.utils.col_name(uuid_col_index)
            
            # Write each UUID fix
            for df_index, new_uuid in uuid_fixes.items():
                # Convert DataFrame index to Excel row number
                # Formula: df_index + header_row_index + 2 (skip header, 1-based indexing)
                excel_row = df_index + header_row_index + 2
                cell_address = f"{col_letter}{excel_row}"
                
                sheet.range(cell_address).value = new_uuid
                print(f"SQLite:   Fixed UUID in row {excel_row}: {new_uuid}")
            
            # Save the workbook
            book.save()
            print(f"SQLite: Successfully wrote {len(uuid_fixes)} UUID fixes to Excel.")
            
            return True
            
        except Exception as e:
            print(f"SQLite: ERROR writing UUID fixes to Excel: {e}")
            traceback.print_exc()
            return False
            
        finally:
            # Only close if we opened a new app
            if opened_new_app:
                if book:
                    try:
                        book.close()
                    except:
                        pass
                if app:
                    try:
                        app.quit()
                    except:
                        pass
    
    
    def _ensure_table_exists(self, df, table_name):
        """
        Create or recreate the SQLite table based on DataFrame structure.
        
        Args:
            df: DataFrame with sanitized column names
            table_name: Name of the table to create
        """
        cursor = self.conn.cursor()
        
        # Check if table exists
        cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{table_name}'")
        table_exists = cursor.fetchone() is not None
        
        if table_exists:
            # Get existing columns
            cursor.execute(f"PRAGMA table_info('{table_name}')")
            existing_cols = {row[1] for row in cursor.fetchall()}
            
            # Get new columns
            new_cols = set(df.columns)
            
            # If columns match, keep the table
            if existing_cols == new_cols:
                print(f"SQLite: Table '{table_name}' already exists with correct structure.")
                return
            
            # Otherwise, drop and recreate
            print(f"SQLite: Table structure changed. Recreating '{table_name}'...")
            cursor.execute(f"DROP TABLE IF EXISTS '{table_name}'")
            self.conn.commit()
        
        # Create new table
        print(f"SQLite: Creating table '{table_name}'...")
        
        # All columns as TEXT for simplicity and flexibility
        col_definitions = [f'"{col}" TEXT' for col in df.columns]
        create_sql = f"CREATE TABLE '{table_name}' ({', '.join(col_definitions)})"
        
        cursor.execute(create_sql)
        self.conn.commit()
        print(f"SQLite: Table '{table_name}' created successfully.")
    
    def full_sync(self, excel_path, header_finder_func, skip_uuid_fixes=False):
        """
        Perform a complete synchronization from Excel to SQLite.
        Makes the database identical to the Excel sheet.
        
        Args:
            excel_path: Path to the Excel file
            header_finder_func: Function to find the header row
            
        Returns:
            bool: True if successful, False otherwise
        """
        print("\n=== SQLite: Starting Full Synchronization ===")
        
        try:
            # 1. Read Excel data
            df, header_row = self._read_excel_data(excel_path, header_finder_func)
            if df is None:
                return False
            
            # 2. Add/validate excel_row column FIRST (before any other processing)
            df = self._validate_and_fix_excel_rows(df, header_row)
            
            # 3. Validate and fix UUIDs if column exists (unless skipped)
            if skip_uuid_fixes:
                print("SQLite: Skipping UUID validation as requested.")
                uuid_fixes = {}
            else:
                df, uuid_fixes = self._validate_and_fix_uuids(df)
                
                # 3.5. Write UUID fixes back to Excel BEFORE sanitizing column names
                if uuid_fixes:
                    # Find original UUID column name (before sanitization)
                    uuid_col_name = None
                    for col in df.columns:
                        if str(col).strip().upper() == 'UUID':
                            uuid_col_name = col
                            break
                    
                    if uuid_col_name:
                        write_success = self._write_uuid_fixes_to_excel(
                            excel_path, 
                            header_row, 
                            uuid_fixes, 
                            uuid_col_name
                        )
                        if not write_success:
                            print("SQLite: WARNING - UUID fixes may not have been written to Excel.")
                            print("SQLite: Excel and database UUIDs may be out of sync!")
           
            # 4. Sanitize column names (spaces to underscores)
            original_columns = df.columns.tolist()
            df.columns = [self._sanitize_column_name(col) for col in df.columns]
            
            # Log column name changes
            for orig, new in zip(original_columns, df.columns):
                if orig != new:
                    print(f"SQLite: Column '{orig}' -> '{new}'")
            
            # 5. Set table name from Excel filename
            self.table_name = Path(excel_path).stem
            
            # 6. Ensure table exists with correct structure
            self._ensure_table_exists(df, self.table_name)
            
            # 7. Clear existing data and insert fresh data
            cursor = self.conn.cursor()
            
            print(f"SQLite: Clearing existing data from '{self.table_name}'...")
            cursor.execute(f"DELETE FROM '{self.table_name}'")
            
            print(f"SQLite: Inserting {len(df)} rows...")
            
            # Prepare INSERT statement
            cols = list(df.columns)
            col_str = ', '.join(f'"{col}"' for col in cols)
            placeholders = ', '.join(['?' for _ in cols])
            insert_sql = f"INSERT INTO '{self.table_name}' ({col_str}) VALUES ({placeholders})"
            
            # Insert all rows
            rows_data = [tuple(row) for row in df.values]
            cursor.executemany(insert_sql, rows_data)
            
            self.conn.commit()
            
            print(f"SQLite: Full sync complete. {len(df)} rows mirrored to database.")
            print("=== SQLite: Synchronization Finished ===\n")
            return True
            
        except sqlite3.Error as e:
            print(f"SQLite: Database error during sync - {e}")
            traceback.print_exc()
            try:
                self.conn.rollback()
            except:
                pass
            return False
        except Exception as e:
            print(f"SQLite: Unexpected error during sync - {e}")
            traceback.print_exc()
            return False
    
    def add_single_row(self, row_data, excel_path=None, excel_row=None):
        """
        Add a single row to the database quickly without full sync.
        Used when logging new events to keep the UI responsive.
        
        Args:
            row_data: Dictionary of {column_name: value}
                     Column names should match Excel headers (will be sanitized)
            excel_path: Optional Excel file path to derive table name if not set
            excel_row: The Excel row number where this data was written (required)
        
        Returns:
            bool: True if successful, False otherwise
        """
        # If table name not set, try to derive it from excel_path
        if not self.table_name and excel_path:
            self.table_name = Path(excel_path).stem
            print(f"SQLite: Set table name to '{self.table_name}'")
        
        if not self.table_name:
            print("SQLite: ERROR - No table name set. Cannot add row.")
            print("SQLite: Run full_sync first or provide excel_path parameter.")
            return False
        
        if excel_row is None:
            print("SQLite: WARNING - No excel_row provided. This row will be missing its Excel position reference.")
            print("SQLite: It's recommended to provide the excel_row parameter.")
        
        try:
            # 1. Sanitize column names
            sanitized_data = {}
            
            # Add excel_row FIRST if provided
            if excel_row is not None:
                sanitized_data['excel_row'] = str(excel_row)
            for key, value in row_data.items():
                sanitized_key = self._sanitize_column_name(key)
                sanitized_data[sanitized_key] = value
            
            # 2. Generate UUID if UUID column exists and value is missing
            uuid_col = self._sanitize_column_name('UUID')
            if uuid_col in sanitized_data:
                current_uuid = sanitized_data[uuid_col]
                if not current_uuid or str(current_uuid).strip() == '':
                    sanitized_data[uuid_col] = str(uuid.uuid4())
                    print(f"SQLite: Generated new UUID for row: {sanitized_data[uuid_col]}")
            
            # 3. Insert row
            cursor = self.conn.cursor()
            
            cols = list(sanitized_data.keys())
            col_str = ', '.join(f'"{col}"' for col in cols)
            placeholders = ', '.join(['?' for _ in cols])
            values = [sanitized_data[col] for col in cols]
            
            insert_sql = f"INSERT INTO '{self.table_name}' ({col_str}) VALUES ({placeholders})"
            
            cursor.execute(insert_sql, values)
            self.conn.commit()
            
            if excel_row is not None:
                print(f"SQLite: Added 1 row to '{self.table_name}' (Excel row {excel_row}).")
            else:
                print(f"SQLite: Added 1 row to '{self.table_name}'.")
            return True
            
        except sqlite3.Error as e:
            print(f"SQLite: Error adding row - {e}")
            traceback.print_exc()
            try:
                self.conn.rollback()
            except:
                pass
            return False
        except Exception as e:
            print(f"SQLite: Unexpected error adding row - {e}")
            traceback.print_exc()
            return False


# --- FolderMonitor Class ---
class FolderMonitor(FileSystemEventHandler):
    '''
    A custom event handler for watchdog that monitors a specified folder for new or  files
    matching a given extension. It updates a global cache with the latest matching file.
    '''
    def __init__(self, path, folder_name, gui_instance, extension=""):
        self.path = path
        self.folder_name = folder_name
        self.gui_instance = gui_instance
        self.extension = extension.lower() if extension else ''
        self.latest_file = None
        #self.update_latest_file() # Initial scan

    def on_(self, event):
        if not event.is_directory:
        # Change self.file_extension to self.extension in these two places
            if not self.extension or event.src_path.lower().endswith(self.extension.lower()):
                self.update_latest_file()

    def on_created(self, event):
        if not event.is_directory and (not self.extension or event.src_path.lower().endswith(self.extension)):
            self._update_if_newer(event.src_path)

    def _update_if_newer(self, file_path):
        current_mtime = os.path.getmtime(file_path)
        cached_file = folder_cache.get(self.folder_name)
        
        if not cached_file or current_mtime > os.path.getmtime(cached_file):
            folder_cache[self.folder_name] = file_path
            # self.gui_instance.update_status(f"Newer file found in {self.folder_name}: {os.path.basename(file_path)}")

    def update_latest_file(self):
        '''Scans the folder AND ALL SUBFOLDERS to find the truly latest file and updates the cache.'''
        latest = None
        latest_mtime = -1
        try:
            # Use os.walk for a recursive search through all directories and subdirectories
            for root, _, files in os.walk(self.path):
                if not files and os.path.normcase(root) == os.path.normcase(self.path):
                    print(f"DEBUG {self.folder_name}: Found folder root, 0 files matched extension '{self.extension}' in root.")
                for f_name in files:
                    # Check if the file matches the extension (if one is specified)
                    if not self.extension or f_name.lower().endswith(self.extension):
                        f_path = os.path.join(root, f_name)
                        try:
                            mtime = os.path.getmtime(f_path)
                            if self.folder_name == "Naviscan": # Only print for the problem folder
                                print(f"DEBUG Naviscan: Candidate found: {os.path.basename(f_path)}, in dir: {os.path.basename(root)}, mtime: {datetime.datetime.fromtimestamp(mtime)}")
                            if mtime > latest_mtime:
                                latest_mtime = mtime
                                latest = f_path
                        except FileNotFoundError:
                            continue # File might have been deleted during the scan
                            
        except FileNotFoundError:
            self.gui_instance.update_status(f"Monitoring error: Folder '{self.path}' not found for '{self.folder_name}'.")
        except Exception as e:
            self.gui_instance.update_status(f"Monitoring error in '{self.folder_name}': {e}")

        latest = self.gui_instance.find_latest_file_in_folder(self.path, self.extension)

        if latest:
            # If a new latest file is found, update the cache
            if folder_cache.get(self.folder_name) != latest:
                print(f"Cache update for '{self.folder_name}': {os.path.basename(latest)}")
                folder_cache[self.folder_name] = latest
        elif self.folder_name in folder_cache:
            # If no files are found, clear the cache entry
            del folder_cache[self.folder_name]

# --- Main Application GUI Class ---
class DataLoggerGUI:
    ''' Main GUI class for the Data Acquisition Logger application.
        This class initializes the main window, sets up styles, variables, and handles user interactions.
        It includes methods for creating buttons, managing settings, and logging events.

        Attributes:
        * master: The root Tkinter window or parent widget.
        * settings_file: Path to the settings file.
        * style: The ttk.Style object for styling widgets.
        * status_var: StringVar for status messages.
        * monitor_status_label: Label to display monitoring status.
        * settings_window_instance: Instance of the settings window to avoid multiple instances.
        * log_file_path: Path to the Excel log file.
        * txt_folder_path: Folder path for TXT files.
        * txt_file_path: Path to the latest found TXT file.
        * txt_field_columns: Dictionary mapping expected field names to their corresponding Excel or DB column names.
        * txt_field_skips: Dictionary for TXT field skips.
        * num_custom_buttons: Number of custom buttons to render for Set 1.
        * custom_button_configs: List of dictionaries containing configurations for custom buttons in Set 1.
        * txt_folder_path_set2: Folder path for the second set of TXT files.
        * txt_file_path_set2: Path to the latest found TXT file for Set 2.
        * txt_field_columns_set2: Dictionary mapping expected field names to their corresponding Excel or DB column names for Set 2.
        * txt_field_skips_set2: Dictionary for TXT field skips for Set 2.
        * num_custom_buttons_set2: Number of custom buttons to render for Set 2.
        * custom_button_configs_set2: List of dictionaries containing configurations for custom buttons in Set 2.
        * folder_paths: Dictionary of monitored folders (e.g., for SVP files).
        * folder_columns: Maps each folder to the corresponding Excel/DB column name.
        * file_extensions: File filters (e.g., .svp, .txt) for each monitored folder.
        * folder_skips: Skip flags for folders.
        * monitors: Holds the actual folder watchers.
        * button_colors: Dictionary mapping button text to their colors.
        * main_frame: The main frame containing all widgets.
        '''

    # --- Initialization ---
    def __init__(self, master):
        '''
        Initializes the main GUI application.
        This method sets up the main window, initializes styles, variables, and loads settings.
        Arguments:
        * master: The root Tkinter window or parent widget.
        '''
        self.calculate_logoff_values = tk.BooleanVar(value=True) # Defaults to enabled
        self.last_log_on_kp = None
        self.log_on_time = None

        print("\n--- Starting Online Logger ---\n")

        # Initialize the main window
        self.master = master
        master.title("Online Logger")
        master.geometry("1400x250")
        master.minsize(800, 200)
        
        # Set window icon
        self.set_window_icon()

        self.init_styles()
        self.init_variables()
        self.static_field_configs = []
        self.init_settings()
        
        # Update window title with version and project name
        self.update_window_title()

        # --- Main Layout ---
        self.main_frame = ttk.Frame(self.master, padding="5")
        self.main_frame.grid(row=0, column=0, sticky="nsew")
        self.master.columnconfigure(0, weight=1)
        self.master.rowconfigure(0, weight=1)

        # Configure the main 3-column layout for the application
        self.main_frame.columnconfigure(0, weight=4) # Custom Buttons area (largest)
        self.main_frame.columnconfigure(1, weight=1) # General Buttons area
        self.main_frame.columnconfigure(2, weight=1) # Configuration area
        self.main_frame.rowconfigure(0, weight=1)    # Main content row
        self.main_frame.rowconfigure(1, weight=0)    # Progress bar row
        self.main_frame.rowconfigure(2, weight=0)    # Status bar row

        # Create container frames for each section
        self.custom_buttons_frame = ttk.Frame(self.main_frame)
        self.custom_buttons_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

        self.general_buttons_frame = ttk.Frame(self.main_frame)
        self.general_buttons_frame.grid(row=0, column=1, sticky="nsew", padx=5)

        self.config_frame = ttk.Frame(self.main_frame)
        self.config_frame.grid(row=0, column=2, sticky="nsew", padx=(5, 0))
        # Configure the config frame to place status indicators at the bottom
        self.config_frame.rowconfigure(0, weight=1) # Buttons will be at the top
        self.config_frame.rowconfigure(1, weight=0) # Indicators will be at the bottom

        
        print("-> Initializing GUI components...")
        # Create all buttons and place them in the correct frames
        self.create_main_buttons()

        # Create status indicators and place them in the config frame
        self.create_status_indicators()

        #Create Progress bar for folder monitoring
        self.create_progress_bar()

        # Create status bar at the very bottom, spanning all columns
        self.create_status_bar()

        # Scheduled tasks
        self.schedule_new_day() # Start the midnight log schedule
        self.schedule_hourly_log() # Start the hourly log schedule
        # self.start_monitoring() 
        self.update_monitor_indicator_text()  # Initial monitor start & status update

        #schedule sutomatic sync routing
        

        # Open the settings window by default when the app starts
        self.startup_settings()

        
    def init_styles(self):
        ''' 
        Initializes the styles for the application using ttk.Style.
        This method sets the theme and configures styles for various widgets.
        It also handles theme availability and sets default styles.
        '''
        self.style = ttk.Style()
        self.style.theme_use("clam")

        self.style.configure("TFrame", background="#f0f0f0")
        self.style.configure("TLabel", background="#f0f0f0", font=("Arial", 10))
        self.style.configure("TButton", font=("Arial", 10, "bold"), padding=4) # Smaller padding for buttons
        self.style.configure("TEntry", font=("Arial", 10), padding=4)
        self.style.configure("StatusBar.TLabel", background="#e0e0e0", font=("Arial", 8), relief=tk.SUNKEN, padding=(3, 1)) # Smaller font/padding for status bar
        self.style.configure("Header.TFrame", background="#dcdcdc")
        # Define styles for selected and unselected rows
        self.style.configure("Row0.TFrame", background="#ffffff")
        self.style.configure("Row1.TFrame", background="#f5f5f5")
        self.style.configure("Selected.TFrame", background="#ADD8E6") # Light blue for selection
        self.style.configure("TLabelframe", background="#f0f0f0", padding=3, relief="flat") # Flat relief for compact
        self.style.configure("TLabelframe.Label", background="#f0f0f0", font=("Arial", 9, "bold")) # Smaller font
        self.style.configure("Large.TCheckbutton", font=("Arial", 10)) # For settings checkbox
        self.style.configure("Small.TButton", font=("Arial", 8), padding=3) # Define a new custom style for smaller buttons that keeps the standard border.
        self.style.configure("Accent.TButton", font=("Arial", 9, "bold"), foreground="white", background="#0078D4") # For settings save

        self.style.map("TButton",
                        foreground=[('pressed', 'darkblue'), ('active', 'blue'), ('disabled', '#999999')],
                        background=[('pressed', '!disabled', '#c0c0c0'), ('active', '#e0e0e0')]
                        )
        self.style.map("TLabel", background=[('selected', '#ADD8E6')]) # Ensure labels in selected row change color

        selection_color = "#cce5ff" # A light blue for highlighting
        self.style.configure("Selected.TLabel", background=selection_color)
        self.style.configure("Selected.TEntry", fieldbackground=selection_color)
        self.style.configure("Selected.TButton", background=selection_color)
        self.style.configure("Selected.TCheckbutton", background=selection_color)

    def init_variables(self):
        '''
        Initializes all key configuration variables, paths, button presets, and GUI state defaults used throughout the application. 
        This method is called when the GUI is first launched.
        '''
        self.log_file_path = None

        # Settings File Configuration (Projects-based)
        self.settings_file = None  # Active project path; None means not yet saved to a project file
        # Track last used project path (for Settings window Projects tab)
        try:
            settings_dir = os.path.join(os.getcwd(), "settings")
            self.current_project_path = os.path.join(settings_dir, "new_settings.json")
        except Exception:
            self.current_project_path = None

        # Try to load the last-used project path from state and use it if valid
        try:
            last_path = self.load_last_project_path()
            if last_path and os.path.exists(last_path):
                self.current_project_path = last_path
                self.settings_file = last_path
                print(f"Using last project from state: {last_path}")
        except Exception as _e:
            # Non-fatal: fall back to defaults if state cannot be read
            pass

        # Event Code Configuration
        self.event_codes_file = EVENT_CODES_FILE
        self.event_codes = {} # Will store {'code': 'description'}

        self.main_button_configs = {
            "Log on": {"event_text": "Log on event occurred", "event_code": ""},
            "Log off": {"event_text": "Log off event occurred", "event_code": ""},
            "Event": {"event_text": "", "event_code": ""}, # Intentionally blank for the "Event" button
            "SVP": {"event_text": "SVP applied", "event_code": ""},
            "Manual KP Log": {"event_text": "Auto generated", "event_code": ""}, 
        }
        
        # Original TXT path for the 'Event' button
        self.txt_folder_path = None 
        # New TXT paths for additional sources
        self.txt_folder_path_set2 = None
        self.txt_folder_path_set3 = None
        self.txt_folder_path_set4 = None
        self.txt_folder_path_set5 = None

        # Dictionary to hold user-defined aliases for TXT sources
        self.txt_source_aliases = {
            "Main TXT": "Main TXT",
            "TXT Source 2": "TXT Source 2",
            "TXT Source 3": "TXT Source 3",
            "TXT Source 4": "TXT Source 4",
            "TXT Source 5": "TXT Source 5"
        }

        self.source_based_colors = {
            "Main TXT": "#BAE1FF",      # Light Blue
            "TXT Source 2": "#BAFFC9",    # Light Green
            "TXT Source 3": "#FFFFBA",    # Light Yellow
            "TXT Source 4": "#FFB3BA",    # Light Red/Pink
            "TXT Source 5": "#E0BBE4",    # Light Purple
            "None": None          # No color for buttons with no source
        }

        self.all_txt_mappings = {
            "Main TXT": [
                {"field": "KP", "column_name": "KP", "skip": False},
                {"field": "DCC", "column_name": "DCC", "skip": False},
                {"field": "Line name", "column_name": "Runline", "skip": False},
                {"field": "Latitude", "column_name": "Latitude", "skip": False},
                {"field": "Longitude", "column_name": "Longitude", "skip": False},
                {"field": "Easting", "column_name": "Easting", "skip": False},
                {"field": "Northing", "column_name": "Northing", "skip": False},
            ],
            "TXT Source 2": [], # Start empty, user will configure
            "TXT Source 3": [],
            "TXT Source 4": [],
            "TXT Source 5": [],
            # No entry for "None" as it means no mapping is needed
        }
        
    
        # NEW: For data generated by the application itself
        self.generated_fields_config = [
            {"field": "Date-Time", "column_name": "UTC Date-Time", "skip": False, "source": "PC Time (UTC)"},
            {"field": "Local Time", "column_name": "Local Time", "skip": False, "source": "PC Time + Offset"},
            {"field": "Event", "column_name": "Event", "skip": False, "source": "Button"},
            {"field": "Code", "column_name": "Code", "skip": False, "source": "Button"},
            {"field": "KP Ref.", "column_name": "KP Ref.", "skip": False, "source": "Source Alias"},
            {"field": "UUID", "column_name": "UUID", "skip": False, "source": "Generated"}
        ]

        # For data from static cells in Excel
        self.static_field_configs = []


        self.folder_paths = {}
        self.folder_columns = {}
        self.file_extensions = {}
        self.folder_skips = {}
        self.folder_log_x_instead = {}
        self.folder_log_ext_vars = {}
        self.monitors = {}
        self.num_custom_buttons = 3
        self.MAX_CUSTOM_BUTTONS = 50 # Define the maximum number of custom buttons
        
        # Each custom button config now includes a 'txt_source_key'
        self.custom_button_configs = [
            {"text": "Custom Event 1", "event_text": "Custom Event 1 Triggered", "txt_source_key": "Main TXT", "tab_group": "Main", "event_code": ""},
            {"text": "Custom Event 2", "event_text": "Custom Event 2 Triggered", "txt_source_key": "None", "tab_group": "Main", "event_code": ""},
            {"text": "Custom Event 3", "event_text": "Custom Event 3 Triggered", "txt_source_key": "None", "tab_group": "Main", "event_code": ""}
        ]
        self.custom_buttons = []
        self.button_colors = {
            "Log on": ("#90EE90", None),  # Format: (background_color, font_color)
            "Log off": ("#FFB6C1", None),
            "Event": ("#FFFFE0", None),
            "SVP": ("#ADD8E6", None),
            "New Day": ("#FFFF99", None),
            "Hourly KP Log": ("#FFFF99", None)
        }
        # Initialize custom button colors to None for both background and font
        for i in range(self.MAX_CUSTOM_BUTTONS):
            self.button_colors[f"Custom {i+1}"] = (None, None)
        
        # Define the three tab groups explicitly
        self.custom_button_tab_groups = ["Main"]
        self.custom_button_tab_frames = {}


        self.time_offset_hours = tk.DoubleVar(value=0.0)
        self.active_logging_threshold_seconds = tk.IntVar(value=15)

        # Variables to control the automatic, timed events
        self.new_day_event_enabled_var = tk.BooleanVar(value=True)
        self.hourly_event_enabled_var = tk.BooleanVar(value=True)
        self.hourly_log_txt_source_key = tk.StringVar(value="Main TXT")

        self.always_on_top_var = tk.BooleanVar(value=False)
        self.settings_window_instance = None # Track settings window
        self.custom_inline_editor_window = None # To track the open inline editor
        self.auto_sync_enabled_var = tk.BooleanVar(value=True)
        self.auto_sync_interval_min_var = tk.IntVar(value=15)
        self._auto_sync_timer_id = None
        self.is_monitoring = False 
        self.monitoring_button = None 

        self.sqlite_mirror_enabled_var = tk.BooleanVar(value=False)
        self.sqlite_manager = None
        self.db_path = None
        self.sqlite_db_path = None

        self.status_var = tk.StringVar()
        self.monitor_status_label = None
        self.settings_window_instance = None # Track settings window
        self.custom_inline_editor_window = None # To track the open inline editor


    def init_settings(self):
        ''' Initialize settings from the active project file if present, otherwise load the blank project template. '''
        if self.settings_file and os.path.exists(self.settings_file):
            self.load_settings()
        else:
            try:
                print("No active project found. Loading from blank project template...")
                self.revert_to_defaults()
            except Exception as e:
                messagebox.showwarning("Initialization Error", "Blank project template not found. Please create settings/config/blank_project.json.", parent=self.master)

    # --- GUI Creation ---
    def create_main_buttons(self):
        """
        Builds and renders all the buttons in the GUI dynamically, grouped for better intuitiveness.
        Custom buttons are now organized into tabs within a ttk.Notebook.
        """
        # (The first part of this function that clears frames is unchanged)
        for frame in [self.custom_buttons_frame, self.general_buttons_frame, self.config_frame]:
            for widget in frame.winfo_children():
                widget.destroy()
        self.custom_buttons = []

        # --- Section 1: Custom Events (Left Side) ---
        # ... (unchanged setup for custom buttons) ...
        custom_lf = ttk.LabelFrame(self.custom_buttons_frame, text="Custom Events")
        custom_lf.pack(fill="both", expand=True)
        self.custom_buttons_notebook = ttk.Notebook(custom_lf)
        self.custom_buttons_notebook.pack(fill="both", expand=True, padx=5, pady=5)
        self.custom_buttons_notebook.bind("<Button-3>", self._show_tab_context_menu)
        self.custom_button_tab_frames = {}
        all_tab_groups = sorted(list(set(self.custom_button_tab_groups)))
        for tab_group_name in all_tab_groups:
            if tab_group_name:
                tab_frame = ttk.Frame(self.custom_buttons_notebook, padding=5)
                self.custom_buttons_notebook.add(tab_frame, text=tab_group_name)
                self.custom_button_tab_frames[tab_group_name] = tab_frame
                tab_frame.bind("<Button-3>", self._show_add_button_context_menu)
        custom_buttons_by_tab = {group: [] for group in all_tab_groups if group}
        for config in self.custom_button_configs[:self.num_custom_buttons]:
            tab_group = config.get("tab_group", "Main")
            if tab_group not in custom_buttons_by_tab:
                custom_buttons_by_tab[tab_group] = []
            custom_buttons_by_tab[tab_group].append(config)
        for tab_group, configs in custom_buttons_by_tab.items():
            if tab_group in self.custom_button_tab_frames:
                tab_frame = self.custom_button_tab_frames[tab_group]
                for i, config in enumerate(configs):
                    button_text = config.get("text", "Custom")
                    event_desc = config.get("event_text", "Triggered")
                    txt_source = config.get("txt_source_key", "None")
                    bg_color_hex, font_color_hex = self.button_colors.get(button_text, (None, None))
                    if not bg_color_hex:
                        bg_color_hex = self.source_based_colors.get(txt_source)
                    cleaned_button_text = ''.join(e for e in button_text if e.isalnum()) 
                    style_name = f"CustomBtn_{cleaned_button_text}.TButton"
                    style_config = {}
                    if bg_color_hex:
                        style_config['background'] = bg_color_hex
                    if font_color_hex:
                        style_config['foreground'] = font_color_hex
                    self.style.configure(style_name, font=("Arial", 10, "bold"), padding=4, **style_config)
                    button = ttk.Button(tab_frame, text=button_text, style=style_name)
                    button.config(command=lambda c=config, b=button: self.log_custom_event(c, b))
                    num_columns = LAYOUT_BUTTON_COLUMNS
                    row = i // num_columns
                    col = i % num_columns
                    button.grid(row=row, column=col, padx=3, pady=3, sticky="nsew")
                    tab_frame.columnconfigure(col, weight=1)
                    tab_frame.rowconfigure(row, weight=1)
                    original_index = self.custom_button_configs.index(config)
                    button.bind("<Button-3>", lambda e, idx=original_index: self._show_custom_button_context_menu(e, idx))
                    ToolTip(button, f"Log '{event_desc}' (Source: {txt_source})")
                    self.custom_buttons.append(button)


        # --- Section 2: General Event Buttons (Middle) ---
        general_lf = ttk.LabelFrame(self.general_buttons_frame, text="General Events")
        general_lf.pack(fill="both", expand=True)
        general_lf.columnconfigure((0, 1), weight=1)
        # Configure for 4 rows: 0, 1, 2 (Manual Log), 3 (Historic Event)
        general_lf.rowconfigure((0, 1, 2, 3), weight=1)

        # Helper function to create styled main buttons ( COLOR LOOKUP)
        def create_main_button(parent, text, command_func, tooltip_text, grid_row, grid_col):
            
            # --- CORRECTION APPLIED HERE (Ensures Manual KP Log uses Hourly KP Log's style) ---
            key_to_use = "Hourly KP Log" if text == "Manual KP Log" else text
            bg_color_hex, font_color_hex = self.button_colors.get(key_to_use, (None, None))
            # ---------------------------------------------------------------------------------
            
            cleaned_text = ''.join(e for e in text if e.isalnum())  
            style_name = f"MainBtn_{cleaned_text}.TButton"
            style_config = {}
            if bg_color_hex:
                style_config['background'] = bg_color_hex
            if font_color_hex:
                style_config['foreground'] = font_color_hex
            self.style.configure(style_name, font=("Arial", 10, "bold"), padding=4, **style_config)
            btn = ttk.Button(parent, text=text, style=style_name, command=command_func)
            btn.grid(row=grid_row, column=grid_col, padx=4, pady=4, sticky="nsew")
            btn.bind("<Button-3>", lambda e, name=text: self._show_main_button_context_menu(e, name))
            ToolTip(btn, tooltip_text)
            return btn

        # Create the standard buttons
        create_main_button(general_lf, "Log on", lambda b=None: self.log_event("Log on", b, "Main TXT"), "Record a 'Log on' marker.", 0, 0)
        create_main_button(general_lf, "Log off", lambda b=None: self.log_event("Log off", b, "Main TXT"), "Record a 'Log off' marker.", 1, 0)
        create_main_button(general_lf, "Event", lambda b=None: self.log_event("Event", b, "Main TXT"), "Record data from the Main TXT source.", 0, 1)
        create_main_button(general_lf, "SVP", lambda b=None: self.log_svp("SVP", b, "Main TXT"), "Record data and insert latest SVP filename.", 1, 1)

        # Add the new Manual Hourly KP Log button (Row 2)
        manual_hourly_btn = create_main_button(
            general_lf, 
            "Manual KP Log", 
            lambda b=None: self.trigger_manual_hourly_log_action(manual_hourly_btn), 
            "Manually trigger the hourly KP log and progress calculation.", 
            2, 
            0 
        )
        manual_hourly_btn.grid(columnspan=2, sticky="nsew")
        
        # Add the new "Add Historic Event" button to the grid
        historic_btn = ttk.Button(general_lf, text="Add Historic Event", command=self.add_historic_event)
        historic_btn.grid(row=3, column=0, columnspan=2, padx=4, pady=4, sticky="nsew") 
        ToolTip(historic_btn, "Add an event from a past date/time by searching the Main data source file.")
        
        # --- Section 3: Configuration Buttons (Right Side) ---
        config_lf = ttk.LabelFrame(self.config_frame, text="Configuration")
        config_lf.grid(row=0, column=0, sticky="nsew")
        self.config_frame.columnconfigure(0, weight=1)
        config_lf.columnconfigure((0, 1), weight=1)
        config_lf.rowconfigure((0, 1), weight=1)
        self.monitoring_button = ttk.Button(config_lf, text="Start Monitoring", style="Small.TButton", command=self.toggle_monitoring)
        self.monitoring_button.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=4, pady=(4, 2))
        ToolTip(self.monitoring_button, "Start or stop monitoring all configured folders for file changes.")
        self.update_monitoring_button_ui()
        btn_settings = ttk.Button(config_lf, text="Settings", style="Small.TButton", command=self.open_settings)
        btn_settings.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=4, pady=(2, 4))
        ToolTip(btn_settings, "Open the configuration window.")

    
    def toggle_sqlite_mirroring(self):
        """Handles enabling or disabling the SQLite mirror feature."""
        if self.sqlite_mirror_enabled_var.get():
            # --- ENABLE ---
            if not self.log_file_path or not os.path.exists(self.log_file_path):
                messagebox.showerror("Error", "Please set a valid Excel Log File path in Settings before enabling SQLite mirroring.", parent=self.master)
                self.sqlite_mirror_enabled_var.set(False)
                return

            if self.sqlite_db_path:
                self.db_path = self.sqlite_db_path
            else:
                self.db_path = str(Path(self.log_file_path).with_suffix('.db'))
            
            self.update_status(f"Enabling SQLite mirror at {self.db_path}...")
            
            try:
                self.sqlite_manager = SQLiteManager(self.db_path)
                
                # Count UUID issues before starting sync
                def _count_and_ask():
                    try:
                        self.update_status("Analyzing Excel data for UUID issues...")
                        header_finder = lambda path: self._find_header_row(path)
                        
                        # Read Excel data to count UUID issues
                        df, header_row = self.sqlite_manager._read_excel_data(
                            self.log_file_path,
                            header_finder
                        )
                        
                        if df is None:
                            raise Exception("Failed to read Excel data")
                        
                        # Quick count of UUID issues
                        num_issues = self.sqlite_manager._count_uuid_issues(df)
                        
                        if num_issues > 0:
                            # Estimate time: approximately 0.001 seconds per UUID fix + Excel write time
                            estimated_seconds = (num_issues * 0.001) + 2  # +2 seconds for Excel operations
                            time_str = f"{estimated_seconds:.1f} seconds" if estimated_seconds < 60 else f"{estimated_seconds/60:.1f} minutes"
                            
                            # Ask user on main thread
                            def _ask_user():
                                result = messagebox.askyesno(
                                    "UUID Fixes Required",
                                    f"Found {num_issues:,} UUIDs that need fixing.\n\n"
                                    f"Estimated time: {time_str}\n\n"
                                    f"Do you want to fix these UUIDs now?\n\n"
                                    f"Note: Choosing 'No' will enable SQLite mirroring without fixing UUIDs, "
                                    f"which may cause synchronization issues.",
                                    icon='warning',
                                    parent=self.master
                                )
                                
                                if result:
                                    # User chose to fix UUIDs
                                    self._perform_initial_sync_with_fixes()
                                else:
                                    # User chose to skip UUID fixes
                                    print("SQLite: User chose to skip UUID fixes. Proceeding without UUID validation.")
                                    self._perform_initial_sync_without_fixes()
                            
                            self.master.after(0, _ask_user)
                        else:
                            # No UUID issues, proceed normally
                            self.master.after(0, self._perform_initial_sync_with_fixes)
                    
                    except Exception as e:
                        print(f"SQLite: Error analyzing UUIDs: {e}")
                        traceback.print_exc()
                        self.master.after(0, lambda: messagebox.showerror(
                            "Error",
                            f"Failed to analyze UUIDs: {e}\n\nSQLite mirroring will be disabled.",
                            parent=self.master
                        ))
                        self.master.after(0, lambda: self.sqlite_mirror_enabled_var.set(False))
                
                # Start counting in background
                count_thread = threading.Thread(target=_count_and_ask, daemon=True)
                count_thread.start()

            except Exception as e:
                messagebox.showerror("SQLite Error", f"Failed to initialize SQLite database: {e}", parent=self.master)
                self.sqlite_mirror_enabled_var.set(False)
                if self.sqlite_manager:
                    self.sqlite_manager.close()
                self.sqlite_manager = None
        
        else:
            # --- DISABLE ---
            self.stop_auto_sync() # Stop the timer when mirroring is disabled
            if self.sqlite_manager:
                self.sqlite_manager.close()
                self.sqlite_manager = None
            self.update_status("SQLite mirroring disabled.")
    
    def _perform_initial_sync_with_fixes(self):
        """Perform initial sync with UUID validation and fixes."""
        def _sync_worker():
            self.update_status("Performing initial full sync to SQLite (with UUID fixes)...")
            header_finder = lambda path: self._find_header_row(path)
            success = self.sqlite_manager.full_sync(self.log_file_path, header_finder, skip_uuid_fixes=False)
            
            if success:
                self.update_status("Initial SQLite sync complete.")
                # Start the auto-sync timer only after successful sync
                self.master.after(0, self.start_auto_sync)
            else:
                self.update_status("Initial SQLite sync failed. Check console for details.")
        
        sync_thread = threading.Thread(target=_sync_worker, daemon=True)
        sync_thread.start()
    
    def _perform_initial_sync_without_fixes(self):
        """Perform initial sync WITHOUT UUID validation (skip UUID fixes)."""
        def _sync_worker():
            self.update_status("Performing initial full sync to SQLite (skipping UUID fixes)...")
            header_finder = lambda path: self._find_header_row(path)
            success = self.sqlite_manager.full_sync(self.log_file_path, header_finder, skip_uuid_fixes=True)
            
            if success:
                self.update_status("Initial SQLite sync complete (UUIDs not fixed).")
                # Start the auto-sync timer only after successful sync
                self.master.after(0, self.start_auto_sync)
            else:
                self.update_status("Initial SQLite sync failed. Check console for details.")
        
        sync_thread = threading.Thread(target=_sync_worker, daemon=True)
        sync_thread.start()
    
    def create_status_indicators(self):
        '''
        Creates the status indicators for monitoring and SQLite connection status.
        This method adds a frame below the main buttons to show the current status of monitoring and SQLite logging.
        '''
        # Create a frame for status indicators
        indicator_lf = ttk.LabelFrame(self.config_frame, text="Status")
        indicator_lf.grid(row=1, column=0, sticky="sew", pady=(10, 0))
        indicator_lf.columnconfigure(1, weight=1)

        # Monitoring Status
        ttk.Label(indicator_lf, text="Monitoring:", font=("Arial", 8, "bold")).grid(row=0, column=0, sticky="w", padx=4, pady=2)
        self.monitor_status_label = ttk.Label(indicator_lf, text="...", foreground="orange", font=("Arial", 8))
        self.monitor_status_label.grid(row=0, column=1, sticky="w", padx=4, pady=2)

        # SQLite Mirror Status
        ttk.Label(indicator_lf, text="SQLite Mirror:", font=("Arial", 8, "bold")).grid(row=1, column=0, sticky="w", padx=4, pady=2)
        sqlite_mirror_check = ttk.Checkbutton(
            indicator_lf,
            text="Enabled",
            variable=self.sqlite_mirror_enabled_var,
            command=self.toggle_sqlite_mirroring
        )
        sqlite_mirror_check.grid(row=1, column=1, sticky='w', padx=4, pady=2)
        ToolTip(sqlite_mirror_check, "If checked, mirrors the Excel log to an SQLite database file in the same folder.")

        # Always on Top Checkbox
        always_on_top_check = ttk.Checkbutton(
            indicator_lf,
            text="Always on Top",
            variable=self.always_on_top_var,
            command=self.toggle_always_on_top
        )
        always_on_top_check.grid(row=2, column=0, columnspan=2, sticky='w', padx=4, pady=(5, 2))
        ToolTip(always_on_top_check, "If checked, this window will always stay on top.")
      

    def create_status_bar(self):
        '''
        Creates a status bar at the bottom of the main window to display status messages.
        This method initializes a label that will show the current status of the application, such as monitoring status, database connection status, and other messages.
        '''
        self.status_var.set("Status: Ready")
        status_bar = ttk.Label(self.main_frame, textvariable=self.status_var, style="StatusBar.TLabel", anchor='w')
        status_bar.grid(row=2, column=0, columnspan=3, sticky="ew")

    def create_progress_bar(self):
        """Creates the progress bar and label, initially hidden."""
        self.progress_frame = ttk.Frame(self.main_frame)
        self.progress_frame.columnconfigure(1, weight=1)

        self.progress_label = ttk.Label(self.progress_frame, text="Scanning folders, please wait...")
        self.progress_label.grid(row=0, column=0, padx=(0, 10), sticky='w')
        
        self.progress_bar = ttk.Progressbar(self.progress_frame, mode='indeterminate')
        self.progress_bar.grid(row=0, column=1, sticky='ew')

        
        # Place the frame in the layout grid and then immediately remove it.
        # This ensures the GUI manager knows where to put it when we un-hide it later.
        self.progress_frame.grid(row=1, column=0, columnspan=3, sticky='ew', pady=(5, 2), padx=5)
        self.progress_frame.grid_remove()
    
    def show_progress_bar(self):
        """Displays and animates the progress bar."""
        if hasattr(self, 'monitoring_button') and self.monitoring_button:
            self.monitoring_button.config(state=tk.DISABLED)
        self.progress_frame.grid(row=1, column=0, columnspan=3, sticky='ew', pady=(5, 2), padx=5)
        self.progress_bar.start(10) # The number controls animation speed

    def hide_progress_bar(self):
        """Stops and hides the progress bar."""
        self.progress_bar.stop()
        self.progress_frame.grid_remove()
        if hasattr(self, 'monitoring_button') and self.monitoring_button:
            self.monitoring_button.config(state=tk.NORMAL)
        self.update_status("Initial folder scan complete.")
    
    # --- Synchronization Excel Log to SQLite Database ---
    def _find_header_row(self, excel_file, max_rows_to_scan=MAX_HEADER_SEARCH_ROW):
        """
        Scans the top N rows of an Excel sheet using xlwings to find the header.
        The header is identified by the presence of a specific required column.
        """
        print(f"Using xlwings to find header in: {excel_file}")
        app, book, sheet = None, None, None
        try:
            app = xw.App(visible=False)
            book = app.books.open(excel_file)
            sheet = book.sheets[0]

            data_block = sheet.range(f'A1:{chr(ord("A")+25)}{max_rows_to_scan}').value

            for idx, row in enumerate(data_block):
                if row is None: continue
                # Comparing as lowercase strings for robustness
                row_values_list = [str(v).lower().strip() for v in row if v is not None]
                current_row_headers = set(row_values_list)
                
                # Check for missing columns
                missing_cols = EXCEL_LOG_REQUIRED_COLS - current_row_headers
                
                if not missing_cols:
                    # All required columns are present in this row
                    return idx  # Return the zero-based index of the header row

            # If the loop finishes, the header was not found
            # Determine which columns are missing overall to provide the most helpful error.
            # (We cannot know for certain which ones were missing, but we'll list all required ones)
            
            
            # If a match isn't found, construct the most informative error message.
            missing_text = ', '.join(sorted(EXCEL_LOG_REQUIRED_COLS))
            
            raise ValueError(
                f"Crucial header columns not found in the first {max_rows_to_scan} rows of the log sheet. "
                f"Ensure the following headers (case-insensitive) exist: {missing_text}"
            )
            
            
        except Exception as e:
            # Re-raise the exception after printing the traceback for better debugging
            traceback.print_exc() 
            raise e
        finally:
            # Ensure the Excel process is always closed
            if book: book.close()
            if app: app.quit()
    
    
    # --- Status Bar and Indicators ---
    def update_status(self, message):
        '''
        Updates the status bar with a new message, including a timestamp.
        This method formats the message with the current time and ensures it does not exceed a certain length.
        Arguments:
        * message: The message to display in the status bar.
        '''

        # FUNCTION DEFINED INLINE
        def _update(): 
            timestamp = datetime.datetime.now().strftime("%H:%M:%S")
            max_len = 100
            display_message = message if len(message) <= max_len else message[:max_len-3] + "..."
            try:
                if self.status_var:
                    self.status_var.set(f"[{timestamp}] {display_message}")
            except tk.TclError:
                pass # Window might be closing

        if hasattr(self, 'master') and self.master.winfo_exists():
            try:
                self.master.after(0, _update)
            except tk.TclError:
                pass # Window might be destroyed between check and after call

    def set_window_icon(self):
        """
        Sets the window icon to OnlineLoggerLogo.ico.
        Handles both development (running from source) and production (PyInstaller executable).
        """
        try:
            # Determine the base path (handles PyInstaller's temporary folder)
            if getattr(sys, 'frozen', False):
                # Running as compiled executable
                base_path = sys._MEIPASS
            else:
                # Running from source - get script directory and go up one level
                base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            
            # Construct path to icon file
            icon_path = os.path.join(base_path, "_repositoryfiles", "OnlineLoggerLogo.ico")
            
            # Set the icon if file exists
            if os.path.exists(icon_path):
                self.master.iconbitmap(icon_path)
                print(f"Window icon set: {icon_path}")
            else:
                print(f"Warning: Icon file not found at {icon_path}")
        except Exception as e:
            print(f"Could not set window icon: {e}")
    
    def update_window_title(self):
        """
        Updates the main window title to show version and current project name.
        Format: 'Online Logger v{VERSION} - {PROJECT_NAME}'
        """
        title_parts = [f"Online Logger v{APP_VERSION}"]
        
        # Extract project name from settings file path if available
        if self.settings_file and os.path.exists(self.settings_file):
            project_name = os.path.splitext(os.path.basename(self.settings_file))[0]
            title_parts.append(project_name)
        
        final_title = " - ".join(title_parts)
        
        try:
            self.master.title(final_title)
        except tk.TclError:
            pass  # Window might be closing
    
    def update_monitor_indicator_text(self):
        """
        Updates the monitoring status label text and color based on the current
        state of the monitor threads, without restarting them.
        """
        # First, ensure the widget exists. This is a safeguard.
        if not hasattr(self, 'monitor_status_label') or not self.monitor_status_label or not self.monitor_status_label.winfo_exists():
            return

        is_active = any(observer.is_alive() for observer in self.monitors.values())
        
        try:
            if is_active:
                self.monitor_status_label.config(text="Active", foreground="green")
            else:
                self.monitor_status_label.config(text="Inactive", foreground="red")
        except tk.TclError:
            # This can happen if the widget is destroyed between the check and the config call
            pass

    def _find_closest_sfile(self, historic_dt):
        """
        Finds the S-File with the timestamp closest to, but before, the historic datetime.
        S-Files are expected in a monitored folder named 'S-File' and formatted as YYYYMMDD_HHMMSS_S.
        """
        # NOTE: This assumes your monitored folder for S-Files is named "S-File" in the settings.
        sfile_folder_key = "S-File" 
        sfile_folder_path = self.folder_paths.get(sfile_folder_key)

        if not sfile_folder_path or not os.path.isdir(sfile_folder_path):
            self.update_status(f"Warning: S-File folder '{sfile_folder_key}' not configured or found.")
            return None

        candidate_files = []
        # Walk through the directory and all subdirectories to find all files
        for root, _, files in os.walk(sfile_folder_path):
            for filename in files:
                # Parse filename to get datetime
                basename, _ = os.path.splitext(filename)
                try:
                    
                    # The format string now includes the literal "_S" at the end.
                    file_dt = datetime.datetime.strptime(basename, "%Y%m%d_%H%M%S_S")
                    
                    # Make the historic_dt timezone-unaware for a direct comparison
                    historic_dt_unaware = historic_dt.replace(tzinfo=None)

                    # Only consider files created *before* the historic event
                    if file_dt < historic_dt_unaware:
                        candidate_files.append((file_dt, os.path.join(root, filename)))
                except ValueError:
                    # Ignore any files that don't match the YYYYMMDD_HHMMSS_S format
                    continue
        
        if not candidate_files:
            self.update_status(f"No S-Files found before {historic_dt.strftime('%Y-%m-%d %H:%M:%S')}.")
            return None

        # From the valid candidates, find the one with the latest timestamp
        closest_file_tuple = max(candidate_files, key=lambda item: item[0])
        return closest_file_tuple[1] # Return the full path of the best match

    # --- Logging Actions ---
    def log_event(self, event_type, button_widget, txt_source_key="Main TXT"):
        '''
        This function is called when a standard event button is pressed (e.g., Log on, Log off, Event).
        It handles the logging of the event by calling the _perform_log_action method with appropriate parameters.
        Arguments:
        * event_type: The type of event being logged (e.g., "Log on", "Log off", "Event").
        * button_widget: The button widget that was pressed, used to temporarily disable it during processing.
        '''
        event_text_for_excel = None
        skip_files = False

        # Get the configuration for this specific main button
        config = self.main_button_configs.get(event_type, {})

        # Get the event text from the configuration, with a fallback
        event_text_for_excel = config.get("event_text", f"Default {event_type}")

        # Get the source key from the configuration
        source_key_for_log = config.get("txt_source_key", "Main TXT")
        
        
        skip_files = (event_type == "Event") # Still skip files only for the main "Event" button
            
        self._perform_log_action(event_type=event_type,
                                 event_text_for_excel=event_text_for_excel,
                                 triggering_button=button_widget,
                                 txt_source_key=source_key_for_log) # Use the new, configurable source key

    def log_custom_event(self, config, button_widget):
        """
        This function is called when a custom event button is pressed.
        It retrieves the button text and event text from the configuration, 
        then calls _perform_log_action to log the event.
        """
        button_text = config.get("text", "Unknown Custom")
        event_text_for_excel = config.get("event_text", f"{button_text} Triggered")
        txt_source_key = config.get("txt_source_key", "None")
        
        self._perform_log_action(
            event_type=button_text,
            event_text_for_excel=event_text_for_excel,
            triggering_button=button_widget,
            txt_source_key=txt_source_key
        )

    def trigger_manual_hourly_log_action(self, button_widget):
        """
        Triggers the hourly log function via a button press, including visual feedback.
        """
        # Ensure the button is enabled in settings before proceeding
        if not self.hourly_event_enabled_var.get():
            messagebox.showinfo("Disabled", "The 'Hourly KP Log' event must be enabled in Settings to use this manual trigger.", parent=self.master)
            return

        # Check if log file is configured/exists (basic check before threading)
        if not self.log_file_path or not os.path.exists(self.log_file_path):
            messagebox.showerror("Error", f"Excel Log file is missing or not configured:\n{self.log_file_path}", parent=self.master)
            return
            
        # Manually trigger the core logic function on a thread to prevent freezing
        # Note: self.trigger_hourly_log does the complex logic AND reschedules the next timer.
        # We only want the logic here, so we copy the core logic's parameters.
        
        # Disable button and update status immediately on the main thread
        original_text = button_widget['text']
        button_widget.config(state=tk.DISABLED, text="Working...")
        self.update_status("Processing 'Manual KP Log'...")
        
        # Reroute to the trigger function on a new thread (similar to other logs)
        def _manual_log_worker():
            try:
                # 1. Call the core function that does the work and generates the log text
                self.trigger_hourly_log_core() 
                
                # 2. Re-enable button and update status on the main thread
                self.master.after(0, lambda: self._re_enable_button_and_update_status(
                    button_widget, original_text, "Manual KP Log completed successfully."
                ))
            except Exception as e:
                self.master.after(0, lambda: self._re_enable_button_and_update_status(
                    button_widget, original_text, f"Manual KP Log failed: {e}"
                ))
                traceback.print_exc()

        log_thread = threading.Thread(target=_manual_log_worker, daemon=True)
        log_thread.start()
    
    def add_historic_event(self):
        """
        Adds a historic event by letting the user choose a file and enter a time
        in a single dialog. The log date is derived from the chosen file's 
        'last ' date.
        """
        # 1. Open the combined dialog to get all details from the user
        details = self._ask_for_historic_event_details_dialog()
        
        if not details:
            self.update_status("Historic event entry cancelled.")
            return

        file_to_search = details['file_path']
        user_time = details['time_obj']
        
        # Use the original time string from the dialog for the file search
        time_str_to_find = details['time_str'] 
        
        insert_sfile = details['insert_sfile']

        # 2. Construct the final datetime (unchanged)
        final_datetime = None
        try:
            file_mtime = os.path.getmtime(file_to_search)
            file_date = datetime.date.fromtimestamp(file_mtime)
            final_datetime = datetime.datetime.combine(file_date, user_time).replace(tzinfo=datetime.timezone.utc)
        except Exception as e:
            messagebox.showerror("File Error", f"Could not read the file's modification date:\n{e}", parent=self.master)
            return

        # 2.5 Handle S-File search (unchanged)
        additional_data_to_log = {}
        if insert_sfile:
            # ... (S-File logic remains the same)
            sfile_folder_key = "S-File" 
            sfile_column_name = self.folder_columns.get(sfile_folder_key)
            if not sfile_column_name:
                messagebox.showwarning("Configuration Error",
                                     f"The target column for '{sfile_folder_key}' is not configured in Folder Settings.",
                                     parent=self.master)
            else:
                self.update_status("Searching for corresponding S-File...")
                closest_sfile_path = self._find_closest_sfile(final_datetime)
                if closest_sfile_path:
                    sfile_name, _ = os.path.splitext(os.path.basename(closest_sfile_path))
                    additional_data_to_log[sfile_column_name] = sfile_name
                    self.update_status(f"Found S-File: {sfile_name}")
                else:
                    additional_data_to_log[sfile_column_name] = "N/A"
                    self.update_status("No matching S-File found.")

        # 3. Search for the time within the chosen file
        
        search_pattern = r"(?:^|[^A-Za-z0-9:])" + re.escape(time_str_to_find)
        
        self.update_status(f"Searching for time '{time_str_to_find}' in {os.path.basename(file_to_search)}...")
        found_line = self._search_file_for_line(file_to_search, search_pattern)
        
        if not found_line:
            messagebox.showinfo("Time Not Found",
                                f"The time '{time_str_to_find}' was not found within the selected file.",
                                parent=self.master)
            self.update_status(f"Time '{time_str_to_find}' not found.")
            return
            
        # 4. Parse and log the data
        parsed_data = self._parse_txt_line(found_line, source_key="Main TXT")
        if not parsed_data:
            messagebox.showerror("Parse Error", "Could not parse the found line. The line might be empty or malformed.", parent=self.master)
            self.update_status("Error parsing the found line.")
            return
        
        self.update_status(f"Found line for '{time_str_to_find}'. Logging...")
        self._perform_log_action(
            event_type="Historic Event",
            event_text_for_excel=f"Historic data for {final_datetime.strftime('%Y-%m-%d %H:%M')}",
            triggering_button=None,
            txt_source_key="Manual/Historic",
            override_txt_data=parsed_data,
            override_utc_datetime=final_datetime,
            skip_monitored_folders=True,
            # Pass the S-File data (if any) to the logging function
            additional_data=additional_data_to_log
           
        )
    
    
    def log_svp(self, event_type, button_widget, txt_source_key="Main TXT"):
        '''
        This function is called when the "SVP" button is pressed.
        It checks if the necessary configurations are set (log file, TXT folder, SVP folder path),
        and if so, it calls _perform_log_action to log the SVP event.
        Arguments:
        * event_type: necessary code to classify the event.
        * button_widget: The button widget that was pressed, used to temporarily disable it during processing.
        * txt_source_key: TXT file where metadata is obtained

        '''
        if not self.log_file_path or not self.txt_folder_path or "SVP" not in self.folder_paths:
            messagebox.showinfo("Info", "Please select log file, TXT folder, and configure SVP folder path/column in Settings.", parent=self.master)
            self.update_status("SVP Error: Configuration missing.")
            return
        if not self.folder_columns.get("SVP"):
            messagebox.showinfo("Info", "Please configure the 'Target Column' for SVP in Folder Settings.", parent=self.master)
            self.update_status("SVP Error: Target column missing.")
            return
        if self.log_file_path and not os.path.exists(self.log_file_path):
            messagebox.showerror("Error", f"Excel Log file does not exist:\n{self.log_file_path}", parent=self.master)
            self.update_status("SVP Error: Excel file missing.")
            return

        event_text = self.main_button_configs.get("SVP", {}).get("event_text", "SVP applied")
        self._perform_log_action(event_type=event_type,
                                 event_text_for_excel=event_text,
                                 triggering_button=button_widget,
                                 txt_source_key=txt_source_key) 

    def _perform_log_action(self, event_type, event_text_for_excel, triggering_button, txt_source_key, override_txt_data=None, override_utc_datetime=None, skip_monitored_folders=False, additional_data=None):
        """Initiates a logging action on a background thread to prevent GUI freezing."""
        # >>> The only change here is adding `additional_data=None` to the function signature
        
        original_text = None
        if triggering_button and isinstance(triggering_button, ttk.Button) and triggering_button.winfo_exists():
            original_text = triggering_button['text']
            triggering_button.config(state=tk.DISABLED, text="Working...")
        
        self.update_status(f"Processing '{event_type}'...")

        def _log_worker():
            """The function that runs on the background thread."""
            try:
                # -----------------------------------------------------------
                # Dynamically derive mapping configs based on source key
                # -----------------------------------------------------------
                # 1. Get the mapping configuration list for the button's source
                mapping_config = self.all_txt_mappings.get(txt_source_key, [])
                
                # 2. Build the lookup dictionaries needed by the rest of the logic
                # Only include entries that aren't marked to skip for TXT data processing
                txt_mapping_for_lookup = {cfg["field"]: cfg["column_name"] for cfg in mapping_config}

                # Combine all lookup dictionaries for a single source of truth for column names
                # This needs to include generated fields and static fields which are common.
                combined_configs = mapping_config + self.generated_fields_config + self.static_field_configs
                txt_field_columns = {cfg["field"]: cfg["column_name"] for cfg in combined_configs}
                
                # Use this dynamically generated map for everything below
                # The old self.txt_field_columns and self.txt_field_skips are no longer globals.

                row_data = {}
                
                # --- GENERATE UUID AT START OF LOGGING ---
                uuid_col = txt_field_columns.get("UUID")
                if uuid_col:
                    row_data[uuid_col] = str(uuid.uuid4())

                # --- DATA GATHERING ---
                if override_txt_data is not None:
                    row_data.update(override_txt_data)
                elif txt_source_key != "None":
                    source_folder_path = self._get_path_from_source_key(txt_source_key)
                    if source_folder_path and os.path.isdir(source_folder_path):
                        #  CALL: Pass the source key to the helper
                        txt_data = self._get_txt_data_from_source(source_folder_path, txt_source_key)
                        if txt_data:
                            row_data.update(txt_data)
                
                static_data_from_cells = self._get_static_excel_data()
                if static_data_from_cells:
                    row_data.update(static_data_from_cells)

                
                if not skip_monitored_folders:
                    latest_files_data = self.get_latest_files_data_fast()
                    if latest_files_data:
                        row_data.update(latest_files_data)

               
                # Add the extra data (like the S-File name) to the row
                if additional_data:
                    row_data.update(additional_data)
               

                # --- PROCESSING AND GENERATED FIELDS ---
                final_event_text = event_text_for_excel

                # >>> Reworked Log On/Log Off logic for robustness and feedback <<<
                if event_type == "Log on":
                    kp_col_name = txt_field_columns.get("KP")
                    if kp_col_name and kp_col_name in row_data and row_data[kp_col_name] is not None:
                        try:
                            # Attempt to parse KP and store it
                            kp_value = float(row_data[kp_col_name])
                            self.last_log_on_kp = kp_value
                            self.log_on_time = datetime.datetime.now(datetime.UTC) # Use timezone-aware time
                            self.update_status(f"Log On successful. Stored KP: {self.last_log_on_kp:.3f}")
                        except (ValueError, TypeError):
                            # Failed to parse the KP value
                            self.last_log_on_kp = None
                            self.log_on_time = None
                            self.update_status("Log On Warning: Could not parse KP value. Calculation disabled.")
                    else:
                        # KP column was not found in the parsed data
                        self.last_log_on_kp = None
                        self.log_on_time = None
                        self.update_status("Log On Warning: 'KP' data not found in source file. Calculation disabled.")

                elif event_type == "Log off" and self.calculate_logoff_values.get():
                    kp_col_name = txt_field_columns.get("KP")
                    # Check if we have the necessary data from a previous "Log on"
                    if self.last_log_on_kp is not None and self.log_on_time is not None:
                        if kp_col_name and kp_col_name in row_data and row_data[kp_col_name] is not None:
                            try:
                                current_kp = float(row_data[kp_col_name])
                                current_time = datetime.datetime.now(datetime.UTC)
                                
                                time_diff_seconds = (current_time - self.log_on_time).total_seconds()
                                distance_km = abs(current_kp - self.last_log_on_kp)
                                speed_knots = 0
                                if time_diff_seconds > 1: # Avoid division by zero
                                    distance_nm = distance_km / 1.852
                                    time_hours = time_diff_seconds / 3600
                                    speed_knots = distance_nm / time_hours
                                
                                final_event_text = f"Log off - Traveled: {distance_km:.3f} km @ {speed_knots:.2f} kts"
                            except (ValueError, TypeError):
                                self.update_status("Log Off Warning: Could not parse current KP. Using default text.")
                        else:
                            self.update_status("Log Off Warning: Could not find current KP data. Using default text.")
                        # Reset values after any Log Off attempt (success or failure) to prevent using stale data
                        self.last_log_on_kp = None 
                        self.log_on_time = None
                    else:
                        # Log on data was never stored, so just log the default text
                        self.update_status("Log Off Info: No prior 'Log On' data to calculate from.")

                # Add all other generated fields
                if override_utc_datetime:
                    utc_now = override_utc_datetime.replace(tzinfo=datetime.timezone.utc)
                else:
                    utc_now = datetime.datetime.now(datetime.UTC)
                
                offset_delta = datetime.timedelta(hours=self.time_offset_hours.get())
                local_time = utc_now + offset_delta

                def get_gen_col_name(field_name):
                    return txt_field_columns.get(field_name)

                dt_col = get_gen_col_name("Date-Time")
                if dt_col: row_data[dt_col] = utc_now.strftime("%Y-%m-%d %H:%M:%S")
                lt_col = get_gen_col_name("Local Time")
                if lt_col: row_data[lt_col] = local_time.strftime("%Y-%m-%d %H:%M:%S")
                event_col = get_gen_col_name("Event")
                if event_col: row_data[event_col] = final_event_text
                
                # (The rest of the function is unchanged)
                code_col = get_gen_col_name("Code")
                if code_col:
                    event_code_to_log = ""
                    if event_type in self.main_button_configs:
                        event_code_to_log = self.main_button_configs[event_type].get("event_code", "")
                    else:
                        for cfg in self.custom_button_configs:
                            if cfg['text'] == event_type:
                                event_code_to_log = cfg.get("event_code", "")
                                break
                    if event_code_to_log:
                        row_data[code_col] = event_code_to_log
                kp_ref_col = get_gen_col_name("KP Ref.")
                if kp_ref_col and txt_source_key != "None":
                    kp_ref_to_log = self.txt_source_aliases.get(txt_source_key, "")
                    if kp_ref_to_log:
                        row_data[kp_ref_col] = kp_ref_to_log
                
                color_tuple = self.button_colors.get(event_type, (None, None))
                row_color = color_tuple[0] if isinstance(color_tuple, tuple) and len(color_tuple) > 0 else None
                font_color = color_tuple[1] if isinstance(color_tuple, tuple) and len(color_tuple) > 1 else None
                excel_success, _, excel_message = self.save_to_excel_and_sqlite(row_data, row_color, font_color, txt_field_columns)
                message = f"'{event_type}' logged successfully. {excel_message}" if excel_success else f"Error logging '{event_type}'. {excel_message}"
                if triggering_button:
                    self.master.after(0, lambda: self._re_enable_button_and_update_status(triggering_button, original_text, message))
                else:
                    self.master.after(0, self.update_status, message)
            
            except Exception as thread_ex:
                traceback.print_exc()
                status_msg = f"'{event_type}' failed due to an unexpected error."
                self.master.after(0, lambda: messagebox.showerror("Thread Error", f"Critical error during logging: {thread_ex}", parent=self.master))
                if triggering_button:
                    self.master.after(0, lambda: self._re_enable_button_and_update_status(triggering_button, original_text, status_msg))
                else:
                    self.master.after(0, self.update_status, status_msg)

        log_thread = threading.Thread(target=_log_worker, daemon=True)
        log_thread.start()
    
    def _ask_for_historic_event_details_dialog(self):
        """
        Opens a single dialog for the user to select a file and enter a time.
        Returns a dictionary {'file_path': str, 'time_obj': datetime.time} or None.
        """
        dialog = Toplevel(self.master)
        dialog.title("Add Historic Event")
        dialog.transient(self.master)
        dialog.grab_set()
        dialog.resizable(False, False)

        frame = ttk.Frame(dialog, padding="15")
        frame.pack(fill='both', expand=True)
        frame.columnconfigure(1, weight=1)

        result = {}
        
        # --- Variables ---
        now = datetime.datetime.now()
        time_var = tk.StringVar(value=now.strftime('%H:%M:%S')) # Default to more precise format
        file_path_var = tk.StringVar(value="No file selected...")
        insert_sfile_var = tk.BooleanVar(value=False)

        # --- Widgets ---
        def browse_file():
            file_path = filedialog.askopenfilename(
                parent=dialog,
                title="Select the log file to search",
                filetypes=[("Data Files", "*.txt *.npd *.csv"), ("All files", "*.*")]
            )
            if file_path:
                file_path_var.set(file_path)

        # Row 0: File Selection (unchanged)
        ttk.Label(frame, text="Log File:").grid(row=0, column=0, sticky='w', pady=5, padx=5)
        file_entry = ttk.Entry(frame, textvariable=file_path_var, state="readonly", width=50)
        file_entry.grid(row=0, column=1, sticky='ew', pady=5, padx=5)
        browse_btn = ttk.Button(frame, text="Browse...", command=browse_file)
        browse_btn.grid(row=0, column=2, sticky='ew', pady=5, padx=5)

        # Row 1: Time Entry
        
        # Update the label to show both accepted formats
        ttk.Label(frame, text="Time to Find (HH:MM:SS or HH:MM):").grid(row=1, column=0, sticky='w', pady=5, padx=5)
        time_entry = ttk.Entry(frame, textvariable=time_var, width=15)
        time_entry.grid(row=1, column=1, sticky='w', pady=5, padx=5)

        # Row 2: S-File Checkbox (unchanged)
        sfile_check = ttk.Checkbutton(frame, text="Insert S-File ID", variable=insert_sfile_var)
        sfile_check.grid(row=2, column=0, columnspan=2, sticky='w', pady=(10, 5), padx=5)
        ToolTip(sfile_check, "If checked, finds the closest S-File before this time and logs its name.\nRequires 'S-File' folder to be configured in Settings.")

        # --- OK / Cancel Logic ---
        def on_ok():
            file_path = file_path_var.get()
            time_str = time_var.get().strip() # Get the user's input

            if not os.path.isfile(file_path):
                messagebox.showwarning("Input Error", "Please select a valid log file first.", parent=dialog)
                return
            
            # New logic to handle both HH:MM:SS and HH:MM
            t_obj = None
            try:
                # First, try parsing the more specific HH:MM:SS format
                t_obj = datetime.datetime.strptime(time_str, '%H:%M:%S').time()
            except ValueError:
                # If that fails, try parsing the HH:MM format
                try:
                    t_obj = datetime.datetime.strptime(time_str, '%H:%M').time()
                except ValueError:
                    # If both formats fail, show a single, clear error message
                    messagebox.showwarning("Invalid Format", "Please enter the time in HH:MM:SS or HH:MM format.", parent=dialog)
                    return
            
            # If either format succeeded, t_obj will have a value.
            # Now, populate the results dictionary.
            result['file_path'] = file_path
            result['time_obj'] = t_obj
            result['time_str'] = time_str # Also pass the original string for the file search
            result['insert_sfile'] = insert_sfile_var.get()
            dialog.destroy()
            
        
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=3, column=0, columnspan=3, pady=(10,0), sticky='e')
        ttk.Button(button_frame, text="OK", command=on_ok, style="Accent.TButton").pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT)
        
        dialog.wait_window()
        return result if result else None

    def _re_enable_button_and_update_status(self, button, original_text, status_message):
        """A helper function to run on the main thread after a background task completes."""
        if button and button.winfo_exists():
            button.config(state=tk.NORMAL)
            if original_text: button.config(text=original_text)
        self.update_status(status_message)

    def _get_path_from_source_key(self, source_key):
        """Translates a txt_source_key string to its corresponding folder path."""
        path_map = {
            "Main TXT": self.txt_folder_path,
            "TXT Source 2": self.txt_folder_path_set2,
            "TXT Source 3": self.txt_folder_path_set3,
            "TXT Source 4": self.txt_folder_path_set4,
            "TXT Source 5": self.txt_folder_path_set5
        }
        return path_map.get(source_key)

    def _search_file_for_line(self, file_path, search_pattern):
        """
        Searches a file line by line using a regular expression pattern.
        Returns the first matching line, or None if not found.
        """
        if not search_pattern:
            return None
        try:
            # Compile the regex for efficiency
            regex = re.compile(search_pattern)
            encodings_to_try = ['utf-8', 'latin-1', 'cp1252']
            for enc in encodings_to_try:
                try:
                    with open(file_path, "r", encoding=enc) as f:
                        for line in f:
                            # Use the more precise regex search
                            if regex.search(line):
                                return line.strip()
                    return None 
                except UnicodeDecodeError:
                    continue 
            self.update_status(f"Error: Could not decode file {os.path.basename(file_path)}.")
            return None
        except FileNotFoundError:
            self.update_status(f"Error: File not found for search: {file_path}")
            return None
        except Exception as e:
            self.update_status(f"Error reading file for search: {e}")
            return None

    def _parse_txt_line(self, line_str, source_key="Main TXT"):
        """
        Parses a single comma-separated line string into a data dictionary
        based on the current txt_mapping_config.
        """
        parsed_data = {}
        if not line_str:
            return parsed_data
        
        latest_line_parts = line_str.split(",")

        mapping_config = self.all_txt_mappings.get(source_key, [])
        if not mapping_config:
            print(f"Warning: No field mapping found for source key '{source_key}'.")
            return parsed_data

        for i, field_config in enumerate(mapping_config):
            excel_col = field_config["column_name"]
            if not field_config.get("skip", False):
                if i < len(latest_line_parts):
                    parsed_data[excel_col] = latest_line_parts[i].strip()
                else:
                    parsed_data[excel_col] = None
        return parsed_data

    # --- TXT Reading and Writting ---
    def _get_txt_data_from_source(self, folder_path, source_key="Main TXT"):
        """
        Reads and parses data from the latest TXT file based on txt_mapping_config.
        Returns a dictionary of the parsed data only.
        """
        mapping_config = self.all_txt_mappings.get(source_key, [])

        parsed_data = {}
        if not mapping_config:
            return parsed_data
        latest_txt_file_path = None
        if folder_path and os.path.exists(folder_path):
            latest_txt_file_path = self.find_latest_file_in_folder(folder_path, ".txt")
            if not latest_txt_file_path:
                latest_txt_file_path = self.find_latest_file_in_folder(folder_path, ".npd")
            if not latest_txt_file_path:
                latest_txt_file_path = self.find_latest_file_in_folder(folder_path, ".csv")

        if not latest_txt_file_path:
            return parsed_data

        try:
            lines = []
            encodings_to_try = ['utf-8', 'latin-1', 'cp1252']
            read_success = False
            for enc in encodings_to_try:
                try:
                    time.sleep(0.05)
                    with open(latest_txt_file_path, "r", encoding=enc) as file:
                        lines = file.readlines()
                    read_success = True
                    break
                except (IOError, UnicodeDecodeError):
                    continue

            if read_success and lines:
                last_line_str = lines[-1].strip()
                latest_line_parts = last_line_str.split(",")

                # Use the new, dedicated config for mapping
                for i, field_config in enumerate(mapping_config):
                    excel_col = field_config["column_name"]
                    if not field_config.get("skip", False):
                        if i < len(latest_line_parts):
                            parsed_data[excel_col] = latest_line_parts[i].strip()
                        else:
                            parsed_data[excel_col] = None
        except Exception as e:
            print(f"Major error during TXT parsing: {e}")

        return parsed_data

    def get_latest_files_data_fast(self):
        """
        FAST version that collects latest file data by reading directly from the
        watchdog's global cache, avoiding slow disk scans.
        Ensures the correct filename format is logged based on the user's 'Log Ext.' setting.
        """
        data = {}
        threshold = self.active_logging_threshold_seconds.get()
        current_time = time.time()

        for folder_name, column_name in self.folder_columns.items():
            
            if self.folder_skips.get(folder_name, False) or not column_name:
                continue

            latest_file_path = folder_cache.get(folder_name)

            if latest_file_path and os.path.exists(latest_file_path):
                
                is_excluded_from_threshold = (folder_name == "SVP")
                file_mtime = os.path.getmtime(latest_file_path)
                is_active = is_excluded_from_threshold or ((current_time - file_mtime) <= threshold)
                
                # Retrieve the state of the "Log Ext.?" BooleanVar
                log_full_filename_var = self.folder_log_ext_vars.get(folder_name, None)
                log_full_filename = False
                if hasattr(log_full_filename_var, 'get'):
                    log_full_filename = log_full_filename_var.get()
                
                # --- LOGGING VALUE DETERMINATION ---
                logged_value = "" # Default to blank if inactive
                
                if is_active:
                    # Case 1: Log 'X' if option is ticked
                    if self.folder_log_x_instead.get(folder_name, False):
                        logged_value = "X"
                    else:
                        # Case 2: Log Filename (with or without extension)
                        filename = os.path.basename(latest_file_path)
                        if log_full_filename:
                            logged_value = filename  # Full name with extension
                        else:
                            # Split and join to get filename without extension
                            filename_without_ext, _ = os.path.splitext(filename)
                            logged_value = filename_without_ext
                
                # Assign the determined value to the output dictionary
                data[column_name] = logged_value
            
            else:
                # If no file is found in the cache or on disk
                data[column_name] = "N/A"
                
        return data

    def find_latest_file_in_folder(self, folder_path, extension=""):
        '''Finds the most recent file with the specified extension in the given folder AND ALL SUBFOLDERS.
            Arguments:
            * folder_path: The path to the top-level folder to search.
            * extension: The file extension to look for. If empty, finds the latest of any file.
            Returns:
            * The path (as a string) to the most recent file, or None if no such file exists.
        '''
        try:
            p = Path(folder_path)
            
            # Handles empty extension and formats the glob pattern
            if extension:
                # Ensure the extension starts with a dot for the glob pattern
                glob_pattern = f'*.{extension.lstrip(".")}'
            else:
                # If no extension is specified, search for any file
                glob_pattern = '*'
            
            # Use rglob to find all matching files recursively, then filter for actual files
            files = [f for f in p.rglob(glob_pattern) if f.is_file()]
            
            if not files:
                return None

            # Find the file with the latest modification time
            latest_file = max(files, key=lambda f: f.stat().st_mtime)
            
            # Return the path as a string to maintain consistency with os.path functions
            return str(latest_file)
        
        except FileNotFoundError:
            return None # Folder doesn't exist
        except Exception:
            return None # Other potential errors

    def _get_txt_file_data_for_preview(self, source_key):
        """
        Reads the latest line from the source folder associated with the key.
        Returns (latest_file_path, latest_line_parts) or (None, None).
        """
        folder_path = self._get_path_from_source_key(source_key)
        if not folder_path or not os.path.isdir(folder_path):
             return None, None
             
        latest_file = self.find_latest_file_in_folder(folder_path, ".txt")
        if not latest_file:
            latest_file = self.find_latest_file_in_folder(folder_path, ".npd")
        if not latest_file:
            latest_file = self.find_latest_file_in_folder(folder_path, ".csv")

        if not latest_file:
            return None, None

        try:
            # Use 'r' mode to open the file
            with open(latest_file, "r", encoding='utf-8', errors='ignore') as f:
                lines = f.readlines()
            if lines:
                # Use ',' as the separator for TXT/NPD/CSV data
                data_parts = lines[-1].strip().split(',') 
                return latest_file, data_parts
            return latest_file, []
        except Exception as e:
            print(f"Error reading preview file for {source_key}: {e}")
            return latest_file, None # Return None for malformed data

    def preview_data_file(self):
        """
        REMOVED: The original preview_data_file used to exist here. 
        It is being deleted as per the project plan, as the function is now handled
        by the new dialog's method. This placeholder must be deleted.
        """
        # This function is now OBSOLETE and should be removed from DataLoggerGUI.
        # If it is still present in your code, remove it entirely.
        pass
    
    def _get_static_excel_data(self):
        """
        Reads data from specific cells in the Excel log file based on the
        "='SheetName'!Cell" syntax in the data mapping configuration.
        """
        static_data = {}
        # Filter for configs that use the static cell lookup syntax
        cell_lookup_configs = self.static_field_configs

        if not cell_lookup_configs:
            return static_data  # Return empty if no lookups are configured

        app, workbook, opened_new_app = None, None, False
        try:
            # This logic connects to an existing instance or opens a new one
            target_norm_path = os.path.normcase(os.path.abspath(self.log_file_path))
            for running_app in xw.apps:
                for wb in running_app.books:
                    try:
                        if os.path.normcase(os.path.abspath(wb.fullname)) == target_norm_path:
                            workbook, app = wb, running_app
                            break
                    except Exception: continue
                if workbook: break
            
            if workbook is None:
                app = xw.App(visible=False)
                opened_new_app = True
                workbook = app.books.open(self.log_file_path, read_only=True)

            # Process each defined cell lookup
            for config in cell_lookup_configs:
                lookup_str = config["column_name"]
                excel_col_key = config["field"] # The desired Excel column name
                description = config.get("description", "") # <-- NEW: Get the description

                if config.get("skip"): # <-- Check for skip flag
                    static_data[excel_col_key] = "Skipped" # Or some other indicator
                    continue

                try:
                    # Parse the syntax: ='SheetName'!CellRef
                    # Using a more robust regex for parsing
                    match = re.match(r"='?([^'!]+)'?!([A-Z]+\d+)", lookup_str, re.IGNORECASE)
                    if not match:
                        print(f"Warning: Invalid cell lookup syntax '{lookup_str}'. Skipping.")
                        continue
                    
                    sheet_name, cell_ref = match.groups()
                    sheet = workbook.sheets[sheet_name]
                    value = sheet.range(cell_ref).value
                    
                    static_data[excel_col_key] = value
                    #print(f"Read '{value}' from {sheet_name}!{cell_ref} for mapping '{excel_col_key}'") #DEBUG

                except Exception as e:
                    print(f"Error reading from Excel cell for lookup '{lookup_str}': {e}")
            
            return static_data

        except Exception as e:
            print(f"Could not open or connect to Excel to read static data: {e}")
            return {} # Return empty on major error
        finally:
            # Only quit the app if this function started it
            if app is not None and opened_new_app:
                try:
                    app.quit()
                except Exception: 
                    pass

    def save_to_excel_and_sqlite(self, row_data, row_color=None, font_color=None, txt_field_columns=None):
        """
        Saves a single row of data to the open Excel file via xlwings
        and mirrors the new row to the SQLite database if enabled.
        
        : Forces date/time column to be written as a datetime object
                  to ensure correct Excel formatting.
        """
        if not self.log_file_path or not os.path.exists(self.log_file_path):
            return False, False, "Excel: Path Invalid."

        excel_message = "Excel: Fail."
        sqlite_message = "SQLite: Skipped."
        success_excel = False
        success_sqlite = False
        next_row = -1
        header_values = [] 
        
        # --- Part 1: Save to Excel ( for datetime write) ---
        try:
            wb = xw.Book(self.log_file_path)
            sheet = wb.sheets[0]
            header_row_index = -1
            
            # Search for the header row (omitted code for brevity)
            for i in range(1, MAX_HEADER_SEARCH_ROW + 1):
                row_values_list = sheet.range(f'A{i}').expand('right').value
                if not row_values_list: continue
                current_row_headers = {str(h).lower().strip() for h in row_values_list if h is not None}
                missing_cols = EXCEL_LOG_REQUIRED_COLS - current_row_headers
                
                if not missing_cols:
                    header_row_index = i
                    header_values = [h if h is not None else '' for h in row_values_list]
                    break
            
            if header_row_index == -1:
                # ... (ValueError raising logic remains unchanged)
                missing_text = ', '.join(sorted(EXCEL_LOG_REQUIRED_COLS))
                raise ValueError(
                    f"Required header row not found. Missing columns: {missing_text}."
                )

            header_map = {str(h).lower(): i for i, h in enumerate(header_values) if h}
            
            # Determine the next empty row for data insertion
            last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
            next_row = max(last_row, header_row_index) + 1

            output_data = [None] * len(header_values)
            
            # Identify the column name for the "Date-Time" field from settings
            dt_col_name = txt_field_columns.get("Date-Time") # <--- 
            dt_col_name_lower = str(dt_col_name).lower()
            dt_col_index = -1

            # --- Map Data to Output Row & Force Date Conversion for Excel ---
            for col_name, value in row_data.items():
                col_name_lower = str(col_name).lower()
                if col_name_lower in header_map:
                    col_idx = header_map[col_name_lower]
                    
                    if col_name_lower == dt_col_name_lower and value:
                        # CRITICAL: Convert the standardized date string back to a Python datetime object
                        # This ensures xlwings writes the correct numeric date value to Excel.
                        try:
                            value = datetime.datetime.strptime(str(value), "%Y-%m-%d %H:%M:%S")
                            dt_col_index = col_idx + 1 # Store 1-based column index for formatting later
                        except (ValueError, TypeError):
                            # If conversion fails, keep the original value (will likely be a string or None)
                            pass 
                    
                    output_data[col_idx] = value
            
            target_range = sheet.range(f"A{next_row}").resize(1, len(output_data))
            target_range.value = output_data

            # --- Apply Date/Time Format to the relevant column ---
            if dt_col_index > 0:
                dt_cell_range = sheet.cells(next_row, dt_col_index)
                
                # Apply the standard Excel date/time format
                # Using this format ensures both date and time are visible
                dt_cell_range.number_format = 'yyyy-mm-dd hh:mm:ss' 
            
            
            if row_color or font_color:
                format_range = sheet.range((next_row, 1), (next_row, len(header_map)))
                if row_color: format_range.color = row_color
                if font_color: format_range.font.color = font_color
                
            wb.save()
            excel_message = "Excel: OK."
            success_excel = True
            
        except ValueError as ve:
            print(f"Header Error during Excel Save: {ve}")
            traceback.print_exc()
            excel_message = f"Excel: Fail (Missing Header: {ve})."
            return False, False, f"{excel_message} {sqlite_message}"
            
        except Exception as e:
            traceback.print_exc()
            excel_message = f"Excel: Fail ({type(e).__name__})."
            return False, False, f"{excel_message} {sqlite_message}"

        # --- Part 2: Mirror to SQLite ---
        if success_excel and self.sqlite_mirror_enabled_var.get() and self.sqlite_manager:
            try:
                # Prepare data dictionary with original Excel headers
                row_data_for_sqlite = {}
                for header, value in zip(header_values, output_data):
                    if header:  # Skip empty headers
                        # Convert datetime objects to strings for SQLite
                        if isinstance(value, datetime.datetime):
                            value = value.strftime("%Y-%m-%d %H:%M:%S")
                        row_data_for_sqlite[header] = value
                
                # Use the new add_single_row method (fast, no full sync)
                # Pass excel_path so table name can be derived if needed
                # Pass excel_row (next_row) so the Excel row position is recorded
                success_sqlite = self.sqlite_manager.add_single_row(
                    row_data_for_sqlite, 
                    excel_path=self.log_file_path,
                    excel_row=next_row
                )
                
                if success_sqlite:
                    sqlite_message = "SQLite: OK."
                else:
                    sqlite_message = "SQLite: Fail."
                    print("SQLite: Failed to add row. Check console for details.")
                    
            except Exception as e:
                sqlite_message = f"SQLite: Fail ({type(e).__name__})."
                print(f"Error during SQLite single row add: {e}")
                traceback.print_exc()

        return success_excel, success_sqlite, f"{excel_message} {sqlite_message}"

    # --- Settings Saving and Loading ---
    def save_settings(self):
        '''Saves the current settings to the JSON file.'''
        print("\n--- Saving Settings ---")
        if not self.settings_file or not isinstance(self.settings_file, str):
            messagebox.showwarning(
                "No Project Selected",
                "There is no active project file to save to. Go to Settings → Projects and use 'Save As...' to create one.",
                parent=self.master
            )
            self.update_status("No active project file. Use Projects → Save As...")
            return
        colors_to_save = {}
        for key, (bg_color, font_color) in self.button_colors.items():
            if bg_color or font_color:
                colors_to_save[key] = (bg_color, font_color)
        settings = {
            "log_file_path": self.log_file_path,
            "sqlite_db_path": self.sqlite_db_path,
            "time_offset_hours": self.time_offset_hours.get(),
            "txt_folder_path": self.txt_folder_path,
            "txt_folder_path_set2": self.txt_folder_path_set2,
            "txt_folder_path_set3": self.txt_folder_path_set3,
            "txt_folder_path_set4": self.txt_folder_path_set4,
            "txt_folder_path_set5": self.txt_folder_path_set5,
            "all_txt_mappings": self.all_txt_mappings,
            "generated_fields_config": self.generated_fields_config,
            "static_field_configs": self.static_field_configs,
            "folder_paths": self.folder_paths,
            "folder_columns": self.folder_columns,
            "file_extensions": self.file_extensions,
            "folder_skips": self.folder_skips,
            "folder_log_x_instead": self.folder_log_x_instead,
            "folder_log_ext_vars": {k: v.get() for k, v in self.folder_log_ext_vars.items()},
            "num_custom_buttons": self.num_custom_buttons,
            "custom_button_configs": self.custom_button_configs,
            "custom_button_tab_groups": self.custom_button_tab_groups,
            "button_colors": colors_to_save,
            "always_on_top": self.always_on_top_var.get(),
            "active_logging_threshold_seconds": self.active_logging_threshold_seconds.get(),
            "new_day_event_enabled": self.new_day_event_enabled_var.get(),
            "hourly_event_enabled": self.hourly_event_enabled_var.get(),
            "hourly_log_txt_source_key": self.hourly_log_txt_source_key.get(),
            "main_button_configs": self.main_button_configs,
            "txt_source_aliases": self.txt_source_aliases,
            "calculate_logoff_values": self.calculate_logoff_values.get(),
            "event_codes": self.event_codes,
            "auto_sync_enabled": self.auto_sync_enabled_var.get(),
            "auto_sync_interval_min": self.auto_sync_interval_min_var.get(),
            "app_version": APP_VERSION  # Save current application version
        }
        try:
            with open(self.settings_file, 'w') as f:
                json.dump(settings, f, indent=4)
            print(f"Settings successfully saved to {self.settings_file}")
            self.update_status("Settings saved.")
        except Exception as e:
            print(f"Error saving settings: {e}")
            messagebox.showerror("Save Error", f"Could not save settings to {self.settings_file}:\n{e}", parent=self.master)
            self.update_status("Error saving settings.")

    def revert_to_defaults(self):
        """
        Loads the blank project template into memory and updates the UI.
        Does not overwrite or create any project file. The active project path
        remains unset until the user chooses Save As in the Projects tab.
        """
        print("\n--- Restoring Blank Project Template ---")
        template_path = os.path.join(os.getcwd(), PROJECT_TEMPLATE_FILE)
        if not os.path.exists(template_path):
            raise FileNotFoundError(
                f"The blank project template '{PROJECT_TEMPLATE_FILE}' was not found."
            )

        # Temporarily point to template to reuse load_settings logic
        prev_settings_path = self.settings_file
        self.settings_file = template_path
        self.load_settings()
        # After loading, clear active project so subsequent saves require explicit path
        self.settings_file = None
        # Refresh the main GUI
        self.update_custom_buttons()
        print("--- Blank Project Template Restored Successfully ---")

    def load_settings(self):
        '''Loads settings from the JSON file and updates the GUI variables accordingly.'''
        print("\n--- Loading Settings ---")

        try:
            if os.path.exists(self.settings_file):
                print(f"Loading Settings from: {self.settings_file}")
                with open(self.settings_file, 'r') as f:
                    settings = json.load(f)
                
                # --- Load Main Settings ---
                self.all_txt_mappings = settings.get("all_txt_mappings", self.all_txt_mappings)
                self.log_file_path = settings.get("log_file_path")
                self.sqlite_db_path = settings.get("sqlite_db_path")
                self.time_offset_hours.set(settings.get("time_offset_hours", 0.0))
                self.txt_folder_path = settings.get("txt_folder_path")
                self.txt_folder_path_set2 = settings.get("txt_folder_path_set2")
                self.txt_folder_path_set3 = settings.get("txt_folder_path_set3")
                self.txt_folder_path_set4 = settings.get("txt_folder_path_set4")
                self.txt_folder_path_set5 = settings.get("txt_folder_path_set5")

                # Check for old format (single config list) and migrate if necessary
                if "txt_mapping_config" in settings and not settings.get("all_txt_mappings"):
                    print("Old single mapping format detected. Migrating to 'Main TXT'...")
                    # Retrieve the old config list and save it to the correct new structure key
                    old_config_list = settings.get("txt_mapping_config") 
                    self.all_txt_mappings["Main TXT"] = old_config_list
                    if old_config_list is not None:
                        self.all_txt_mappings["Main TXT"] = old_config_list
                    
                # Ensure all required keys are present, even if empty
                for key in TXT_FILES_KEYS[1:]:
                    if key not in self.all_txt_mappings:
                        self.all_txt_mappings[key] = []

                # --- Load the three separate config lists ---
                self.generated_fields_config = settings.get("generated_fields_config", self.generated_fields_config)
                self.static_field_configs = settings.get("static_field_configs", [])

                # --- Handle backwards compatibility for old settings files ---
                if "txt_field_columns_config" in settings and not settings.get("txt_mapping_config"):
                    all_configs = settings["txt_field_columns_config"]
                    print("Old settings format detected. Migrating to new format...")
                    generated_fields_set = {"Date-Time", "Local Time", "Event", "Code", "KP Ref."}
                    self.generated_fields_config = [c for c in all_configs if c.get("field") in generated_fields_set]
                    main_txt_config = [c for c in all_configs if c.get("field") not in generated_fields_set and not str(c.get("column_name", "")).startswith('=')]
                    self.all_txt_mappings["Main TXT"] = main_txt_config # <--- 
                    
                    self.static_field_configs = [c for c in all_configs if str(c.get("column_name", "")).startswith('=')]
                
                # --- Add missing 'source' keys for robustness ---
                default_sources = {
                    "Date-Time": "PC Time (UTC)", "Local Time": "PC Time + Offset",
                    "Event": "Button", "Code": "Button", "KP Ref.": "Source Alias"
                }
                for item in self.generated_fields_config:
                    if 'source' not in item or not item['source']:
                        item['source'] = default_sources.get(item['field'], 'Unknown')

                # --- Re-derive the lookup dictionaries from ALL THREE lists ---
                main_txt_config = self.all_txt_mappings.get("Main TXT", [])
                combined_configs = main_txt_config + self.generated_fields_config + self.static_field_configs
                self.txt_field_columns = {cfg["field"]: cfg["column_name"] for cfg in combined_configs}
                self.txt_field_skips = {cfg["field"]: cfg.get("skip", False) for cfg in combined_configs}

                # --- Load Remaining Settings (Folder, Button, etc.) ---
                # Load event codes from the same project JSON; fallback to separate file if not present (backward compatible)
                loaded_event_codes = settings.get("event_codes", None)
                if isinstance(loaded_event_codes, dict):
                    self.event_codes = loaded_event_codes
                else:
                    self.load_event_codes()
                loaded_main_configs = settings.get("main_button_configs", {})
                for btn_name, default_conf in self.main_button_configs.items():
                    default_conf.update(loaded_main_configs.get(btn_name, {}))

                self.folder_paths.clear()
                self.folder_paths.update(settings.get("folder_paths", {}))
                self.folder_columns.clear()
                self.folder_columns.update(settings.get("folder_columns", {}))
                self.file_extensions.clear()
                self.file_extensions.update(settings.get("file_extensions", {}))
                self.folder_skips.clear()
                self.folder_skips.update(settings.get("folder_skips", {}))
                self.folder_log_x_instead = settings.get("folder_log_x_instead", {})
                loaded_log_exts = settings.get("folder_log_ext_vars", {})
                self.folder_log_ext_vars = {k: tk.BooleanVar(value=v) for k, v in loaded_log_exts.items()}
                self.num_custom_buttons = settings.get("num_custom_buttons", 3)
                loaded_configs = settings.get("custom_button_configs", [])
                
                updated_custom_configs = []
                for i in range(self.num_custom_buttons):
                    config = loaded_configs[i] if i < len(loaded_configs) else {"text": f"Custom {i+1}", "event_text": f"Custom {i+1} Event"}
                    config["txt_source_key"] = config.get("txt_source_key", "None")
                    config["tab_group"] = config.get("tab_group", "Main")
                    config["event_code"] = config.get("event_code", "")
                    updated_custom_configs.append(config)
                self.custom_button_configs = updated_custom_configs

                self.custom_button_tab_groups = sorted(list(set(["Main"] + settings.get("custom_button_tab_groups", []))))
                self.custom_button_tab_groups = [g for g in self.custom_button_tab_groups if g]

                # --- COMPLETE: This block now correctly loads and applies colors ---
                loaded_colors = settings.get("button_colors", {})
                # JSON saves tuples as lists, so we convert them back
                for key, value in loaded_colors.items():
                    if isinstance(value, list):
                        loaded_colors[key] = tuple(value)
                self.button_colors.update(loaded_colors) # Merge saved colors into defaults

                # --- Load Miscellaneous Settings ---
                self.always_on_top_var.set(settings.get("always_on_top", False))
                self.active_logging_threshold_seconds.set(settings.get("active_logging_threshold_seconds", 15))
                self.calculate_logoff_values.set(settings.get("calculate_logoff_values", True))
                self.new_day_event_enabled_var.set(settings.get("new_day_event_enabled", True))
                self.hourly_event_enabled_var.set(settings.get("hourly_event_enabled", True))
                self.hourly_log_txt_source_key.set(settings.get("hourly_log_txt_source_key", "Main TXT"))
                self.txt_source_aliases = settings.get("txt_source_aliases", self.txt_source_aliases)
                self.auto_sync_enabled_var.set(settings.get("auto_sync_enabled", True))
                self.auto_sync_interval_min_var.set(settings.get("auto_sync_interval_min", 15))

                print("Settings loaded successfully")
                self.update_status("Settings loaded.")
            else:
                self.update_status("Settings file not found. Using defaults.")
                print("Settings file not found, using defaults.")
                # When no file is found, derive the lookup dictionaries from defaults
                main_txt_config = self.all_txt_mappings.get("Main TXT", [])
                combined_configs = main_txt_config + self.generated_fields_config + self.static_field_configs
                self.txt_field_columns = {cfg["field"]: cfg["column_name"] for cfg in combined_configs}
                self.txt_field_skips = {cfg["field"]: cfg.get("skip", False) for cfg in combined_configs}

        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Load Error", f"Could not load settings from {self.settings_file}:\n{e}\n\nUsing default settings.", parent=self.master)
            self.init_variables() # Fallback to hard-coded defaults on critical error
        finally:
            # Ensure UI state is correct after loading
            self.master.wm_attributes("-topmost", self.always_on_top_var.get())
            if hasattr(self, 'main_frame') and self.main_frame:
                self.update_custom_buttons()
                        

    def load_event_codes(self):
        """Loads the event codes from its dedicated JSON file."""

        print(f"Loading Event Codes from {self.event_codes_file}...")
        if os.path.exists(self.event_codes_file):
            try:
                with open(self.event_codes_file, 'r') as f:
                    self.event_codes = json.load(f)
            except (json.JSONDecodeError, Exception) as e:
                print(f"Error loading event codes file: {e}")
                self.event_codes = {} # Reset to empty on error
                messagebox.showerror("Load Error", f"Could not load or parse the event codes file:\n{self.event_codes_file}\n\nError: {e}", parent=self.master)
        else:
            print("Event codes file not found. Using empty set.")
            self.event_codes = {}

    # --- Settings Window Interaction ---
    def open_settings(self):
        '''Open the settings window. If it already exists, bring it to the front.'''
        if hasattr(self, 'settings_window_instance') and self.settings_window_instance and self.settings_window_instance.winfo_exists():
            self.settings_window_instance.lift()
            self.settings_window_instance.focus_set()
            return

        settings_top_level = tk.Toplevel(self.master)
        if self.custom_inline_editor_window and self.custom_inline_editor_window.winfo_exists():
            settings_top_level.transient(self.custom_inline_editor_window)
        else:
            settings_top_level.transient(self.master)
        
        settings_top_level.grab_set()

        # Store references to both the Toplevel widget and the class instance
        self.settings_window_instance = settings_top_level
        self.settings_gui_instance = SettingsWindow(settings_top_level, self)
        self.settings_gui_instance.load_settings()
        
        self.master.wait_window(settings_top_level)

        # Clear references after the window is closed
        self.settings_window_instance = None
        self.settings_gui_instance = None

        def startup_settings(self):
            '''Open settings by default in the startup of the app'''

            self.open_settings()
    
    def startup_settings(self):
        '''Open settings by default in the startup of the app'''
        self.open_settings()

    def update_custom_buttons(self):
        '''Update the custom buttons in the main GUI based on current settings.'''

        # Check for one of the new main frames to ensure the UI has been initialized.
        if hasattr(self, 'custom_buttons_frame'):
            # Call create_main_buttons without arguments, as it now handles all frames.
            self.create_main_buttons()
            self.create_status_indicators()
            # update the text of the newly created monitor status label
            self.update_monitor_indicator_text() 
            self.master.update_idletasks()

    def toggle_always_on_top(self):
        """Toggles the 'always on top' state of the main window based on the checkbox."""
        is_on_top = self.always_on_top_var.get()
        self.master.wm_attributes("-topmost", is_on_top)

    # --- Profile state persistence ---
    def load_last_project_path(self):
        """Reads the last-used project path from a small state file. Returns a path or None."""
        state_path = os.path.join(os.getcwd(), PROJECT_STATE_FILE)
        if not os.path.exists(state_path):
            return None
        try:
            with open(state_path, 'r', encoding='utf-8') as f:
                obj = json.load(f)
            path = obj.get('current_project_path')
            return path if path else None
        except Exception:
            return None

    def persist_current_project_path(self, path=None):
        """Persists the provided (or current) project path to the state file."""
        state_path = os.path.join(os.getcwd(), PROJECT_STATE_FILE)
        try:
            to_write = {
                'current_project_path': (path or getattr(self, 'current_project_path', None))
            }
            os.makedirs(os.path.dirname(state_path), exist_ok=True)
            with open(state_path, 'w', encoding='utf-8') as f:
                json.dump(to_write, f, indent=4)
        except Exception:
            # Non-fatal if we cannot persist; ignore
            pass

    def set_active_project(self, path):
        """Sets the active project for this session and persists it for next startup."""
        if not path:
            return
        self.settings_file = path
        self.current_project_path = path
        self.persist_current_project_path(path)
        self.update_window_title()  # Refresh title with new project name

    # --- Monitoring ---

    def _manage_scan_threads(self, threads_to_watch):
        """Waits for all initial scan threads to complete, then hides the progress bar."""
        for t in threads_to_watch:
            t.join() # This waits for a thread to finish
        
        # When all threads are done, schedule the UI update on the main thread
        self.master.after(0, self.hide_progress_bar)
        
    def toggle_monitoring(self):
        """Starts or stops the folder monitoring threads based on the current state."""
        if self.is_monitoring:
            self.stop_monitoring()
        else:
            self.start_monitoring()
        self.update_monitoring_button_ui()

    def update_monitoring_button_ui(self):
        """Updates the text of the monitoring button."""
        if not hasattr(self, 'monitoring_button') or not self.monitoring_button or not self.monitoring_button.winfo_exists():
            return

        if self.is_monitoring:
            self.monitoring_button.config(text="Stop Folder Monitoring")
        else:
            self.monitoring_button.config(text="Start Folder Monitoring")

    def _clear_monitors(self):
        """Stops and clears all active monitoring threads."""
        active_monitors = list(self.monitors.items())
        if not active_monitors:
            return

        print("Stopping all active monitors...")
        for name, monitor_observer in active_monitors:
            try:
                if monitor_observer.is_alive():
                    monitor_observer.stop()
            except Exception as e:
                print(f"Error signalling monitor {name} to stop: {e}")

        for name, monitor_observer in active_monitors:
            try:
                monitor_observer.join(timeout=1.0)
            except Exception as e:
                print(f"Error joining monitor {name}: {e}")

        self.monitors.clear()
        folder_cache.clear()
        print("Cleared existing monitors and folder cache.")

    def start_folder_monitoring(self, folder_name, folder_path, file_extension):
        '''Start monitoring a specific folder...'''
        try: 
            if not os.path.isdir(folder_path):
                print(f"Error: Path '{folder_path}' is not a valid directory for monitoring '{folder_name}'.")
                return None # 
            os.listdir(folder_path) # Check permissions
        except Exception as e: 
            print(f"Error accessing directory '{folder_path}' for monitoring '{folder_name}': {e}")
            return None 
        
        try:
            event_handler = FolderMonitor(folder_path, folder_name, self, file_extension)
            observer = PollingObserver(timeout=1)
            observer.schedule(event_handler, folder_path, recursive=True)
            observer.start()
            self.monitors[folder_name] = observer
            print(f"Successfully started recursive monitoring for {folder_name} at {folder_path} (ext: {file_extension}).")
            return event_handler 
        except Exception as e: 
            print(f"Failed to start watchdog monitor for {folder_name} at {folder_path}: {e}")
            return None # 

    def stop_monitoring(self):
        """Public method to stop all monitoring."""
        if hasattr(self, 'progress_bar') and self.progress_bar.winfo_ismapped():
             self.hide_progress_bar() 
        self._clear_monitors()
        self.is_monitoring = False
        self.update_status("Monitoring stopped.")
        self.update_monitor_indicator_text()
        print("--- Monitoring Stopped ---")

    def start_monitoring(self):
        '''Stops any existing monitors and starts new ones based on current settings.'''
        print("\n--- Starting Monitoring ---")
        
        self.show_progress_bar()
        self._clear_monitors() # Stop any running monitors first

        count = 0
        scan_threads = [] # A list to hold our background scanning threads
        
        monitored_sources_data = {
            "Main TXT File": self.txt_folder_path,
            "TXT Source 2": self.txt_folder_path_set2,
            "TXT Source 3": self.txt_folder_path_set3,
            "TXT Source 4": self.txt_folder_path_set4,
            "TXT Source 5": self.txt_folder_path_set5
        }
        
        for source_name, source_path in monitored_sources_data.items():
            if source_path and os.path.isdir(source_path) and source_name not in self.folder_paths:
                self.folder_paths[source_name] = source_path
                self.folder_columns[source_name] = self.folder_columns.get(source_name, source_name.replace(" ", "_") + "_File")
                self.file_extensions[source_name] = self.file_extensions.get(source_name, "txt")
                self.folder_skips[source_name] = self.folder_skips.get(source_name, False)

        for folder_name, folder_path in self.folder_paths.items():
            if folder_path and os.path.isdir(folder_path) and not self.folder_skips.get(folder_name, False):
                file_extension = self.file_extensions.get(folder_name, "")
                event_handler = self.start_folder_monitoring(folder_name, folder_path, file_extension)
                
                if event_handler:
                    print(f"Queueing background scan for {folder_name}...")
                    scan_thread = threading.Thread(
                        target=event_handler.update_latest_file,
                        daemon=True
                    )
                    scan_thread.start()
                    scan_threads.append(scan_thread) # Add thread to our list
                    count += 1
            
            elif self.folder_skips.get(folder_name):
                print(f"Monitor for {folder_name} skipped by setting.")
            elif folder_path:
                print(f"Monitor for {folder_name} not started: path invalid or not a directory ({folder_path}).")

        self.is_monitoring = count > 0

        # If we started any scans, launch the supervisor to watch them
        if scan_threads:
            supervisor_thread = threading.Thread(
                target=self._manage_scan_threads,
                args=(scan_threads,),
                daemon=True
            )
            supervisor_thread.start()
        else:
            # If no scans were started, just hide the progress bar immediately
            self.hide_progress_bar()
        
        if self.is_monitoring:
            self.update_status(f"Scanning {count} folders in the background...")
        else:
            self.update_status("Monitoring not started. Check folder paths are valid.")
            
        self.update_monitor_indicator_text()
        
    # --- Programmed Events Scheduling ---
    def start_auto_sync(self):
        """Schedules the first periodic sync and stores the timer ID."""
        self.stop_auto_sync() # Ensure no other timers are running
        if self.auto_sync_enabled_var.get() and self.sqlite_mirror_enabled_var.get():
            interval_minutes = self.auto_sync_interval_min_var.get()
            if interval_minutes > 0:
                delay_ms = interval_minutes * 60 * 1000
                self._auto_sync_timer_id = self.master.after(delay_ms, self._periodic_sync_worker)
                print(f"Automatic DB sync scheduled to run every {interval_minutes} minutes.")
                self.update_status(f"Auto-sync scheduled every {interval_minutes} mins.")

    def stop_auto_sync(self, log_message=True):
        """Stops the scheduled periodic sync."""
        if self._auto_sync_timer_id:
            self.master.after_cancel(self._auto_sync_timer_id)
            self._auto_sync_timer_id = None
            if log_message:
                print("Automatic DB sync stopped.")

    def _periodic_sync_worker(self):
        """The worker function that runs on the timer, offloading the heavy work."""
        if not self.sqlite_mirror_enabled_var.get() or not self.sqlite_manager:
            self.stop_auto_sync()
            return

        def _sync_in_thread():
            """This function contains the slow code that runs on a separate thread."""
            self.master.after(0, self.update_status, "Starting periodic background sync...")
            print("\n--- Running Periodic Background Sync ---")
            try:
                header_finder = lambda path: self._find_header_row(path)
                success = self.sqlite_manager.full_sync(self.log_file_path, header_finder)
                
                if success:
                    self.master.after(0, self.update_status, "Periodic background sync complete.")
                    print("--- Periodic Sync Complete ---\n")
                else:
                    self.master.after(0, self.update_status, "Periodic sync had errors.")
                    print("--- Periodic Sync Failed ---\n")
            except Exception as e:
                error_msg = f"Auto-sync failed: {e}"
                self.master.after(0, self.update_status, error_msg)
                print(error_msg)
                traceback.print_exc()
            finally:
                # IMPORTANT: Reschedule the next sync *after* the current one finishes
                # Don't log cancellation message since we're just rescheduling
                self.stop_auto_sync(log_message=False)
                if self.auto_sync_enabled_var.get() and self.sqlite_mirror_enabled_var.get():
                    interval_minutes = self.auto_sync_interval_min_var.get()
                    if interval_minutes > 0:
                        delay_ms = interval_minutes * 60 * 1000
                        self._auto_sync_timer_id = self.master.after(delay_ms, self._periodic_sync_worker)
                        print(f"Next automatic sync scheduled in {interval_minutes} minutes.")


        # Run the actual sync in a background thread to not freeze the GUI
        sync_thread = threading.Thread(target=_sync_in_thread, daemon=True)
        sync_thread.start()

    def schedule_new_day(self):
        '''Schedule the next "New Day" log to trigger at midnight.'''

        now = datetime.datetime.now()
        tomorrow = now.date() + datetime.timedelta(days=1)
        midnight = datetime.datetime.combine(tomorrow, datetime.time.min)
        time_until_midnight_ms = int((midnight - now).total_seconds() * 1000)
        trigger_delay_ms = time_until_midnight_ms + 1000

        self._new_day_timer_id = self.master.after(trigger_delay_ms, self.trigger_new_day) # Set the timer to trigger at midnight - .after(delay in ms, callback function)
        print(f"Next 'New Day' event scheduled for {midnight} (in {time_until_midnight_ms/1000:.1f} seconds).") #DEBUG

    def trigger_new_day(self):
        """Triggers the Midnight position event and reschedules the next one."""

        if self.new_day_event_enabled_var.get():
            self._perform_log_action(event_type="New Day",
                            event_text_for_excel="Midnight Position",
                            triggering_button=None,
                            txt_source_key="Main TXT")
        else:
            print("'New Day' event is disabled, skipping log.")

        # After logging the new day, reschedule the next trigger
        self.schedule_new_day()

    def schedule_hourly_log(self):
        """Schedules the next hourly KP log to trigger on the hour."""
        # ... (This function remains unchanged as in your original code) ...
        now = datetime.datetime.now()
        next_hour = (now + datetime.timedelta(hours=1)).replace(minute=0, second=0, microsecond=0)
        time_until_next_hour_ms = int((next_hour - now).total_seconds() * 1000)

        # Add a small buffer (e.g., 1 second) to ensure it triggers after the hour
        trigger_delay_ms = time_until_next_hour_ms + 1000

        self._hourly_log_timer_id = self.master.after(trigger_delay_ms, self.trigger_hourly_log)
        print(f"Next 'Hourly KP Log' scheduled for {next_hour} (in {time_until_next_hour_ms/1000:.1f} seconds).")

    # --- Wrapper for the automatic timer event ---
    def trigger_hourly_log(self):
        """Triggers the core hourly log, then reschedules the next one."""
        self.trigger_hourly_log_core()
        # Reschedule for the following hour
        self.schedule_hourly_log()

    # --- CORE LOGIC
    def trigger_hourly_log_core(self):
        """
        Calculates and performs the hourly log without modifying the timer schedule.
        Includes user feedback if the required KP or column configuration is missing.
        """

        if not self.hourly_event_enabled_var.get():
            print("'Hourly KP Log' event is disabled, skipping log.")
            return

        # Get column names from settings
        kp_col_name = self.txt_field_columns.get("KP")
        event_col_name = self.txt_field_columns.get("Event")
        line_field_name = "Line name"
        line_col_name = self.txt_field_columns.get(line_field_name)

        if not kp_col_name or not event_col_name or not line_col_name:
            error_msg = f"Error: 'KP', 'Event', or '{line_field_name}' column not configured in TXT Data Columns settings. Skipping log."
            print(error_msg)
            self.update_status(error_msg)
            return
            
        # 1. Get current KP and Line Name value
        current_kp = None
        current_line = None
        try:
            txt_source_key = self.hourly_log_txt_source_key.get()
            source_folder_path = self._get_path_from_source_key(txt_source_key)
            
            txt_data = self._get_txt_data_from_source(source_folder_path, txt_source_key)
            
            current_kp_str = txt_data.get(kp_col_name)
            current_line = txt_data.get(line_col_name)

            if current_kp_str and str(current_kp_str).strip():
                current_kp = float(current_kp_str)
            else:
                # Raise an error if KP data is missing/invalid
                raise ValueError(f"KP data is empty or non-numeric in source file ({current_kp_str})")
                
        except (ValueError, TypeError, AttributeError, Exception) as e:
            # Catch file read errors, parse errors, or the ValueError raised above
            error_msg = f"Automatic KP Log skipped: Could not retrieve valid KP/Line from source. ({e})"
            print(error_msg)
            self.update_status(error_msg)
            return # Skip the log entry and return

        if current_kp is None or current_line is None:
            # This check is mostly redundant due to the try/except block above but remains as a safeguard
            error_msg = "Automatic KP Log skipped: Could not retrieve a valid current KP or Line Name."
            print(error_msg)
            self.update_status(error_msg)
            return

        # 2. Find the last hourly KP log from the Excel file
        last_kp = None
        last_line = None
        
        try:
            df = pd.read_excel(self.log_file_path)
            
            # Filter for previous hourly logs
            hourly_logs_df = df[df[event_col_name].str.startswith("Current KP:", na=False)].copy()
            
            # Ensure the KP and Line name columns are usable
            hourly_logs_df[kp_col_name] = pd.to_numeric(hourly_logs_df[kp_col_name], errors='coerce')
            hourly_logs_df.dropna(subset=[kp_col_name], inplace=True)
            
            # Ensure the Line name column is present and usable
            if line_col_name not in hourly_logs_df.columns:
                print(f"Error: Line Name column '{line_col_name}' not found in log file.")
            
            if not hourly_logs_df.empty:
                # --- Get the last logged KP and Line Name ---
                last_log = hourly_logs_df.iloc[-1]
                last_kp = last_log[kp_col_name]
                # Safely get the line, use a placeholder if the column is somehow missing from the log data
                last_line = last_log.get(line_col_name, "N/A_LINE_ERROR")
                                
        except Exception as e:
            print(f"Could not read or find last KP/Line from Excel file: {e}")

        # 3. Format the event text string
        event_text = ""
        
        if last_kp is not None and last_line is not None:
            
            if current_line == last_line:
                # SCENARIO 1: Line Name is the same (Simple Calculation)
                progress = current_kp - last_kp
                event_text = (
                    f"Current KP: {current_kp:.3f} | "
                    f"Progress last hour: {progress:+.3f} km | "
                    f"Line: {current_line}"
                )
            else:
                # SCENARIO 2: Line Name has changed (Reset Calculation)
                progress = current_kp
                
                event_text = (
                    f"Current KP: {current_kp:.3f} | "
                    f"**LINE CHANGED** from {last_line} to {current_line}. "
                    f"Progress on new line: {progress:.3f} km"
                )
                
        else:
            # SCENARIO 3: First hourly log
            event_text = f"Current KP: {current_kp:.3f} | First hourly log on Line: {current_line}"

        # 4. Call the logging function with the generated text
        self._perform_log_action(event_type="Hourly KP Log",
                                 event_text_for_excel=event_text,
                                 triggering_button=None, # No button is associated with the core logic
                                 txt_source_key=self.hourly_log_txt_source_key.get())


    # --- Custom Button Management ---
    def _show_custom_button_context_menu(self, event, button_index):
        """Shows a context menu for the clicked custom button."""
        # Check if the right-click was on one of the custom button tab frames
        # Iterate through custom_button_tab_frames values
        for tab_frame_widget in self.custom_button_tab_frames.values():
            # Check if event.widget is the tab_frame_widget itself, or a child of it (not necessarily the button)
            if str(event.widget) == str(tab_frame_widget) or tab_frame_widget.winfo_containing(event.x_root, event.y_root) == tab_frame_widget:
                # If right-click is on the tab frame itself or its background, show add button menu
                self._show_add_button_context_menu(event)
                return

        # If an inline editor is already open, focus it instead of opening another or a context menu
        if self.custom_inline_editor_window and self.custom_inline_editor_window.winfo_exists():
            self.custom_inline_editor_window.lift()
            self.custom_inline_editor_window.focus_set()
            return

        context_menu = tk.Menu(self.master, tearoff=0)
        current_button_text = self.custom_button_configs[button_index].get("text", f"Custom {button_index+1}")
        # Right Click edit button command
        context_menu.add_command(label=f"Edit \"{current_button_text}\" Settings...",
                              command=lambda: self._edit_custom_button_inline(button_index))
    # Add a separator for visual clarity
        context_menu.add_separator()
    # Add the new "Delete" command
        context_menu.add_command(label=f"Delete \"{current_button_text}\"",
                              command=lambda: self._delete_custom_button(button_index))
        try:
            context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            context_menu.grab_release()

    def _show_main_button_context_menu(self, event, button_name):
        """Shows a context menu for the clicked main button."""
        # If an inline editor is already open, focus it instead of opening a new one
        if self.custom_inline_editor_window and self.custom_inline_editor_window.winfo_exists():
            self.custom_inline_editor_window.lift()
            self.custom_inline_editor_window.focus_set()
            return

        context_menu = tk.Menu(self.master, tearoff=0)
        
        # Add the command to edit the button's settings
        context_menu.add_command(label=f"Edit \"{button_name}\" Settings...",
                                 command=lambda: self._edit_main_button_inline(button_name))
        
        try:
            context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            context_menu.grab_release()

    def _show_add_button_context_menu(self, event):
        """Shows a context menu specifically for adding a new button."""
        if self.num_custom_buttons >= self.MAX_CUSTOM_BUTTONS:
            messagebox.showinfo("Limit Reached", f"You have reached the maximum number of {self.MAX_CUSTOM_BUTTONS} custom buttons.", parent=self.master)
            return

        context_menu = tk.Menu(self.master, tearoff=0)
        context_menu.add_command(label="Add New Custom Button",
                                 command=self._add_new_custom_button)
        
        try:
            context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            context_menu.grab_release()

    def _edit_main_button_inline(self, button_name):
        """
        Opens a small Toplevel window to edit settings for a specific main button.
        """
        if self.custom_inline_editor_window and self.custom_inline_editor_window.winfo_exists():
            self.custom_inline_editor_window.lift()
            self.custom_inline_editor_window.focus_set()
            return

        # Fetch the complete configuration for the button
        button_config = self.main_button_configs.get(button_name, {})
        
        editor_window = tk.Toplevel(self.master)
        self.custom_inline_editor_window = editor_window
        editor_window.title(f"Edit \"{button_name}\"")
        editor_window.transient(self.master)
        editor_window.grab_set()
        editor_window.resizable(False, False)

        frame = ttk.Frame(editor_window, padding="15")
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)

        # --- Get current values ---
        current_event_text = button_config.get("event_text", "")
        current_event_code = button_config.get("event_code", "")
        
        #  Determine the source key based on the button name
        if button_name == "Manual KP Log":
            current_source_key = self.hourly_log_txt_source_key.get()
        else:
            current_source_key = button_config.get("txt_source_key", "Main TXT")
        
        # NOTE: Colors for Manual/Hourly KP Log are now retrieved from the Programmed Events setting's storage
        if button_name in ["Manual KP Log", "Hourly KP Log"]:
            # Retrieve colors using the fixed 'Hourly KP Log' key
            current_bg_color, current_font_color = self.button_colors.get("Hourly KP Log", (None, None))
        else:
            # For all other buttons, use their own key
            current_bg_color, current_font_color = self.button_colors.get(button_name, (None, None))
        
        # --- Create StringVars ---
        event_text_var = tk.StringVar(value=current_event_text)
       
        # Find the full "Code - Description" string for the current code to display it initially
        initial_display_value = ""
        if current_event_code and current_event_code in self.event_codes:
            initial_display_value = f"{current_event_code} - {self.event_codes[current_event_code]}"
        elif current_event_code:
            initial_display_value = f"{current_event_code} - <no description>"
        event_code_display_var = tk.StringVar(value=initial_display_value)
        
        # Create a StringVar for the source name to display in the Combobox
        # Find the display name from the alias map
        current_display_name = self.txt_source_aliases.get(current_source_key, current_source_key)
        source_display_var = tk.StringVar(value=current_display_name)
        
        # Use separate vars for the dialog's color widgets to track changes before saving
        button_bg_color_var = tk.StringVar(value=current_bg_color if current_bg_color else "")
        button_font_color_var = tk.StringVar(value=current_font_color if current_font_color else "")
        
        # --- UI Elements for the editor ---
        row_idx = 0
        
        # Event Text Entry
        ttk.Label(frame, text="Event Text:").grid(row=row_idx, column=0, sticky="w", pady=5, padx=5)
        event_text_entry = ttk.Entry(frame, textvariable=event_text_var, width=40)
        
        # Make text read-only for auto-generated events
        if button_name in ["Manual KP Log", "Hourly KP Log"]:
            event_text_var.set("Auto generated")
            event_text_entry.config(state="readonly")
            ToolTip(event_text_entry, "This event text is automatically generated and cannot be manually edited.")

        event_text_entry.grid(row=row_idx, column=1, sticky="ew", pady=5, padx=5)
        if button_name not in ["Manual KP Log", "Hourly KP Log"]:
            ToolTip(event_text_entry, "Text written to the 'Event' column in the log.")

        row_idx += 1
        # Event Code Combobox
        ttk.Label(frame, text="Event Code:").grid(row=row_idx, column=0, sticky="w", pady=5, padx=5)
        
        
        # Create a list of "Code - Description" strings for the dropdown
        event_code_display_list = [""] # Start with a blank option
        for code, desc in sorted(self.event_codes.items()):
            event_code_display_list.append(f"{code} - {desc}")
        
        event_code_combobox = ttk.Combobox(frame, textvariable=event_code_display_var, # Use the new display variable
                                             values=event_code_display_list,           # Use the new display list
                                             state="readonly", width=37)
        

        event_code_combobox.grid(row=row_idx, column=1, sticky="ew", pady=5, padx=5)
        ToolTip(event_code_combobox, "Select an event code to write to the 'Code' column when this button is pressed.")
        
        row_idx += 1
        # Add Event Source Combobox for main buttons
        ttk.Label(frame, text="Event Source:").grid(row=row_idx, column=0, sticky="w", pady=5, padx=5)
        
        # Build list for the dropdown using aliases for display
        aliases = self.txt_source_aliases
        internal_keys = TXT_FILES_KEYS
        display_names = [aliases.get(key, key) for key in internal_keys] # Get names from aliases or use defaults
        
        source_combobox = ttk.Combobox(frame, textvariable=source_display_var,
                                             values=display_names, state="readonly", width=37)
        source_combobox.grid(row=row_idx, column=1, sticky="ew", pady=5, padx=5)
        
        # Set the state and tooltip for the manual log button source
        if button_name in ["Manual KP Log", "Hourly KP Log"]:
             source_combobox.config(state="readonly") 
             ToolTip(source_combobox, "Source is linked to the 'KP Data Source' setting in the Programmed Events tab.")
        else:
             ToolTip(source_combobox, "Select which data source this button should use. Names are configured in Settings -> File Paths.")
        
        # --- CONDITIONAL COLOR PICKERS (NEW BLOCK) ---
        # The color is derived from the "Programmed Events" tab for these, so editing here is redundant.
        if button_name not in ["Manual KP Log", "Hourly KP Log"]:
            row_idx += 1
            # Button Background Color Picker
            ttk.Label(frame, text="Button Background:").grid(row=row_idx, column=0, sticky="w", pady=5, padx=5)
            
            bg_color_widget_frame = ttk.Frame(frame)
            bg_color_widget_frame.grid(row=row_idx, column=1, sticky="w", pady=5, padx=5)

            bg_color_display_label = tk.Label(bg_color_widget_frame, width=4, relief="solid", borderwidth=1,
                                                 background=button_bg_color_var.get() if button_bg_color_var.get() else 'SystemButtonFace')
            bg_color_display_label.pack(side="left", padx=(0, 5))

            clear_bg_btn = ttk.Button(bg_color_widget_frame, text="X", width=2,
                                         command=lambda: self._set_color_on_widget(button_bg_color_var, bg_color_display_label, None, editor_window))
            clear_bg_btn.pack(side="left", padx=1)
            ToolTip(clear_bg_btn, "Clear button background color.")

            choose_bg_btn = ttk.Button(bg_color_widget_frame, text="...", width=3,
                                        command=lambda v=button_bg_color_var, l=bg_color_display_label: self._choose_color_dialog(v, l, editor_window, button_name + " Background"))
            choose_bg_btn.pack(side="left", padx=1)
            ToolTip(choose_bg_btn, "Choose a custom background color.")

            row_idx += 1
            # Button Font Color Picker
            ttk.Label(frame, text="Button Font Color:").grid(row=row_idx, column=0, sticky="w", pady=5, padx=5)
            
            font_color_widget_frame = ttk.Frame(frame)
            font_color_widget_frame.grid(row=row_idx, column=1, sticky="w", pady=5, padx=5)

            font_color_display_label = tk.Label(font_color_widget_frame, width=4, relief="solid", borderwidth=1,
                                                 background=button_font_color_var.get() if button_font_color_var.get() else 'SystemButtonFace')
            font_color_display_label.pack(side="left", padx=(0, 5))

            clear_font_btn = ttk.Button(font_color_widget_frame, text="X", width=2,
                                           command=lambda: self._set_color_on_widget(button_font_color_var, font_color_display_label, None, editor_window))
            clear_font_btn.pack(side="left", padx=1)
            ToolTip(clear_font_btn, "Clear button font color.")

            choose_font_btn = ttk.Button(font_color_widget_frame, text="...", width=3,
                                           command=lambda v=button_font_color_var, l=font_color_display_label: self._choose_color_dialog(v, l, editor_window, button_name + " Font"))
            choose_font_btn.pack(side="left", padx=1)
            ToolTip(choose_font_btn, "Choose a custom font color.")
     
        # --- Save and Cancel buttons ---
        row_idx += 1
        button_controls_frame = ttk.Frame(frame)
        button_controls_frame.grid(row=row_idx, column=0, columnspan=2, pady=(15,0), sticky="e")

        def save_main_button_changes():
            # Save the new event text
            self.main_button_configs[button_name]['event_text'] = event_text_var.get()
                        
            # Get the full "Code - Description" string and parse it to save only the code
            selected_display_string = event_code_display_var.get()
            code_to_save = ""
            if " - " in selected_display_string:
                code_to_save = selected_display_string.split(" - ", 1)[0]
            self.main_button_configs[button_name]['event_code'] = code_to_save
                   
            #Save the selected source key
            # We need to map the display name back to the internal key
            # Recreate maps using the global constant
            internal_keys_for_map = TXT_FILES_KEYS
            display_names_for_map = [self.txt_source_aliases.get(key, key) for key in internal_keys_for_map]
            internal_to_display_map = {internal: display for display, internal in zip(display_names_for_map, internal_keys_for_map)}
            
            selected_display_name = source_combobox.get() 
            # Reverse lookup (Display Name -> Internal Key)
            selected_source_key = next((key for key, display in internal_to_display_map.items() if display == selected_display_name), "None")
            
            # Check button name and save to the correct location
            if button_name == "Manual KP Log":
                # Manual KP Log source is linked to the hourly event source setting
                self.hourly_log_txt_source_key.set(selected_source_key)
            else:
                self.main_button_configs[button_name]['txt_source_key'] = selected_source_key
            
            # Save the new colors as a tuple (Only if the pickers were shown)
            if button_name not in ["Manual KP Log", "Hourly KP Log"]:
                new_bg_color_hex = button_bg_color_var.get() if button_bg_color_var.get() else None
                new_font_color_hex = button_font_color_var.get() if button_font_color_var.get() else None
                self.button_colors[button_name] = (new_bg_color_hex, new_font_color_hex)
            
            # Persist all settings and redraw the UI
            self.save_settings()
            
            # Call the comprehensive update function (which re-creates buttons with new styles)
            self.update_custom_buttons() # This method name is a bit misleading, it updates all buttons

            editor_window.destroy()

        ttk.Button(button_controls_frame, text="Save", command=save_main_button_changes, style="Accent.TButton").pack(side="right", padx=5)
        ttk.Button(button_controls_frame, text="Cancel", command=editor_window.destroy).pack(side="right")

        editor_window.protocol("WM_DELETE_WINDOW", editor_window.destroy)
        editor_window.wait_window(editor_window)
        self.custom_inline_editor_window = None
        
    def _add_new_custom_button(self):
        """Adds a new custom button configuration and updates the GUI."""
        if self.num_custom_buttons < self.MAX_CUSTOM_BUTTONS:
            self.num_custom_buttons += 1
            new_button_idx = self.num_custom_buttons
            new_config = {
                "text": f"Custom {new_button_idx}",
                "event_text": f"Custom {new_button_idx} Event",
                "txt_source_key": "None",
                "tab_group": "Main" # **:** Default to "Main" tab
            }
            self.custom_button_configs.append(new_config)
            
            # Ensure the new button gets a default color entry if it doesn't exist
            if new_config["text"] not in self.button_colors:
                self.button_colors[new_config["text"]] = (None, None)

            self.save_settings()
            self.update_custom_buttons()
            self.update_status(f"Added new button: '{new_config['text']}'.")
            # Optionally, open the inline editor for the newly added button
            self._edit_custom_button_inline(self.num_custom_buttons - 1)
        else:
            messagebox.showinfo("Limit Reached", f"You have reached the maximum number of {self.MAX_CUSTOM_BUTTONS} custom buttons.", parent=self.master)
  
    def _delete_custom_button(self, button_index):
        """Deletes a custom button after confirmation."""
        
        # Safely get the button text for the confirmation message
        try:
            button_text = self.custom_button_configs[button_index].get("text", f"Custom {button_index + 1}")
        except IndexError:
            messagebox.showerror("Error", "Cannot delete button. Index is out of range.", parent=self.master)
            return

        # Ask for user confirmation before deleting
        if not messagebox.askyesno(
            "Confirm Deletion",
            f"Are you sure you want to permanently delete the button '{button_text}'?",
            parent=self.master):
            self.update_status(f"Deletion of '{button_text}' cancelled.")
            return

        # --- Deletion Logic ---
        # 1. Remove the button's configuration from the list
        if button_index < len(self.custom_button_configs):
            # Also remove any associated color from the button_colors dictionary
            if button_text in self.button_colors:
                del self.button_colors[button_text]
            
            del self.custom_button_configs[button_index]
            
            # 2. Decrement the total number of custom buttons
            self.num_custom_buttons -= 1

            # 3. Save the updated settings to the JSON file
            self.save_settings()

            # 4. Refresh the buttons on the main UI
            self.update_custom_buttons()
            
            self.update_status(f"Button '{button_text}' was deleted.")
        else:
            self.update_status("Error: Could not delete button (invalid index).")
    
    def _show_tab_context_menu(self, event):
        """Shows a context menu for adding, renaming, or deleting notebook tabs."""
        context_menu = tk.Menu(self.master, tearoff=0)
        
        # Add the "Add New Tab" command, which is always available
        context_menu.add_command(label="Add New Tab...", command=self._add_new_tab_dialog)
        
        try:
            # Check if the click was on an existing tab label
            tab_index = self.custom_buttons_notebook.index(f"@{event.x},{event.y}")
            tab_name = self.custom_buttons_notebook.tab(tab_index, "text")
            
            # If so, add commands for renaming and deleting that specific tab
            context_menu.add_separator()
            context_menu.add_command(
                label=f"Rename '{tab_name}' Tab...",
                command=lambda: self._rename_tab_dialog(tab_name)
            )
            context_menu.add_command(
                label=f"Delete '{tab_name}' Tab",
                command=lambda: self._delete_tab(tab_name)
            )
            
            # Protect the "Main" tab from being renamed or deleted
            if tab_name == "Main":
                context_menu.entryconfigure(f"Rename '{tab_name}' Tab...", state=tk.DISABLED)
                context_menu.entryconfigure(f"Delete '{tab_name}' Tab", state=tk.DISABLED)

        except tk.TclError:
            # This error means the click was not on a tab label, so we just show the "Add" menu.
            pass

        try:
            context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            context_menu.grab_release()

    def _rename_tab_dialog(self, old_name):
        """Opens a dialog to get the new name for a tab."""

        new_name = simpledialog.askstring(
            "Rename Tab",
            f"Enter the new name for the '{old_name}' tab:",
            parent=self.master,
            initialvalue=old_name
        )

        if new_name and new_name.strip() and new_name != old_name:
            self._rename_tab_group(old_name, new_name.strip())
        elif new_name and new_name == old_name:
            self.update_status("Tab rename cancelled (name is the same).")
        else:
            self.update_status("Tab rename cancelled.")

    def _rename_tab_group(self, old_name, new_name):
        """Renames a tab group and updates all related configurations."""
        if new_name in self.custom_button_tab_groups:
            messagebox.showerror("Rename Error", f"The tab name '{new_name}' already exists.", parent=self.master)
            return

        # Update the master list of tab groups
        try:
            # Find and replace the old name with the new name
            index = self.custom_button_tab_groups.index(old_name)
            self.custom_button_tab_groups[index] = new_name
        except ValueError:
            # If not found (shouldn't happen with this workflow), just add the new one
            self.custom_button_tab_groups.append(new_name)

        # Update all custom button configurations that use the old tab name
        for config in self.custom_button_configs:
            if config.get("tab_group") == old_name:
                config["tab_group"] = new_name
        
        self.update_status(f"Renamed tab '{old_name}' to '{new_name}'.")

        # Save the settings to persist the change
        self.save_settings()

        # Re-render the main buttons to show the change immediately
        self.update_custom_buttons() 

    def _add_new_tab_dialog(self):
        """Opens a dialog to get the name for a new tab."""
        from tkinter import simpledialog
        new_name = simpledialog.askstring(
            "Add New Tab",
            "Enter the name for the new tab:",
            parent=self.master
        )

        if not new_name or not new_name.strip():
            self.update_status("Add tab cancelled.")
            return

        new_name = new_name.strip()
        
        # Check for duplicates
        existing_groups = [group.lower() for group in self.custom_button_tab_groups]
        if new_name.lower() in existing_groups:
            messagebox.showerror("Creation Error", f"The tab name '{new_name}' already exists.", parent=self.master)
            return

        # Add the new tab, save, and refresh
        self.custom_button_tab_groups.append(new_name)
        self.update_status(f"Added new tab: '{new_name}'.")
        self.save_settings()
        self.update_custom_buttons()

    def _delete_tab(self, tab_name):
        """Deletes a tab and moves its buttons to the 'Main' tab."""
        if tab_name == "Main":
            messagebox.showerror("Delete Error", "The 'Main' tab cannot be deleted.", parent=self.master)
            return

        # Confirm deletion with the user
        if not messagebox.askyesno(
            "Confirm Deletion",
            f"Are you sure you want to delete the '{tab_name}' tab?\n\n"
            f"All buttons in this tab will be moved to the 'Main' tab.",
            parent=self.master):
            self.update_status("Delete tab cancelled.")
            return

        # Move all buttons from the deleted tab to the 'Main' tab
        for config in self.custom_button_configs:
            if config.get("tab_group") == tab_name:
                config["tab_group"] = "Main"
        
        # Remove the tab from the master list
        if tab_name in self.custom_button_tab_groups:
            self.custom_button_tab_groups.remove(tab_name)

        self.update_status(f"Deleted tab '{tab_name}'.")
        self.save_settings()
        self.update_custom_buttons()           

    def _edit_custom_button_inline(self, button_index):
        """
        Opens a small Toplevel window to edit settings for a specific custom button.
        """
        if self.custom_inline_editor_window and self.custom_inline_editor_window.winfo_exists():
            self.custom_inline_editor_window.lift()
            self.custom_inline_editor_window.focus_set()
            return

        button_config = self.custom_button_configs[button_index]
        
        editor_window = tk.Toplevel(self.master)
        self.custom_inline_editor_window = editor_window
        editor_window.title(f"Edit Custom Button {button_index + 1}")
        editor_window.transient(self.master)
        editor_window.grab_set()
        editor_window.resizable(False, False)
        
        self.master.update_idletasks()
        main_x = self.master.winfo_x()
        main_y = self.master.winfo_y()
        main_width = self.master.winfo_width()
        main_height = self.master.winfo_height()

        editor_window.update_idletasks()
        dialog_width = editor_window.winfo_reqwidth() or 350
        dialog_height = editor_window.winfo_reqheight() or 300 

        center_x = main_x + (main_width // 2) - (dialog_width // 2)
        center_y = main_y + (main_height // 2) - (dialog_height // 2)
        editor_window.geometry(f"+{center_x}+{center_y}")

        frame = ttk.Frame(editor_window, padding="15")
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1)

        current_bg_color, current_font_color = self.button_colors.get(button_config.get("text"), (None, None))

        button_text_var = tk.StringVar(value=button_config.get("text", f"Custom {button_index+1}"))
        event_text_var = tk.StringVar(value=button_config.get("event_text", f"{button_config.get('text', f'Custom {button_index+1}')} Triggered"))
        tab_group_var = tk.StringVar(value=button_config.get("tab_group", "Main"))
        
        current_event_code = button_config.get("event_code", "")
        initial_display_value = ""
        if current_event_code and current_event_code in self.event_codes:
            initial_display_value = f"{current_event_code} - {self.event_codes[current_event_code]}"
        elif current_event_code:
            initial_display_value = f"{current_event_code} - <no description>"
        event_code_display_var = tk.StringVar(value=initial_display_value)
        
        button_bg_color_var = tk.StringVar(value=current_bg_color if current_bg_color else "")
        button_font_color_var = tk.StringVar(value=current_font_color if current_font_color else "")
        
        row_idx = 0

        # Button Text
        ttk.Label(frame, text="Button Text:").grid(row=row_idx, column=0, sticky="w", pady=2, padx=5)
        text_entry = ttk.Entry(frame, textvariable=button_text_var, width=30)
        text_entry.grid(row=row_idx, column=1, columnspan=2, sticky="ew", pady=2, padx=5)
        ToolTip(text_entry, "Text displayed on the button.")

        row_idx += 1
        # Event Text
        ttk.Label(frame, text="Event Text:").grid(row=row_idx, column=0, sticky="w", pady=2, padx=5)
        event_entry = ttk.Entry(frame, textvariable=event_text_var, width=30)
        event_entry.grid(row=row_idx, column=1, columnspan=2, sticky="ew", pady=2, padx=5)
        ToolTip(event_entry, "Text written to the 'Event' column in the log.")

        row_idx += 1
        # Event Code Combobox
        ttk.Label(frame, text="Event Code:").grid(row=row_idx, column=0, sticky="w", pady=2, padx=5)
        event_code_display_list = [""]
        for code, desc in sorted(self.event_codes.items()):
            event_code_display_list.append(f"{code} - {desc}")
        event_code_combobox = ttk.Combobox(frame, textvariable=event_code_display_var,
                                           values=event_code_display_list, state="readonly", width=27)
        event_code_combobox.grid(row=row_idx, column=1, columnspan=2, sticky="ew", pady=2, padx=5)
        ToolTip(event_code_combobox, "Select an event code to write to the 'Code' column when this button is pressed.")

        row_idx += 1
        # Event Source Combobox
        ttk.Label(frame, text="Event Source:").grid(row=row_idx, column=0, sticky="w", pady=2, padx=5)
        aliases = self.txt_source_aliases
        internal_keys = TXT_FILES_KEYS
        display_names = ["None"] + [aliases.get(key, key) for key in internal_keys[1:]]
        display_to_internal_map = {display: internal for display, internal in zip(display_names, internal_keys)}
        internal_to_display_map = {internal: display for display, internal in zip(display_names, internal_keys)}
        current_internal_key = button_config.get("txt_source_key", "None")
        txt_source_display_var = tk.StringVar(value=internal_to_display_map.get(current_internal_key, "None"))
        source_combobox = ttk.Combobox(frame, textvariable=txt_source_display_var,
                                       values=display_names, state="readonly", width=27)
        source_combobox.grid(row=row_idx, column=1, columnspan=2, sticky="ew", pady=2, padx=5)
        ToolTip(source_combobox, "Select which data source this button should use.")

        row_idx += 1
        # Tab Group selection
        ttk.Label(frame, text="Tab Group:").grid(row=row_idx, column=0, sticky="w", pady=2, padx=5)
        all_tab_groups = sorted(self.custom_button_tab_groups[:])
        tab_group_combobox = ttk.Combobox(frame, textvariable=tab_group_var,
                                          values=all_tab_groups, width=27)
        tab_group_combobox.grid(row=row_idx, column=1, columnspan=2, sticky="ew", pady=2, padx=5)
        ToolTip(tab_group_combobox, "Assign this button to a tab group.")

        row_idx += 1
        # Color Pickers and Save button... (This part is unchanged)
        # Button Background Color Picker
        ttk.Label(frame, text="Button Background:").grid(row=row_idx, column=0, sticky="w", pady=2, padx=5)
        bg_color_widget_frame = ttk.Frame(frame)
        bg_color_widget_frame.grid(row=row_idx, column=1, sticky="w", pady=2, padx=5)
        bg_color_display_label = tk.Label(bg_color_widget_frame, width=4, relief="solid", borderwidth=1,
                                            background=button_bg_color_var.get() if button_bg_color_var.get() else 'SystemButtonFace')
        bg_color_display_label.pack(side="left", padx=(0, 5))
        clear_bg_btn = ttk.Button(bg_color_widget_frame, text="X", width=2, style="Toolbutton",
                                  command=lambda: self._set_color_on_widget(button_bg_color_var, bg_color_display_label, None, editor_window))
        clear_bg_btn.pack(side="left", padx=1)
        ToolTip(clear_bg_btn, "Clear button background color (use default appearance).")
        pastel_colors_for_picker = ["#FFB3BA", "#FFDFBA", "#FFFFBA", "#BAFFC9", "#BAE1FF"]
        for p_color in pastel_colors_for_picker:
            try:
                b = tk.Button(bg_color_widget_frame, bg=p_color, width=1, height=1, relief="raised", bd=1,
                                  command=lambda c=p_color: self._set_color_on_widget(button_bg_color_var, bg_color_display_label, c, editor_window))
                b.pack(side=tk.LEFT, padx=1)
            except tk.TclError: pass
        choose_bg_btn = ttk.Button(bg_color_widget_frame, text="...", width=3, style="Toolbutton",
                                   command=lambda v=button_bg_color_var, l=bg_color_display_label, n=button_text_var.get(): self._choose_color_dialog(v, l, editor_window, n + " Background"))
        choose_bg_btn.pack(side="left", padx=1)
        ToolTip(choose_bg_btn, "Choose a custom background color.")
        row_idx += 1
        # Button Font Color Picker
        ttk.Label(frame, text="Button Font Color:").grid(row=row_idx, column=0, sticky="w", pady=2, padx=5)
        font_color_widget_frame = ttk.Frame(frame)
        font_color_widget_frame.grid(row=row_idx, column=1, sticky="w", pady=2, padx=5)
        font_color_display_label = tk.Label(font_color_widget_frame, width=4, relief="solid", borderwidth=1,
                                              background=button_font_color_var.get() if button_font_color_var.get() else 'SystemButtonFace')
        font_color_display_label.pack(side="left", padx=(0, 5))
        clear_font_btn = ttk.Button(font_color_widget_frame, text="X", width=2, style="Toolbutton",
                                      command=lambda: self._set_color_on_widget(button_font_color_var, font_color_display_label, None, editor_window))
        clear_font_btn.pack(side="left", padx=1)
        ToolTip(clear_font_btn, "Clear button font color (use default appearance).")
        default_font_colors_for_picker = ["#000000", "#FFFFFF"]
        for f_color in default_font_colors_for_picker:
            try:
                b = tk.Button(font_color_widget_frame, bg=f_color, width=1, height=1, relief="raised", bd=1,
                                  fg='white' if f_color == '#000000' else 'black',
                                  command=lambda c=f_color: self._set_color_on_widget(button_font_color_var, font_color_display_label, c, editor_window))
                b.pack(side=tk.LEFT, padx=1)
            except tk.TclError: pass
        choose_font_btn = ttk.Button(font_color_widget_frame, text="...", width=3, style="Toolbutton",
                                       command=lambda v=button_font_color_var, l=font_color_display_label, n=button_text_var.get(): self._choose_color_dialog(v, l, editor_window, n + " Font"))
        choose_font_btn.pack(side="left", padx=1)
        ToolTip(choose_font_btn, "Choose a custom font color.")
        row_idx += 1
        
        button_controls_frame = ttk.Frame(frame)
        button_controls_frame.grid(row=row_idx, column=0, columnspan=3, pady=(15,0), sticky="e")

        def save_changes():
            old_button_text = button_config.get("text")
            button_config["text"] = button_text_var.get().strip() or f"Custom {button_index+1}"
            button_config["event_text"] = event_text_var.get().strip() or f"{button_config['text']} Triggered"
            selected_display_string = event_code_display_var.get()
            code_to_save = ""
            if " - " in selected_display_string:
                code_to_save = selected_display_string.split(" - ", 1)[0]
            button_config["event_code"] = code_to_save
            button_config["tab_group"] = tab_group_var.get().strip() or "Main"
            selected_display_name = txt_source_display_var.get()
            button_config["txt_source_key"] = display_to_internal_map.get(selected_display_name, "None")
            # The line for saving "search_for_time" has been removed.
            new_bg_color_hex = button_bg_color_var.get() if button_bg_color_var.get() else None
            new_font_color_hex = button_font_color_var.get() if button_font_color_var.get() else None
            if old_button_text in self.button_colors and old_button_text != button_config["text"]:
                del self.button_colors[old_button_text]
            self.button_colors[button_config["text"]] = (new_bg_color_hex, new_font_color_hex)
            new_group = button_config["tab_group"]
            if new_group not in self.custom_button_tab_groups:
                self.custom_button_tab_groups.append(new_group)
                self.custom_button_tab_groups.sort()
            self.save_settings()
            self.update_custom_buttons()
            editor_window.destroy()

        ttk.Button(button_controls_frame, text="Save", command=save_changes, style="Accent.TButton").pack(side="right", padx=5)
        ttk.Button(button_controls_frame, text="Cancel", command=editor_window.destroy).pack(side="right")
        
        editor_window.protocol("WM_DELETE_WINDOW", editor_window.destroy)
        editor_window.wait_window(editor_window)
        self.custom_inline_editor_window = None

    def _set_color_on_widget(self, color_str_var, display_label, color_hex, parent_toplevel):
        """Internal helper to validate and set the color for a color picker display Label."""
        valid_color = None
        if color_hex:
            # Tkinter's Label widget supports direct background color setting
            try:
                # Test if the color is valid by trying to set it on a temporary widget
                temp_label = tk.Label(parent_toplevel, background=color_hex)
                valid_color = color_hex
                temp_label.destroy() # Clean up temp widget
            except tk.TclError:
                valid_color = None # Color was invalid
        
        color_str_var.set(valid_color if valid_color else "")
        
        try:
            # Update the actual display label
            display_label.config(background=valid_color if valid_color else 'SystemButtonFace')
        except tk.TclError:
            # If the widget is destroyed, just ignore
            pass

    def _choose_color_dialog(self, color_str_var, display_label, parent_toplevel, name="Item"):
        """Opens color chooser dialog and updates the color_str_var and display_label."""
        current_color = color_str_var.get()
        color_code = colorchooser.askcolor(color=current_color if current_color else None,
                                           title=f"Choose Color for {name}",
                                           parent=parent_toplevel)
        if color_code and color_code[1]:
            self._set_color_on_widget(color_str_var, display_label, color_code[1], parent_toplevel)

class TxtMappingDialog:
    """
    A separate window for editing the column mappings for a single TXT data source.
    """
    def __init__(self, master, parent_gui, source_key):
        self.master = master
        self.parent_gui = parent_gui
        self.source_key = source_key
        self.initial_config = self.parent_gui.all_txt_mappings.get(self.source_key, [])[:] # Deep copy the list for rollback
        
        display_name = self.parent_gui.txt_source_aliases.get(self.source_key, self.source_key)
        
        self.master.title(f"Field Mapping for: {display_name}")
        self.master.transient(self.parent_gui.settings_window_instance or self.parent_gui.master)
        self.master.grab_set()
        self.master.geometry("900x600")
        self.master.minsize(600, 400)
        
        # Main frame for padding and structure
        main_frame = ttk.Frame(self.master, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        # Description
        ttk.Label(main_frame, text=f"Configure the comma-separated order of fields for '{display_name}' ({source_key}). The order must match the data in the TXT/NPD/CSV file.", wraplength=750, justify=tk.LEFT).pack(pady=(0, 10), anchor='w')

        # Controls Frame
        controls_frame = ttk.Frame(main_frame)
        controls_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Button(controls_frame, text="Preview Latest Data", command=self.preview_latest_data).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(controls_frame, text="Clear Preview", command=self.clear_preview).pack(side=tk.LEFT, padx=(0, 20))

        use_default_btn = ttk.Button(controls_frame, text="Use Main Mapping", command=self.use_main_mapping)
        use_default_btn.pack(side=tk.LEFT, padx=(0, 10))
        ToolTip(use_default_btn, "Copies the field mapping settings from 'Main TXT' to this source.")
        
        # Only enable the button if the current source is NOT 'Main TXT'
        if source_key == "Main TXT":
            use_default_btn.config(state=tk.DISABLED)

        spacer = ttk.Frame(controls_frame)
        spacer.pack(side=tk.LEFT, expand=True, fill='x')

        self.txt_move_up_btn = ttk.Button(controls_frame, text="Move Up", command=lambda: self.move_selected_field("up"), state=tk.DISABLED)
        self.txt_move_up_btn.pack(side=tk.RIGHT, padx=5)

        self.txt_move_down_btn = ttk.Button(controls_frame, text="Move Down", command=lambda: self.move_selected_field("down"), state=tk.DISABLED)
        self.txt_move_down_btn.pack(side=tk.RIGHT, padx=5)

        ttk.Button(controls_frame, text="Add New Field", command=self.add_field_row).pack(side=tk.RIGHT, padx=5)
        
        # Scrollable Area for Fields
        self.txt_fields_canvas = tk.Canvas(main_frame, borderwidth=0, background="#ffffff")
        txt_scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.txt_fields_canvas.yview)
        self.txt_fields_scrollable_frame = ttk.Frame(self.txt_fields_canvas)
        
        self.txt_fields_scrollable_frame.bind("<Configure>", lambda e: self.txt_fields_canvas.configure(scrollregion=self.txt_fields_canvas.bbox("all")))
        self.txt_fields_canvas.create_window((0, 0), window=self.txt_fields_scrollable_frame, anchor="nw")
        self.txt_fields_canvas.configure(yscrollcommand=txt_scrollbar.set)
        
        self.txt_fields_canvas.pack(side="left", fill="both", expand=True)
        txt_scrollbar.pack(side="right", fill="y")
        
        # Button Frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x', pady=(10, 0))
        ttk.Button(button_frame, text="Save Mappings", command=self.save_mappings, style="Accent.TButton").pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.master.destroy).pack(side=tk.RIGHT)

        self.txt_field_row_widgets = []
        self.selected_row_index = -1
        
        self.add_field_header(self.txt_fields_scrollable_frame)
        self.recreate_field_rows()
        self.master.protocol("WM_DELETE_WINDOW", self.master.destroy)

    def use_main_mapping(self):
        """
        Copies the mapping configuration from 'Main TXT' to the current source,
        saves it, and reloads the display rows.
        """
        # 1. Get the configuration of the main source
        main_mapping = self.parent_gui.all_txt_mappings.get("Main TXT")
        
        if not main_mapping:
            messagebox.showerror("Error", "The 'Main TXT' mapping is empty or not configured. Cannot copy.", parent=self.master)
            return

        if not messagebox.askyesno(
            "Confirm Copy",
            f"Are you sure you want to overwrite all current mappings for '{self.source_key}' with the mappings from 'Main TXT'?",
            parent=self.master):
            return

        # 2. Deep copy the configuration and overwrite the current source's mapping
        new_config = [item.copy() for item in main_mapping]
        self.parent_gui.all_txt_mappings[self.source_key] = new_config
        
        # 3. Save to JSON and refresh the dialog display
        try:
            # We don't call save_mappings, as that method also destroys the window.
            # We must manually save the parent's data and refresh the display.
            self.parent_gui.save_settings()
            self.recreate_field_rows() 
            self.parent_gui.update_status(f"Mapping for {self.source_key} updated by copying 'Main TXT' settings.")
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save copied mapping: {e}", parent=self.master)
        
    def add_field_header(self, parent):
        parent.grid_columnconfigure(0, weight=0, minsize=50) 
        parent.grid_columnconfigure(1, weight=2, minsize=150) 
        parent.grid_columnconfigure(2, weight=2, minsize=150)
        parent.grid_columnconfigure(3, weight=2, minsize=150) 
        parent.grid_columnconfigure(4, weight=0, minsize=50) 
        parent.grid_columnconfigure(5, weight=0, minsize=80) 

        header_frame = ttk.Frame(parent, style="Header.TFrame", padding=(5,3))
        header_frame.grid(row=0, column=0, columnspan=7, sticky="ew") 

        ttk.Label(header_frame, text="Order", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=6, sticky='w')
        ttk.Label(header_frame, text="TXT Column", font=("Arial", 10, "bold")).grid(row=0, column=1, padx=6, sticky='w')
        ttk.Label(header_frame, text="Preview Data", font=("Arial", 10, "bold")).grid(row=0, column=2, padx=8, sticky='w')
        ttk.Label(header_frame, text="Excel Column", font=("Arial", 10, "bold")).grid(row=0, column=3, padx=6, sticky='w')
        ttk.Label(header_frame, text="Skip?", font=("Arial", 10, "bold")).grid(row=0, column=4, padx=6, sticky='w')
        ttk.Label(header_frame, text="Actions", font=("Arial", 10, "bold")).grid(row=0, column=5, padx=6, sticky='w')

        for i in range(6):
             header_frame.grid_columnconfigure(i, weight=parent.grid_columnconfigure(i).get('weight', 0), minsize=parent.grid_columnconfigure(i).get('minsize', 0))

    def recreate_field_rows(self, reselect_index=None):
        for widget in self.txt_fields_scrollable_frame.winfo_children():
            if int(widget.grid_info()["row"]) > 0:
                widget.destroy()
        self.txt_field_row_widgets.clear()
        
        mapping_config = self.parent_gui.all_txt_mappings.get(self.source_key, [])
        
        for i, config in enumerate(mapping_config):
            grid_row_index = i + 1
            parent_frame = self.txt_fields_scrollable_frame
            widgets_in_row = []

            # Order Label (0)
            order_label = ttk.Label(parent_frame, text=str(i + 1), anchor='center')
            order_label.grid(row=grid_row_index, column=0, padx=5, pady=2, sticky='ew')
            widgets_in_row.append(order_label)

            # TXT Field Entry (1)
            field_widget = ttk.Entry(parent_frame)
            field_widget.insert(0, config["field"])
            field_widget.grid(row=grid_row_index, column=1, padx=5, pady=2, sticky='ew')
            widgets_in_row.append(field_widget)

            # Preview Data Label (2)
            preview_label = ttk.Label(parent_frame, text="", anchor='w', foreground="blue")
            preview_label.grid(row=grid_row_index, column=2, padx=5, pady=2, sticky='ew')
            widgets_in_row.append(preview_label)
            
            # Excel Column Entry (3)
            column_entry = ttk.Entry(parent_frame)
            column_entry.insert(0, config.get("column_name", config["field"]))
            column_entry.grid(row=grid_row_index, column=3, padx=5, pady=2, sticky="ew")
            widgets_in_row.append(column_entry)
            
            # Skip Checkbox (4)
            skip_var = tk.BooleanVar(value=config.get("skip", False))
            skip_checkbox = ttk.Checkbutton(parent_frame, variable=skip_var)
            skip_checkbox.grid(row=grid_row_index, column=4, padx=(15,5), pady=2, sticky='w')
            widgets_in_row.append(skip_checkbox)

            # Remove Button (5)
            remove_btn = ttk.Button(parent_frame, text="Remove", width=8, style="Toolbutton",
                                    command=lambda idx=i: self.remove_field_row(idx))
            remove_btn.grid(row=grid_row_index, column=5, padx=5, pady=2, sticky='w')
            widgets_in_row.append(remove_btn)

            click_handler = lambda e, idx=i: self._select_row(idx)
            for widget in widgets_in_row:
                widget.bind("<Button-1>", click_handler)

            self.txt_field_row_widgets.append({
                "field_entry_widget": field_widget,
                "column_entry": column_entry,
                "skip_var": skip_var,
                "preview_label": preview_label,
                "all_widgets": widgets_in_row
            })
        
        if reselect_index is not None:
            self._select_row(reselect_index)
        else:
            self._select_row(-1)

        self.master.after_idle(lambda: self.txt_fields_canvas.config(scrollregion=self.txt_fields_canvas.bbox("all")))

    def _select_row(self, index):
        if self.selected_row_index != -1 and self.selected_row_index < len(self.txt_field_row_widgets):
            prev_row_info = self.txt_field_row_widgets[self.selected_row_index]
            for widget in prev_row_info.get("all_widgets", []):
                try: widget.configure(style=f"T{type(widget).__name__}")
                except tk.TclError: pass
                
        self.selected_row_index = index
        
        if index != -1 and index < len(self.txt_field_row_widgets):
            current_row_info = self.txt_field_row_widgets[index]
            for widget in current_row_info.get("all_widgets", []):
                try: widget.configure(style=f"Selected.T{type(widget).__name__}")
                except tk.TclError: pass

        self._update_move_buttons_state()

    def _update_move_buttons_state(self):
        config_list = self.parent_gui.all_txt_mappings.get(self.source_key, [])
        can_move_up = self.selected_row_index > 0
        can_move_down = self.selected_row_index != -1 and self.selected_row_index < len(config_list) - 1

        self.txt_move_up_btn.config(state=tk.NORMAL if can_move_up else tk.DISABLED)
        self.txt_move_down_btn.config(state=tk.NORMAL if can_move_down else tk.DISABLED)
        
    def move_selected_field(self, direction):
        current_index = self.selected_row_index
        if current_index == -1: return
        
        config_list = self.parent_gui.all_txt_mappings.get(self.source_key, [])
        total_items = len(config_list)

        if direction == "up" and current_index > 0:
            config_list[current_index], config_list[current_index - 1] = config_list[current_index - 1], config_list[current_index]
            self.recreate_field_rows(reselect_index=current_index - 1)
        elif direction == "down" and current_index < total_items - 1:
            config_list[current_index], config_list[current_index + 1] = config_list[current_index + 1], config_list[current_index]
            self.recreate_field_rows(reselect_index=current_index + 1)

    def add_field_row(self):
        config_list = self.parent_gui.all_txt_mappings.get(self.source_key, [])
        new_field_index = len(config_list) + 1
        
        config_list.append({
            "field": f"Custom_Field_{new_field_index}",
            "column_name": f"Custom_Col_{new_field_index}",
            "skip": False
        })
        self.recreate_field_rows(reselect_index=len(config_list) - 1)

    def remove_field_row(self, index_to_remove):
        config_list = self.parent_gui.all_txt_mappings.get(self.source_key, [])
        if not (0 <= index_to_remove < len(config_list)): return
        
        if messagebox.askyesno("Confirm Deletion", f"Are you sure you want to remove this field?", parent=self.master):
            del config_list[index_to_remove]
            
            new_selection = -1
            if self.selected_row_index == index_to_remove: new_selection = -1
            elif self.selected_row_index > index_to_remove: new_selection = self.selected_row_index - 1
            else: new_selection = self.selected_row_index

            self.recreate_field_rows(reselect_index=new_selection)
    
    def preview_latest_data(self):
        """Finds the latest file for the source and displays the data in the preview labels."""
        latest_file, data_parts = self.parent_gui._get_txt_file_data_for_preview(self.source_key)

        if not latest_file:
            messagebox.showinfo("File Not Found", f"No .txt/.npd/.csv files found for '{self.source_key}'.", parent=self.master)
            self.clear_preview()
            return
            
        if data_parts is None:
            messagebox.showerror("Read Error", f"An error occurred while reading or parsing the latest file:\n{os.path.basename(latest_file)}", parent=self.master)
            return

        for i, row_widgets in enumerate(self.txt_field_row_widgets):
            preview_label = row_widgets.get("preview_label")
            if preview_label:
                preview_label.config(text=data_parts[i].strip() if i < len(data_parts) else "<no data>")
        
        self.parent_gui.update_status(f"Preview loaded from {os.path.basename(latest_file)} for {self.source_key}")

    def clear_preview(self):
        for row_widgets in self.txt_field_row_widgets:
            preview_label = row_widgets.get("preview_label")
            if preview_label:
                preview_label.config(text="")
        self.parent_gui.update_status("Preview cleared.")

    def save_mappings(self):
        """
        Gathers data from entry widgets and updates the parent GUI's self.all_txt_mappings.
        """
        new_config = []
        
        # We save the data directly back to the active list in the parent GUI to retain order changes
        # However, we build a fresh list here and update the master list at the end for consistency
        for i, row_info in enumerate(self.txt_field_row_widgets):
            field_name = row_info["field_entry_widget"].get().strip()
            column_name = row_info["column_entry"].get().strip()
            skip_value = row_info["skip_var"].get()

            if not field_name:
                messagebox.showerror("Input Error", f"Field name at row {i+1} cannot be empty.", parent=self.master)
                return # Stop save if any field name is missing

            new_config.append({
                "field": field_name,
                "column_name": column_name if column_name else field_name, # Use field name as default column name
                "skip": skip_value
            })
            
        self.parent_gui.all_txt_mappings[self.source_key] = new_config
        self.parent_gui.save_settings() # Trigger the main save function
        self.parent_gui.update_status(f"Mappings for {self.source_key} saved.")
        self.master.destroy()

# --- Settings Window Class ---
class SettingsWindow:

    def __init__(self, master, parent_gui):
        self.master = master
        self.parent_gui = parent_gui
        self.master.title("Settings")
        # Keep the default size, scrolling will handle overflow
        self.master.geometry("1150x850") 
        self.master.minsize(800, 500)
        self.style = parent_gui.style

        # Main frame now uses grid for canvas and scrollbar
        self.main_frame = ttk.Frame(self.master)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.main_frame.rowconfigure(0, weight=1)    # Row for canvas/scrollbar
        self.main_frame.rowconfigure(1, weight=0)    # Row for buttons
        self.main_frame.columnconfigure(0, weight=1) # Column for canvas
        self.main_frame.columnconfigure(1, weight=0) # Column for scrollbar

        # --- Create Scrollable Area ---
        self.canvas = tk.Canvas(self.main_frame, borderwidth=0, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.main_frame, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # This frame goes INSIDE the canvas and will contain the notebook
        self.scrollable_content_frame = ttk.Frame(self.canvas)
        self.canvas_frame_id = self.canvas.create_window((0, 0), window=self.scrollable_content_frame, anchor="nw")

        # Bind events to make scrolling work
        self.scrollable_content_frame.bind("<Configure>", self.on_frame_configure)
        self.canvas.bind("<Configure>", self.on_canvas_configure)
        self.canvas.bind_all("<MouseWheel>", self.on_mousewheel) # Use bind_all for better capture

        # Place canvas and scrollbar on the grid
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")

        # --- Place Notebook inside the SCROLLABLE frame ---
        self.notebook = ttk.Notebook(self.scrollable_content_frame)
        self.notebook.pack(fill="both", expand=True, padx=5, pady=5)

        # Initialize selection tracking
        self.selected_txt_row_index = -1
        self.txt_move_up_btn = None
        self.txt_move_down_btn = None
        self.selected_static_row_index = -1
        self.static_move_up_btn = None
        self.static_move_down_btn = None
        
        # Initialize color picker variables
        self.new_day_bg_color_var = tk.StringVar()
        self.new_day_font_color_var = tk.StringVar()
        self.hourly_bg_color_var = tk.StringVar()
        self.hourly_font_color_var = tk.StringVar()

        # --- Create tabs (no changes here) ---
        self.create_file_paths_tab()
        self.create_generated_fields_tab()
        self.create_static_fields_tab()
        self.create_button_configuration_tab()
        self.create_event_codes_tab()
        self.create_monitored_folders_tab()
        self.create_auto_events_tab()
        self.create_timezone_tab()
        self.create_database_sync_tab()
        self.create_projects_tab()
        self._load_programmed_events_ui_state()

        # --- Bottom Buttons (remain in the main_frame) ---
        button_frame = ttk.Frame(self.main_frame)
        # Span both columns (canvas and scrollbar)
        button_frame.grid(row=1, column=0, columnspan=2, pady=(10, 0), sticky="e")
        ttk.Button(button_frame, text="Save", command=self.save_settings).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Save and Close", command=self.save_and_close, style="Accent.TButton").pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.master.destroy).pack(side=tk.RIGHT)

    # helper function that reads the file for a given source key
    
    def on_frame_configure(self, event=None):
        """Updates the canvas scroll region when the inner frame's size changes."""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def on_canvas_configure(self, event=None):
        """Ensures the inner frame width matches the canvas width."""
        self.canvas.itemconfig(self.canvas_frame_id, width=event.width)

    def on_mousewheel(self, event):
        """Handles cross-platform mouse wheel scrolling."""
        # Check if the mouse is over the main canvas before scrolling
        if self.canvas.winfo_containing(event.x_root, event.y_root) == self.canvas:
            if sys.platform == "win32":
                delta = -1 * int(event.delta / 120)
            elif event.num == 4: # Linux scroll up
                delta = -1
            elif event.num == 5: # Linux scroll down
                delta = 1
            else: # Fallback for other systems/events
                delta = 0
            self.canvas.yview_scroll(delta, "units")

    def save_and_close(self):
        self.save_settings()
        self.master.destroy()

    # --- Tab Creation Methods ---

    def create_database_sync_tab(self):
        """Creates the UI tab for managing automatic database synchronization."""
        tab = ttk.Frame(self.notebook, padding=20)
        self.notebook.add(tab, text="Database Sync")

        # --- Description ---
        desc_frame = ttk.Frame(tab)
        desc_frame.pack(fill='x', pady=(0, 10))
        ttk.Label(
            desc_frame,
            text="Configure a periodic synchronization to keep the SQLite database mirror up-to-date with any manual changes made to the Excel file.",
            wraplength=900
        ).pack(anchor='w')

        # --- Main content frame ---
        content_frame = ttk.LabelFrame(tab, text="Automatic Synchronization", padding=15)
        content_frame.pack(fill='x', expand=False, anchor='n')
        content_frame.columnconfigure(1, weight=0)

        # --- Enable/Disable Checkbox ---
        enable_check = ttk.Checkbutton(
            content_frame,
            text="Enable automatic sync",
            variable=self.parent_gui.auto_sync_enabled_var,
            style="Large.TCheckbutton"
        )
        enable_check.grid(row=0, column=0, columnspan=3, padx=5, pady=(5, 10), sticky="w")
        ToolTip(enable_check, "When checked, the application will periodically re-sync the entire Excel file to the database.")

        # --- Interval Spinbox ---
        ttk.Label(content_frame, text="Sync Interval (minutes):").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        interval_spinbox = ttk.Spinbox(
            content_frame,
            from_=1,
            to=1440, # 24 hours
            increment=1,
            textvariable=self.parent_gui.auto_sync_interval_min_var,
            width=10
        )
        interval_spinbox.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        ToolTip(interval_spinbox, "How often (in minutes) the sync should run.")
    
    def create_static_fields_tab(self):
        """
        Creates a new tab for configuring static fields read directly from Excel cells.
        """
        tab = ttk.Frame(self.notebook, padding=20)
        self.notebook.add(tab, text="Static Fields")

        ttk.Label(tab, text="Map static values from specific Excel cells to new columns in your log. Use the format: ='SheetName'!A1. Check 'Skip' to ignore a field entirely.", wraplength=900, justify=tk.LEFT).pack(pady=(0, 10), anchor='w')

        # Control buttons for adding/removing fields
        controls_frame = ttk.Frame(tab)
        controls_frame.pack(fill='x', pady=(0, 10))

        ttk.Button(controls_frame, text="Add New Field", command=self.add_static_field_row).pack(side=tk.LEFT, padx=5)
        
        # Keep the spacer to push the next button to the right
        spacer = ttk.Frame(controls_frame)
        spacer.pack(side=tk.LEFT, expand=True, fill='x')

        # Remove the Move Up and Move Down buttons
        self.static_move_up_btn = None
        self.static_move_down_btn = None

        # Canvas and Scrollbar for the dynamic field list
        self.static_fields_canvas = tk.Canvas(tab, borderwidth=0, background="#ffffff")
        static_scrollbar = ttk.Scrollbar(tab, orient="vertical", command=self.static_fields_canvas.yview)
        self.static_fields_scrollable_frame = ttk.Frame(self.static_fields_canvas, style="Row0.TFrame")
        self.static_fields_scrollable_frame.bind("<Configure>", lambda e: self.static_fields_canvas.configure(scrollregion=self.static_fields_canvas.bbox("all")))
        self.static_fields_canvas_window = self.static_fields_canvas.create_window((0, 0), window=self.static_fields_scrollable_frame, anchor="nw")
        self.static_fields_canvas.configure(yscrollcommand=static_scrollbar.set)
        self.static_fields_canvas.pack(side="left", fill="both", expand=True, padx=(0,0), pady=0)
        static_scrollbar.pack(side="right", fill="y", padx=(0,0), pady=0)
        
        def _on_mousewheel_static(event):
            if event.num == 4: delta = -1
            elif event.num == 5: delta = 1
            elif hasattr(event, 'delta'): delta = -int(event.delta / 120)
            else: delta = 0
            self.static_fields_canvas.yview_scroll(delta, "units")
        self.static_fields_canvas.bind("<MouseWheel>", _on_mousewheel_static)
        self.static_fields_canvas.bind("<Button-4>", _on_mousewheel_static)
        self.static_fields_canvas.bind("<Button-5>", _on_mousewheel_static)

        # Store widgets for each row dynamically
        self.static_field_row_widgets = []
        self.add_static_field_header(self.static_fields_scrollable_frame)
        self.recreate_static_field_rows()
        self._update_static_move_buttons_state()

    # --- New helper methods for Static Fields tab ---
    def add_static_field_header(self, parent):
        """Adds a header row to the static field mapping section."""
        # New column configuration
        parent.grid_columnconfigure(0, weight=2, minsize=150) # Description
        parent.grid_columnconfigure(1, weight=2, minsize=150) # Excel Column Name
        parent.grid_columnconfigure(2, weight=2, minsize=250) # Static Cell Reference
        parent.grid_columnconfigure(3, weight=0, minsize=50)  # Skip?
        parent.grid_columnconfigure(4, weight=0, minsize=80)  # Actions

        header_frame = ttk.Frame(parent, style="Header.TFrame", padding=(5,3))
        header_frame.grid(row=0, column=0, columnspan=5, sticky="ew") # Update columnspan

        # Update the header labels with the new 'Description' column
        ttk.Label(header_frame, text="Description", font=("Arial", 10, "bold")).grid(row=0, column=1, padx=6, sticky='w') 
        ttk.Label(header_frame, text="Excel Column", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=6, sticky='w')
        ttk.Label(header_frame, text="Static Cell Reference", font=("Arial", 10, "bold")).grid(row=0, column=2, padx=6, sticky='w')
        ttk.Label(header_frame, text="Skip?", font=("Arial", 10, "bold")).grid(row=0, column=3, padx=6, sticky='w')
        ttk.Label(header_frame, text="Actions", font=("Arial", 10, "bold")).grid(row=0, column=4, padx=6, sticky='w')

        for i in range(5): # Update range to 5
            header_frame.grid_columnconfigure(i, weight=parent.grid_columnconfigure(i).get('weight', 0), minsize=parent.grid_columnconfigure(i).get('minsize', 0))

    def recreate_static_field_rows(self, reselect_index=None):
        # Clear existing widgets except the header
        for widget in self.static_fields_scrollable_frame.winfo_children():
            if int(widget.grid_info()["row"]) > 0:
                widget.destroy()
        self.static_field_row_widgets.clear()
        
        # Iterate over the static_field_configs list to create each row
        for i, config in enumerate(self.parent_gui.static_field_configs):
            # All of the following lines must be indented
            grid_row_index = i + 1
            parent_frame = self.static_fields_scrollable_frame
            widgets_in_row = []

            # Create the widget for the 'Excel Column' field
            column_entry = ttk.Entry(parent_frame)
            column_entry.insert(0, config.get("field", ""))
            column_entry.grid(row=grid_row_index, column=0, padx=5, pady=2, sticky='ew')
            widgets_in_row.append(column_entry)
            ToolTip(column_entry, "The header for the column in your Excel log.")

            # Create the widget for the 'Description' field
            description_entry = ttk.Entry(parent_frame)
            description_entry.insert(0, config.get("description", ""))
            description_entry.grid(row=grid_row_index, column=1, padx=5, pady=2, sticky='ew')
            widgets_in_row.append(description_entry)
            ToolTip(description_entry, "A brief description of this static field.")

            # Create the widget for the 'Static Cell Reference' field
            cell_entry = ttk.Entry(parent_frame)
            cell_entry.insert(0, config.get("column_name", ""))
            cell_entry.grid(row=grid_row_index, column=2, padx=5, pady=2, sticky="ew")
            widgets_in_row.append(cell_entry)
            ToolTip(cell_entry, "Enter the static cell reference, e.g., ='SheetName'!A1")

            # Create the widget for the 'Skip' checkbox
            skip_var = tk.BooleanVar(value=config.get("skip", False))
            skip_checkbox = ttk.Checkbutton(parent_frame, variable=skip_var)
            skip_checkbox.grid(row=grid_row_index, column=3, padx=5, pady=2, sticky='w')
            widgets_in_row.append(skip_checkbox)

            # Create the widget for the 'Remove' button
            remove_btn = ttk.Button(parent_frame, text="Remove", width=8, style="Toolbutton",
                                    command=lambda idx=i: self.remove_static_field_row(idx))
            remove_btn.grid(row=grid_row_index, column=4, padx=5, pady=2, sticky='w')
            widgets_in_row.append(remove_btn)

            # Bind the click handler to all widgets in the row for selection
            click_handler = lambda e, idx=i: self._select_static_row(idx)
            for widget in widgets_in_row:
                widget.bind("<Button-1>", click_handler)

            # Store the widget references in the instance variable
            self.static_field_row_widgets.append({
                "column_entry": column_entry,
                "description_entry": description_entry,
                "cell_entry": cell_entry,
                "skip_var": skip_var,
                "all_widgets": widgets_in_row
            })
        if reselect_index is not None:
            self._select_static_row(reselect_index)
        else:
            self._select_static_row(-1)

        self.master.after_idle(lambda: self.static_fields_canvas.config(scrollregion=self.static_fields_canvas.bbox("all")))

    def _select_static_row(self, index):
        """Highlights the selected row in the Static Fields tab."""
        if self.selected_static_row_index != -1 and self.selected_static_row_index < len(self.static_field_row_widgets):
            prev_row_info = self.static_field_row_widgets[self.selected_static_row_index]
            for widget in prev_row_info.get("all_widgets", []):
                try: widget.configure(style=f"T{type(widget).__name__}")
                except tk.TclError: pass
        self.selected_static_row_index = index
        if index != -1 and index < len(self.static_field_row_widgets):
            current_row_info = self.static_field_row_widgets[index]
            for widget in current_row_info.get("all_widgets", []):
                try: widget.configure(style=f"Selected.T{type(widget).__name__}")
                except tk.TclError: pass
        self._update_static_move_buttons_state()

    def _update_static_move_buttons_state(self):
        """Enables/disables move buttons for the Static Fields tab."""
        can_move_up = (self.selected_static_row_index > 0)
        can_move_down = (self.selected_static_row_index != -1 and self.selected_static_row_index < len(self.parent_gui.static_field_configs) - 1)
        if self.static_move_up_btn: self.static_move_up_btn.config(state=tk.NORMAL if can_move_up else tk.DISABLED)
        if self.static_move_down_btn: self.static_move_down_btn.config(state=tk.NORMAL if can_move_down else tk.DISABLED)
    
    def add_static_field_row(self):
        """Adds a new row for a custom static field."""
        new_field_index = len(self.parent_gui.static_field_configs) + 1
        
        # Correctly append the new configuration to the parent GUI's list
        # The 'skip' field needs to be explicitly included with a default value.
        self.parent_gui.static_field_configs.append({
            "field": f"Static_Col_{new_field_index}",
            "description": "", 
            "column_name": f"='SheetName'!A{new_field_index}",
            "skip": False  
        })
        
        newly_added_index = len(self.parent_gui.static_field_configs) - 1
        self.recreate_static_field_rows(reselect_index=newly_added_index)
        self.parent_gui.update_status(f"Added new static field 'Static_Col_{new_field_index}'.")

    def remove_static_field_row(self, index_to_remove):
        """Removes a static field row by index."""
        if not (0 <= index_to_remove < len(self.parent_gui.static_field_configs)):
            return
        if messagebox.askyesno("Confirm Deletion", f"Are you sure you want to remove this static field?", parent=self.master):
            del self.parent_gui.static_field_configs[index_to_remove]
            if self.selected_static_row_index == index_to_remove: self.selected_static_row_index = -1
            elif self.selected_static_row_index > index_to_remove: self.selected_static_row_index -= 1
            self.recreate_static_field_rows(reselect_index=self.selected_static_row_index)
            self.parent_gui.update_status("Static field removed.")
    
    def move_selected_static_field(self, direction):
        """Moves the selected static field up or down."""
        current_index = self.selected_static_row_index
        if current_index == -1:
            messagebox.showinfo("No Selection", "Please select a row to move.", parent=self.master)
            return
        total_items = len(self.parent_gui.static_field_configs)
        if direction == "up" and current_index > 0:
            self.parent_gui.static_field_configs[current_index], self.parent_gui.static_field_configs[current_index - 1] = self.parent_gui.static_field_configs[current_index - 1], self.parent_gui.static_field_configs[current_index]
            self.selected_static_row_index -= 1
            self.recreate_static_field_rows(reselect_index=self.selected_static_row_index)
        elif direction == "down" and current_index < total_items - 1:
            self.parent_gui.static_field_configs[current_index], self.parent_gui.static_field_configs[current_index + 1] = self.parent_gui.static_field_configs[current_index + 1], self.parent_gui.static_field_configs[current_index]
            self.selected_static_row_index += 1
            self.recreate_static_field_rows(reselect_index=self.selected_static_row_index)

    def create_timezone_tab(self):
        """Creates the UI tab for managing the UTC offset."""
        tab = ttk.Frame(self.notebook, padding=20)
        self.notebook.add(tab, text="Timezone")

        # --- Description ---
        desc_frame = ttk.Frame(tab)
        desc_frame.pack(fill='x', pady=(0, 10))
        ttk.Label(desc_frame, text="Configure the local time offset from UTC. This will be used to populate the 'Local Time' column.", wraplength=900).pack(anchor='w')

        # --- Main content frame ---
        content_frame = ttk.LabelFrame(tab, text="UTC Offset", padding=15)
        content_frame.pack(fill='x', expand=False, anchor='n')
        content_frame.columnconfigure(1, weight=0)

        ttk.Label(content_frame, text="Offset (hours):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        
        # Use a Spinbox for better user experience
        offset_spinbox = ttk.Spinbox(
            content_frame,
            from_=-14.0,
            to=14.0,
            increment=0.5,
            textvariable=self.parent_gui.time_offset_hours, # Bind directly to the DoubleVar
            width=10
        )
        offset_spinbox.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        ToolTip(offset_spinbox, "Enter the local time offset from UTC in hours (e.g., -5 for EST, +1 for BST). Decimals are allowed.")

    def create_event_codes_tab(self):
        """Creates the UI tab for managing the event codes configuration."""
        tab = ttk.Frame(self.notebook, padding=20)
        self.notebook.add(tab, text="Event Codes")

        # --- Description ---
        desc_frame = ttk.Frame(tab)
        desc_frame.pack(fill='x', pady=(0, 10))
        ttk.Label(desc_frame, text="Create and manage event codes here. These codes can be assigned to buttons to be logged in the 'Code' column.", wraplength=900).pack(anchor='w')
        
        # --- Main content frame ---
        content_frame = ttk.Frame(tab)
        content_frame.pack(fill='both', expand=True)
        content_frame.columnconfigure(0, weight=1)
        content_frame.rowconfigure(0, weight=1)

        # --- Treeview to display codes ---
        tree_frame = ttk.Frame(content_frame)
        tree_frame.grid(row=0, column=0, sticky='nsew', pady=(0, 10))
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.event_codes_tree = ttk.Treeview(tree_frame, columns=('Code', 'Description'), show='headings', height=10)
        self.event_codes_tree.heading('Code', text='Code')
        self.event_codes_tree.heading('Description', text='Description')
        self.event_codes_tree.column('Code', width=150, stretch=False)
        self.event_codes_tree.column('Description', width=400, stretch=True)
        
        tree_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.event_codes_tree.yview)
        self.event_codes_tree.configure(yscrollcommand=tree_scrollbar.set)

        self.event_codes_tree.grid(row=0, column=0, sticky='nsew')
        tree_scrollbar.grid(row=0, column=1, sticky='ns')

        # --- Buttons for managing codes ---
        button_frame = ttk.Frame(content_frame)
        button_frame.grid(row=1, column=0, sticky='e')

        ttk.Button(button_frame, text="Add Code...", command=self.add_event_code).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Edit Selected...", command=self.edit_event_code).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Delete Selected", command=self.delete_event_code).pack(side=tk.LEFT, padx=5)

    def populate_event_codes_tree(self):
        """Clears and re-populates the event codes treeview from the parent GUI's data."""
        # Clear existing items
        for item in self.event_codes_tree.get_children():
            self.event_codes_tree.delete(item)
        
        # Populate with new data
        codes = self.parent_gui.event_codes
        for code, description in sorted(codes.items()):
            self.event_codes_tree.insert('', 'end', values=(code, description))

    def save_event_codes_to_file(self):
        """Persists the current event codes into the active project JSON via the main save."""
        try:
            # Save entire settings (including event codes) to the active project JSON
            self.parent_gui.save_settings()
            self.parent_gui.update_status("Event codes saved to project.")
        except Exception as e:
            messagebox.showerror("Save Error", f"Could not save event codes to project file:\n{e}", parent=self.master)

    def _show_event_code_dialog(self, title, initial_code="", initial_desc=""):
        """Helper dialog for adding/editing event codes."""
        dialog = Toplevel(self.master)
        dialog.title(title)
        dialog.transient(self.master)
        dialog.grab_set()
        dialog.resizable(False, False)

        frame = ttk.Frame(dialog, padding="15")
        frame.pack(fill='both', expand=True)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Code:").grid(row=0, column=0, sticky='w', pady=5)
        code_entry = ttk.Entry(frame, width=40)
        code_entry.grid(row=0, column=1, sticky='ew', pady=5)
        code_entry.insert(0, initial_code)
        
        ttk.Label(frame, text="Description:").grid(row=1, column=0, sticky='w', pady=5)
        desc_entry = ttk.Entry(frame, width=40)
        desc_entry.grid(row=1, column=1, sticky='ew', pady=5)
        desc_entry.insert(0, initial_desc)

        result = {}
        def on_ok():
            result['code'] = code_entry.get().strip()
            result['desc'] = desc_entry.get().strip()
            if not result['code']:
                messagebox.showwarning("Input Error", "Code cannot be empty.", parent=dialog)
                return
            dialog.destroy()

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=(10,0), sticky='e')
        ttk.Button(button_frame, text="OK", command=on_ok).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT)

        dialog.wait_window()
        return result

    def add_event_code(self):
        result = self._show_event_code_dialog("Add New Event Code")
        if result and result.get('code'):
            if result['code'] in self.parent_gui.event_codes:
                messagebox.showwarning("Duplicate Code", "This event code already exists.", parent=self.master)
                return
            self.parent_gui.event_codes[result['code']] = result['desc']
            self.save_event_codes_to_file()
            self.populate_event_codes_tree()

    def edit_event_code(self):
        selected_item = self.event_codes_tree.focus()
        if not selected_item:
            messagebox.showinfo("No Selection", "Please select an event code to edit.", parent=self.master)
            return
        
        item_values = self.event_codes_tree.item(selected_item, 'values')
        old_code, old_desc = item_values[0], item_values[1]

        result = self._show_event_code_dialog("Edit Event Code", initial_code=old_code, initial_desc=old_desc)
        if result and result.get('code'):
            new_code = result['code']
            new_desc = result['desc']
            # Remove old code first
            del self.parent_gui.event_codes[old_code]
            # Add new/updated code
            self.parent_gui.event_codes[new_code] = new_desc
            self.save_event_codes_to_file()
            self.populate_event_codes_tree()

    def delete_event_code(self):
        selected_item = self.event_codes_tree.focus()
        if not selected_item:
            messagebox.showinfo("No Selection", "Please select an event code to delete.", parent=self.master)
            return

        code_to_delete = self.event_codes_tree.item(selected_item, 'values')[0]
        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete the code '{code_to_delete}'?", parent=self.master):
            del self.parent_gui.event_codes[code_to_delete]
            self.save_event_codes_to_file()
            self.populate_event_codes_tree()

    def create_file_paths_tab(self):
        tab = ttk.Frame(self.notebook, padding=20)
        self.notebook.add(tab, text="File Paths")
        
        # --- Excel Log File ---
        log_frame = ttk.LabelFrame(tab, text="Excel Log File (.xlsx or .xlsb)", padding=15)
        log_frame.pack(fill="x", pady=(0, 15))
        log_frame.columnconfigure(1, weight=1)
        self.log_file_label = ttk.Label(log_frame, text="Path:", anchor='e')
        self.log_file_label.grid(row=0, column=0, padx=(0, 5), pady=5, sticky='w')
        self.log_file_entry = ttk.Entry(log_frame, width=80)
        self.log_file_entry.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        log_browse_btn = ttk.Button(log_frame, text="Browse...", command=self.select_excel_file)
        log_browse_btn.grid(row=0, column=2, padx=(5, 0), pady=5)
        ToolTip(log_browse_btn, "Select the main Excel file for logging.")
        ToolTip(self.log_file_entry, "Full path to the .xlsx or .xlsb file where all log entries will be written.")

        # --- SQLite Database File ---
        db_frame = ttk.LabelFrame(tab, text="SQLite Database Mirror File (.db)", padding=15)
        db_frame.pack(fill="x", pady=(0, 15))
        db_frame.columnconfigure(1, weight=1)
        
        # --- Red Warning Text ---
        ttk.Label(
            db_frame, 
            text="⚠️ To be used when Reach Horizon Spreadsheet is not in use", 
            foreground="red",
            font=("Arial", 9, "bold")
        ).grid(row=0, column=0, columnspan=3, padx=5, pady=(0, 10), sticky='w')
        
        self.db_file_label = ttk.Label(db_frame, text="Path:", anchor='e')
        self.db_file_label.grid(row=1, column=0, padx=(0, 5), pady=5, sticky='w')
        
        self.db_file_entry = ttk.Entry(db_frame, width=80)
        self.db_file_entry.grid(row=1, column=1, padx=5, pady=5, sticky='ew')
        
        db_browse_btn = ttk.Button(db_frame, text="Browse...", command=self.select_sqlite_file)
        db_browse_btn.grid(row=1, column=2, padx=(5, 0), pady=5)
        
        ToolTip(db_browse_btn, "Select the location to save the SQLite database file.")
        ToolTip(self.db_file_entry, "Full path to the .db file where the Excel data will be mirrored. If left blank, it will be created next to the Excel file.")

        # --- Main Navigation TXT Data Folder ---
        txt_sources_container = ttk.Frame(tab)
        txt_sources_container.pack(fill='x', expand=True, anchor='n')
        txt_sources_container.columnconfigure(0, weight=1)

        # --- Helper function to create each TXT source entry ( ---
        def create_txt_source_frame(parent, title, source_key, name_entry_var, path_entry_var):
            frame = ttk.LabelFrame(parent, text=title, padding=15)
            frame.grid(sticky='ew', pady=(0, 15))
            parent.columnconfigure(0, weight=1)
            frame.columnconfigure(1, weight=1) # Allow path entry to expand

            ttk.Label(frame, text="Source Name:").grid(row=0, column=0, padx=(0, 5), pady=5, sticky='w')
            name_entry = ttk.Entry(frame, textvariable=name_entry_var, width=25)
            name_entry.grid(row=0, column=1, padx=5, pady=5, sticky='w')
            ToolTip(name_entry, "Set a custom, user-friendly name for this data source (e.g., 'Vessel', 'WROV').")

            ttk.Label(frame, text="Folder Path:").grid(row=1, column=0, padx=(0, 5), pady=5, sticky='w')
            path_entry = ttk.Entry(frame, textvariable=path_entry_var, width=80)
            path_entry.grid(row=1, column=1, padx=5, pady=5, sticky='ew')
            browse_btn = ttk.Button(frame, text="Browse...", command=lambda: self.select_txt_folder(path_entry))
            browse_btn.grid(row=1, column=2, padx=(5, 0), pady=5)
            ToolTip(browse_btn, "Select the folder containing the navigation TXT files for this source.")
            
            # --- MAPPING BUTTON ---
            map_btn = ttk.Button(frame, text="Map Fields", command=lambda key=source_key: self.open_mapping_dialog(key))
            map_btn.grid(row=0, column=2, padx=(5, 0), pady=5) # Place on the same row as Source Name
            ToolTip(map_btn, f"Open the field mapping configuration for {source_key}. (Per-file mapping).")
            
            return name_entry, path_entry

        # Create StringVars (unchanged)
        self.txt_name_main_var = tk.StringVar()
        self.txt_path_main_var = tk.StringVar()
        self.txt_name_set2_var = tk.StringVar()
        self.txt_path_set2_var = tk.StringVar()
        self.txt_name_set3_var = tk.StringVar()
        self.txt_path_set3_var = tk.StringVar()
        self.txt_name_set4_var = tk.StringVar()
        self.txt_path_set4_var = tk.StringVar()
        self.txt_name_set5_var = tk.StringVar()
        self.txt_path_set5_var = tk.StringVar()

        # Create the five source blocks using the helper ( to pass source_key)
        create_txt_source_frame(txt_sources_container, "Main Vehicle Navigation (Main TXT Data)", "Main TXT", self.txt_name_main_var, self.txt_path_main_var)
        create_txt_source_frame(txt_sources_container, "Additional Vehicle Navigation Data (TXT Source 2)", "TXT Source 2", self.txt_name_set2_var, self.txt_path_set2_var)
        create_txt_source_frame(txt_sources_container, "Additional Vehicle Navigation Data (TXT Source 3)", "TXT Source 3", self.txt_name_set3_var, self.txt_path_set3_var)
        create_txt_source_frame(txt_sources_container, "Additional Vehicle Navigation Data (TXT Source 4)", "TXT Source 4", self.txt_name_set4_var, self.txt_path_set4_var)
        create_txt_source_frame(txt_sources_container, "Additional Vehicle Navigation Data (TXT Source 5)", "TXT Source 5", self.txt_name_set5_var, self.txt_path_set5_var)
        

    def open_mapping_dialog(self, source_key):
        """Opens a dialog to configure the mapping for a single source key."""
        # This check is crucial to ensure the parent_gui object is available and the window is the current one
        if not self.parent_gui or not self.master.winfo_exists():
            messagebox.showerror("Error", "Application state invalid.", parent=self.master)
            return

        dialog_window = tk.Toplevel(self.master)
        
        # Pass the parent_gui instance, not the settings window instance
        TxtMappingDialog(dialog_window, self.parent_gui, source_key) 
    
    def select_excel_file(self):
        initial_dir = os.path.dirname(self.log_file_entry.get()) if self.log_file_entry.get() else os.getcwd()
        file_path = filedialog.askopenfilename(initialdir=initial_dir, filetypes=[("Excel files", ["*.xlsx",".xlsb"])], parent=self.master, title="Select Excel Log File")
        if file_path: self.log_file_entry.delete(0, tk.END); self.log_file_entry.insert(0, file_path)

    def select_txt_folder(self, entry_widget):
        current_path = entry_widget.get()
        initial_dir = current_path if os.path.isdir(current_path) else os.path.dirname(current_path) if current_path else os.getcwd()
        folder_path = filedialog.askdirectory(initialdir=initial_dir, parent=self.master, title="Select Navigation TXT Folder")
        if folder_path: entry_widget.delete(0, tk.END); entry_widget.insert(0, folder_path)

    def select_sqlite_file(self):
        """Opens a 'save as' dialog to choose the SQLite DB file location."""
        initial_dir = os.path.dirname(self.db_file_entry.get()) if self.db_file_entry.get() else os.getcwd()
        
        # The file_path variable is created here
        file_path = filedialog.asksaveasfilename(
            initialdir=initial_dir,
            filetypes=[("SQLite Database", "*.db"), ("All files", "*.*")],
            parent=self.master,
            title="Select SQLite Database File"
        )
        
        # This check now works correctly because file_path is defined above
        if file_path:
            # Ensure the file has the .db extension
            if not file_path.lower().endswith('.db'):
                file_path += '.db'
            self.db_file_entry.delete(0, tk.END)
            self.db_file_entry.insert(0, file_path)

    
            
    def create_generated_fields_tab(self):
        tab = ttk.Frame(self.notebook, padding=20)
        self.notebook.add(tab, text="Generated Fields")

        ttk.Label(tab, text="Assign Excel column names to data generated by the application (e.g., timestamps). The 'Source' column is for information only.", wraplength=900, justify=tk.LEFT).pack(pady=(0, 10), anchor='w')

        fields_frame = ttk.Frame(tab)
        fields_frame.pack(fill='both', expand=True)

        header = ttk.Frame(fields_frame, style="Header.TFrame", padding=(5,3))
        header.pack(fill='x', pady=(0, 5))
        header.columnconfigure(0, weight=1)
        header.columnconfigure(1, weight=1)
        header.columnconfigure(2, weight=1)
        
        ttk.Label(header, text="Field Name", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky='w', padx=5)
        ttk.Label(header, text="Data Source", font=("Arial", 10, "bold")).grid(row=0, column=1, sticky='w', padx=5)
        ttk.Label(header, text="Excel Column", font=("Arial", 10, "bold")).grid(row=0, column=2, sticky='w', padx=5)

        self.generated_field_widgets = []
        for i, config in enumerate(self.parent_gui.generated_fields_config):
            style_name = f"Row{i % 2}.TFrame"
            row_frame = ttk.Frame(fields_frame, padding=(5, 2), style=style_name)
            row_frame.pack(fill='x')
            row_frame.columnconfigure(0, weight=1)
            row_frame.columnconfigure(1, weight=1)
            row_frame.columnconfigure(2, weight=1)
            
            ttk.Label(row_frame, text=config["field"], style=style_name.replace("Frame","Label")).grid(row=0, column=0, sticky='w', padx=5)
            ttk.Label(row_frame, text=config.get("source", "N/A")).grid(row=0, column=1, sticky='w', padx=5)

            entry = ttk.Entry(row_frame)
            entry.insert(0, config["column_name"])
            entry.grid(row=0, column=2, sticky='ew', padx=5)

            skip_var = tk.BooleanVar(value=config.get("skip", False))
            # You could add a skip checkbox here if desired, similar to other tabs

            self.generated_field_widgets.append({'entry': entry, 'skip_var': skip_var})
    
    def add_txt_field_header(self, parent):
        """Adds a header row to the TXT field mapping section."""
        
        # Apply column configuration to the single, shared parent frame
        parent.grid_columnconfigure(0, weight=2, minsize=50) # TXT Field Name
        parent.grid_columnconfigure(1, weight=2, minsize=150) # TXT Field Name
        parent.grid_columnconfigure(2, weight=2, minsize=150) # Target Excel Column
        parent.grid_columnconfigure(3, weight=2, minsize=150) # Preview Data
        parent.grid_columnconfigure(4, weight=0, minsize=50)  # Skip
        parent.grid_columnconfigure(5, weight=0, minsize=80)  # Actions

        header_frame = ttk.Frame(parent, style="Header.TFrame", padding=(5,3))
        header_frame.grid(row=0, column=0, columnspan=8, sticky="ew") # Span all columns

        # Place labels inside the header_frame, but they will align because the parent of header_frame has the config
        ttk.Label(header_frame, text="Order", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=6, sticky='w')
        ttk.Label(header_frame, text="TXT Column", font=("Arial", 10, "bold")).grid(row=0, column=1, padx=6, sticky='w')
        ttk.Label(header_frame, text="Excel Column", font=("Arial", 10, "bold")).grid(row=0, column=3, padx=6, sticky='w')
        ttk.Label(header_frame, text="Preview TXT Data", font=("Arial", 10, "bold")).grid(row=0, column=2, padx=8, sticky='w')
        ttk.Label(header_frame, text="Skip?", font=("Arial", 10, "bold")).grid(row=0, column=5, padx=6, sticky='w')
        ttk.Label(header_frame, text="Actions", font=("Arial", 10, "bold")).grid(row=0, column=6, padx=6, sticky='w')

        # Also apply the same column configure to the header_frame itself so its internal labels space out correctly
        for i in range(6):
            header_frame.grid_columnconfigure(i, weight=parent.grid_columnconfigure(i).get('weight', 0), minsize=parent.grid_columnconfigure(i).get('minsize', 0))

    def create_button_configuration_tab(self):
        tab = ttk.Frame(self.notebook, padding=20)
        self.notebook.add(tab, text="Button Configuration")

        num_buttons_frame = ttk.Frame(tab)
        num_buttons_frame.pack(pady=5, anchor='w')
        ttk.Label(num_buttons_frame, text=f"Number of Custom Buttons (0-{self.parent_gui.MAX_CUSTOM_BUTTONS}):").pack(side='left', padx=5)
        self.num_buttons_entry = ttk.Entry(num_buttons_frame, width=5)
        self.num_buttons_entry.pack(side='left', padx=5)
        ToolTip(self.num_buttons_entry, "Enter the number of custom event buttons needed (max 10).")

        update_btn = ttk.Button(num_buttons_frame, text="Update List", command=self.update_num_custom_buttons)
        update_btn.pack(side='left', padx=5)
        ToolTip(update_btn, "Update the list below to show the specified number of button configurations.")

        # Header
        header_frame = ttk.Frame(tab, style="Header.TFrame", padding=(5,3))
        header_frame.pack(anchor='w', pady=(15,5))

        header_frame.grid_columnconfigure(0, weight=0, minsize=40) 
        header_frame.grid_columnconfigure(1, weight=1, minsize=135) 
        header_frame.grid_columnconfigure(2, weight=2, minsize=200) 
        header_frame.grid_columnconfigure(3, weight=0, minsize=80) 
        header_frame.grid_columnconfigure(4, weight=0, minsize=80) 
        header_frame.grid_columnconfigure(5, weight=0, minsize=80) 

        ttk.Label(header_frame, text="Button #", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=(5,0), sticky='w')
        ttk.Label(header_frame, text="Button Text", font=("Arial", 10, "bold")).grid(row=0, column=1, padx=5, sticky='ew')
        ttk.Label(header_frame, text="Event Text (for Log)", font=("Arial", 10, "bold")).grid(row=0, column=2, padx=5, sticky='ew')
        ttk.Label(header_frame, text="Event Code", font=("Arial", 10, "bold")).grid(row=0, column=3, padx=5, sticky='w')
        ttk.Label(header_frame, text="Event Source", font=("Arial", 10, "bold")).grid(row=0, column=4, padx=5, sticky='w')
        ttk.Label(header_frame, text="Tab Group", font=("Arial", 10, "bold")).grid(row=0, column=5, padx=5, sticky='w')

        # Scrollable region
        canvas = tk.Canvas(tab)
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )
        window = canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Replace direct frame with scrollable one
        self.custom_button_entries_frame = self.scrollable_frame
        self.custom_button_widgets = []
    
    def _select_txt_row(self, index):
        # This helper function should already be correct, but is included for completeness
        if hasattr(self, 'selected_txt_row_index') and self.selected_txt_row_index != -1 and self.selected_txt_row_index < len(self.txt_field_row_widgets):
            # Deselect previous row
            # ... (omitting deselection logic for brevity, your existing logic is likely fine) ...
            pass
        self.selected_txt_row_index = index
        # ... (omitting selection logic for brevity, your existing logic is likely fine) ...
        self._update_txt_move_buttons_state()

    def _update_txt_move_buttons_state(self):
        # CORRECTED: Use the new all_txt_mappings structure
        main_txt_config = self.parent_gui.all_txt_mappings.get("Main TXT", [])
        can_move_up = hasattr(self, 'selected_txt_row_index') and self.selected_txt_row_index > 0
        can_move_down = hasattr(self, 'selected_txt_row_index') and self.selected_txt_row_index != -1 and self.selected_txt_row_index < len(main_txt_config) - 1

        if self.txt_move_up_btn:
            self.txt_move_up_btn.config(state=tk.NORMAL if can_move_up else tk.DISABLED)
        if self.txt_move_down_btn:
            self.txt_move_down_btn.config(state=tk.NORMAL if can_move_down else tk.DISABLED)

    def move_selected_txt_field(self, direction):
        current_index = self.selected_txt_row_index
        if current_index == -1: return

        # CORRECTED: Use the new all_txt_mappings structure
        config_list = self.parent_gui.all_txt_mappings.get("Main TXT", [])
        total_items = len(config_list)

        if direction == "up" and current_index > 0:
            config_list[current_index], config_list[current_index - 1] = config_list[current_index - 1], config_list[current_index]
            self.parent_gui.all_txt_mappings["Main TXT"] = config_list
            self.recreate_txt_field_rows(reselect_index=current_index - 1)
        elif direction == "down" and current_index < total_items - 1:
            config_list[current_index], config_list[current_index + 1] = config_list[current_index + 1], config_list[current_index]
            self.parent_gui.all_txt_mappings["Main TXT"] = config_list
            self.recreate_txt_field_rows(reselect_index=current_index + 1)

    def add_txt_field_row(self):
        # CORRECTED: Add to the new all_txt_mappings structure
        main_txt_config = self.parent_gui.all_txt_mappings.get("Main TXT", [])
        new_field_index = len(main_txt_config) + 1
        main_txt_config.append({
            "field": f"Custom_Field_{new_field_index}",
            "column_name": f"Custom_Col_{new_field_index}",
            "skip": False
        })
        self.parent_gui.all_txt_mappings["Main TXT"] = main_txt_config
        self.recreate_txt_field_rows(reselect_index=len(main_txt_config) - 1)

    def remove_txt_field_row(self, index_to_remove):
        # CORRECTED: Remove from the new all_txt_mappings structure
        main_txt_config = self.parent_gui.all_txt_mappings.get("Main TXT", [])
        if not (0 <= index_to_remove < len(main_txt_config)):
            return
        
        config_to_remove = main_txt_config[index_to_remove]
        if messagebox.askyesno("Confirm Deletion", f"Are you sure you want to remove field '{config_to_remove['field']}'?", parent=self.master):
            del main_txt_config[index_to_remove]
            self.parent_gui.all_txt_mappings["Main TXT"] = main_txt_config
            
            new_selection = -1
            if self.selected_txt_row_index == index_to_remove:
                new_selection = -1
            elif self.selected_txt_row_index > index_to_remove:
                new_selection = self.selected_txt_row_index -1
            else:
                new_selection = self.selected_txt_row_index

            self.recreate_txt_field_rows(reselect_index=new_selection)

    def update_num_custom_buttons(self):
        try:
            num_buttons = int(self.num_buttons_entry.get())
            if not (0 <= num_buttons <= self.parent_gui.MAX_CUSTOM_BUTTONS): raise ValueError(f"Number must be between 0 and {self.parent_gui.MAX_CUSTOM_BUTTONS}")
            if self.parent_gui.num_custom_buttons != num_buttons:
                self.parent_gui.num_custom_buttons = num_buttons; current_configs = self.parent_gui.custom_button_configs
                if num_buttons < len(current_configs): self.parent_gui.custom_button_configs = current_configs[:num_buttons]
                else:
                    while len(self.parent_gui.custom_button_configs) < num_buttons:
                        idx = len(self.parent_gui.custom_button_configs) + 1
                        self.parent_gui.custom_button_configs.append({"text": f"Custom {idx}", "event_text": f"Custom {idx} Event", "txt_source_key": "None", "tab_group": "Main"}) # Default to Main
                self.recreate_custom_button_settings()
        except ValueError as e:
            messagebox.showerror("Invalid Number", f"Please enter a whole number between 0 and {self.parent_gui.MAX_CUSTOM_BUTTONS}. Error: {e}", parent=self.master)
            self.num_buttons_entry.delete(0, tk.END); self.num_buttons_entry.insert(0, str(self.parent_gui.num_custom_buttons))

    def recreate_custom_button_settings(self):
        """Clears and redraws the custom button configuration rows (text, event, source, tab group)."""
        for widget in self.custom_button_entries_frame.winfo_children():
            widget.destroy()
        self.custom_button_widgets = []

        num_buttons = self.parent_gui.num_custom_buttons
        configs = self.parent_gui.custom_button_configs
        txt_source_options = TXT_FILES_KEYS
        
                # Use the parent GUI's master list of tab groups as the single source of truth.
        all_tab_groups = sorted(self.parent_gui.custom_button_tab_groups[:])

        for i in range(num_buttons):
            config = configs[i] if i < len(configs) else {}
            initial_text = config.get("text", f"Custom {i+1}")
            initial_event = config.get("event_text", f"{initial_text} Event")
            initial_txt_source = config.get("txt_source_key", "None")
            initial_tab_group = config.get("tab_group", "Main") # **:** Default to "Main"

            style_name = f"Row{i % 2}.TFrame"
            row_frame = ttk.Frame(self.custom_button_entries_frame, style=style_name, padding=(0, 2))
            row_frame.pack(anchor='w', pady=0)

            # Configure columns for each row frame
            row_frame.grid_columnconfigure(0, weight=0)  # Button # Label
            row_frame.grid_columnconfigure(1, weight=1)  # Button Text Entry
            row_frame.grid_columnconfigure(2, weight=2)  # Event Text Entry
            row_frame.grid_columnconfigure(3, weight=0)  # Event Code
            row_frame.grid_columnconfigure(4, weight=0)  # Event Source
            row_frame.grid_columnconfigure(5, weight=0)  # Tab Group

            # Get initial values
            initial_event_code = config.get("event_code", "")

            ttk.Label(row_frame, text=f"{i+1}", width=7, style=style_name.replace("Frame","Label")).grid(row=0, column=0, padx=(5,0), sticky='w')
            text_entry = ttk.Entry(row_frame, width=20); text_entry.insert(0, initial_text); text_entry.grid(row=0, column=1, padx=5, sticky='ew'); ToolTip(text_entry, "Text displayed on the button.")
            event_entry = ttk.Entry(row_frame, width=30); event_entry.insert(0, initial_event); event_entry.grid(row=0, column=2, padx=5, sticky='ew'); ToolTip(event_entry, "Text written to the 'Event' column in the log.")

            # Event Code Combobox
            event_code_var = tk.StringVar(value=initial_event_code)
            event_code_options = [""] + sorted(list(self.parent_gui.event_codes.keys()))
            event_code_combobox = ttk.Combobox(row_frame, textvariable=event_code_var, values=event_code_options, state="readonly", width=12)
            event_code_combobox.grid(row=0, column=3, padx=5, sticky='w')
            ToolTip(event_code_combobox, "Select an event code to write to the 'Code' column.")

            # Event Source Combobox
            txt_source_var = tk.StringVar(value=initial_txt_source)
            txt_source_combobox = ttk.Combobox(row_frame, textvariable=txt_source_var, values=txt_source_options, state="readonly", width=12)
            txt_source_combobox.grid(row=0, column=4, padx=5, sticky='w')
            ToolTip(txt_source_combobox, "Select which TXT file source this button should read data from. 'None' means no TXT data will be logged by this button.")

            # Tab Group Combobox
            tab_group_var = tk.StringVar(value=initial_tab_group)
            tab_group_combobox = ttk.Combobox(row_frame, textvariable=tab_group_var, values=all_tab_groups, width=12)
            tab_group_combobox.grid(row=0, column=5, padx=5, sticky='w')
            ToolTip(tab_group_combobox, "Assign this button to a tab group. You can type a new group name or select an existing one.")

            self.custom_button_widgets.append( (text_entry, event_entry, event_code_var, txt_source_var, tab_group_var) )

    def create_monitored_folders_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Monitored Folders")

        #Tip for Monitored folders
        warning_frame = ttk.Frame(tab, padding=5)
        warning_frame.pack(fill='x', pady=(0, 10))
        ttk.Label(warning_frame, text="⚠ When changing directories, stop folder monitoring before making the change or the program will not start monitoring for any changes you make.",
                  wraplength=900, justify=tk.LEFT, foreground='red').pack(fill='x')
        
        ttk.Label(tab, text="Configure additional folders to monitor for their latest file names. The latest file name will be logged in the specified Excel/DB column.", wraplength=900, justify=tk.LEFT).pack(pady=(0, 10), anchor='w')

        threshold_frame = ttk.Frame(tab)
        threshold_frame.pack(fill='x', pady=5, padx=10)

        ttk.Label(threshold_frame, text="Active Logging Threshold (seconds):").pack(side=tk.LEFT, padx=(0, 5))
        
        threshold_spinbox = ttk.Spinbox(
            threshold_frame,
            from_=1,
            to=3600,
            increment=1,
            textvariable=self.parent_gui.active_logging_threshold_seconds,
            width=8
        )
        threshold_spinbox.pack(side=tk.LEFT)
        ToolTip(threshold_spinbox, "A file is considered 'active' if it was  within this many seconds.\nIf inactive, the cell will be left blank.")

        controls_frame = ttk.Frame(tab)
        controls_frame.pack(fill='x', pady=(0, 10), padx=10)

        add_btn = ttk.Button(controls_frame, text="Add New Folder", command=self.add_new_folder_row)
        add_btn.pack(side=tk.LEFT, padx=(0, 5))
        ToolTip(add_btn, "Add a new custom folder to monitor.")

        self.remove_folder_btn = ttk.Button(controls_frame, text="Remove Selected Folder", command=self.remove_selected_folder_row, state=tk.DISABLED)
        self.remove_folder_btn.pack(side=tk.LEFT, padx=5)
        ToolTip(self.remove_folder_btn, "Removes the currently selected folder row from the list.")

        self.folder_canvas = tk.Canvas(tab, borderwidth=0, background="#ffffff")
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=self.folder_canvas.yview)
        self.scrollable_frame = ttk.Frame(self.folder_canvas, style="Row0.TFrame")
        self.scrollable_frame.bind("<Configure>", lambda e: self.folder_canvas.configure(scrollregion=self.folder_canvas.bbox("all")))
        self.folder_canvas_window = self.folder_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.folder_canvas.configure(yscrollcommand=scrollbar.set)
        self.folder_canvas.pack(side="left", fill="both", expand=True, padx=(10,0), pady=10)
        scrollbar.pack(side="right", fill="y", padx=(0,10), pady=10)
        def _on_mousewheel(event):
            if event.num == 4: delta = -1
            elif event.num == 5: delta = 1
            elif hasattr(event, 'delta'): delta = -int(event.delta / 120)
            else: delta = 0
            self.folder_canvas.yview_scroll(delta, "units")
        self.folder_canvas.bind("<MouseWheel>", _on_mousewheel) 
        self.folder_canvas.bind("<Button-4>", _on_mousewheel)   
        self.folder_canvas.bind("<Button-5>", _on_mousewheel)  
        
        self.folder_entries = {}
        self.folder_column_entries = {}
        self.folder_db_column_entries = {}
        self.file_extension_entries = {}
        self.folder_skip_vars = {}
        self.folder_log_x_vars = {}
        self.folder_log_ext_vars = {}
        self.folder_row_widgets = {}
        self.add_folder_header(self.scrollable_frame)

    def add_folder_header(self, parent):
        # Configure the grid columns on the parent frame
        parent.columnconfigure(0, weight=2, minsize=140)  # Folder Type
        parent.columnconfigure(1, weight=4, minsize=250)  # Monitor Path
        parent.columnconfigure(2, weight=0)               # ... button
        parent.columnconfigure(3, weight=2, minsize=150)  # Excel Column
        parent.columnconfigure(4, weight=1, minsize=80)   # File Ext.
        parent.columnconfigure(5, weight=0, minsize=50)   # Skip?
        parent.columnconfigure(6, weight=0, minsize=70)   # Log 'X'?
        parent.columnconfigure(7, weight=0, minsize=70)   # Log Ext

        # 1. Create a header frame to contain the labels
        header_frame = ttk.Frame(parent, style="Header.TFrame")
        header_frame.grid(row=0, column=0, columnspan=8, sticky="ew")

        # 2. Add header labels to the new frame
        ttk.Label(header_frame, text="Folder Type", font=("Arial", 10, "bold"), style="Header.TLabel").grid(row=0, column=0, sticky='w', padx=(15, 5))
        ttk.Label(header_frame, text="Monitor Path", font=("Arial", 10, "bold"), style="Header.TLabel").grid(row=0, column=1, sticky='w', padx=5)
        ttk.Label(header_frame, text="", style="Header.TLabel").grid(row=0, column=2) # Space for browse button
        ttk.Label(header_frame, text="Excel Column", font=("Arial", 10, "bold"), style="Header.TLabel").grid(row=0, column=3, sticky='w', padx=5)
        ttk.Label(header_frame, text="File Ext.", font=("Arial", 10, "bold"), style="Header.TLabel").grid(row=0, column=4, sticky='w', padx=5)
        ttk.Label(header_frame, text="Skip?", font=("Arial", 10, "bold"), style="Header.TLabel").grid(row=0, column=5, sticky='w', padx=5)
        
        # --- FIX: Log 'X'? is now in column 6 ---
        ttk.Label(header_frame, text="Log 'X'?", font=("Arial", 10, "bold"), style="Header.TLabel").grid(row=0, column=6, sticky='w', padx=5) 
        # ----------------------------------------
        
        # --- Log Ext.? is in column 7 ---
        ttk.Label(header_frame, text="Log Ext.?", font=("Arial", 10, "bold"), style="Header.TLabel").grid(row=0, column=7, sticky='w', padx=5)
        # Apply the same column configure to the header_frame itself so its internal labels space out correctly
        for i in range(8):
             header_frame.grid_columnconfigure(i, weight=parent.grid_columnconfigure(i).get('weight', 0), minsize=parent.grid_columnconfigure(i).get('minsize', 0))

    def add_initial_folder_rows(self):
        default_folders = DEFAULT_MONITORED_FOLDERS
        
        ordered_specific_txt_folders = [
            ("Main TXT File", self.parent_gui.txt_folder_path),
            ("TXT Source 2", self.parent_gui.txt_folder_path_set2),
            ("TXT Source 3", self.parent_gui.txt_folder_path_set3),
            ("TXT Source 4", self.parent_gui.txt_folder_path_set4),
            ("TXT Source 5", self.parent_gui.txt_folder_path_set5)
        ] 
        all_folder_names = []
        processed_set = set()

        for name, path in ordered_specific_txt_folders:
            all_folder_names.append(name)
            processed_set.add(name)
            if path:
                self.parent_gui.folder_paths[name] = path  
                if name == "TXT Source 2" and not self.parent_gui.folder_columns.get(name):
                    self.parent_gui.folder_columns[name] = "TXT_Set2_File"
                    self.parent_gui.file_extensions[name] = "txt"
                if name == "TXT Source 3" and not self.parent_gui.folder_columns.get(name):
                    self.parent_gui.folder_columns[name] = "TXT_Set3_File"
                    self.parent_gui.file_extensions[name] = "txt"
                if name == "Main TXT File" and not self.parent_gui.folder_columns.get(name):
                    self.parent_gui.folder_columns[name] = "Main_TXT_File"
                    self.parent_gui.file_extensions[name] = "txt"
                if name == "TXT Source 4" and not self.parent_gui.folder_columns.get(name):
                    self.parent_gui.folder_columns[name] = "TXT_Set4_File"
                    self.parent_gui.file_extensions[name] = "txt"
                if name == "TXT Source 5" and not self.parent_gui.folder_columns.get(name):
                    self.parent_gui.folder_columns[name] = "TXT_Set5_File"
                    self.parent_gui.file_extensions[name] = "txt"

        for name in default_folders:
            if name not in processed_set:
                all_folder_names.append(name)
                processed_set.add(name)
        
        for name in self.parent_gui.folder_paths:
            if name not in processed_set:
                all_folder_names.append(name)
                processed_set.add(name)

        for folder_name in all_folder_names:
            folder_path_to_use = self.parent_gui.folder_paths.get(folder_name, "")
            if folder_name == "Main TXT File": folder_path_to_use = self.parent_gui.txt_folder_path or ""
            elif folder_name == "TXT Source 2": folder_path_to_use = self.parent_gui.txt_folder_path_set2 or ""
            elif folder_name == "TXT Source 3": folder_path_to_use = self.parent_gui.txt_folder_path_set3 or ""
            elif folder_name == "TXT Source 4": folder_path_to_use = self.parent_gui.txt_folder_path_set4 or ""
            elif folder_name == "TXT Source 5": folder_path_to_use = self.parent_gui.txt_folder_path_set5 or ""

            column_name_to_use = self.parent_gui.folder_columns.get(folder_name, folder_name)

            extension_to_use = self.parent_gui.file_extensions.get(folder_name, "")

            if folder_name in ["Main TXT File", "TXT Source 2", "TXT Source 3", "TXT Source 4", "TXT Source 5"]:
                if not column_name_to_use or column_name_to_use == folder_name:
                    column_name_to_use = folder_name.replace(" ", "_")
                if not extension_to_use:
                    extension_to_use = "txt"

            self.add_folder_row(folder_name=folder_name, folder_path=folder_path_to_use, column_name=column_name_to_use, extension=extension_to_use, skip=self.parent_gui.folder_skips.get(folder_name, False))
        self.master.after_idle(self.update_scroll_region)

    def _select_folder_row(self, folder_name):
        """Highlights all widgets in a selected folder row."""
        # Deselect the previously selected row by resetting widget styles
        if hasattr(self, 'selected_folder_name') and self.selected_folder_name:
            prev_widgets = self.folder_row_widgets.get(self.selected_folder_name)
            if prev_widgets:
                for widget in prev_widgets:
                    try:
                        # Resets to the default style for each widget type
                        widget.configure(style=f"T{type(widget).__name__}")
                    except tk.TclError:
                        pass # Widget may have been destroyed

        # Select the new row by applying the "Selected" style
        self.selected_folder_name = folder_name
        current_widgets = self.folder_row_widgets.get(folder_name)
        if current_widgets:
            for widget in current_widgets:
                try:
                    # Applies the 'Selected' style variant, e.g., "Selected.TEntry"
                    widget.configure(style=f"Selected.T{type(widget).__name__}")
                except tk.TclError:
                    pass

        self.remove_folder_btn.config(state=tk.NORMAL)

    def add_new_folder_row(self):
        """Asks for a new folder type name and adds a new row to the UI."""
        new_name = simpledialog.askstring("New Folder Type", "Enter a unique name for the new folder type (e.g., 'WROV Data'):", parent=self.master)

        if not new_name or not new_name.strip():
            return # User cancelled

        new_name = new_name.strip()
        if new_name in self.folder_entries:
            messagebox.showerror("Duplicate Name", f"A folder type named '{new_name}' already exists.", parent=self.master)
            return
        
        # Add a new, empty row for the user to configure
        self.add_folder_row(folder_name=new_name)
        self.master.after_idle(self.update_scroll_region)

    def remove_selected_folder_row(self):
        """Removes the currently selected folder row and its associated data."""
        if not hasattr(self, 'selected_folder_name') or not self.selected_folder_name:
            messagebox.showinfo("No Selection", "Please select a folder row to remove.", parent=self.master)
            return

        folder_to_remove = self.selected_folder_name
        
        if messagebox.askyesno("Confirm Deletion", f"Are you sure you want to remove the '{folder_to_remove}' folder configuration?", parent=self.master):
            # Pop the list of widgets from the dictionary
            widgets_to_destroy = self.folder_row_widgets.pop(folder_to_remove, None)
            
            # If the list exists, iterate through it and destroy each widget
            if widgets_to_destroy:
                for widget in widgets_to_destroy:
                    if widget and widget.winfo_exists():
                        widget.destroy()

            # Remove the data entries from the other dictionaries
            self.folder_entries.pop(folder_to_remove, None)
            self.folder_column_entries.pop(folder_to_remove, None)
            self.folder_db_column_entries.pop(folder_to_remove, None)
            self.file_extension_entries.pop(folder_to_remove, None)
            self.folder_skip_vars.pop(folder_to_remove, None)

            # Reset the selection state
            self.selected_folder_name = None
            self.remove_folder_btn.config(state=tk.DISABLED)
            self.parent_gui.update_status(f"Removed '{folder_to_remove}' configuration.")

    def add_folder_row(self, folder_name="", folder_path="", column_name="", extension="", skip=False):
        row_index = len(self.folder_row_widgets) + 1
        parent = self.scrollable_frame # The single grid container

        # --- Create Widgets ---
        label = ttk.Label(parent, text=f"{folder_name}:", anchor='w')
        entry = ttk.Entry(parent)
        entry.insert(0, folder_path)
        
        def select_folder(e=entry, name=folder_name):
            current_path = e.get()
            initial = current_path if os.path.isdir(current_path) else os.getcwd()
            folder = filedialog.askdirectory(parent=self.master, initialdir=initial, title=f"Select Folder for {name}")
            if folder:
                e.delete(0, tk.END)
                e.insert(0, folder)

        button = ttk.Button(parent, text="...", width=3, command=select_folder)
        column_entry = ttk.Entry(parent)
        column_entry.insert(0, column_name if column_name else folder_name)
        extension_entry = ttk.Entry(parent, width=10)
        extension_entry.insert(0, extension)
        skip_var = tk.BooleanVar(value=skip)
        skip_checkbox = ttk.Checkbutton(parent, variable=skip_var)

        log_x_var = tk.BooleanVar(value=self.parent_gui.folder_log_x_instead.get(folder_name, False))
        log_x_checkbox = ttk.Checkbutton(parent, variable=log_x_var)

        if folder_name not in self.parent_gui.folder_log_ext_vars:
            self.parent_gui.folder_log_ext_vars[folder_name] = tk.BooleanVar(value=False)
            
        log_ext_var = self.parent_gui.folder_log_ext_vars[folder_name] # Retrieve the shared BooleanVar
        log_ext_checkbox = ttk.Checkbutton(parent, variable=log_ext_var)

        # --- Place Widgets on the Shared Grid ---
        label.grid(row=row_index, column=0, padx=5, pady=2, sticky="ew")
        entry.grid(row=row_index, column=1, padx=5, pady=2, sticky="ew")
        button.grid(row=row_index, column=2, padx=(0,5), pady=2, sticky='w')
        column_entry.grid(row=row_index, column=3, padx=5, pady=2, sticky="ew")
        extension_entry.grid(row=row_index, column=4, padx=5, pady=2, sticky="ew") 
        skip_checkbox.grid(row=row_index, column=5, padx=(15, 5), pady=2, sticky='w')
        log_x_checkbox.grid(row=row_index, column=6, padx=(15, 5), pady=2, sticky='w') 
        log_ext_checkbox.grid(row=row_index, column=7, padx=(15, 5), pady=2, sticky='w')

        # --- Selection and Tooltip Logic ---
        widgets_in_row = [label, entry, button, column_entry, extension_entry, skip_checkbox]
        click_handler = lambda e, name=folder_name: self._select_folder_row(name)
        for widget in widgets_in_row:
            widget.bind("<Button-1>", click_handler)
        ToolTip(entry, f"Enter the full path to the '{folder_name}' data folder.")
        ToolTip(button, "Browse for the folder.")
        ToolTip(column_entry, f"Enter the Excel/DB column name for the latest '{folder_name}' filename.")
        ToolTip(extension_entry, "Optional: Monitor only files with this extension (e.g., 'svp').")
        ToolTip(skip_checkbox, f"Check to disable monitoring for the '{folder_name}' folder.")
        ToolTip(log_x_checkbox, "If a file is logging, insert 'X' into the Excel column.\nThe database will still receive the actual filename.")
        ToolTip(log_ext_checkbox, "If checked, the full filename, including the extension, will be logged to Excel (to prevent scientific notation formatting).")
    
        # Store references for selection, saving, and removal
        self.folder_entries[folder_name] = entry
        self.folder_column_entries[folder_name] = column_entry
        self.file_extension_entries[folder_name] = extension_entry
        self.folder_skip_vars[folder_name] = skip_var
        self.folder_log_x_vars[folder_name] = log_x_var
        # Store all widgets in the row for highlighting
        self.folder_row_widgets[folder_name] = widgets_in_row
        self.folder_log_ext_vars[folder_name] = log_ext_var

        widgets_in_row = [label, entry, button, column_entry, extension_entry, skip_checkbox, log_x_checkbox, log_ext_checkbox]
        self.folder_row_widgets[folder_name] = widgets_in_row

    def update_scroll_region(self):
        self.scrollable_frame.update_idletasks()
        self.folder_canvas.configure(scrollregion=self.folder_canvas.bbox("all"))

    def save_settings_to_parent_vars(self):
        """
        A lightweight helper that only updates the parent GUI's in-memory
        column configurations from the current state of the settings UI. 
        This ensures the SQL generation uses the most up-to-date names.
        """
        new_txt_field_configs = []
        for i, row_info in enumerate(self.txt_field_row_widgets):
            field_name = ""
            # For non-fixed fields, read from the entry widget
            if row_info["field_entry_widget"]:
                field_name = row_info["field_entry_widget"].get().strip()
            # For fixed fields, get the name from the original config
            else:
                if i < len(self.parent_gui.txt_field_columns_config):
                    field_name = self.parent_gui.txt_field_columns_config[i]["field"]
            
            column_name = row_info["column_entry"].get().strip()
            skip_value = row_info["skip_var"].get()

            if not field_name and not (field_name in DEFAULT_DATA_FIELDS):
                field_name = f"Custom_Field_{i+1}"

            new_txt_field_configs.append({
                "field": field_name,
                "column_name": column_name if column_name else field_name,
                "skip": skip_value
            })
        self.parent_gui.txt_field_columns_config = new_txt_field_configs

    def _sanitize_for_sql(self, name):
        """Removes symbols and converts spaces to underscores for a valid SQL column name."""
        if not name:
            return ""
        # Remove any character that is not a letter, number, or space
        s = re.sub(r'[^\w\s]', '', name)
        # Replace one or more spaces with a single underscore
        s = re.sub(r'\s+', '_', s.strip())
        return s    


    def create_auto_events_tab(self):
        """
        Creates the tab for configuring automatic timed events with an improved layout,
        including a configurable source for the Hourly KP Log.
        """
        tab = ttk.Frame(self.notebook, padding=20)
        self.notebook.add(tab, text="Programmed Events")

        # Use a main grid to structure the tab content
        tab.columnconfigure(0, weight=1)
        
        # 1. Midnight 'New Day' Event Configuration
        # CORRECTED: Define as instance attribute
        self.new_day_frame = ttk.LabelFrame(tab, text="Midnight 'New Day' Event", padding=15)
        self.new_day_frame.grid(row=0, column=0, sticky='ew', pady=(0, 15))
        self.new_day_frame.columnconfigure(1, weight=1) # Allow second column to expand

        # Row 0: Enable Checkbox
        new_day_check = ttk.Checkbutton(self.new_day_frame, text="Enable this automatic event", 
                                         variable=self.parent_gui.new_day_event_enabled_var,
                                         style="Large.TCheckbutton")
        new_day_check.grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 10))
        ToolTip(new_day_check, "If checked, an event will be logged automatically at midnight.")

        # Rows 1-2: Color Pickers
        ttk.Label(self.new_day_frame, text="Excel Row Colors:").grid(row=1, column=0, sticky='w', padx=5, pady=(2, 0))
        self._create_color_picker_widgets(self.new_day_frame, 1, "New Day")


        # 2. Hourly KP Log Event Configuration
        # CORRECTED: Define as instance attribute
        self.hourly_frame = ttk.LabelFrame(tab, text="Hourly KP Log Event", padding=15)
        self.hourly_frame.grid(row=1, column=0, sticky='ew', pady=5)
        self.hourly_frame.columnconfigure(1, weight=1)

        # Row 0: Enable Checkbox
        hourly_check = ttk.Checkbutton(self.hourly_frame, text="Enable this automatic event",
                                     variable=self.parent_gui.hourly_event_enabled_var,
                                     style="Large.TCheckbutton")
        hourly_check.grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 10))
        ToolTip(hourly_check, "If checked, the current KP will be logged automatically every hour.")
        
        # --- ROW 1: Source Selection ---
        ttk.Label(self.hourly_frame, text="KP Data Source:").grid(row=1, column=0, sticky='w', padx=5, pady=(10, 5))

        # Build map for translation (Key -> Display Name)
        aliases = self.parent_gui.txt_source_aliases
        # Access global constant directly
        internal_keys = TXT_FILES_KEYS
        
        # Create map and lists for the combobox
        key_to_display_map = {k: aliases.get(k, k) for k in internal_keys if k != "None"}
        display_names = list(key_to_display_map.values()) 

        # Retrieve the saved internal key from the parent GUI instance
        initial_key = self.parent_gui.hourly_log_txt_source_key.get()
        # Find the corresponding display name, defaulting if the key is somehow missing
        initial_display = key_to_display_map.get(initial_key, "Main TXT")
        
        # Use a local StringVar to hold the *display name* for the ComboBox
        hourly_source_display_var = tk.StringVar(value=initial_display)

        hourly_source_combobox = ttk.Combobox(self.hourly_frame, textvariable=hourly_source_display_var,
                                              values=display_names, state="readonly", width=15)
        
        hourly_source_combobox.grid(row=1, column=1, sticky='w', padx=5, pady=(10, 5))
        ToolTip(hourly_source_combobox, "Select which data source to use for the hourly KP and line check.")
        
        # --- Store the combobox and the reverse map for saving later ---
        self.hourly_source_combobox = hourly_source_combobox
        # IMPORTANT: Store the map (Display Name -> Internal Key) for saving
        self.hourly_source_map = {v: k for k, v in key_to_display_map.items()}

        # Rows 2-3: Color Pickers
        #ttk.Label(self.hourly_frame, text="Excel Row Colors:").grid(row=2, column=0, sticky='w', padx=5, pady=(2, 0))
        #self._create_color_picker_widgets(self.hourly_frame, 2, "Hourly KP Log")


        # 3. Log off Distance/Speed Calculation
        # CORRECTED: Define as instance attribute
        self.logoff_frame = ttk.LabelFrame(tab, text="Log off Distance/Speed Calculation", padding=15)
        self.logoff_frame.grid(row=2, column=0, sticky='ew', pady=5)
        self.logoff_frame.columnconfigure(1, weight=1)
        
        # Row 0: Enable Checkbox
        logoff_check = ttk.Checkbutton(self.logoff_frame, text="Calculate distance & speed on Log off",
                                     variable=self.parent_gui.calculate_logoff_values,
                                     style="Large.TCheckbutton")
        logoff_check.grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 10))
        ToolTip(logoff_check, "If checked, the Log off button will calculate and display distance and speed based on the last Log on event's KP.")


    def _create_color_picker_widgets(self, parent_frame, grid_row, event_name):
        """
        Helper to create and place the color picker widgets for both background and font colors.
        """
        initial_bg_color, initial_font_color = self.parent_gui.button_colors.get(event_name, (None, None))
        
        bg_color_var = tk.StringVar(value=initial_bg_color if initial_bg_color else "")
        font_color_var = tk.StringVar(value=initial_font_color if initial_font_color else "")
        
        # Frame for Background Color picker
        bg_color_frame = ttk.Frame(parent_frame)
        bg_color_frame.grid(row=grid_row, column=1, sticky='w', padx=5, pady=(2, 0))

        bg_display_label = tk.Label(bg_color_frame, width=4, relief="solid", borderwidth=1,
                                    background=bg_color_var.get() if bg_color_var.get() else 'SystemButtonFace')
        bg_display_label.pack(side="left", padx=(0, 5))

        clear_bg_btn = ttk.Button(bg_color_frame, text="X", width=2, style="Toolbutton",
                                command=lambda: self.parent_gui._set_color_on_widget(bg_color_var, bg_display_label, None, self.master))
        clear_bg_btn.pack(side="left", padx=1)
        
        choose_bg_btn = ttk.Button(bg_color_frame, text="...", width=3, style="Toolbutton",
                                command=lambda: self.parent_gui._choose_color_dialog(bg_color_var, bg_display_label, self.master, f"{event_name} Background"))
        choose_bg_btn.pack(side="left", padx=1)

        # Frame for Font Color picker
        font_color_frame = ttk.Frame(parent_frame)
        font_color_frame.grid(row=grid_row + 1, column=1, sticky='w', padx=5, pady=(0, 2))

        font_display_label = tk.Label(font_color_frame, width=4, relief="solid", borderwidth=1,
                                    background=font_color_var.get() if font_color_var.get() else 'SystemButtonFace')
        font_display_label.pack(side="left", padx=(0, 5))

        clear_font_btn = ttk.Button(font_color_frame, text="X", width=2, style="Toolbutton",
                                    command=lambda: self.parent_gui._set_color_on_widget(font_color_var, font_display_label, None, self.master))
        clear_font_btn.pack(side="left", padx=1)
        
        choose_font_btn = ttk.Button(font_color_frame, text="...", width=3, style="Toolbutton",
                                    command=lambda: self.parent_gui._choose_color_dialog(font_color_var, font_display_label, self.master, f"{event_name} Font"))
        choose_font_btn.pack(side="left", padx=1)

    def _load_programmed_events_ui_state(self):
        """
        Synchronizes the combobox display variables and color labels for Programmed Events
        after all widgets have been created.
        """
        # --- Hourly KP Log Source Synchronization ---
        if hasattr(self, 'hourly_source_combobox') and hasattr(self, 'hourly_source_map'):
            # 1. Get the internal key stored in the parent GUI (loaded from JSON)
            initial_key = self.parent_gui.hourly_log_txt_source_key.get()
            
            # 2. Re-map to find the Display Name (Alias) from the Internal Key
            display_to_internal_map = self.hourly_source_map
            
            initial_display_name = None
            for display_name, internal_key in display_to_internal_map.items():
                if internal_key == initial_key:
                    initial_display_name = display_name
                    break
            
            # 3. Update the combobox's display
            if initial_display_name:
                self.hourly_source_combobox.set(initial_display_name)

        # --- Color Label Synchronization (Requires accessing children of the instance attributes) ---
        if hasattr(self, 'new_day_frame') and hasattr(self, 'hourly_frame'):
            # New Day Background Color Label: Grid row 1, column 1, is the frame holding the label. Label is the first child [0].
            new_day_bg_frame = self.new_day_frame.grid_slaves(row=1, column=1)[0]
            new_day_bg_label = new_day_bg_frame.winfo_children()[0]
            
            # Hourly Log Background Color Label: Grid row 2, column 1, is the frame holding the label. Label is the first child [0].
            #hourly_bg_frame = self.hourly_frame.grid_slaves(row=2, column=1)[0]
            #hourly_bg_label = hourly_bg_frame.winfo_children()[0]

            # Trigger color update using the stored StringVar value
            self.parent_gui._set_color_on_widget(
                self.new_day_bg_color_var, 
                new_day_bg_label, 
                self.new_day_bg_color_var.get(), 
                self.master
            )
            #self.parent_gui._set_color_on_widget(
                #self.hourly_bg_color_var, 
                #hourly_bg_label, 
                #self.hourly_bg_color_var.get(), 
                #self.master
            #)
            # You may need similar logic for font colors if they are configured to show a background color
            # font_frame = self.new_day_frame.grid_slaves(row=2, column=1)[0]
            # font_label = font_frame.winfo_children()[0]
            # self.parent_gui._set_color_on_widget(self.new_day_font_color_var, font_label, self.new_day_font_color_var.get(), self.master)

    def create_projects_tab(self):
        """Creates the Settings Projects tab for managing project JSON files."""
        projects_tab = ttk.Frame(self.notebook)
        self.notebook.add(projects_tab, text="Projects")
        
        # Initialize variables for this tab
        self.current_project_path = tk.StringVar()
        self.projects_tree = None
        
        # Set default project path: use parent's last used project if available, otherwise blank
        default_project = getattr(self.parent_gui, "current_project_path", None) or ""
        self.current_project_path.set(default_project)
        # Ensure parent keeps track of last used project path as well
        self.parent_gui.current_project_path = self.current_project_path.get()
        
        # Main container with padding
        main_container = ttk.Frame(projects_tab)
        main_container.pack(fill="both", expand=True, padx=10, pady=10)

        # --- Project Path Section ---
        path_frame = ttk.LabelFrame(main_container, text="Current Project", padding="10")
        path_frame.pack(fill="x", pady=(0, 10))
        path_frame.columnconfigure(1, weight=1)

        # Current project path display
        ttk.Label(path_frame, text="Project Path:").grid(row=0, column=0, sticky="w", padx=(0, 10))
        self.project_path_entry = ttk.Entry(path_frame, textvariable=self.current_project_path,
                                          state="readonly", width=60)
        self.project_path_entry.grid(row=0, column=1, sticky="ew", padx=(0, 10))

        # Browse button
        browse_btn = ttk.Button(path_frame, text="Browse...", command=self.browse_project)
        browse_btn.grid(row=0, column=2, sticky="e")

        # --- Control Buttons Section ---
        control_frame = ttk.Frame(main_container)
        control_frame.pack(fill="x", pady=(0, 10))

        # Load Project button
        load_btn = ttk.Button(control_frame, text="Load Project", command=self.load_project)
        load_btn.pack(side="left", padx=(0, 5))

        # Save Project button  
        save_btn = ttk.Button(control_frame, text="Save Project", command=self.save_project)
        save_btn.pack(side="left", padx=(0, 5))

        # Save As button
        save_as_btn = ttk.Button(control_frame, text="Save As...", command=self.save_project_as)
        save_as_btn.pack(side="left", padx=(0, 5))

        # Restore Blank Project button (uses blank template)
        style = ttk.Style()
        style.configure("Danger.TButton", foreground="white", background="#C00000")
        style.map("Danger.TButton", background=[("active", "#A00000")], foreground=[("active", "white")])

        restore_btn = ttk.Button(control_frame, text="Restore Blank Project", command=self.load_blank_project, style="Danger.TButton")
        restore_btn.pack(side="left")
        ToolTip(restore_btn, "Restore the in-memory settings from the blank project template.")

        # --- JSON Structure Viewer ---
        viewer_frame = ttk.LabelFrame(main_container, text="Project Structure", padding="5")
        viewer_frame.pack(fill="both", expand=True)
        viewer_frame.columnconfigure(0, weight=1)
        viewer_frame.rowconfigure(0, weight=1)
        
        # Create treeview with scrollbars
        tree_container = ttk.Frame(viewer_frame)
        tree_container.grid(row=0, column=0, sticky="nsew")
        tree_container.columnconfigure(0, weight=1)
        tree_container.rowconfigure(0, weight=1)
        
        # Treeview for JSON structure
        self.projects_tree = ttk.Treeview(tree_container, columns=("value",), show="tree headings")
        self.projects_tree.heading("#0", text="Setting")
        self.projects_tree.heading("value", text="Value")
        self.projects_tree.column("#0", width=300, minwidth=200)
        self.projects_tree.column("value", width=400, minwidth=200)
        
        # Scrollbars for treeview
        v_scrollbar = ttk.Scrollbar(tree_container, orient="vertical", command=self.projects_tree.yview)
        h_scrollbar = ttk.Scrollbar(tree_container, orient="horizontal", command=self.projects_tree.xview)
        self.projects_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Grid layout for treeview and scrollbars
        self.projects_tree.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        
        # Bind treeview selection event
        self.projects_tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        
        # Load the current project structure
        self.refresh_project_structure()

    def load_blank_project(self):
        """Loads the blank project template into the application without selecting a file."""
        template_path = os.path.join(os.getcwd(), PROJECT_TEMPLATE_FILE)
        if not os.path.exists(template_path):
            messagebox.showerror(
                "Template Missing",
                f"Blank project template not found at:\n{template_path}\n\nPlease create it or reinstall.",
                parent=self.master,
            )
            return

        try:
            with open(template_path, 'r', encoding='utf-8') as f:
                project_data = json.load(f)
            # Apply to parent GUI
            self._apply_project_to_gui(project_data)
            # Clear active project path; this is a template-in-memory state
            self.current_project_path.set("")
            self.parent_gui.settings_file = None
            # Reload UI tabs and structure tree
            self._reload_all_tabs()
            self.refresh_project_structure()
            messagebox.showinfo("Blank Project Restored", "Blank project template restored in memory. Use 'Save As...' to create a new project file.", parent=self.master)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load blank project template:\n{e}", parent=self.master)
    
    def browse_project(self):
        """Opens a file dialog to browse for project JSON files."""
        # Start in settings directory
        initial_dir = os.path.join(os.getcwd(), "settings")
        if not os.path.exists(initial_dir):
            initial_dir = os.getcwd()
            
        file_path = filedialog.askopenfilename(
            parent=self.master,
            title="Select Project",
            initialdir=initial_dir,
            filetypes=[
                ("JSON files", "*.json"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self.current_project_path.set(file_path)
            # Persist selection in parent so it shows next time
            self.parent_gui.current_project_path = file_path
            self.refresh_project_structure()
    
    def load_project(self):
        """Loads the selected project into the current settings."""
        project_path = self.current_project_path.get()
        
        if not project_path or not os.path.exists(project_path):
            messagebox.showerror("Error", "Project file not found!", parent=self.master)
            return
            
        try:
            # Read the project JSON
            with open(project_path, 'r', encoding='utf-8') as f:
                project_data = json.load(f)
            
            # Apply the loaded settings to the parent GUI
            self._apply_project_to_gui(project_data)
            
            # Reload all tabs to reflect the new settings
            self._reload_all_tabs()
            
            # Refresh the structure viewer
            self.refresh_project_structure()
            # Remember last used project (in-memory and persisted) and make it active for saving
            self.parent_gui.set_active_project(project_path)
            
            messagebox.showinfo("Success", f"Project loaded successfully from:\n{project_path}", parent=self.master)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load project:\n{str(e)}", parent=self.master)
    
    def save_project(self):
        """Saves current settings to the current project path."""
        project_path = self.current_project_path.get()
        
        if not project_path:
            self.save_project_as()
            return
            
        try:
            # Collect current settings from the GUI
            current_settings = self._collect_current_settings()
            
            # Save to file
            with open(project_path, 'w', encoding='utf-8') as f:
                json.dump(current_settings, f, indent=4, ensure_ascii=False)
            
            # Refresh the structure viewer
            self.refresh_project_structure()
            # Make this project the active one going forward
            self.parent_gui.set_active_project(project_path)
            
            messagebox.showinfo("Success", f"Project saved to:\n{project_path}", parent=self.master)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save project:\n{str(e)}", parent=self.master)
            print(f"Failed to save project:\n{str(e)}")
    
    def save_project_as(self):
        """Saves current settings to a new project file."""
        # Show rename dialog with proper parent window to ensure it appears on top
        result = simpledialog.askstring(
            "Save Project As", 
            "Enter a name for the project:",
            parent=self.master,
            initialvalue="new_settings"
        )
        
        if not result:
            return
            
        # Ensure .json extension
        if not result.endswith('.json'):
            result += '.json'
            
        # Construct full path in settings directory
        settings_dir = os.path.join(os.getcwd(), "settings")
        if not os.path.exists(settings_dir):
            os.makedirs(settings_dir)
            
        new_path = os.path.join(settings_dir, result)
        
        # Check if file already exists
        if os.path.exists(new_path):
            if not messagebox.askyesno("File Exists", 
                                     f"File '{result}' already exists. Overwrite?",
                                     parent=self.master):
                return
        
        # Update current path and save
        self.current_project_path.set(new_path)
        # Persist selection in parent and make it the active project
        self.parent_gui.set_active_project(new_path)
        self.save_project()
    
    def refresh_project_structure(self):
        """Refreshes the JSON structure tree view."""
        if not self.projects_tree:
            return
            
        # Clear existing items
        for item in self.projects_tree.get_children():
            self.projects_tree.delete(item)
            
        project_path = self.current_project_path.get()
        
        if not project_path or not os.path.exists(project_path):
            # Show empty message
            self.projects_tree.insert("", "end", text="No project loaded", values=("",))
            return
            
        try:
            # Load and display JSON structure
            with open(project_path, 'r', encoding='utf-8') as f:
                project_data = json.load(f)
            
            self._populate_tree("", project_data)
            
        except Exception as e:
            self.projects_tree.insert("", "end", text=f"Error loading project: {str(e)}", values=("",))
    
    def _populate_tree(self, parent, data, key_prefix=""):
        """Recursively populates the tree with JSON data."""
        if isinstance(data, dict):
            for key, value in data.items():
                full_key = f"{key_prefix}.{key}" if key_prefix else key
                
                if isinstance(value, (dict, list)):
                    # Create parent node for nested structures
                    item_id = self.projects_tree.insert(parent, "end", text=key, 
                                                       values=(f"{type(value).__name__} ({len(value)} items)",))
                    self._populate_tree(item_id, value, full_key)
                else:
                    # Leaf node with value
                    display_value = str(value)
                    if len(display_value) > 100:
                        display_value = display_value[:100] + "..."
                    self.projects_tree.insert(parent, "end", text=key, values=(display_value,))
                    
        elif isinstance(data, list):
            for i, item in enumerate(data):
                full_key = f"{key_prefix}[{i}]" if key_prefix else f"[{i}]"
                
                if isinstance(item, (dict, list)):
                    item_id = self.projects_tree.insert(parent, "end", text=f"[{i}]", 
                                                       values=(f"{type(item).__name__} ({len(item)} items)",))
                    self._populate_tree(item_id, item, full_key)
                else:
                    display_value = str(item)
                    if len(display_value) > 100:
                        display_value = display_value[:100] + "..."
                    self.projects_tree.insert(parent, "end", text=f"[{i}]", values=(display_value,))
    
    def on_tree_select(self, event):
        """Handles tree selection events."""
        selection = self.projects_tree.selection()
        if selection:
            item = selection[0]
            # Could implement additional functionality here, like showing detailed value in a tooltip
            pass
    
    def _apply_project_to_gui(self, project_data):
        """Applies loaded project data to the parent GUI variables."""
        # This mirrors the load_settings logic but from the project data instead of file
        parent = self.parent_gui
        
        # File paths
        if "log_file_path" in project_data:
            parent.log_file_path = project_data["log_file_path"]
        if "sqlite_db_path" in project_data:
            parent.sqlite_db_path = project_data["sqlite_db_path"]
        if "txt_folder_path" in project_data:
            parent.txt_folder_path = project_data["txt_folder_path"]
        if "txt_folder_path_set2" in project_data:
            parent.txt_folder_path_set2 = project_data["txt_folder_path_set2"]
        if "txt_folder_path_set3" in project_data:
            parent.txt_folder_path_set3 = project_data["txt_folder_path_set3"]
        if "txt_folder_path_set4" in project_data:
            parent.txt_folder_path_set4 = project_data["txt_folder_path_set4"]
        if "txt_folder_path_set5" in project_data:
            parent.txt_folder_path_set5 = project_data["txt_folder_path_set5"]
        
        # TXT mappings
        if "all_txt_mappings" in project_data:
            parent.all_txt_mappings = project_data["all_txt_mappings"]
        elif "txt_mapping_config" in project_data:
            # Handle legacy format
            parent.all_txt_mappings["Main TXT"] = project_data["txt_mapping_config"]
        
        # Generated fields
        if "generated_fields_config" in project_data:
            parent.generated_fields_config = project_data["generated_fields_config"]
        
        # Static fields
        if "static_field_configs" in project_data:
            parent.static_field_configs = project_data["static_field_configs"]
        
        # Monitored folders
        if "folder_paths" in project_data:
            parent.folder_paths = project_data["folder_paths"]
        if "folder_columns" in project_data:
            parent.folder_columns = project_data["folder_columns"]
        if "file_extensions" in project_data:
            parent.file_extensions = project_data["file_extensions"]
        if "folder_skips" in project_data:
            parent.folder_skips = project_data["folder_skips"]
        if "folder_log_x_instead" in project_data:
            parent.folder_log_x_instead = project_data["folder_log_x_instead"]
        if "folder_log_ext_vars" in project_data:
            loaded_log_exts = project_data["folder_log_ext_vars"]
            parent.folder_log_ext_vars = {k: tk.BooleanVar(value=v) for k, v in loaded_log_exts.items()}
        
        # Custom buttons
        if "num_custom_buttons" in project_data:
            parent.num_custom_buttons = project_data["num_custom_buttons"]
        if "custom_button_configs" in project_data:
            parent.custom_button_configs = project_data["custom_button_configs"]
        if "custom_button_tab_groups" in project_data:
            parent.custom_button_tab_groups = project_data["custom_button_tab_groups"]
        
        # Colors and UI settings
        if "button_colors" in project_data:
            parent.button_colors = project_data["button_colors"]
        if "always_on_top" in project_data:
            parent.always_on_top_var.set(project_data["always_on_top"])
        
        # Timing and behavior settings
        if "time_offset_hours" in project_data:
            parent.time_offset_hours.set(project_data["time_offset_hours"])
        if "active_logging_threshold_seconds" in project_data:
            parent.active_logging_threshold_seconds.set(project_data["active_logging_threshold_seconds"])
        
        # Auto events
        if "new_day_event_enabled" in project_data:
            parent.new_day_event_enabled_var.set(project_data["new_day_event_enabled"])
        if "hourly_event_enabled" in project_data:
            parent.hourly_event_enabled_var.set(project_data["hourly_event_enabled"])
        if "hourly_log_txt_source_key" in project_data:
            parent.hourly_log_txt_source_key.set(project_data["hourly_log_txt_source_key"])
        
        # Main button configs
        if "main_button_configs" in project_data:
            parent.main_button_configs = project_data["main_button_configs"]
        
        # TXT source aliases
        if "txt_source_aliases" in project_data:
            parent.txt_source_aliases = project_data["txt_source_aliases"]
        # Event codes (embedded in project)
        if "event_codes" in project_data and isinstance(project_data["event_codes"], dict):
            parent.event_codes = project_data["event_codes"]
        
        # Other settings
        if "calculate_logoff_values" in project_data:
            parent.calculate_logoff_values.set(project_data["calculate_logoff_values"])
        if "auto_sync_enabled" in project_data:
            parent.auto_sync_enabled_var.set(project_data["auto_sync_enabled"])
        if "auto_sync_interval_min" in project_data:
            parent.auto_sync_interval_min_var.set(project_data["auto_sync_interval_min"])
    
    def _collect_current_settings(self):
        """Collects current settings from GUI to create a project."""
        # This mirrors the save_settings logic to gather all current settings
        parent = self.parent_gui
        
        settings = {
            # File paths
            "log_file_path": parent.log_file_path or "",
            "sqlite_db_path": parent.sqlite_db_path or "",
            "txt_folder_path": parent.txt_folder_path or "",
            "txt_folder_path_set2": parent.txt_folder_path_set2 or "",
            "txt_folder_path_set3": parent.txt_folder_path_set3 or "",
            "txt_folder_path_set4": parent.txt_folder_path_set4 or "",
            "txt_folder_path_set5": parent.txt_folder_path_set5 or "",
            
            # TXT mappings
            "all_txt_mappings": parent.all_txt_mappings,
            
            # Generated fields
            "generated_fields_config": parent.generated_fields_config,
            
            # Static fields
            "static_field_configs": parent.static_field_configs,
            
            # Monitored folders
            "folder_paths": parent.folder_paths,
            "folder_columns": parent.folder_columns,
            "file_extensions": parent.file_extensions,
            "folder_skips": parent.folder_skips,
            "folder_log_x_instead": parent.folder_log_x_instead,
            "folder_log_ext_vars": {k: v.get() for k, v in parent.folder_log_ext_vars.items()},
            
            # Custom buttons
            "num_custom_buttons": parent.num_custom_buttons,
            "custom_button_configs": parent.custom_button_configs,
            "custom_button_tab_groups": parent.custom_button_tab_groups,
            
            # Colors and UI
            "button_colors": parent.button_colors,
            "always_on_top": parent.always_on_top_var.get(),
            
            # Timing settings
            "time_offset_hours": parent.time_offset_hours.get(),
            "active_logging_threshold_seconds": parent.active_logging_threshold_seconds.get(),
            
            # Auto events
            "new_day_event_enabled": parent.new_day_event_enabled_var.get(),
            "hourly_event_enabled": parent.hourly_event_enabled_var.get(),
            "hourly_log_txt_source_key": parent.hourly_log_txt_source_key.get(),
            
            # Main button configs
            "main_button_configs": parent.main_button_configs,
            
            # TXT source aliases
            "txt_source_aliases": parent.txt_source_aliases,
            # Event codes
            "event_codes": parent.event_codes,
            
            # Other settings
            "calculate_logoff_values": parent.calculate_logoff_values.get(),
            "auto_sync_enabled": parent.auto_sync_enabled_var.get(),
            "auto_sync_interval_min": parent.auto_sync_interval_min_var.get(),
        }
        
        return settings
    
    def _reload_all_tabs(self):
        """Reloads all UI elements to reflect the newly loaded settings."""
        # Repopulate all tabs from the parent GUI's in-memory values
        # so the UI truly reflects the newly loaded project.
        self.load_settings()


    # --- Settings Save/Load Logic ---
    def save_settings(self):
        """
        Gathers all current settings from the UI fields and updates the parent
        GUI's in-memory variables, then forces a save to the JSON file.
        The window remains open.
        """
        try:
            # --- File Paths Tab ---
            self.parent_gui.log_file_path = self.log_file_entry.get().strip()
            self.parent_gui.sqlite_db_path = self.db_file_entry.get().strip()
            
            # --- TXT Aliases and Paths ---
            self.parent_gui.txt_source_aliases["Main TXT"] = self.txt_name_main_var.get().strip()
            self.parent_gui.txt_folder_path = self.txt_path_main_var.get().strip()
            self.parent_gui.txt_source_aliases["TXT Source 2"] = self.txt_name_set2_var.get().strip()
            self.parent_gui.txt_folder_path_set2 = self.txt_path_set2_var.get().strip()
            self.parent_gui.txt_source_aliases["TXT Source 3"] = self.txt_name_set3_var.get().strip()
            self.parent_gui.txt_folder_path_set3 = self.txt_path_set3_var.get().strip()
            self.parent_gui.txt_source_aliases["TXT Source 4"] = self.txt_name_set4_var.get().strip()
            self.parent_gui.txt_folder_path_set4 = self.txt_path_set4_var.get().strip()
            self.parent_gui.txt_source_aliases["TXT Source 5"] = self.txt_name_set5_var.get().strip()
            self.parent_gui.txt_folder_path_set5 = self.txt_path_set5_var.get().strip()
            
            # --- Generated Fields Saving ---
            for i, widget_info in enumerate(self.generated_field_widgets):
                self.parent_gui.generated_fields_config[i]["column_name"] = widget_info["entry"].get().strip()

            # --- Static Fields Saving ---
            new_static_configs = []
            for i, row_info in enumerate(self.static_field_row_widgets):
                new_static_configs.append({
                    "field": row_info["column_entry"].get().strip(), 
                    "description": row_info["description_entry"].get().strip(), 
                    "column_name": row_info["cell_entry"].get().strip(), 
                    "skip": row_info["skip_var"].get()
                })
            self.parent_gui.static_field_configs = new_static_configs

            # --- Monitored Folders Saving ---
            # Local Dictionaries to collect data from UI inputs
            parent_folder_paths, parent_folder_cols, parent_folder_exts, parent_folder_skips, parent_folder_log_x_instead, parent_folder_log_exts = {}, {}, {}, {}, {}, {}
            
            # Loop over ALL configured folder names (including newly added ones)
            for folder_name in self.folder_entries.keys():
                folder_path = self.folder_entries[folder_name].get().strip()
                
                # Only save configuration for folders that have a path set
                if folder_path:
                    parent_folder_paths[folder_name] = folder_path
                    parent_folder_cols[folder_name] = self.folder_column_entries[folder_name].get().strip()
                    parent_folder_exts[folder_name] = self.file_extension_entries[folder_name].get().strip().lstrip('.')
                    parent_folder_skips[folder_name] = self.folder_skip_vars[folder_name].get()
                    parent_folder_log_x_instead[folder_name] = self.folder_log_x_vars[folder_name].get()
                    
                    # --- CRITICAL FIX: Get the value from the local BooleanVar ---
                    # The BooleanVar exists in self.folder_log_ext_vars because it was added in add_folder_row
                    parent_folder_log_exts[folder_name] = self.folder_log_ext_vars[folder_name].get()

            # --- Transfer collected data back to DataLoggerGUI instance ---
            self.parent_gui.folder_paths = parent_folder_paths
            self.parent_gui.folder_columns = parent_folder_cols
            self.parent_gui.file_extensions = parent_folder_exts
            self.parent_gui.folder_skips = parent_folder_skips
            self.parent_gui.folder_log_x_instead = parent_folder_log_x_instead
            
            # --- CRITICAL: Update the parent's persistent BooleanVar references ---
            # We must map the current configuration back to the parent's object, 
            # ensuring only active folders exist in the parent's dictionary.
            new_log_ext_vars_for_parent = {}
            for k, v in parent_folder_paths.items():
                # We reuse the existing BooleanVar if it exists in the UI, or create a temporary one if needed
                if k in self.parent_gui.folder_log_ext_vars:
                    # Update the existing BooleanVar's value using the collected boolean value
                    self.parent_gui.folder_log_ext_vars[k].set(parent_folder_log_exts[k])
                else:
                    # For new entries, create and set the value
                    self.parent_gui.folder_log_ext_vars[k] = tk.BooleanVar(value=parent_folder_log_exts[k])
                
                # Add the reference back to the new structure
                new_log_ext_vars_for_parent[k] = self.parent_gui.folder_log_ext_vars[k]
            
            # Finally, replace the parent's dictionary with only the active folder references
            self.parent_gui.folder_log_ext_vars = new_log_ext_vars_for_parent
            # ------------------------------------------------------------------------

            # --- Button Configuration Saving (Remains unchanged and correct) ---
            self.parent_gui.num_custom_buttons = int(self.num_buttons_entry.get())
            for i, (text_entry, event_entry, event_code_var, txt_source_var, tab_group_var) in enumerate(self.custom_button_widgets):
                
                # We need a robust way to map the ComboBox Display Name back to the Internal Key (e.g., 'Main TXT')
                internal_keys_for_map = TXT_FILES_KEYS
                aliases = self.parent_gui.txt_source_aliases
                display_to_internal_map = {aliases.get(k, k): k for k in internal_keys_for_map}
                
                selected_display_name = txt_source_var.get()
                selected_source_key = display_to_internal_map.get(selected_display_name, selected_display_name)
                
                # Handle Event Code (parse 'Code - Description')
                selected_code_string = event_code_var.get()
                code_to_save = selected_code_string.split(" - ", 1)[0] if " - " in selected_code_string else selected_code_string
                
                # Save changes back to the in-memory list (important for persisting the config)
                config = self.parent_gui.custom_button_configs[i]
                config["text"] = text_entry.get().strip()
                config["event_text"] = event_entry.get().strip()
                config["event_code"] = code_to_save
                config["txt_source_key"] = selected_source_key
                config["tab_group"] = tab_group_var.get().strip() or "Main"

                # Ensure new tab group is registered
                new_group = config["tab_group"]
                if new_group not in self.parent_gui.custom_button_tab_groups:
                    self.parent_gui.custom_button_tab_groups.append(new_group)
                    self.parent_gui.custom_button_tab_groups.sort()

            # --- Programmed Events / Timezone ---
            new_day_bg = self.new_day_bg_color_var.get() if self.new_day_bg_color_var.get() else None
            new_day_font = self.new_day_font_color_var.get() if self.new_day_font_color_var.get() else None
            self.parent_gui.button_colors["New Day"] = (new_day_bg, new_day_font)

            if hasattr(self, 'hourly_source_combobox'):
                selected_display_name = self.hourly_source_combobox.get()
                
                internal_keys_for_map = TXT_FILES_KEYS
                aliases = self.parent_gui.txt_source_aliases
                display_to_internal_map = {aliases.get(k, k): k for k in internal_keys_for_map if k != "None"}

                current_internal_key = self.parent_gui.hourly_log_txt_source_key.get()
                internal_key_to_save = display_to_internal_map.get(selected_display_name, current_internal_key)
                
                self.parent_gui.hourly_log_txt_source_key.set(internal_key_to_save)

            # --- Final Save and UI Refresh ---
            self.parent_gui.save_settings() # This saves all instance attributes to JSON
            self.parent_gui.update_custom_buttons()
            
            self.parent_gui.update_status("Settings saved successfully (window remains open).")

        except Exception as e:
            error_message = f"Error saving settings: {e}"
            messagebox.showerror("Save Error", error_message, parent=self.master)
            self.parent_gui.update_status(error_message)
            print(error_message)
    def save_and_close(self):
        """Saves settings and closes the window."""
        self.save_settings()
        self.master.destroy()
        

    def load_settings(self):
        """Loads settings from the parent DataLoggerGUI instance and populates the UI."""
        
        # --- File Paths Tab ---
        self.log_file_entry.delete(0, tk.END)
        self.log_file_entry.insert(0, self.parent_gui.log_file_path or "")
        self.db_file_entry.delete(0, tk.END)  
        self.db_file_entry.insert(0, self.parent_gui.sqlite_db_path or "")
        
        aliases = self.parent_gui.txt_source_aliases
        self.txt_name_main_var.set(aliases.get("Main TXT", "Main TXT"))
        self.txt_name_set2_var.set(aliases.get("TXT Source 2", "TXT Source 2"))
        self.txt_name_set3_var.set(aliases.get("TXT Source 3", "TXT Source 3"))
        self.txt_name_set4_var.set(aliases.get("TXT Source 4", "TXT Source 4"))
        self.txt_name_set5_var.set(aliases.get("TXT Source 5", "TXT Source 5"))

        self.txt_path_main_var.set(self.parent_gui.txt_folder_path or "")
        self.txt_path_set2_var.set(self.parent_gui.txt_folder_path_set2 or "")
        self.txt_path_set3_var.set(self.parent_gui.txt_folder_path_set3 or "")
        self.txt_path_set4_var.set(self.parent_gui.txt_folder_path_set4 or "")
        self.txt_path_set5_var.set(self.parent_gui.txt_folder_path_set5 or "")

    # --- Data Columns, Monitoring, and Button Configuration loading ---
    # IMPORTANT: Do not reload from disk here; just reflect current parent values.
    # Previously this called parent_gui.load_settings(), which re-read JSON
    # and could overwrite the just-loaded project values. We now only
    # refresh the UI from the current in-memory state.
        self.recreate_static_field_rows()
        
        # --- Monitored Folders loading
        self.folder_entries.clear()
        self.folder_column_entries.clear()
        self.file_extension_entries.clear()
        self.folder_skip_vars.clear()
        self.folder_log_x_vars.clear()
        self.add_initial_folder_rows()
        self.master.after_idle(self.update_scroll_region)
        
        self.num_buttons_entry.delete(0, tk.END)
        self.num_buttons_entry.insert(0, str(self.parent_gui.num_custom_buttons))
        self.recreate_custom_button_settings()
        
        
        self.master.after_idle(lambda: self.static_fields_canvas.config(scrollregion=self.static_fields_canvas.bbox("all")))
        
        # --- Programmed Events Tab Loading ---
        self.parent_gui.new_day_event_enabled_var.set(self.parent_gui.new_day_event_enabled_var.get())
        self.parent_gui.hourly_event_enabled_var.set(self.parent_gui.hourly_event_enabled_var.get())
        self.parent_gui.calculate_logoff_values.set(self.parent_gui.calculate_logoff_values.get())
        
        # Load the color values into the local UI variables
        new_day_bg_color, new_day_font_color = self.parent_gui.button_colors.get("New Day", (None, None))
        self.new_day_bg_color_var.set(new_day_bg_color or "")
        self.new_day_font_color_var.set(new_day_font_color or "")
        
        hourly_bg_color, hourly_font_color = self.parent_gui.button_colors.get("Hourly KP Log", (None, None))
        self.hourly_bg_color_var.set(hourly_bg_color or "")
        self.hourly_font_color_var.set(hourly_font_color or "")

        # --- Timezone Tab ---
        self.parent_gui.time_offset_hours.set(self.parent_gui.time_offset_hours.get())
    
        self.populate_event_codes_tree()
# --- Main Execution ---
if __name__ == "__main__":
    root = tk.Tk()

    gui = DataLoggerGUI(root)

    def on_closing():
        """Handles application closing: stops monitors."""
        active_monitors = list(gui.monitors.items())
        if not active_monitors:
            pass
        else:
            for name, monitor_observer in active_monitors:
                try:
                    if monitor_observer.is_alive():
                        monitor_observer.stop()
                except Exception: pass
            for name, monitor_observer in active_monitors:
                try:
                    if monitor_observer.is_alive():
                        monitor_observer.join(timeout=0.5)
                except Exception: pass
                finally:
                    if name in gui.monitors: del gui.monitors[name]
                if gui._auto_sync_timer_id: #REMOVE
                    gui.master.after_cancel(gui._auto_sync_timer_id)

        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()