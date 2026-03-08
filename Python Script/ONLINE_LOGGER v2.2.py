import sys
import os
import json
import sqlite3
import datetime
import traceback
import re
import uuid
import threading
import socket     
import select     
from pathlib import Path
import pandas as pd
import xlwings as xw
from watchdog.observers.polling import PollingObserver
from watchdog.events import FileSystemEventHandler

from PySide6.QtCore import Qt, QObject, Signal, QTimer, QThread, QPoint
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QDialog,
                               QHBoxLayout, QPushButton, QLabel, QStatusBar, QLineEdit,
                               QProgressBar, QTabWidget, QMessageBox, QFileDialog, QScrollArea, QStackedWidget,
                               QGridLayout, QFrame, QMenu, QInputDialog, QComboBox, QTableWidget,
                               QTableWidgetItem, QHeaderView, QColorDialog, QSpinBox, QGroupBox, QTextEdit, QTextBrowser,
                               QCheckBox, QDialogButtonBox, QTreeWidget, QTreeWidgetItem, QDoubleSpinBox, QSizePolicy)
from PySide6.QtGui import QAction, QColor, QFont, QFontMetrics, QPainter, QPolygon, QBrush, QPainterPath, QPalette, QIcon, QTextCursor
from PySide6.QtWidgets import QAbstractItemView

# --- CONSTANTS ---
APP_VERSION = "2.2"
PROJECT_STATE_FILE = "settings/config/last_project.json"
PROJECT_TEMPLATE_FILE = "settings/config/blank_project.json"
EVENT_CODES_FILE = "settings/event_codes.json"
EXCEL_LOG_REQUIRED_COLS = {'runline', 'kp', 'event'} 
TXT_FILES_KEYS = ["None", "Main TXT", "TXT Source 2", "TXT Source 3", "TXT Source 4", "TXT Source 5"]
DEFAULT_MONITORED_FOLDERS = ["Qinsy DB", "Naviscan", "SIS", "SSS", "SBP", "Mag", "Grad", "SVP", "SpintINS", "Video", "Cathx", "Hypack RAW", "Eiva NaviPac"]

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# =====================================================================
# GLOBAL EXCEPTION HANDLER
# =====================================================================
def global_exception_handler(exc_type, exc_value, exc_traceback):
    """Catches fatal errors, saves history to disk, and displays a GUI popup."""
    traceback.print_exception(exc_type, exc_value, exc_traceback)
    traceback_msg = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
    
    # 1. Gather recent console history
    recent_history = ""
    full_history = ""
    if 'console_logger' in globals():
        full_history = "".join(console_logger.history)
        recent_history = "".join(console_logger.history[-50:]) # Show last 50 lines in the UI
        
    full_msg = f"--- LAST SYSTEM MESSAGES ---\n{recent_history}\n\n--- FATAL CRASH ---\n{traceback_msg}"
    
    # 2. Automatically dump a permanent crash log to the hard drive
    try:
        log_path = os.path.join(os.getcwd(), "crash_log.txt")
        with open(log_path, "w", encoding="utf-8") as f:
            f.write(f"CRASH DATE: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("="*60 + "\n")
            f.write(full_history)
            f.write("\n" + "="*60 + "\n")
            f.write("FATAL TRACEBACK:\n")
            f.write(traceback_msg)
    except Exception:
        pass

    # 3. Show the UI Popup (Forced to the absolute front)
    box = QMessageBox()
    # -Force the crash window to sit on top of everything else
    box.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True) 
    
    box.setIcon(QMessageBox.Icon.Critical)
    box.setWindowTitle("Critical Application Error")
    box.setText("The application encountered a fatal error and must close.\n\nA complete diagnostic log has been saved to 'crash_log.txt' in the application folder.")
    box.setDetailedText(full_msg)
    box.setStyleSheet("QTextEdit { min-width: 600px; min-height: 300px; font-family: monospace; font-size: 12px; }")
    
    # Bring window to focus
    box.raise_()
    box.activateWindow()
    box.exec()

sys.excepthook = global_exception_handler

# =====================================================================
# CORE LOGIC & MANAGERS
# =====================================================================

class ConfigManager:
    """Centralizes all application state, loading, and saving."""
    def __init__(self):
        self.current_project_path = self._load_last_project_path()
        self.settings = self._get_default_settings()
        self.load_project(self.current_project_path)

    def _get_default_settings(self):
        return {
            "log_file_path": "", "sqlite_db_path": "",
            "txt_folder_paths": {k: "" for k in TXT_FILES_KEYS[1:]},
            "txt_source_aliases": {k: k for k in TXT_FILES_KEYS[1:]},
            "all_txt_mappings": {k: [] for k in TXT_FILES_KEYS[1:]},
            "custom_button_configs": [], "num_custom_buttons": 3,
            "button_colors": {}, "custom_button_tab_groups": ["Main"],
            "main_button_configs": {
                "Log on": {"event_text": "Log on event occurred", "event_code": ""},
                "Log off": {"event_text": "Log off event occurred", "event_code": ""},
                "Event": {"event_text": "", "event_code": ""},
                "SVP": {"event_text": "SVP applied", "event_code": ""},
                "Manual KP Log": {"event_text": "Auto generated", "event_code": ""}
            },
            "time_offset_hours": 0.0, "hourly_log_txt_source_key": "Main TXT",
            "active_logging_threshold_seconds": 15, "new_day_event_enabled": True,
            "new_day_event_code": "", "hourly_event_enabled": True, "calculate_logoff_values": True,
            "auto_sync_interval_min": 15,
            "generated_fields_config": [
                {"field": "Date-Time", "column_name": "UTC Date-Time", "skip": False, "source": "PC Time (UTC)"},
                {"field": "Local Time", "column_name": "Local Time", "skip": False, "source": "PC Time + Offset"},
                {"field": "Event", "column_name": "Event", "skip": False, "source": "Button"},
                {"field": "Code", "column_name": "Code", "skip": False, "source": "Button"},
                {"field": "KP Ref.", "column_name": "KP Ref.", "skip": False, "source": "Source Alias"},
                {"field": "UUID", "column_name": "UUID", "skip": False, "source": "Generated"}
            ],
            "static_field_configs": [],
            "folder_paths": {}, "folder_columns": {}, "file_extensions": {},
            "folder_skips": {}, "folder_log_x_instead": {}, "folder_log_ext_vars": {},
            "udp_trigger_enabled": False, "udp_trigger_port": 5999, 
            "udp_payload_recording": "RECORDING", "udp_payload_idle": "IDLE", 
            "always_on_top": False,
            "dark_mode": False,
            "event_codes": self._load_event_codes()
        }

    def _load_last_project_path(self):
        state_path = os.path.join(os.getcwd(), PROJECT_STATE_FILE)
        if os.path.exists(state_path):
            try:
                with open(state_path, 'r', encoding='utf-8') as f:
                    return json.load(f).get('current_project_path')
            except Exception: pass
        return os.path.join(os.getcwd(), "settings", "default_project.json")

    def persist_current_project_path(self, path):
        state_path = os.path.join(os.getcwd(), PROJECT_STATE_FILE)
        os.makedirs(os.path.dirname(state_path), exist_ok=True)
        try:
            with open(state_path, 'w', encoding='utf-8') as f:
                json.dump({'current_project_path': path}, f, indent=4)
        except Exception: pass

    def _load_event_codes(self):
        if os.path.exists(EVENT_CODES_FILE):
            try:
                with open(EVENT_CODES_FILE, 'r') as ecf: return json.load(ecf)
            except Exception: pass
        return {}

    def load_project(self, path):
        if not path: return False, "No path provided."
        self.current_project_path = path
        self.persist_current_project_path(path)
        if os.path.exists(path):
            try:
                with open(path, 'r') as f:
                    loaded = json.load(f)
                    self.settings.update(loaded)
                    
                    # Only apply legacy migration if the new dictionary DOES NOT exist yet
                    if "txt_folder_paths" not in loaded and "txt_folder_path" in loaded:
                        self.settings["txt_folder_paths"]["Main TXT"] = loaded.get("txt_folder_path", "")
                        self.settings["txt_folder_paths"]["TXT Source 2"] = loaded.get("txt_folder_path_set2", "")
                        self.settings["txt_folder_paths"]["TXT Source 3"] = loaded.get("txt_folder_path_set3", "")
                        self.settings["txt_folder_paths"]["TXT Source 4"] = loaded.get("txt_folder_path_set4", "")
                        self.settings["txt_folder_paths"]["TXT Source 5"] = loaded.get("txt_folder_path_set5", "")
                return True, "Loaded successfully."
            except Exception as e: return False, f"Failed to load settings: {e}"
        return False, "File does not exist."

    def save(self):
        os.makedirs(os.path.dirname(self.current_project_path), exist_ok=True)
        try:
            with open(self.current_project_path, 'w') as f: json.dump(self.settings, f, indent=4)
            return True, "Saved successfully."
        except Exception as e: 
            return False, f"Failed to save settings: {e}"
            
    def get(self, key, default=None): return self.settings.get(key, default)
    def set(self, key, value): self.settings[key] = value


class ExcelLogger:
    """Handles isolated Excel COM manipulation to prevent GUI lockups."""
    def __init__(self, log_file_path, sqlite_manager=None):
        self.log_file_path = log_file_path
        self.sqlite_manager = sqlite_manager

    def log_entry(self, row_data, bg_color=None, date_columns=None, static_configs=None):
        if not self.log_file_path or not os.path.exists(self.log_file_path): 
            return False, False, "Excel file path is invalid or missing."

        try:
            clean_row_data = {str(k).strip(): v for k, v in row_data.items()}
            date_cols = [str(c).lower() for c in (date_columns or []) if c]

            wb = xw.Book(self.log_file_path)
            
            # --- Safely read static cells on the isolated background thread ---
            if static_configs:
                for config in static_configs:
                    excel_col_key = config.get("field", "").strip()
                    lookup_str = config.get("column_name", "").strip()
                    try:
                        match = re.match(r"='?([^'!]+)'?!([A-Z]+\d+)", lookup_str)
                        if match: clean_row_data[excel_col_key] = wb.sheets[match.group(1)].range(match.group(2)).value
                        elif re.match(r"=?([A-Z]+\d+)", lookup_str): clean_row_data[excel_col_key] = wb.sheets[0].range(re.match(r"=?([A-Z]+\d+)", lookup_str).group(1)).value
                    except Exception as e:
                        pass # Ignore errors if cell is empty or locked
            # ----------------------------------------------------------------------

            sheet = wb.sheets[0]

            headers, header_row_idx = self._find_headers(sheet, clean_row_data)
            last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row + 1
            if last_row <= header_row_idx: last_row = header_row_idx + 1

            if not headers:
                headers = list(clean_row_data.keys())
                sheet.range('A1').value = headers
                header_row_idx = 1
                last_row = 2

            new_row, dt_format_indices = self._prepare_row(headers, clean_row_data, date_cols)

            sheet.range(f'A{last_row}').value = new_row
            self._apply_formatting(sheet, last_row, dt_format_indices, headers, bg_color)
            wb.save()

            sqlite_success = False
            if self.sqlite_manager:
                sqlite_success = self.sqlite_manager.add_single_row(clean_row_data, excel_path=self.log_file_path, excel_row=last_row)

            return True, sqlite_success, "Data mapped and saved successfully."
        except Exception as e:
            traceback.print_exc()
            err_msg = str(e)
            
            # Translate common confusing Windows COM errors into plain English
            if "Permission denied" in err_msg or "used by another process" in err_msg or "0x800A03EC" in err_msg:
                err_msg = "Permission Denied. The file is likely open in Excel on another computer."
            elif "No such file" in err_msg:
                err_msg = "File not found. Please check the Excel file path."
                
            return False, False, err_msg

    def _find_headers(self, sheet, clean_row_data):
        target_keys = {k.lower() for k in clean_row_data.keys()}
        best_match = -1
        header_row = 1
        headers = []

        for i in range(1, 11):
            row_vals = sheet.range(f'A{i}:AZ{i}').value
            if not row_vals: continue
            
            cleaned = [str(h).strip().lower() if h else "" for h in row_vals]
            match_count = sum(1 for h in cleaned if h in target_keys)
            
            if match_count > best_match and match_count > 0:
                best_match = match_count
                header_row = i
                headers = [str(h).strip() if h else "" for h in row_vals]
            if match_count >= 2: break

        while headers and not headers[-1]: headers.pop()
        return headers, header_row

    def _prepare_row(self, headers, clean_row_data, date_cols):
        new_row = [""] * len(headers)
        dt_format_indices = []

        for key, val in clean_row_data.items():
            key_str = key.lower()
            for col_idx, h in enumerate(headers):
                if h.lower() == key_str:
                    if key_str in date_cols and val:
                        try:
                            dt_obj = datetime.datetime.strptime(str(val), "%Y-%m-%d %H:%M:%S")
                            delta = dt_obj - datetime.datetime(1899, 12, 30)
                            val = delta.days + (delta.seconds / 86400.0)
                            dt_format_indices.append(col_idx + 1)
                        except ValueError: pass
                    new_row[col_idx] = val
                    break
        return new_row, dt_format_indices

    def _apply_formatting(self, sheet, target_row, dt_indices, headers, bg_color):
        for idx in dt_indices:
            sheet.cells(target_row, idx).number_format = 'yyyy-mm-dd hh:mm:ss'
        if bg_color:
            rgb_bg = tuple(int(bg_color.lstrip('#')[i:i+2], 16) for i in (0, 2, 4)) if isinstance(bg_color, str) and bg_color.startswith('#') else bg_color
            col_count = len(headers)
            if col_count > 0: sheet.range((target_row, 1), (target_row, col_count)).color = rgb_bg


class SQLiteManager:
    def __init__(self, db_path):
        self.db_path = db_path
        self.conn = None
        self.table_name = None
        try:
            self.conn = sqlite3.connect(db_path, check_same_thread=False, timeout=30.0)
            self.conn.execute("PRAGMA journal_mode=WAL")
            self.conn.execute("PRAGMA busy_timeout=30000")
        except sqlite3.Error as e: print(f"SQLite Connection error: {e}")

    def close(self):
        if self.conn:
            self.conn.close()
            self.conn = None

    def verify_wal_mode(self):
        try:
            cursor = self.conn.cursor()
            cursor.execute("PRAGMA journal_mode")
            if cursor.fetchone()[0].lower() != 'wal':
                cursor.execute("PRAGMA journal_mode=WAL")
            return True
        except Exception: return False

    def _sanitize_column_name(self, name):
        return re.sub(r'[^A-Za-z0-9_]', '', str(name).replace(' ', '_').replace('-', '_'))

    def _read_excel_data(self, excel_path, header_finder_func):
        try:
            header_row_index = header_finder_func(excel_path)
            if header_row_index == -1: return None, -1
            df = pd.read_excel(excel_path, sheet_name=0, header=header_row_index, skiprows=header_row_index)
            df.dropna(how='all', inplace=True)
            df.reset_index(drop=True, inplace=True)
            
            date_keywords = ['date', 'time', 'datetime', 'timestamp', 'utc', 'local']
            for col in df.columns:
                if any(k in str(col).lower() for k in date_keywords):
                    def robust_date_convert(val):
                        if pd.isna(val): return ""
                        if hasattr(val, 'strftime'): return val.strftime('%Y-%m-%d %H:%M:%S')
                        if "datetime64" in str(type(val)): return pd.Timestamp(val).strftime('%Y-%m-%d %H:%M:%S')
                        if isinstance(val, (float, int)):
                            try:
                                return (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(val))).strftime('%Y-%m-%d %H:%M:%S')
                            except Exception: return str(val)
                        return str(val).strip()
                    df[col] = df[col].apply(robust_date_convert)
            return df, header_row_index
        except Exception: return None, -1

    def full_sync(self, excel_path, header_finder_func, skip_uuid_fixes=True):
        try:
            df, header_row = self._read_excel_data(excel_path, header_finder_func)
            if df is None: return False
            
            index_col = next((c for c in df.columns if str(c).strip().lower() == 'index'), None)
            index_numbers = [i + header_row + 2 for i in range(len(df))]
            if index_col is None: df.insert(0, 'index', index_numbers)
            else: df[index_col] = index_numbers

            df.columns = [self._sanitize_column_name(col) for col in df.columns]
            self.table_name = Path(excel_path).stem
            
            cursor = self.conn.cursor()
            cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name='{self.table_name}'")
            if cursor.fetchone():
                cursor.execute(f"PRAGMA table_info('{self.table_name}')")
                if {row[1] for row in cursor.fetchall()} != set(df.columns):
                    cursor.execute(f"DROP TABLE IF EXISTS '{self.table_name}'")
            
            cursor.execute(f"CREATE TABLE IF NOT EXISTS '{self.table_name}' ({', '.join([f'\"{c}\" TEXT' for c in df.columns])})")
            cursor.execute(f"DELETE FROM '{self.table_name}'")
            
            cols = list(df.columns)
            placeholders = ', '.join(['?' for _ in cols])
            cursor.executemany(f"INSERT INTO '{self.table_name}' ({', '.join([f'\"{c}\"' for c in cols])}) VALUES ({placeholders})", 
                               [tuple(row) for row in df.values])
            self.conn.commit()
            return True
        except Exception: return False

    def add_single_row(self, row_data, excel_path=None, excel_row=None):
        if not self.table_name and excel_path: self.table_name = Path(excel_path).stem
        if not self.table_name: return False

        try:
            sanitized_data = {'index': str(excel_row)} if excel_row is not None else {}
            for key, value in row_data.items(): sanitized_data[self._sanitize_column_name(key)] = str(value) if value is not None else ""

            uuid_key = self._sanitize_column_name('UUID')
            if uuid_key in sanitized_data and not sanitized_data[uuid_key]: sanitized_data[uuid_key] = str(uuid.uuid4())

            cursor = self.conn.cursor()
            cols = list(sanitized_data.keys())
            sql = f"INSERT INTO '{self.table_name}' ({', '.join([f'\"{c}\"' for c in cols])}) VALUES ({', '.join(['?' for _ in cols])})"
            cursor.execute(sql, list(sanitized_data.values()))
            self.conn.commit()
            return True
        except Exception: return False


class FolderMonitor(FileSystemEventHandler):
    def __init__(self, path, folder_name, cache, cache_lock, extension=""):
        self.path = path
        self.folder_name = folder_name
        self.cache = cache            
        self.cache_lock = cache_lock 
        self.extension = extension.lower()
        self._initial_full_scan()

    def _is_valid_file(self, file_path):
        if not os.path.isfile(file_path): return False
        if self.extension and not file_path.lower().endswith(self.extension): return False
        return True

    def _handle_file_event(self, file_path):
        if self._is_valid_file(file_path):
            with self.cache_lock:
                # 1. Get the current cache list (or empty list if new)
                current_list = self.cache.get(self.folder_name, [])
                
                # 2. Add the new/modified file to the list if it isn't there already
                if file_path not in current_list:
                    current_list.append(file_path)
                
                # 3. Clean up deleted files and sort by Creation Time (Newest First)
                valid_files = [f for f in current_list if os.path.exists(f)]
                valid_files.sort(key=lambda x: os.path.getctime(x), reverse=True)
                
                # 4. Keep ONLY the latest 3 files in the cache!
                self.cache[self.folder_name] = valid_files[:3]

    def on_created(self, event):
        if not event.is_directory: self._handle_file_event(event.src_path)

    def on_modified(self, event):
        if not event.is_directory: self._handle_file_event(event.src_path)
            
    def on_moved(self, event):
        if not event.is_directory: self._handle_file_event(event.dest_path)

    def _initial_full_scan(self):
        all_files = []
        try:
            for root, _, files in os.walk(self.path):
                for f_name in files:
                    f_path = os.path.join(root, f_name)
                    if self._is_valid_file(f_path):
                        all_files.append(f_path)
            
            # Sort all found files by Creation Time (Newest First)
            all_files.sort(key=lambda x: os.path.getctime(x), reverse=True)
            
            with self.cache_lock:
                # Store the 3 most recent files in the cache
                self.cache[self.folder_name] = all_files[:3]
        except Exception: pass

class MonitorManager:
    def __init__(self):
        self.observer = None
        self.cache = {}
        self.cache_lock = threading.Lock()
        self.is_monitoring = False

    def start_monitoring(self, folder_configs):
        if self.is_monitoring: self.stop_monitoring()
        self.cache.clear()
        self.observer = PollingObserver()
        monitors_added = 0
        
        for name, data in folder_configs.items():
            path = data.get("path")
            if not path or not os.path.exists(path) or data.get("skip", False): continue
            handler = FolderMonitor(path, name, self.cache, self.cache_lock, data.get("ext", ""))
            self.observer.schedule(handler, path, recursive=True)
            monitors_added += 1

        if monitors_added > 0:
            self.observer.start()
            self.is_monitoring = True
            return True, f"Started monitoring {monitors_added} folders."
        return False, "No valid folders found to monitor. Check your paths."

    def stop_monitoring(self):
        if self.observer:
            self.observer.stop()
            self.observer.join()
            self.observer = None
        self.is_monitoring = False

    def get_latest_file(self, folder_name):
        """Pulls the absolutely newest valid file from the Top 3 cache."""
        with self.cache_lock: 
            file_list = self.cache.get(folder_name, [])
            # Iterate through the cached 3 files. Return the first one that STILL exists.
            for f in file_list:
                if os.path.exists(f):
                    return f
            return None

# =====================================================================
# THREAD WORKERS
# =====================================================================

class MonitorSetupWorker(QObject):
    finished = Signal(bool, str)
    def __init__(self, manager, configs):
        super().__init__()
        self.manager = manager
        self.configs = configs
        
    def run(self):
        try:
            # The heavy os.walk scanning happens safely inside here now
            success, msg = self.manager.start_monitoring(self.configs)
            self.finished.emit(success, msg)
        except Exception as e:
            self.finished.emit(False, f"Error: {str(e)}")

class SqliteSyncWorker(QObject):
    finished = Signal(bool, str)
    def __init__(self, manager, excel_path, header_func):
        super().__init__()
        self.manager, self.excel_path, self.header_func = manager, excel_path, header_func
    def run(self):
        try:
            try: import pythoncom; pythoncom.CoInitialize()
            except ImportError: pass
            success = self.manager.full_sync(self.excel_path, self.header_func, skip_uuid_fixes=False)
            self.finished.emit(True, "SQLite full sync complete.") if success else self.finished.emit(False, "SQLite full sync failed.")
        except Exception as e: self.finished.emit(False, f"SQLite Sync Error: {str(e)}")
        finally:
            try: import pythoncom; pythoncom.CoUninitialize()
            except ImportError: pass

class UdpListenerWorker(QObject):
    trigger_log = Signal(str)
    
    def __init__(self, port, payload_rec="RECORDING", payload_idle="IDLE"):
        super().__init__()
        #  Start with current_state as None instead of "IDLE"
        self.port, self.running, self.sock, self.current_state = port, True, None, None
        self.payload_rec = payload_rec.upper()
        self.payload_idle = payload_idle.upper()
        
    def run(self):
        try:
            self.sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            self.sock.bind(('', self.port))
            self.sock.setblocking(0)
        except Exception as e: return
        while self.running:
            try:
                ready = select.select([self.sock], [], [], 0.5)
                if ready[0]:
                    data, _ = self.sock.recvfrom(1024)
                    payload = data.decode('utf-8', errors='ignore').strip().upper()
                    
                    is_recording_payload = self.payload_rec in payload
                    is_idle_payload = self.payload_idle in payload

                    # 1. Silently calibrate the state on the very first packet after boot
                    if self.current_state is None:
                        if is_recording_payload:
                            self.current_state = "RECORDING"
                        elif is_idle_payload:
                            self.current_state = "IDLE"
                        continue # Skip the trigger logic below for this first packet

                    # 2. Only trigger events on actual state CHANGES
                    if is_recording_payload and self.current_state != "RECORDING":
                        self.current_state = "RECORDING"
                        self.trigger_log.emit("Log on")
                    elif is_idle_payload and self.current_state != "IDLE":
                        self.current_state = "IDLE"
                        self.trigger_log.emit("Log off")
            except Exception: pass
        if self.sock: self.sock.close()
        
    def stop(self): self.running = False

class LogWorker(QObject):
    finished = Signal(bool, bool, str, object, str) 
    
    def __init__(self, excel_logger, task_data, btn=None, orig_text=None):
        super().__init__()
        self.logger = excel_logger
        self.data = task_data
        self.btn = btn
        self.orig_text = orig_text if orig_text else ""
        
    def run(self):
        print(f"[DEBUG - LOG WORKER] Thread started for action: '{self.orig_text}'")
        try:
            try: 
                import pythoncom; pythoncom.CoInitialize()
                print(f"[DEBUG - LOG WORKER] COM Initialized. Sending data to Excel...")
            except ImportError: pass
            
            s_ex, s_sq, msg = self.logger.log_entry(
                self.data['row_data'], 
                self.data['bg_color'], 
                self.data.get('date_columns', ["UTC Date-Time", "Local Time"]),
                self.data.get('static_configs', [])
            )
            print(f"[DEBUG - LOG WORKER] Excel write complete. Success: {s_ex}")
            self.finished.emit(s_ex, s_sq, msg, self.btn, self.orig_text)
        except Exception as e: 
            print(f"[DEBUG - LOG WORKER] CRASHED: {str(e)}")
            self.finished.emit(False, False, f"Thread Error: {str(e)}", self.btn, self.orig_text)
        finally:
            print(f"[DEBUG - LOG WORKER] Releasing COM lock and closing thread.")
            try: import pythoncom; pythoncom.CoUninitialize()
            except ImportError: pass

class HourlyCalcWorker(QObject):
    finished = Signal(bool, str, object, object, object) 
    def __init__(self, log_file_path, event_col, kp_col, line_col, dt_col, curr_kp, curr_line, curr_time, last_kp, last_line, last_time):
        super().__init__()
        self.log_file_path, self.event_col, self.kp_col, self.line_col, self.dt_col = log_file_path, event_col, kp_col, line_col, dt_col
        self.curr_kp, self.curr_line, self.curr_time = curr_kp, curr_line, curr_time
        self.last_kp, self.last_line, self.last_time = last_kp, last_line, last_time
        
    def run(self):
        print("\n[DEBUG - HOURLY WORKER] Thread started.")
        try:
            try: 
                import pythoncom; pythoncom.CoInitialize()
                print("[DEBUG - HOURLY WORKER] COM Initialized.")
            except ImportError: pass
            
            if self.last_kp is None or self.last_line is None: 
                print("[DEBUG - HOURLY WORKER] Memory cache empty. Must read Excel history...")
                self._lookup_history_in_excel()
                print(f"[DEBUG - HOURLY WORKER] History lookup complete! Found Last KP: {self.last_kp}")
            else:
                print("[DEBUG - HOURLY WORKER] Memory cache hit! Skipping Excel history read.")
                
            event_text = self._generate_event_text()
            self.finished.emit(True, event_text, self.curr_kp, self.curr_line, self.curr_time)
            
        except Exception as e: 
            print(f"[DEBUG - HOURLY WORKER] CRASHED: {str(e)}")
            self.finished.emit(False, str(e), None, None, None)
        finally:
            print("[DEBUG - HOURLY WORKER] Releasing COM lock and closing thread.\n")
            try: import pythoncom; pythoncom.CoUninitialize()
            except ImportError: pass

    def _lookup_history_in_excel(self):
        try:
            print("[DEBUG - HOURLY WORKER] Connecting to xlwings App...")
            wb = xw.Book(self.log_file_path)
            sheet = wb.sheets[0]
            h_idx, h_vals = -1, []
            
            print("[DEBUG - HOURLY WORKER] Scanning for headers...")
            for i in range(1, 31):
                row_vals = sheet.range(f'A{i}:AZ{i}').value
                if not row_vals: continue
                curr_h = [str(h).lower().strip() if h else '' for h in row_vals]
                if str(self.event_col).lower().strip() in curr_h:
                    h_idx, h_vals = i, curr_h
                    break
                    
            if h_idx != -1:
                e_idx = h_vals.index(str(self.event_col).lower().strip())
                k_idx = h_vals.index(str(self.kp_col).lower().strip()) if str(self.kp_col).lower().strip() in h_vals else -1
                l_idx = h_vals.index(str(self.line_col).lower().strip()) if str(self.line_col).lower().strip() in h_vals else -1
                t_idx = h_vals.index(str(self.dt_col).lower().strip()) if str(self.dt_col).lower().strip() in h_vals else -1
                
                if k_idx != -1:
                    last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
                    start_row = max(h_idx + 1, last_row - 500)
                    
                    if last_row >= start_row:
                        print(f"[DEBUG - HOURLY WORKER] Pulling rows {start_row} to {last_row} from Excel...")
                        data_block = sheet.range((start_row, 1), (last_row, len(h_vals))).value
                        if data_block and not isinstance(data_block[0], list): data_block = [data_block]
                        for row in reversed(data_block):
                            if row and len(row) > e_idx and row[e_idx] and str(row[e_idx]).startswith("Current KP:"):
                                try:
                                    self.last_kp = float(row[k_idx])
                                    if l_idx != -1 and len(row) > l_idx: self.last_line = row[l_idx]
                                    if t_idx != -1 and len(row) > t_idx:
                                        tv = row[t_idx]
                                        if isinstance(tv, datetime.datetime): self.last_time = tv
                                        elif isinstance(tv, str):
                                            try: self.last_time = datetime.datetime.strptime(tv, "%Y-%m-%d %H:%M:%S")
                                            except Exception: pass
                                    break
                                except Exception: continue
        except Exception as e: 
            print(f"[DEBUG - HOURLY WORKER] History Lookup Failed: {e}")

    def _generate_event_text(self):
        if self.last_kp is not None and self.last_line is not None:
            if str(self.curr_line).strip() == str(self.last_line).strip():
                progress = self.curr_kp - self.last_kp
                time_str, speed_str = "since last log", ""
                if self.last_time:
                    diff = self.curr_time - self.last_time
                    time_diff_seconds = diff.total_seconds()
                    hours, rem = divmod(int(time_diff_seconds), 3600)
                    mins, secs = divmod(rem, 60)
                    time_str = f"in {hours}h {mins}m" if hours > 0 else f"in {mins}m {secs}s"
                    if time_diff_seconds > 1: speed_str = f" (@ {(abs(progress) / 1.852) / (time_diff_seconds / 3600):.1f} kts)"
                return f"Current KP: {self.curr_kp:.3f} | Progress {time_str}: {progress:+.3f} km{speed_str} | Line: {self.curr_line}"
            else: return f"Current KP: {self.curr_kp:.3f} | **LINE CHANGED** from {self.last_line} to {self.curr_line}. Progress on new line: {self.curr_kp:.3f} km"
        else: return f"Current KP: {self.curr_kp:.3f} | First KP log on Line: {self.curr_line}"


# =====================================================================
# GUI COMPONENTS
# =====================================================================

# --- Dynamic Font Button with Color Wedge ---
class AutoScalingButton(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self._min_size, self._max_size = 8, 36
        self.wedge_color = None
        self._is_updating = False
        
        # Stop the button from pushing the layout around when its text changes:
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

    def sizeHint(self):
        # Override sizeHint so dynamic text/font changes NEVER trigger layout shifts!
        from PySide6.QtCore import QSize
        return QSize(100, 65)

    def set_wedge_color(self, color_hex):
        self.wedge_color = color_hex
        self.update()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._scale_font()

    def setText(self, text):
        super().setText(text)
        self._scale_font()

    def _scale_font(self):
        # Prevent recursive calculation loops
        if getattr(self, '_is_updating', False): return
        self._is_updating = True
        
        # 1. Clean the text of any existing artificial newlines so we can test it fresh
        current_text = self.text().replace('\n', ' ')
        if not current_text or self.width() <= 0 or self.height() <= 0: 
            self._is_updating = False
            return
        
        margin_x, margin_y = 34, 30
        target_w = self.width() - margin_x
        target_h = self.height() - margin_y
        
        best_size = 0
        best_text = current_text
        
        # Test A: Standard single-line text
        for size in range(self._max_size, self._min_size - 1, -1):
            font = self.font()
            font.setPointSize(size)
            fm = QFontMetrics(font)
            rect = fm.boundingRect(0, 0, 10000, 10000, int(Qt.AlignmentFlag.AlignCenter), current_text)
            if rect.width() <= target_w and rect.height() <= target_h:
                best_size = size
                break
                
        # Test B: Multi-line text (Find the space closest to the center and break the line)
        if ' ' in current_text:
            mid = len(current_text) // 2
            spaces = [i for i, c in enumerate(current_text) if c == ' ']
            closest_space = min(spaces, key=lambda x: abs(x - mid))
            split_text = current_text[:closest_space] + '\n' + current_text[closest_space+1:]
            
            # Check if this split allows for a BIGGER font size than Test A
            for size in range(self._max_size, best_size, -1): 
                font = self.font()
                font.setPointSize(size)
                fm = QFontMetrics(font)
                rect = fm.boundingRect(0, 0, 10000, 10000, int(Qt.AlignmentFlag.AlignCenter), split_text)
                if rect.width() <= target_w and rect.height() <= target_h:
                    best_size = size
                    best_text = split_text
                    break
        
        # 3. Apply the winning font size
        final_font = self.font()
        final_font.setPointSize(max(best_size, self._min_size))
        self.setFont(final_font)
        
        # 4. Apply the winning text format
        if self.text() != best_text:
            self.blockSignals(True)
            super().setText(best_text)
            self.blockSignals(False)
            
        self._is_updating = False

    def paintEvent(self, event):
        # Let PySide6 draw the dark background and text
        super().paintEvent(event)

        # Draw the wedge on top
        if hasattr(self, 'wedge_color') and self.wedge_color:
            color = QColor(self.wedge_color)
            if color.isValid():
                painter = QPainter(self)
                painter.setRenderHint(QPainter.RenderHint.Antialiasing)
                w, h = self.width(), self.height()
                
                # Match the 6px rounded corners from the CSS
                path = QPainterPath()
                path.addRoundedRect(0, 0, w, h, 6, 6)
                painter.setClipPath(path)
                
                # Size of the triangle in pixels
                wedge_size = 25
                polygon = QPolygon([QPoint(w - wedge_size, h), QPoint(w, h), QPoint(w, h - wedge_size)])
                
                painter.setPen(Qt.PenStyle.NoPen)
                painter.setBrush(QBrush(color))
                painter.drawPolygon(polygon)
                painter.end()


class HistoricEventDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)
        self.setWindowTitle("Add Historic Event")
        self.setFixedSize(500, 180)
        self.setStyleSheet("""
            QDialog { background-color: palette(window); }
            QLabel { color: palette(window-text); font-weight: bold; }
            QLineEdit { background-color: palette(base); color: palette(text); border: 1px solid rgba(128, 128, 128, 0.6); border-radius: 4px; padding: 6px; }
            QLineEdit:focus { border: 2px solid #0078D4; }
            QPushButton { border: 1px solid rgba(128, 128, 128, 0.5); border-radius: 4px; padding: 6px 12px; background-color: rgba(128, 128, 128, 0.05); color: palette(text); font-weight: bold; }
            QPushButton:hover { background-color: rgba(128, 128, 128, 0.15); }
            QCheckBox { spacing: 8px; color: palette(window-text); font-weight: bold; }
            QCheckBox::indicator { width: 18px; height: 18px; border: 2px solid palette(text); border-radius: 4px; background-color: transparent; }
            QCheckBox::indicator:hover { border: 2px solid #0078D4; }
            QCheckBox::indicator:checked { background-color: #0078D4; border: 2px solid #0078D4; image: url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>'); }
        """)
        layout = QVBoxLayout(self); grid = QGridLayout(); layout.addLayout(grid)
        grid.addWidget(QLabel("Log File:"), 0, 0)
        self.edit_file_path = QLineEdit("No file selected..."); self.edit_file_path.setReadOnly(True)
        grid.addWidget(self.edit_file_path, 0, 1)
        btn_browse = QPushButton("Browse..."); btn_browse.setCursor(Qt.CursorShape.PointingHandCursor); btn_browse.clicked.connect(self.browse_file)
        grid.addWidget(btn_browse, 0, 2)
        grid.addWidget(QLabel("Time to Find (HH:MM:SS or HH:MM):"), 1, 0)
        self.time_edit = QLineEdit(datetime.datetime.now().strftime("%H:%M:%S"))
        grid.addWidget(self.time_edit, 1, 1, 1, 2)
        self.chk_sfile = QCheckBox("Insert S-File ID")
        grid.addWidget(self.chk_sfile, 2, 0, 1, 3)
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.button_box.accepted.connect(self.validate_and_accept); self.button_box.rejected.connect(self.reject)
        layout.addWidget(self.button_box)
        self.result_data = None
        
    def browse_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select log file", "", "Data (*.txt *.npd *.csv);;All (*.*)")
        if path: self.edit_file_path.setText(path)
        
    def validate_and_accept(self):
        path = self.edit_file_path.text()
        time_str = self.time_edit.text().strip()
        if not os.path.isfile(path): return QMessageBox.warning(self, "Input Error", "Select a valid log file first.")
        try: t_obj = datetime.datetime.strptime(time_str, '%H:%M:%S').time()
        except ValueError:
            try: t_obj = datetime.datetime.strptime(time_str, '%H:%M').time()
            except ValueError: return QMessageBox.warning(self, "Invalid Format", "Please enter the time in HH:MM:SS or HH:MM format.")
        self.result_data = {'file_path': path, 'time_str': time_str, 'time_obj': t_obj, 'insert_sfile': self.chk_sfile.isChecked()}
        self.accept()

class HistoricPreviewDialog(QDialog):
    def __init__(self, parent, raw_line, parsed_data):
        super().__init__(parent)
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)
        self.setWindowTitle("Confirm Historic Data Mapping")
        self.setMinimumSize(500, 400)
        self.setStyleSheet("""
            QDialog { background-color: palette(window); }
            QLabel { color: palette(window-text); font-size: 13px; }
            QTextEdit { background-color: palette(base); color: palette(text); border: 1px solid rgba(128, 128, 128, 0.4); border-radius: 4px; padding: 4px; font-family: monospace; }
            QTableWidget { background-color: palette(base); color: palette(text); border: 1px solid rgba(128, 128, 128, 0.4); border-radius: 4px; gridline-color: rgba(128, 128, 128, 0.2); }
            QHeaderView::section { background-color: rgba(128, 128, 128, 0.1); color: palette(window-text); padding: 5px; border: 1px solid rgba(128, 128, 128, 0.2); font-weight: bold; }
            QPushButton { border: 1px solid rgba(128, 128, 128, 0.5); border-radius: 4px; padding: 6px 12px; background-color: rgba(128, 128, 128, 0.05); color: palette(text); font-weight: bold; }
            QPushButton:hover { background-color: rgba(0, 120, 212, 0.1); border-color: #0078D4; color: #0078D4;}
        """)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("<b>1. Raw Data Line Found:</b>"))
        raw_text = QTextEdit(); raw_text.setPlainText(raw_line); raw_text.setReadOnly(True); raw_text.setMaximumHeight(65)
        layout.addWidget(raw_text)
        layout.addWidget(QLabel("<b>2. Mapped Data to be Logged:</b>"))
        self.table = QTableWidget(len(parsed_data), 2)
        self.table.setHorizontalHeaderLabels(["Excel Column", "Extracted Value"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table.verticalHeader().setVisible(False); self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers); self.table.setAlternatingRowColors(True)
        for row, (col, val) in enumerate(parsed_data.items()):
            self.table.setItem(row, 0, QTableWidgetItem(str(col)))
            self.table.setItem(row, 1, QTableWidgetItem(str(val)))
        layout.addWidget(self.table)
        bbox = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        bbox.accepted.connect(self.accept); bbox.rejected.connect(self.reject)
        layout.addWidget(bbox)

class HandoverReportDialog(QDialog):
    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)
        self.gui = parent
        self.setWindowTitle("Shift Handover Report Builder")
        self.resize(950, 650)
        self.setStyleSheet("""
            QDialog { background-color: palette(window); }
            QLabel { color: palette(window-text); font-weight: bold; font-size: 13px; }
            QTextEdit { background-color: palette(base); color: palette(text); border: 1px solid rgba(128, 128, 128, 0.4); border-radius: 4px; padding: 10px; font-family: monospace; font-size: 13px; }
            QPushButton { border: 1px solid rgba(128, 128, 128, 0.5); border-radius: 4px; padding: 8px 15px; background-color: rgba(128, 128, 128, 0.05); color: palette(text); font-weight: bold; }
            QPushButton:hover { background-color: rgba(0, 120, 212, 0.1); border-color: #0078D4; color: #0078D4;}
            QPushButton#ActionBtn { background-color: #0078D4; color: white; border: none; }
            QPushButton#ActionBtn:hover { background-color: #106ebe; }
            QComboBox { background-color: palette(base); color: palette(text); border: 1px solid rgba(128, 128, 128, 0.6); border-radius: 4px; padding: 4px; }
            QGroupBox { font-weight: bold; border: 1px solid rgba(128, 128, 128, 0.4); border-radius: 6px; margin-top: 10px; padding-top: 15px; }
            QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top left; padding: 0 5px; left: 10px; }
            QCheckBox { spacing: 8px; color: palette(window-text); font-weight: bold; }
            QCheckBox::indicator { width: 16px; height: 16px; border: 2px solid palette(text); border-radius: 4px; background-color: transparent; }
            QCheckBox::indicator:hover { border: 2px solid #0078D4; }
            QCheckBox::indicator:checked { background-color: #0078D4; border: 2px solid #0078D4; image: url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>'); }
        """)
        
        main_layout = QHBoxLayout(self)
        
        # --- Load Saved Preferences ---
        prefs = self.gui.config.get("handover_report_prefs", {})
        
        # ================= LEFT PANEL (Controls) =================
        left_panel = QWidget()
        left_panel.setFixedWidth(350)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 10, 0)
        
        left_layout.addWidget(QLabel("Timeframe:"))
        self.combo_time = QComboBox()
        self.combo_time.addItems([
            "Current 12h Shift (0600-1800 or 1800-0600)", 
            "Previous 12h Shift", 
            "Current Day (00:00 - 23:59)", 
            "Previous Day",
            "All Time (Entire Log)"
        ])
        
        # Restore saved dropdown index
        self.combo_time.setCurrentIndex(prefs.get("timeframe_index", 0))
        self.combo_time.currentIndexChanged.connect(self.on_preferences_changed)
        left_layout.addWidget(self.combo_time)
        
        # --- Metrics Configuration Group ---
        metrics_group = QGroupBox("Included Metrics & Events")
        metrics_layout = QVBoxLayout(metrics_group)
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setStyleSheet("background: transparent;")
        
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        
        # Standard Checkboxes (Restoring saved state)
        self.chk_dist = QCheckBox("Total Distance Surveyed")
        self.chk_dist.setChecked(prefs.get("chk_dist", True))
        
        self.chk_active_lines = QCheckBox("List of Active Lines")
        self.chk_active_lines.setChecked(prefs.get("chk_active_lines", True))
        
        self.chk_completed_lines = QCheckBox("List of Completed Lines (EOL)")
        self.chk_completed_lines.setChecked(prefs.get("chk_completed_lines", True))
        
        for chk in [self.chk_dist, self.chk_active_lines, self.chk_completed_lines]:
            chk.toggled.connect(self.on_preferences_changed)
            scroll_layout.addWidget(chk)
            
        line1 = QFrame(); line1.setFrameShape(QFrame.Shape.HLine); line1.setStyleSheet("background-color: rgba(128,128,128,0.3); margin: 5px 0;"); scroll_layout.addWidget(line1)
        scroll_layout.addWidget(QLabel("Background System Events:"))

        # Internal Event Checkboxes (Restoring saved state)
        self.system_events = {
            "_SYS_LOG_ON": ("Log ons", "Log on"),
            "_SYS_LOG_OFF": ("Log offs", "Log off"),
            "_SYS_HOURLY": ("Hourly/Manual KP Logs", "Current KP"),
            "_SYS_MIDNIGHT": ("Midnight Logs", "New Day")
        }
        self.sys_checkboxes = {}
        sys_prefs = prefs.get("sys_events", {})
        
        for sys_key, (display_name, search_term) in self.system_events.items():
            chk = QCheckBox(display_name)
            chk.setChecked(sys_prefs.get(sys_key, False)) # Default to false so report isn't cluttered
            chk.toggled.connect(self.on_preferences_changed)
            self.sys_checkboxes[sys_key] = chk
            scroll_layout.addWidget(chk)

        line2 = QFrame(); line2.setFrameShape(QFrame.Shape.HLine); line2.setStyleSheet("background-color: rgba(128,128,128,0.3); margin: 5px 0;"); scroll_layout.addWidget(line2)
        scroll_layout.addWidget(QLabel("Your Custom Event Codes:"))

        # Custom Event Code Checkboxes (Restoring saved state)
        self.code_checkboxes = {}
        event_codes = self.gui.config.get("event_codes", {})
        code_prefs = prefs.get("custom_codes", {})
        
        if not event_codes:
            lbl = QLabel("<i>No custom codes defined.</i>"); lbl.setStyleSheet("color: gray;")
            scroll_layout.addWidget(lbl)
        else:
            for code, desc in event_codes.items():
                chk = QCheckBox(f"[{code}] {desc}")
                chk.setChecked(code_prefs.get(code, True)) # Default true
                chk.toggled.connect(self.on_preferences_changed)
                self.code_checkboxes[code] = chk
                scroll_layout.addWidget(chk)

        scroll_layout.addStretch()
        scroll.setWidget(scroll_widget)
        metrics_layout.addWidget(scroll)
        left_layout.addWidget(metrics_group)
        
        btn_gen = QPushButton("🔄 Refresh Report")
        btn_gen.setObjectName("ActionBtn")
        btn_gen.clicked.connect(self.generate_report)
        left_layout.addWidget(btn_gen)
        
        main_layout.addWidget(left_panel)
        
        # ================= RIGHT PANEL (Text Area) =================
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        
        self.text_out = QTextEdit()
        self.text_out.setReadOnly(True)
        right_layout.addWidget(self.text_out)
        
        bot_lay = QHBoxLayout()
        btn_copy = QPushButton("📋 Copy to Clipboard")
        btn_copy.clicked.connect(self.copy_to_clipboard)
        bot_lay.addStretch()
        bot_lay.addWidget(btn_copy)
        right_layout.addLayout(bot_lay)
        
        main_layout.addWidget(right_panel)
        
        # Auto-generate on open
        self.generate_report()

    def on_preferences_changed(self):
        """Saves the layout preferences to settings.json and instantly regenerates the report."""
        prefs = {
            "timeframe_index": self.combo_time.currentIndex(),
            "chk_dist": self.chk_dist.isChecked(),
            "chk_active_lines": self.chk_active_lines.isChecked(),
            "chk_completed_lines": self.chk_completed_lines.isChecked(),
            "sys_events": {key: chk.isChecked() for key, chk in self.sys_checkboxes.items()},
            "custom_codes": {code: chk.isChecked() for code, chk in self.code_checkboxes.items()}
        }
        
        # Update config in memory and write to the JSON file
        self.gui.config.set("handover_report_prefs", prefs)
        self.gui.config.save()
        
        # Refresh the text view
        self.generate_report()


    def generate_report(self):
        excel_path = self.gui.config.get("log_file_path")
        if not excel_path or not os.path.exists(excel_path):
            self.text_out.setPlainText("Error: Excel log file not found. Please check your settings.")
            return
            
        self.text_out.setPlainText("Reading Excel file... please wait.")
        QApplication.processEvents()
        
        try:
            # Read enough rows to find the header
            df_preview = pd.read_excel(excel_path, nrows=30, header=None)
            header_idx = 0
            for i, row in df_preview.iterrows():
                if any(isinstance(x, str) and 'event' in x.lower() for x in row.values):
                    header_idx = i; break
                    
            df = pd.read_excel(excel_path, header=header_idx)
            df.dropna(how='all', inplace=True)
            
            # Smart Column Identification
            date_col = next((c for c in df.columns if 'local' in str(c).lower() and 'time' in str(c).lower()), None)
            if not date_col:
                date_col = next((c for c in df.columns if any(k in str(c).lower() for k in ['date', 'time', 'utc'])), None)
                
            event_col = next((c for c in df.columns if 'event' in str(c).lower()), None)
            line_col = next((c for c in df.columns if 'line' in str(c).lower()), None)
            kp_col = next((c for c in df.columns if str(c).strip().lower() == 'kp'), None)
            code_col = next((c for c in df.columns if 'code' in str(c).lower()), None)
            
            if not date_col or not event_col:
                self.text_out.setPlainText("Error: Could not identify Date or Event columns in the Excel file.")
                return
                
            # Robust Date Parsing
            def parse_excel_date(val):
                if pd.isna(val): return pd.NaT
                if isinstance(val, (int, float)):
                    return pd.Timestamp('1899-12-30') + pd.to_timedelta(val, unit='D')
                return pd.to_datetime(val, errors='coerce')
                
            df[date_col] = df[date_col].apply(parse_excel_date)
            df.dropna(subset=[date_col], inplace=True)
            
            if df.empty:
                self.text_out.setPlainText("Error: No valid dates could be parsed from the log.")
                return

            if hasattr(df[date_col].dt, 'tz_localize'):
                df[date_col] = df[date_col].dt.tz_localize(None)
            
            # --- True Real-Time Shift Boundaries ---
            now = datetime.datetime.now()
            time_txt = self.combo_time.currentText().lower() 
            
            # Determine current 12h shift start (0600 or 1800)
            if now.hour >= 18: current_shift_start = now.replace(hour=18, minute=0, second=0, microsecond=0)
            elif now.hour >= 6: current_shift_start = now.replace(hour=6, minute=0, second=0, microsecond=0)
            else: current_shift_start = (now - datetime.timedelta(days=1)).replace(hour=18, minute=0, second=0, microsecond=0)

            # Apply exact boundaries
            if "current 12" in time_txt or "current shift" in time_txt:
                start_time = pd.Timestamp(current_shift_start)
                end_time = start_time + pd.Timedelta(hours=12)
            elif "previous 12" in time_txt or "prior" in time_txt:
                start_time = pd.Timestamp(current_shift_start) - pd.Timedelta(hours=12)
                end_time = pd.Timestamp(current_shift_start)
            elif "current day" in time_txt or "24" in time_txt:
                # Exactly 00:00:00 to 23:59:59 of TODAY
                start_time = pd.Timestamp(now).replace(hour=0, minute=0, second=0, microsecond=0)
                end_time = start_time + pd.Timedelta(days=1)
            elif "previous day" in time_txt:
                # Exactly 00:00:00 to 23:59:59 of YESTERDAY
                start_time = (pd.Timestamp(now) - pd.Timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
                end_time = start_time + pd.Timedelta(days=1)
            else:
                start_time = df[date_col].min() if not df.empty else pd.Timestamp(now)
                end_time = pd.Timestamp(now) + pd.Timedelta(days=1)
                
            df_shift = df[(df[date_col] >= start_time) & (df[date_col] < end_time)].copy()
            
            if df_shift.empty:
                self.text_out.setPlainText(f"No events found in the selected timeframe:\n{self.combo_time.currentText()}\nDate Bounds: {start_time.strftime('%Y-%m-%d %H:%M')} to {end_time.strftime('%Y-%m-%d %H:%M')}")
                return
                
            # --- Optional Metric: Macro-Segment Distance Calculation ---
            total_distance = 0.0
            if self.chk_dist.isChecked() and kp_col and line_col:
                df_before = df[df[date_col] < start_time].sort_values(by=date_col).tail(1)
                df_calc = pd.concat([df_before, df_shift.sort_values(by=date_col)])
                active_line, start_kp, last_kp = None, None, None
                
                for idx, row in df_calc.iterrows():
                    # Safely parse values
                    try:
                        kp = float(row[kp_col])
                        if pd.isna(kp): continue
                    except (ValueError, TypeError): continue
                        
                    line = str(row[line_col]).strip()
                    if not line or line.lower() == 'nan': line = "Unknown Line"
                    ev_str = str(row.get(event_col, "")).lower()

                    # Pre-shift anchor
                    if row[date_col] < start_time:
                        if "log off" not in ev_str:
                            active_line, start_kp, last_kp = line, kp, kp
                        continue
                        
                    # --- Processing Rows INSIDE the Shift ---
                    # A. Explicit Log off
                    if "log off" in ev_str:
                        if active_line is not None and start_kp is not None: 
                            total_distance += abs(kp - start_kp)
                        active_line, start_kp, last_kp = None, None, None
                        continue
                        
                    # B. Explicit Log on
                    if "log on" in ev_str:
                        if active_line is not None and start_kp is not None and last_kp is not None: 
                            total_distance += abs(last_kp - start_kp)
                        active_line, start_kp, last_kp = line, kp, kp
                        continue
                        
                    # C. Line Change
                    if active_line is not None and line != active_line:
                        if start_kp is not None and last_kp is not None: 
                            total_distance += abs(last_kp - start_kp)
                        active_line, start_kp, last_kp = line, kp, kp
                        continue
                        
                    # D. Standard ongoing log event
                    if active_line is None: 
                        active_line, start_kp = line, kp
                    last_kp = kp
                        
                # End of Shift
                if active_line is not None and start_kp is not None and last_kp is not None:
                    total_distance += abs(last_kp - start_kp)

            # --- Extract General Metrics ---
            total_entries = len(df_shift)
            
            active_lines_list = []
            if self.chk_active_lines.isChecked() and line_col:
                active_lines_list = [str(x) for x in df_shift[line_col].dropna().unique() if str(x).strip() and str(x).lower() != 'nan']
                
            completed_lines_list = []
            if self.chk_completed_lines.isChecked() and line_col:
                for idx, row in df_shift.iterrows():
                    ev_str = str(row.get(event_col, "")).lower()
                    code_str = str(row.get(code_col, "")).strip().lower() if code_col else ""
                    
                    # Consider a line completed if the user explicitly clicked Log off OR logged an End Of Line code
                    if "log off" in ev_str or "eol" in code_str:
                        ln = str(row.get(line_col, "")).strip()
                        if ln and ln.lower() != 'nan' and ln not in completed_lines_list:
                            completed_lines_list.append(ln)
                
            # --- Flexible Event Counting STRICTLY by Code ---
            event_counts = {}
            target_sys = {key: val for key, val in self.system_events.items() if self.sys_checkboxes[key].isChecked()}
            target_codes = {code: chk for code, chk in self.code_checkboxes.items() if chk.isChecked()}
            
            for idx, row in df_shift.iterrows():
                ev_str = str(row.get(event_col, ""))
                code_str = str(row.get(code_col, "")).strip() if code_col else ""
                
                # Check requested system events (Log on, Log off, Hourly)
                for sys_key, (display_name, search_term) in target_sys.items():
                    if search_term.lower() in ev_str.lower():
                        event_counts[display_name] = event_counts.get(display_name, 0) + 1
                        
                # STRICTLY Check requested custom codes from the Code column
                for code in target_codes.keys():
                    if not code: continue
                    # Match if the Code Column exactly equals the Code, OR if the Code is baked into the Event text
                    if (code_str.lower() == code.lower()) or (f"[{code.upper()}]" in ev_str.upper()):
                        display_name = f"[{code}] {self.gui.config.get('event_codes', {}).get(code, '')}"
                        event_counts[display_name] = event_counts.get(display_name, 0) + 1
                
            # --- Build Text Output ---
            report = f"--- SHIFT HANDOVER REPORT ---\n"
            
            # Format the Date nicely. If shift spans two days, show both dates.
            date_str_start = start_time.strftime('%Y-%m-%d')
            date_str_end = (end_time - pd.Timedelta(seconds=1)).strftime('%Y-%m-%d') # Subtract 1 second so 00:00 falls on the correct day
            if date_str_start == date_str_end:
                report += f"Date: {date_str_start}\n"
            else:
                report += f"Date: {date_str_start} to {date_str_end}\n"
                
            report += f"Timeframe: {self.combo_time.currentText()}\n"
            report += f"Shift Bounds: {start_time.strftime('%H:%M')} to {end_time.strftime('%H:%M')} (Local Time)\n\n"
            
            report += f"📊 LOG SUMMARY:\n"
            report += f"• Total Log Entries: {total_entries}\n"
            
            if self.chk_dist.isChecked() and total_distance > 0:
                report += f"• Exact Distance Surveyed: {total_distance:.3f} km\n"
            
            if self.chk_active_lines.isChecked() and active_lines_list:
                report += f"• Unique Lines Surveyed ({len(active_lines_list)}):\n"
                for ln in active_lines_list:
                    report += f"    ◦ {ln}\n"
                
            if self.chk_completed_lines.isChecked() and completed_lines_list:
                report += f"• Lines Completed / EOL ({len(completed_lines_list)}):\n"
                for ln in completed_lines_list:
                    report += f"    ◦ {ln}\n"
                
            if event_counts:
                report += f"\n🔖 EVENT BREAKDOWN:\n"
                for k, v in sorted(event_counts.items()):
                    report += f"• {k}: {v}\n"
                
            self.text_out.setPlainText(report)
            
        except Exception as e:
            traceback.print_exc()
            self.text_out.setPlainText(f"An error occurred while generating the report:\n\n{str(e)}")

    def copy_to_clipboard(self):
        QApplication.clipboard().setText(self.text_out.toPlainText())
        QMessageBox.information(self, "Copied", "Report copied to clipboard!")

class ButtonEditDialog(QDialog):
    def __init__(self, parent, button_index, config, is_main=False, button_name=""):
        super().__init__(parent)
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)
        self.gui, self.config_data, self.is_main, self.button_name = parent, config, is_main, button_name
        self.setWindowTitle(f"Edit Button: {button_name if is_main else config.get('text', 'Custom')}")
        self.setMinimumWidth(400)
        layout = QVBoxLayout(self); grid = QGridLayout(); layout.addLayout(grid)
        row = 0
        self.edit_text = QLineEdit(config.get("text", ""))
        self.edit_event = QLineEdit(config.get("event_text", ""))
        if is_main:
            self.edit_text.setReadOnly(True)
            if button_name in ["Manual KP Log", "Hourly KP Log"]: self.edit_event.setReadOnly(True)
        grid.addWidget(QLabel("Button Text:"), row, 0); grid.addWidget(self.edit_text, row, 1); row += 1
        grid.addWidget(QLabel("Event Text:"), row, 0); grid.addWidget(self.edit_event, row, 1); row += 1
        
        self.combo_code = QComboBox()
        self.combo_code.addItems([""] + [f"{c} - {d}" for c, d in self.gui.config.get("event_codes", {}).items()])
        current_code = config.get("event_code", "")
        if current_code:
            match_index = next((i for i in range(self.combo_code.count()) if self.combo_code.itemText(i).startswith(f"{current_code} - ")), 0)
            self.combo_code.setCurrentIndex(match_index)
        grid.addWidget(QLabel("Event Code:"), row, 0); grid.addWidget(self.combo_code, row, 1); row += 1

        self.combo_source = QComboBox()
        txt_aliases = self.gui.config.get("txt_source_aliases", {})
        self.combo_source.addItems(["None"] + list(txt_aliases.values()))
        internal_key = self.gui.config.get("hourly_log_txt_source_key") if button_name in ["Manual KP Log", "Hourly KP Log"] else config.get("txt_source_key", "None")
        self.combo_source.setCurrentText(txt_aliases.get(internal_key, "None"))
        grid.addWidget(QLabel("Event Source:"), row, 0); grid.addWidget(self.combo_source, row, 1); row += 1

        if not self.is_main:
            self.combo_group = QComboBox(); self.combo_group.setEditable(True)
            self.combo_group.addItems(self.gui.config.get("custom_button_tab_groups", ["Main"]))
            self.combo_group.setCurrentText(config.get("tab_group", "Main"))
            grid.addWidget(QLabel("Tab Group:"), row, 0); grid.addWidget(self.combo_group, row, 1); row += 1

        bg, fg = self.gui.config.get("button_colors", {}).get(button_name if is_main else config.get("text"), (None, None))
        self.bg_color, self.fg_color = bg, fg
        self.lbl_bg = QLabel("BG Color"); self.lbl_bg.setStyleSheet(f"background-color: {bg or 'none'}; border: 1px solid black;")
        btn_bg = QPushButton("Pick BG"); btn_bg.clicked.connect(lambda: self.pick_color('bg'))
        grid.addWidget(btn_bg, row, 0); grid.addWidget(self.lbl_bg, row, 1); row += 1
        self.lbl_fg = QLabel("Text Color"); self.lbl_fg.setStyleSheet(f"background-color: {fg or 'none'}; border: 1px solid black;")
        btn_fg = QPushButton("Pick FG"); btn_fg.clicked.connect(lambda: self.pick_color('fg'))
        grid.addWidget(btn_fg, row, 0); grid.addWidget(self.lbl_fg, row, 1); row += 1

        bbox = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        bbox.accepted.connect(self.save_and_accept); bbox.rejected.connect(self.reject)
        layout.addWidget(bbox)

    def pick_color(self, target):
        color = QColorDialog.getColor()
        if color.isValid():
            hex_col = color.name()
            if target == 'bg': 
                self.bg_color = hex_col
                self.lbl_bg.setStyleSheet(f"background-color: {hex_col}; border: 1px solid black;")
            else: 
                self.fg_color = hex_col
                self.lbl_fg.setStyleSheet(f"background-color: {hex_col}; border: 1px solid black;")

    def save_and_accept(self):
        old_text = self.config_data.get("text")
        self.config_data["text"] = self.edit_text.text().strip()
        self.config_data["event_text"] = self.edit_event.text().strip()
        code_text = self.combo_code.currentText()
        self.config_data["event_code"] = code_text.split(" - ")[0] if " - " in code_text else code_text
        txt_aliases = self.gui.config.get("txt_source_aliases", {})
        display_to_internal = {v: k for k, v in txt_aliases.items()}
        selected_internal = display_to_internal.get(self.combo_source.currentText(), "None")
        
        if self.button_name in ["Manual KP Log", "Hourly KP Log"]: self.gui.config.set("hourly_log_txt_source_key", selected_internal)
        else: self.config_data["txt_source_key"] = selected_internal

        if not self.is_main:
            new_group = self.combo_group.currentText().strip() or "Main"
            self.config_data["tab_group"] = new_group
            groups = self.gui.config.get("custom_button_tab_groups", ["Main"])
            if new_group not in groups:
                groups.append(new_group)
                self.gui.config.set("custom_button_tab_groups", groups)

        colors = self.gui.config.get("button_colors", {})
        if old_text and old_text in colors and old_text != self.config_data["text"]: del colors[old_text]
        colors[self.button_name if self.is_main else self.config_data["text"]] = (self.bg_color, self.fg_color)
        self.gui.config.set("button_colors", colors)
        self.gui.config.save()
        self.gui.refresh_custom_buttons() 
        if hasattr(self.gui, 'refresh_main_buttons'): self.gui.refresh_main_buttons()
        self.accept()

class TxtMappingDialog(QDialog):
    def __init__(self, parent, source_key):
        super().__init__(parent)
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)
        self.gui, self.source_key = parent, source_key
        self.setWindowTitle(f"Field Mapping: {source_key}")
        self.resize(800, 500)
        self.setStyleSheet("""
            QDialog { background-color: palette(window); }
            QLabel { color: palette(window-text); }
            QLineEdit { background-color: palette(base); color: palette(text); border: 1px solid rgba(128, 128, 128, 0.6); border-radius: 4px; padding: 6px; min-height: 18px; }
            QLineEdit:focus { border: 2px solid #0078D4; padding: 5px; }
            QCheckBox { spacing: 8px; color: palette(window-text); font-weight: bold; }
            QCheckBox::indicator { width: 18px; height: 18px; border: 2px solid palette(text); border-radius: 4px; background-color: transparent; }
            QCheckBox::indicator:hover { border: 2px solid #28a745; }
            QCheckBox::indicator:checked { background-color: #28a745; border: 2px solid #28a745; image: url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>'); }
            QFrame#Card { background-color: rgba(128, 128, 128, 0.08); border: 1px solid rgba(128, 128, 128, 0.3); border-radius: 8px; margin-bottom: 6px; }
        """)
        layout = QVBoxLayout(self)
        controls = QHBoxLayout()
        btn_add = QPushButton("+ Add Field Map"); btn_add.setStyleSheet("QPushButton { background-color: #0078D4; color: white; padding: 6px 15px; border-radius: 4px; font-weight: bold; border: none; } QPushButton:hover { background-color: #106ebe; }"); btn_add.clicked.connect(lambda: self.add_card())
        btn_copy = QPushButton("Copy Main TXT"); btn_copy.setStyleSheet("QPushButton { background-color: rgba(128, 128, 128, 0.15); color: palette(text); padding: 6px 15px; border-radius: 4px; font-weight: bold; border: 1px solid rgba(128, 128, 128, 0.4); }"); btn_copy.clicked.connect(self.copy_main)
        if source_key == "Main TXT": btn_copy.setEnabled(False)
        btn_test = QPushButton("Test Mapping"); btn_test.setStyleSheet("QPushButton { background-color: #28a745; color: white; padding: 6px 15px; border-radius: 4px; font-weight: bold; border: none; } QPushButton:hover { background-color: #218838; }"); btn_test.clicked.connect(self.test_mapping)
        controls.addWidget(btn_add); controls.addWidget(btn_copy); controls.addWidget(btn_test); controls.addStretch()
        layout.addLayout(controls)

        self.scroll = QScrollArea(); self.scroll.setWidgetResizable(True); self.scroll.setFrameShape(QFrame.Shape.NoFrame); self.scroll.setStyleSheet("background: transparent;")
        self.container = QWidget(); self.card_layout = QVBoxLayout(self.container); self.card_layout.setAlignment(Qt.AlignmentFlag.AlignTop); self.card_layout.setSpacing(10)
        self.scroll.setWidget(self.container); layout.addWidget(self.scroll)

        self.input_refs = []
        for c in self.gui.config.get("all_txt_mappings", {}).get(self.source_key, []): self.add_card(c.get("field", ""), c.get("column_name", ""), c.get("skip", False))
        
        bbox = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        bbox.accepted.connect(self.save_and_accept); bbox.rejected.connect(self.reject)
        layout.addWidget(bbox)

    def add_card(self, field="", col="", skip=False):
        card = QFrame(); card.setObjectName("Card"); card_layout = QVBoxLayout(card); card_layout.setContentsMargins(15, 15, 15, 15); card_layout.setSpacing(10)
        top_lay = QHBoxLayout()
        edit_field = QLineEdit(field); edit_field.setPlaceholderText("TXT Raw Field Name")
        edit_col = QLineEdit(col); edit_col.setPlaceholderText("Excel Target Column")
        btn_del = QPushButton("Remove"); btn_del.setCursor(Qt.CursorShape.PointingHandCursor); btn_del.setStyleSheet("QPushButton { color: #FF6B6B; border: 1px solid #FF6B6B; background: transparent; padding: 4px 10px; border-radius: 4px; font-weight: bold; } QPushButton:hover { background-color: rgba(255, 107, 107, 0.15); }")
        top_lay.addWidget(QLabel("<b>TXT Field:</b>")); top_lay.addWidget(edit_field); top_lay.addSpacing(20); top_lay.addWidget(QLabel("<b>Excel Column:</b>")); top_lay.addWidget(edit_col); top_lay.addStretch(); top_lay.addWidget(btn_del)
        card_layout.addLayout(top_lay)
        chk_skip = QCheckBox("Skip this field"); chk_skip.setChecked(skip); chk_skip.setStyleSheet("color: #FF6B6B; font-weight: bold;")
        card_layout.addWidget(chk_skip)
        self.card_layout.addWidget(card)
        refs = {'card': card, 'field': edit_field, 'col': edit_col, 'skip': chk_skip}
        self.input_refs.append(refs)
        btn_del.clicked.connect(lambda checked=False, r=refs: self._remove_card(r))

    def _remove_card(self, refs):
        self.card_layout.removeWidget(refs['card'])
        refs['card'].deleteLater()
        if refs in self.input_refs: self.input_refs.remove(refs)

    def copy_main(self):
        if QMessageBox.question(self, "Confirm", "Overwrite with Main TXT settings?") == QMessageBox.StandardButton.Yes:
            for refs in self.input_refs[:]: self._remove_card(refs)
            for c in self.gui.config.get("all_txt_mappings", {}).get("Main TXT", []): self.add_card(c.get("field", ""), c.get("column_name", ""), c.get("skip", False))

    def test_mapping(self):
        folder_path = self.gui.config.get("txt_folder_paths", {}).get(self.source_key, "")
        if not folder_path or not os.path.exists(folder_path): return QMessageBox.warning(self, "Error", f"Folder path for {self.source_key} is invalid.")
        latest_file, latest_time = None, -1
        for root, _, files in os.walk(folder_path):
            for f in files:
                if f.lower().endswith(('.txt', '.csv', '.npd')):
                    fp = os.path.join(root, f)
                    try:
                        mtime = os.path.getmtime(fp)
                        if mtime > latest_time: latest_time, latest_file = mtime, fp
                    except Exception: pass
        if not latest_file: return QMessageBox.warning(self, "Error", "No valid data files found.")
        try:
            with open(latest_file, 'r') as f:
                lines = [l for l in f.readlines() if l.strip()]
                if not lines: return QMessageBox.warning(self, "Error", "The file is empty.")
                last_line = lines[-1].strip()
            parts = [p.strip() for p in last_line.split(',')]
            result = f"<b>File:</b> {os.path.basename(latest_file)}<br><b>Raw:</b> {last_line}<br><br><b><u>Parsed Mapping:</u></b><br>"
            for i, refs in enumerate(self.input_refs):
                field, col, skip = refs['field'].text().strip(), refs['col'].text().strip() or refs['field'].text().strip(), refs['skip'].isChecked()
                val = parts[i] if i < len(parts) else "<span style='color:red;'>NULL (No Data)</span>"
                if skip: result += f"<span style='color:#FF6B6B;'>[SKIPPED] {field} -> {val}</span><br>"
                else: result += f"<b>{col}</b> = {val}<br>"
            msg = QMessageBox(self); msg.setWindowTitle("Mapping Test Results"); msg.setText(result); msg.exec()
        except Exception as e: QMessageBox.critical(self, "Error", f"Failed to parse file: {e}")

    def save_and_accept(self):
        new_config = []
        for refs in self.input_refs:
            field, col, skip = refs['field'].text().strip(), refs['col'].text().strip(), refs['skip'].isChecked()
            if field: new_config.append({"field": field, "column_name": col or field, "skip": skip})
        maps = self.gui.config.get("all_txt_mappings", {})
        maps[self.source_key] = new_config
        self.gui.config.set("all_txt_mappings", maps)
        self.gui.config.save()
        self.accept()

class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)
        self.gui = parent 
        self.setWindowTitle("System Configuration")
        self.resize(1100, 750)
        
        # --- Check Theme to determine Checkbox Border ---
        is_dark = self.gui.config.get("dark_mode", False)
        cb_border = "#FFFFFF" if is_dark else "#000000"
        
        self.setStyleSheet(f"""
            QDialog {{ background-color: palette(window); }}
            QLabel {{ color: palette(window-text); }}
            QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox {{ background-color: palette(base); color: palette(text); border: 1px solid rgba(128, 128, 128, 0.6); border-radius: 4px; padding: 6px; min-height: 18px; }}
            QLineEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus, QComboBox:focus {{ border: 2px solid #0078D4; padding: 5px; }}
            
            /* --- The fixed Checkbox CSS --- */
            QCheckBox {{ spacing: 8px; color: palette(window-text); }}
            QCheckBox::indicator {{ width: 18px; height: 18px; border: 2px solid {cb_border}; border-radius: 4px; background-color: transparent; }}
            QCheckBox::indicator:hover {{ border: 2px solid #28a745; }}
            QCheckBox::indicator:checked {{ background-color: #28a745; border: 2px solid #28a745; image: url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>'); }}
            
            QFrame#Card {{ background-color: rgba(128, 128, 128, 0.08); border: 1px solid rgba(128, 128, 128, 0.3); border-radius: 8px; margin-bottom: 6px; }}
            QTreeWidget {{ background-color: rgba(128, 128, 128, 0.05); border: 1px solid rgba(128, 128, 128, 0.2); border-radius: 6px; }}
            QTreeWidget::item {{ padding: 6px; }}
            QTreeWidget::item:selected {{ background-color: #0078d4; color: white; border-radius: 4px; }}
        """)
        
        self.main_layout = QHBoxLayout(self)
        self.nav_list = QTreeWidget(); self.nav_list.setFixedWidth(220); self.nav_list.setHeaderHidden(True); self.nav_list.setIndentation(10)
        self.stack = QStackedWidget()
        nav_items = [
            ("📊 Projects", self.build_projects_page), ("📁 File Paths", self.build_file_paths_page),
            ("⚙️ Generated Fields", self.build_generated_fields_page), ("📊 Static Fields", self.build_static_fields_page),
            ("📂 Folder Monitors", self.build_folders_page), ("🔖 Event Codes", self.build_event_codes_page),
            ("⏰ Timed Events", self.build_events_page), ("🌍 Timezone", self.build_timezone_page)
        ]
        for name, func in nav_items:
            item = QTreeWidgetItem([name])
            self.nav_list.addTopLevelItem(item)
            self.stack.addWidget(func())
        self.main_layout.addWidget(self.nav_list)
        right_area = QVBoxLayout(); right_area.addWidget(self.stack)
        self.btn_save_all = QPushButton("SAVE ALL AND CLOSE"); self.btn_save_all.setFixedHeight(45); self.btn_save_all.setStyleSheet("QPushButton { background-color: #0078d4; color: white; font-weight: bold; font-size: 14px; border-radius: 6px; border: none; } QPushButton:hover { background-color: #106ebe; }")
        self.btn_save_all.clicked.connect(self.accept)
        right_area.addWidget(self.btn_save_all)
        self.main_layout.addLayout(right_area)
        self.nav_list.itemClicked.connect(lambda it: self.stack.setCurrentIndex(self.nav_list.indexOfTopLevelItem(it)))
        self.load_settings_into_ui()

    def build_projects_page(self):
        page = QWidget(); layout = QVBoxLayout(page)
        title = QLabel("<b>Project Management</b>"); title.setStyleSheet("font-size: 18px;")
        layout.addWidget(title); layout.addWidget(QLabel("Manage your configuration profiles. Different vessels or jobs can have their own project files."))
        curr_frame = QFrame(); curr_frame.setObjectName("Card"); curr_lay = QVBoxLayout(curr_frame)
        curr_lay.addWidget(QLabel("<b>Active Project File:</b>"))
        self.lbl_active_project = QLineEdit(self.gui.config.current_project_path)
        self.lbl_active_project.setReadOnly(True); self.lbl_active_project.setStyleSheet("background-color: rgba(0, 120, 212, 0.1); color: #0078D4; font-weight: bold; border: 1px solid #0078D4;")
        curr_lay.addWidget(self.lbl_active_project); layout.addWidget(curr_frame)
        btn_lay = QGridLayout(); btn_lay.setSpacing(10)
        btn_load = QPushButton("📂 Load Existing Project"); btn_load.clicked.connect(self.load_project)
        btn_save_as = QPushButton("📝 Save As New Project"); btn_save_as.clicked.connect(self.save_project_as)
        btn_template = QPushButton("⚠️ Restore Blank Template"); btn_template.setStyleSheet("QPushButton { color: #FF6B6B; border: 1px solid #FF6B6B; background: transparent; } QPushButton:hover { background-color: rgba(255, 107, 107, 0.15); }"); btn_template.clicked.connect(self.load_blank_project)
        btn_lay.addWidget(btn_load, 0, 0); btn_lay.addWidget(btn_save_as, 0, 1); btn_lay.addWidget(btn_template, 1, 0, 1, 2)
        layout.addLayout(btn_lay)
        sum_frame = QFrame(); sum_frame.setObjectName("Card"); sum_lay = QVBoxLayout(sum_frame)
        sum_lay.addWidget(QLabel("<b>Project Summary:</b>"))
        self.lbl_sum_excel = QLabel("Excel Log: Loading..."); self.lbl_sum_folders = QLabel("Active Folders: Loading..."); self.lbl_sum_buttons = QLabel("Custom Buttons: Loading...")
        sum_lay.addWidget(self.lbl_sum_excel); sum_lay.addWidget(self.lbl_sum_folders); sum_lay.addWidget(self.lbl_sum_buttons)
        layout.addWidget(sum_frame); layout.addStretch(); self.refresh_project_summary()
        return page

    def refresh_project_summary(self):
        if hasattr(self, 'lbl_sum_excel'):
            excel = self.gui.config.get('log_file_path', '')
            self.lbl_sum_excel.setText(f"Excel Log: {os.path.basename(excel) if excel else 'Not Set'}")
            folders = [k for k, v in self.gui.config.get('folder_paths', {}).items() if v and not self.gui.config.get('folder_skips', {}).get(k, False)]
            self.lbl_sum_folders.setText(f"Active Folder Monitors: {len(folders)}")
            self.lbl_sum_buttons.setText(f"Custom Buttons: {self.gui.config.get('num_custom_buttons', 0)}")

    def load_project(self):
        initial_dir = os.path.join(os.getcwd(), "settings")
        path, _ = QFileDialog.getOpenFileName(self, "Select Project", initial_dir, "JSON files (*.json);;All files (*.*)")
        if path:
            success, msg = self.gui.config.load_project(path)
            if success:
                self.lbl_active_project.setText(path)
                self.load_settings_into_ui()
                self.refresh_project_summary()
                QMessageBox.information(self, "Success", f"Project loaded successfully:\n{os.path.basename(path)}")
            else:
                QMessageBox.critical(self, "Load Error", msg)

    def save_project_as(self):
        initial_dir = os.path.join(os.getcwd(), "settings")
        name, ok = QInputDialog.getText(self, "Save Project As", "Enter a new name for this project:")
        if ok and name.strip():
            name = name.strip()
            if not name.endswith('.json'): name += '.json'
            new_path = os.path.join(initial_dir, name)
            
            # Store old path in case of failure
            old_path = self.gui.config.current_project_path
            self.gui.config.current_project_path = new_path
            
            success, msg = self.gui.config.save()
            if success:
                self.accept()
                self.gui.setWindowTitle(f"Online Logger - {name}")
                QMessageBox.information(self, "Success", f"Project saved as '{name}'.")
            else:
                self.gui.config.current_project_path = old_path # Revert
                QMessageBox.critical(self, "Save Error", msg)

    def load_blank_project(self):
        template_path = os.path.join(os.getcwd(), PROJECT_TEMPLATE_FILE)
        if not os.path.exists(template_path): return QMessageBox.critical(self, "Error", f"Blank template not found at:\n{template_path}")
        if QMessageBox.question(self, "Confirm Reset", "Are you sure you want to load the blank template?") == QMessageBox.StandardButton.Yes:
            success, msg = self.gui.config.load_project(template_path)
            if success:
                self.lbl_active_project.setText("BLANK TEMPLATE (Save As to create new project)")
                self.load_settings_into_ui()
                self.refresh_project_summary()
            else:
                QMessageBox.critical(self, "Error", msg)

    def build_file_paths_page(self):
        page = QWidget(); layout = QVBoxLayout(page)
        scroll = QScrollArea(); scroll.setWidgetResizable(True); scroll.setFrameShape(QFrame.Shape.NoFrame); scroll.setStyleSheet("background: transparent;")
        inner = QWidget(); inner_lay = QVBoxLayout(inner)
        browse_style = "QPushButton { background-color: rgba(128, 128, 128, 0.05); color: palette(text); border: 1px solid rgba(128, 128, 128, 0.3); border-radius: 4px; padding: 6px 12px; font-weight: bold; } QPushButton:hover { background-color: rgba(0, 120, 212, 0.1); border-color: #0078D4; color: #0078D4; }"
        map_style = "QPushButton { background-color: rgba(0, 120, 212, 0.05); color: #0078D4; border: 1px solid #0078D4; border-radius: 4px; padding: 6px 12px; font-weight: bold; } QPushButton:hover { background-color: #0078D4; color: white; }"
        grp1 = QGroupBox("Master Logging Files"); glay1 = QGridLayout(grp1); glay1.setSpacing(10)
        self.edit_log_path = QLineEdit(); btn_l = QPushButton("Browse..."); btn_l.setStyleSheet(browse_style); btn_l.clicked.connect(lambda: self.browse_file(self.edit_log_path, "Excel (*.xlsx *.xlsb)"))
        glay1.addWidget(QLabel("Excel Log Path:"), 0, 0); glay1.addWidget(self.edit_log_path, 0, 1); glay1.addWidget(btn_l, 0, 2)
        self.edit_db_path = QLineEdit(); btn_d = QPushButton("Browse..."); btn_d.setStyleSheet(browse_style); btn_d.clicked.connect(lambda: self.browse_file(self.edit_db_path, "Database (*.db)", save=True))
        glay1.addWidget(QLabel("SQLite Mirror:"), 1, 0); glay1.addWidget(self.edit_db_path, 1, 1); glay1.addWidget(btn_d, 1, 2)
        inner_lay.addWidget(grp1)
        grp2 = QGroupBox("Navigation Data Sources"); glay2 = QGridLayout(grp2); glay2.setSpacing(10)
        self.txt_source_widgets = {}
        row_idx = 0
        for i, key in enumerate(TXT_FILES_KEYS[1:]):
            a_w = QLineEdit(); p_w = QLineEdit()
            btn_p = QPushButton("Browse Folder..."); btn_p.setStyleSheet(browse_style); btn_p.setFixedWidth(130); btn_p.clicked.connect(lambda checked=False, e=p_w: self.browse_folder(e))
            btn_m = QPushButton("⚙️ Field Mapping"); btn_m.setStyleSheet(map_style); btn_m.clicked.connect(lambda checked=False, k=key: TxtMappingDialog(self.gui, k).exec())
            glay2.addWidget(QLabel(f"<b>{key}</b> Name:"), row_idx, 0); glay2.addWidget(a_w, row_idx, 1); glay2.addWidget(btn_m, row_idx, 2)
            glay2.addWidget(QLabel("Folder Path:"), row_idx+1, 0); glay2.addWidget(p_w, row_idx+1, 1); glay2.addWidget(btn_p, row_idx+1, 2)
            self.txt_source_widgets[key] = (a_w, p_w)
            row_idx += 2
            if i < len(TXT_FILES_KEYS[1:]) - 1:
                separator = QFrame(); separator.setFrameShape(QFrame.Shape.HLine); separator.setStyleSheet("background-color: rgba(128,128,128,0.2); margin: 8px 0px;")
                glay2.addWidget(separator, row_idx, 0, 1, 3); row_idx += 1
        inner_lay.addWidget(grp2); inner_lay.addStretch(); scroll.setWidget(inner); layout.addWidget(scroll)
        return page

    def build_generated_fields_page(self):
        page = QWidget(); layout = QVBoxLayout(page)
        title = QLabel("<b>Generated Data Mapping</b>"); title.setStyleSheet("font-size: 18px;")
        layout.addWidget(title); layout.addWidget(QLabel("Assign Excel column names to data generated natively by the application."))
        self.gen_scroll = QScrollArea(); self.gen_scroll.setWidgetResizable(True); self.gen_scroll.setFrameShape(QFrame.Shape.NoFrame); self.gen_scroll.setStyleSheet("background: transparent;")
        self.gen_container = QWidget(); self.gen_layout = QVBoxLayout(self.gen_container); self.gen_layout.setAlignment(Qt.AlignmentFlag.AlignTop); self.gen_layout.setSpacing(12)
        self.gen_scroll.setWidget(self.gen_container); layout.addWidget(self.gen_scroll)
        self.generated_input_refs = []
        return page

    def build_static_fields_page(self):
        page = QWidget(); layout = QVBoxLayout(page)
        header_lay = QHBoxLayout()
        title = QLabel("<b>Static Excel Lookups</b>"); title.setStyleSheet("font-size: 18px;")
        header_lay.addWidget(title); header_lay.addStretch()
        btn_add = QPushButton("+ Add Static Field"); btn_add.clicked.connect(lambda: self.add_static_card())
        header_lay.addWidget(btn_add); layout.addLayout(header_lay)
        self.static_scroll = QScrollArea(); self.static_scroll.setWidgetResizable(True); self.static_scroll.setFrameShape(QFrame.Shape.NoFrame); self.static_scroll.setStyleSheet("background: transparent;")
        self.static_container = QWidget(); self.static_layout = QVBoxLayout(self.static_container); self.static_layout.setAlignment(Qt.AlignmentFlag.AlignTop); self.static_layout.setSpacing(12)
        self.static_scroll.setWidget(self.static_container); layout.addWidget(self.static_scroll)
        self.static_input_refs = []
        return page

    def build_folders_page(self):
        page = QWidget(); layout = QVBoxLayout(page)
        header_lay = QHBoxLayout()
        title = QLabel("<b>Monitored Folders</b>"); title.setStyleSheet("font-size: 18px;")
        header_lay.addWidget(title); header_lay.addStretch()
        btn_add = QPushButton("+ Add Folder"); btn_add.clicked.connect(lambda: self.add_folder_row("NewFolder"))
        header_lay.addWidget(btn_add); layout.addLayout(header_lay)
        self.folders_scroll = QScrollArea(); self.folders_scroll.setWidgetResizable(True); self.folders_scroll.setFrameShape(QFrame.Shape.NoFrame); self.folders_scroll.setStyleSheet("background: transparent;")
        self.folders_container = QWidget(); self.folders_layout = QVBoxLayout(self.folders_container); self.folders_layout.setAlignment(Qt.AlignmentFlag.AlignTop); self.folders_layout.setSpacing(12)
        self.folders_scroll.setWidget(self.folders_container); layout.addWidget(self.folders_scroll)
        self.folder_input_refs = [] 
        return page

    def build_event_codes_page(self):
        page = QWidget(); layout = QVBoxLayout(page)
        header_lay = QHBoxLayout()
        title = QLabel("<b>Event Codes</b>"); title.setStyleSheet("font-size: 18px;")
        header_lay.addWidget(title); header_lay.addStretch()
        btn_add = QPushButton("+ Add Code"); btn_add.setStyleSheet("QPushButton { background-color: #0078D4; color: white; padding: 6px 15px; border-radius: 4px; font-weight: bold; border: none; } QPushButton:hover { background-color: #106ebe; }"); btn_add.clicked.connect(lambda: self.add_event_code_row())
        header_lay.addWidget(btn_add); layout.addLayout(header_lay)
        self.codes_table = QTableWidget(0, 2); self.codes_table.setHorizontalHeaderLabels(["Code", "Description"]); self.codes_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch); self.codes_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows); self.codes_table.verticalHeader().setVisible(False)
        btn_remove = QPushButton("Remove Selected"); btn_remove.setStyleSheet("QPushButton { color: #FF6B6B; border: 1px solid #FF6B6B; background: transparent; padding: 6px 15px; border-radius: 4px; font-weight: bold; } QPushButton:hover { background-color: rgba(255, 107, 107, 0.15); }"); btn_remove.clicked.connect(self.remove_event_code_row)
        layout.addWidget(self.codes_table); bot_lay = QHBoxLayout(); bot_lay.addStretch(); bot_lay.addWidget(btn_remove); layout.addLayout(bot_lay)
        return page

    def build_events_page(self):
        page = QWidget(); layout = QVBoxLayout(page)
        std_frame = QFrame(); std_frame.setObjectName("Card"); std_lay = QVBoxLayout(std_frame)
        
        # --- New Day Layout with Code Dropdown ---
        nd_lay = QHBoxLayout()
        self.chk_new_day = QCheckBox("Enable Midnight 'New Day' Event")
        self.combo_new_day_code = QComboBox()
        self.combo_new_day_code.addItems([""] + [f"{c} - {d}" for c, d in self.gui.config.get('event_codes', {}).items()])
        nd_lay.addWidget(self.chk_new_day); nd_lay.addWidget(QLabel("  Code:")); nd_lay.addWidget(self.combo_new_day_code); nd_lay.addStretch()
        
        self.chk_hourly = QCheckBox("Enable Hourly KP Log Event")
        self.chk_logoff = QCheckBox("Calculate distance/speed on Log off")
        
        thresh_layout = QHBoxLayout(); self.spin_threshold = QSpinBox(); self.spin_threshold.setRange(0, 3600)
        thresh_layout.addWidget(QLabel("Active Logging Threshold (secs):")); thresh_layout.addWidget(self.spin_threshold); thresh_layout.addStretch()
        
        std_lay.addLayout(nd_lay); std_lay.addWidget(self.chk_hourly); std_lay.addWidget(self.chk_logoff); std_lay.addLayout(thresh_layout); layout.addWidget(std_frame)
        
        udp_frame = QFrame(); udp_frame.setObjectName("Card"); udp_lay = QVBoxLayout(udp_frame)
        udp_lay.addWidget(QLabel("<b>UDP Logging Automation</b>")); self.chk_udp_enabled = QCheckBox("Enable UDP Automation")
        port_lay = QHBoxLayout(); self.spin_udp_port = QSpinBox(); self.spin_udp_port.setRange(1024, 65535)
        port_lay.addWidget(QLabel("<b>Listening Port:</b>")); port_lay.addWidget(self.spin_udp_port); port_lay.addStretch()
        payload_lay = QHBoxLayout()
        self.edit_udp_rec = QLineEdit(); self.edit_udp_rec.setPlaceholderText("e.g. RECORDING")
        self.edit_udp_idle = QLineEdit(); self.edit_udp_idle.setPlaceholderText("e.g. IDLE")
        payload_lay.addWidget(QLabel("Recording Payload:")); payload_lay.addWidget(self.edit_udp_rec)
        payload_lay.addSpacing(15)
        payload_lay.addWidget(QLabel("Idle Payload:")); payload_lay.addWidget(self.edit_udp_idle)
        
        udp_lay.addWidget(self.chk_udp_enabled); udp_lay.addLayout(port_lay); udp_lay.addLayout(payload_lay); layout.addWidget(udp_frame); layout.addStretch()
        return page

    def build_timezone_page(self):
        page = QWidget(); layout = QVBoxLayout(page)
        self.spin_tz = QDoubleSpinBox(); self.spin_tz.setRange(-14, 14); self.spin_tz.setSingleStep(0.5)
        layout.addWidget(QLabel("UTC Offset (hours):")); layout.addWidget(self.spin_tz); layout.addStretch()
        return page

    def browse_file(self, line_edit, filter_str, save=False):
        if save: path, _ = QFileDialog.getSaveFileName(self, "Save", "", filter_str)
        else: path, _ = QFileDialog.getOpenFileName(self, "Open", "", filter_str)
        if path: line_edit.setText(path)

    def browse_folder(self, line_edit):
        path = QFileDialog.getExistingDirectory(self, "Select Folder")
        if path: line_edit.setText(path)

    def add_generated_card(self, field="", source="", col=""):
        card = QFrame(); card.setObjectName("Card"); card_layout = QVBoxLayout(card); card_layout.setContentsMargins(0,0,0,0); card_layout.setSpacing(0)
        header_widget = QFrame(); header_widget.setStyleSheet("QFrame#CardHeader { background-color: rgba(128, 128, 128, 0.15); border-bottom: 1px solid rgba(128, 128, 128, 0.4); border-radius: 6px; border-bottom-left-radius: 0px; border-bottom-right-radius: 0px; }")
        header_lay = QHBoxLayout(header_widget); header_lay.setContentsMargins(15, 8, 15, 8); header_lay.addWidget(QLabel(f"<b>System Field:</b> {field}")); card_layout.addWidget(header_widget)
        content_widget = QWidget(); content_layout = QHBoxLayout(content_widget); content_layout.setContentsMargins(15, 15, 15, 15)
        edit_col = QLineEdit(col); edit_col.setPlaceholderText("Excel Column Name")
        content_layout.addWidget(QLabel(f"<b>Internal Source:</b> {source}")); content_layout.addSpacing(30); content_layout.addWidget(QLabel("<b>Target Excel Column:</b>")); content_layout.addWidget(edit_col); content_layout.addStretch(); card_layout.addWidget(content_widget)
        self.gen_layout.addWidget(card); self.generated_input_refs.append({'field': field, 'source': source, 'col': edit_col})

    def add_static_card(self, field="", desc="", cell=""):
        card = QFrame(); card.setObjectName("Card"); card_layout = QVBoxLayout(card); card_layout.setContentsMargins(0,0,0,0); card_layout.setSpacing(0)
        header_widget = QFrame(); header_lay = QHBoxLayout(header_widget); header_lay.setContentsMargins(10, 8, 10, 8)
        btn_toggle = QPushButton("▼"); btn_toggle.setFixedSize(32, 32); btn_toggle.setCheckable(True); btn_toggle.setChecked(True); btn_toggle.setStyleSheet("QPushButton { border: none; font-size: 16px; color: palette(text); background: transparent; } QPushButton:hover { color: #0078d4; background-color: rgba(128, 128, 128, 0.15); border-radius: 16px; }")
        edit_field = QLineEdit(field); edit_field.setPlaceholderText("Excel Target Column"); edit_field.setMinimumWidth(200)
        btn_del = QPushButton("Remove"); btn_del.setStyleSheet("QPushButton { color: #FF6B6B; border: 1px solid #FF6B6B; background: transparent; padding: 5px 15px; border-radius: 4px; font-weight: bold; } QPushButton:hover { background-color: rgba(255, 107, 107, 0.15); }")
        header_lay.addWidget(btn_toggle); header_lay.addWidget(QLabel("<b>Excel Column:</b>")); header_lay.addWidget(edit_field); header_lay.addStretch(); header_lay.addWidget(btn_del); card_layout.addWidget(header_widget)
        content_widget = QWidget(); content_layout = QHBoxLayout(content_widget); content_layout.setContentsMargins(45, 15, 15, 15)
        edit_desc = QLineEdit(desc); edit_desc.setPlaceholderText("Description"); edit_cell = QLineEdit(cell); edit_cell.setPlaceholderText("='Sheet'!A1")
        content_layout.addWidget(QLabel("<b>Description:</b>")); content_layout.addWidget(edit_desc); content_layout.addSpacing(20); content_layout.addWidget(QLabel("<b>Excel Cell Reference:</b>")); content_layout.addWidget(edit_cell); card_layout.addWidget(content_widget)
        def toggle_content(checked):
            content_widget.setVisible(checked); btn_toggle.setText("▼" if checked else "▶")
            header_widget.setStyleSheet("QFrame { background-color: rgba(128, 128, 128, 0.15); border-bottom: 1px solid rgba(128, 128, 128, 0.4); }" if checked else "QFrame { background-color: rgba(128, 128, 128, 0.08); border-bottom: none; }")
        btn_toggle.toggled.connect(toggle_content); toggle_content(True); self.static_layout.addWidget(card)
        refs = {'card': card, 'field': edit_field, 'desc': edit_desc, 'cell': edit_cell}; self.static_input_refs.append(refs); btn_del.clicked.connect(lambda checked=False, r=refs: self._remove_static_card(r))

    def _remove_static_card(self, refs):
        self.static_layout.removeWidget(refs['card']); refs['card'].deleteLater()
        if refs in self.static_input_refs: self.static_input_refs.remove(refs)

    def add_folder_row(self, name="", path="", col="", ext="", skip=False, log_x=False, log_ext=False):
        card = QFrame(); card.setObjectName("Card"); card_layout = QVBoxLayout(card); card_layout.setContentsMargins(0,0,0,0); card_layout.setSpacing(0)
        header_widget = QFrame(); header_lay = QHBoxLayout(header_widget); header_lay.setContentsMargins(10, 8, 10, 8)

        btn_toggle = QPushButton("▶"); btn_toggle.setFixedSize(32, 32); btn_toggle.setCheckable(True); btn_toggle.setStyleSheet("QPushButton { border: none; font-size: 16px; color: palette(text); background: transparent; } QPushButton:hover { color: #0078d4; background-color: rgba(128, 128, 128, 0.15); border-radius: 16px; }")
        edit_name = QLineEdit(name); edit_name.setPlaceholderText("Folder Name..."); edit_name.setMinimumWidth(160)
        edit_col = QLineEdit(col); edit_col.setPlaceholderText("Target Column..."); edit_col.setMinimumWidth(160)
        btn_del = QPushButton("Remove"); btn_del.setStyleSheet("QPushButton { color: #FF6B6B; border: 1px solid #FF6B6B; background: transparent; padding: 4px 12px; border-radius: 4px; font-weight: bold; } QPushButton:hover { background-color: rgba(255, 107, 107, 0.15); }")
        header_lay.addWidget(btn_toggle); header_lay.addWidget(QLabel("<b>Name:</b>")); header_lay.addWidget(edit_name); header_lay.addSpacing(20); header_lay.addWidget(QLabel("<b>Excel Column:</b>")); header_lay.addWidget(edit_col); header_lay.addStretch(); header_lay.addWidget(btn_del); card_layout.addWidget(header_widget)

        content_widget = QWidget(); content_layout = QVBoxLayout(content_widget); content_layout.setContentsMargins(45, 15, 15, 15); content_layout.setSpacing(12)
        path_lay = QHBoxLayout(); edit_path = QLineEdit(path)
        
        btn_browse = QPushButton("Browse...")
        btn_browse.setStyleSheet("QPushButton { background-color: rgba(128, 128, 128, 0.05); color: palette(text); border: 1px solid rgba(128, 128, 128, 0.3); border-radius: 4px; padding: 6px 12px; font-weight: bold; } QPushButton:hover { background-color: rgba(0, 120, 212, 0.1); border-color: #0078D4; color: #0078D4; }")
        btn_browse.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_browse.clicked.connect(lambda checked=False, e=edit_path: self.browse_folder(e))
        
        path_lay.addWidget(QLabel("<b>Directory Path:</b>")); path_lay.addWidget(edit_path); path_lay.addWidget(btn_browse); content_layout.addLayout(path_lay)
        
        target_lay = QHBoxLayout(); edit_ext = QLineEdit(ext); edit_ext.setFixedWidth(100); target_lay.addWidget(QLabel("<b>File Ext Filter:</b>")); target_lay.addWidget(edit_ext); target_lay.addStretch(); content_layout.addLayout(target_lay)
        chk_group = QFrame(); chk_group.setStyleSheet("QFrame { border-top: 1px solid rgba(128,128,128,0.2); margin-top: 5px; padding-top: 10px; }"); chk_lay = QHBoxLayout(chk_group); chk_lay.setContentsMargins(0, 0, 0, 0)
        chk_skip = QCheckBox("Skip Monitoring"); chk_skip.setChecked(skip); chk_skip.setStyleSheet("color: #FF6B6B; font-weight: bold;")
        chk_log_x = QCheckBox("Log 'X' instead of filename"); chk_log_x.setChecked(log_x)
        chk_log_ext = QCheckBox("Include extension in log"); chk_log_ext.setChecked(log_ext)
        chk_lay.addWidget(chk_skip); chk_lay.addSpacing(20); chk_lay.addWidget(chk_log_x); chk_lay.addSpacing(20); chk_lay.addWidget(chk_log_ext); chk_lay.addStretch(); content_layout.addWidget(chk_group); card_layout.addWidget(content_widget)
        content_widget.setVisible(False)
        def toggle_content(checked):
            content_widget.setVisible(checked); btn_toggle.setText("▼" if checked else "▶")
        btn_toggle.toggled.connect(toggle_content)
        if name == "NewFolder": btn_toggle.setChecked(True)
        self.folders_layout.addWidget(card)
        card_refs = {'card': card, 'name': edit_name, 'path': edit_path, 'col': edit_col, 'ext': edit_ext, 'skip': chk_skip, 'log_x': chk_log_x, 'log_ext': chk_log_ext}
        if not hasattr(self, 'folder_input_refs'): self.folder_input_refs = []
        self.folder_input_refs.append(card_refs); btn_del.clicked.connect(lambda checked=False, r=card_refs: self._remove_folder_card(r))

    def _remove_folder_card(self, refs):
        self.folders_layout.removeWidget(refs['card']); refs['card'].deleteLater()
        if refs in self.folder_input_refs: self.folder_input_refs.remove(refs)

    def add_event_code_row(self, code="", desc=""):
        row = self.codes_table.rowCount()
        self.codes_table.insertRow(row)
        self.codes_table.setItem(row, 0, QTableWidgetItem(code))
        self.codes_table.setItem(row, 1, QTableWidgetItem(desc))

    def remove_event_code_row(self):
        current_row = self.codes_table.currentRow()
        if current_row >= 0: self.codes_table.removeRow(current_row)

    def load_settings_into_ui(self):
        try:
            c = self.gui.config
            if hasattr(self, 'edit_log_path'): self.edit_log_path.setText(c.get('log_file_path', ""))
            if hasattr(self, 'edit_db_path'): self.edit_db_path.setText(c.get('sqlite_db_path', ""))
            txt_aliases = c.get('txt_source_aliases', {})
            txt_paths = c.get('txt_folder_paths', {})
            if hasattr(self, 'txt_source_widgets'):
                for k, (a_w, p_w) in self.txt_source_widgets.items():
                    a_w.setText(txt_aliases.get(k, k))
                    p_w.setText(txt_paths.get(k, ""))
            if hasattr(self, 'gen_layout'):
                for i in reversed(range(self.gen_layout.count())):
                    w = self.gen_layout.itemAt(i).widget()
                    if w: w.deleteLater()
                self.generated_input_refs.clear()
                for config in c.get('generated_fields_config', []): self.add_generated_card(config.get("field", ""), config.get("source", ""), config.get("column_name", ""))
            if hasattr(self, 'static_layout'):
                for i in reversed(range(self.static_layout.count())):
                    w = self.static_layout.itemAt(i).widget()
                    if w: w.deleteLater()
                self.static_input_refs.clear()
                for config in c.get('static_field_configs', []): self.add_static_card(config.get("field", ""), config.get("description", ""), config.get("column_name", ""))
            if hasattr(self, 'folders_layout'):
                for refs in getattr(self, 'folder_input_refs', [])[:]: self._remove_folder_card(refs)
                
                f_paths = c.get('folder_paths', {})
                # Only inject defaults if this is a brand new project, otherwise use EXACTLY what is saved
                if not f_paths and not c.get('_folders_initialized', False):
                    all_folders = list(DEFAULT_MONITORED_FOLDERS)
                    c.set('_folders_initialized', True)
                else:
                    all_folders = list(f_paths.keys())
                    
                f_cols = c.get('folder_columns', {}); f_exts = c.get('file_extensions', {}); f_skips = c.get('folder_skips', {})
                f_logx = c.get('folder_log_x_instead', {}); f_logext = c.get('folder_log_ext_vars', {})
                for name in all_folders:
                    self.add_folder_row(name, f_paths.get(name, ""), f_cols.get(name, name.replace(" ", "_")), f_exts.get(name, ""), f_skips.get(name, False), f_logx.get(name, False), f_logext.get(name, False))

            if hasattr(self, 'codes_table'):
                self.codes_table.setRowCount(0)
                for code, desc in c.get('event_codes', {}).items(): self.add_event_code_row(code, desc)
            if hasattr(self, 'chk_new_day'): 
                self.chk_new_day.setChecked(c.get('new_day_event_enabled', True))
                if hasattr(self, 'combo_new_day_code'):
                    code = c.get('new_day_event_code', "")
                    if code:
                        match_idx = next((i for i in range(self.combo_new_day_code.count()) if self.combo_new_day_code.itemText(i).startswith(f"{code} - ")), 0)
                        self.combo_new_day_code.setCurrentIndex(match_idx)
            if hasattr(self, 'chk_hourly'): self.chk_hourly.setChecked(c.get('hourly_event_enabled', True))
            if hasattr(self, 'chk_logoff'): self.chk_logoff.setChecked(c.get('calculate_logoff_values', True))
            if hasattr(self, 'spin_tz'): self.spin_tz.setValue(c.get('time_offset_hours', 0.0))
            if hasattr(self, 'chk_udp_enabled'): self.chk_udp_enabled.setChecked(c.get('udp_trigger_enabled', False))
            if hasattr(self, 'spin_udp_port'): self.spin_udp_port.setValue(c.get('udp_trigger_port', 5999))
            if hasattr(self, 'edit_udp_rec'): self.edit_udp_rec.setText(c.get('udp_payload_recording', 'RECORDING'))
            if hasattr(self, 'edit_udp_idle'): self.edit_udp_idle.setText(c.get('udp_payload_idle', 'IDLE'))
            if hasattr(self, 'spin_threshold'): self.spin_threshold.setValue(min(int(c.get('active_logging_threshold_seconds', 15)), 3600))
        except Exception: traceback.print_exc()

    def save_settings(self):
        # Prevent redundant saves if multiple close triggers fire at once
        if getattr(self, '_is_saved', False): return True
        
        try:
            c = self.gui.config
            c.set('log_file_path', self.edit_log_path.text().strip())
            c.set('sqlite_db_path', self.edit_db_path.text().strip())
            
            txt_aliases, txt_paths = {}, {}
            for k, (a_w, p_w) in self.txt_source_widgets.items():
                txt_aliases[k] = a_w.text().strip()
                txt_paths[k] = p_w.text().strip()
            c.set('txt_source_aliases', txt_aliases)
            c.set('txt_folder_paths', txt_paths)
            
            new_gen = [{"field": ref['field'], "source": ref['source'], "column_name": ref['col'].text().strip()} for ref in getattr(self, 'generated_input_refs', [])]
            c.set('generated_fields_config', new_gen)
            
            new_static = [{"field": ref['field'].text().strip(), "description": ref['desc'].text().strip(), "column_name": ref['cell'].text().strip(), "skip": False} for ref in getattr(self, 'static_input_refs', []) if ref['field'].text().strip()]
            c.set('static_field_configs', new_static)
            
            f_paths, f_cols, f_exts, f_skips, f_logx, f_logext = {}, {}, {}, {}, {}, {}
            for refs in getattr(self, 'folder_input_refs', []):
                name = refs['name'].text().strip()
                path = refs['path'].text().strip()
                # Save the folder as long as it has a name. It's okay if the path is left blank!
                if name:
                    f_paths[name] = path
                    f_cols[name] = refs['col'].text().strip() or name
                    f_exts[name] = refs['ext'].text().strip()
                    f_skips[name] = refs['skip'].isChecked()
                    f_logx[name] = refs['log_x'].isChecked()
                    f_logext[name] = refs['log_ext'].isChecked()
            c.set('folder_paths', f_paths)
            c.set('_folders_initialized', True) 
            c.set('folder_columns', f_cols)
            c.set('file_extensions', f_exts)
            c.set('folder_skips', f_skips)
            c.set('folder_log_x_instead', f_logx)
            c.set('folder_log_ext_vars', f_logext)
            
            c.set('time_offset_hours', self.spin_tz.value())
            c.set('active_logging_threshold_seconds', self.spin_threshold.value())
            c.set('new_day_event_enabled', self.chk_new_day.isChecked())
            c.set('hourly_event_enabled', self.chk_hourly.isChecked())
            c.set('calculate_logoff_values', self.chk_logoff.isChecked())
            
            if hasattr(self, 'combo_new_day_code'):
                code_text = self.combo_new_day_code.currentText()
                c.set('new_day_event_code', code_text.split(" - ")[0] if " - " in code_text else code_text)
                
            if hasattr(self, 'codes_table'):
                new_codes = {}
                for row in range(self.codes_table.rowCount()):
                    c_item, d_item = self.codes_table.item(row, 0), self.codes_table.item(row, 1)
                    if c_item and c_item.text().strip(): 
                        new_codes[c_item.text().strip()] = d_item.text().strip() if d_item else ""
                c.set('event_codes', new_codes)
                
            if hasattr(self, 'chk_udp_enabled'):
                c.set('udp_trigger_enabled', self.chk_udp_enabled.isChecked())
                c.set('udp_trigger_port', self.spin_udp_port.value())
                c.set('udp_payload_recording', self.edit_udp_rec.text().strip() or "RECORDING")
                c.set('udp_payload_idle', self.edit_udp_idle.text().strip() or "IDLE")
                if hasattr(self.gui, 'restart_udp_listener'): 
                    self.gui.restart_udp_listener()
                
            success, msg = c.save()
            if not success:
                QMessageBox.critical(self, "Save Error", f"Failed to save settings file:\n{msg}")
                return False
                
            if hasattr(self.gui, 'refresh_custom_buttons'): self.gui.refresh_custom_buttons()
            if hasattr(self.gui, 'refresh_main_buttons'): self.gui.refresh_main_buttons()
            
            self._is_saved = True
            print("[SYSTEM] Settings successfully saved to file.")
            return True
            
        except Exception as e:
            traceback.print_exc()
            QMessageBox.critical(self, "Settings Crash", f"A critical error occurred while trying to save your settings:\n\n{str(e)}")
            return False

    # Intercept all window close events to force an auto-save
    def closeEvent(self, event):
        if self.save_settings():
            event.accept()
        else:
            event.ignore() # Don't close the window if the save failed!

    def accept(self):
        if self.save_settings():
            super().accept()

    def reject(self):
        if self.save_settings():
            super().reject()


# =====================================================================
# MAIN WINDOW
# =====================================================================

class DataLoggerMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config = ConfigManager()
        self.monitor_manager = MonitorManager()
        self.sqlite_manager = None
        self._active_threads = []
        self.main_button_widgets = {}

        self.setWindowTitle(f"Online Logger v{APP_VERSION} - {os.path.basename(self.config.current_project_path)}")
        self.resize(1350, 380); self.setMinimumSize(1050, 350)
        
        self.central_widget = QWidget(); self.setCentralWidget(self.central_widget)
        self.main_v_layout = QVBoxLayout(self.central_widget); self.main_v_layout.setSpacing(15); self.main_v_layout.setContentsMargins(15, 15, 15, 15)
        self.columns_layout = QHBoxLayout(); self.columns_layout.setSpacing(15); self.main_v_layout.addLayout(self.columns_layout)

        self.custom_events_frame = QFrame(); self.custom_events_frame.setObjectName("MainPanel"); self.custom_layout = QVBoxLayout(self.custom_events_frame); self.custom_layout.setContentsMargins(15, 15, 15, 15); self.custom_layout.setSpacing(10)
        h1 = QLabel("Custom Events"); h1.setObjectName("PanelHeader"); self.custom_layout.addWidget(h1)
        self.custom_notebook = QTabWidget(); self.custom_notebook.setTabPosition(QTabWidget.TabPosition.South); self.custom_notebook.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.custom_notebook.tabBar().setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu); self.custom_notebook.tabBar().customContextMenuRequested.connect(self._show_tab_ctx_menu)
        self.custom_layout.addWidget(self.custom_notebook); self.columns_layout.addWidget(self.custom_events_frame, stretch=10)

        self.general_events_frame = QFrame(); self.general_events_frame.setObjectName("MainPanel"); self.general_layout = QVBoxLayout(self.general_events_frame); self.general_layout.setContentsMargins(15, 15, 15, 15); self.general_layout.setSpacing(10)
        h2 = QLabel("Standard Logging"); h2.setObjectName("PanelHeader"); self.general_layout.addWidget(h2); self.columns_layout.addWidget(self.general_events_frame, stretch=5)

        self.config_frame = QFrame(); self.config_frame.setObjectName("MainPanel"); self.config_layout = QVBoxLayout(self.config_frame); self.config_layout.setContentsMargins(15, 15, 15, 15); self.config_layout.setSpacing(10)
        h3 = QLabel("System & Status"); h3.setObjectName("PanelHeader"); self.config_layout.addWidget(h3); self.columns_layout.addWidget(self.config_frame, stretch=5)

        self.progress_bar = QProgressBar(); self.progress_bar.setRange(0, 0); self.progress_bar.setFixedHeight(12); self.progress_bar.setVisible(False); self.main_v_layout.addWidget(self.progress_bar)
        self.status_bar = QStatusBar(); self.setStatusBar(self.status_bar)

        self.setup_static_ui()
        self.refresh_custom_buttons() 
        if self.config.get("sqlite_enabled", False): self.chk_sqlite_enabled.setChecked(True)
        if self.config.get("always_on_top"): self.chk_always_on_top.setChecked(True)
        if self.config.get("new_day_event_enabled", True): self.schedule_new_day()
        if self.config.get("hourly_event_enabled", True): self.schedule_hourly_log()
        self.restart_udp_listener()

        is_dark = self.config.get("dark_mode", False)
        self.apply_theme(is_dark)
        self.chk_dark_mode.blockSignals(True) 
        self.chk_dark_mode.setChecked(is_dark)
        self.chk_dark_mode.blockSignals(False)

        self.update_status("System Ready.")
    
    def toggle_theme(self, checked):
        self.config.set("dark_mode", checked)
        self.config.save()
        self.apply_theme(checked)

    def apply_theme(self, is_dark):
        app = QApplication.instance()
        palette = app.palette()
        
        if is_dark:
            palette.setColor(QPalette.ColorRole.Window, QColor(40, 40, 40))
            palette.setColor(QPalette.ColorRole.WindowText, QColor(255, 255, 255))
            palette.setColor(QPalette.ColorRole.Base, QColor(30, 30, 30))
            palette.setColor(QPalette.ColorRole.AlternateBase, QColor(40, 40, 40))
            palette.setColor(QPalette.ColorRole.Text, QColor(255, 255, 255))
            palette.setColor(QPalette.ColorRole.Button, QColor(50, 50, 50))
            palette.setColor(QPalette.ColorRole.ButtonText, QColor(255, 255, 255))
            
            cb_border = "#FFFFFF" 
            panel_border = "rgba(255, 255, 255, 0.2)"
            tab_bg = "rgba(255, 255, 255, 0.05)"
        else:
            palette.setColor(QPalette.ColorRole.Window, QColor(240, 240, 240))
            palette.setColor(QPalette.ColorRole.WindowText, QColor(0, 0, 0))
            palette.setColor(QPalette.ColorRole.Base, QColor(255, 255, 255))
            palette.setColor(QPalette.ColorRole.AlternateBase, QColor(233, 231, 227))
            palette.setColor(QPalette.ColorRole.Text, QColor(0, 0, 0))
            palette.setColor(QPalette.ColorRole.Button, QColor(240, 240, 240))
            palette.setColor(QPalette.ColorRole.ButtonText, QColor(0, 0, 0))
            
            cb_border = "#000000" 
            panel_border = "rgba(0, 0, 0, 0.25)"
            tab_bg = "rgba(0, 0, 0, 0.05)"
            
        app.setPalette(palette)
        
        self.setStyleSheet(f"""
            QMainWindow {{ background-color: palette(window); }}
            QFrame#MainPanel {{ background-color: palette(base); border: 1px solid {panel_border}; border-radius: 8px; }}
            QLabel#PanelHeader {{ color: #0078D4; font-size: 12px; font-weight: 900; letter-spacing: 1px; text-transform: uppercase; padding-bottom: 6px; border-bottom: 1px solid {panel_border}; }}
            QLabel#StatusLabel {{ font-size: 13px; color: palette(text); font-weight: bold; }}
            QTabWidget::pane {{ border: none; background-color: transparent; }}
            QTabBar::tab {{ background: {tab_bg}; border: 1px solid {panel_border}; border-radius: 4px; padding: 4px 15px; margin-right: 4px; margin-top: 8px; color: palette(text); font-weight: bold; }}
            QTabBar::tab:selected {{ background: #0078D4; color: white; border: 1px solid #0078D4; }}
            QTabBar::tab:hover:!selected {{ background: rgba(128, 128, 128, 0.15); }}
            QPushButton#ActionBtn {{ border: 1px solid {panel_border}; border-radius: 6px; padding: 8px 15px; font-weight: bold; font-size: 13px; background-color: {tab_bg}; color: palette(text); }}
            QPushButton#ActionBtn:hover {{ background-color: rgba(0, 120, 212, 0.1); border-color: #0078D4; color: #0078D4; }}
            QStatusBar {{ background-color: #0078D4; color: white; font-weight: bold; }}
            QProgressBar {{ border: none; background-color: rgba(128, 128, 128, 0.2); border-radius: 3px; text-align: center; color: transparent; }}
            QProgressBar::chunk {{ background-color: #28a745; border-radius: 3px; }}
            
            QCheckBox {{ spacing: 8px; color: palette(window-text); font-weight: bold; }}
            QCheckBox::indicator {{ width: 18px; height: 18px; border: 2px solid {cb_border}; border-radius: 4px; background-color: transparent; }}
            QCheckBox::indicator:hover {{ border: 2px solid #0078D4; }}
            QCheckBox::indicator:checked {{ background-color: #0078D4; border: 2px solid #0078D4; image: url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>'); }}
        """)
        
        if hasattr(self, 'main_button_widgets') and self.main_button_widgets:
            self.refresh_main_buttons()
            self.refresh_custom_buttons()

    def setup_static_ui(self):
        btn_grid = QGridLayout(); btn_grid.setSpacing(8)
        row, col = 0, 0
        for btn_name in ["Log on", "Log off", "Event", "SVP", "Manual KP Log"]:
            btn = AutoScalingButton(btn_name); btn.setMinimumHeight(60)
            bg, fg = self._get_color_tuple(btn_name); btn.setStyleSheet(self._get_button_stylesheet(bg, fg))
            
            btn.set_wedge_color(bg) 
            
            btn.clicked.connect(lambda checked=False, t=btn_name, b=btn: self.log_event(t, b))
            btn.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu); btn.customContextMenuRequested.connect(lambda pos, b=btn_name, w=btn: self._show_ctx_menu(pos, 0, True, b, w))
            btn_grid.addWidget(btn, row, col); self.main_button_widgets[btn_name] = btn
            col += 1
            if col > 1: col, row = 0, row + 1
            
        btn_hist = QPushButton("Add Historic Event"); btn_hist.setObjectName("ActionBtn"); btn_hist.setCursor(Qt.CursorShape.PointingHandCursor); btn_hist.clicked.connect(self.add_historic_event); btn_hist.setMinimumHeight(60); btn_grid.addWidget(btn_hist, row, col)
        
        # --- Shift Report Button ---
        col += 1
        if col > 1: col, row = 0, row + 1
        btn_handover = QPushButton("📄 Shift Report"); btn_handover.setObjectName("ActionBtn"); btn_handover.setCursor(Qt.CursorShape.PointingHandCursor); btn_handover.clicked.connect(lambda: HandoverReportDialog(self).exec()); btn_handover.setMinimumHeight(60); btn_grid.addWidget(btn_handover, row, col)

        self.general_layout.addLayout(btn_grid); self.general_layout.addStretch()

       # --- System & Status Buttons (2-Column Grid) ---
        sys_btn_layout = QGridLayout(); sys_btn_layout.setSpacing(8)
        
        self.btn_toggle_monitor = QPushButton("Start Monitoring")
        self.btn_toggle_monitor.setObjectName("ActionBtn")
        self.btn_toggle_monitor.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_toggle_monitor.clicked.connect(self.toggle_monitoring)
        
        btn_set = QPushButton("⚙️ Settings")
        btn_set.setObjectName("ActionBtn")
        btn_set.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_set.clicked.connect(lambda: SettingsDialog(self).exec())
        
        btn_help = QPushButton("❓ Help")
        btn_help.setObjectName("ActionBtn")
        btn_help.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_help.clicked.connect(lambda: HelpDialog(self).exec())
        
        btn_debug = QPushButton("🐞 Debug")
        btn_debug.setObjectName("ActionBtn")
        btn_debug.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_debug.clicked.connect(lambda: DebugDialog(self).exec())
        
        # Add to Grid: widget, row, column
        sys_btn_layout.addWidget(self.btn_toggle_monitor, 0, 0)
        sys_btn_layout.addWidget(btn_set, 0, 1)
        sys_btn_layout.addWidget(btn_help, 1, 0)
        sys_btn_layout.addWidget(btn_debug, 1, 1)
        
        self.config_layout.addLayout(sys_btn_layout)
        
        separator = QFrame(); separator.setFrameShape(QFrame.Shape.HLine); separator.setStyleSheet("background-color: rgba(128,128,128,0.2); margin: 5px 0px;"); self.config_layout.addWidget(separator)

        status_container = QWidget(); s_layout = QGridLayout(status_container); s_layout.setContentsMargins(0, 0, 0, 0); s_layout.setSpacing(8)
        lbl_mon_title = QLabel("Folder Monitor:"); lbl_mon_title.setObjectName("StatusLabel"); self.lbl_mon = QLabel("● Inactive"); self.lbl_mon.setStyleSheet("color: #FF6B6B; font-weight: bold; font-size: 13px;")
        lbl_sql_title = QLabel("SQLite DB:"); lbl_sql_title.setObjectName("StatusLabel"); self.chk_sqlite_enabled = QCheckBox("Enable"); self.chk_sqlite_enabled.setCursor(Qt.CursorShape.PointingHandCursor)
        self.spin_auto_sync = QSpinBox(); self.spin_auto_sync.setRange(0, 1440); self.spin_auto_sync.setSuffix(" min Auto-Sync"); self.spin_auto_sync.setValue(self.config.get('auto_sync_interval_min', 15))
        self.btn_manual_sync = QPushButton("Force Resync"); self.btn_manual_sync.setObjectName("ActionBtn"); self.btn_manual_sync.setEnabled(False)
        s_layout.addWidget(lbl_mon_title, 0, 0); s_layout.addWidget(self.lbl_mon, 0, 1); s_layout.addWidget(lbl_sql_title, 1, 0); s_layout.addWidget(self.chk_sqlite_enabled, 1, 1); s_layout.addWidget(self.spin_auto_sync, 2, 0); s_layout.addWidget(self.btn_manual_sync, 2, 1)
        lbl_win_title = QLabel("Window:"); lbl_win_title.setObjectName("StatusLabel"); self.chk_always_on_top = QCheckBox("Always On Top"); self.chk_always_on_top.setCursor(Qt.CursorShape.PointingHandCursor); 
        
        lbl_theme_title = QLabel("Theme:")
        lbl_theme_title.setObjectName("StatusLabel")
        self.chk_dark_mode = QCheckBox("Dark Mode")
        self.chk_dark_mode.setCursor(Qt.CursorShape.PointingHandCursor)

        s_layout.addWidget(lbl_win_title, 3, 0)
        s_layout.addWidget(self.chk_always_on_top, 3, 1)
        s_layout.addWidget(lbl_theme_title, 4, 0)     
        s_layout.addWidget(self.chk_dark_mode, 4, 1)
        self.chk_dark_mode.toggled.connect(self.toggle_theme)
        
        self.chk_always_on_top.toggled.connect(self.toggle_always_on_top); self.chk_sqlite_enabled.toggled.connect(self.toggle_sqlite_mirroring); self.spin_auto_sync.valueChanged.connect(self.update_auto_sync_interval); self.btn_manual_sync.clicked.connect(self.run_manual_sqlite_sync)
        self.config_layout.addWidget(status_container); self.config_layout.addStretch()

    def update_status(self, message): 
        self.statusBar().showMessage(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {message}")

    def refresh_custom_buttons(self, focus_tab=None):
        active_tab_name = focus_tab or ("Main" if self.custom_notebook.count() == 0 else self.custom_notebook.tabText(max(0, self.custom_notebook.currentIndex())))
        while self.custom_notebook.count() > 0:
            w = self.custom_notebook.widget(0)
            self.custom_notebook.removeTab(0)
            if w: w.deleteLater()
        tabs = sorted(list(set(self.config.get("custom_button_tab_groups", ["Main"])))); 
        if "Main" not in tabs: tabs.insert(0, "Main")
        configs = self.config.get("custom_button_configs", [])
        
        for tab_name in tabs:
            t_widget = QWidget(); t_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding); t_widget.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu); t_widget.customContextMenuRequested.connect(lambda pos, w=t_widget: self._show_add_button_ctx_menu(pos, w))
            t_grid = QGridLayout(t_widget); t_grid.setSpacing(12); t_grid.setContentsMargins(10, 10, 10, 10)
            tab_configs = [c for c in configs[:self.config.get("num_custom_buttons", 3)] if c.get("tab_group", "Main") == tab_name][:15]
            if not tab_configs: self.custom_notebook.addTab(t_widget, tab_name); continue
            cols, rows = 5, min(3, (len(tab_configs) + 4) // 5)
            for c in range(cols): t_grid.setColumnStretch(c, 1)
            for r in range(rows): t_grid.setRowStretch(r, 1)
            for i, c in enumerate(tab_configs):
                btn_txt = c.get("text", "Custom"); btn = AutoScalingButton(btn_txt); btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
                bg, fg = self._get_color_tuple(btn_txt); btn.setStyleSheet(self._get_button_stylesheet(bg, fg))
                
                btn.set_wedge_color(bg) 
                
                btn.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu); m_idx = configs.index(c)
                btn.customContextMenuRequested.connect(lambda pos, idx=m_idx, b=btn_txt, w=btn: self._show_ctx_menu(pos, idx, False, b, w))
                btn.clicked.connect(lambda checked=False, conf=c, b=btn: self.log_custom_event(conf, b))
                t_grid.addWidget(btn, i // cols, i % cols)
            self.custom_notebook.addTab(t_widget, tab_name)
        for i in range(self.custom_notebook.count()):
            if self.custom_notebook.tabText(i) == active_tab_name: self.custom_notebook.setCurrentIndex(i); break

    def refresh_main_buttons(self):
        for text, btn in self.main_button_widgets.items():
            bg, fg = self._get_color_tuple(text)
            btn.setStyleSheet(self._get_button_stylesheet(bg, fg))
            btn.set_wedge_color(bg)

    def _get_color_tuple(self, text):
        val = self.config.get("button_colors", {}).get(text)
        if isinstance(val, (list, tuple)) and len(val) == 2: return val[0], val[1]
        return {"Log on": "#90EE90", "Log off": "#FFB6C1", "SVP": "#ADD8E6"}.get(text, None), None

    def _get_button_stylesheet(self, bg_color, fg_color):
        app = QApplication.instance()
        is_light_mode = app.palette().color(QPalette.ColorRole.Window).lightness() > 128
        
        if is_light_mode:
            bg, fg, border_col = '#FFFFFF', '#000000', '#D1D1D1'
            hover_bg, pressed_bg = '#F4F4F4', '#E5E5E5'
        else:
            bg, fg, border_col = '#2A2A2D', '#FFFFFF', '#3F3F46'
            hover_bg, pressed_bg = '#3E3E42', '#1E1E20'
            
        return f"""
            QPushButton {{ 
                background-color: {bg}; 
                color: {fg}; 
                font-weight: bold; 
                border: 1px solid {border_col}; 
                border-radius: 6px; 
                padding: 10px 12px; 
            }} 
            QPushButton:hover {{ border: 1px solid #0078D4; background-color: {hover_bg}; }} 
            QPushButton:pressed {{ background-color: {pressed_bg}; padding-top: 12px; padding-bottom: 8px; }}
            QPushButton:disabled {{ background-color: rgba(128, 128, 128, 0.05); color: rgba(128, 128, 128, 0.4); border: 1px dashed rgba(128, 128, 128, 0.2); }}
        """

    def _show_ctx_menu(self, pos, index, is_main, btn_name, widget):
        menu = QMenu(); act_edit = QAction("Edit Settings...", self)
        config = self.config.get("main_button_configs", {}).get(btn_name, {}) if is_main else self.config.get("custom_button_configs", [])[index]
        act_edit.triggered.connect(lambda: ButtonEditDialog(self, index, config, is_main, btn_name).exec()); menu.addAction(act_edit)
        if not is_main: menu.addSeparator(); act_delete = QAction("Delete Button", self); act_delete.triggered.connect(lambda: self._delete_custom_button(index)); menu.addAction(act_delete)
        menu.exec(widget.mapToGlobal(pos))

    def _show_add_button_ctx_menu(self, pos, widget):
        menu = QMenu(); act_add = QAction("Add New Custom Button", self); act_add.triggered.connect(self._add_new_custom_button); menu.addAction(act_add); menu.exec(widget.mapToGlobal(pos))

    def _show_tab_ctx_menu(self, pos):
        menu = QMenu(); act_add = QAction("Add New Tab...", self); act_add.triggered.connect(self._add_new_tab); menu.addAction(act_add)
        tab_bar = self.custom_notebook.tabBar(); t_idx = tab_bar.tabAt(pos)
        if t_idx >= 0:
            name = self.custom_notebook.tabText(t_idx); menu.addSeparator()
            act_ren = QAction(f"Rename '{name}'...", self); act_ren.triggered.connect(lambda: self._rename_tab(name)); menu.addAction(act_ren)
            act_del = QAction(f"Delete '{name}'", self); act_del.triggered.connect(lambda: self._delete_tab(name)); menu.addAction(act_del)
            if name == "Main": act_ren.setEnabled(False); act_del.setEnabled(False)
        menu.exec(tab_bar.mapToGlobal(pos))

    def _add_new_custom_button(self):
        if self.config.get("num_custom_buttons", 3) >= 90: return QMessageBox.information(self, "Limit", "Max custom buttons reached.")
        self.config.set("num_custom_buttons", self.config.get("num_custom_buttons", 3) + 1)
        curr_tab = self.custom_notebook.tabText(self.custom_notebook.currentIndex()) if self.custom_notebook.currentIndex() >= 0 else "Main"
        new_config = {"text": f"Custom {self.config.get('num_custom_buttons')}", "event_text": "Triggered", "txt_source_key": "None", "tab_group": curr_tab, "event_code": ""}
        configs = self.config.get("custom_button_configs", []); configs.append(new_config); self.config.set("custom_button_configs", configs); self.config.save()
        ButtonEditDialog(self, len(configs) - 1, new_config, is_main=False, button_name=new_config["text"]).exec()
        self.refresh_custom_buttons()

    def _delete_custom_button(self, index):
        configs = self.config.get("custom_button_configs", [])
        if 0 <= index < len(configs):
            txt = configs[index].get("text")
            if QMessageBox.question(self, "Confirm", f"Delete '{txt}'?") == QMessageBox.StandardButton.Yes:
                del configs[index]; self.config.set("custom_button_configs", configs); self.config.set("num_custom_buttons", self.config.get("num_custom_buttons", 3) - 1)
                colors = self.config.get("button_colors", {})
                if txt in colors: del colors[txt]; self.config.set("button_colors", colors)
                self.config.save(); self.refresh_custom_buttons()

    def _add_new_tab(self):
        name, ok = QInputDialog.getText(self, "Add New Tab", "Enter tab name:")
        if ok and name.strip():
            groups = self.config.get("custom_button_tab_groups", ["Main"])
            if name.strip() not in groups: groups.append(name.strip()); self.config.set("custom_button_tab_groups", groups); self.config.save(); self.refresh_custom_buttons(focus_tab=name.strip())

    def _rename_tab(self, old):
        new, ok = QInputDialog.getText(self, "Rename Tab", f"Rename '{old}' to:", QLineEdit.EchoMode.Normal, old)
        if ok and new.strip() and new != old:
            groups = self.config.get("custom_button_tab_groups", ["Main"]); groups[groups.index(old)] = new.strip()
            configs = self.config.get("custom_button_configs", [])
            for cfg in configs: 
                if cfg.get("tab_group") == old: cfg["tab_group"] = new.strip()
            self.config.set("custom_button_tab_groups", groups); self.config.set("custom_button_configs", configs); self.config.save(); self.refresh_custom_buttons(focus_tab=new.strip())

    def _delete_tab(self, name):
        if name == "Main": return
        if QMessageBox.question(self, "Confirm", f"Delete tab '{name}'?") == QMessageBox.StandardButton.Yes:
            groups = self.config.get("custom_button_tab_groups", ["Main"]); groups.remove(name)
            configs = self.config.get("custom_button_configs", [])
            for cfg in configs:
                if cfg.get("tab_group") == name: cfg["tab_group"] = "Main"
            self.config.set("custom_button_tab_groups", groups); self.config.set("custom_button_configs", configs); self.config.save(); self.refresh_custom_buttons()

    def toggle_always_on_top(self, checked):
        self.config.set("always_on_top", checked); self.config.save()
        try:
            import ctypes; from ctypes import wintypes
            SetWindowPos = ctypes.windll.user32.SetWindowPos
            SetWindowPos.argtypes = [wintypes.HWND, wintypes.HWND, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_uint]
            SetWindowPos(int(self.winId()), -1 if checked else -2, 0, 0, 0, 0, 3)
        except Exception: self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, checked); self.show(); self.raise_(); self.activateWindow()

    def toggle_monitoring(self):
        if not self.monitor_manager.is_monitoring:
            folder_configs = {}
            for name, path in self.config.get("folder_paths", {}).items():
                folder_configs[name] = {
                    "path": path, 
                    "ext": self.config.get("file_extensions", {}).get(name, ""), 
                    "skip": self.config.get("folder_skips", {}).get(name, False)
                }
            
            # 1. Update UI for the scanning phase
            self.btn_toggle_monitor.setEnabled(False)
            self.btn_toggle_monitor.setText("Scanning...")
            self.progress_bar.setVisible(True)
            self.update_status("Scanning folders to build initial file cache. Please wait...")
            
            # 2. Push the heavy folder scanning to a background QThread
            self.mon_thread = QThread()
            self.mon_worker = MonitorSetupWorker(self.monitor_manager, folder_configs)
            self.mon_worker.moveToThread(self.mon_thread)
            self._active_threads.append((self.mon_thread, self.mon_worker))
            
            self.mon_thread.started.connect(self.mon_worker.run)
            self.mon_worker.finished.connect(self._on_monitor_started)
            self.mon_worker.finished.connect(self.mon_thread.quit)
            self.mon_worker.finished.connect(self.mon_worker.deleteLater)
            self.mon_thread.finished.connect(self.mon_thread.deleteLater)
            self.mon_thread.finished.connect(lambda t=self.mon_thread, w=self.mon_worker: self._active_threads.remove((t, w)) if (t, w) in self._active_threads else None)
            
            self.mon_thread.start()
            
        else:
            self.monitor_manager.stop_monitoring()
            self.lbl_mon.setText("● Inactive")
            self.lbl_mon.setStyleSheet("color: #FF6B6B; font-weight: bold; font-size: 13px;")
            self.btn_toggle_monitor.setText("Start Monitoring")
            self.btn_toggle_monitor.setStyleSheet("")
            self.update_status("Folder monitoring stopped.")

    def _on_monitor_started(self, success, msg):
        """Callback triggered when the background scan finishes."""
        self.progress_bar.setVisible(False)
        self.btn_toggle_monitor.setEnabled(True)
        
        if success:
            self.lbl_mon.setText("● LIVE")
            self.lbl_mon.setStyleSheet("color: #28a745; font-weight: 900; font-size: 14px;")
            self.btn_toggle_monitor.setText("Stop Monitoring")
            self.btn_toggle_monitor.setStyleSheet("QPushButton { background-color: rgba(40, 167, 69, 0.15); border: 2px solid #28a745; color: #28a745; border-radius: 6px; padding: 8px 15px; font-weight: bold; }")
            self.update_status(msg)
        else:
            self.btn_toggle_monitor.setText("Start Monitoring")
            self.btn_toggle_monitor.setStyleSheet("")
            QMessageBox.warning(self, "Monitor Failed", msg)
            self.update_status("Monitor failed to start.")

    def toggle_sqlite_mirroring(self, checked):
        self.config.set("sqlite_enabled", checked)
        self.config.save()
        
        if hasattr(self, 'btn_manual_sync') and self.btn_manual_sync: 
            self.btn_manual_sync.setEnabled(checked)
            
        if checked:
            excel_path = self.config.get("log_file_path")
            
            # Check if the path is valid
            if not excel_path or not os.path.exists(excel_path):
                error_msg = "You must set a valid Excel Log Path in Settings before enabling SQLite."
                self.update_status(f"SQLite disabled: {error_msg}")
                
                # ONLY show the loud popup if the app is fully visible (meaning the user manually clicked it)
                # If it's not visible yet, it means the app is still booting up, so we skip the popup!
                if self.isVisible():
                    QMessageBox.critical(self, "Error", error_msg)
                
                # Silently uncheck the box
                self.chk_sqlite_enabled.blockSignals(True)
                self.chk_sqlite_enabled.setChecked(False)
                self.chk_sqlite_enabled.blockSignals(False)
                return
                
            db_path = self.config.get("sqlite_db_path") or str(Path(excel_path).with_suffix('.db'))
            self.update_status(f"Enabling SQLite mirror at {db_path}")
            self.sqlite_manager = SQLiteManager(db_path)
            self.run_manual_sqlite_sync()
            self.start_auto_sync()
        else:
            self.stop_auto_sync()
            if self.sqlite_manager: 
                self.sqlite_manager.close()
                self.sqlite_manager = None
            self.update_status("SQLite mirroring disabled.")

    def run_manual_sqlite_sync(self):
        if not self.chk_sqlite_enabled.isChecked() or not self.sqlite_manager or getattr(self, 'is_syncing_db', False): return
        
        excel_path = self.config.get("log_file_path")
        if not excel_path or not os.path.exists(excel_path):
            QMessageBox.critical(self, "Sync Error", "Cannot sync to SQLite: Excel file path is missing or invalid.")
            return
            
        self.is_syncing_db = True; self.update_status("Starting full SQLite sync..."); self.btn_manual_sync.setEnabled(False); self.btn_manual_sync.setText("Syncing...")
        def _find_header_row(excel_file):
            try:
                wb = xw.Book(excel_file); sheet = wb.sheets[0]
                for i in range(1, 31):
                    row_vals = sheet.range(f'A{i}:AZ{i}').value
                    if row_vals and any(h in [str(x).strip().lower() for x in row_vals if x] for h in ["event", "kp"]): return i - 1 
                return 0
            except Exception: return 0
        self.sync_thread = QThread()
        self.sync_worker = SqliteSyncWorker(self.sqlite_manager, excel_path, _find_header_row)
        self.sync_worker.moveToThread(self.sync_thread)
        self._active_threads.append((self.sync_thread, self.sync_worker))
        self.sync_thread.started.connect(self.sync_worker.run); self.sync_worker.finished.connect(self._on_sync_finished); self.sync_worker.finished.connect(self.sync_thread.quit); self.sync_worker.finished.connect(self.sync_worker.deleteLater); self.sync_thread.finished.connect(self.sync_thread.deleteLater)
        self.sync_thread.start()

    def _on_sync_finished(self, success, msg):
        self.is_syncing_db = False
        if self.chk_sqlite_enabled.isChecked(): self.btn_manual_sync.setEnabled(True)
        self.btn_manual_sync.setText("Force Resync")
        if success: self.update_status(msg)
        else:
            self.update_status(f"Sync Failed: {msg}")
            QMessageBox.warning(self, "SQLite Sync Failed", f"Failed to synchronize database.\n\nError: {msg}")

    def update_auto_sync_interval(self):
        self.config.set('auto_sync_interval_min', self.spin_auto_sync.value()); self.config.save()
        if self.chk_sqlite_enabled.isChecked(): self.start_auto_sync()

    def start_auto_sync(self):
        self.stop_auto_sync()
        if self.chk_sqlite_enabled.isChecked() and self.spin_auto_sync.value() > 0:
            self.auto_sync_timer = QTimer(self); self.auto_sync_timer.timeout.connect(self.run_manual_sqlite_sync); self.auto_sync_timer.start(self.spin_auto_sync.value() * 60 * 1000)

    def stop_auto_sync(self):
        if hasattr(self, 'auto_sync_timer') and self.auto_sync_timer.isActive(): self.auto_sync_timer.stop()

    def restart_udp_listener(self):
        if hasattr(self, 'udp_worker') and self.udp_worker:
            self.udp_worker.stop()
            if hasattr(self, 'udp_thread') and self.udp_thread: self.udp_thread.quit(); self.udp_thread.wait()
            self.udp_worker, self.udp_thread = None, None
            
        if self.config.get("udp_trigger_enabled", False):
            self.udp_thread = QThread()
            
            # --- Pass Custom Payloads to the Worker ---
            self.udp_worker = UdpListenerWorker(
                self.config.get("udp_trigger_port", 5999),
                self.config.get("udp_payload_recording", "RECORDING"),
                self.config.get("udp_payload_idle", "IDLE")
            )
            
            self.udp_worker.moveToThread(self.udp_thread)
            self.udp_thread.started.connect(self.udp_worker.run)
            self.udp_worker.trigger_log.connect(self._handle_udp_trigger)
            self.udp_thread.start()

    def _handle_udp_trigger(self, action):
        self.update_status(f"UDP Payload Detected: Executing '{action}'...")
        self.log_event(action, self.main_button_widgets.get(action))

    def _get_parsed_txt_data(self, src_key):
        if not src_key or src_key == "None": return {}
        
        folder_path = self.config.get("txt_folder_paths", {}).get(src_key, "")
        
        # 1. Soft Warning: Path not set
        if not folder_path: 
            self.update_status(f"Warning: No folder configured in Settings for '{src_key}'.")
            return {}
            
        # 2. Soft Warning: Path missing/disconnected
        if not os.path.exists(folder_path): 
            self.update_status(f"Warning: Folder path does not exist for '{src_key}': {folder_path}")
            return {}
            
        latest_file, latest_time = None, -1
        for root, _, files in os.walk(folder_path):
            for f in files:
                if f.lower().endswith(('.txt', '.csv', '.npd')):
                    fp = os.path.join(root, f)
                    try:
                        mtime = os.path.getmtime(fp)
                        if mtime > latest_time: latest_time, latest_file = mtime, fp
                    except Exception: pass
                    
        # 3. Soft Warning: No files found
        if not latest_file: 
            self.update_status(f"Warning: No valid data files (.txt, .csv) found in '{src_key}' folder.")
            return {}
            
        try:
            with open(latest_file, 'r') as f:
                lines = [l for l in f.readlines() if l.strip()]
                # 4. Soft Warning: Empty file
                if not lines: 
                    self.update_status(f"Warning: The latest file in '{src_key}' is completely empty.")
                    return {}
                last_line = lines[-1].strip()
                
            parts = [p.strip() for p in last_line.split(',')]
            parsed_data = {}
            for i, m in enumerate(self.config.get("all_txt_mappings", {}).get(src_key, [])):
                if not m.get('skip', False) and m.get('column_name', '').strip() and i < len(parts): 
                    parsed_data[m.get('column_name', '').strip()] = parts[i]
            return parsed_data
            
        # 5. Soft Warning: File locks or IO exceptions
        except PermissionError:
            self.update_status(f"Error: Permission denied to read file in '{src_key}'.")
            return {}
        except Exception as e: 
            self.update_status(f"Error reading '{src_key}': {str(e)}")
            return {}

    def _get_static_excel_data(self):
        static_data = {}
        static_configs = self.config.get("static_field_configs", [])
        if not static_configs: return static_data
        try:
            wb = xw.Book(self.config.get("log_file_path"))
            for config in static_configs:
                excel_col_key = config.get("field", "").strip()
                lookup_str = config.get("column_name", "").strip() 
                match = re.match(r"='?([^'!]+)'?!([A-Z]+\d+)", lookup_str)
                if match: static_data[excel_col_key] = wb.sheets[match.group(1)].range(match.group(2)).value
                elif re.match(r"=?([A-Z]+\d+)", lookup_str): static_data[excel_col_key] = wb.sheets[0].range(re.match(r"=?([A-Z]+\d+)", lookup_str).group(1)).value
        except Exception as e:
            self.update_status(f"Warning: Could not fetch Static Excel Lookups: {e}")
        return static_data

    def log_event(self, ev_type, btn):
        if ev_type == "Manual KP Log": return self.trigger_manual_hourly_log_action(btn)
        conf = self.config.get("main_button_configs", {}).get(ev_type, {})
        self._perform_log_action(ev_type, conf.get("event_text", ev_type), btn, conf.get("txt_source_key", "Main TXT"))

    def log_custom_event(self, config, btn):
        self._perform_log_action(config.get("text", "Custom"), config.get("event_text", "Triggered"), btn, config.get("txt_source_key", "None"))

    def _perform_log_action(self, event_type, event_text_for_excel, triggering_button, txt_source_key, override_txt_data=None, override_utc_datetime=None, skip_monitored_folders=False, additional_data=None):
        """Initiates a logging action on a background thread to prevent GUI freezing."""
        original_text = None
        
        # Native PySide6 button handling
        if triggering_button:
            original_text = triggering_button.text()
            if original_text == "Working...":
                original_text = getattr(self, 'original_manual_btn_text', event_type)
                
            triggering_button.setEnabled(False)
            triggering_button.setText("Working...")
            
        self.update_status(f"Processing '{event_type}'...")
            
        try:
            # 1. Gather all data synchronously first
            row_data = {}
            base_time = override_utc_datetime if override_utc_datetime else datetime.datetime.now()
            
            txt_data = self._get_parsed_txt_data(txt_source_key)
            for k, v in txt_data.items(): 
                row_data[str(k).strip()] = v

            kp_col_name = "KP"
            for cfg in self.config.get("all_txt_mappings", {}).get(txt_source_key, []):
                if cfg.get("field") == "KP" and not cfg.get("skip"): 
                    kp_col_name = cfg.get("column_name", "KP")
                    break

            if event_type == "Log on":
                try:
                    self._cached_log_on_kp = float(row_data.get(kp_col_name, ""))
                    self._cached_log_on_time = base_time
                    self.update_status("Log On successful. Stored KP.")
                except ValueError: 
                    self._cached_log_on_kp, self._cached_log_on_time = None, None
                    
            elif event_type == "Log off" and self.config.get("calculate_logoff_values", True):
                log_on_kp = getattr(self, '_cached_log_on_kp', None)
                log_on_time = getattr(self, '_cached_log_on_time', None)
                if log_on_kp is not None and log_on_time is not None:
                    try:
                        time_diff_secs = (base_time - log_on_time).total_seconds()
                        distance_km = abs(float(row_data.get(kp_col_name, "")) - log_on_kp)
                        speed_knots = (distance_km / 1.852) / (time_diff_secs / 3600.0) if time_diff_secs > 1 else 0.0
                        event_text_for_excel = f"Log off - Traveled: {distance_km:.3f} km @ {speed_knots:.2f} kts"
                    except ValueError: pass
                self._cached_log_on_kp, self._cached_log_on_time = None, None

            for config in self.config.get("generated_fields_config", []):
                target_col = config.get("column_name", config.get("field")).strip()
                src = config.get("source", "")
                if "UUID" in src or config["field"] == "UUID": row_data[target_col] = str(uuid.uuid4())
                elif "UTC" in src or config["field"] == "Date-Time": row_data[target_col] = base_time.strftime("%Y-%m-%d %H:%M:%S")
                elif "Local" in src or config["field"] == "Local Time": row_data[target_col] = (base_time + datetime.timedelta(hours=self.config.get("time_offset_hours", 0.0))).strftime("%Y-%m-%d %H:%M:%S")
                elif config["field"] == "Event": row_data[target_col] = event_text_for_excel
                elif config["field"] == "Code":
                    code = self.config.get("main_button_configs", {}).get(event_type, {}).get("event_code", "")
                    if not code:
                        for cb in self.config.get("custom_button_configs", []):
                            if cb.get("text") == event_type: 
                                code = cb.get("event_code", "")
                                break
                    row_data[target_col] = code
                elif config["field"] == "KP Ref.": row_data[target_col] = self.config.get("txt_source_aliases", {}).get(txt_source_key, txt_source_key)

            for k, v in self._get_static_excel_data().items(): row_data[k.strip()] = v

            if self.monitor_manager.is_monitoring and not skip_monitored_folders:
                paths = self.config.get("folder_paths", {})
                skips = self.config.get("folder_skips", {})
                cols = self.config.get("folder_columns", {})
                log_x = self.config.get("folder_log_x_instead", {})
                log_ext = self.config.get("folder_log_ext_vars", {})
                for name in paths.keys():
                    if skips.get(name, False): continue
                    latest_file_path = self.monitor_manager.get_latest_file(name)
                    if latest_file_path: 
                        row_data[cols.get(name, name).strip()] = "X" if log_x.get(name, False) else (os.path.basename(latest_file_path) if log_ext.get(name, False) else os.path.splitext(os.path.basename(latest_file_path))[0])

            if override_txt_data: 
                for k, v in override_txt_data.items(): row_data[str(k).strip()] = v
                
            if additional_data:
                for k, v in additional_data.items(): row_data[str(k).strip()] = v

            if not any(k for k in row_data.keys() if "UUID" in str(k).upper()): row_data["UUID"] = str(uuid.uuid4())
            if not any(k for k in row_data.keys() if "Event" in str(k)): row_data["Event"] = event_text_for_excel

            bg, _ = self._get_color_tuple(event_type)
            
            # 2. PySide6 Threading 
            logger = ExcelLogger(self.config.get("log_file_path"), self.sqlite_manager)
            thread = QThread()
            
            # Pass button, text, and static configs into the worker so it handles them safely
            worker_payload = {
                'row_data': row_data, 
                'bg_color': bg,
                'static_configs': self.config.get("static_field_configs", [])
            }
            worker = LogWorker(logger, worker_payload, triggering_button, original_text)
            worker.moveToThread(thread)
            self._active_threads.append((thread, worker))
            
            # ---> THIS IS THE MISSING LINE WE NEED TO ADD BACK <---
            thread.started.connect(worker.run) 
            
            # Connect directly to the method. Qt will automatically route this to the Main UI Thread!
            worker.finished.connect(self._on_log_complete)
            
            worker.finished.connect(thread.quit)
            worker.finished.connect(worker.deleteLater)
            thread.finished.connect(thread.deleteLater)
            thread.finished.connect(lambda t=thread, w=worker: self._active_threads.remove((t, w)) if (t, w) in self._active_threads else None)
            
            thread.start()
            
        except Exception as e:
            traceback.print_exc()
            self.update_status(f"Log Prep Failed: {e}")
            if triggering_button:
                triggering_button.setEnabled(True)
                triggering_button.setText(original_text)

    def _on_log_complete(self, excel_success, sqlite_success, msg, btn, orig_text):
        if btn: btn.setEnabled(True); btn.setText(orig_text)
        if excel_success:
            self.update_status(f"Log Success: {msg}")
        else:
            self.update_status(f"Log FAILED: {msg}")
            # --- LOUD Error handling for Excel Write Failures ---
            QMessageBox.critical(
                self, 
                "Logging Failed", 
                f"An error occurred while writing to Excel:\n\n{msg}\n\n"
                "Common Causes:\n"
                "• The Excel file is open and locked by another user.\n"
                "• The file path in Settings is incorrect or disconnected.\n"
                "• The file is set to Read-Only."
            )

    def schedule_new_day(self):
        """Calculates exact milliseconds until midnight and schedules the trigger."""
        now = datetime.datetime.now()
        
        # Calculate exactly when tomorrow's midnight is
        tomorrow = now + datetime.timedelta(days=1)
        midnight = datetime.datetime(
            year=tomorrow.year, 
            month=tomorrow.month, 
            day=tomorrow.day, 
            hour=0, minute=0, second=0, microsecond=0
        )
        
        # Calculate the exact difference in milliseconds
        time_until_midnight_ms = int((midnight - now).total_seconds() * 1000)
        
        # Add a tiny 1-second buffer to ensure it triggers ON the new day, not right before it
        trigger_delay_ms = time_until_midnight_ms + 1000

        print(f"[SYSTEM] Next Midnight 'New Day' log scheduled in {time_until_midnight_ms/1000/3600:.2f} hours.")
        
        # Use a persistent QTimer
        if hasattr(self, 'midnight_timer') and self.midnight_timer:
            self.midnight_timer.stop()
            
        self.midnight_timer = QTimer(self)
        self.midnight_timer.setSingleShot(True) # Only run once, we will manually reschedule it
        self.midnight_timer.timeout.connect(self.trigger_new_day)
        self.midnight_timer.start(trigger_delay_ms)

    def trigger_new_day(self):
        print("\n[SYSTEM] Executing Midnight 'New Day' Event...")
        if self.config.get("new_day_event_enabled", True):
            overrides = {}
            code = self.config.get("new_day_event_code", "")
            
            if code:
                for cfg in self.config.get("generated_fields_config", []):
                    if cfg.get("field") == "Code": 
                        overrides[cfg.get("column_name", "Code").strip()] = code
                        break
                        
            # Perform the log action in the background
            self._perform_log_action("New Day", "Midnight Position", None, "Main TXT", override_txt_data=overrides)
        else:
            print("[SYSTEM] 'New Day' event is disabled in settings. Skipping.")
            
        # Reschedule for tomorrow's midnight
        self.schedule_new_day()

    def schedule_hourly_log(self):
        now = datetime.datetime.now()
        next_hour = (now + datetime.timedelta(hours=1)).replace(minute=0, second=0, microsecond=0)
        QTimer.singleShot(int((next_hour - now).total_seconds() * 1000) + 1000, self.trigger_hourly_log)

    def trigger_hourly_log(self):
        self.schedule_hourly_log() 
        if self.config.get("hourly_event_enabled", True): 
            self.trigger_manual_hourly_log_action(None)

    def trigger_manual_hourly_log_action(self, btn):
        print("\n[DEBUG - HOURLY TRIGGER] Hourly Log Initiated.")
        if not self.config.get("hourly_event_enabled", True): 
            if btn: QMessageBox.information(self, "Disabled", "The 'Hourly KP Log' event is disabled.")
            return
            
        excel_path = self.config.get("log_file_path")
        if not excel_path or not os.path.exists(excel_path): 
            if btn: QMessageBox.critical(self, "Error", "Excel Log file missing. Please configure it in settings.")
            self.update_status("Manual KP Log skipped: Excel file missing.")
            return 
            
        if getattr(self, 'is_calculating_kp', False): 
            print("[DEBUG - HOURLY TRIGGER] Blocked: An hourly calculation is already running!")
            return
            
        self.is_calculating_kp = True
        
        if btn: 
            self.original_manual_btn_text = btn.text()
            btn.setEnabled(False)
            btn.setText("Working...")
            
        self.update_status("Processing 'Manual KP Log'...")
        
        src_key = self.config.get("hourly_log_txt_source_key", "Main TXT")
        print(f"[DEBUG - HOURLY TRIGGER] Gathering data from: {src_key}")
        
        kp_col_name, line_col_name, event_col_name, dt_col_name = "KP", "Runline", "Event", "UTC Date-Time"
        for cfg in self.config.get("all_txt_mappings", {}).get(src_key, []):
            if cfg.get("field") == "KP" and not cfg.get("skip"): kp_col_name = cfg.get("column_name", "KP")
            if cfg.get("field") == "Line name" and not cfg.get("skip"): line_col_name = cfg.get("column_name", "Line name")
        for cfg in self.config.get("generated_fields_config", []):
            if cfg.get("field") == "Event": event_col_name = cfg.get("column_name", "Event")
            if cfg.get("field") == "Date-Time": dt_col_name = cfg.get("column_name", "UTC Date-Time")

        txt_data = self._get_parsed_txt_data(src_key)
        try:
            current_kp = float(txt_data.get(kp_col_name))
            current_line = txt_data.get(line_col_name)
            if current_line is None: raise ValueError()
            print(f"[DEBUG - HOURLY TRIGGER] Captured Live Data -> KP: {current_kp}, Line: {current_line}")
        except Exception:
            msg = f"KP Log skipped: Failed to extract valid KP or Line Name from '{src_key}'."
            print(f"[DEBUG - HOURLY TRIGGER] FAILED: {msg}")
            self.update_status(msg)
            self.is_calculating_kp = False
            if btn: 
                btn.setEnabled(True)
                btn.setText(self.original_manual_btn_text)
                QMessageBox.warning(self, "Missing Source Data", f"Cannot perform Manual KP Log.\n\n{msg}\n\nPlease ensure your text data folder has valid files and your mapping is correct.")
            return

        print("[DEBUG - HOURLY TRIGGER] Starting background QThread...")
        thread = QThread()
        worker = HourlyCalcWorker(
            excel_path, event_col_name, kp_col_name, line_col_name, dt_col_name,
            current_kp, current_line, datetime.datetime.now(datetime.timezone.utc).replace(tzinfo=None), 
            getattr(self, '_cached_last_hourly_kp', None), getattr(self, '_cached_last_hourly_line', None), getattr(self, '_cached_last_hourly_time', None)
        )
        worker.moveToThread(thread)
        self._active_threads.append((thread, worker))
        self._current_manual_btn = btn
        
        thread.started.connect(worker.run)
        worker.finished.connect(self._on_calc_finished)
        worker.finished.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        thread.finished.connect(lambda: self._active_threads.remove((thread, worker)) if (thread, worker) in self._active_threads else None)
        
        thread.start()

    def _on_calc_finished(self, success, event_text, new_kp, new_line, new_time):
        print(f"[DEBUG - HOURLY TRIGGER] Background calculation returned. Success: {success}")
        self.is_calculating_kp = False
        btn = getattr(self, '_current_manual_btn', None)
        if success:
            self._cached_last_hourly_kp, self._cached_last_hourly_line, self._cached_last_hourly_time = new_kp, new_line, new_time
            print("[DEBUG - HOURLY TRIGGER] Passing generated text to LogWorker...")
            self._perform_log_action("Hourly KP Log", event_text, btn, self.config.get('hourly_log_txt_source_key', 'Main TXT'))
        else:
            if btn: 
                btn.setEnabled(True)
                btn.setText(getattr(self, 'original_manual_btn_text', 'Manual KP Log'))
            self.update_status(f"KP Log Calculation failed: {event_text}")

    def _on_calc_finished(self, success, event_text, new_kp, new_line, new_time):
        self.is_calculating_kp = False
        btn = getattr(self, '_current_manual_btn', None)
        if success:
            self._cached_last_hourly_kp, self._cached_last_hourly_line, self._cached_last_hourly_time = new_kp, new_line, new_time
            self._perform_log_action("Hourly KP Log", event_text, btn, self.config.get('hourly_log_txt_source_key', 'Main TXT'))
        else:
            if btn: 
                btn.setEnabled(True)
                btn.setText(getattr(self, 'original_manual_btn_text', 'Manual KP Log'))
            self.update_status(f"KP Log Calculation failed: {event_text}")

    
    def add_historic_event(self):
        d = HistoricEventDialog(self)
        if d.exec() == QDialog.DialogCode.Accepted:
            self.update_status(f"Searching for time '{d.result_data['time_str']}'...")
            found_line = self._search_file_for_line(d.result_data['file_path'], r"(?:^|[^A-Za-z0-9:])" + re.escape(d.result_data['time_str']))
            if not found_line: return QMessageBox.information(self, "Not Found", "Time not found in file."); self.update_status("Historic search failed.")
            parsed_data = self._parse_txt_line(found_line, "Main TXT")
            if not parsed_data: return QMessageBox.warning(self, "Parse Error", "Found the line, but could not map any data.")
            if HistoricPreviewDialog(self, found_line, parsed_data).exec() == QDialog.DialogCode.Accepted:
                try: final_dt = datetime.datetime.combine(datetime.date.fromtimestamp(os.path.getmtime(d.result_data['file_path'])), d.result_data['time_obj']).replace(tzinfo=datetime.timezone.utc)
                except Exception: final_dt = datetime.datetime.now(datetime.timezone.utc)
                if d.result_data['insert_sfile']:
                    sfile_col = self.config.get("folder_columns", {}).get("S-File")
                    if sfile_col:
                        closest_sfile = self._find_closest_sfile(final_dt)
                        parsed_data[sfile_col] = os.path.splitext(os.path.basename(closest_sfile))[0] if closest_sfile else "N/A"
                self.update_status("Logging historic event...")
                self._perform_log_action("Historic Event", f"Historic data for {final_dt.strftime('%Y-%m-%d %H:%M')}", None, "None", parsed_data, final_dt)
            else: self.update_status("Historic event cancelled by user.")

    def _search_file_for_line(self, file_path, search_pattern):
        if not search_pattern: return None
        try:
            regex = re.compile(search_pattern)
            for enc in ['utf-8', 'latin-1', 'cp1252']:
                try:
                    with open(file_path, "r", encoding=enc) as f:
                        for line in f:
                            if regex.search(line): return line.strip()
                except UnicodeDecodeError: continue
        except Exception: pass
        return None

    def _parse_txt_line(self, line_str, source_key="Main TXT"):
        if not line_str: return {}
        parts = [p.strip() for p in line_str.split(",")]
        parsed_data = {}
        for i, m in enumerate(self.config.get("all_txt_mappings", {}).get(source_key, [])):
            if not m.get('skip', False) and m.get('column_name', '').strip() and i < len(parts): parsed_data[m.get('column_name', '').strip()] = parts[i]
        return parsed_data

    def _find_closest_sfile(self, historic_dt):
        sfile_folder_path = self.config.get("folder_paths", {}).get("S-File")
        if not sfile_folder_path or not os.path.isdir(sfile_folder_path): return None
        candidate_files = []
        for root, _, files in os.walk(sfile_folder_path):
            for filename in files:
                try:
                    file_dt = datetime.datetime.strptime(os.path.splitext(filename)[0], "%Y%m%d_%H%M%S_S")
                    if file_dt < historic_dt.replace(tzinfo=None): candidate_files.append((file_dt, os.path.join(root, filename)))
                except ValueError: continue
        return max(candidate_files, key=lambda item: item[0])[1] if candidate_files else None

    def closeEvent(self, event):
        print("\n[SYSTEM] Shutting down application...")
        
        # 1. Stop Folder Monitoring
        if hasattr(self, 'monitor_manager') and self.monitor_manager:
            try: self.monitor_manager.stop_monitoring()
            except Exception: pass
            
        # 2. Close SQLite Connection
        if hasattr(self, 'sqlite_manager') and self.sqlite_manager:
            try: self.sqlite_manager.close()
            except Exception: pass
            
        # 3. Stop UDP Listener
        if hasattr(self, 'udp_worker') and self.udp_worker:
            try: self.udp_worker.stop()
            except Exception: pass
            
        if hasattr(self, 'udp_thread') and self.udp_thread:
            try:
                self.udp_thread.quit()
                # 600ms because the UDP network check takes 500ms to cycle!
                self.udp_thread.wait(600) 
            except Exception: pass

        # 4. Clean up any stuck background worker threads
        for thread, worker in getattr(self, '_active_threads', []):
            try:
                if thread.isRunning(): 
                    thread.quit()
                    # Give Excel writers a little more time to cleanly detach
                    thread.wait(600) 
            except Exception: pass
            
        print("[SYSTEM] Safe to close.")
        event.accept()

class DebugDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Application Debug Console")
        self.resize(800, 450)
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)
        
        layout = QVBoxLayout(self)
        self.text_browser = QTextBrowser()
        # Matrix-style console look
        self.text_browser.setStyleSheet("background-color: #1e1e1e; color: #00ff00; font-family: Consolas, monospace; font-size: 13px; padding: 10px;")
        layout.addWidget(self.text_browser)
        
        # Connect live output and pre-fill history
        if 'console_logger' in globals():
            self.text_browser.insertPlainText("".join(console_logger.history))
            self.text_browser.moveCursor(QTextCursor.MoveOperation.End)
            console_logger.written.connect(self.append_text)
            
        bot_lay = QHBoxLayout()
        
        # --- NEW: Force Crash Button ---
        btn_crash = QPushButton("☢️ Force Test Crash")
        btn_crash.setStyleSheet("QPushButton { border: 1px solid #FF6B6B; border-radius: 6px; padding: 8px 20px; font-weight: bold; background-color: rgba(255, 107, 107, 0.1); color: #FF6B6B; } QPushButton:hover { background-color: rgba(255, 107, 107, 0.3); }")
        btn_crash.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_crash.clicked.connect(self.trigger_fake_crash)
        bot_lay.addWidget(btn_crash)
        
        bot_lay.addStretch()
        
        btn_close = QPushButton("Close Console")
        btn_close.setStyleSheet("QPushButton { border: 1px solid rgba(128, 128, 128, 0.5); border-radius: 6px; padding: 8px 20px; font-weight: bold; background-color: #333333; color: white; } QPushButton:hover { background-color: #555555; }")
        btn_close.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_close.clicked.connect(self.accept)
        bot_lay.addWidget(btn_close)
        
        layout.addLayout(bot_lay)
        
    def append_text(self, text):
        # Auto-scroll to bottom when new text arrives
        self.text_browser.moveCursor(QTextCursor.MoveOperation.End)
        self.text_browser.insertPlainText(text)
        self.text_browser.moveCursor(QTextCursor.MoveOperation.End)

    def trigger_fake_crash(self):
        print("\n[DEBUG] User triggered a fake crash to test the exception handler!")
        # This will instantly halt the thread and send it to the global_exception_handler
        raise RuntimeError("This is a deliberately triggered test crash from the Debug console.")

class HelpDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Setup Guide & Help")
        self.resize(850, 750)
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint)
        self.setStyleSheet("""
            QDialog { background-color: palette(window); }
            QTextBrowser { background-color: palette(base); color: palette(text); font-family: 'Segoe UI', Arial, sans-serif; font-size: 14px; border: 1px solid rgba(128, 128, 128, 0.4); border-radius: 6px; padding: 15px; }
            h2 { color: #0078D4; margin-bottom: 5px; }
            h3 { color: #0078D4; border-bottom: 1px solid rgba(128, 128, 128, 0.2); padding-bottom: 4px; margin-top: 20px;}
            li { margin-bottom: 8px; line-height: 1.4; }
            QPushButton { border: 1px solid rgba(128, 128, 128, 0.5); border-radius: 6px; padding: 8px 20px; background-color: #0078D4; color: white; font-weight: bold; }
            QPushButton:hover { background-color: #106ebe; }
        """)

        layout = QVBoxLayout(self)

        browser = QTextBrowser()
        browser.setOpenExternalLinks(True)

        html_content = """
        <h2>Online Logger - Setup Guide & Help</h2>
        <p>Follow these steps to configure the logger correctly for your operations.</p>

        <h3>1. Excel Log File Setup</h3>
        <ul>
            <li>Create an empty Excel file with a header row at the top. <b>Supported formats:</b> <code>.xlsx</code>, <code>.xlsb</code>, or <code>.xlsm</code>. (Do not use the outdated <code>.xls</code> format, as it has row limits and lacks modern COM support).</li>
            <li><b>Important:</b> The program requires you to have columns named exactly (case-insensitive): <b>Runline</b>, <b>KP</b>, and <b>Event</b>.</li>
            <li>Go to <b>⚙️ Settings > File Paths</b> and browse to select this Excel file.</li>
            <li><i>Note:</i> Make sure the file is not opened in "Protected View" (e.g., downloaded from an email). If Excel shows a yellow bar at the top, click "Enable Editing" before running the logger.</li>
        </ul>

        <h3>2. Dashboard Shortcuts (Right-Click Menus)</h3>
        <ul>
            <li><b>Edit Buttons Instantly:</b> Right-click any button (Standard or Custom) directly on the main dashboard to quickly change its text, colors, event code, or data source without digging through the main settings menu.</li>
            <li><b>Manage Tabs:</b> Right-click the custom tab headers (e.g., "Main") to <b>Add</b>, <b>Rename</b>, or <b>Delete</b> tabs. This allows you to organize your custom buttons by operation type (e.g., 'ROV Ops', 'Deployment', 'Transit').</li>
            <li><b>Add New Buttons:</b> Right-click the empty background space inside any custom tab to instantly spawn a new custom button.</li>
        </ul>

        <h3>3. Navigation Data (TXT Sources)</h3>
        <ul>
            <li>The logger reads real-time survey strings from text files (comma-separated) generated by Qinsy, Eiva, Naviscan, etc.</li>
            <li>In <b>Settings > File Paths</b>, browse to the folder containing your live navigation output.</li>
            <li>Click <b>⚙️ Field Mapping</b>. You must tell the program which comma-separated value corresponds to which Excel column (e.g., Value 1 = KP, Value 3 = Line name). Check "Skip" for values you don't need.</li>
        </ul>

        <h3>4. Monitored Folders (SVP, Video, Sensors)</h3>
        <ul>
            <li>To automatically log the latest file generated by another system (like an SVP drop or a new Video file), go to <b>Settings > Monitored Folders</b>.</li>
            <li>Add the folder path and state the exact name of the Excel column it should write the filename to.</li>
            <li>Click <b>Start Monitoring</b> on the main dashboard to activate the background watchers. The monitor keeps a cache of the top 3 newest files, so if the acquisition software splits a file, the logger automatically tracks the newest one.</li>
        </ul>

        <h3>5. SQLite Database Mirror (Optional but Recommended)</h3>
        <ul>
            <li>To prevent data loss if Excel crashes or is closed, mirror your log to an SQLite database.</li>
            <li>In <b>Settings > File Paths</b>, set a path for the <code>.db</code> file. Check the <b>"Enable"</b> box on the main dashboard to activate it.</li>
        </ul>

        <h3>6. Custom Buttons & Event Codes</h3>
        <ul>
            <li>Go to <b>Settings > Event Codes</b> to define standard survey abbreviations (e.g., SOL, EOL, SVP). These will appear in dropdowns when editing buttons.</li>
            <li>Go to <b>Settings > Button Configuration</b> to mass-create custom buttons if you don't want to use the right-click method mentioned above.</li>
        </ul>

        <h3>7. Automation (Hourly, Midnight, UDP)</h3>
        <ul>
            <li><b>Hourly / Midnight Logs:</b> In <b>Settings > Programmed Events</b>, enable automatic logging of the Midnight position and Hourly KP progress. You must select the correct "KP Data Source" for the hourly math to work.</li>
            <li><b>UDP Triggers:</b> If your acquisition software broadcasts its status via UDP, enter the Listening Port and payload strings (e.g., 'RECORDING' or 'IDLE') to trigger hands-free 'Log on' and 'Log off' events.</li>
        </ul>

        <hr>
        <p><b>💡 Troubleshooting Tip - "Excel: Fail":</b> If the program fails to log or the UI seems unresponsive, ensure the Excel file is not locked in <i>Edit Mode</i>. If a user double-clicks an Excel cell and the cursor is actively blinking inside it, Windows blocks all external programs from writing. Simply press <b>Enter</b> or <b>Escape</b> in Excel to unlock it!</p>
        """
        browser.setHtml(html_content)
        layout.addWidget(browser)

        btn_close = QPushButton("Close")
        btn_close.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_close.clicked.connect(self.accept)
        layout.addWidget(btn_close, alignment=Qt.AlignmentFlag.AlignRight)

# =====================================================================
# INTERCEPTOR & LAUNCH APP
# =====================================================================

# =====================================================================
# INTERCEPTOR & LAUNCH APP
# =====================================================================

class ConsoleLogger(QObject):
    """Intercepts terminal print() statements and errors, sending them to the GUI."""
    written = Signal(str)
    
    def __init__(self):
        super().__init__()
        self.history = []
        # Use sys.__stdout__ which is the true original console, bypassing PyInstaller wrappers
        self.original_stdout = sys.__stdout__
        self.original_stderr = sys.__stderr__
        
    def write(self, text):
        # 1. Hyper-safe check: Only write if the terminal exists AND has a write function
        if self.original_stdout and hasattr(self.original_stdout, 'write'):
            try:
                self.original_stdout.write(text)
            except Exception:
                pass
                
        # 2. Add to our GUI Debug history
        self.history.append(text)
        
        # Keep memory clean (limit to last 5000 lines)
        if len(self.history) > 5000: 
            self.history.pop(0)
            
        # 3. Emit to the live UI safely
        self.written.emit(text)
        
    def flush(self):
        if self.original_stdout and hasattr(self.original_stdout, 'flush'):
            try:
                self.original_stdout.flush()
            except Exception:
                pass

if __name__ == "__main__":
    try:
        import ctypes
        myappid = f"Online Logger {APP_VERSION}"
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception: pass

    app = QApplication(sys.argv)
    
    # --- ACTIVATE PRINT INTERCEPTOR ---
    global console_logger
    console_logger = ConsoleLogger()
    sys.stdout = console_logger
    sys.stderr = console_logger
    
    app_icon = QIcon(resource_path("Logo.ico"))
    app.setWindowIcon(app_icon)
    app.setStyle("Fusion") 
    
    window = DataLoggerMainWindow()
    window.setWindowIcon(app_icon)
    window.show()
    sys.exit(app.exec())