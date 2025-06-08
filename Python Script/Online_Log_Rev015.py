import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser, Toplevel, Label
import os
import xlwings as xw # Keep xlwings for Excel interaction
import threading
import time
from watchdog.observers.polling import PollingObserver
from watchdog.events import FileSystemEventHandler
import datetime
import json
import traceback
import sqlite3 # Use Python's built-in SQLite module
import uuid
import pandas as pd

# Global cache
folder_cache = {}

# --- Tooltip Class (IMPROVED with Delays) ---
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

        # Bind events to intermediate handlers
        self.widget.bind("<Enter>", self.on_enter, add='+') # Use add='+ to coexist with button bindings
        self.widget.bind("<Leave>", self.on_leave, add='+')
        # self.widget.bind("<Destroy>", self.on_leave, add='+') # Might cause issues if triggered too often

    def on_enter(self, event=None):
        # When mouse enters, cancel any scheduled hide and schedule a show
        self.cancel_scheduled_hide()
        self.schedule_show()

    def on_leave(self, event=None):
        # When mouse leaves, cancel any scheduled show and schedule a hide
        self.cancel_scheduled_show()
        self.cancel_scheduled_show() # Ensure no show is pending
        self.schedule_hide()

    def schedule_show(self):
        # Cancel previous show timer if any
        self.cancel_scheduled_show()
        # Schedule the tooltip to appear after delay
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
        # Cancel any scheduled hide first (prevents duplicate calls)
        self.cancel_scheduled_hide()
        tw = self.tooltip_window
        self.tooltip_window = None # Set to None first
        if tw:
            try:
                tw.destroy()
            except tk.TclError:
                pass # Ignore if already destroyed


# --- FolderMonitor Class (Unchanged) ---
class FolderMonitor(FileSystemEventHandler):
    def __init__(self, folder_path, folder_name, gui_instance, file_extension=""):
        self.folder_path = folder_path
        self.folder_name = folder_name
        self.gui_instance = gui_instance
        self.file_extension = ("." + file_extension.lstrip('.')) if file_extension else ""

    def on_modified(self, event):
        if not event.is_directory:
            if not self.file_extension or event.src_path.lower().endswith(self.file_extension.lower()):
                self.update_latest_file()

    def on_created(self, event):
        if not event.is_directory:
            if not self.file_extension or event.src_path.lower().endswith(self.file_extension.lower()):
                self.update_latest_file()

    def update_latest_file(self):
        if not os.path.exists(self.folder_path):
            if self.folder_name in folder_cache: del folder_cache[self.folder_name]
            return
        try:
            files = []
            for f_name in os.listdir(self.folder_path):
                full_path = os.path.join(self.folder_path, f_name)
                try:
                    is_file = os.path.isfile(full_path)
                except OSError: continue
                if is_file and not f_name.endswith(".tmp"): # Ignore temp files
                    if not self.file_extension or f_name.lower().endswith(self.file_extension.lower()):
                        files.append(full_path)
            if files:
                accessible_files = []
                for file_path in files:
                    retries = 3; delay = 0.5
                    for _ in range(retries):
                        try:
                            mod_time = os.path.getmtime(file_path)
                            accessible_files.append((file_path, mod_time)); break
                        except (OSError, IOError) as e:
                            if _ == retries - 1: print(f"Warning: File {os.path.basename(file_path)} not accessible. Error: {e}")
                            else: time.sleep(delay)
                if accessible_files:
                    latest_file_path = max(accessible_files, key=lambda x: x[1])[0]
                    file_name = os.path.basename(latest_file_path)
                    if folder_cache.get(self.folder_name) != file_name:
                        folder_cache[self.folder_name] = file_name
                        # Update status via GUI thread safely
                        #if self.gui_instance and hasattr(self.gui_instance, 'master') and self.gui_instance.master.winfo_exists():
                        #   self.gui_instance.master.after(0, self.gui_instance.update_status, f"Latest {self.folder_name} file: {file_name}")
        except Exception as e:
            print(f"Error updating cache for {self.folder_name}: {e}")
            traceback.print_exc()


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
        * db_status_label: Label to display SQLite database status.
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
        * sqlite_enabled: Whether SQLite logging is enabled.
        * sqlite_db_path: Path to the SQLite database file.
        * sqlite_table: Default table name for SQLite logging.
        * main_frame: The main frame containing all widgets.
        [...]
    '''


    def __init__(self, master):
        '''
        Initializes the main GUI application.
        This method sets up the main window, initializes styles, variables, and loads settings.
        Arguments:
        * master: The root Tkinter window or parent widget.
        '''
        self.master = master
        master.title("Data Acquisition Logger (SQLite Mode)")
        master.geometry("550x550") # Slightly larger default size
        master.minsize(450, 400)
        self.settings_file = "logger_settings.json"
        self.init_styles()
        self.init_variables()
        self.load_settings()

        # --- Main Layout ---
        self.main_frame = ttk.Frame(self.master, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.master.columnconfigure(0, weight=1)
        self.master.rowconfigure(0, weight=1) # Allow button area to expand if needed

        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(0, weight=1) # Button frame row
        self.main_frame.rowconfigure(1, weight=0) # Indicator row
        self.main_frame.rowconfigure(2, weight=0) # Status bar row

        # Create button frame FIRST
        self.button_frame = ttk.Frame(self.main_frame, padding="10")
        self.button_frame.grid(row=0, column=0, sticky=(tk.N, tk.W, tk.E, tk.S))
        # Create buttons inside the frame (now happens in two stages)
        self.create_main_buttons(self.button_frame)

        # Create indicators and status bar AFTER button frame
        self.create_status_indicators(self.main_frame) # Status Indicators
        self.create_status_bar(self.main_frame)        # Status Bar

        # --- Final Setup ---
        self.schedule_new_day()
        self.start_monitoring()  # Initial monitor start & status update

        # Open the settings window by default when the app starts
        self.startup_settings()

    def init_styles(self):
        ''' 
        Initializes the styles for the application using ttk.Style.
        This method sets the theme and configures styles for various widgets.
        It also handles theme availability and sets default styles.
        '''
        self.style = ttk.Style()
        try:
            available_themes = self.style.theme_names()
            if 'vista' in available_themes: self.style.theme_use('vista')
            elif 'aqua' in available_themes: self.style.theme_use('aqua')
            else: self.style.theme_use("clam")
        except tk.TclError:
            self.style.theme_use("clam")

        self.style.configure("TFrame", background="#f0f0f0")
        self.style.configure("TLabel", background="#f0f0f0", font=("Arial", 10))
        self.style.configure("TButton", font=("Arial", 10, "bold"), padding=6)
        self.style.configure("TEntry", font=("Arial", 10), padding=6)
        self.style.configure("StatusBar.TLabel", background="#e0e0e0", font=("Arial", 9), relief=tk.SUNKEN, padding=(5, 2))
        self.style.configure("Header.TFrame", background="#dcdcdc")
        # Define styles for selected and unselected rows
        self.style.configure("Row0.TFrame", background="#ffffff")
        self.style.configure("Row1.TFrame", background="#f5f5f5")
        self.style.configure("Selected.TFrame", background="#ADD8E6") # Light blue for selection
        self.style.configure("TLabelframe", background="#f0f0f0", padding=5)
        self.style.configure("TLabelframe.Label", background="#f0f0f0", font=("Arial", 10, "bold"))
        self.style.configure("Large.TCheckbutton", font=("Arial", 11)) # For settings checkbox
        self.style.configure("Toolbutton", padding=2) # For settings color picker buttons
        self.style.configure("Accent.TButton", font=("Arial", 10, "bold"), foreground="white", background="#0078D4") # For settings save

        self.style.map("TButton",
                        foreground=[('pressed', 'darkblue'), ('active', 'blue'), ('disabled', '#999999')],
                        background=[('pressed', '!disabled', '#c0c0c0'), ('active', '#e0e0e0')]
                        )
        self.style.map("TLabel", background=[('selected', '#ADD8E6')]) # Ensure labels in selected row change color

    def init_variables(self):
        '''
        Initializes all key configuration variables, paths, button presets, and GUI state defaults used throughout the application. 
        This method is called when the GUI is first launched.
        '''
        self.log_file_path = None
        
        # Original TXT path for the 'Event' button
        self.txt_folder_path = None 
        # New TXT paths for additional sources
        self.txt_folder_path_set2 = None
        self.txt_folder_path_set3 = None

        self.source_based_colors = {
            "Main TXT": "#BAE1FF",      # Light Blue
            "TXT Source 2": "#BAFFC9",    # Light Green
            "TXT Source 3": "#FFFFBA",    # Light Yellow
            "None": None                # No color for buttons with no source
        }

        self.txt_file_path = None # This will now be dynamic based on source

        # Modified: Use a list of dicts for TXT field columns to preserve order
        self.txt_field_columns_config = [
            {"field": "Date", "column_name": "Date", "skip": False},
            {"field": "Time", "column_name": "Time", "skip": False},
            {"field": "KP", "column_name": "KP", "skip": False},
            {"field": "DCC", "column_name": "DCC", "skip": False},
            {"field": "Line name", "column_name": "Line name", "skip": False},
            {"field": "Latitude", "column_name": "Latitude", "skip": False},
            {"field": "Longitude", "column_name": "Longitude", "skip": False},
            {"field": "Easting", "column_name": "Easting", "skip": False},
            {"field": "Northing", "column_name": "Northing", "skip": False},
            {"field": "Event", "column_name": "Event", "skip": False} # Default "Event" field is still here
        ]
        # These two will be derived from txt_field_columns_config for backwards compatibility/easier lookup
        self.txt_field_columns = {cfg["field"]: cfg["column_name"] for cfg in self.txt_field_columns_config}
        self.txt_field_skips = {cfg["field"]: cfg["skip"] for cfg in self.txt_field_columns_config}


        self.folder_paths = {}
        self.folder_columns = {}
        self.file_extensions = {}
        self.folder_skips = {}
        self.monitors = {}
        self.num_custom_buttons = 3
        self.MAX_CUSTOM_BUTTONS = 20 # Define the maximum number of custom buttons
        
        # Each custom button config now includes a 'txt_source_key'
        # This key maps to a folder path variable in the GUI instance
        # 'None' means no TXT data is read for this button
        # 'Main TXT' maps to self.txt_folder_path
        # 'TXT Source 2' maps to self.txt_folder_path_set2
        # 'TXT Source 3' maps to self.txt_folder_path_set3
        self.custom_button_configs = [
            {"text": "Custom Event 1", "event_text": "Custom Event 1 Triggered", "txt_source_key": "Main TXT"},
            {"text": "Custom Event 2", "event_text": "Custom Event 2 Triggered", "txt_source_key": "None"},
            {"text": "Custom Event 3", "event_text": "Custom Event 3 Triggered", "txt_source_key": "None"},
        ]
        self.custom_buttons = []
        self.button_colors = {
            "Log on": (None, "#90EE90"), "Log off": (None, "#FFB6C1"),
            "Event": (None, "#FFFFE0"), "SVP": (None, "#ADD8E6"),
            "New Day": (None, "#FFFF99")
        }
        # Initialize custom button colors to None
        for i in range(self.MAX_CUSTOM_BUTTONS): self.button_colors[f"Custom {i+1}"] = (None, None)
        

        self.sqlite_enabled = False
        self.sqlite_db_path = None
        self.sqlite_table = "EventLog"

        self.status_var = tk.StringVar()
        self.monitor_status_label = None
        self.db_status_label = None
        self.settings_window_instance = None # Track settings window
        self.custom_inline_editor_window = None # To track the open inline editor

    def create_main_buttons(self, parent_frame):
        '''
        Builds and renders all the buttons in the GUI dynamically: the main event buttons (e.g. "Log on", "SVP") and both sets of custom event buttons
        '''
        # Clear existing buttons in the parent frame
        for widget in parent_frame.winfo_children(): widget.destroy()
        self.custom_buttons = [] # List to hold custom button widgets for reference

        # --- Logging Buttons Area ---
        logging_frame = ttk.LabelFrame(parent_frame, text="Logging Actions", padding=10)
        logging_frame.grid(row=0, column=0, padx=5, pady=(0, 5), sticky="nsew")
        logging_frame.columnconfigure(0, weight=1)

        # --- Custom Events Area ---
        custom_frame = ttk.LabelFrame(parent_frame, text="Custom Events", padding=10)
        custom_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
        num_custom_cols = 2
        for i in range(num_custom_cols):
            custom_frame.columnconfigure(i, weight=1)
            
        custom_frame.bind("<Button-3>", self._show_add_button_context_menu)
        for child in custom_frame.winfo_children():
            child.bind("<Button-3>", self._show_add_button_context_menu)

        other_frame = ttk.LabelFrame(parent_frame, text="Other Actions", padding=10)
        other_frame.grid(row=0, column=1, padx=5, pady=(0, 5), sticky="nsew")
        other_frame.columnconfigure(0, weight=1)

        parent_frame.columnconfigure(0, weight=1)
        parent_frame.columnconfigure(1, weight=1)
        parent_frame.rowconfigure(0, weight=0)
        parent_frame.rowconfigure(1, weight=1)

        # --- Define Button Data ---
        all_buttons_data = {
            "Log on":  {"command_ref": self.log_event, "frame": logging_frame, "tooltip": "Record a 'Log on' marker with current data.", "txt_source_key": "Main TXT"},
            "Log off": {"command_ref": self.log_event, "frame": logging_frame, "tooltip": "Record a 'Log off' marker with current data.", "txt_source_key": "Main TXT"},
            "Event":   {"command_ref": self.log_event, "frame": logging_frame, "tooltip": "Record data from the Main TXT source. Text field is left blank in log.", "txt_source_key": "Main TXT"},
            "SVP":     {"command_ref": self.apply_svp, "frame": other_frame, "tooltip": "Record data and insert latest SVP filename.", "txt_source_key": "Main TXT"},
            "New Day": {"command_ref": self.log_new_day, "frame": other_frame, "tooltip": "Manually trigger the 'New Day' log entry.", "txt_source_key": "Main TXT"},
            "Settings":{"command_ref": self.open_settings, "frame": other_frame, "tooltip": "Open the configuration window."},
            "Sync Excel->DB":{"command_ref": self.sync_excel_to_sqlite_triggered, "frame": other_frame, "tooltip": "Read the Excel log and update the SQLite DB with changes (Requires RecordID column)."}
        }

        # --- Prepare and Sort Custom Button Data ---
        custom_buttons_data = []
        valid_custom_configs = self.custom_button_configs[:self.num_custom_buttons]
        source_order = {"Main TXT": 0, "TXT Source 2": 1, "TXT Source 3": 2, "None": 3}
        try:
            sorted_configs = sorted(valid_custom_configs, key=lambda cfg: source_order.get(cfg.get('txt_source_key', 'None'), 99))
        except Exception as e:
            print(f"Warning: Could not sort custom buttons. Using default order. Error: {e}")
            sorted_configs = valid_custom_configs
            
        for config in sorted_configs:
            button_text = config.get("text", f"Custom Button")
            event_desc = config.get("event_text", f"{button_text} Triggered")
            txt_source = config.get("txt_source_key", "None")
            custom_buttons_data.append({
                "text": button_text, "config": config, "frame": custom_frame,
                "tooltip": f"Log '{event_desc}' with current data. (Source: {txt_source})",
                "txt_source_key": txt_source
            })

        # --- Layout Buttons (2 Stages: Create+Store, then Grid+Assign Command) ---
        buttons_dict = {}

        # Stage 1: Create standard buttons and store references
        for text, data in all_buttons_data.items():
            style_name = f"{text.replace(' ', '').replace('->','')}.TButton"
            _, bg_color = self.button_colors.get(text, (None, None))
            final_style = "TButton"
            if bg_color:
                try:
                    self.style.configure(style_name, background=bg_color)
                    final_style = style_name
                except tk.TclError as e_style:
                    print(f"Warning: Could not apply style color '{bg_color}' for button '{text}'. Error: {e_style}")
            button = ttk.Button(data["frame"], text=text, style=final_style)
            buttons_dict[text] = {"widget": button, "data": data}

        # Create custom buttons and store them
        for button_config_data in custom_buttons_data:
            button_text = button_config_data["text"]
            txt_source = button_config_data.get("txt_source_key", "None")
            style_name = f"{button_text.replace(' ', '')}.TButton"
            
            # --- COLOR LOGIC ---
            # 1. Check for a user-defined color from settings.
            _, user_set_color = self.button_colors.get(button_text, (None, None))
            bg_color = user_set_color if user_set_color else self.source_based_colors.get(txt_source)
            # --- END COLOR LOGIC ---
            
            final_style = "TButton"
            if bg_color:
                try:
                    self.style.configure(style_name, background=bg_color)
                    final_style = style_name
                except tk.TclError as e_style:
                    print(f"Warning: Could not apply style color '{bg_color}' for button '{button_text}'. Error: {e_style}")

            button = ttk.Button(button_config_data["frame"], text=button_text, style=final_style)
            buttons_dict[button_text] = {"widget": button, "data": button_config_data}
            if button_text.startswith("Custom") or button in self.custom_buttons:
                if button not in self.custom_buttons:
                    self.custom_buttons.append(button)
            
            # Add right-click binding using the original index for editing
            config_object = button_config_data.get("config")
            if config_object:
                try:
                    original_index = self.custom_button_configs.index(config_object)
                    button.bind("<Button-3>", lambda e, idx=original_index: self._show_custom_button_context_menu(e, idx))
                except ValueError:
                    print(f"Warning: Could not find config for button '{button_text}' in main list. Context menu may fail.")
                    button.bind("<Button-3>", lambda e, txt=button_text: messagebox.showerror("Binding Error", f"Cannot edit '{txt}'. Configuration is out of sync."))

        # Stage 2: Grid buttons and assign commands
        btn_row_log, btn_row_other, custom_idx = 0, 0, 0
        for text, item in buttons_dict.items():
            button, data = item["widget"], item["data"]
            frame = data["frame"]
            
            cmd = None
            command_ref = data.get("command_ref")
            config_ref = data.get("config")
            txt_source_key = data.get("txt_source_key", "None")

            if command_ref: # Standard buttons
                if text in ["Log on", "Log off", "Event"]: cmd = lambda t=text, b=button, src_key=txt_source_key: self.log_event(t, b, src_key)
                elif text in ["SVP", "New Day"]: cmd = lambda b=button, ref=command_ref, src_key=txt_source_key: ref(b, src_key)
                else: cmd = command_ref
            elif config_ref: # Custom Buttons
                cmd = lambda cfg=config_ref, b=button: self.log_custom_event(cfg, b)
            else: cmd = lambda t=text: print(f"Error: No command or config found for {t}")

            button.config(command=cmd)

            if frame == logging_frame:
                button.grid(row=btn_row_log, column=0, padx=5, pady=4, sticky="ew")
                frame.rowconfigure(btn_row_log, weight=1); btn_row_log += 1
            elif frame == other_frame:
                button.grid(row=btn_row_other, column=0, padx=5, pady=4, sticky="ew")
                frame.rowconfigure(btn_row_other, weight=1); btn_row_other += 1
            elif frame == custom_frame:
                custom_row, custom_col = divmod(custom_idx, num_custom_cols)
                button.grid(row=custom_row, column=custom_col, padx=5, pady=4, sticky="nsew")
                frame.rowconfigure(custom_row, weight=1); custom_idx += 1

            ToolTip(button, data["tooltip"])

        parent_frame.update_idletasks()

    def create_status_indicators(self, parent_frame):
        '''
        Creates the status indicators for monitoring and SQLite connection status.
        This method adds a frame below the main buttons to show the current status of monitoring and SQLite logging.
        '''
        # Create a frame for status indicators
        indicator_frame = ttk.Frame(parent_frame, padding="5 0")
        indicator_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        indicator_frame.columnconfigure(1, weight=0)
        indicator_frame.columnconfigure(3, weight=0)
        indicator_frame.columnconfigure(4, weight=1) # Spacer

        # Create labels for monitoring status
        ttk.Label(indicator_frame, text="Monitoring:", font=("Arial", 9, "bold")).grid(row=0, column=0, sticky=tk.W, padx=(0, 2))
        self.monitor_status_label = ttk.Label(indicator_frame, text="Initializing...", foreground="orange", font=("Arial", 9))
        self.monitor_status_label.grid(row=0, column=1, sticky=tk.W)

        # Create labels for SQLite status
        ttk.Label(indicator_frame, text="SQLite:", font=("Arial", 9, "bold")).grid(row=0, column=2, sticky=tk.W, padx=(15, 2))
        self.db_status_label = ttk.Label(indicator_frame, text="Initializing...", foreground="orange", font=("Arial", 9))
        self.db_status_label.grid(row=0, column=3, sticky=tk.W)
        ttk.Frame(indicator_frame).grid(row=0, column=4) # Spacer

        self.update_db_indicator()

    def sync_excel_to_sqlite_triggered(self):
        '''
        This function is triggered when the "Sync Excel->DB" button is pressed in the GUI.
        Its job is to:
        * Check if SQLite logging is enabled.
        * Validate that the Excel file and SQLite DB paths are properly configured.
        * Disable the button during syncing to avoid repeated clicks.
        * Run the sync operation in a background thread (non-blocking).
        * Re-enable the button and update the status when done.
        '''

        # Check if SQLite logging is enabled
        if not self.sqlite_enabled:
            messagebox.showwarning("Sync Skipped", "SQLite logging is not enabled in Settings.", parent=self.master)
            self.update_status("Sync Error: SQLite disabled.")
            return
        # Validate paths for Excel log file
        if not self.log_file_path or not os.path.exists(self.log_file_path):
            messagebox.showerror("Sync Error", "Excel log file path is not set or the file does not exist.", parent=self.master)
            self.update_status("Sync Error: Excel file path invalid.")
            return
        # Validate SQLite database path
        if not self.sqlite_db_path:
            messagebox.showerror("Sync Error", "SQLite database path is not set.", parent=self.master)
            self.update_status("Sync Error: SQLite DB path missing.")
            return

        sync_button = None
        target_button_text = "Sync Excel->DB"
        try:
            # Searches the GUI for the button labeled "Sync Excel->DB"
            if hasattr(self, 'button_frame') and self.button_frame:
                for child in self.button_frame.winfo_children():
                    if isinstance(child, ttk.LabelFrame):
                        for btn in child.winfo_children():
                            if isinstance(btn, ttk.Button) and btn.cget('text') == target_button_text:
                                sync_button = btn
                                break
                        if sync_button: break
        except Exception: pass

        original_text = None
        if sync_button:
            try:
                if sync_button.winfo_exists():
                    original_text = sync_button['text']
                    sync_button.config(state=tk.DISABLED, text="Syncing...")
            except tk.TclError:
                sync_button = None

        self.update_status("Starting sync from Excel to SQLite...")

        def _sync_worker():
            nonlocal original_text
            success, message = self.perform_excel_to_sqlite_sync()

            self.master.after(0, self.update_status, message)

            if sync_button:
                def re_enable_sync_button(btn=sync_button, txt=original_text):
                    try:
                        if btn and btn.winfo_exists():
                            btn.config(state=tk.NORMAL)
                            if txt: btn.config(text=txt)
                    except tk.TclError: pass
                self.master.after(0, re_enable_sync_button)

        sync_thread = threading.Thread(target=_sync_worker, daemon=True)
        sync_thread.start()

    def perform_excel_to_sqlite_sync(self):
        '''
        This functions ensures the SQLite database reflects the latest data from the Excel log, without overwriting or duplicating unchanged records.
        It performs the following steps:
        * Reads the Excel file specified in self.log_file_path.
        * Parses the data, ensuring date and time formats are handled correctly.
        * Connects to the SQLite database specified in self.sqlite_db_path.
        * Checks if the specified table exists, creating it if necessary.
        * Compares the Excel data against existing records in the SQLite table using RecordID.
        * Inserts new records and updates existing ones based on RecordID.
        * Cleans up resources and returns success status.
        '''
        # Initialization
        excel_file = self.log_file_path
        db_file = self.sqlite_db_path
        db_table = self.sqlite_table
        record_id_column = "RecordID"
        date_col_name = "Date"
        time_col_name = "Time"

        if not excel_file or not db_file or not db_table:
            return False, "Sync Error: Configuration paths or table missing."

        excel_data = {}
        app = None
        wb = None
        sheet = None
        header = None
        df_excel = None

        try:
            # Open the Excel file using xlwings
            app = xw.App(visible=False, add_book=False)

            if not excel_file: raise ValueError("excel_file path is empty")
            wb = app.books.open(excel_file, update_links=False, read_only=True)

            if not wb.sheets: raise ValueError("Workbook has no sheets")
            sheet = wb.sheets[0]

            header_range = sheet.range('A1').expand('right')
            if header_range is None: raise ValueError("Cannot find header range")
            header = header_range.value
            # Check if header is None or includes the RecordID column
            if header is None or record_id_column not in header:
                raise ValueError(f"Column '{record_id_column}' not found in Excel header or header is empty.")

            record_id_col_index = header.index(record_id_column) + 1
            last_row = 1
            try:
                last_row = sheet.range(sheet.api.Rows.Count, record_id_col_index).end('up').row
            except Exception:
                try:
                    max_row = sheet.cells.last_cell.row
                    last_row_A = sheet.range(f'A{max_row}').end('up').row
                    last_row = max(1, last_row_A)
                except Exception:
                    last_row = 1

            if last_row <= 1:
                if wb: wb.close()
                if app: app.quit()
                wb = None; app = None
                return True, "Sync Info: Excel sheet is empty, nothing to sync."
            
            # Read data from the Excel sheet, starting from row 2 to skip header
            data_range = sheet.range((2, 1), (last_row, len(header)))
            # Convert the data range to a DataFrame
            df_excel = pd.DataFrame(data_range.value, columns=header)

            # Parse date and time columns if they exist
            if date_col_name in df_excel.columns:
                df_excel[date_col_name] = pd.to_datetime(df_excel[date_col_name], errors='coerce')
            if time_col_name in df_excel.columns:
                df_excel[time_col_name] = pd.to_numeric(df_excel[time_col_name], errors='coerce')

            # Filter out rows where RecordID is NaN, None, or empty
            if record_id_column in df_excel.columns:
                df_excel[record_id_column] = df_excel[record_id_column].astype(str)
                df_excel[record_id_column] = df_excel[record_id_column].replace({'nan': '', 'None': '', None: ''})
                df_excel = df_excel[df_excel[record_id_column].str.strip() != '']
                df_excel = df_excel.dropna(subset=[record_id_column])
            else:
                raise ValueError(f"'{record_id_column}' column disappeared after initial read.")

            if df_excel.empty:
                if wb: wb.close()
                if app: app.quit()
                wb = None; app = None
                return True, "Sync Info: No valid Excel rows found to sync."

            # Set RecordID as index and convert to dictionary:  Clean Excel data into a usable format keyed by RecordID
            df_excel = df_excel.set_index(record_id_column, drop=False)
            excel_data = df_excel.to_dict('index')

            if wb: wb.close(); wb = None
            if app: app.quit(); app = None

        except xw.XlwingsError as xw_err:
            traceback.print_exc()
            try:
                if wb is not None: wb.close()
            except Exception: pass
            try:
                if app is not None: app.quit()
            except Exception: pass
            return False, f"Sync Error: Excel interaction failed ({type(xw_err).__name__})"

        except Exception as e_excel:
            traceback.print_exc()
            try:
                if wb is not None: wb.close()
            except Exception: pass
            try:
                if app is not None: app.quit()
            except Exception: pass
            return False, f"Sync Error: Reading Excel failed ({type(e_excel).__name__})"

        sqlite_data = {}
        conn_sqlite = None
        db_cols = []
        try:
            # Connect to the SQLite database
            conn_sqlite = sqlite3.connect(db_file, timeout=10)
            conn_sqlite.row_factory = sqlite3.Row

            # Check if the table exists 
            cursor = conn_sqlite.cursor()

            cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name=?;", (db_table,))
            if cursor.fetchone() is None:
                conn_sqlite.close()
                return False, f"Sync Error: SQLite table '{db_table}' does not exist."
            
            # Check if RecordID exists as a column
            cursor.execute(f"PRAGMA table_info([{db_table}])")
            cols_info = cursor.fetchall()
            db_cols = [col['name'] for col in cols_info]
            if record_id_column not in db_cols:
                conn_sqlite.close()
                return False, f"Sync Error: Column '{record_id_column}' not found in SQLite table."

            # Read all rows from the database table into a dictionary (sqlite_data), keyed by RecordID
            quoted_db_cols = ", ".join([f"[{c}]" for c in db_cols])
            cursor.execute(f"SELECT {quoted_db_cols} FROM [{db_table}]")
            rows = cursor.fetchall()
            for row in rows:
                row_dict = dict(row)
                rec_id = str(row_dict.get(record_id_column, '')).strip()
                if rec_id:
                    sqlite_data[rec_id] = row_dict

        except sqlite3.Error as e_sqlite:
            traceback.print_exc()
            if conn_sqlite: conn_sqlite.close()
            return False, f"Sync Error: Reading SQLite failed - {type(e_sqlite).__name__}"
        except Exception as e:
            traceback.print_exc()
            if conn_sqlite: conn_sqlite.close()
            return False, f"Sync Error: Unexpected error reading SQLite - {type(e).__name__}"

        # Compare Excel vs SQLite
        updates_to_apply = [] # Store all update operations to perform
        records_processed = 0 # Total Excel records examined
        records_updated = 0 # How many rows were actually different (and need updating)
        db_cols_set = set(db_cols) # Set of column names in the SQLite table (used for quick lookup)

        # Iterate over each row from Excel
        for rec_id, excel_row_dict in excel_data.items():
            records_processed += 1
            if rec_id in sqlite_data:
                sqlite_row_dict = sqlite_data[rec_id]
                row_needs_update = False

                # Compare Column by Column
                for excel_col_name, excel_val in excel_row_dict.items():
                    if excel_col_name in db_cols_set and excel_col_name != record_id_column:
                        # Get the corresponding value from SQLite
                        sqlite_val = sqlite_row_dict.get(excel_col_name)
                        # Handle Special Formatting (Date/Time)
                        formatted_excel_val = excel_val
                        comparison_val = excel_val

                        if excel_col_name == date_col_name and isinstance(excel_val, pd.Timestamp):
                            try:
                                # Normalize both values to strings
                                formatted_excel_val = excel_val.strftime('%Y-%m-%d')
                                comparison_val = formatted_excel_val
                            except Exception: pass
                        elif excel_col_name == time_col_name and isinstance(excel_val, float):
                            try:
                                total_seconds = int(excel_val * 24 * 60 * 60)
                                hours = (total_seconds % (24 * 3600)) // 3600
                                minutes = (total_seconds % 3600) // 60
                                seconds = total_seconds % 60
                                formatted_excel_val = f"{hours:02}:{minutes:02}:{seconds:02}"
                                comparison_val = formatted_excel_val
                            except Exception: pass

                        str_comparison_val = str(comparison_val) if comparison_val is not None else ""
                        str_sqlite_val = str(sqlite_val) if sqlite_val is not None else ""
                        # Compare: If they are not equal, add the field to the update dictionary
                        if str_comparison_val != str_sqlite_val:
                            updates_to_apply.append((rec_id, excel_col_name, formatted_excel_val))
                            row_needs_update = True

                # If Differences Found, Queue Update
                if row_needs_update:
                    records_updated += 1
        #Applying Updates to SQLite

        # If no updates are needed, return early
        if not updates_to_apply:
            if conn_sqlite: conn_sqlite.close()
            return True, f"Sync complete. No changes detected in {records_processed} Excel rows."

        try:
            cursor = conn_sqlite.cursor()
            updates_by_record = {}
            for rec_id, col, val in updates_to_apply:
                if rec_id not in updates_by_record: updates_by_record[rec_id] = {}
                updates_by_record[rec_id][col] = val

            for rec_id, col_val_dict in updates_by_record.items():
                set_clauses = []
                values = []
                for col, val in col_val_dict.items():
                    set_clauses.append(f"[{col}] = ?")
                    values.append(val)
                if set_clauses:
                    values.append(rec_id)
                    sql_update = f"UPDATE [{db_table}] SET {', '.join(set_clauses)} WHERE [{record_id_column}] = ?"
                    cursor.execute(sql_update, values)
            conn_sqlite.commit()
            if conn_sqlite: conn_sqlite.close()
            return True, f"Sync successful. Updated {len(updates_by_record)} records ({cursor.rowcount} rows affected) in SQLite."

        except sqlite3.Error as e_update:
            traceback.print_exc()
            if conn_sqlite:
                try: conn_sqlite.rollback()
                except Exception: pass
                conn_sqlite.close()
            return False, f"Sync Error: Updating SQLite failed - {type(e_update).__name__}"
        except Exception as e:
            traceback.print_exc()
            if conn_sqlite:
                try: conn_sqlite.rollback()
                except Exception: pass
                conn_sqlite.close()
            return False, f"Sync Error: Unexpected error updating SQLite - {type(e).__name__}"

    def create_status_bar(self, parent_frame):
        '''
        Creates a status bar at the bottom of the main window to display status messages.
        This method initializes a label that will show the current status of the application, such as monitoring status, database connection status, and other messages.
        Arguments:
        * parent_frame: The frame where the status bar will be placed.
        '''
        self.status_var.set("Status: Ready")
        status_bar = ttk.Label(parent_frame, textvariable=self.status_var, style="StatusBar.TLabel", anchor='w')
        status_bar.grid(row=2, column=0, sticky=(tk.W, tk.E), padx=0, pady=(5,0))

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


    def update_db_indicator(self):
        '''
        Updates the SQLite database status indicator label based on the current configuration.
        This method checks if SQLite logging is enabled, verifies the database file path, and updates the label text and color accordingly.
        It also handles cases where the database file is missing or the path is not set.
        '''
        if not hasattr(self, 'db_status_label') or not self.db_status_label:
            return
        if not self.master.winfo_exists():
            return

        # Corrected indentation for the following block
        status_text = "Disabled"
        status_color = "gray"
        if self.sqlite_enabled:
            if self.sqlite_db_path and os.path.exists(self.sqlite_db_path):
                status_text = "Enabled"
                status_color = "green"
            elif self.sqlite_db_path:
                status_text = "File Missing"
                status_color = "#E65C00"
            else:
                status_text = "Path Missing"
                status_color = "#E65C00"
        try:
            self.db_status_label.config(text=status_text, foreground=status_color)
        except tk.TclError:
            pass # Widget might be destroyed


    # --- Logging Actions (using threading) ---

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
        if event_type in ["Log on", "Log off"]:
            event_text_for_excel = f"{event_type} event occurred"
        elif event_type == "Event":
            skip_files = True
            event_text_for_excel = ""
            
        self._perform_log_action(event_type=event_type,
                                 event_text_for_excel=event_text_for_excel,
                                 skip_latest_files=skip_files,
                                 triggering_button=button_widget,
                                 txt_source_key=txt_source_key)

    def log_custom_event(self, config, button_widget):
        '''
        This function is called when a custom event button is pressed.
        It retrieves the button text and event text from the configuration, then calls _perform_log_action to log the event.
        Arguments:
        * config: The configuration dictionary for the custom button, containing "text" and "event_text".
        * button_widget: The button widget that was pressed, used to temporarily disable it during processing.
        * txt_source_set: The set number (1 or 2) indicating which TXT source to use for logging.
        '''
        button_text = config.get("text", "Unknown Custom")
        event_text_for_excel = config.get("event_text", f"{button_text} Triggered")
        txt_source_key = config.get("txt_source_key", "None")
        
        self._perform_log_action(event_type=button_text,
                                 event_text_for_excel=event_text_for_excel,
                                 triggering_button=button_widget,
                                 txt_source_key=txt_source_key)

    def log_new_day(self, button_widget=None, txt_source_key="Main TXT"):
        '''
        This function is called when the "New Day" button is pressed.
        It logs a "New Day" event by calling _perform_log_action with the appropriate parameters.
        Arguments:
        * button_widget: The button widget that was pressed, used to temporarily disable it during processing.
        '''
        self._perform_log_action(event_type="New Day",
                                 event_text_for_excel="New Day",
                                 triggering_button=button_widget,
                                 txt_source_key=txt_source_key)

    def apply_svp(self, button_widget, txt_source_key="Main TXT"):
        '''
        This function is called when the "Apply SVP" button is pressed.
        It checks if the necessary configurations are set (log file, TXT folder, SVP folder path),
        and if so, it calls _perform_log_action to log the SVP event.
        Arguments:
        * button_widget: The button widget that was pressed, used to temporarily disable it during processing.
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

        self._perform_log_action(event_type="SVP",
                                 event_text_for_excel="SVP applied",
                                 svp_specific_handling=True,
                                 triggering_button=button_widget,
                                 txt_source_set=1) # SVP uses primary TXT source

    def _perform_log_action(self, event_type, event_text_for_excel, skip_latest_files=False, svp_specific_handling=False, triggering_button=None, txt_source_key="Main TXT"):
        '''This is the main entry point for handling an event (e.g., button press). 
        It collects all necessary data (from TXT files and folder monitors), 
        then logs the event to Excel and/or SQLite in a background thread.
        
        Arguments:
        * event_type: The label of the event, e.g., "Log on", "Event", "Custom Event 1".
        * event_text_for_excel: The actual text that goes into the "Event" column in Excel.
        * skip_latest_files: Whether to skip checking monitored folders (used for basic events).
        * svp_specific_handling: Enables special logic if the event is "SVP".
        * triggering_button: The button that was pressed (used to temporarily disable it).
        * txt_source_set: Specifies which TXT file set (1 or 2) to use for extracting data.

        Workflow explanation:
        When you click a button in the GUI:
        * _perform_log_action() is triggered.
        * It calls insert_txt_data() to extract latest TXT data.
        * Then it appends other folder-monitor-based data.
        * Then it logs everything to:
        * Excel: with optional row color
        * SQLite: if enabled
        * Updates the GUI status with success/failure feedback.



        '''
        self.update_status(f"Processing '{event_type}'...")

        original_text = None
        # Disable the button if it exists and is a ttk.Button
        if triggering_button and isinstance(triggering_button, ttk.Button):
            try:
                if triggering_button.winfo_exists():
                    original_text = triggering_button['text']
                    triggering_button.config(state=tk.DISABLED, text="Working...")
            except tk.TclError:
                triggering_button = None

        # Define a background thread to avoid blocking the GUI
        def _worker_thread_func():
            nonlocal original_text 
            # Prepares an empty data row with a RecordID
            row_data = {}
            excel_success = False
            sqlite_logged = False
            excel_save_exception = None
            sqlite_save_exception_type = None
            status_msg = f"'{event_type}' processed with errors."

            record_id = str(uuid.uuid4())
            row_data['RecordID'] = record_id

            try:
                event_col_name = self.txt_field_columns.get("Event", "Event")
                row_data["EventType"] = event_type
                if event_text_for_excel is not None:
                    row_data[event_col_name] = event_text_for_excel

                if txt_source_key and txt_source_key != "None":
                    source_folder_path = None
                    if txt_source_key == "Main TXT":
                        source_folder_path = self.txt_folder_path
                    elif txt_source_key == "TXT Source 2":
                        source_folder_path = self.txt_folder_path_set2
                    elif txt_source_key == "TXT Source 3":
                        source_folder_path = self.txt_folder_path_set3
                    
                    if source_folder_path:
                        try:
                            txt_data = self._get_txt_data_from_source(source_folder_path)
                            if txt_data: row_data.update(txt_data)
                        except Exception as e_txt:
                            self.master.after(0, lambda e=e_txt: messagebox.showerror("Error", f"Failed to read TXT data from {txt_source_key}:\n{e}", parent=self.master))
                
                if not skip_latest_files:
                    try:
                        latest_files_data = self.get_latest_files_data()
                        if latest_files_data: row_data.update(latest_files_data)
                    except Exception as e_files:
                        self.master.after(0, lambda e=e_files: messagebox.showerror("Error", f"Failed to get latest file data:\n{e}", parent=self.master))

                # Adds SVP file info if applicable
                if svp_specific_handling: # SVP logic also global
                    svp_folder_path = self.folder_paths.get("SVP")
                    svp_col_name = self.folder_columns.get("SVP", "SVP")
                    if svp_folder_path and svp_col_name:
                        latest_svp_file = folder_cache.get("SVP")
                        row_data[svp_col_name] = latest_svp_file if latest_svp_file else "N/A"
                    elif svp_col_name:
                        row_data[svp_col_name] = "Config Error"
                
                # If row_data has data, proceed to save to Excel and SQLite (save_to_excel()) and log_to_sqlite())
                if row_data:
                    # Get the color for the row based on the event type
                    color_tuple = self.button_colors.get(event_type, (None, None))
                    row_color_for_excel = color_tuple[1] if isinstance(color_tuple, tuple) and len(color_tuple) > 1 else None

                    excel_data = {k: v for k, v in row_data.items() if k != 'EventType'}

                    try:
                        if not self.log_file_path: excel_save_exception = ValueError("Excel path missing")
                        elif not os.path.exists(self.log_file_path): excel_save_exception = FileNotFoundError("Excel file missing")
                        else:
                            # Save the data to Excel
                            self.save_to_excel(excel_data, row_color=row_color_for_excel)
                            excel_success = True
                    except Exception as e_excel:
                        excel_save_exception = e_excel
                        traceback.print_exc()
                        self.master.after(0, lambda e=e_excel: messagebox.showerror("Excel Error", f"Failed to save to Excel:\n{e}", parent=self.master))
                    
                    # If Excel save was successful, log to SQLite
                    sqlite_logged, sqlite_save_exception_type = self.log_to_sqlite(row_data)

                    # Constructs a status message to show whether Excel and SQLite logging succeeded or failed.
                    status_parts = []
                    if excel_success: status_parts.append("Excel: OK")
                    elif excel_save_exception: status_parts.append(f"Excel: Fail ({type(excel_save_exception).__name__})")
                    else: status_parts.append("Excel: Fail (Check Path)")

                    if self.sqlite_enabled:
                        if sqlite_logged: status_parts.append("SQLite: OK")
                        else:
                            err_detail = f" ({sqlite_save_exception_type})" if sqlite_save_exception_type else ""
                            status_parts.append(f"SQLite: Fail{err_detail}")

                    if not excel_success and not (self.sqlite_enabled and sqlite_logged):
                        status_msg = f"'{event_type}' log FAILED. " + ", ".join(status_parts)
                    elif not status_parts: status_msg = f"Error logging '{event_type}' - No status."
                    else: status_msg = f"'{event_type}' logged. " + ", ".join(status_parts) + "."
                else:
                    status_msg = f"'{event_type}' pressed, but no data was collected/generated."

            except Exception as thread_ex:
                traceback.print_exc()
                status_msg = f"'{event_type}' - Unexpected thread error: {thread_ex}"
                self.master.after(0, lambda e=thread_ex: messagebox.showerror("Thread Error", f"Critical error during logging action '{event_type}':\n{e}", parent=self.master))

            finally:
                self.master.after(0, self.update_status, status_msg)

            # Re-enables the button if it was disabled
                if triggering_button and isinstance(triggering_button, ttk.Button):
                    def re_enable_button(btn=triggering_button, txt=original_text):
                        try:
                            if btn and btn.winfo_exists():
                                btn.config(state=tk.NORMAL)
                                if txt:
                                    btn.config(text=txt)
                        except tk.TclError: pass
                    self.master.after(0, re_enable_button)

        log_thread = threading.Thread(target=_worker_thread_func, daemon=True)
        log_thread.start()

    # --- Data Fetching Methods (Refactored to support multiple TXT sources) ---
    def _get_txt_data_from_source(self, folder_path):
        """
        Reads and parses data from the latest TXT file in the specified folder.
        Uses global self.txt_field_columns and self.txt_field_skips for mapping.
        Returns a dictionary of parsed data or empty dict if no data/errors.
        """
        row_data = {}
        current_dt = datetime.datetime.now()

        # Always use PC time, overriding any file-based time logic
        use_pc_time = True
        reason_for_pc_time = "Code configured to always use PC time."

        latest_txt_file_path = None
        if folder_path and os.path.exists(folder_path):
            latest_txt_file_path = self.find_latest_file_in_folder(folder_path, ".txt")
        
        temp_txt_data = {}

        # Even though we're using PC time for Date/Time, we still attempt to
        # read other data from the TXT file if it exists and is readable.
        if latest_txt_file_path:
            try:
                lines = []
                encodings_to_try = ['utf-8', 'latin-1', 'cp1252']
                read_success = False
                for enc in encodings_to_try:
                    try:
                        for attempt in range(3):
                            try:
                                with open(latest_txt_file_path, "r", encoding=enc) as file:
                                    lines = file.readlines()
                                read_success = True
                                break
                            except IOError:
                                if attempt < 2:
                                    time.sleep(0.1)
                                    continue
                                else:
                                    raise
                        if read_success:
                            break
                    except UnicodeDecodeError:
                        continue
                    except Exception:
                        lines = []
                        break

                if lines:
                    latest_line_str = lines[-1].strip()
                    latest_line_parts = latest_line_str.split(",")
                    
                    # Iterate through the ordered config for TXT fields
                    for i, field_config in enumerate(self.txt_field_columns_config):
                        field_key = field_config["field"]
                        excel_col = field_config["column_name"]
                        skip_field = field_config["skip"]

                        if excel_col and not skip_field:
                            # If it's a Date or Time field, we will override it with PC time later,
                            # so we can skip reading it from the file here.
                            if field_key in ["Date", "Time"]:
                                continue

                            # The 'Event' field is typically for event descriptions triggered by buttons,
                            # not usually directly parsed from TXT data in a specific position.
                            if field_key == "Event" and (field_key not in ["Date", "Time", "KP", "DCC", "Line name", "Latitude", "Longitude", "Easting", "Northing"]):
                                continue

                            try:
                                if i < len(latest_line_parts):
                                    value = latest_line_parts[i].strip()
                                    temp_txt_data[excel_col] = value
                                else:
                                    temp_txt_data[excel_col] = None # Field not found at expected index
                            except Exception:
                                temp_txt_data[excel_col] = None

            except Exception as e:
                # Log or handle unexpected errors during file processing, but still use PC time
                print(f"Warning: Error reading TXT file for other data: {e}. Still using PC time for Date/Time.")


        # When using PC time, populate configured Date/Time columns
        date_col = None
        time_col = None
        skip_date = False
        skip_time = False

        # Find configured Date/Time columns and skip status
        for cfg in self.txt_field_columns_config:
            if cfg["field"] == "Date":
                date_col = cfg["column_name"]
                skip_date = cfg["skip"]
            elif cfg["field"] == "Time":
                time_col = cfg["column_name"]
                skip_time = cfg["skip"]

        if date_col and not skip_date:
            row_data[date_col] = current_dt.strftime("%Y-%m-%d")
        if time_col and not skip_time:
            row_data[time_col] = current_dt.strftime("%H:%M:%S")

        # Add other data that might have been partially parsed from the file
        for col, val in temp_txt_data.items():
            if col not in row_data: # Don't overwrite PC date/time if already set
                row_data[col] = val
        
        # You might want to log the reason for using PC time for debugging
        # print(f"Using PC time: {reason_for_pc_time}")

        return row_data
    def get_latest_files_data(self): # This is global for monitored folders
        '''Collects the latest files from all monitored folders and returns a dictionary of column names to file paths.
        Returns:
        * A dictionary where keys are column names (from folder_columns) and values are the latest file paths.
        '''

        latest_files = {}
        for folder_name, folder_path in self.folder_paths.items():
            if not folder_path or self.folder_skips.get(folder_name, False): continue
            latest_file = folder_cache.get(folder_name)
            column_name = self.folder_columns.get(folder_name)
            if not column_name: continue
            if latest_file: latest_files[column_name] = latest_file
            else: latest_files[column_name] = "N/A"
        return latest_files

    def find_latest_file_in_folder(self, folder_path, extension=".txt"):
        '''Finds the most recent file with the specified extension in the given folder.
            Arguments:
            * folder_path: The path to the folder where files are searched.
            * extension: The file extension to look for (default is ".txt").
            Returns:
            * The path to the most recent file with the specified extension, or None if no such file exists.
        '''
        try:
            files = []
            ext_lower = extension.lower()
            for f in os.listdir(folder_path):
                f_path = os.path.join(folder_path, f)
                try:
                    if os.path.isfile(f_path) and f.lower().endswith(ext_lower): files.append(f_path)
                except OSError: continue
            return max(files, key=os.path.getmtime) if files else None
        except FileNotFoundError: return None
        except Exception: return None

    def save_to_excel(self, row_data, row_color=None, next_row=None):
        '''Saves the provided row_data to the specified Excel log file.
        Arguments:
        * row_data: A dictionary containing the data to log, where keys are column names and values are the data to insert.
        * row_color: An optional RGB tuple (R, G, B) to apply as the background color for the row in Excel.
        * next_row: The row number where the data should be written. If None, finds the next empty row automatically.
        Returns:
        * None if successful, raises an exception if there is an error.
        '''
        # Check Excel file path and existence
        if not self.log_file_path:
            raise ValueError("Excel log file path is missing.")
        if not os.path.exists(self.log_file_path):
            raise FileNotFoundError(f"Excel log file not found: {self.log_file_path}")

        app = None
        workbook = None
        opened_new_app = False
        opened_workbook = False

        try:
            # Try to connect to an existing Excel instance first
            try:
                # Check if an Excel app is already running
                app = xw.apps.active
                if app is None: raise Exception("No active Excel instance")
            except Exception:
                # If no instance is active, create a new invisible one
                app = xw.App(visible=False)
                opened_new_app = True

            # Normalize paths for reliable comparison
            target_norm_path = os.path.normcase(os.path.abspath(self.log_file_path))
            
            # Check if the workbook is already open in this Excel instance
            for wb in app.books:
                try:
                    # Check if the workbook is already open by comparing normalized paths
                    if os.path.normcase(os.path.abspath(wb.fullname)) == target_norm_path:
                        workbook = wb
                        break
                except Exception:
                    continue # Ignore workbooks that might cause errors (e.g., protected)

            # If the workbook was not found open, open it
            if workbook is None:
                workbook = app.books.open(self.log_file_path, read_only=False)
                opened_workbook = True

            # Check if the workbook has at least one sheet
            sheet = workbook.sheets[0]
            # Get the header row (A1)
            header_range_obj = sheet.range("A1").expand("right")
            header_values = header_range_obj.value
            if not header_values or not any(h is not None for h in header_values):
                raise ValueError("Excel header row is missing or empty.")
            
            record_id_col_name = "RecordID"
            if record_id_col_name not in header_values:
                raise ValueError(f"Excel header missing required '{record_id_col_name}' column.")

            header_map_lower = {str(h).lower(): i + 1 for i, h in enumerate(header_values) if h is not None}
            last_header_col_index = max(header_map_lower.values()) if header_map_lower else 1

            if next_row is None:
                try:
                    # Find the last used row in the first column and add 1
                    last_row_a = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
                    next_row = last_row_a + 1
                    if sheet.range(f'A{last_row_a}').value is None: # Handle case where sheet is empty
                         next_row = 2
                except Exception:
                    next_row = 2 # Fallback for completely empty sheets

            # --- Write data to cells ---
            written_cols = []
            for col_name, value in row_data.items():
                col_name_lower = str(col_name).lower()
                if col_name_lower in header_map_lower:
                    col_index = header_map_lower[col_name_lower]
                    try:
                        target_cell = sheet.range(next_row, col_index)
                        if col_name == record_id_col_name:
                            target_cell.number_format = '@' # Ensure RecordID is treated as text
                        target_cell.value = value
                        written_cols.append(col_index)
                    except Exception as e_write:
                        print(f"Warning: Could not write to column '{col_name}'. Error: {e_write}")

            # If row_color is specified, apply it to the entire row
            if row_color and written_cols:
                try:
                    target_range = sheet.range((next_row, 1), (next_row, last_header_col_index))
                    target_range.color = row_color
                except Exception as e_color:
                     print(f"Warning: Could not apply color to row. Error: {e_color}")

            # --- CRITICAL SAVE OPERATION ---
            try:
                workbook.save()
            except Exception as e_save:
                # This is the most likely point of failure causing the "Save As" conflict.
                # We provide a specific, helpful error message.
                traceback.print_exc()
                error_msg = (
                    "Failed to save the Excel file. This is usually because the file is locked.\n\n"
                    "1. Check Task Manager for any lingering 'EXCEL.EXE' processes and end them.\n"
                    "2. Ensure you have permissions to write to the file.\n\n"
                    f"(Details: {e_save})"
                )
                # Show the error in the main thread's GUI
                self.master.after(0, lambda: messagebox.showerror("Excel Save Conflict", error_msg, parent=self.master))
                # Re-raise the exception to stop the current logging process
                raise IOError(f"Failed to save Excel workbook: {e_save}")

        except Exception as e:
            traceback.print_exc()
            # Ensure the exception is propagated so the calling thread knows it failed
            raise e
        finally:
            # This cleanup is vital. It ensures we only close what we opened.
            if workbook is not None and opened_workbook:
                try:
                    workbook.close(save_changes=False)
                except Exception:
                    pass # Ignore errors on close
            if app is not None and opened_new_app:
                try:
                    app.quit()
                except Exception:
                    pass # Ignore errors on quit

    def log_to_sqlite(self, row_data):
        '''Logs the provided row_data to the SQLite database.
            Arguments:
            * row_data: A dictionary containing the data to log, where keys are column names and values are the data to insert.
            Returns:
            * success: True if logging was successful, False otherwise.
            * error_type: A string indicating the type of error if logging failed, or None if successful.
        '''
        success = False
        error_type = None

        # Check if SQLite logging is enabled and if the necessary configuration is present
        if not self.sqlite_enabled:
            return False, "Disabled"

        if not self.sqlite_db_path or not self.sqlite_table:
            self.master.after(0, self.update_status, "SQLite Log Error: DB Path or Table Name missing in settings.")
            return False, "ConfigurationMissing"

        conn = None
        cursor = None
        try:
            # Connect to the SQLite database
            conn = sqlite3.connect(self.sqlite_db_path, timeout=5)
            cursor = conn.cursor()

            table_columns_info = {}
            try:
                # Retrieve the table schema to get column names
                pragma_sql = f"PRAGMA table_info([{self.sqlite_table}]);"
                cursor.execute(pragma_sql)
                results = cursor.fetchall()
                if not results: raise sqlite3.OperationalError(f"No such table: {self.sqlite_table}")
                table_columns_info = {row[1].lower(): row[1] for row in results}
            except sqlite3.Error as e_meta:
                self.master.after(0, self.update_status, f"SQLite Log Error: Cannot get columns for '{self.sqlite_table}'")
                error_type = f"MetadataError_{type(e_meta).__name__}"
                raise e_meta 
            
            # Prepare the data to insert into the SQLite table
            data_to_insert = {}
            provided_keys_lower = {str(k).lower(): k for k, v in row_data.items()}
            for lower_key, actual_key in provided_keys_lower.items():
                if lower_key in table_columns_info:
                    db_col_name = table_columns_info[lower_key]
                    data_to_insert[db_col_name] = row_data[actual_key]

            if not data_to_insert:
                self.master.after(0, self.update_status, "SQLite Log Info: No data matched DB columns.")
                success = True
                error_type = None
                return success, error_type

            cols = list(data_to_insert.keys())
            placeholders = ", ".join(["?"] * len(cols))
            col_name_string = ", ".join([f"[{c}]" for c in cols])
            sql_insert = f"INSERT INTO [{self.sqlite_table}] ({col_name_string}) VALUES ({placeholders})"
            values = [data_to_insert[c] for c in cols]
            
            # Execute the insert statement 
            cursor.execute(sql_insert, values)
            conn.commit()
            success = True
            error_type = None

        except sqlite3.OperationalError as op_err:
            error_message = str(op_err); error_type = "OperationalError"
            self.master.after(0, self.update_status, f"SQLite Log Error: {error_message}")
            if conn:
                try: conn.rollback()
                except Exception: pass
            if "no such table" in error_message.lower():
                error_type = "NoSuchTable"
            elif "has no column named" in error_message.lower():
                error_type = "NoSuchColumn"
            elif "database is locked" in error_message.lower():
                error_type = "DatabaseLocked"
            self.master.after(0, lambda em=error_message, et=error_type: self.show_sqlite_error_message(em, et))
            success = False

        except sqlite3.Error as ex:
            error_message = str(ex); error_type = type(ex).__name__
            self.master.after(0, self.update_status, f"SQLite Log Error: {error_message}")
            if conn:
                try: conn.rollback()
                except Exception: pass
            self.master.after(0, lambda et=error_type, em=error_message: messagebox.showerror("SQLite Error", f"Failed to log to SQLite database.\nType: {et}\nMessage: {em}", parent=self.master))
            success = False

        except Exception as e:
            error_message = str(e); error_type = type(e).__name__
            self.master.after(0, self.update_status, f"SQLite Log Error: Unexpected error ({error_type}).")
            if conn:
                try: conn.rollback()
                except Exception: pass
            self.master.after(0, lambda em=error_message: messagebox.showerror("Application Error", f"Unexpected error during SQLite logging:\n{em}", parent=self.master))
            success = False

        finally:
            if cursor:
                try: cursor.close()
                except Exception: pass
            if conn:
                try: conn.close()
                except Exception: pass
        return success, error_type

    def show_sqlite_error_message(self, error_message, error_type):
        ''' Displays an error message box for SQLite errors with specific handling based on the error type.
            Arguments:
            * error_message: The error message string from the SQLite operation.
            * error_type: A string indicating the type of error (e.g., "NoSuchTable", "NoSuchColumn", "DatabaseLocked", etc.).
        '''

        parent_window = self.settings_window_instance if (hasattr(self, 'settings_window_instance') and self.settings_window_instance and self.settings_window_instance.winfo_exists()) else self.master

        if error_type == "NoSuchTable":
            messagebox.showerror("SQLite Error", f"Table '{self.sqlite_table}' not found.\nPlease check table name or create table.\nDB: {self.sqlite_db_path}", parent=parent_window)
        elif error_type == "NoSuchColumn":
            try:
                missing_col = error_message.split("column named")[-1].strip().split(":")[0].strip()
                missing_col = missing_col.strip("'\"[]")
            except Exception:
                missing_col = "[unknown]"
            messagebox.showerror("SQLite Error", f"Column '{missing_col}' not found in table '{self.sqlite_table}'.\nCheck Settings (TXT Columns / Folder Columns) vs. DB table structure.\n\n(Original error: {error_message})", parent=parent_window)
        elif error_type == "DatabaseLocked":
            messagebox.showerror("SQLite Error", f"Database file is locked.\nAnother program might be using it.\nDB: {self.sqlite_db_path}\n\n(Original error: {error_message})", parent=parent_window)
        else:
            messagebox.showerror("SQLite Operational Error", f"Error interacting with database:\n{error_message}", parent=parent_window)

    def save_settings(self):
        '''Saves the current settings to the JSON file. Cleans up button colors for custom buttons that no longer exist.'''
        colors_to_save = {}
        for key, (_, color_hex) in self.button_colors.items():
            if color_hex: colors_to_save[key] = color_hex
        settings = {
            "log_file_path": self.log_file_path,
            "txt_folder_path": self.txt_folder_path,
            "txt_folder_path_set2": self.txt_folder_path_set2,
            "txt_folder_path_set3": self.txt_folder_path_set3,
            "txt_field_columns_config": self.txt_field_columns_config, # Save the new ordered list
            "folder_paths": self.folder_paths, "folder_columns": self.folder_columns,
            "file_extensions": self.file_extensions, "folder_skips": self.folder_skips,
            "num_custom_buttons": self.num_custom_buttons, "custom_button_configs": self.custom_button_configs,
            "button_colors": colors_to_save, "sqlite_enabled": self.sqlite_enabled,
            "sqlite_db_path": self.sqlite_db_path, "sqlite_table": self.sqlite_table,
        }
        try:
            with open(self.settings_file, 'w') as f: json.dump(settings, f, indent=4)
            self.update_status("Settings saved.")
        except Exception as e:
            messagebox.showerror("Save Error", f"Could not save settings to {self.settings_file}:\n{e}", parent=self.master)
            self.update_status("Error saving settings.")

    def load_settings(self):
        '''Loads settings from the JSON file and updates the GUI variables accordingly.'''
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r') as f: settings = json.load(f)
                self.log_file_path = settings.get("log_file_path")
                
                self.txt_folder_path = settings.get("txt_folder_path")
                self.txt_folder_path_set2 = settings.get("txt_folder_path_set2")
                self.txt_folder_path_set3 = settings.get("txt_folder_path_set3")

                # Load new txt_field_columns_config, with fallback for old settings
                loaded_txt_config = settings.get("txt_field_columns_config")
                if loaded_txt_config:
                    self.txt_field_columns_config = loaded_txt_config
                else: # Fallback for old settings structure
                    old_txt_cols = settings.get("txt_field_columns", {"Event": "Event"})
                    old_txt_skips = settings.get("txt_field_skips", {})
                    # Reconstruct the ordered list from old dicts, prioritizing new fields
                    new_config = []
                    default_order_fields = ["Date", "Time", "KP", "DCC", "Line name", "Latitude", "Longitude", "Easting", "Northing", "Event"]
                    for field in default_order_fields:
                        new_config.append({
                            "field": field,
                            "column_name": old_txt_cols.get(field, field),
                            "skip": old_txt_skips.get(field, False)
                        })
                    # Add any custom fields that might have been saved in old structure but aren't default
                    # This loop needs to be careful not to re-add default fields if their name was changed.
                    # A better way is to iterate over all existing keys in old_txt_cols and add them if not already added.
                    for field_key, col_name in old_txt_cols.items():
                        # Check if this field_key already exists in our new_config based on its 'field' value
                        if not any(cfg['field'] == field_key for cfg in new_config):
                            new_config.append({
                                "field": field_key,
                                "column_name": col_name,
                                "skip": old_txt_skips.get(field_key, False)
                            })
                    self.txt_field_columns_config = new_config
                
                # Re-derive these for backward compatibility
                self.txt_field_columns = {cfg["field"]: cfg["column_name"] for cfg in self.txt_field_columns_config}
                self.txt_field_skips = {cfg["field"]: cfg["skip"] for cfg in self.txt_field_columns_config}


                self.folder_paths.clear(); self.folder_paths.update(settings.get("folder_paths", {}))
                self.folder_columns.clear(); self.folder_columns.update(settings.get("folder_columns", {}))
                self.file_extensions.clear(); self.file_extensions.update(settings.get("file_extensions", {}))
                self.folder_skips.clear(); self.folder_skips.update(settings.get("folder_skips", {}))
                self.num_custom_buttons = settings.get("num_custom_buttons", 3)
                loaded_configs = settings.get("custom_button_configs", [])
                
                updated_custom_configs = []
                for i in range(self.num_custom_buttons):
                    if i < len(loaded_configs):
                        config = loaded_configs[i]
                    else:
                        config = {"text": f"Custom {i+1}", "event_text": f"Custom {i+1} Event"}
                    config["txt_source_key"] = config.get("txt_source_key", "None")  
                    updated_custom_configs.append(config)
                self.custom_button_configs = updated_custom_configs

                loaded_colors = settings.get("button_colors", {})
                default_colors = {"Log on": (None, "#90EE90"), "Log off": (None, "#FFB6C1"), "Event": (None, "#FFFFE0"), "SVP": (None, "#ADD8E6"), "New Day": (None, "#FFFF99")}
                self.button_colors = default_colors
                for config in self.custom_button_configs:
                    btn_text = config.get("text")
                    if btn_text and btn_text not in self.button_colors: self.button_colors[btn_text] = (None, None)
                for key, color_hex in loaded_colors.items():
                    if color_hex: self.button_colors[key] = (None, color_hex)
                self.sqlite_enabled = settings.get("sqlite_enabled", False)
                self.sqlite_db_path = settings.get("sqlite_db_path")
                self.sqlite_table = settings.get("sqlite_table", "EventLog")
                self.update_status("Settings loaded.")
            else:
                self.update_status("Settings file not found. Using defaults.")
        except json.JSONDecodeError as e:
            messagebox.showerror("Load Error", f"Settings file '{self.settings_file}' has invalid format:\n{e}\n\nUsing default settings.", parent=self.master)
            self.update_status("Error loading settings: Invalid format."); self.init_variables()
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Load Error", f"Could not load settings from {self.settings_file}:\n{e}\n\nUsing default settings.", parent=self.master)
            self.update_status("Error loading settings."); self.init_variables()
        finally:
            if hasattr(self, 'button_frame') and self.button_frame: self.update_custom_buttons()
            if hasattr(self, 'db_status_label') and self.db_status_label: self.update_db_indicator()

    # --- Settings Window Interaction ---
    def open_settings(self):
        '''Open the settings window. If it already exists, bring it to the front.'''

        # Check if the settings window already exists and is open
        if hasattr(self, 'settings_window_instance') and self.settings_window_instance and self.settings_window_instance.winfo_exists():
            self.settings_window_instance.lift(); self.settings_window_instance.focus_set()
        else:
            settings_top_level = tk.Toplevel(self.master)
            if self.custom_inline_editor_window and self.custom_inline_editor_window.winfo_exists():
                settings_top_level.transient(self.custom_inline_editor_window)
                settings_top_level.grab_set()
            else:
                settings_top_level.transient(self.master)
                settings_top_level.grab_set()

            self.settings_window_instance = settings_top_level
            settings_gui = SettingsWindow(settings_top_level, self)
            settings_gui.load_settings()
            self.master.wait_window(settings_top_level)
            try:
                del self.settings_window_instance
            except AttributeError: pass

    def startup_settings(self):
        '''Open settings by default in the startup of the app'''

        self.open_settings()

    def update_custom_buttons(self):
        '''Update the custom buttons in the main GUI based on current settings.'''

        if hasattr(self, 'button_frame') and self.button_frame:
            self.create_main_buttons(self.button_frame)
            self.master.update_idletasks()
        else: pass

    # --- Monitoring ---
    def start_monitoring(self):
        '''Function to read the last version of a file in several folders'''

        for name, monitor_observer in list(self.monitors.items()):
            try:
                if monitor_observer.is_alive():
                    monitor_observer.stop()
            except Exception: pass
        self.monitors.clear(); folder_cache.clear()

        count = 0; monitoring_active = False
        
        monitored_sources_data = {
            "Main TXT File": self.txt_folder_path,
            "TXT Source 2": self.txt_folder_path_set2,
            "TXT Source 3": self.txt_folder_path_set3
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
                success = self.start_folder_monitoring(folder_name, folder_path, file_extension)
                if success: count += 1; monitoring_active = True
            elif self.folder_skips.get(folder_name): pass
            elif folder_path: pass

        self.update_status(f"Monitoring {count} active folders.")

        if hasattr(self, 'monitor_status_label') and self.monitor_status_label:
            if monitoring_active: self.monitor_status_label.config(text="Active", foreground="green")
            else: self.monitor_status_label.config(text="Inactive", foreground="red")
        self.update_db_indicator()

    def start_folder_monitoring(self, folder_name, folder_path, file_extension):
        '''Start monitoring a specific folder for changes in files with a given extension.
           Arguments:
           * folder_name: Name of the folder to monitor.
           * folder_path: Full path to the folder to monitor.
           * file_extension: File extension to filter files (e.g., ".txt"). If empty, monitors all files.
           
           Returns True if monitoring started successfully, False otherwise.
        '''
        try: os.listdir(folder_path)
        except Exception: return False
        try:
            event_handler = FolderMonitor(folder_path, folder_name, self, file_extension)
            observer = PollingObserver(timeout=1)
            observer.schedule(event_handler, folder_path, recursive=False)
            observer.start()
            self.monitors[folder_name] = observer
            threading.Thread(target=event_handler.update_latest_file, daemon=True).start()
            return True
        except Exception: return False

    # --- Automatic New Day Scheduling ---
    def schedule_new_day(self):
        '''Schedule the next "New Day" log to trigger at midnight.'''

        now = datetime.datetime.now()
        tomorrow = now.date() + datetime.timedelta(days=1)
        midnight = datetime.datetime.combine(tomorrow, datetime.time.min)
        time_until_midnight_ms = int((midnight - now).total_seconds() * 1000)
        trigger_delay_ms = time_until_midnight_ms + 1000

        self._new_day_timer_id = self.master.after(trigger_delay_ms, self.trigger_new_day)

    def trigger_new_day(self):
        '''Trigger the "New Day" log manually. This can be called automatically at midnight.'''

        self.log_new_day(button_widget=None, txt_source_key="Main TXT")
        # After logging the new day, reschedule the next trigger
        self.schedule_new_day()

    # --- Inline Custom Button Editor ---
    def _show_custom_button_context_menu(self, event, button_index):
        """Shows a context menu for the clicked custom button."""
        # Check if the right-click was on the frame or a button inside
        if str(event.widget).startswith(str(event.widget.winfo_toplevel()) + ".!frame.!frame.!labelframe"): # Check if it's the custom_frame itself
            # This is a right-click on the frame background or empty space
            self._show_add_button_context_menu(event)
            return

        # If an inline editor is already open, focus it instead of opening another or a context menu
        if self.custom_inline_editor_window and self.custom_inline_editor_window.winfo_exists():
            self.custom_inline_editor_window.lift()
            self.custom_inline_editor_window.focus_set()
            return

        context_menu = tk.Menu(self.master, tearoff=0)
        current_button_text = self.custom_button_configs[button_index].get("text", f"Custom {button_index+1}")
        context_menu.add_command(label=f"Edit \"{current_button_text}\" Settings...",
                                 command=lambda: self._edit_custom_button_inline(button_index))
        
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

    def _add_new_custom_button(self):
        """Adds a new custom button configuration and updates the GUI."""
        if self.num_custom_buttons < self.MAX_CUSTOM_BUTTONS:
            self.num_custom_buttons += 1
            new_button_idx = self.num_custom_buttons
            new_config = {
                "text": f"Custom {new_button_idx}",
                "event_text": f"Custom {new_button_idx} Event",
                "txt_source_key": "None" # Default to no TXT source
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
        dialog_height = editor_window.winfo_reqheight() or 250

        center_x = main_x + (main_width // 2) - (dialog_width // 2)
        center_y = main_y + (main_height // 2) - (dialog_height // 2)
        editor_window.geometry(f"+{center_x}+{center_y}")

        frame = ttk.Frame(editor_window, padding="15")
        frame.pack(fill="both", expand=True)

        button_text_var = tk.StringVar(value=button_config.get("text", f"Custom {button_index+1}"))
        event_text_var = tk.StringVar(value=button_config.get("event_text", f"{button_config.get('text', f'Custom {button_index+1}')} Triggered"))
        txt_source_var = tk.StringVar(value=button_config.get("txt_source_key", "None"))
        current_color = self.button_colors.get(button_config.get("text"), (None, None))[1]
        button_color_var = tk.StringVar(value=current_color if current_color else "")
        
        row_idx = 0
        ttk.Label(frame, text="Button Text:").grid(row=row_idx, column=0, sticky="w", pady=2, padx=5)
        text_entry = ttk.Entry(frame, textvariable=button_text_var, width=30)
        text_entry.grid(row=row_idx, column=1, columnspan=2, sticky="ew", pady=2, padx=5)
        ToolTip(text_entry, "Text displayed on the button.")

        row_idx += 1
        ttk.Label(frame, text="Event Text:").grid(row=row_idx, column=0, sticky="w", pady=2, padx=5)
        event_entry = ttk.Entry(frame, textvariable=event_text_var, width=30)
        event_entry.grid(row=row_idx, column=1, columnspan=2, sticky="ew", pady=2, padx=5)
        ToolTip(event_entry, "Text written to the 'Event' column in the log.")

        row_idx += 1
        ttk.Label(frame, text="Event Source:").grid(row=row_idx, column=0, sticky="w", pady=2, padx=5)
        txt_source_options = ["None", "Main TXT", "TXT Source 2", "TXT Source 3"]
        source_combobox = ttk.Combobox(frame, textvariable=txt_source_var,
                                       values=txt_source_options, state="readonly", width=27)
        source_combobox.grid(row=row_idx, column=1, columnspan=2, sticky="ew", pady=2, padx=5)
        ToolTip(source_combobox, "Select which TXT file source this button should read data from. 'None' means no TXT data will be logged by this button.")

        row_idx += 1
        ttk.Label(frame, text="Button Color:").grid(row=row_idx, column=0, sticky="w", pady=2, padx=5)
        
        color_display_label = tk.Label(frame, width=4, relief="solid", borderwidth=1,
                                       background=button_color_var.get() if button_color_var.get() else 'SystemButtonFace')
        color_display_label.grid(row=row_idx, column=1, sticky="w", pady=2, padx=5)

        color_buttons_frame = ttk.Frame(frame)
        color_buttons_frame.grid(row=row_idx, column=2, sticky="e", pady=2, padx=5)

        clear_btn = ttk.Button(color_buttons_frame, text="X", width=2, style="Toolbutton",
                               command=lambda: self._set_color_on_widget(button_color_var, color_display_label, None, editor_window))
        clear_btn.pack(side="left", padx=1)
        ToolTip(clear_btn, "Clear button/row color (use default appearance).")

        pastel_colors_for_picker = ["#FFB3BA", "#FFDFBA", "#FFFFBA", "#BAFFC9", "#BAE1FF"]
        for p_color in pastel_colors_for_picker:
            try:
                b = tk.Button(color_buttons_frame, bg=p_color, width=1, height=1, relief="raised", bd=1,
                                  command=lambda c=p_color: self._set_color_on_widget(button_color_var, color_display_label, c, editor_window))
                b.pack(side=tk.LEFT, padx=1)
            except tk.TclError: pass

        choose_btn = ttk.Button(color_buttons_frame, text="...", width=3, style="Toolbutton",
                                  command=lambda v=button_color_var, l=color_display_label, n=button_text_var.get(): self._choose_color_dialog(v, l, editor_window, n))
        choose_btn.pack(side="left", padx=1)
        ToolTip(choose_btn, "Choose a custom color.")

        row_idx += 1
        button_controls_frame = ttk.Frame(frame)
        button_controls_frame.grid(row=row_idx, column=0, columnspan=3, pady=(15,0), sticky="e")

        def save_changes():
            old_button_text = button_config.get("text")
            
            button_config["text"] = button_text_var.get().strip() or f"Custom {button_index+1}"
            button_config["event_text"] = event_text_var.get().strip() or f"{button_config['text']} Triggered"
            button_config["txt_source_key"] = txt_source_var.get()

            new_color_hex = button_color_var.get() if button_color_var.get() else None
            
            if old_button_text in self.button_colors and old_button_text != button_config["text"]:
                del self.button_colors[old_button_text]
            
            self.button_colors[button_config["text"]] = (None, new_color_hex)

            self.save_settings()
            self.update_custom_buttons()
            editor_window.destroy()

        ttk.Button(button_controls_frame, text="Save", command=save_changes, style="Accent.TButton").pack(side="right", padx=5)
        ttk.Button(button_controls_frame, text="Cancel", command=editor_window.destroy).pack(side="right")
        
        editor_window.protocol("WM_DELETE_WINDOW", editor_window.destroy)
        editor_window.wait_window(editor_window)
        self.custom_inline_editor_window = None

    def _set_color_on_widget(self, color_str_var, display_label, color_hex, parent_toplevel):
        """Internal helper to validate and set the color for a color picker widget."""
        valid_color = None
        if color_hex:
            temp_label = None
            try:
                temp_label = tk.Label(parent_toplevel)
                temp_label.config(background=color_hex)
                valid_color = color_hex
            except tk.TclError:
                valid_color = None
            finally:
                if temp_label is not None:
                    try: temp_label.destroy()
                    except Exception: pass
            
        color_str_var.set(valid_color if valid_color else "")
        try:
            display_label.config(background=valid_color if valid_color else 'SystemButtonFace')
        except tk.TclError:
            display_label.config(background='SystemButtonFace')

    def _choose_color_dialog(self, color_str_var, display_label, parent_toplevel, name="Item"):
        """Opens color chooser dialog and updates the color_str_var and display_label."""
        current_color = color_str_var.get()
        color_code = colorchooser.askcolor(color=current_color if current_color else None,
                                           title=f"Choose Color for {name}",
                                           parent=parent_toplevel)
        if color_code and color_code[1]:
            self._set_color_on_widget(color_str_var, display_label, color_code[1], parent_toplevel)


# --- Settings Window Class (MODIFIED) ---
class SettingsWindow:
    def __init__(self, master, parent_gui):
        self.master = master
        self.parent_gui = parent_gui
        self.master.title("Settings")
        self.master.geometry("1000x650")
        self.master.minsize(700, 500)
        self.style = parent_gui.style

        self.main_frame = ttk.Frame(self.master)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.main_frame.rowconfigure(0, weight=1); self.main_frame.columnconfigure(0, weight=1)

        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        # Initialize selection tracking for TXT data columns
        self.selected_txt_row_index = -1 # -1 means no row is selected
        self.txt_move_up_btn = None
        self.txt_move_down_btn = None


        # Create tabs
        self.create_file_paths_tab() # Renamed and reorganized
        self.create_txt_column_mapping_tab()
        self.create_button_configuration_tab() # New tab
        self.create_monitored_folders_tab() # Renamed
        self.create_button_colors_tab() # New tab
        self.create_sqlite_tab()

        # Bottom Buttons
        button_frame = ttk.Frame(self.main_frame)
        button_frame.grid(row=1, column=0, pady=(10, 0), sticky="e")
        ttk.Button(button_frame, text="Save and Close", command=self.save_and_close, style="Accent.TButton").pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.master.destroy).pack(side=tk.RIGHT)

    def save_and_close(self):
        self.save_settings()
        self.master.destroy()

    # --- Tab Creation Methods ---

    def create_file_paths_tab(self):
        tab = ttk.Frame(self.notebook, padding=20)
        self.notebook.add(tab, text="File Paths")
        
        # Excel Log File
        log_frame = ttk.LabelFrame(tab, text="Excel Log File (.xlsx)", padding=15)
        log_frame.pack(fill="x", pady=(0, 15))
        log_frame.columnconfigure(1, weight=1)
        self.log_file_label = ttk.Label(log_frame, text="Path:", anchor='e')
        self.log_file_label.grid(row=0, column=0, padx=(0, 5), pady=5, sticky='w')
        self.log_file_entry = ttk.Entry(log_frame, width=80)
        self.log_file_entry.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        log_browse_btn = ttk.Button(log_frame, text="Browse...", command=self.select_excel_file)
        log_browse_btn.grid(row=0, column=2, padx=(5, 0), pady=5)
        ToolTip(log_browse_btn, "Select the main Excel file for logging."); ToolTip(self.log_file_entry, "Full path to the .xlsx file where all log entries will be written.")

        # Main Navigation TXT Data Folder
        txt_main_frame = ttk.LabelFrame(tab, text="Main Navigation TXT Data Folder (for general events)", padding=15)
        txt_main_frame.pack(fill="x", pady=(0, 15))
        txt_main_frame.columnconfigure(1, weight=1)
        self.txt_folder_label_main = ttk.Label(txt_main_frame, text="Folder:", anchor='e')
        self.txt_folder_label_main.grid(row=0, column=0, padx=(0, 5), pady=5, sticky='w')
        self.txt_folder_entry_main = ttk.Entry(txt_main_frame, width=80)
        self.txt_folder_entry_main.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        txt_browse_btn_main = ttk.Button(txt_main_frame, text="Browse...", command=lambda: self.select_txt_folder(self.txt_folder_entry_main))
        txt_browse_btn_main.grid(row=0, column=2, padx=(5, 0), pady=5)
        ToolTip(txt_browse_btn_main, "Select the primary folder containing navigation TXT files. Used by 'Event' button and can be selected by custom buttons."); ToolTip(self.txt_folder_entry_main, "Path to the main folder containing navigation TXT files.")

        # Navigation TXT Data Folder - Source 2
        txt_set2_frame = ttk.LabelFrame(tab, text="Additional TXT Data Folder - Source 2 (for custom buttons)", padding=15)
        txt_set2_frame.pack(fill="x", pady=(0, 15))
        txt_set2_frame.columnconfigure(1, weight=1)
        self.txt_folder_label_set2 = ttk.Label(txt_set2_frame, text="Folder:", anchor='e')
        self.txt_folder_label_set2.grid(row=0, column=0, padx=(0, 5), pady=5, sticky='w')
        self.txt_folder_entry_set2 = ttk.Entry(txt_set2_frame, width=80)
        self.txt_folder_entry_set2.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        txt_browse_btn_set2 = ttk.Button(txt_set2_frame, text="Browse...", command=lambda: self.select_txt_folder(self.txt_folder_entry_set2))
        txt_browse_btn_set2.grid(row=0, column=2, padx=(5, 0), pady=5)
        ToolTip(txt_browse_btn_set2, "Select a secondary folder for navigation TXT files. Can be assigned to custom buttons."); ToolTip(self.txt_folder_entry_set2, "Path to a secondary folder for navigation TXT files.")

        # Navigation TXT Data Folder - Source 3
        txt_set3_frame = ttk.LabelFrame(tab, text="Additional TXT Data Folder - Source 3 (for custom buttons)", padding=15)
        txt_set3_frame.pack(fill="x", pady=(0, 0))
        txt_set3_frame.columnconfigure(1, weight=1)
        self.txt_folder_label_set3 = ttk.Label(txt_set3_frame, text="Folder:", anchor='e')
        self.txt_folder_label_set3.grid(row=0, column=0, padx=(0, 5), pady=5, sticky='w')
        self.txt_folder_entry_set3 = ttk.Entry(txt_set3_frame, width=80)
        self.txt_folder_entry_set3.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        txt_browse_btn_set3 = ttk.Button(txt_set3_frame, text="Browse...", command=lambda: self.select_txt_folder(self.txt_folder_entry_set3))
        txt_browse_btn_set3.grid(row=0, column=2, padx=(5, 0), pady=5)
        ToolTip(txt_browse_btn_set3, "Select a third folder for navigation TXT files. Can be assigned to custom buttons."); ToolTip(self.txt_folder_entry_set3, "Path to a third folder for navigation TXT files.")


    def select_excel_file(self):
        initial_dir = os.path.dirname(self.log_file_entry.get()) if self.log_file_entry.get() else os.getcwd()
        file_path = filedialog.askopenfilename(initialdir=initial_dir, filetypes=[("Excel files", "*.xlsx")], parent=self.master, title="Select Excel Log File")
        if file_path: self.log_file_entry.delete(0, tk.END); self.log_file_entry.insert(0, file_path)

    def select_txt_folder(self, entry_widget):
        current_path = entry_widget.get()
        initial_dir = current_path if os.path.isdir(current_path) else os.path.dirname(current_path) if current_path else os.getcwd()
        folder_path = filedialog.askdirectory(initialdir=initial_dir, parent=self.master, title="Select Navigation TXT Folder")
        if folder_path: entry_widget.delete(0, tk.END); entry_widget.insert(0, folder_path)

    def create_txt_column_mapping_tab(self):
        tab = ttk.Frame(self.notebook, padding=20)
        self.notebook.add(tab, text="TXT Data Columns")
        
        ttk.Label(tab, text="Map fields found in TXT files to your desired Excel/Database column names. Check 'Skip' to ignore a field entirely. Click on a row to select it, then use the Move Up/Down buttons to reorder.", wraplength=900, justify=tk.LEFT).pack(pady=(0, 10), anchor='w')

        # Control buttons for adding/removing/reordering fields
        controls_frame = ttk.Frame(tab)
        controls_frame.pack(fill='x', pady=(0, 10))
        
        # --- NEW: Button to Preview TXT data ---
        preview_btn = ttk.Button(controls_frame, text="Preview Latest TXT Data", command=self.preview_txt_data)
        preview_btn.pack(side=tk.LEFT, padx=(0, 10))
        ToolTip(preview_btn, "Load the last line from the latest file in the 'Main Navigation TXT Data Folder' to see the data for each column index.")
        
        clear_preview_btn = ttk.Button(controls_frame, text="Clear Preview", command=self.clear_txt_preview)
        clear_preview_btn.pack(side=tk.LEFT, padx=(0, 20))
        ToolTip(clear_preview_btn, "Clear the preview data from the view.")

        # Spacer to push other buttons to the right
        spacer = ttk.Frame(controls_frame)
        spacer.pack(side=tk.LEFT, expand=True, fill='x')

        self.txt_move_up_btn = ttk.Button(controls_frame, text="Move Up", command=lambda: self.move_selected_txt_field("up"))
        self.txt_move_up_btn.pack(side=tk.RIGHT, padx=5)
        ToolTip(self.txt_move_up_btn, "Move the selected field up in the list.")

        self.txt_move_down_btn = ttk.Button(controls_frame, text="Move Down", command=lambda: self.move_selected_txt_field("down"))
        self.txt_move_down_btn.pack(side=tk.RIGHT, padx=5)
        ToolTip(self.txt_move_down_btn, "Move the selected field down in the list.")

        ttk.Button(controls_frame, text="Add New Field", command=self.add_txt_field_row).pack(side=tk.RIGHT, padx=5)
        
        # Canvas and Scrollbar for the dynamic field list
        self.txt_fields_canvas = tk.Canvas(tab, borderwidth=0, background="#ffffff")
        txt_scrollbar = ttk.Scrollbar(tab, orient="vertical", command=self.txt_fields_canvas.yview)
        self.txt_fields_scrollable_frame = ttk.Frame(self.txt_fields_canvas, style="Row0.TFrame")
        self.txt_fields_scrollable_frame.bind("<Configure>", lambda e: self.txt_fields_canvas.configure(scrollregion=self.txt_fields_canvas.bbox("all")))
        self.txt_fields_canvas_window = self.txt_fields_canvas.create_window((0, 0), window=self.txt_fields_scrollable_frame, anchor="nw")
        self.txt_fields_canvas.configure(yscrollcommand=txt_scrollbar.set)
        self.txt_fields_canvas.pack(side="left", fill="both", expand=True, padx=(0,0), pady=0)
        txt_scrollbar.pack(side="right", fill="y", padx=(0,0), pady=0)
        
        def _on_mousewheel_txt(event):
            if event.num == 4: delta = -1
            elif event.num == 5: delta = 1
            elif hasattr(event, 'delta'): delta = -int(event.delta / 120)
            else: delta = 0
            self.txt_fields_canvas.yview_scroll(delta, "units")
        self.txt_fields_canvas.bind_all("<MouseWheel>", _on_mousewheel_txt); self.txt_fields_canvas.bind_all("<Button-4>", _on_mousewheel_txt); self.txt_fields_canvas.bind_all("<Button-5>", _on_mousewheel_txt)
        
        # Store widgets for each row dynamically
        self.txt_field_row_widgets = [] # List of dictionaries, each holding refs for a row
        self.add_txt_field_header(self.txt_fields_scrollable_frame) # Initial header

        # Populate with existing fields
        self.recreate_txt_field_rows() # Will be called by load_settings too
        self._update_txt_move_buttons_state() # Initial state of move buttons

    # --- NEW: Method to preview TXT data ---
    def preview_txt_data(self):
        """Finds the latest TXT file, reads the last line, and displays the parts in the preview column."""
        txt_folder = self.txt_folder_entry_main.get()
        if not txt_folder or not os.path.isdir(txt_folder):
            messagebox.showerror("Path Error", "The 'Main Navigation TXT Data Folder' path is not set or is invalid. Please set it in the 'File Paths' tab.", parent=self.master)
            return

        latest_file = self.parent_gui.find_latest_file_in_folder(txt_folder, ".txt")

        if not latest_file:
            messagebox.showinfo("File Not Found", f"No .txt files were found in:\n{txt_folder}", parent=self.master)
            return

        try:
            lines = []
            encodings_to_try = ['utf-8', 'latin-1', 'cp1252']
            read_success = False
            for enc in encodings_to_try:
                try:
                    with open(latest_file, "r", encoding=enc) as f:
                        lines = f.readlines()
                    read_success = True
                    break
                except UnicodeDecodeError:
                    continue
            
            if not read_success or not lines:
                messagebox.showinfo("File Empty", f"The latest file is empty or could not be read:\n{os.path.basename(latest_file)}", parent=self.master)
                return

            last_line = lines[-1].strip()
            data_parts = last_line.split(',')
            
            # Update the preview labels
            for i, row_widgets in enumerate(self.txt_field_row_widgets):
                preview_label = row_widgets.get("preview_label")
                if preview_label:
                    if i < len(data_parts):
                        preview_label.config(text=data_parts[i].strip())
                    else:
                        preview_label.config(text="<no data>")
            
            self.parent_gui.update_status(f"Preview loaded from {os.path.basename(latest_file)}")

        except Exception as e:
            messagebox.showerror("Read Error", f"An error occurred while reading the file:\n{e}", parent=self.master)

    # --- NEW: Method to clear the preview data ---
    def clear_txt_preview(self):
        """Clears the text from all preview labels."""
        for row_widgets in self.txt_field_row_widgets:
            preview_label = row_widgets.get("preview_label")
            if preview_label:
                preview_label.config(text="")
        self.parent_gui.update_status("Preview cleared.")

    def add_txt_field_header(self, parent):
        header_frame = ttk.Frame(parent, style="Header.TFrame", padding=(5,3))
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 5))
        
        # --- MODIFIED: Added Preview Data column ---
        header_frame.grid_columnconfigure(0, weight=1) # TXT Field Name
        header_frame.grid_columnconfigure(1, weight=1) # Target Column
        header_frame.grid_columnconfigure(2, weight=1) # Preview Data (NEW)
        header_frame.grid_columnconfigure(3, weight=0) # Skip
        header_frame.grid_columnconfigure(4, weight=0) # Actions

        ttk.Label(header_frame, text="TXT Field (Name / Index)", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=5, sticky='w')
        ttk.Label(header_frame, text="Target Excel/DB Column Name", font=("Arial", 10, "bold")).grid(row=0, column=1, padx=5, sticky='w')
        ttk.Label(header_frame, text="Preview Data", font=("Arial", 10, "bold")).grid(row=0, column=2, padx=5, sticky='w') # NEW
        ttk.Label(header_frame, text="Skip?", font=("Arial", 10, "bold")).grid(row=0, column=3, padx=5, sticky='w')
        ttk.Label(header_frame, text="Actions", font=("Arial", 10, "bold")).grid(row=0, column=4, padx=5, sticky='w')

    def _select_txt_row(self, index):
        """Highlights the selected row and updates the selected index."""
        if self.selected_txt_row_index != -1 and self.selected_txt_row_index < len(self.txt_field_row_widgets):
            # Deselect previous row
            prev_row_info = self.txt_field_row_widgets[self.selected_txt_row_index]
            prev_row_frame = prev_row_info["row_frame"]
            original_style = f"Row{self.selected_txt_row_index % 2}.TFrame"
            prev_row_frame.config(style=original_style)
            # Re-apply styles to child widgets if they were overridden
            for child in prev_row_frame.winfo_children():
                if isinstance(child, ttk.Label):
                    child.config(style=original_style.replace("Frame", "Label"))
                elif isinstance(child, ttk.Checkbutton):
                    child.config(style=original_style.replace("Frame", "Checkbutton"))


        self.selected_txt_row_index = index
        if index != -1 and index < len(self.txt_field_row_widgets):
            # Select current row
            current_row_info = self.txt_field_row_widgets[index]
            current_row_frame = current_row_info["row_frame"]
            current_row_frame.config(style="Selected.TFrame")
            for child in current_row_frame.winfo_children():
                if isinstance(child, ttk.Label):
                    child.config(style="Selected.TLabel") # Apply a specific style for selected labels
                elif isinstance(child, ttk.Checkbutton):
                    child.config(style="Selected.TCheckbutton")

        self._update_txt_move_buttons_state()

    def _update_txt_move_buttons_state(self):
        """Enables/disables move buttons based on selected row index."""
        can_move_up = (self.selected_txt_row_index > 0)
        can_move_down = (self.selected_txt_row_index != -1 and self.selected_txt_row_index < len(self.parent_gui.txt_field_columns_config) - 1)

        if self.txt_move_up_btn:
            self.txt_move_up_btn.config(state=tk.NORMAL if can_move_up else tk.DISABLED)
        if self.txt_move_down_btn:
            self.txt_move_down_btn.config(state=tk.NORMAL if can_move_down else tk.DISABLED)

    def move_selected_txt_field(self, direction):
        """Moves the selected TXT field up or down."""
        current_index = self.selected_txt_row_index
        if current_index == -1:
            messagebox.showinfo("No Selection", "Please select a row to move.", parent=self.master)
            return

        if direction == "up":
            if current_index > 0:
                self.parent_gui.txt_field_columns_config[current_index], self.parent_gui.txt_field_columns_config[current_index - 1] = \
                    self.parent_gui.txt_field_columns_config[current_index - 1], self.parent_gui.txt_field_columns_config[current_index]
                self.selected_txt_row_index -= 1 # Update selected index
                self.recreate_txt_field_rows(reselect_index=self.selected_txt_row_index)
                self.parent_gui.update_status(f"Moved field '{self.parent_gui.txt_field_columns_config[self.selected_txt_row_index]['field']}' up.")
        elif direction == "down":
            if current_index < len(self.parent_gui.txt_field_columns_config) - 1:
                self.parent_gui.txt_field_columns_config[current_index], self.parent_gui.txt_field_columns_config[current_index + 1] = \
                    self.parent_gui.txt_field_columns_config[current_index + 1], self.parent_gui.txt_field_columns_config[current_index]
                self.selected_txt_row_index += 1 # Update selected index
                self.recreate_txt_field_rows(reselect_index=self.selected_txt_row_index)
                self.parent_gui.update_status(f"Moved field '{self.parent_gui.txt_field_columns_config[self.selected_txt_row_index]['field']}' down.")


    def add_txt_field_row(self):
        """Adds a new row for a custom TXT field."""
        new_field_index = len(self.parent_gui.txt_field_columns_config) + 1
        new_field_name = f"Custom_Field_{new_field_index}"
        new_column_name = f"Custom_Col_{new_field_index}" # Propose a default column name too

        self.parent_gui.txt_field_columns_config.append({
            "field": new_field_name,
            "column_name": new_column_name,
            "skip": False
        })
        newly_added_index = len(self.parent_gui.txt_field_columns_config) - 1
        self.recreate_txt_field_rows(reselect_index=newly_added_index) # Redraw all rows and select new one
        self.parent_gui.update_status(f"Added new TXT field '{new_field_name}'.")

    def remove_txt_field_row(self, index_to_remove):
        if messagebox.askyesno("Confirm Deletion", f"Are you sure you want to remove field '{self.parent_gui.txt_field_columns_config[index_to_remove]['field']}'?", parent=self.master):
            del self.parent_gui.txt_field_columns_config[index_to_remove]
            # Adjust selected index if the removed row was before it or was the selected one
            if self.selected_txt_row_index == index_to_remove:
                self.selected_txt_row_index = -1 # No longer selected
            elif self.selected_txt_row_index > index_to_remove:
                self.selected_txt_row_index -= 1

            self.recreate_txt_field_rows(reselect_index=self.selected_txt_row_index) # Redraw all rows
            self.parent_gui.update_status("TXT field removed.")


    def recreate_txt_field_rows(self, reselect_index=None):
        # Clear existing widgets except the header
        for widget_info in self.txt_field_row_widgets:
            if "row_frame" in widget_info and widget_info["row_frame"].winfo_exists():
                widget_info["row_frame"].destroy()
        self.txt_field_row_widgets.clear()

        # Define the set of default fields that should not be editable as 'TXT Field' or removable
        default_fixed_fields = {"Date", "Time", "KP", "DCC", "Line name", "Latitude", "Longitude", "Easting", "Northing", "Event"}

        # Recreate rows based on the current self.parent_gui.txt_field_columns_config
        for i, config in enumerate(self.parent_gui.txt_field_columns_config):
            row_index = i + 1 # +1 because of the header row
            
            # Determine initial style based on row index and reselect_index
            if reselect_index is not None and i == reselect_index:
                row_style = "Selected.TFrame"
                label_style = "Selected.TLabel"
                checkbox_style = "Selected.TCheckbutton"
                self.selected_txt_row_index = i # Ensure internal state is updated for reselection
            else:
                row_style = f"Row{i % 2}.TFrame"
                label_style = row_style.replace("Frame", "Label")
                checkbox_style = row_style.replace("Frame", "Checkbutton")
                try: self.style.configure(label_style, background=self.style.lookup(row_style, 'background'))
                except Exception: pass
                try: self.style.configure(checkbox_style, background=self.style.lookup(row_style, 'background'))
                except Exception: pass

            row_frame = ttk.Frame(self.txt_fields_scrollable_frame, style=row_style, padding=(0, 2))
            row_frame.grid(row=row_index, column=0, sticky="ew", pady=0)
            
            # --- MODIFIED: Added Preview Data column ---
            row_frame.grid_columnconfigure(0, weight=1) # TXT Field Name
            row_frame.grid_columnconfigure(1, weight=1) # Target Column
            row_frame.grid_columnconfigure(2, weight=1) # Preview Data (NEW)
            row_frame.grid_columnconfigure(3, weight=0) # Skip
            row_frame.grid_columnconfigure(4, weight=0) # Actions

            # Bind click event to all elements in the row for selection
            click_handler = lambda e, idx=i: self._select_txt_row(idx)
            row_frame.bind("<Button-1>", click_handler)

            # TXT Field Label/Entry (fixed or editable)
            if config["field"] in default_fixed_fields:
                field_label = ttk.Label(row_frame, text=f"{config['field']}:", anchor='w', style=label_style)
                field_label.grid(row=0, column=0, padx=5, sticky='ew')
                field_label.bind("<Button-1>", click_handler)
                current_field_entry_widget = None
            else: # Allow custom fields to be edited
                field_entry = ttk.Entry(row_frame)
                field_entry.insert(0, config["field"])
                field_entry.grid(row=0, column=0, padx=5, sticky='ew')
                ToolTip(field_entry, "Enter the exact name of the field as it appears in the TXT file. E.g., 'Depth'.")
                current_field_entry_widget = field_entry

            # Target Excel/DB Column Name
            column_entry = ttk.Entry(row_frame)
            column_entry.insert(0, config.get("column_name", config["field"])) # Default to field name if not set
            column_entry.grid(row=0, column=1, padx=5, sticky="ew")
            ToolTip(column_entry, f"Enter the exact column name in your Excel/DB where '{config['field']}' data should be written.")
            
            # --- NEW: Preview Data Label ---
            preview_label = ttk.Label(row_frame, text="", style=label_style, anchor='w', foreground="blue")
            preview_label.grid(row=0, column=2, padx=5, sticky='ew')
            preview_label.bind("<Button-1>", click_handler)
            
            # Skip Checkbox
            skip_var = tk.BooleanVar(value=config.get("skip", False))
            skip_checkbox = ttk.Checkbutton(row_frame, variable=skip_var, style=checkbox_style)
            skip_checkbox.grid(row=0, column=3, padx=5, sticky='w')
            ToolTip(skip_checkbox, f"Check this box to ignore the '{config['field']}' field entirely when logging TXT data.")

            # Remove Button
            remove_button_frame = ttk.Frame(row_frame, style=row_style)
            remove_button_frame.grid(row=0, column=4, padx=5, sticky='w')
            remove_btn = ttk.Button(remove_button_frame, text="Remove", width=8, style="Toolbutton",
                                      command=lambda idx=i: self.remove_txt_field_row(idx))
            if config["field"] in default_fixed_fields:
                remove_btn.config(state=tk.DISABLED) # Disable removing default fields
            remove_btn.pack(side=tk.LEFT, padx=1)
            ToolTip(remove_btn, "Remove this custom field.")

            # Store references to widgets for later retrieval
            self.txt_field_row_widgets.append({
                "field_entry_widget": current_field_entry_widget,
                "column_entry": column_entry,
                "skip_var": skip_var,
                "preview_label": preview_label, # NEW
                "row_frame": row_frame
            })
        
        # After recreating all rows, ensure the selection state is correct
        if reselect_index is None or not (0 <= reselect_index < len(self.parent_gui.txt_field_columns_config)):
            self.selected_txt_row_index = -1 
        if self.selected_txt_row_index != -1:
            pass 

        self._update_txt_move_buttons_state() # Update button states
        self.master.after_idle(lambda: self.txt_fields_canvas.config(scrollregion=self.txt_fields_canvas.bbox("all")))


    def create_button_configuration_tab(self):
        tab = ttk.Frame(self.notebook, padding=20)
        self.notebook.add(tab, text="Button Configuration")

        num_buttons_frame = ttk.Frame(tab); num_buttons_frame.pack(pady=5, anchor='w')
        ttk.Label(num_buttons_frame, text=f"Number of Custom Buttons (0-{self.parent_gui.MAX_CUSTOM_BUTTONS}):").pack(side='left', padx=5)
        self.num_buttons_entry = ttk.Entry(num_buttons_frame, width=5); self.num_buttons_entry.pack(side='left', padx=5); ToolTip(self.num_buttons_entry, "Enter the number of custom event buttons needed (max 10).")
        update_btn = ttk.Button(num_buttons_frame, text="Update List", command=self.update_num_custom_buttons); update_btn.pack(side='left', padx=5); ToolTip(update_btn, "Update the list below to show the specified number of button configurations.")

        # Header for the custom button configuration table
        header_frame = ttk.Frame(tab, style="Header.TFrame", padding=(5,3))
        header_frame.pack(fill='x', pady=(15,5))
        
        # Configure columns for the header frame to match the rows
        header_frame.grid_columnconfigure(0, weight=0) # Button #
        header_frame.grid_columnconfigure(1, weight=1) # Button Text
        header_frame.grid_columnconfigure(2, weight=2) # Event Text (longer)
        header_frame.grid_columnconfigure(3, weight=0) # Event Source

        ttk.Label(header_frame, text="Button #", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=(5,0), sticky='w')
        ttk.Label(header_frame, text="Button Text", font=("Arial", 10, "bold")).grid(row=0, column=1, padx=5, sticky='ew')
        ttk.Label(header_frame, text="Event Text (for Log)", font=("Arial", 10, "bold")).grid(row=0, column=2, padx=5, sticky='ew')
        ttk.Label(header_frame, text="Event Source", font=("Arial", 10, "bold")).grid(row=0, column=3, padx=5, sticky='w')

        self.custom_button_entries_frame = ttk.Frame(tab); self.custom_button_entries_frame.pack(pady=0, fill='both', expand=True)
        self.custom_button_widgets = []

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
                        self.parent_gui.custom_button_configs.append({"text": f"Custom {idx}", "event_text": f"Custom {idx} Event", "txt_source_key": "None"})
                self.recreate_custom_button_settings()
        except ValueError as e:
            messagebox.showerror("Invalid Number", f"Please enter a whole number between 0 and {self.parent_gui.MAX_CUSTOM_BUTTONS}. Error: {e}", parent=self.master)
            self.num_buttons_entry.delete(0, tk.END); self.num_buttons_entry.insert(0, str(self.parent_gui.num_custom_buttons))

    def recreate_custom_button_settings(self):
        """Clears and redraws the custom button configuration rows (text, event, source)."""
        for widget in self.custom_button_entries_frame.winfo_children():
            widget.destroy()
        self.custom_button_widgets = []

        num_buttons = self.parent_gui.num_custom_buttons
        configs = self.parent_gui.custom_button_configs
        txt_source_options = ["None", "Main TXT", "TXT Source 2", "TXT Source 3"]

        for i in range(num_buttons):
            config = configs[i] if i < len(configs) else {}
            initial_text = config.get("text", f"Custom {i+1}")
            initial_event = config.get("event_text", f"{initial_text} Event")
            initial_txt_source = config.get("txt_source_key", "None")

            style_name = f"Row{i % 2}.TFrame"
            row_frame = ttk.Frame(self.custom_button_entries_frame, style=style_name, padding=(0, 2))
            row_frame.pack(fill='x', pady=0)

            # Configure columns for each row frame
            row_frame.grid_columnconfigure(0, weight=0) # Button # Label (fixed width)
            row_frame.grid_columnconfigure(1, weight=1) # Button Text Entry
            row_frame.grid_columnconfigure(2, weight=2) # Event Text Entry (expands more)
            row_frame.grid_columnconfigure(3, weight=0) # Event Source Combobox (fixed width)


            ttk.Label(row_frame, text=f"{i+1}", width=7, style=style_name.replace("Frame","Label")).grid(row=0, column=0, padx=(5,0), sticky='w')
            text_entry = ttk.Entry(row_frame, width=20); text_entry.insert(0, initial_text); text_entry.grid(row=0, column=1, padx=5, sticky='ew'); ToolTip(text_entry, "Text displayed on the button.")
            event_entry = ttk.Entry(row_frame, width=30); event_entry.insert(0, initial_event); event_entry.grid(row=0, column=2, padx=5, sticky='ew'); ToolTip(event_entry, "Text written to the 'Event' column in the log.")

            txt_source_var = tk.StringVar(value=initial_txt_source)
            txt_source_combobox = ttk.Combobox(row_frame, textvariable=txt_source_var,
                                                 values=txt_source_options, state="readonly", width=12)
            txt_source_combobox.grid(row=0, column=3, padx=5, sticky='w')
            ToolTip(txt_source_combobox, "Select which TXT file source this button should read data from. 'None' means no TXT data will be logged by this button.")

            self.custom_button_widgets.append( (text_entry, event_entry, txt_source_var) )

    def create_monitored_folders_tab(self): # Renamed
        tab = ttk.Frame(self.notebook); self.notebook.add(tab, text="Monitored Folders")
        
        ttk.Label(tab, text="Configure additional folders to monitor for their latest file names. The latest file name will be logged in the specified Excel/DB column.", wraplength=900, justify=tk.LEFT).pack(pady=(0, 10), anchor='w')

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
        self.folder_canvas.bind_all("<MouseWheel>", _on_mousewheel); self.folder_canvas.bind_all("<Button-4>", _on_mousewheel); self.folder_canvas.bind_all("<Button-5>", _on_mousewheel)
        self.folder_entries = {}; self.folder_column_entries = {}; self.file_extension_entries = {}; self.folder_skip_vars = {}; self.folder_row_widgets = {}
        self.add_folder_header(self.scrollable_frame)

    def add_folder_header(self, parent):
        header_frame = ttk.Frame(parent, style="Header.TFrame", padding=(5,3))
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 5))
        
        # Define column weights and sticky for the header to match row layout
        # Col 0: Folder Type - fixed width
        header_frame.grid_columnconfigure(0, weight=0)
        # Col 1: Monitor Path - should expand
        header_frame.grid_columnconfigure(1, weight=1)
        # Col 2: "..." button - fixed width
        header_frame.grid_columnconfigure(2, weight=0)
        # Col 3: Target Column - fixed width
        header_frame.grid_columnconfigure(3, weight=0)
        # Col 4: File Ext. - fixed width
        header_frame.grid_columnconfigure(4, weight=0)
        # Col 5: Skip? - fixed width
        header_frame.grid_columnconfigure(5, weight=0)

        ttk.Label(header_frame, text="Folder Type", width=15, anchor="w", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=(5,0), sticky='w')
        ttk.Label(header_frame, text="Monitor Path", anchor="w", font=("Arial", 10, "bold")).grid(row=0, column=1, padx=5, sticky='ew') # Changed to 'ew'
        ttk.Label(header_frame, text="...", width=4, anchor="center", font=("Arial", 10, "bold")).grid(row=0, column=2, padx=1, sticky='w') 
        ttk.Label(header_frame, text="Target Column", width=20, anchor="w", font=("Arial", 10, "bold")).grid(row=0, column=3, padx=5, sticky='w')
        ttk.Label(header_frame, text="File Ext.", width=10, anchor="w", font=("Arial", 10, "bold")).grid(row=0, column=4, padx=5, sticky='w')
        ttk.Label(header_frame, text="Skip?", width=5, anchor="center", font=("Arial", 10, "bold")).grid(row=0, column=5, padx=(10,5), sticky='w')

    def add_initial_folder_rows(self):
        default_folders = ["Qinsy DB", "Naviscan", "SIS", "SSS", "SBP", "Mag", "Grad", "SVP", "SpintINS", "Video", "Cathx", "Hypack RAW", "Eiva NaviPac"]
        
        ordered_specific_txt_folders = [
            ("Main TXT File", self.parent_gui.txt_folder_path),
            ("TXT Source 2", self.parent_gui.txt_folder_path_set2),
            ("TXT Source 3", self.parent_gui.txt_folder_path_set3)
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

            column_name_to_use = self.parent_gui.folder_columns.get(folder_name, folder_name)
            extension_to_use = self.parent_gui.file_extensions.get(folder_name, "")

            if folder_name in ["Main TXT File", "TXT Source 2", "TXT Source 3"]:
                if not column_name_to_use or column_name_to_use == folder_name:
                    column_name_to_use = folder_name.replace(" ", "_")
                if not extension_to_use:
                    extension_to_use = "txt"

            self.add_folder_row(folder_name=folder_name, folder_path=folder_path_to_use,
                                  column_name=column_name_to_use,
                                  extension=extension_to_use,
                                  skip=self.parent_gui.folder_skips.get(folder_name, False))
        self.master.after_idle(self.update_scroll_region)

    def add_folder_row(self, folder_name="", folder_path="", column_name="", extension="", skip=False):
        row_index = len(self.folder_row_widgets) + 1; style_name = f"Row{row_index % 2}.TFrame"
        try: self.style.configure(style_name)
        except tk.TclError: bg = "#ffffff" if row_index % 2 == 0 else "#f5f5f5"; self.style.configure(style_name, background=bg)
        row_frame = ttk.Frame(self.scrollable_frame, style=style_name, padding=(0, 2)); row_frame.grid(row=row_index, column=0, sticky="ew", pady=0); 
        
        # Add columnconfigure to each row frame to match the header
        row_frame.grid_columnconfigure(0, weight=0) # Folder Type (fixed width)
        row_frame.grid_columnconfigure(1, weight=1) # Monitor Path (expands)
        row_frame.grid_columnconfigure(2, weight=0) # "..." button (fixed width)
        row_frame.grid_columnconfigure(3, weight=0) # Target Column (fixed width)
        row_frame.grid_columnconfigure(4, weight=0) # File Ext. (fixed width)
        row_frame.grid_columnconfigure(5, weight=0) # Skip? (fixed width)


        label_style = style_name.replace("Frame","Label")
        try: self.style.configure(label_style, background=self.style.lookup(style_name, 'background'))
        except Exception: pass
        label = ttk.Label(row_frame, text=f"{folder_name}:", width=15, anchor='w', style=label_style); label.grid(row=0, column=0, padx=(5,0), pady=1, sticky="w")
        entry = ttk.Entry(row_frame, width=50); entry.insert(0, folder_path); entry.grid(row=0, column=1, padx=5, pady=1, sticky="ew"); ToolTip(entry, f"Enter the full path to the '{folder_name}' data folder.")
        def select_folder(e=entry, name=folder_name):
            current_path = e.get(); initial = current_path if os.path.isdir(current_path) else (os.path.dirname(current_path) if current_path else os.getcwd())
            folder = filedialog.askdirectory(parent=self.master, initialdir=initial, title=f"Select Folder for {name}")
            if folder:
                e.delete(0, tk.END); e.insert(0, folder)
                if name == "Main TXT File": self.parent_gui.txt_folder_path = folder
                elif name == "TXT Source 2": self.parent_gui.txt_folder_path_set2 = folder
                elif name == "TXT Source 3": self.parent_gui.txt_folder_path_set3 = folder

        button = ttk.Button(row_frame, text="...", width=3, command=select_folder); button.grid(row=0, column=2, padx=(0,5), pady=1, sticky='w'); ToolTip(button, "Browse for the folder.") 
        column_entry = ttk.Entry(row_frame, width=20); column_entry.insert(0, column_name if column_name else folder_name); column_entry.grid(row=0, column=3, padx=5, pady=1, sticky="w"); ToolTip(column_entry, f"Enter the Excel/DB column name for the latest '{folder_name}' filename.")
        extension_entry = ttk.Entry(row_frame, width=10); extension_entry.insert(0, extension); extension_entry.grid(row=0, column=4, padx=5, pady=1, sticky="w"); ToolTip(extension_entry, f"Optional: Monitor only files ending with this extension (e.g., 'svp', 'log'). Leave blank for any file.")
        skip_var = tk.BooleanVar(value=skip); skip_checkbox = ttk.Checkbutton(row_frame, variable=skip_var); skip_checkbox.grid(row=0, column=5, padx=(15,5), pady=1, sticky="w"); ToolTip(skip_checkbox, f"Check to disable monitoring for the '{folder_name}' folder.")
        self.folder_entries[folder_name] = entry; self.folder_column_entries[folder_name] = column_entry; self.file_extension_entries[folder_name] = extension_entry; self.folder_skip_vars[folder_name] = skip_var; self.folder_row_widgets[folder_name] = row_frame

    def update_scroll_region(self):
        self.scrollable_frame.update_idletasks()
        self.folder_canvas.configure(scrollregion=self.folder_canvas.bbox("all"))

    def create_button_colors_tab(self): # New tab for all button colors
        tab = ttk.Frame(self.notebook, padding=20)
        self.notebook.add(tab, text="Button Colors")
        self.pastel_colors = ["#FFB3BA", "#FFDFBA", "#FFFFBA", "#BAFFC9", "#BAE1FF", "#E0BBE4", "#FFC8A2", "#D4A5A5", "#A2D4AB", "#A2C4D4"]

        canvas = tk.Canvas(tab, borderwidth=0, background=self.style.lookup("TFrame", "background"))
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        self.color_scrollable_frame = ttk.Frame(canvas)
        self.color_scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.color_scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True, padx=(0,0), pady=0)
        scrollbar.pack(side="right", fill="y", padx=(0,0), pady=0)

        # Headers for color tab
        header_frame = ttk.Frame(self.color_scrollable_frame, style="Header.TFrame", padding=(5,3))
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 5))
        # Ensure consistent column configurations with the rows below
        header_frame.grid_columnconfigure(0, weight=0) # Button Name (fixed width)
        header_frame.grid_columnconfigure(1, weight=1) # Button/Row Color (expands)

        ttk.Label(header_frame, text="Button Name", font=("Arial", 10, "bold"), width=20, anchor='w').grid(row=0, column=0, padx=5, sticky='w')
        ttk.Label(header_frame, text="Button/Row Color", font=("Arial", 10, "bold"), width=30, anchor='w').grid(row=0, column=1, padx=5, sticky='w')

        self.all_button_color_widgets = {} # Stores (StringVar, Label) for all buttons

        # Add standard button color rows
        standard_buttons_to_color = ["Log on", "Log off", "Event", "SVP", "New Day"]
        for i, btn_name in enumerate(standard_buttons_to_color):
            self._add_color_row(self.color_scrollable_frame, i + 1, btn_name, is_custom=False)
        
        # Add custom button color rows (up to max possible, even if not all are currently configured)
        for i in range(self.parent_gui.MAX_CUSTOM_BUTTONS): # Max 10 custom buttons
            btn_name = f"Custom {i+1}"
            self._add_color_row(self.color_scrollable_frame, len(standard_buttons_to_color) + i + 1, btn_name, is_custom=True)

        self.master.after_idle(lambda: canvas.config(scrollregion=canvas.bbox("all")))

    def _add_color_row(self, parent_frame, row_index, btn_name, is_custom=False):
        """Helper to create a single row for color setting."""
        style_name = f"Row{row_index % 2}.TFrame"
        row_f = ttk.Frame(parent_frame, style=style_name, padding=(0, 2))
        row_f.grid(row=row_index, column=0, sticky='ew', pady=0)
        
        # Configure columns for this row frame to match the header
        row_f.grid_columnconfigure(0, weight=0) # Button Name (fixed width)
        row_f.grid_columnconfigure(1, weight=1) # Color widgets frame (expands)


        label_style = style_name.replace("Frame", "Label")
        try: self.style.configure(label_style, background=self.style.lookup(style_name, 'background'))
        except Exception: pass

        ttk.Label(row_f, text=f"{btn_name}:", width=20, style=label_style, anchor='w').grid(row=0, column=0, padx=(5,0), sticky='w')

        color_widget_frame = ttk.Frame(row_f, style=style_name)
        color_widget_frame.grid(row=0, column=1, padx=5, sticky='ew') 

        initial_color = self.parent_gui.button_colors.get(btn_name, (None, None))[1]
        selected_color_var = tk.StringVar(value=initial_color if initial_color else "")

        color_display_label = tk.Label(color_widget_frame, width=4, relief="solid", borderwidth=1)
        color_display_label.pack(side="left", padx=(0, 5))
        try: color_display_label.config(background=initial_color if initial_color else 'SystemButtonFace')
        except tk.TclError: color_display_label.config(background='SystemButtonFace')

        clear_btn = ttk.Button(color_widget_frame, text="X", width=2, style="Toolbutton",
                                command=lambda v=selected_color_var, l=color_display_label: self.parent_gui._set_color_on_widget(v, l, None, self.master))
        clear_btn.pack(side="left", padx=1); ToolTip(clear_btn, f"Clear color for {btn_name}.")

        presets_frame = ttk.Frame(color_widget_frame, style=style_name)
        presets_frame.pack(side="left", padx=(2, 2))
        for p_color in self.pastel_colors[:5]:
            try:
                b = tk.Button(presets_frame, bg=p_color, width=1, height=1, relief="raised", bd=1,
                                  command=lambda c=p_color, v=selected_color_var, l=color_display_label: self.parent_gui._set_color_on_widget(v, l, c, self.master))
                b.pack(side=tk.LEFT, padx=1)
            except tk.TclError: pass

        choose_btn = ttk.Button(color_widget_frame, text="...", width=3, style="Toolbutton",
                                  command=lambda v=selected_color_var, l=color_display_label, n=btn_name: self.parent_gui._choose_color_dialog(v, l, self.master, n))
        choose_btn.pack(side="left", padx=1); ToolTip(choose_btn, f"Choose a custom color for {btn_name}.")

        self.all_button_color_widgets[btn_name] = (selected_color_var, color_display_label)


    def create_sqlite_tab(self):
        tab = ttk.Frame(self.notebook, padding=20); self.notebook.add(tab, text="SQLite Log")
        
        ttk.Label(tab, text="Configure settings for logging data to a SQLite database file.", wraplength=900, justify=tk.LEFT).pack(pady=(0, 10), anchor='w')

        enable_frame = ttk.Frame(tab); enable_frame.pack(fill='x', pady=(0, 15))
        self.sqlite_enabled_var = tk.BooleanVar()
        enable_check = ttk.Checkbutton(enable_frame, text="Enable SQLite Database Logging", variable=self.sqlite_enabled_var, style="Large.TCheckbutton")
        enable_check.pack(side=tk.LEFT, pady=(5, 10)); ToolTip(enable_check, "Check to enable logging events to an SQLite database file.")
        
        config_frame = ttk.LabelFrame(tab, text="SQLite Configuration", padding=15)
        config_frame.pack(fill='x'); config_frame.columnconfigure(1, weight=1)
        
        ttk.Label(config_frame, text="Database File (.db):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.sqlite_db_path_entry = ttk.Entry(config_frame, width=70)
        self.sqlite_db_path_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew"); ToolTip(self.sqlite_db_path_entry, "Full path to the SQLite database file. It will be created if it doesn't exist.")
        db_browse_btn = ttk.Button(config_frame, text="Browse/Create...", command=self.select_sqlite_file)
        db_browse_btn.grid(row=0, column=2, padx=5, pady=5); ToolTip(db_browse_btn, "Browse for an existing SQLite file or specify a name/location for a new one.")
        
        ttk.Label(config_frame, text="Table Name:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.sqlite_table_entry = ttk.Entry(config_frame, width=40)
        self.sqlite_table_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w"); ToolTip(self.sqlite_table_entry, "The name of the table within the database where logs will be written (e.g., 'EventLog'). This table must exist or be created by you.")
        
        test_button = ttk.Button(config_frame, text="Test Connection & Table", command=self.test_sqlite_connection)
        test_button.grid(row=2, column=1, padx=5, pady=15, sticky="w"); ToolTip(test_button, "Verify connection to the database file and check if the specified table exists.")
        
        self.test_result_label = ttk.Label(config_frame, text="", font=("Arial", 9), wraplength=500)
        self.test_result_label.grid(row=3, column=0, columnspan=3, padx=5, pady=2, sticky="w")

    def select_sqlite_file(self):
        filetypes = [("SQLite Database", "*.db"), ("SQLite Database", "*.sqlite"), ("SQLite3 Database", "*.sqlite3"), ("All Files", "*.*")]
        current_path = self.sqlite_db_path_entry.get(); initial_dir = os.path.dirname(current_path) if current_path else os.getcwd()
        filepath = filedialog.asksaveasfilename(parent=self.master, title="Select or Create SQLite Database File", initialdir=initial_dir, initialfile="DataLoggerLog.db", filetypes=filetypes, defaultextension=".db")
        if filepath: self.sqlite_db_path_entry.delete(0, tk.END); self.sqlite_db_path_entry.insert(0, filepath)
        if hasattr(self, 'test_result_label'): self.test_result_label.config(text="")

    def test_sqlite_connection(self):
        db_path = self.sqlite_db_path_entry.get().strip(); table_name = self.sqlite_table_entry.get().strip() or "EventLog"
        if not db_path: self.test_result_label.config(text="❌ Error: Database path is empty.", foreground="red"); return
        conn = None; result_text = ""; result_color = "red"
        try:
            conn = sqlite3.connect(db_path, timeout=3); cursor = conn.cursor()
            result_text = f"✔️ Connection to '{os.path.basename(db_path)}' successful.\n"
            try:
                cursor.execute(f"SELECT 1 FROM [{table_name}] LIMIT 1;");
                result_text += f"✔️ Table '{table_name}' found."; result_color = "green"
            except sqlite3.OperationalError as e_table:
                if "no such table" in str(e_table).lower(): result_text += f"⚠️ Warning: Table '{table_name}' not found. It needs to be created."; result_color = "#E67E00"
                else: raise e_table
            except Exception as e: result_text += f"❌ Error checking table: {e}"; result_color = "red"
        except sqlite3.Error as e: result_text = f"❌ Error connecting/checking DB: {e}"; result_color = "red"
        except Exception as e: result_text = f"❌ Unexpected Error: {e}"; result_color = "red"
        finally:
            if conn: conn.close()
            self.test_result_label.config(text=result_text, foreground=result_color)
            self.master.after(15000, lambda: self.test_result_label.config(text=""))

    # --- Settings Save/Load Logic ---
    def save_settings(self):
        self.parent_gui.log_file_path = self.log_file_entry.get().strip()
        self.parent_gui.txt_folder_path = self.txt_folder_entry_main.get().strip()
        self.parent_gui.txt_folder_path_set2 = self.txt_folder_entry_set2.get().strip()
        self.parent_gui.txt_folder_path_set3 = self.txt_folder_entry_set3.get().strip()

        # Save TXT field columns from the UI
        new_txt_field_configs = []
        for i, row_info in enumerate(self.txt_field_row_widgets):
            field_name = ""
            # For non-fixed fields, read from the entry widget
            if row_info["field_entry_widget"]:
                field_name = row_info["field_entry_widget"].get().strip()
            else: # For fixed fields, get the name from the original config based on its index
                if i < len(self.parent_gui.txt_field_columns_config):
                    field_name = self.parent_gui.txt_field_columns_config[i]["field"]
                else:
                    # Fallback, though this case should ideally not be reached if recreation is consistent
                    field_name = f"Unknown_Field_{i}" 

            column_name = row_info["column_entry"].get().strip()
            skip_value = row_info["skip_var"].get()
            
            # Ensure field_name is not empty for custom fields, assign a default if it is
            # This is important for saving valid data.
            if not field_name and not (field_name in {"Date", "Time", "KP", "DCC", "Line name", "Latitude", "Longitude", "Easting", "Northing", "Event"}):
                field_name = f"Custom_Field_{i+1}"

            new_txt_field_configs.append({
                "field": field_name,
                "column_name": column_name if column_name else field_name, # Default to field name if column is empty
                "skip": skip_value
            })
        self.parent_gui.txt_field_columns_config = new_txt_field_configs
        # Also update the derived dicts for runtime use
        self.parent_gui.txt_field_columns = {cfg["field"]: cfg["column_name"] for cfg in self.parent_gui.txt_field_columns_config}
        self.parent_gui.txt_field_skips = {cfg["field"]: cfg["skip"] for cfg in self.parent_gui.txt_field_columns_config}


        parent_folder_paths = {}; parent_folder_cols = {}; parent_folder_exts = {}; parent_folder_skips = {}
        for folder_name, entry_widget in self.folder_entries.items():
            folder_path = entry_widget.get().strip()
            if folder_path and folder_name not in ["Main TXT File", "TXT Source 2", "TXT Source 3"]:
                parent_folder_paths[folder_name] = folder_path; col_entry = self.folder_column_entries.get(folder_name); ext_entry = self.file_extension_entries.get(folder_name); skip_var = self.folder_skip_vars.get(folder_name)
                parent_folder_cols[folder_name] = col_entry.get().strip() if col_entry and col_entry.get().strip() else folder_name
                parent_folder_exts[folder_name] = ext_entry.get().strip().lstrip('.') if ext_entry else ""
                parent_folder_skips[folder_name] = skip_var.get() if skip_var else False
            elif folder_name in ["Main TXT File", "TXT Source 2", "TXT Source 3"]:
                current_txt_path = ""
                if folder_name == "Main TXT File": current_txt_path = self.parent_gui.txt_folder_path
                elif folder_name == "TXT Source 2": current_txt_path = self.parent_gui.txt_folder_path_set2
                elif folder_name == "TXT Source 3": current_txt_path = self.parent_gui.txt_folder_path_set3

                if current_txt_path:
                    parent_folder_paths[folder_name] = current_txt_path
                    col_entry = self.folder_column_entries.get(folder_name)
                    ext_entry = self.file_extension_entries.get(folder_name)
                    skip_var = self.folder_skip_vars.get(folder_name)
                    
                    parent_folder_cols[folder_name] = col_entry.get().strip() if col_entry and col_entry.get().strip() else folder_name.replace(" ", "_")
                    parent_folder_exts[folder_name] = ext_entry.get().strip().lstrip('.') if ext_entry else "txt"
                    parent_folder_skips[folder_name] = skip_var.get() if skip_var else False
                else:
                    for d in [self.parent_gui.folder_paths, self.parent_gui.folder_columns, self.parent_gui.file_extensions, self.parent_gui.folder_skips]:
                        d.pop(folder_name, None)

        self.parent_gui.folder_paths = parent_folder_paths; self.parent_gui.folder_columns = parent_folder_cols; self.parent_gui.file_extensions = parent_folder_exts; self.parent_gui.folder_skips = parent_folder_skips

        parent_custom_configs = []
        for i, (text_widget, event_widget, txt_source_var) in enumerate(self.custom_button_widgets):
            text = text_widget.get().strip();
            event_text = event_widget.get().strip();
            txt_source_key = txt_source_var.get();

            default_text = f"Custom {i + 1}";
            final_text = text if text else default_text;
            final_event_text = event_text if event_text else f"{final_text} Triggered"

            parent_custom_configs.append({"text": final_text, "event_text": final_event_text, "txt_source_key": txt_source_key})

        self.parent_gui.num_custom_buttons = len(parent_custom_configs)
        self.parent_gui.custom_button_configs = parent_custom_configs

        new_button_colors = {}
        for btn_name, (color_var, _) in self.all_button_color_widgets.items():
            color_hex = color_var.get()
            new_button_colors[btn_name] = (None, color_hex if color_hex else None)
        self.parent_gui.button_colors = new_button_colors

        self.parent_gui.sqlite_enabled = self.sqlite_enabled_var.get(); self.parent_gui.sqlite_db_path = self.sqlite_db_path_entry.get().strip(); self.parent_gui.sqlite_table = self.sqlite_table_entry.get().strip() or "EventLog"

        self.parent_gui.save_settings();
        self.parent_gui.update_custom_buttons();
        self.parent_gui.start_monitoring();
        self.parent_gui.update_db_indicator()

    def load_settings(self):
        self.log_file_entry.delete(0, tk.END); self.log_file_entry.insert(0, self.parent_gui.log_file_path or "")
        
        self.txt_folder_entry_main.delete(0, tk.END); self.txt_folder_entry_main.insert(0, self.parent_gui.txt_folder_path or "")
        self.txt_folder_entry_set2.delete(0, tk.END); self.txt_folder_entry_set2.insert(0, self.parent_gui.txt_folder_path_set2 or "")
        self.txt_folder_entry_set3.delete(0, tk.END); self.txt_folder_entry_set3.insert(0, self.parent_gui.txt_folder_path_set3 or "")

        # Reload TXT field rows based on the (potentially newly loaded) config
        self.recreate_txt_field_rows()
        self.master.after_idle(lambda: self.txt_fields_canvas.config(scrollregion=self.txt_fields_canvas.bbox("all")))


        for name, frame in list(self.folder_row_widgets.items()):
            if frame and frame.winfo_exists(): frame.destroy()
        self.folder_row_widgets.clear(); self.folder_entries.clear(); self.folder_column_entries.clear(); self.file_extension_entries.clear(); self.folder_skip_vars.clear()
        self.add_initial_folder_rows(); self.master.after_idle(self.update_scroll_region)

        self.num_buttons_entry.delete(0, tk.END); self.num_buttons_entry.insert(0, str(self.parent_gui.num_custom_buttons))
        self.recreate_custom_button_settings()

        for btn_name, (color_var, display_label) in self.all_button_color_widgets.items():
            loaded_color_hex = self.parent_gui.button_colors.get(btn_name, (None, None))[1]
            self.parent_gui._set_color_on_widget(color_var, display_label, loaded_color_hex, self.master)

        self.sqlite_enabled_var.set(self.parent_gui.sqlite_enabled); self.sqlite_db_path_entry.delete(0, tk.END); self.sqlite_db_path_entry.insert(0, self.parent_gui.sqlite_db_path or "")
        self.sqlite_table_entry.delete(0, tk.END); self.sqlite_table_entry.insert(0, self.parent_gui.sqlite_table or "EventLog")
        if hasattr(self, 'test_result_label'): self.test_result_label.config(text="")
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

        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()