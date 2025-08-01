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
import sqlite3 # Use Python's built-in SQLite module
import uuid
import pandas as pd
import openpyxl
import re
from pathlib import Path

# --- DEFINED CONSTANTS ---
# PATHS
DEFAULT_SETTINGS_FILE = "settings/default_settings.json"
CUSTOM_SETTINGS_FILE = "settings/custom_settings.json"
EVENT_CODES_FILE = "settings/event_codes.json"

# DICCTIONARY KEYS #NEEDS TO BE REVIEWED
EXCEL_LOG_REQUIRED_COLS = {'runline', 'kp', 'event'} 
DEFAULT_DATA_FIELDS = {"Date", "Time", "KP", "DCC", "Line name", "Latitude", "Longitude", "Easting", "Northing", "Event", "Code"} 
TXT_FILES_KEYS = ["None", "Main TXT", "TXT Source 2", "TXT Source 3"] 
DEFAULT_MONITORED_FOLDERS = ["Qinsy DB", "Naviscan", "SIS", "SSS", "SBP", "Mag", "Grad", "SVP", "SpintINS", "Video", "Cathx", "Hypack RAW", "Eiva NaviPac"]

# DB SETTINGS
DEFAULT_TABLE_NAME = "fieldlog" #NEEDS TO BE REVIEWED

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

# --- FolderMonitor Class ---
class FolderMonitor(FileSystemEventHandler):
    '''
    A custom event handler for watchdog that monitors a specified folder for new or modified files
    matching a given extension. It updates a global cache with the latest matching file.
    '''
    def __init__(self, path, folder_name, gui_instance, extension=""):
        self.path = path
        self.folder_name = folder_name
        self.gui_instance = gui_instance
        self.extension = extension.lower() if extension else ''
        self.latest_file = None
        self.update_latest_file() # Initial scan

    def on_modified(self, event):
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
                for f_name in files:
                    # Check if the file matches the extension (if one is specified)
                    if not self.extension or f_name.lower().endswith(self.extension):
                        f_path = os.path.join(root, f_name)
                        try:
                            mtime = os.path.getmtime(f_path)
                            if mtime > latest_mtime:
                                latest_mtime = mtime
                                latest = f_path
                        except FileNotFoundError:
                            continue # File might have been deleted during the scan
                            
        except FileNotFoundError:
            self.gui_instance.update_status(f"Monitoring error: Folder '{self.path}' not found for '{self.folder_name}'.")
        except Exception as e:
            self.gui_instance.update_status(f"Monitoring error in '{self.folder_name}': {e}")

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
        '''

    # --- Initialization ---
    def __init__(self, master):
        '''
        Initializes the main GUI application.
        This method sets up the main window, initializes styles, variables, and loads settings.
        Arguments:
        * master: The root Tkinter window or parent widget.
        '''

        print("\n--- Starting Online Logger ---\n")

        # Initialize the main window
        self.master = master
        master.title("Online Logger")
        master.geometry("1400x250")
        master.minsize(800, 200)

        self.init_styles()
        self.init_variables()
        self.init_settings()

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
        self.main_frame.rowconfigure(1, weight=0)    # Status bar row

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

        # Create status bar at the very bottom, spanning all columns
        self.create_status_bar()

        # Scheduled tasks
        self.schedule_new_day() # Start the midnight log schedule
        self.schedule_hourly_log() # Start the hourly log schedule
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

        # Settings File Configuration
        self.default_settings_file = DEFAULT_SETTINGS_FILE
        self.settings_file = CUSTOM_SETTINGS_FILE       

       # Event Code Configuration
        self.event_codes_file = EVENT_CODES_FILE
        self.event_codes = {} # Will store {'code': 'description'}

        self.main_button_configs = {
            "Log on": {"event_text": "Log on event occurred", "event_code": ""},
            "Log off": {"event_text": "Log off event occurred", "event_code": ""},
            "Event": {"event_text": "", "event_code": ""}, # Intentionally blank for the "Event" button
            "SVP": {"event_text": "SVP applied", "event_code": ""}
        }
        
        # Original TXT path for the 'Event' button
        self.txt_folder_path = None 
        # New TXT paths for additional sources
        self.txt_folder_path_set2 = None
        self.txt_folder_path_set3 = None

        # Dictionary to hold user-defined aliases for TXT sources
        self.txt_source_aliases = {
            "Main TXT": "Main TXT",
            "TXT Source 2": "TXT Source 2",
            "TXT Source 3": "TXT Source 3"
        }

        self.source_based_colors = {
            "Main TXT": "#BAE1FF",    # Light Blue
            "TXT Source 2": "#BAFFC9",    # Light Green
            "TXT Source 3": "#FFFFBA",    # Light Yellow
            "None": None         # No color for buttons with no source
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
            {"field": "Event", "column_name": "Event", "skip": False}, # Default "Event" field is still here
            {"field": "Code", "column_name": "Code", "skip": False}

        ]
        # These two will be derived from txt_field_columns_config for backwards compatibility/easier lookup
        self.txt_field_columns = {cfg["field"]: cfg["column_name"] for cfg in self.txt_field_columns_config}
        self.txt_field_skips = {cfg["field"]: cfg["skip"] for cfg in self.txt_field_columns_config}


        self.folder_paths = {}
        self.folder_columns = {}
        self.folder_db_columns = {} # --- NEW ---
        self.file_extensions = {}
        self.folder_skips = {}
        self.monitors = {}
        self.num_custom_buttons = 3
        self.MAX_CUSTOM_BUTTONS = 50 # Define the maximum number of custom buttons
        
        # Each custom button config now includes a 'txt_source_key'
        # This key maps to a folder path variable in the GUI instance
        # 'None' means no TXT data is read for this button
        # 'Main TXT' maps to self.txt_folder_path
        # 'TXT Source 2' maps to self.txt_folder_path_set2
        # 'TXT Source 3' maps to self.txt_folder_path_set3
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


        self.sqlite_enabled = False
        self.sqlite_db_path = None
        self.sqlite_table = DEFAULT_TABLE_NAME

        # Variables to control the automatic, timed events
        self.new_day_event_enabled_var = tk.BooleanVar(value=True)
        self.hourly_event_enabled_var = tk.BooleanVar(value=True)

        self.always_on_top_var = tk.BooleanVar(value=False)
        self.settings_window_instance = None # Track settings window
        self.custom_inline_editor_window = None # To track the open inline editor

        self.status_var = tk.StringVar()
        self.monitor_status_label = None
        self.db_status_label = None
        self.settings_window_instance = None # Track settings window
        self.custom_inline_editor_window = None # To track the open inline editor

    def init_settings(self):
        ''' Check if the custom settings file exists and loads it. If not, it load the default settings file.'''
        # Determine which settings file to load
        if not os.path.exists(self.settings_file):
            try:
                print(f"Custom settings not found. Loading from default file: {self.default_settings_file}")
                self.revert_to_defaults()
            except Exception as e:
                messagebox.showwarning("Error in the settings memory", "Paths for custom or default settings files not found", parent=self.master)
        else:
            self.load_settings()

    # --- GUI Creation ---
    def create_main_buttons(self):
        '''
        Builds and renders all the buttons in the GUI dynamically, grouped for better intuitiveness.
        Custom buttons are now organized into tabs within a ttk.Notebook.
        '''
        # Clear existing widgets from all three frames
        for frame in [self.custom_buttons_frame, self.general_buttons_frame, self.config_frame]:
            for widget in frame.winfo_children():
                widget.destroy()
        self.custom_buttons = [] # Reset custom_buttons list

        # --- Section 1: Custom Events (Left Side) ---
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

        # Prepare and sort custom button data by tab
        custom_buttons_by_tab = {group: [] for group in all_tab_groups if group}
        for config in self.custom_button_configs[:self.num_custom_buttons]:
            tab_group = config.get("tab_group", "Main")
            if tab_group not in custom_buttons_by_tab:
                custom_buttons_by_tab[tab_group] = []
            custom_buttons_by_tab[tab_group].append(config)

        # Create and grid custom buttons inside their tabs
        for tab_group, configs in custom_buttons_by_tab.items():
            if tab_group in self.custom_button_tab_frames:
                tab_frame = self.custom_button_tab_frames[tab_group]
                for i, config in enumerate(configs):
                    button_text = config.get("text", "Custom")
                    event_desc = config.get("event_text", "Triggered")
                    txt_source = config.get("txt_source_key", "None")

                    # Retrieve configured background and font colors for this specific button
                    # Fallback to source_based_colors for background if button-specific not set
                    bg_color_hex, font_color_hex = self.button_colors.get(button_text, (None, None))
                    
                    # If button-specific background is not set, try source_based_colors
                    if not bg_color_hex:
                        bg_color_hex = self.source_based_colors.get(txt_source)
                    
                    # Create a unique style name for this button
                    # Use a clean version of button_text for the style name
                    cleaned_button_text = ''.join(e for e in button_text if e.isalnum()) 
                    style_name = f"CustomBtn_{cleaned_button_text}.TButton"
                    
                    # Configure the specific style for this button
                    style_config = {}
                    if bg_color_hex:
                        style_config['background'] = bg_color_hex
                    if font_color_hex:
                        style_config['foreground'] = font_color_hex # This is where font color is applied

                    # Configure or re-configure the style based on collected colors
                    # Ensure font is always set, and padding is maintained
                    self.style.configure(style_name, font=("Arial", 10, "bold"), padding=4, **style_config)
                    
                    button = ttk.Button(tab_frame, text=button_text, style=style_name)
                    # Corrected: lambda function for command
                    button.config(command=lambda c=config, b=button: self.log_custom_event(c, b))
                    

                    # Calculate row and column based on the number of columns
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
        general_lf.rowconfigure((0, 1), weight=1)

        # --- Helper function to create styled main buttons ---
        def create_main_button(parent, text, command_func, tooltip_text, grid_row, grid_col):
            # 1. Get the configured colors (background, font)
            bg_color_hex, font_color_hex = self.button_colors.get(text, (None, None))
            
            # 2. Create a unique style for this button
            cleaned_text = ''.join(e for e in text if e.isalnum()) 
            style_name = f"MainBtn_{cleaned_text}.TButton"
            
            # 3. Configure the style with the colors, if they are set
            style_config = {}
            if bg_color_hex:
                style_config['background'] = bg_color_hex
            if font_color_hex:
                style_config['foreground'] = font_color_hex
            
            # Ensure font is always set, and padding is maintained
            self.style.configure(style_name, font=("Arial", 10, "bold"), padding=4, **style_config)
            
            # 4. Create the button with the dynamic style
            btn = ttk.Button(parent, text=text, style=style_name, command=command_func) # Command is now correctly passed
            btn.grid(row=grid_row, column=grid_col, padx=4, pady=4, sticky="nsew")
            
            # 5. Add right-click menu and tooltip
            btn.bind("<Button-3>", lambda e, name=text: self._show_main_button_context_menu(e, name))
            ToolTip(btn, tooltip_text)
            return btn

        # --- Create the buttons using the helper function ---
        # The lambda for the command needs to wrap the function call to ensure the button itself is passed
        # and that the logging function is called *when the button is clicked*, not when it's created.
        create_main_button(general_lf, "Log on", lambda b=None: self.log_event("Log on", b, "Main TXT"), "Record a 'Log on' marker.", 0, 0)
        create_main_button(general_lf, "Log off", lambda b=None: self.log_event("Log off", b, "Main TXT"), "Record a 'Log off' marker.", 1, 0)
        create_main_button(general_lf, "Event", lambda b=None: self.log_event("Event", b, "Main TXT"), "Record data from the Main TXT source.", 0, 1)
        create_main_button(general_lf, "SVP", lambda b=None: self.log_svp("SVP", b, "Main TXT"), "Record data and insert latest SVP filename.", 1, 1)


        # --- Section 3: Configuration Buttons (Right Side) ---
        config_lf = ttk.LabelFrame(self.config_frame, text="Configuration")
        config_lf.grid(row=0, column=0, sticky="new")
        self.config_frame.columnconfigure(0, weight=1)
        config_lf.columnconfigure(0, weight=1)

        btn_settings = ttk.Button(config_lf, text="Settings", style="Small.TButton", command=self.open_settings)
        btn_settings.grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 2))
        ToolTip(btn_settings, "Open the configuration window.")

        btn_sync = ttk.Button(config_lf, text="Sync DB", style="Small.TButton", command=self.sync_excel_to_sqlite_triggered)
        btn_sync.grid(row=1, column=0, sticky="ew", padx=4, pady=2)
        ToolTip(btn_sync, "Update SQLite DB from the Excel log.")

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

        # SQLite Status
        ttk.Label(indicator_lf, text="SQLite:", font=("Arial", 8, "bold")).grid(row=1, column=0, sticky="w", padx=4, pady=2)
        self.db_status_label = ttk.Label(indicator_lf, text="...", foreground="orange", font=("Arial", 8))
        self.db_status_label.grid(row=1, column=1, sticky="w", padx=4, pady=2)
        
        # Always on Top Checkbox
        always_on_top_check = ttk.Checkbutton(
            indicator_lf,
            text="Always on Top",
            variable=self.always_on_top_var,
            command=self.toggle_always_on_top
        )
        always_on_top_check.grid(row=2, column=0, columnspan=2, sticky='w', padx=4, pady=(5, 2))
        ToolTip(always_on_top_check, "If checked, this window will always stay on top.")

        self.update_db_indicator()

    def create_status_bar(self):
        '''
        Creates a status bar at the bottom of the main window to display status messages.
        This method initializes a label that will show the current status of the application, such as monitoring status, database connection status, and other messages.
        '''
        self.status_var.set("Status: Ready")
        status_bar = ttk.Label(self.main_frame, textvariable=self.status_var, style="StatusBar.TLabel", anchor='w')
        status_bar.grid(row=1, column=0, columnspan=3, sticky="ew")
    
    # --- Synchronization Excel Log to SQLite Database ---
    def _find_header_row(self, excel_file, engine, required_column='event', max_rows_to_scan=MAX_HEADER_SEARCH_ROW):
        """
        Scans the top N rows of an Excel sheet to find the row index of the header.
        The header is identified by the presence of a specific required column (e.g., 'GUID').

        Args:
            excel_file (str): Path to the Excel file.
            engine (str): The pandas engine to use ('openpyxl' or 'pyxlsb').
            required_column (str): A column name that MUST be in the header row.
            max_rows_to_scan (int): The number of rows to scan from the top.

        Returns:
            int: The zero-based index of the header row.

        Raises:
            ValueError: If the required column is not found in the scanned rows.
        """
        # Read only the top part of the file without assuming any header
        df_top = pd.read_excel(
            excel_file,
            engine=engine,
            header=None,
            nrows=max_rows_to_scan
        )
        for idx, row in df_top.iterrows():
            # Check if the required column name is in the current row's values
            # Comparing as lowercase strings for robustness
            row_values = [str(v).lower() for v in row.values]
            if required_column.lower() in row_values:
                return idx  # Return the index of the found header row

        # If the loop finishes, the header was not found
        raise ValueError(f"Crucial '{required_column}' column not found in the first {max_rows_to_scan} rows.")
    
    # Add this new helper method inside the DataLoggerGUI class
    def _values_are_different(self, val1, val2):
        """Robustly compares two values, handling None, pandas NaT/NaN, and numeric types."""
        # If both are considered null, they are not different
        if (pd.isna(val1) or val1 is None) and (pd.isna(val2) or val2 is None):
            return False
        # If one is null and the other is not, they are different
        if (pd.isna(val1) or val1 is None) or (pd.isna(val2) or val2 is None):
            return True
        
        # Try a numeric comparison with a small tolerance for floating point issues
        try:
            if abs(float(val1) - float(val2)) < 0.00001:
                return False
        except (ValueError, TypeError):
            # If they can't be converted to float, proceed to string comparison
            pass

        # Fallback to string comparison for all other types (dates, text, etc.)
        return str(val1) != str(val2)

    
    def update_or_insert_record(self, record_to_process, cursor, db_table, excel_to_db_map, db_cols_set, orphaned_rows_df):
        """
        Intelligently updates an existing orphaned record or inserts a new one.
        """
        time_fix_val = record_to_process.get('time_fix')
        excel_event_col = self.txt_field_columns.get("Event")
        event_val = record_to_process.get(excel_event_col, "") if excel_event_col else ""
        db_event_col = excel_to_db_map.get(excel_event_col) if excel_event_col else None

        match = None
        if time_fix_val and not orphaned_rows_df.empty:
            if db_event_col and db_event_col in orphaned_rows_df.columns:
                match_mask = (orphaned_rows_df['time_fix'] == time_fix_val) & \
                            (orphaned_rows_df[db_event_col].fillna('') == (event_val or ''))
            else:
                match_mask = (orphaned_rows_df['time_fix'] == time_fix_val)

            potential_matches = orphaned_rows_df[match_mask]
            if len(potential_matches) == 1:
                match = potential_matches.iloc[0]

        cols_to_use = [excel_to_db_map[k] for k in record_to_process.keys() if k in excel_to_db_map and excel_to_db_map[k] in db_cols_set]
        vals_to_use = [record_to_process[k] for k in record_to_process.keys() if k in excel_to_db_map and excel_to_db_map[k] in db_cols_set]

        if db_event_col in cols_to_use:
            try:
                event_col_index = cols_to_use.index(db_event_col)
                if vals_to_use[event_col_index] is None:
                    vals_to_use[event_col_index] = ''
            except (ValueError, IndexError):
                pass

        placeholders = ', '.join(['?'] * len(cols_to_use))
        sql_insert = f"INSERT INTO \"{db_table}\" ({', '.join(f'\"{c}\"' for c in cols_to_use)}) VALUES ({placeholders})"
        row_num_db_col_name = excel_to_db_map.get("original_excel_row")

        if match is not None:
            old_row_to_delete = match[row_num_db_col_name]
            sql_delete = f"DELETE FROM \"{db_table}\" WHERE \"{row_num_db_col_name}\" = ?"
            cursor.execute(sql_delete, (old_row_to_delete,))
            cursor.execute(sql_insert, vals_to_use)
            return "REPLACE"
        else:
            cursor.execute(sql_insert, vals_to_use)
            return "INSERT"

    def sync_excel_to_sqlite_triggered(self):
        should_proceed = messagebox.askokcancel(
            "Save Before Syncing",
            "Please ensure the Excel log file has been saved before proceeding. You can save it now if you forgot it.\n\n"
            "Click OK to continue with the sync.\n"
            "Click Cancel to stop the operation.",
            icon='warning',
            parent=self.master
        )

        if not should_proceed:
            self.update_status("Sync cancelled by user.")
            return

        if not self.sqlite_enabled:
            messagebox.showwarning("Sync Skipped", "SQLite logging is not enabled in Settings.", parent=self.master)
            return
        if not self.log_file_path or not os.path.exists(self.log_file_path):
            messagebox.showerror("Sync Error", "Excel log file path is not set or the file does not exist.", parent=self.master)
            return
        if not self.sqlite_db_path:
            messagebox.showerror("Sync Error", "SQLite database path is not set.", parent=self.master)
            return

        try:
            excel_file = self.log_file_path
            excel_engine = 'pyxlsb' if excel_file.lower().endswith('.xlsb') else 'openpyxl'
            header_row = self._find_header_row(excel_file, excel_engine)
            df_excel = pd.read_excel(excel_file, engine=excel_engine, header=header_row)
            df_excel['original_excel_row'] = df_excel.index + header_row + 2
        except Exception as e_check:
            traceback.print_exc()
            messagebox.showerror("Pre-check Error", f"An unexpected error occurred while checking the file:\n{e_check}", parent=self.master)
            return

        sync_button = next((btn for lf in self.config_frame.winfo_children() if isinstance(lf, ttk.LabelFrame) for btn in lf.winfo_children() if isinstance(btn, ttk.Button) and btn.cget('text') == "Sync DB"), None)

        if sync_button:
            original_text = sync_button['text']
            sync_button.config(state=tk.DISABLED, text="Syncing...")

        self.update_status("Starting sync from Excel to SQLite...")

        def _sync_worker():
            success, message = self.perform_excel_to_sqlite_sync()
            self.master.after(0, self.update_status, message)
            if sync_button:
                self.master.after(0, lambda: sync_button.config(state=tk.NORMAL, text=original_text))

        threading.Thread(target=_sync_worker, daemon=True).start()

    
    def perform_excel_to_sqlite_sync(self):
        print("\n--- Starting FULL Excel to SQLite Sync (Inserts & Updates) ---")
        
        if not all([self.log_file_path, self.sqlite_db_path, self.sqlite_table]):
            return False, "Sync Error: Configuration missing."
        if not os.path.exists(self.log_file_path):
            return False, "Sync Error: Excel file not found."

        conn_sqlite = None
        try:
            # --- 1. Read data from Excel using xlwings ---
            wb = xw.Book(self.log_file_path)
            sheet = wb.sheets[0]
            
            header_row_index = -1
            header_values = []
            for i in range(1, MAX_HEADER_SEARCH_ROW + 1):
                row_values = sheet.range(f'A{i}').expand('right').value
                if row_values and all(item in [str(v).lower().strip() for v in row_values if v] for item in EXCEL_LOG_REQUIRED_COLS):
                    header_row_index = i
                    header_values = row_values
                    break
            if header_row_index == -1: return False, "Sync Error: Could not find header row."

            used_range = sheet.used_range
            last_col_letter = xw.utils.col_name(used_range.last_cell.column)
            data_range = sheet.range(f'A{header_row_index}:{last_col_letter}{used_range.last_cell.row}')
            
            # --- FIX STARTS HERE ---
            # Read data as raw objects to prevent pandas from auto-detecting dates incorrectly.
            df_excel = data_range.options(pd.DataFrame, header=1, dtype=object).value 
            df_excel = df_excel.reset_index(drop=True) # Ensure a clean, simple integer index.
            # --- FIX ENDS HERE ---

            df_excel.columns = [str(c) for c in df_excel.columns]
            df_excel['original_excel_row'] = df_excel.index + header_row_index + 1
            print(f"Read {len(df_excel)} data rows from Excel.")
            
            # --- 2. Build the Excel-to-DB column map ---
            excel_to_db_map = {item["column_name"]: item.get("db_column_name") for item in self.txt_field_columns_config if item.get("column_name") and item.get("db_column_name")}
            for folder, excel_col in self.folder_columns.items():
                db_col = self.folder_db_columns.get(folder)
                if excel_col and db_col:
                    excel_to_db_map[excel_col] = db_col
            excel_to_db_map["original_excel_row"] = 'original_excel_row'
            
            # --- 3. Connect to SQLite and get ALL existing data ---
            conn_sqlite = sqlite3.connect(self.sqlite_db_path, timeout=10)
            db_cols_set = set(pd.read_sql_query(f"PRAGMA table_info('{self.sqlite_table}')", conn_sqlite)['name'])
            row_num_db_col_name = excel_to_db_map.get("original_excel_row")
            
            if not row_num_db_col_name or row_num_db_col_name not in db_cols_set:
                return False, f"Sync Error: Crucial column '{row_num_db_col_name}' missing in DB."

            df_sqlite = pd.read_sql_query(f'SELECT * FROM "{self.sqlite_table}"', conn_sqlite)
            print(f"Read {len(df_sqlite)} rows from SQLite for comparison.")
            
            # --- 4. Compare datasets to find inserts AND updates ---
            excel_rows_set = set(df_excel['original_excel_row'])
            db_rows_set = set(df_sqlite[row_num_db_col_name]) if not df_sqlite.empty else set()
            
            new_row_indices = list(excel_rows_set - db_rows_set)
            records_to_insert = df_excel[df_excel['original_excel_row'].isin(new_row_indices)].to_dict('records')
            
            records_to_update = []
            common_row_indices = list(excel_rows_set.intersection(db_rows_set))
            
            if common_row_indices:
                df_excel_common = df_excel[df_excel['original_excel_row'].isin(common_row_indices)].set_index('original_excel_row')
                df_sqlite_common = df_sqlite[df_sqlite[row_num_db_col_name].isin(common_row_indices)].set_index(row_num_db_col_name)

                for row_num in common_row_indices:
                    excel_row = df_excel_common.loc[row_num]
                    sqlite_row = df_sqlite_common.loc[row_num]
                    
                    updates = {}
                    for excel_col, db_col in excel_to_db_map.items():
                        if excel_col in excel_row.index and db_col in sqlite_row.index:
                            excel_val = excel_row[excel_col]
                            sqlite_val = sqlite_row[db_col]
                            if self._values_are_different(excel_val, sqlite_val):
                                updates[db_col] = excel_val
                    
                    if updates:
                        updates[row_num_db_col_name] = row_num 
                        records_to_update.append(updates)

            # --- 5. Write changes to SQLite ---
            cursor = conn_sqlite.cursor()
            inserted_count = 0
            updated_count = 0

            if records_to_insert:
                print(f"Inserting {len(records_to_insert)} new rows...")
                for record in records_to_insert:
                    db_record = {excel_to_db_map.get(k): v for k, v in record.items() if excel_to_db_map.get(k) in db_cols_set and pd.notna(v)}
                    date_excel_col = self.txt_field_columns.get("Date")
                    time_excel_col = self.txt_field_columns.get("Time")
                    if date_excel_col in record and time_excel_col in record:
                        date_val = record.get(date_excel_col)
                        time_val = record.get(time_excel_col)
                        if isinstance(date_val, (int, float)) and isinstance(time_val, (int, float)):
                            dt_val = pd.to_datetime(date_val + time_val, unit='D', origin='1899-12-30')
                            db_record['time_fix'] = dt_val.strftime('%Y-%m-%d %H:%M:%S')
                        elif date_val and time_val:
                            db_record['time_fix'] = f"{date_val} {time_val}"
                    event_excel_col = self.txt_field_columns.get("Event")
                    event_db_col = excel_to_db_map.get(event_excel_col)
                    if event_db_col and event_db_col not in db_record:
                        db_record[event_db_col] = ""
                    if not db_record or 'time_fix' not in db_record: continue
                    cols = ', '.join([f'"{c}"' for c in db_record.keys()])
                    placeholders = ', '.join(['?'] * len(db_record))
                    sql = f'INSERT INTO "{self.sqlite_table}" ({cols}) VALUES ({placeholders})'
                    cursor.execute(sql, list(db_record.values()))
                    inserted_count += cursor.rowcount

            if records_to_update:
                print(f"Updating {len(records_to_update)} modified rows...")
                for update_record in records_to_update:
                    row_val = update_record.pop(row_num_db_col_name) 
                    
                    event_excel_col = self.txt_field_columns.get("Event")
                    event_db_col = excel_to_db_map.get(event_excel_col)
                    if event_db_col in update_record and update_record[event_db_col] is None:
                        update_record[event_db_col] = ""

                    if not update_record: continue 

                    set_clauses = [f'"{col}" = ?' for col in update_record.keys()]
                    values = list(update_record.values()) + [row_val]
                    sql = f'UPDATE "{self.sqlite_table}" SET {", ".join(set_clauses)} WHERE "{row_num_db_col_name}" = ?'
                    cursor.execute(sql, values)
                    updated_count += cursor.rowcount

            conn_sqlite.commit()
            print("Sync complete.")
            return True, f"Sync successful. Inserted: {inserted_count}. Updated: {updated_count}."
            
        except Exception as e:
            traceback.print_exc()
            if conn_sqlite: conn_sqlite.rollback()
            return False, f"Sync Error: {e}"
        finally:
            if conn_sqlite:
                conn_sqlite.close()

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
        event_text_for_excel = self.main_button_configs.get(event_type, {}).get("event_text", f"Default {event_type}")
        skip_files = (event_type == "Event") # Still skip files only for the main "Event" button
            
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
        txt_source_key = config.get("txt_source_key", "None") # This is correctly getting the key
        
        self._perform_log_action(event_type=button_text,
                                 event_text_for_excel=event_text_for_excel,
                                 triggering_button=button_widget,
                                 txt_source_key=txt_source_key) # This is correctly passing it

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
                                 svp_specific_handling=True,
                                 triggering_button=button_widget,
                                 txt_source_key=txt_source_key) 

    def _perform_log_action(self, event_type, event_text_for_excel, skip_latest_files=False, svp_specific_handling=False, triggering_button=None, txt_source_key="Main TXT"):
        self.update_status(f"Processing '{event_type}'...")
        print(f"\n--- Log Action Initiated for '{event_type}' ---")

        original_text = None
        if triggering_button and isinstance(triggering_button, ttk.Button):
            try:
                if triggering_button.winfo_exists():
                    original_text = triggering_button['text']
                    triggering_button.config(state=tk.DISABLED, text="Working...")
            except tk.TclError:
                triggering_button = None

        def _log_thread_func():
            status_msg = f"'{event_type}' failed."
            try:
                # --- This part of the logic for gathering data remains the same ---
                row_data = {}

                # 1. Get data from the selected TXT source
                if txt_source_key and txt_source_key != "None":
                    source_folder_path = None
                    if txt_source_key == "Main TXT":
                        source_folder_path = self.txt_folder_path
                    elif txt_source_key == "TXT Source 2":
                        source_folder_path = self.txt_folder_path_set2
                    elif txt_source_key == "TXT Source 3":
                        source_folder_path = self.txt_folder_path_set3

                    if source_folder_path and os.path.isdir(source_folder_path):
                        try:
                            txt_data = self._get_txt_data_from_source(source_folder_path)
                            if txt_data:
                                row_data.update(txt_data)
                        except Exception as e_txt:
                            print(f"Error getting TXT data from source '{txt_source_key}': {e_txt}")

                # 2. Get static data from Excel cells (this correctly uses xlwings)
                try:
                    static_data_from_cells = self._get_static_excel_data()
                    if static_data_from_cells:
                        row_data.update(static_data_from_cells)
                except Exception as e_static:
                    print(f"Error getting static data from Excel cells: {e_static}")
                
                # 3. Get latest files from monitored folders (using the OPTIMIZED function)
                if not skip_latest_files:
                    try:
                        latest_files_data = self.get_latest_files_data_fast() # Use the new fast function
                        if latest_files_data:
                            row_data.update(latest_files_data)
                    except Exception as e_files:
                        print(f"Error getting latest file data (monitored folders): {e_files}")

                # 4. Handle SVP specific logic
                if svp_specific_handling:
                    svp_col_name = self.folder_columns.get("SVP")
                    if svp_col_name:
                        latest_svp = folder_cache.get("SVP")
                        if latest_svp:
                            row_data[svp_col_name] = os.path.splitext(os.path.basename(latest_svp))[0]
                        else:
                            row_data[svp_col_name] = "N/A"

                # 5. Add event text and code
                event_column_name = self.txt_field_columns.get("Event")
                if event_column_name:
                    row_data[event_column_name] = event_text_for_excel

                event_code_to_log = ""
                if event_type in [cfg['text'] for cfg in self.custom_button_configs]:
                    for cfg in self.custom_button_configs:
                        if cfg['text'] == event_type:
                            event_code_to_log = cfg.get("event_code", "")
                            break
                elif event_type in self.main_button_configs:
                    event_code_to_log = self.main_button_configs[event_type].get("event_code", "")

                code_column_name = self.txt_field_columns.get("Code")
                if code_column_name and event_code_to_log:
                    row_data[code_column_name] = event_code_to_log

                # --- 🚀 REFACTORED SAVING LOGIC ---
                if row_data:
                    color_tuple = self.button_colors.get(event_type, (None, None))
                    row_color = color_tuple[0] if isinstance(color_tuple, tuple) and len(color_tuple) > 0 else None
                    font_color = color_tuple[1] if isinstance(color_tuple, tuple) and len(color_tuple) > 1 else None

                    # Call the new, combined save function
                    success, message = self.save_to_excel_and_sqlite(row_data, row_color, font_color)
                    if success:
                        status_msg = f"'{event_type}' logged. {message}"
                    else:
                        status_msg = f"'{event_type}' failed. {message}"
                else:
                    status_msg = f"'{event_type}' - No data collected."

            except Exception as thread_ex:
                traceback.print_exc()
                status_msg = f"'{event_type}' - Thread error: {thread_ex}"
                self.master.after(0, lambda e=thread_ex: messagebox.showerror("Thread Error", f"Critical error in logging thread:\n{e}", parent=self.master))
            finally:
                self.master.after(0, self.update_status, status_msg)
                if triggering_button and isinstance(triggering_button, ttk.Button):
                    def re_enable_button(btn=triggering_button, txt=original_text):
                        try:
                            if btn and btn.winfo_exists():
                                btn.config(state=tk.NORMAL)
                                if txt:
                                    btn.config(text=txt)
                        except tk.TclError:
                            pass
                    self.master.after(0, re_enable_button)

        threading.Thread(target=_log_thread_func, daemon=True).start()

    # --- TXT Reading and Writting ---
    def _get_txt_data_from_source(self, folder_path):
        """
        Reads and parses data from the latest TXT file in the specified folder.
        Returns a dictionary of parsed data or empty dict if no data/errors.
        """
        row_data = {}
        current_dt = datetime.datetime.now()

        latest_txt_file_path = None
        if folder_path and os.path.exists(folder_path):
            latest_txt_file_path = self.find_latest_file_in_folder(folder_path, ".txt")
            if not latest_txt_file_path:
                latest_txt_file_path = self.find_latest_file_in_folder(folder_path, ".csv") # Try CSV as fallback
        else:
            print(f"TXT folder path is invalid or empty: {folder_path}")
            return row_data # Return empty if path is invalid

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
                                # Add a small delay to avoid file contention if file is being written to
                                time.sleep(0.05) 
                                with open(latest_txt_file_path, "r", encoding=enc) as file:
                                    lines = file.readlines()
                                read_success = True
                                break
                            except IOError as e_io:
                                print(f"IOError reading TXT file '{latest_txt_file_path}' on attempt {attempt+1} with encoding {enc}: {e_io}") #DEBUG
                                if attempt < 2:
                                    time.sleep(0.1) # Wait a bit longer if file is busy
                                    continue
                                else:
                                    raise
                        if read_success:
                            break
                    except UnicodeDecodeError:
                        print(f"UnicodeDecodeError with encoding: {enc} for file {latest_txt_file_path}") #DEBUG
                        continue
                    except Exception as e_read:
                        print(f"Unexpected error during TXT file read for {latest_txt_file_path}: {e_read}") #DEBUG
                        lines = []
                        break

                if not read_success or not lines:
                    print(f"Failed to read any lines from TXT file: {latest_txt_file_path}")
                    return row_data

                last_line_str = lines[-1].strip() # Get the last line of the file
                latest_line_parts = last_line_str.split(",")
                
                # Iterate through the ordered config for TXT fields
                for i, field_config in enumerate(self.txt_field_columns_config):
                    field_key = field_config["field"]
                    excel_col = field_config["column_name"]
                    skip_field = field_config["skip"]

                    if excel_col and not skip_field:
                        if field_key in ["Date", "Time"]:
                            continue # PC time will be used for these

                        if i < len(latest_line_parts):
                            value = latest_line_parts[i].strip()
                            temp_txt_data[excel_col] = value
                            #print(f"Parsed TXT data: '{field_key}' -> '{excel_col}': '{value}'") #DEBUG
                        else:
                            temp_txt_data[excel_col] = None # Field not found at expected index
                            #print(f"TXT field '{field_key}' (index {i}) has no corresponding data in line parts.")  #DEBUG
            except Exception as e:
                print(f"Major error during TXT parsing: {e}")

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

        return row_data

    def get_latest_files_data_fast(self):
        """
        FAST version that collects latest file data by reading directly from the
        watchdog's global cache, avoiding slow disk scans.
        """
        data = {}
        for folder_name, column_name in self.folder_columns.items():
            if not self.folder_skips.get(folder_name, False) and column_name:
                # Read directly from the cache
                latest_file_path = folder_cache.get(folder_name)
                
                if latest_file_path and os.path.exists(latest_file_path):
                    # Store the filename without the extension
                    filename_without_ext, _ = os.path.splitext(os.path.basename(latest_file_path))
                    data[column_name] = filename_without_ext
                else:
                    data[column_name] = "N/A" # No file found/cached yet
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
            
            # --- MODIFIED: Handles empty extension and formats the glob pattern ---
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

    def _get_static_excel_data(self):
        """
        Reads data from specific cells in the Excel log file based on the
        "='SheetName'!Cell" syntax in the data mapping configuration.
        """
        static_data = {}
        # Filter for configs that use the static cell lookup syntax
        cell_lookup_configs = [
            item for item in self.txt_field_columns_config
            if str(item.get("column_name")).startswith('=') and not item.get("skip")
        ]

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
                # The key for the returned dictionary is the "Excel Column" name itself,
                # as this is what the rest of the pipeline expects.
                excel_col_key = config["column_name"]
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

    def save_to_excel_and_sqlite(self, row_data, row_color=None, font_color=None):
        """
        Efficiently saves a single row of data to the open Excel file via xlwings
        and directly inserts the corresponding record into the SQLite database.
        This avoids slow, full-file reads.
        """
        if not self.log_file_path or not os.path.exists(self.log_file_path):
            return False, "Excel: Path Invalid."

        excel_message = "Excel: Fail."
        sqlite_message = "SQLite: Skipped."
        next_row = -1

        # --- 1. Write to Excel using xlwings ---
        try:
            wb = xw.Book(self.log_file_path)
            sheet = wb.sheets[0]

            header_row_index = -1
            header_values = []
            for i in range(1, MAX_HEADER_SEARCH_ROW + 1):
                row_values_list = sheet.range(f'A{i}').expand('right').value
                if not row_values_list: continue
                current_row_headers = {str(h).lower().strip() for h in row_values_list if h is not None}
                if EXCEL_LOG_REQUIRED_COLS.issubset(current_row_headers):
                    header_row_index = i
                    header_values = row_values_list
                    break
            
            if header_row_index == -1:
                raise ValueError("Could not find header row in Excel.")

            header_map = {str(h).lower(): i for i, h in enumerate(header_values) if h}
            
            last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
            next_row = max(last_row, header_row_index) + 1

            output_data = [None] * len(header_map)
            for col_name, value in row_data.items():
                col_name_lower = str(col_name).lower()
                if col_name_lower in header_map:
                    col_idx = header_map[col_name_lower]
                    output_data[col_idx] = value
            
            target_range = sheet.range(f"A{next_row}").resize(1, len(output_data))
            target_range.value = output_data

            if row_color or font_color:
                format_range = sheet.range((next_row, 1), (next_row, len(header_map)))
                if row_color:
                    format_range.color = row_color
                if font_color:
                    format_range.font.color = font_color
            
            excel_message = "Excel: OK."
            
        except Exception as e:
            traceback.print_exc()
            excel_message = f"Excel: Fail ({type(e).__name__})."
            return False, f"{excel_message} {sqlite_message}"

        # --- 2. Write to SQLite ---
        if self.sqlite_enabled and next_row != -1:
            sqlite_message = "SQLite: Fail."
            conn = None
            try:
                db_data = {}
                for item in self.txt_field_columns_config:
                    excel_col = item.get("column_name")
                    db_col = item.get("db_column_name")
                    if excel_col and db_col and excel_col in row_data:
                        db_data[db_col] = row_data[excel_col]
                
                for folder, excel_col in self.folder_columns.items():
                    db_col = self.folder_db_columns.get(folder)
                    if excel_col and db_col and excel_col in row_data:
                        db_data[db_col] = row_data[excel_col]

                db_data['original_excel_row'] = next_row

                # --- FIX STARTS HERE ---
                # Manually construct the 'time_fix' field required by the database.
                date_excel_col = self.txt_field_columns.get("Date")
                time_excel_col = self.txt_field_columns.get("Time")
                if date_excel_col in row_data and time_excel_col in row_data:
                    date_val = row_data.get(date_excel_col)
                    time_val = row_data.get(time_excel_col)
                    if date_val and time_val:
                        db_data['time_fix'] = f"{date_val} {time_val}"
                # --- FIX ENDS HERE ---

                if db_data:
                    # Ensure the mandatory time_fix column exists before trying to insert
                    if 'time_fix' not in db_data or not db_data['time_fix']:
                        raise sqlite3.IntegrityError("NOT NULL constraint failed: fieldlog.time_fix (generated value was empty)")

                    conn = sqlite3.connect(self.sqlite_db_path, timeout=10)
                    cursor = conn.cursor()
                    
                    cols = ', '.join([f'"{c}"' for c in db_data.keys()])
                    placeholders = ', '.join(['?'] * len(db_data))
                    sql = f'INSERT INTO "{self.sqlite_table}" ({cols}) VALUES ({placeholders})'
                    
                    cursor.execute(sql, list(db_data.values()))
                    conn.commit()
                    sqlite_message = "SQLite: OK."
                else:
                    sqlite_message = "SQLite: No data mapped."

            except Exception as e:
                if conn: conn.rollback()
                traceback.print_exc()
                sqlite_message = f"SQLite: Fail ({type(e).__name__})."
            finally:
                if conn: conn.close()

        return True, f"{excel_message} {sqlite_message}"

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

    # --- Settings Saving and Loading ---
    def save_settings(self):
        '''Saves the current settings to the JSON file. Cleans up button colors for custom buttons that no longer exist.'''
        print("\n--- Saving Settings ---") 
        colors_to_save = {}
        for key, (bg_color, font_color) in self.button_colors.items(): # Iterate over (bg_color, font_color) tuples
            # Save the tuple if at least one color is set
            if bg_color or font_color:
                colors_to_save[key] = (bg_color, font_color)
        settings = {
            "log_file_path": self.log_file_path,
            "txt_folder_path": self.txt_folder_path,
            "txt_folder_path_set2": self.txt_folder_path_set2,
            "txt_folder_path_set3": self.txt_folder_path_set3,
            "txt_field_columns_config": self.txt_field_columns_config,
            "folder_paths": self.folder_paths,
            "folder_columns": self.folder_columns,
            "folder_db_columns": self.folder_db_columns, # --- NEW ---
            "file_extensions": self.file_extensions, 
            "folder_skips": self.folder_skips,
            "num_custom_buttons": self.num_custom_buttons,
            "custom_button_configs": self.custom_button_configs,
            "custom_button_tab_groups": self.custom_button_tab_groups, # NEW: Save tab groups
            "button_colors": colors_to_save, 
            "sqlite_enabled": self.sqlite_enabled,
            "sqlite_db_path": self.sqlite_db_path, 
            "sqlite_table": self.sqlite_table,
            "always_on_top": self.always_on_top_var.get(),
            "new_day_event_enabled": self.new_day_event_enabled_var.get(),
            "hourly_event_enabled": self.hourly_event_enabled_var.get(),
            "main_button_configs": self.main_button_configs,
            "txt_source_aliases": self.txt_source_aliases
        }
        try:
            with open(self.settings_file, 'w') as f: 
                json.dump(settings, f, indent=4)
            print(f"Settings successfully saved to {self.settings_file}")
            self.update_status("Settings saved.")
        except Exception as e:
            print(f"Error saving settings: {e}") #DEBUG
            messagebox.showerror("Save Error", f"Could not save settings to {self.settings_file}:\n{e}", parent=self.master)
            self.update_status("Error saving settings.")

    def revert_to_defaults(self):
        """
        Deletes the user settings file, then forces a reload from the default
        settings file, updates the UI, and restarts services.
        """
        print("\n--- Reverting to Default Settings ---")

        # Check if the default settings file exists before proceeding
        if not os.path.exists(self.default_settings_file):
            raise FileNotFoundError(f"The default settings file '{self.default_settings_file}' was not found. Cannot restore.")

        # Delete the current user settings file if it exists
        if os.path.exists(self.settings_file):
            try:
                os.remove(self.settings_file)
            except OSError as e:
                print(f"Error deleting user settings file: {e}")
                raise e # Re-raise the exception to be caught by the caller
            
        # Define default settings file
        self.settings_file = self.default_settings_file

        # Reload settings (this will now use the defaults) and re-save
        self.load_settings() # This will now load from default_settings.json

        # Refresh the main GUI and restart monitoring
        self.update_custom_buttons()

        # Save a new custom_settings.json from the loaded defaults
        self.settings_file = CUSTOM_SETTINGS_FILE
        self.save_settings()

        print("--- Default Settings Restored Successfully ---")

    def load_settings(self):
        '''Loads settings from the JSON file and updates the GUI variables accordingly.'''
        print("\n--- Loading Settings ---") 

        try:
            if os.path.exists(self.settings_file):
                print("Loading Settings from: {self.settings_file}")
                with open(self.settings_file, 'r') as f: settings = json.load(f)
                self.log_file_path = settings.get("log_file_path")

                # Load main button configs, merging with defaults to handle new settings

                self.load_event_codes() 

                loaded_main_configs = settings.get("main_button_configs", {})
                for btn_name, default_conf in self.main_button_configs.items():
                    # Update the default config with any saved values
                    default_conf.update(loaded_main_configs.get(btn_name, {}))
                
                self.txt_folder_path = settings.get("txt_folder_path")
                self.txt_folder_path_set2 = settings.get("txt_folder_path_set2")
                self.txt_folder_path_set3 = settings.get("txt_folder_path_set3")
                loaded_txt_config = settings.get("txt_field_columns_config")

                if loaded_txt_config:
                    self.txt_field_columns_config = loaded_txt_config
                else: # Fallback for old settings structure
                    old_txt_cols = settings.get("txt_field_columns", {"Event": "Event"})
                    old_txt_skips = settings.get("txt_field_skips", {})
                    # Reconstruct the ordered list from old dicts, prioritizing new fields
                    new_config = []
                    default_order_fields = ["Date", "Time", "KP", "DCC", "Line name", "Latitude", "Longitude", "Easting", "Northing", "Event", "Code"]
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


                self.folder_paths.clear()
                self.folder_paths.update(settings.get("folder_paths", {}))
                self.folder_columns.clear()
                self.folder_columns.update(settings.get("folder_columns", {}))
                self.folder_db_columns.clear() # --- NEW ---
                self.folder_db_columns.update(settings.get("folder_db_columns", {})) # --- NEW ---
                self.file_extensions.clear()
                self.file_extensions.update(settings.get("file_extensions", {}))
                self.folder_skips.clear()
                self.folder_skips.update(settings.get("folder_skips", {}))
                self.num_custom_buttons = settings.get("num_custom_buttons", 3)
                loaded_configs = settings.get("custom_button_configs", [])
                loaded_folders = settings.get("folder_paths", {})
                # If the loaded dictionary is empty, populate it with the defaults
                if not loaded_folders:
                    for folder in DEFAULT_MONITORED_FOLDERS:
                        loaded_folders[folder] = "" # Path will be empty by default
                self.folder_paths.update(loaded_folders)
                
                
                updated_custom_configs = []
                for i in range(self.num_custom_buttons):
                    if i < len(loaded_configs):
                        config = loaded_configs[i]
                    else:
                        config = {"text": f"Custom {i+1}", "event_text": f"Custom {i+1} Event"}
                    
                    # Ensure all required keys exist with defaults
                    config["txt_source_key"] = config.get("txt_source_key", "None")
                    config["tab_group"] = config.get("tab_group", "Main")
                    config["event_code"] = config.get("event_code", "") # Ensure event_code is loaded for each button
                    
                    updated_custom_configs.append(config)

                self.custom_button_configs = updated_custom_configs

                # Load custom button tab groups
                # Start with the fixed groups and add any others found in settings
                self.custom_button_tab_groups = sorted(list(set(["Main"] + settings.get("custom_button_tab_groups", []))))
                # Filter out empty string, if any might appear
                self.custom_button_tab_groups = [g for g in self.custom_button_tab_groups if g]


                loaded_colors = settings.get("button_colors", {})
                default_colors = {
                    "Log on": ("#90EE90", None),
                    "Log off": ("#FFB6C1", None),
                    "Event": ("#FFFFE0", None),
                    "SVP": ("#ADD8E6", None),
                    "New Day": ("#FFFF99", None),
                    "Hourly KP Log": ("#FFFF99", None)
                        }
                self.button_colors = default_colors.copy() # Start with defaults
                self.button_colors = default_colors
                for config in self.custom_button_configs:
                    btn_text = config.get("text")
                    if btn_text and btn_text not in self.button_colors:
                        self.button_colors[btn_text] = (None, None) # Default to no colors for new custom buttons

                for key, color_value in loaded_colors.items():
                    if isinstance(color_value, list) and len(color_value) == 2: # Check if it's the new tuple format (JSON loads lists)
                        self.button_colors[key] = (color_value[0], color_value[1])
                    elif isinstance(color_value, str): # Handle old format (only background hex string) for backward compatibility
                        self.button_colors[key] = (color_value, None) # Assume it's a background color, no font color
                    else:
                        self.button_colors[key] = (None, None) # Fallback for unknown formats
                self.sqlite_enabled = settings.get("sqlite_enabled", False)
                self.sqlite_db_path = settings.get("sqlite_db_path")
                always_on_top_setting = settings.get("always_on_top", False)
                self.always_on_top_var.set(always_on_top_setting)
                self.master.wm_attributes("-topmost", always_on_top_setting)
                self.sqlite_table = settings.get("sqlite_table", "fieldlog")
                self.always_on_top_var.set(settings.get("always_on_top", True))
                self.new_day_event_enabled_var.set(settings.get("new_day_event_enabled", True))
                self.hourly_event_enabled_var.set(settings.get("hourly_event_enabled", True))

                self.txt_source_aliases = settings.get("txt_source_aliases", {
                    "Main TXT": "Main TXT",
                    "TXT Source 2": "TXT Source 2",
                    "TXT Source 3": "TXT Source 3"
                })
                print("Settings loaded succesfully")
                self.update_status("Settings loaded.")
            else:
                self.update_status("Settings file not found. Using defaults.")
                print("Settings file not found, using defaults.") #DEBUG
        except json.JSONDecodeError as e:
            messagebox.showerror("Load Error", f"Settings file '{self.settings_file}' has invalid format:\n{e}\n\nUsing default settings.", parent=self.master)
            self.update_status("Error loading settings: Invalid format."); self.init_variables()
            print(f"JSON Decode Error: {e}") #DEBUG
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Load Error", f"Could not load settings from {self.settings_file}:\n{e}\n\nUsing default settings.", parent=self.master)
            self.update_status("Error loading settings."); self.init_variables()
            print(f"General Error loading settings: {e}") #DEBUG
        finally:
            if hasattr(self, 'button_frame') and self.button_frame: self.update_custom_buttons()
            if hasattr(self, 'db_status_label') and self.db_status_label: self.update_db_indicator()

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

        # Determine which settings file to load
        if not os.path.exists(self.settings_file):
            try:
                print(f"Custom settings not found. Loading from default file: {self.default_settings_file}")
                self.revert_to_defaults()
            except Exception as e:
                messagebox.showwarning("Error in the settings memory", "Paths for custom or default settings files not found", parent=self.master)


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

        # Check for one of the new main frames to ensure the UI has been initialized.
        if hasattr(self, 'custom_buttons_frame'):
            # Call create_main_buttons without arguments, as it now handles all frames.
            self.create_main_buttons()
            self.create_status_indicators()
            # FIX: Add this line to update the text of the newly created monitor status label
            self.update_monitor_indicator_text() 
            self.master.update_idletasks()

    def toggle_always_on_top(self):
        """Toggles the 'always on top' state of the main window based on the checkbox."""
        is_on_top = self.always_on_top_var.get()
        self.master.wm_attributes("-topmost", is_on_top)

    # --- Monitoring ---
    def start_monitoring(self):
        '''Function to read the last version of a file in several folders'''
        print("\n--- Starting Monitoring ---") #DEBUG
        
        active_monitors = list(self.monitors.items()) # Get a copy of the items

    # Step 1: Signal all threads to stop without blocking indefinitely
        for name, monitor_observer in active_monitors:
            try:
                if monitor_observer.is_alive():
                    monitor_observer.stop()
            except Exception as e:
                print(f"Error signalling monitor {name} to stop: {e}")

    # Step 2: Wait for all threads to terminate with a timeout
        for name, monitor_observer in active_monitors:
            try:
            # The join() call is implicitly part of stop(), but doing it
            # separately with a timeout prevents one stuck thread from
            # hanging the entire application.
                monitor_observer.join(timeout=1.0) 
            except Exception as e:
                print(f"Error joining monitor {name}: {e}")
            self.monitors.clear(); folder_cache.clear()
            print("Cleared existing monitors and folder cache.") #DEBUG

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
                #print(f"Attempting to start monitor for {folder_name}: {folder_path} (ext: {file_extension})") #DEBUG
                success = self.start_folder_monitoring(folder_name, folder_path, file_extension)
                if success: 
                    count += 1 
            
            elif self.folder_skips.get(folder_name): 
                print(f"Monitor for {folder_name} skipped by setting.")
            elif folder_path: 
                print(f"Monitor for {folder_name} not started: path invalid or not a directory ({folder_path}).")

        self.update_status(f"Monitoring {count} active folders.")

        # --- REFACTOR ---
        # Replace the old logic with a call to the new, reusable functions
        self.update_monitor_indicator_text()
        self.update_db_indicator()

    def start_folder_monitoring(self, folder_name, folder_path, file_extension):
        '''Start monitoring a specific folder for changes in files with a given extension.
            
            Arguments:
            * folder_name: Name of the folder to monitor.
            * folder_path: Full path to the folder to monitor.
            * file_extension: File extension to filter files (e.g., ".txt"). If empty, monitors all files.
            
            Returns True if monitoring started successfully, False otherwise.
        '''
        try: 
            if not os.path.isdir(folder_path):
                print(f"Error: Path '{folder_path}' is not a valid directory for monitoring '{folder_name}'.")
                return False
            os.listdir(folder_path) # Check permissions
        except Exception as e: 
            print(f"Error accessing directory '{folder_path}' for monitoring '{folder_name}': {e}")
            return False
        
        try:
            event_handler = FolderMonitor(folder_path, folder_name, self, file_extension)
            observer = PollingObserver(timeout=1)
            # --- THE FIX IS HERE: Set recursive=True ---
            observer.schedule(event_handler, folder_path, recursive=True)
            observer.start()
            self.monitors[folder_name] = observer
            # Trigger an immediate scan on a separate thread to populate cache right away
            threading.Thread(target=event_handler.update_latest_file, daemon=True).start()
            print(f"Successfully started recursive monitoring for {folder_name} at {folder_path} (ext: {file_extension}).")
            return True
        except Exception as e: 
            print(f"Failed to start watchdog monitor for {folder_name} at {folder_path}: {e}")
            return False

    # --- Programmed Events Scheduling ---
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
        now = datetime.datetime.now()
        next_hour = (now + datetime.timedelta(hours=1)).replace(minute=0, second=0, microsecond=0)
        #next_hour = (now + datetime.timedelta(minutes=1)).replace(second=0, microsecond=0) # Delta time modified to 1 minute for debugging
        time_until_next_hour_ms = int((next_hour - now).total_seconds() * 1000)

        # Add a small buffer (e.g., 1 second) to ensure it triggers after the hour
        trigger_delay_ms = time_until_next_hour_ms + 1000

        self._hourly_log_timer_id = self.master.after(trigger_delay_ms, self.trigger_hourly_log)
        print(f"Next 'Hourly KP Log' scheduled for {next_hour} (in {time_until_next_hour_ms/1000:.1f} seconds).")

    def trigger_hourly_log(self):
        """Triggers the hourly log and reschedules the next one."""

        if self.hourly_event_enabled_var.get():
            # Get column names from settings
            kp_col_name = self.txt_field_columns.get("KP")
            event_col_name = self.txt_field_columns.get("Event")

            if not kp_col_name or not event_col_name:
                print("Error: 'KP' column not configured in TXT Data Columns settings.")
                self.schedule_hourly_log()
                return
            
            # 1. Get current KP value
            current_kp = None
            try:
                txt_data = self._get_txt_data_from_source(self.txt_folder_path)
                current_kp_str = txt_data.get(kp_col_name)
                if current_kp_str is not None:
                    current_kp = float(current_kp_str)
            except (ValueError, TypeError, AttributeError) as e:
                print(f"Could not parse current KP value: {e}")

            if current_kp is None:
                print("Could not retrieve a valid current KP. Skipping hourly log.")
                self.schedule_hourly_log()
                return

            # 2. Find the last hourly KP log from the Excel file
            last_kp = None
            try:
                df = pd.read_excel(self.log_file_path)
                # Filter for previous hourly logs, ensuring the KP column is numeric
                hourly_logs_df = df[df[event_col_name].str.startswith("Current KP:", na=False)].copy()
                print(f"Found {len(hourly_logs_df)} previous hourly logs in Excel file.") #DEBUG
                hourly_logs_df[kp_col_name] = pd.to_numeric(hourly_logs_df[kp_col_name], errors='coerce')
                hourly_logs_df.dropna(subset=[kp_col_name], inplace=True)

                if not hourly_logs_df.empty:
                    last_kp = current_kp # Get the current KP value
            except Exception as e:
                print(f"Could not read or find last KP from Excel file: {e}")

            # 3. Format the event text string
            if last_kp is not None:
                progress = current_kp - last_kp
                event_text = f"Current KP: {current_kp:.3f} | Progress last hour: {progress:+.3f} km"
            else:
                event_text = f"Current KP: {current_kp:.3f} | First hourly log"

            # 4. Call the logging function with the generated text
            self._perform_log_action(event_type="Hourly KP Log",
                            event_text_for_excel=event_text,
                            triggering_button=None,  # No button is associated
                            txt_source_key="Main TXT") # Use the primary TXT source for KP data
        else:
            print("'Hourly KP Log' event is disabled, skipping log.")
        # Reschedule for the following hour
        self.schedule_hourly_log()

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
        current_bg_color, current_font_color = self.button_colors.get(button_name, (None, None))
        
        # --- Create StringVars ---
        event_text_var = tk.StringVar(value=current_event_text)
        event_code_var = tk.StringVar(value=current_event_code)
        button_bg_color_var = tk.StringVar(value=current_bg_color if current_bg_color else "")
        button_font_color_var = tk.StringVar(value=current_font_color if current_font_color else "")
        
        # --- UI Elements for the editor ---
        row_idx = 0
        
        # Event Text Entry
        ttk.Label(frame, text="Event Text:").grid(row=row_idx, column=0, sticky="w", pady=5, padx=5)
        event_text_entry = ttk.Entry(frame, textvariable=event_text_var, width=40)
        event_text_entry.grid(row=row_idx, column=1, sticky="ew", pady=5, padx=5)
        ToolTip(event_text_entry, "Text written to the 'Event' column in the log.")

        row_idx += 1
        # Event Code Combobox
        ttk.Label(frame, text="Event Code:").grid(row=row_idx, column=0, sticky="w", pady=5, padx=5)
        event_code_options = [""] + sorted(list(self.event_codes.keys()))
        event_code_combobox = ttk.Combobox(frame, textvariable=event_code_var,
                                            values=event_code_options, state="readonly", width=37)
        event_code_combobox.grid(row=row_idx, column=1, sticky="ew", pady=5, padx=5)
        ToolTip(event_code_combobox, "Select an event code to write to the 'Code' column when this button is pressed.")

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
            # Save the new event text and code
            self.main_button_configs[button_name]['event_text'] = event_text_var.get()
            self.main_button_configs[button_name]['event_code'] = event_code_var.get()

            # Save the new colors as a tuple
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
                "tab_group": "Main" # **MODIFIED:** Default to "Main" tab
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
        dialog_height = editor_window.winfo_reqheight() or 300 # Slightly increased height for new field

        center_x = main_x + (main_width // 2) - (dialog_width // 2)
        center_y = main_y + (main_height // 2) - (dialog_height // 2)
        editor_window.geometry(f"+{center_x}+{center_y}")

        frame = ttk.Frame(editor_window, padding="15")
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(1, weight=1) # Allow column 1 to expand for entry fields

        # Get current colors
        current_bg_color, current_font_color = self.button_colors.get(button_config.get("text"), (None, None))

        button_text_var = tk.StringVar(value=button_config.get("text", f"Custom {button_index+1}"))
        event_text_var = tk.StringVar(value=button_config.get("event_text", f"{button_config.get('text', f'Custom {button_index+1}')} Triggered"))
        tab_group_var = tk.StringVar(value=button_config.get("tab_group", "Main"))
        event_code_var = tk.StringVar(value=button_config.get("event_code", ""))
        
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
        event_code_options = [""] + sorted(list(self.event_codes.keys()))
        event_code_combobox = ttk.Combobox(frame, textvariable=event_code_var, 
                                             values=event_code_options, state="readonly", width=27)
        event_code_combobox.grid(row=row_idx, column=1, columnspan=2, sticky="ew", pady=2, padx=5)
        ToolTip(event_code_combobox, "Select an event code to write to the 'Code' column when this button is pressed.")

        row_idx += 1
        # --- Event Source Combobox (Now Dynamic) ---
        ttk.Label(frame, text="Event Source:").grid(row=row_idx, column=0, sticky="w", pady=2, padx=5)

        # 1. Get the aliases and build the lists for the dropdown
        aliases = self.txt_source_aliases
        internal_keys = TXT_FILES_KEYS
        
        # This list will be shown to the user in the dropdown
        display_names = ["None"] + [aliases.get(key, key) for key in internal_keys[1:]]

        # 2. Create translation maps to go between the display name and internal key
        display_to_internal_map = {display: internal for display, internal in zip(display_names, internal_keys)}
        internal_to_display_map = {internal: display for display, internal in zip(display_names, internal_keys)}

        # 3. Set the combobox's initial value based on the current configuration
        current_internal_key = button_config.get("txt_source_key", "None")
        txt_source_display_var = tk.StringVar(value=internal_to_display_map.get(current_internal_key, "None"))
        
        source_combobox = ttk.Combobox(frame, textvariable=txt_source_display_var,
                                        values=display_names, state="readonly", width=27)
        source_combobox.grid(row=row_idx, column=1, columnspan=2, sticky="ew", pady=2, padx=5)
        ToolTip(source_combobox, "Select which data source this button should use. Names are configured in Settings -> File Paths.")


        row_idx += 1
        # Tab Group selection
        ttk.Label(frame, text="Tab Group:").grid(row=row_idx, column=0, sticky="w", pady=2, padx=5)
        all_tab_groups = sorted(self.custom_button_tab_groups[:])
        tab_group_combobox = ttk.Combobox(frame, textvariable=tab_group_var,
                                            values=all_tab_groups, width=27) # Not readonly, allows user to type new group
        tab_group_combobox.grid(row=row_idx, column=1, columnspan=2, sticky="ew", pady=2, padx=5)
        ToolTip(tab_group_combobox, "Assign this button to a tab group. You can type a new group name or select an existing one.")

        row_idx += 1
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

        default_font_colors_for_picker = ["#000000", "#FFFFFF"] # Black, White
        for f_color in default_font_colors_for_picker:
            try:
                b = tk.Button(font_color_widget_frame, bg=f_color, width=1, height=1, relief="raised", bd=1,
                                fg='white' if f_color == '#000000' else 'black', # Make text visible on button
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
            button_config["event_code"] = event_code_var.get()
            button_config["tab_group"] = tab_group_var.get().strip() or "Main"

            # Translate the selected display name back to its internal key before saving
            selected_display_name = txt_source_display_var.get()
            button_config["txt_source_key"] = display_to_internal_map.get(selected_display_name, "None")
            

            new_bg_color_hex = button_bg_color_var.get() if button_bg_color_var.get() else None
            new_font_color_hex = button_font_color_var.get() if button_font_color_var.get() else None
            
            if old_button_text in self.button_colors and old_button_text != button_config["text"]:
                del self.button_colors[old_button_text]
            
            # Save the color as a tuple (background_color, font_color)
            self.button_colors[button_config["text"]] = (new_bg_color_hex, new_font_color_hex)

            # Tab Saving Logic 
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

# --- Settings Window Class ---
class SettingsWindow:

    # In class SettingsWindow...
    def __init__(self, master, parent_gui):
        self.master = master
        self.parent_gui = parent_gui
        self.master.title("Settings")
        self.master.geometry("1150x850") # --- MODIFIED: Wider for new column ---
        self.master.minsize(800, 500) # --- MODIFIED: Wider for new column ---
        self.style = parent_gui.style

        self.main_frame = ttk.Frame(self.master)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.main_frame.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)

        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        # Initialize selection tracking for TXT data columns
        self.selected_txt_row_index = -1  # -1 means no row is selected
        self.txt_move_up_btn = None
        self.txt_move_down_btn = None

        # --- Create tabs (ensure each is called only ONCE) ---
        self.create_file_paths_tab()
        self.create_txt_column_mapping_tab()
        self.create_button_configuration_tab()
        self.create_event_codes_tab()  # For the feature added previously
        self.create_monitored_folders_tab()
        self.create_sqlite_tab()
        self.create_auto_events_tab()

        # Bottom Buttons
        button_frame = ttk.Frame(self.main_frame)
        button_frame.grid(row=1, column=0, pady=(10, 0), sticky="e")
        ttk.Button(button_frame, text="Save and Close", command=self.save_and_close, style="Accent.TButton").pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.master.destroy).pack(side=tk.RIGHT)

    def save_and_close(self):
        self.save_settings()
        self.master.destroy()

    # --- Tab Creation Methods ---

    # Add this entire new method to the SettingsWindow class

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
        """Saves the current event codes from the parent GUI to the JSON file."""
        try:
            with open(self.parent_gui.event_codes_file, 'w') as f:
                json.dump(self.parent_gui.event_codes, f, indent=4)

            self.parent_gui.update_status("Event codes configuration saved.")
            # Also reload them in the parent GUI to ensure consistency
            self.parent_gui.load_event_codes()
        except Exception as e:
            messagebox.showerror("Save Error", f"Could not save event codes file:\n{e}", parent=self.master)

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
        # --- Main Container for TXT Sources ---
        txt_sources_container = ttk.Frame(tab)
        txt_sources_container.pack(fill='x', expand=True, anchor='n')
        txt_sources_container.columnconfigure(0, weight=1)

        # --- Helper function to create each TXT source entry ---
        def create_txt_source_frame(parent, title, name_entry_var, path_entry_var):
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
            return name_entry, path_entry

        # Create StringVars to hold the UI data
        self.txt_name_main_var = tk.StringVar()
        self.txt_path_main_var = tk.StringVar()
        self.txt_name_set2_var = tk.StringVar()
        self.txt_path_set2_var = tk.StringVar()
        self.txt_name_set3_var = tk.StringVar()
        self.txt_path_set3_var = tk.StringVar()

        # Create the three source blocks using the helper
        create_txt_source_frame(txt_sources_container, "Main Vehicle Navigation (Main TXT Data)", self.txt_name_main_var, self.txt_path_main_var)
        create_txt_source_frame(txt_sources_container, "Additional Vehicle Navigation Data (TXT Source 2)", self.txt_name_set2_var, self.txt_path_set2_var)
        create_txt_source_frame(txt_sources_container, "Additional Vehicle Navigation Data (TXT Source 3)", self.txt_name_set3_var, self.txt_path_set3_var)
   
        # Frame for restoring default settings ---
        restore_frame = ttk.LabelFrame(tab, text="Restore Default Settings", padding=15)
        restore_frame.pack(fill="x", pady=(20, 0), side="bottom") # Place it at the bottom
        restore_frame.columnconfigure(0, weight=1)

        restore_desc = ttk.Label(restore_frame, text="This will delete your current custom settings and restore the application's original defaults. This action cannot be undone.", wraplength=800)
        restore_desc.grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 10))

        style = ttk.Style()
        style.configure("Danger.TButton", foreground="white", background="red")
        style.map("Danger.TButton", background=[("active", "#cc0000")], foreground=[("active", "white")])

        restore_button = ttk.Button(restore_frame, text="Restore Default Settings", command=self.restore_default_settings, style="Danger.TButton")
        restore_button.grid(row=1, column=0, sticky='w')
        ToolTip(restore_button, "WARNING: Deletes 'custom_settings.json' and loads defaults from 'default_settings.json'.")

    def select_excel_file(self):
        initial_dir = os.path.dirname(self.log_file_entry.get()) if self.log_file_entry.get() else os.getcwd()
        file_path = filedialog.askopenfilename(initialdir=initial_dir, filetypes=[("Excel files", ["*.xlsx",".xlsb"])], parent=self.master, title="Select Excel Log File")
        if file_path: self.log_file_entry.delete(0, tk.END); self.log_file_entry.insert(0, file_path)

    def select_txt_folder(self, entry_widget):
        current_path = entry_widget.get()
        initial_dir = current_path if os.path.isdir(current_path) else os.path.dirname(current_path) if current_path else os.getcwd()
        folder_path = filedialog.askdirectory(initialdir=initial_dir, parent=self.master, title="Select Navigation TXT Folder")
        if folder_path: entry_widget.delete(0, tk.END); entry_widget.insert(0, folder_path)

    def restore_default_settings(self):
        """
        Handles the user confirmation and initiates the process of restoring default settings.
        """
        # Ask for user confirmation as this is a destructive action
        is_confirmed = messagebox.askyesno(
            "Confirm Restore Defaults",
            "Are you sure you want to restore all settings to their defaults?\n\n"
            "Your current 'custom_settings.json' file will be permanently deleted.",
            parent=self.master
        )

        if is_confirmed:
            try:
                # Call the main GUI's method to perform the core logic
                self.parent_gui.revert_to_defaults()

                # Refresh the settings window UI with the newly loaded default values
                self.load_settings()

                messagebox.showinfo(
                    "Success",
                    "Default settings have been restored.\n\n"
                    "Your custom settings file has been deleted. New settings will be saved to 'custom_settings.json'.",
                    parent=self.master
                )
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred while restoring defaults:\n{e}", parent=self.master)

    def create_txt_column_mapping_tab(self):
        tab = ttk.Frame(self.notebook, padding=20)
        self.notebook.add(tab, text="Data Columns")
        
        ttk.Label(tab, text="Map header found in TXT files to your desired Excel and Database column names. Check 'Skip' to ignore a field entirely. Click on a row to select it, then use the Move Up/Down buttons to reorder.", wraplength=900, justify=tk.LEFT).pack(pady=(0, 10), anchor='w')

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

    def preview_txt_data(self):
        """Finds the latest TXT file, reads the last line, and displays the parts in the preview column."""
        txt_folder = self.parent_gui.txt_folder_path # Use the main TXT folder for preview
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

    def clear_txt_preview(self):
        """Clears the text from all preview labels."""
        for row_widgets in self.txt_field_row_widgets:
            preview_label = row_widgets.get("preview_label")
            if preview_label:
                preview_label.config(text="")
        self.parent_gui.update_status("Preview cleared.")

    def add_txt_field_header(self, parent):
        """Adds a header row to the TXT field mapping section."""
        
        # Apply column configuration to the single, shared parent frame
        parent.grid_columnconfigure(0, weight=2, minsize=50) # TXT Field Name
        parent.grid_columnconfigure(1, weight=2, minsize=150) # TXT Field Name
        parent.grid_columnconfigure(2, weight=2, minsize=150) # Target Excel Column
        parent.grid_columnconfigure(3, weight=2, minsize=150) # Target DB Column
        parent.grid_columnconfigure(4, weight=2, minsize=150) # Preview Data
        parent.grid_columnconfigure(5, weight=0, minsize=50)  # Skip
        parent.grid_columnconfigure(6, weight=0, minsize=80)  # Actions

        header_frame = ttk.Frame(parent, style="Header.TFrame", padding=(5,3))
        header_frame.grid(row=0, column=0, columnspan=8, sticky="ew") # Span all columns

        # Place labels inside the header_frame, but they will align because the parent of header_frame has the config
        ttk.Label(header_frame, text="Order", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=6, sticky='w')
        ttk.Label(header_frame, text="TXT Column", font=("Arial", 10, "bold")).grid(row=0, column=1, padx=6, sticky='w')
        ttk.Label(header_frame, text="Preview TXT Data", font=("Arial", 10, "bold")).grid(row=0, column=2, padx=8, sticky='w')
        ttk.Label(header_frame, text="Excel Column / Cell", font=("Arial", 10, "bold")).grid(row=0, column=3, padx=6, sticky='w')
        ttk.Label(header_frame, text="SQLITE DB Column", font=("Arial", 10, "bold")).grid(row=0, column=4, padx=6, sticky='w')
        ttk.Label(header_frame, text="Skip?", font=("Arial", 10, "bold")).grid(row=0, column=5, padx=6, sticky='w')
        ttk.Label(header_frame, text="Actions", font=("Arial", 10, "bold")).grid(row=0, column=6, padx=6, sticky='w')

        # Also apply the same column configure to the header_frame itself so its internal labels space out correctly
        for i in range(6):
            header_frame.grid_columnconfigure(i, weight=parent.grid_columnconfigure(i).get('weight', 0))
            header_frame.grid_columnconfigure(i, minsize=parent.grid_columnconfigure(i).get('minsize', 0))

    def _select_txt_row(self, index):
        """Highlights the selected row by changing the background of all its widgets."""
        # First, deselect the previously selected row
        if self.selected_txt_row_index != -1 and self.selected_txt_row_index < len(self.txt_field_row_widgets):
            prev_row_info = self.txt_field_row_widgets[self.selected_txt_row_index]
            original_bg = "#f5f5f5" if self.selected_txt_row_index % 2 else "#ffffff"
            for widget in prev_row_info.get("all_widgets", []):
                try:
                    # For ttk widgets, changing style is preferred but complex. Direct config is simpler.
                    if isinstance(widget, (ttk.Entry, ttk.Label, ttk.Checkbutton)):
                        widget.configure(style=f"Row{self.selected_txt_row_index % 2}.T{type(widget).__name__}")
                    else: # Fallback for non-ttk or custom widgets
                        widget.configure(background=original_bg)
                except tk.TclError:
                    pass # Widget might be gone

        # Now, select the new row
        self.selected_txt_row_index = index
        if index != -1 and index < len(self.txt_field_row_widgets):
            current_row_info = self.txt_field_row_widgets[index]
            selection_color = "#ADD8E6" # Light blue
            for widget in current_row_info.get("all_widgets", []):
                try:
                    # For ttk widgets, it's best to configure a style for selection
                    # but for simplicity here we can try direct configuration
                    widget.configure(style=f"Selected.T{type(widget).__name__}")
                except tk.TclError:
                    pass # Widget might be gone or style doesn't apply

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
        
        # Add validation guard clauses at the beginning ---
        if current_index == -1:
            messagebox.showinfo("No Selection", "Please select a row to move.", parent=self.master)
            return

        total_items = len(self.parent_gui.txt_field_columns_config)

        # Now, perform the move only if the operation is valid
        if direction == "up" and current_index > 0:
            # Swap with the item above
            self.parent_gui.txt_field_columns_config[current_index], self.parent_gui.txt_field_columns_config[current_index - 1] = \
                self.parent_gui.txt_field_columns_config[current_index - 1], self.parent_gui.txt_field_columns_config[current_index]
            
            self.selected_txt_row_index -= 1
            self.recreate_txt_field_rows(reselect_index=self.selected_txt_row_index)
            self.parent_gui.update_status(f"Moved field up.")

        elif direction == "down" and current_index < total_items - 1:
            # Swap with the item below
            self.parent_gui.txt_field_columns_config[current_index], self.parent_gui.txt_field_columns_config[current_index + 1] = \
                self.parent_gui.txt_field_columns_config[current_index + 1], self.parent_gui.txt_field_columns_config[current_index]
            
            self.selected_txt_row_index += 1
            self.recreate_txt_field_rows(reselect_index=self.selected_txt_row_index)
            self.parent_gui.update_status(f"Moved field down.")


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
        """Removes a TXT field row by index."""
        # Before doing anything, check if the index is valid for the CURRENT list size.
        if not (0 <= index_to_remove < len(self.parent_gui.txt_field_columns_config)):
            print(f"Warning: remove_txt_field_row called with invalid index {index_to_remove}. Ignoring.")
            return

        # The rest of the function can now proceed safely
        if messagebox.askyesno("Confirm Deletion", f"Are you sure you want to remove field '{self.parent_gui.txt_field_columns_config[index_to_remove]['field']}'?", parent=self.master):
            del self.parent_gui.txt_field_columns_config[index_to_remove]
            if self.selected_txt_row_index == index_to_remove:
                self.selected_txt_row_index = -1
            elif self.selected_txt_row_index > index_to_remove:
                self.selected_txt_row_index -= 1

            self.recreate_txt_field_rows(reselect_index=self.selected_txt_row_index)
            self.parent_gui.update_status("TXT field removed.")


    def recreate_txt_field_rows(self, reselect_index=None):
        # Clear existing widgets except the header
        for widget in self.txt_fields_scrollable_frame.winfo_children():
            # The header is at row 0. We only want to destroy the data rows (row > 0).
            if int(widget.grid_info()["row"]) > 0:
                widget.destroy()
        self.txt_field_row_widgets.clear()

        # Define the set of default fields that should not be editable as 'TXT Field' or removable
        default_fixed_fields = DEFAULT_DATA_FIELDS #NEEDS TO BE REVIEWED

        # Recreate rows based on the current configuration
        for i, config in enumerate(self.parent_gui.txt_field_columns_config):
            # Each data row starts from grid row 1 (header is at row 0)
            grid_row_index = i + 1 

            # All widgets are placed directly into self.txt_fields_scrollable_frame ---
            parent_frame = self.txt_fields_scrollable_frame
            
            # Create a list to hold all widgets for this row for easy selection/styling
            widgets_in_row = []

            order_label = ttk.Label(parent_frame, text=str(i + 1), anchor='center')
            order_label.grid(row=grid_row_index, column=0, padx=5, pady=2, sticky='ew')
            widgets_in_row.append(order_label)

            # TXT Field Label/Entry
            current_field_entry_widget = None
            if config["field"] in default_fixed_fields:
                field_widget = ttk.Label(parent_frame, text=f"{config['field']}:", anchor='w')
                field_widget.grid(row=grid_row_index, column=1, padx=5, pady=2, sticky='ew')
            else:
                field_widget = ttk.Entry(parent_frame)
                field_widget.insert(0, config["field"])
                field_widget.grid(row=grid_row_index, column=1, padx=5, pady=2, sticky='ew')
                ToolTip(field_widget, "Enter the exact name of the field as it appears in the TXT file.")
                current_field_entry_widget = field_widget
            widgets_in_row.append(field_widget)

            # Preview Data Label
            preview_label = ttk.Label(parent_frame, text="", anchor='w', foreground="blue")
            preview_label.grid(row=grid_row_index, column=2, padx=5, pady=2, sticky='ew')
            widgets_in_row.append(preview_label)

            # Target Excel Column Name
            column_entry = ttk.Entry(parent_frame)
            column_entry.insert(0, config.get("column_name", config["field"]))
            column_entry.grid(row=grid_row_index, column=3, padx=5, pady=2, sticky="ew")
            ToolTip(column_entry, "Enter the column header for the Excel Log, OR a static cell reference using the format: ='SheetName'!A1")
            widgets_in_row.append(column_entry)

            # Target DB Column Name
            db_column_entry = ttk.Entry(parent_frame)
            db_column_entry.insert(0, config.get("db_column_name", ""))
            db_column_entry.grid(row=grid_row_index, column=4, padx=5, pady=2, sticky="ew")
            ToolTip(db_column_entry, "Enter the target column name for the SQLite Database.")
            widgets_in_row.append(db_column_entry)
            
            # Skip Checkbox
            skip_var = tk.BooleanVar(value=config.get("skip", False))
            skip_checkbox = ttk.Checkbutton(parent_frame, variable=skip_var)
            skip_checkbox.grid(row=grid_row_index, column=5, padx=(15,5), pady=2, sticky='w')
            widgets_in_row.append(skip_checkbox)

            # Remove Button
            remove_btn = ttk.Button(parent_frame, text="Remove", width=8, style="Toolbutton",
                                     command=lambda idx=i: self.remove_txt_field_row(idx))
            if config["field"] in default_fixed_fields:
                remove_btn.config(state=tk.DISABLED)
            remove_btn.grid(row=grid_row_index, column=6, padx=5, pady=2, sticky='w')
            widgets_in_row.append(remove_btn)

            # Bind click event to all widgets in the row for selection
            click_handler = lambda e, idx=i: self._select_txt_row(idx)
            for widget in widgets_in_row:
                widget.bind("<Button-1>", click_handler)

            # Store references
            self.txt_field_row_widgets.append({
                "field_entry_widget": current_field_entry_widget,
                "column_entry": column_entry,
                "db_column_entry": db_column_entry,
                "skip_var": skip_var,
                "preview_label": preview_label,
                "all_widgets": widgets_in_row # Store list of all widgets in the row
            })
        
        # After recreating, re-apply selection highlighting if needed
        if reselect_index is not None:
             self._select_txt_row(reselect_index)
        else:
            self._select_txt_row(-1) # Deselect all

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
        header_frame.pack(anchor='w', pady=(15,5))
        
        # Configure columns for the header frame to match the rows
        header_frame.grid_columnconfigure(0, weight=0, minsize=40) # Button #
        header_frame.grid_columnconfigure(1, weight=1, minsize=135) # Button Text
        header_frame.grid_columnconfigure(2, weight=2, minsize=200) # Event Text (longer)
        header_frame.grid_columnconfigure(3, weight=0, minsize=80) # Event Code
        header_frame.grid_columnconfigure(4, weight=0, minsize=80) # Event Source
        header_frame.grid_columnconfigure(5, weight=0, minsize=80) # Tab Group

        ttk.Label(header_frame, text="Button #", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=(5,0), sticky='w')
        ttk.Label(header_frame, text="Button Text", font=("Arial", 10, "bold")).grid(row=0, column=1, padx=5, sticky='ew')
        ttk.Label(header_frame, text="Event Text (for Log)", font=("Arial", 10, "bold")).grid(row=0, column=2, padx=5, sticky='ew')
        ttk.Label(header_frame, text="Event Code", font=("Arial", 10, "bold")).grid(row=0, column=3, padx=5, sticky='w')
        ttk.Label(header_frame, text="Event Source", font=("Arial", 10, "bold")).grid(row=0, column=4, padx=5, sticky='w')
        ttk.Label(header_frame, text="Tab Group", font=("Arial", 10, "bold")).grid(row=0, column=5, padx=5, sticky='w')

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
            initial_tab_group = config.get("tab_group", "Main") # **MODIFIED:** Default to "Main"

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
        
        ttk.Label(tab, text="Configure additional folders to monitor for their latest file names. The latest file name will be logged in the specified Excel/DB column.", wraplength=900, justify=tk.LEFT).pack(pady=(0, 10), anchor='w')

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
        self.folder_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        self.folder_canvas.bind_all("<Button-4>", _on_mousewheel)
        self.folder_canvas.bind_all("<Button-5>", _on_mousewheel)
        
        self.folder_entries = {}
        self.folder_column_entries = {}
        self.folder_db_column_entries = {} # --- NEW ---
        self.file_extension_entries = {}
        self.folder_skip_vars = {}
        self.folder_row_widgets = {}
        self.add_folder_header(self.scrollable_frame)

    def add_folder_header(self, parent):
        # Configure the grid columns directly on the main scrollable frame
        parent.columnconfigure(0, weight=2, minsize=140)  # Folder Type
        parent.columnconfigure(1, weight=4, minsize=250)  # Monitor Path
        parent.columnconfigure(2, weight=0)               # ... button
        parent.columnconfigure(3, weight=2, minsize=150)  # Excel Column
        parent.columnconfigure(4, weight=2, minsize=150)  # SQLite DB Column
        parent.columnconfigure(5, weight=1, minsize=80)   # File Ext.
        parent.columnconfigure(6, weight=0, minsize=50)   # Skip?

        # Header labels placed directly into the parent grid for perfect alignment
        ttk.Label(parent, text="Folder Type", style="Header.TLabel", padding=5).grid(row=0, column=0, sticky='w')
        ttk.Label(parent, text="Monitor Path", style="Header.TLabel", padding=5).grid(row=0, column=1, sticky='w')
        # Empty label for browse button column to maintain spacing
        ttk.Label(parent, text="", style="Header.TLabel").grid(row=0, column=2)
        ttk.Label(parent, text="Excel Column", style="Header.TLabel", padding=5).grid(row=0, column=3, sticky='w')
        ttk.Label(parent, text="SQLite DB Column", style="Header.TLabel", padding=5).grid(row=0, column=4, sticky='w')
        ttk.Label(parent, text="File Ext.", style="Header.TLabel", padding=5).grid(row=0, column=5, sticky='w')
        ttk.Label(parent, text="Skip?", style="Header.TLabel", padding=5).grid(row=0, column=6, sticky='w')

    def add_initial_folder_rows(self):
        default_folders = DEFAULT_MONITORED_FOLDERS
        
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
            db_column_name_to_use = self.parent_gui.folder_db_columns.get(folder_name, "")
            extension_to_use = self.parent_gui.file_extensions.get(folder_name, "")

            if folder_name in ["Main TXT File", "TXT Source 2", "TXT Source 3"]:
                if not column_name_to_use or column_name_to_use == folder_name:
                    column_name_to_use = folder_name.replace(" ", "_")
                if not extension_to_use:
                    extension_to_use = "txt"

            self.add_folder_row(folder_name=folder_name, folder_path=folder_path_to_use,
                                 column_name=column_name_to_use,
                                 db_column_name=db_column_name_to_use,
                                 extension=extension_to_use,
                                 skip=self.parent_gui.folder_skips.get(folder_name, False))
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
            
            # **FIX:** If the list exists, iterate through it and destroy each widget
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

    def add_folder_row(self, folder_name="", folder_path="", column_name="", db_column_name="", extension="", skip=False):
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
        db_column_entry = ttk.Entry(parent)
        db_column_entry.insert(0, db_column_name) 
        extension_entry = ttk.Entry(parent, width=10)
        extension_entry.insert(0, extension)
        skip_var = tk.BooleanVar(value=skip)
        skip_checkbox = ttk.Checkbutton(parent, variable=skip_var)

        # --- Place Widgets on the Shared Grid ---
        label.grid(row=row_index, column=0, padx=5, pady=2, sticky="ew")
        entry.grid(row=row_index, column=1, padx=5, pady=2, sticky="ew")
        button.grid(row=row_index, column=2, padx=(0,5), pady=2, sticky='w')
        column_entry.grid(row=row_index, column=3, padx=5, pady=2, sticky="ew")
        db_column_entry.grid(row=row_index, column=4, padx=5, pady=2, sticky="ew")
        extension_entry.grid(row=row_index, column=5, padx=5, pady=2, sticky="ew")
        skip_checkbox.grid(row=row_index, column=6, padx=(15, 5), pady=2, sticky='w')

        # --- Selection and Tooltip Logic ---
        widgets_in_row = [label, entry, button, column_entry, db_column_entry, extension_entry, skip_checkbox]
        click_handler = lambda e, name=folder_name: self._select_folder_row(name)
        for widget in widgets_in_row:
            widget.bind("<Button-1>", click_handler)
        ToolTip(entry, f"Enter the full path to the '{folder_name}' data folder.")
        ToolTip(button, "Browse for the folder.")
        ToolTip(column_entry, f"Enter the Excel/DB column name for the latest '{folder_name}' filename.")
        ToolTip(extension_entry, "Optional: Monitor only files with this extension (e.g., 'svp').")
        ToolTip(skip_checkbox, f"Check to disable monitoring for the '{folder_name}' folder.")
        ToolTip(db_column_entry, f"Enter the SQLite DB column name for the latest '{folder_name}' filename.")

        # Store references for selection, saving, and removal
        self.folder_entries[folder_name] = entry
        self.folder_column_entries[folder_name] = column_entry
        self.folder_db_column_entries[folder_name] = db_column_entry
        self.file_extension_entries[folder_name] = extension_entry
        self.folder_skip_vars[folder_name] = skip_var
        # Store all widgets in the row for highlighting
        self.folder_row_widgets[folder_name] = widgets_in_row


    def update_scroll_region(self):
        self.scrollable_frame.update_idletasks()
        self.folder_canvas.configure(scrollregion=self.folder_canvas.bbox("all"))



    def create_sqlite_tab(self):
        tab = ttk.Frame(self.notebook, padding=20); self.notebook.add(tab, text="SQLite Log")
        
        ttk.Label(tab, text="Configure settings for logging data to a SQLite database file.", wraplength=900, justify=tk.LEFT).pack(pady=(0, 10), anchor='w')

        enable_frame = ttk.Frame(tab); enable_frame.pack(fill='x', pady=(0, 15))
        self.sqlite_enabled_var = tk.BooleanVar()
        enable_check = ttk.Checkbutton(enable_frame, text="Enable SQLite Database Logging", variable=self.sqlite_enabled_var, style="Large.TCheckbutton")
        enable_check.pack(side=tk.LEFT, pady=(5, 10)); ToolTip(enable_check, "Check to enable logging events to an SQLite database file.")
        
        config_frame = ttk.LabelFrame(tab, text="SQLite Configuration", padding=15)
        config_frame.pack(fill='x'); config_frame.columnconfigure(1, weight=1)

        # Frame for action buttons
        action_button_frame = ttk.Frame(config_frame)
        action_button_frame.grid(row=3, column=1, columnspan=2, sticky="w", pady=15)
        
        ttk.Label(config_frame, text="Database File (.db):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.sqlite_db_path_entry = ttk.Entry(config_frame, width=70)
        self.sqlite_db_path_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew"); ToolTip(self.sqlite_db_path_entry, "Full path to the SQLite database file")
        db_browse_btn = ttk.Button(config_frame, text="Browse", command=self.select_sqlite_file)
        db_browse_btn.grid(row=0, column=2, padx=5, pady=5); ToolTip(db_browse_btn, "Browse for an existing SQLite file")
        
        ttk.Label(config_frame, text="Table Name:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.sqlite_table_entry = ttk.Entry(config_frame, width=40)
        self.sqlite_table_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w"); ToolTip(self.sqlite_table_entry, "The name of the table within the database where logs will be written (e.g., 'fieldlog'). This table must exist or be created by you.")
        
        test_button = ttk.Button(config_frame, text="Test Connection & Table", command=self.test_sqlite_connection)
        test_button.grid(row=2, column=1, padx=5, pady=15, sticky="w"); ToolTip(test_button, "Verify connection to the database file and check if the specified table exists.")

        generate_sql_button = ttk.Button(action_button_frame, text="Generate Table SQL", command=self.generate_create_table_sql)
        generate_sql_button.pack(side=tk.LEFT, padx=5)
        ToolTip(generate_sql_button, "Beta for testing")
        
        self.test_result_label = ttk.Label(config_frame, text="", font=("Arial", 9), wraplength=500)
        self.test_result_label.grid(row=3, column=0, columnspan=3, padx=5, pady=2, sticky="w")

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
            db_column_name = row_info["db_column_entry"].get().strip()
            skip_value = row_info["skip_var"].get()

            if not field_name and not (field_name in DEFAULT_DATA_FIELDS):
                field_name = f"Custom_Field_{i+1}"

            new_txt_field_configs.append({
                "field": field_name,
                "column_name": column_name if column_name else field_name,
                "db_column_name": db_column_name,
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

    def generate_create_table_sql(self):
        """
        Generates a CREATE TABLE SQL statement based on the "TXT Column" field
        and displays it in a new window for the user to copy.
        """
        # Ensure we have the latest column names from the UI before proceeding
        self.save_settings_to_parent_vars()

        table_name = self.sqlite_table_entry.get().strip()
        if not table_name:
            messagebox.showerror("Missing Table Name", "Please enter a table name before generating SQL.", parent=self.master)
            return

        column_defs = []
        processed_cols = set()

        # 1. Add the primary key column
        pk_col = "original_excel_row"
        column_defs.append(f"    [{pk_col}] INTEGER PRIMARY KEY")
        processed_cols.add(pk_col.lower())

        # 2. Iterate through the configured columns from the "Data Columns" tab
        for config_item in self.parent_gui.txt_field_columns_config:
            # CHANGE: Use the 'field' value (from the "TXT Column") as the source
            source_col_name = config_item.get("field", "").strip()
            
            if source_col_name:
                # CHANGE: Sanitize the name for use in SQL
                db_col = self._sanitize_for_sql(source_col_name)
                
                # Add the column if it has a valid sanitized name and we haven't already added it
                if db_col and db_col.lower() not in processed_cols:
                    column_defs.append(f"    [{db_col}] TEXT")
                    processed_cols.add(db_col.lower())
        
        # 3. Add any other special columns that are handled programmatically
        special_cols = ["time_fix"]
        for col in special_cols:
             if col.lower() not in processed_cols:
                # Sanitize these as well for consistency
                sanitized_col = self._sanitize_for_sql(col)
                if sanitized_col:
                    column_defs.append(f"    [{sanitized_col}] TEXT")
                    processed_cols.add(sanitized_col.lower())

        # 4. Construct the final SQL statement
        column_sql = ",\n".join(column_defs)
        full_sql = f"CREATE TABLE IF NOT EXISTS [{table_name}] (\n{column_sql}\n);"

        # 5. Display the SQL in a new pop-up window
        sql_window = Toplevel(self.master)
        sql_window.title("Generated SQL for Table Creation")
        sql_window.transient(self.master)
        sql_window.grab_set()
        sql_window.geometry("600x400")

        main_frame = ttk.Frame(sql_window, padding=10)
        main_frame.pack(fill="both", expand=True)
        main_frame.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)

        sql_text = tk.Text(main_frame, wrap="word", font=("Courier New", 10), height=10, width=70)
        sql_text.insert("1.0", full_sql)
        sql_text.config(state="disabled") # Make it read-only
        sql_text.grid(row=0, column=0, columnspan=2, sticky="nsew")

        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=sql_text.yview)
        scrollbar.grid(row=0, column=2, sticky="ns")
        sql_text['yscrollcommand'] = scrollbar.set

        def copy_to_clipboard():
            self.master.clipboard_clear()
            self.master.clipboard_append(full_sql)
            self.parent_gui.update_status("SQL copied to clipboard.")

        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=1, column=0, columnspan=3, pady=(10, 0), sticky="e")

        copy_btn = ttk.Button(button_frame, text="Copy to Clipboard", command=copy_to_clipboard)
        copy_btn.pack(side="left", padx=5)

        close_btn = ttk.Button(button_frame, text="Close", command=sql_window.destroy)
        close_btn.pack(side="left")

    def select_sqlite_file(self):
        filetypes = [("SQLite Database", "*.db"), ("SQLite Database", "*.sqlite"), ("SQLite3 Database", "*.sqlite3"), ("All Files", "*.*")]
        current_path = self.sqlite_db_path_entry.get()
        initial_dir = os.path.dirname(current_path) if current_path else os.getcwd()
        filepath = filedialog.askopenfilename(parent=self.master, title="Select SQLite Database File", initialdir=initial_dir, filetypes=filetypes, defaultextension="*.sqlite")
        if filepath: 
            self.sqlite_db_path_entry.delete(0, tk.END)
            self.sqlite_db_path_entry.insert(0, filepath)
        if hasattr(self, 'test_result_label'): 
            self.test_result_label.config(text="")

    def test_sqlite_connection(self):
        db_path = self.sqlite_db_path_entry.get().strip()
        table_name = self.sqlite_table_entry.get().strip() or DEFAULT_TABLE_NAME
        if not db_path: 
            self.test_result_label.config(text="❌ Error: Database path is empty.", foreground="red")
            return
        conn = None
        result_text = ""
        result_color = "red"
        try:
            conn = sqlite3.connect(db_path, timeout=3)
            cursor = conn.cursor()
            result_text = f"✔️ Connection to '{os.path.basename(db_path)}' successful.\n"
            try:
                cursor.execute(f"SELECT 1 FROM [{table_name}] LIMIT 1;")
                result_text += f"✔️ Table '{table_name}' found."
                result_color = "green"
            except sqlite3.OperationalError as e_table:
                if "no such table" in str(e_table).lower(): 
                    result_text += f"⚠️ Warning: Table '{table_name}' not found. It needs to be created."
                    result_color = "#E67E00"
                else: 
                    raise e_table
            except Exception as e: result_text += f"❌ Error checking table: {e}"; result_color = "red"
        except sqlite3.Error as e: result_text = f"❌ Error connecting/checking DB: {e}"; result_color = "red"
        except Exception as e: result_text = f"❌ Unexpected Error: {e}"; result_color = "red"
        finally:
            if conn: conn.close()
            self.test_result_label.config(text=result_text, foreground=result_color)
            self.master.after(15000, lambda: self.test_result_label.config(text=""))

    def create_auto_events_tab(self):
            """Creates the tab for configuring automatic timed events."""
            tab = ttk.Frame(self.notebook, padding=20)
            self.notebook.add(tab, text="Programmed Events")

            # --- "New Day" Event Configuration ---
            new_day_frame = ttk.LabelFrame(tab, text="Midnight 'New Day' Event", padding=15)
            new_day_frame.pack(fill='x', pady=(0, 15))
            new_day_frame.columnconfigure(1, weight=1)

            new_day_check = ttk.Checkbutton(
                new_day_frame,
                text="Enable this automatic event",
                variable=self.parent_gui.new_day_event_enabled_var,
                style="Large.TCheckbutton"
            )
            new_day_check.grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 10))
            ToolTip(new_day_check, "If checked, an event will be logged automatically at midnight.")

            # Color picker for New Day event
            self.new_day_bg_color_var, self.new_day_bg_color_label, \
            self.new_day_font_color_var, self.new_day_font_color_label = self._create_color_picker_row(
                new_day_frame, 1, "Excel Row Colors:", "New Day" # Changed label text
        )

            # --- "Hourly KP Log" Event Configuration ---
            hourly_frame = ttk.LabelFrame(tab, text="Hourly KP Log Event", padding=15)
            hourly_frame.pack(fill='x', pady=5)
            hourly_frame.columnconfigure(1, weight=1)

            hourly_check = ttk.Checkbutton(
                hourly_frame,
                text="Enable this automatic event",
                variable=self.parent_gui.hourly_event_enabled_var,
                style="Large.TCheckbutton"
            )
            hourly_check.grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 10))
            ToolTip(hourly_check, "If checked, the current KP will be logged automatically every hour.")

            # Color picker for Hourly event
            self.hourly_bg_color_var, self.hourly_bg_color_label, \
            self.hourly_font_color_var, self.hourly_font_color_label = self._create_color_picker_row(
                hourly_frame, 1, "Excel Row Colors:", "Hourly KP Log" # Changed label text
        )

    def _create_color_picker_row(self, parent_frame, row, label_text, event_name):
        """Helper to create color picker widgets for both background and font colors for the Auto Events tab."""
        ttk.Label(parent_frame, text=label_text).grid(row=row, column=0, sticky='w', padx=5)

        # Frame for Background Color picker
        bg_color_frame = ttk.Frame(parent_frame)
        bg_color_frame.grid(row=row, column=1, sticky='w', padx=5, pady=(2,0)) # Add some top padding

        # Get initial colors from the master button_colors dictionary
        initial_bg_color, initial_font_color = self.parent_gui.button_colors.get(event_name, (None, None))
        
        bg_color_var = tk.StringVar(value=initial_bg_color if initial_bg_color else "")
        bg_display_label = tk.Label(bg_color_frame, width=4, relief="solid", borderwidth=1,
                                    background=bg_color_var.get() if bg_color_var.get() else 'SystemButtonFace')
        bg_display_label.pack(side="left", padx=(0, 5))

        # Background Clear button
        clear_bg_btn = ttk.Button(bg_color_frame, text="X", width=2, style="Toolbutton",
                                 command=lambda: self.parent_gui._set_color_on_widget(bg_color_var, bg_display_label, None, self.master))
        clear_bg_btn.pack(side="left", padx=1)
        ToolTip(clear_bg_btn, f"Clear background color for {event_name}.")

        # Background Choose button
        choose_bg_btn = ttk.Button(bg_color_frame, text="...", width=3, style="Toolbutton",
                                  command=lambda: self.parent_gui._choose_color_dialog(bg_color_var, bg_display_label, self.master, f"{event_name} Background"))
        choose_bg_btn.pack(side="left", padx=1)
        ToolTip(choose_bg_btn, f"Choose a custom background color for {event_name}.")


        # Frame for Font Color picker
        font_color_frame = ttk.Frame(parent_frame)
        font_color_frame.grid(row=row + 1, column=1, sticky='w', padx=5, pady=(0,2)) # Place below background, with some bottom padding

        font_color_var = tk.StringVar(value=initial_font_color if initial_font_color else "")
        font_display_label = tk.Label(font_color_frame, width=4, relief="solid", borderwidth=1,
                                      background=font_color_var.get() if font_color_var.get() else 'SystemButtonFace')
        font_display_label.pack(side="left", padx=(0, 5))

        # Font Clear button
        clear_font_btn = ttk.Button(font_color_frame, text="X", width=2, style="Toolbutton",
                                   command=lambda: self.parent_gui._set_color_on_widget(font_color_var, font_display_label, None, self.master))
        clear_font_btn.pack(side="left", padx=1)
        ToolTip(clear_font_btn, f"Clear font color for {event_name}.")

        # Font Choose button
        choose_font_btn = ttk.Button(font_color_frame, text="...", width=3, style="Toolbutton",
                                    command=lambda: self.parent_gui._choose_color_dialog(font_color_var, font_display_label, self.master, f"{event_name} Font"))
        choose_font_btn.pack(side="left", padx=1)
        ToolTip(choose_font_btn, f"Choose a custom font color for {event_name}.")
        
        # Adjust row span for the main label
        parent_frame.grid_rowconfigure(row, weight=0) # Make sure the label row doesn't expand
        parent_frame.grid_rowconfigure(row+1, weight=0) # Make sure the font color row doesn't expand


        return bg_color_var, bg_display_label, font_color_var, font_display_label # Return all four  

    # --- Settings Save/Load Logic ---
    def save_settings(self):
        # --- File Paths Tab ---
        self.parent_gui.log_file_path = self.log_file_entry.get().strip()
        self.parent_gui.txt_source_aliases["Main TXT"] = self.txt_name_main_var.get().strip()
        self.parent_gui.txt_folder_path = self.txt_path_main_var.get().strip()
        self.parent_gui.txt_source_aliases["TXT Source 2"] = self.txt_name_set2_var.get().strip()
        self.parent_gui.txt_folder_path_set2 = self.txt_path_set2_var.get().strip()
        self.parent_gui.txt_source_aliases["TXT Source 3"] = self.txt_name_set3_var.get().strip()
        self.parent_gui.txt_folder_path_set3 = self.txt_path_set3_var.get().strip()

        # --- Data Columns Tab ---
        new_txt_field_configs = []
        for i, row_info in enumerate(self.txt_field_row_widgets):
            field_name = ""
            if row_info["field_entry_widget"]:
                field_name = row_info["field_entry_widget"].get().strip()
            else: # For fixed fields, get the name from the original config
                if i < len(self.parent_gui.txt_field_columns_config):
                    field_name = self.parent_gui.txt_field_columns_config[i]["field"]
            
            column_name = row_info["column_entry"].get().strip()
            db_column_name = row_info["db_column_entry"].get().strip()
            skip_value = row_info["skip_var"].get()

            if not field_name and not (field_name in DEFAULT_DATA_FIELDS):
                field_name = f"Custom_Field_{i+1}"

            new_txt_field_configs.append({
                "field": field_name,
                "column_name": column_name if column_name else field_name,
                "db_column_name": db_column_name,
                "skip": skip_value
            })
        self.parent_gui.txt_field_columns_config = new_txt_field_configs
        self.parent_gui.txt_field_columns = {cfg["field"]: cfg["column_name"] for cfg in self.parent_gui.txt_field_columns_config}
        self.parent_gui.txt_field_skips = {cfg["field"]: cfg["skip"] for cfg in self.parent_gui.txt_field_columns_config}

        # --- Monitored Folders Tab ---
        parent_folder_paths = {}
        parent_folder_cols = {}
        parent_folder_db_cols = {} # --- NEW ---
        parent_folder_exts = {}
        parent_folder_skips = {}
        for folder_name in self.folder_entries.keys():
            entry_widget = self.folder_entries[folder_name]
            col_entry = self.folder_column_entries[folder_name]
            db_col_entry = self.folder_db_column_entries[folder_name] # --- NEW ---
            ext_entry = self.file_extension_entries[folder_name]
            skip_var = self.folder_skip_vars[folder_name]

            folder_path = entry_widget.get().strip()
            parent_folder_paths[folder_name] = folder_path
            parent_folder_cols[folder_name] = col_entry.get().strip() if col_entry.get().strip() else folder_name
            parent_folder_db_cols[folder_name] = db_col_entry.get().strip() # --- NEW ---
            parent_folder_exts[folder_name] = ext_entry.get().strip().lstrip('.')
            parent_folder_skips[folder_name] = skip_var.get()
        
        self.parent_gui.folder_paths = parent_folder_paths
        self.parent_gui.folder_columns = parent_folder_cols
        self.parent_gui.folder_db_columns = parent_folder_db_cols # --- NEW ---
        self.parent_gui.file_extensions = parent_folder_exts
        self.parent_gui.folder_skips = parent_folder_skips

        # --- Button Configuration Tab ---
        parent_custom_configs = []
        all_new_tab_groups = set()
        for i, (text_widget, event_widget, event_code_var, txt_source_var, tab_group_var) in enumerate(self.custom_button_widgets):
            text = text_widget.get().strip()
            event_text = event_widget.get().strip()
            event_code = event_code_var.get()
            txt_source_key = txt_source_var.get()
            tab_group = tab_group_var.get().strip() or "Main"

            default_text = f"Custom {i + 1}"
            final_text = text if text else default_text
            final_event_text = event_text if event_text else f"{final_text} Triggered"

            parent_custom_configs.append({
                "text": final_text,
                "event_text": final_event_text,
                "event_code": event_code,
                "txt_source_key": txt_source_key,
                "tab_group": tab_group
            })
            all_new_tab_groups.add(tab_group)

        self.parent_gui.num_custom_buttons = len(parent_custom_configs)
        self.parent_gui.custom_button_configs = parent_custom_configs
        final_tab_groups = set(self.parent_gui.custom_button_tab_groups)
        final_tab_groups.update(all_new_tab_groups)
        self.parent_gui.custom_button_tab_groups = sorted(list(final_tab_groups))
        
        # --- Programmed Events / Timezone Tab ---
        # Save colors
        new_day_bg_color_hex = self.new_day_bg_color_var.get()
        new_day_font_color_hex = self.new_day_font_color_var.get()
        self.parent_gui.button_colors["New Day"] = (new_day_bg_color_hex if new_day_bg_color_hex else None, new_day_font_color_hex if new_day_font_color_hex else None)

        hourly_bg_color_hex = self.hourly_bg_color_var.get()
        hourly_font_color_hex = self.hourly_font_color_var.get()
        self.parent_gui.button_colors["Hourly KP Log"] = (hourly_bg_color_hex if hourly_bg_color_hex else None, hourly_font_color_hex if hourly_font_color_hex else None)

        
        
        # --- SQLite Log Tab ---
        self.parent_gui.sqlite_enabled = self.sqlite_enabled_var.get()
        self.parent_gui.sqlite_db_path = self.sqlite_db_path_entry.get().strip()
        self.parent_gui.sqlite_table = self.sqlite_table_entry.get().strip() or DEFAULT_TABLE_NAME

        # --- Final Actions ---
        # Trigger the master save function in the main GUI
        self.parent_gui.save_settings()
        # Refresh the main GUI and services with the new settings
        self.parent_gui.update_custom_buttons()
        self.parent_gui.start_monitoring()
        self.parent_gui.update_db_indicator()

    def load_settings(self):
        self.log_file_entry.delete(0, tk.END)
        self.populate_event_codes_tree()
        self.log_file_entry.insert(0, self.parent_gui.log_file_path or "")
        
        # Load the custom names into the new entry fields
        aliases = self.parent_gui.txt_source_aliases
        self.txt_name_main_var.set(aliases.get("Main TXT", "Main TXT"))
        self.txt_name_set2_var.set(aliases.get("TXT Source 2", "TXT Source 2"))
        self.txt_name_set3_var.set(aliases.get("TXT Source 3", "TXT Source 3"))

        # Load the paths into the path entry fields
        self.txt_path_main_var.set(self.parent_gui.txt_folder_path or "")
        self.txt_path_set2_var.set(self.parent_gui.txt_folder_path_set2 or "")
        self.txt_path_set3_var.set(self.parent_gui.txt_folder_path_set3 or "")

        # Reload TXT field rows based on the (potentially newly loaded) config
        self.recreate_txt_field_rows()
        self.master.after_idle(lambda: self.txt_fields_canvas.config(scrollregion=self.txt_fields_canvas.bbox("all")))


        for name, frame in list(self.folder_row_widgets.items()):
            if frame and frame.winfo_exists(): frame.destroy()
        self.folder_row_widgets.clear()
        self.folder_entries.clear()
        self.folder_column_entries.clear()
        self.file_extension_entries.clear()
        self.folder_skip_vars.clear()

        self.add_initial_folder_rows()
        self.master.after_idle(self.update_scroll_region)

        self.num_buttons_entry.delete(0, tk.END)
        self.num_buttons_entry.insert(0, str(self.parent_gui.num_custom_buttons))
        self.recreate_custom_button_settings()
        
        # Load colors for the automatic events tab
        new_day_bg_color, new_day_font_color = self.parent_gui.button_colors.get("New Day", (None, None))
        self.parent_gui._set_color_on_widget(self.new_day_bg_color_var, self.new_day_bg_color_label, new_day_bg_color, self.master)
        self.parent_gui._set_color_on_widget(self.new_day_font_color_var, self.new_day_font_color_label, new_day_font_color, self.master)

        hourly_bg_color, hourly_font_color = self.parent_gui.button_colors.get("Hourly KP Log", (None, None))
        self.parent_gui._set_color_on_widget(self.hourly_bg_color_var, self.hourly_bg_color_label, hourly_bg_color, self.master)
        self.parent_gui._set_color_on_widget(self.hourly_font_color_var, self.hourly_font_color_label, hourly_font_color, self.master)
        

        self.sqlite_enabled_var.set(self.parent_gui.sqlite_enabled)
        self.sqlite_db_path_entry.delete(0, tk.END)
        self.sqlite_db_path_entry.insert(0, self.parent_gui.sqlite_db_path or "")
        self.sqlite_table_entry.delete(0, tk.END)
        self.sqlite_table_entry.insert(0, self.parent_gui.sqlite_table or DEFAULT_TABLE_NAME)
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