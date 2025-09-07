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


if sys.platform == "win32":
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

#DEBUG
timings = {}
start_time = time.perf_counter()

# --- DEFINED CONSTANTS ---
# PATHS
DEFAULT_SETTINGS_FILE = "settings/default_settings.json"
CUSTOM_SETTINGS_FILE = "settings/custom_settings.json"
EVENT_CODES_FILE = "settings/event_codes.json"

# DICCTIONARY KEYS #NEEDS TO BE REVIEWED
EXCEL_LOG_REQUIRED_COLS = {'runline', 'kp', 'event'} 
DEFAULT_DATA_FIELDS = {"Date-Time", "KP", "DCC", "Line name", "Latitude", "Longitude", "Easting", "Northing", "Event", "Code", "KP Ref."} 
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
        #self.update_latest_file() # Initial scan

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

        self.init_styles()
        self.init_variables()
        self.static_field_configs = []
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

        self.txt_file_path = None # This will now be dynamic based on source

        # NEW: For data parsed directly from the TXT file's columns
        self.txt_mapping_config = [
            {"field": "KP", "column_name": "KP", "skip": False},
            {"field": "DCC", "column_name": "DCC", "skip": False},
            {"field": "Line name", "column_name": "Line name", "skip": False},
            {"field": "Latitude", "column_name": "Latitude", "skip": False},
            {"field": "Longitude", "column_name": "Longitude", "skip": False},
            {"field": "Easting", "column_name": "Easting", "skip": False},
            {"field": "Northing", "column_name": "Northing", "skip": False},
        ]

        # NEW: For data generated by the application itself
        self.generated_fields_config = [
            {"field": "Date-Time", "column_name": "UTC Date-Time", "skip": False, "source": "PC Time (UTC)"},
            {"field": "Local Time", "column_name": "Local Time", "skip": False, "source": "PC Time + Offset"},
            {"field": "Event", "column_name": "Event", "skip": False, "source": "Button"},
            {"field": "Code", "column_name": "Code", "skip": False, "source": "Button"},
            {"field": "KP Ref.", "column_name": "KP Ref.", "skip": False, "source": "Source Alias"}
        ]

        # For data from static cells in Excel
        self.static_field_configs = []
        
        # These dictionaries will be derived from the three config lists above for easier lookup
        self.txt_field_columns = {}
        self.txt_field_skips = {}


        self.folder_paths = {}
        self.folder_columns = {}
        self.file_extensions = {}
        self.folder_skips = {}
        self.folder_log_x_instead = {}
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

        self.always_on_top_var = tk.BooleanVar(value=False)
        self.settings_window_instance = None # Track settings window
        self.custom_inline_editor_window = None # To track the open inline editor
        self.is_monitoring = False 
        self.monitoring_button = None 

        self.status_var = tk.StringVar()
        self.monitor_status_label = None
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
        general_lf.rowconfigure((0, 1, 2), weight=1)

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
        config_lf.grid(row=0, column=0, sticky="nsew")
        self.config_frame.columnconfigure(0, weight=1)

        # CHANGE 1: Configure a 2x2 grid layout with equal weighting.
        config_lf.columnconfigure((0, 1), weight=1)
        config_lf.rowconfigure((0, 1), weight=1)

        self.monitoring_button = ttk.Button(config_lf, text="Start Monitoring", style="Small.TButton", command=self.toggle_monitoring)
        # CHANGE 2: Place the monitoring button on the top row, spanning both columns.
        self.monitoring_button.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=4, pady=(4, 2))
        ToolTip(self.monitoring_button, "Start or stop monitoring all configured folders for file changes.")
        self.update_monitoring_button_ui() # Set initial button text and style

        btn_settings = ttk.Button(config_lf, text="Settings", style="Small.TButton", command=self.open_settings)
        # CHANGE 3: Place the settings button on the bottom row, first column.
        btn_settings.grid(row=1, column=0, columnspan=2, sticky="nsew", padx=4, pady=(2, 4))
        ToolTip(btn_settings, "Open the configuration window.")

    # PASTE THESE TWO METHODS INTO THE DataLoggerGUI CLASS

    def preview_data_file(self):
        """Finds the latest TXT or NPD file and displays the data in the settings window preview."""
        if not self.settings_gui_instance:
            self.update_status("Settings window is not open.")
            return

        # Assuming self.txt_folder_path stores the path to the folder containing the data files.
        # It might be beneficial to rename this attribute to something like self.data_folder_path.
        data_folder = self.txt_folder_path
        if not data_folder or not os.path.isdir(data_folder):
            messagebox.showerror("Path Error", "The 'Main Navigation Data Folder' is not set.", parent=self.settings_window_instance)
            return

        # Find the latest file of each type
        latest_txt = self.find_latest_file_in_folder(data_folder, ".txt")
        latest_npd = self.find_latest_file_in_folder(data_folder, ".npd")

        latest_file = None
        # Determine which file is the most recent
        if latest_txt and latest_npd:
            latest_file = latest_txt if os.path.getmtime(latest_txt) > os.path.getmtime(latest_npd) else latest_npd
        else:
            # This will select whichever file exists, or remain None if neither exists
            latest_file = latest_txt or latest_npd

        if not latest_file:
            messagebox.showinfo("File Not Found", f"No .txt or .npd files were found in:\n{data_folder}", parent=self.settings_window_instance)
            return

        try:
            with open(latest_file, "r", encoding='utf-8', errors='ignore') as f:
                lines = f.readlines()
            
            if not lines:
                messagebox.showinfo("File Empty", "The latest file is empty.", parent=self.settings_window_instance)
                return

            data_parts = lines[-1].strip().split(',')
            
            # Use the correct reference to the SettingsWindow instance's widgets
            for i, row_widgets in enumerate(self.settings_gui_instance.txt_field_row_widgets):
                preview_label = row_widgets.get("preview_label")
                if preview_label:
                    preview_label.config(text=data_parts[i].strip() if i < len(data_parts) else "<no data>")
            
            self.update_status(f"Preview loaded from {os.path.basename(latest_file)}")

        except Exception as e:
            messagebox.showerror("Read Error", f"An error occurred while reading the file:\n{e}", parent=self.settings_window_instance)

    def clear_data_preview(self):
        """Clears the text from all preview labels in the settings window."""
        if self.settings_gui_instance:
            # Use the correct reference to the SettingsWindow instance's widgets
            for row_widgets in self.settings_gui_instance.txt_field_row_widgets:
                preview_label = row_widgets.get("preview_label")
                if preview_label:
                    preview_label.config(text="")
        self.update_status("Preview cleared.")


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

        # vvvvvv ADD THIS vvvvvv
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

        # Get the configuration for this specific main button
        config = self.main_button_configs.get(event_type, {})

        # Get the event text from the configuration, with a fallback
        event_text_for_excel = config.get("event_text", f"Default {event_type}")

        # START NEW CODE: Get the source key from the configuration
        source_key_for_log = config.get("txt_source_key", "Main TXT")
        # END NEW CODE
        
        skip_files = (event_type == "Event") # Still skip files only for the main "Event" button
            
        self._perform_log_action(event_type=event_type,
                                 event_text_for_excel=event_text_for_excel,
                                 triggering_button=button_widget,
                                 txt_source_key=source_key_for_log) # Use the new, configurable source key

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
                                 triggering_button=button_widget,
                                 txt_source_key=txt_source_key) 

    def _perform_log_action(self, event_type, event_text_for_excel, triggering_button, txt_source_key):
        """Initiates a logging action on a background thread to prevent GUI freezing."""
        # Check for invalid TXT source early on the main thread
        if txt_source_key == "None":
            # Allow events with no source to proceed but they might be missing data
            pass

        original_text = None
        if triggering_button and isinstance(triggering_button, ttk.Button) and triggering_button.winfo_exists():
            original_text = triggering_button['text']
            triggering_button.config(state=tk.DISABLED, text="Working...")
        
        self.update_status(f"Processing '{event_type}'...")

        def _log_worker():
            """The function that runs on the background thread."""
            try:
                # Start with an empty row
                row_data = {}

                # --- DATA GATHERING ---

                # 1. Get data parsed from the TXT file if a source is specified
                if txt_source_key != "None":
                    source_folder_path = None
                    if txt_source_key == "Main TXT": source_folder_path = self.txt_folder_path
                    elif txt_source_key == "TXT Source 2": source_folder_path = self.txt_folder_path_set2
                    elif txt_source_key == "TXT Source 3": source_folder_path = self.txt_folder_path_set3
                    elif txt_source_key == "TXT Source 4": source_folder_path = self.txt_folder_path_set4
                    elif txt_source_key == "TXT Source 5": source_folder_path = self.txt_folder_path_set5

                    if source_folder_path and os.path.isdir(source_folder_path):
                        txt_data = self._get_txt_data_from_source(source_folder_path)
                        if txt_data:
                            row_data.update(txt_data)

                # 2. Get static data from Excel cells
                static_data_from_cells = self._get_static_excel_data()
                if static_data_from_cells:
                    row_data.update(static_data_from_cells)

                # 3. Get latest files from monitored folders
                latest_files_data = self.get_latest_files_data_fast()
                if latest_files_data:
                    row_data.update(latest_files_data)

                # --- PROCESSING AND GENERATED FIELDS ---

                final_event_text = event_text_for_excel # Start with default event text

                # NEW: Add back the "Log on" data capture logic
                if event_type == "Log on":
                    kp_col_name = self.txt_field_columns.get("KP")
                    if kp_col_name and kp_col_name in row_data:
                        try:
                            self.last_log_on_kp = float(row_data[kp_col_name])
                            self.log_on_time = datetime.datetime.now()
                            self.update_status(f"KP for Log on event stored: {self.last_log_on_kp}")
                        except (ValueError, TypeError):
                            self.last_log_on_kp = None
                            self.log_on_time = None
                            self.update_status("Could not parse KP for Log on. Calculations disabled.")

                # Check for "Log off" and perform calculation
                elif event_type == "Log off" and self.calculate_logoff_values.get():
                    kp_col_name = self.txt_field_columns.get("KP")
                    if self.last_log_on_kp is not None and self.log_on_time is not None and kp_col_name in row_data:
                        try:
                            current_kp = float(row_data[kp_col_name])
                            current_time = datetime.datetime.now()
                            time_diff_seconds = (current_time - self.log_on_time).total_seconds()
                            distance_km = abs(current_kp - self.last_log_on_kp)
                            speed_knots = 0
                            if time_diff_seconds > 0:
                                distance_nm = distance_km / 1.852
                                time_hours = time_diff_seconds / 3600
                                speed_knots = distance_nm / time_hours
                            
                            final_event_text = f"Log off - Distance travelled: {distance_km:.2f}km - Speed: {speed_knots:.2f} Knots"
                            self.last_log_on_kp = None # Reset after calculation
                            self.log_on_time = None
                        except (ValueError, TypeError):
                            pass # Keep default text if calculation fails
                
                # Add all other generated fields
                utc_now = datetime.datetime.now(datetime.UTC)
                offset_delta = datetime.timedelta(hours=self.time_offset_hours.get())
                local_time = utc_now + offset_delta

                def get_gen_col_name(field_name):
                    return self.txt_field_columns.get(field_name)

                dt_col = get_gen_col_name("Date-Time")
                if dt_col: row_data[dt_col] = utc_now.strftime("%Y-%m-%d %H:%M:%S")
                lt_col = get_gen_col_name("Local Time")
                if lt_col: row_data[lt_col] = local_time.strftime("%Y-%m-%d %H:%M:%S")
                event_col = get_gen_col_name("Event")
                if event_col: row_data[event_col] = final_event_text
                
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
                
                # --- FINAL ACTIONS ---
                color_tuple = self.button_colors.get(event_type, (None, None))
                row_color = color_tuple[0] if isinstance(color_tuple, tuple) and len(color_tuple) > 0 else None
                font_color = color_tuple[1] if isinstance(color_tuple, tuple) and len(color_tuple) > 1 else None

                excel_success, _, excel_message = self.save_to_excel_and_sqlite(row_data, row_color, font_color)
                
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

        # Start the background thread
        log_thread = threading.Thread(target=_log_worker, daemon=True)
        log_thread.start()

    def _re_enable_button_and_update_status(self, button, original_text, status_message):
        """A helper function to run on the main thread after a background task completes."""
        if button and button.winfo_exists():
            button.config(state=tk.NORMAL)
            if original_text: button.config(text=original_text)
        self.update_status(status_message)

    # --- TXT Reading and Writting ---
    def _get_txt_data_from_source(self, folder_path):
        """
        Reads and parses data from the latest TXT file based on txt_mapping_config.
        Returns a dictionary of the parsed data only.
        """
        parsed_data = {}
        latest_txt_file_path = None
        if folder_path and os.path.exists(folder_path):
            latest_txt_file_path = self.find_latest_file_in_folder(folder_path, ".txt")
            if not latest_txt_file_path:
                latest_txt_file_path = self.find_latest_file_in_folder(folder_path, ".npd") # Add this line
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
                for i, field_config in enumerate(self.txt_mapping_config):
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
        It now checks if the file is 'active' based on its modification time.
        """
        data = {}
        # Get the threshold value once, outside the loop
        threshold = self.active_logging_threshold_seconds.get()
        current_time = time.time()

        for folder_name, column_name in self.folder_columns.items():
            if not self.folder_skips.get(folder_name, False) and column_name:
                latest_file_path = folder_cache.get(folder_name)

                if latest_file_path and os.path.exists(latest_file_path):
                    # A file exists. Check if it's recent enough to be "active".
                    file_mtime = os.path.getmtime(latest_file_path)
                    if (current_time - file_mtime) <= threshold:
                        # ACTIVE: The file is recent. Log "X" or the filename.
                        if self.folder_log_x_instead.get(folder_name, False):
                            data[column_name] = "X"
                        else:
                            filename_without_ext, _ = os.path.splitext(os.path.basename(latest_file_path))
                            data[column_name] = filename_without_ext
                    else:
                        # INACTIVE: The file is old, so log an empty string.
                        data[column_name] = ""
                else:
                    # NOT FOUND: No file has been found yet for this folder.
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

    def save_to_excel_and_sqlite(self, row_data, row_color=None, font_color=None):
        """
        Saves a single row of data to the open Excel file via xlwings.
        """
        if not self.log_file_path or not os.path.exists(self.log_file_path):
            return False, False, "Excel: Path Invalid."

        excel_message = "Excel: Fail."
        success_excel = False

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
            wb.save()
            
            excel_message = "Excel: OK."
            success_excel = True
            
        except Exception as e:
            traceback.print_exc()
            excel_message = f"Excel: Fail ({type(e).__name__})."
            return False, False, f"{excel_message}"

        return success_excel, True, excel_message # NOTE: returns true for SQL part to avoid errors, but it won't be used


    # --- Settings Saving and Loading ---
    def save_settings(self):
        '''Saves the current settings to the JSON file.'''
        print("\n--- Saving Settings ---")
        colors_to_save = {}
        for key, (bg_color, font_color) in self.button_colors.items():
            if bg_color or font_color:
                colors_to_save[key] = (bg_color, font_color)
        settings = {
            "log_file_path": self.log_file_path,
            "time_offset_hours": self.time_offset_hours.get(),
            "txt_folder_path": self.txt_folder_path,
            "txt_folder_path_set2": self.txt_folder_path_set2,
            "txt_folder_path_set3": self.txt_folder_path_set3,
            "txt_folder_path_set4": self.txt_folder_path_set4,
            "txt_folder_path_set5": self.txt_folder_path_set5,
            # NEW: Save the three separate config lists
            "txt_mapping_config": self.txt_mapping_config,
            "generated_fields_config": self.generated_fields_config,
            "static_field_configs": self.static_field_configs,
            "folder_paths": self.folder_paths,
            "folder_columns": self.folder_columns,
            "file_extensions": self.file_extensions,
            "folder_skips": self.folder_skips,
            "folder_log_x_instead": self.folder_log_x_instead,
            "num_custom_buttons": self.num_custom_buttons,
            "custom_button_configs": self.custom_button_configs,
            "custom_button_tab_groups": self.custom_button_tab_groups,
            "button_colors": colors_to_save,
            "always_on_top": self.always_on_top_var.get(),
            "active_logging_threshold_seconds": self.active_logging_threshold_seconds.get(),
            "new_day_event_enabled": self.new_day_event_enabled_var.get(),
            "hourly_event_enabled": self.hourly_event_enabled_var.get(),
            "main_button_configs": self.main_button_configs,
            "txt_source_aliases": self.txt_source_aliases,
            "calculate_logoff_values": self.calculate_logoff_values.get()
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
                print(f"Loading Settings from: {self.settings_file}")
                with open(self.settings_file, 'r') as f:
                    settings = json.load(f)
                
                # --- Load Main Settings ---
                self.log_file_path = settings.get("log_file_path")
                self.time_offset_hours.set(settings.get("time_offset_hours", 0.0))
                self.txt_folder_path = settings.get("txt_folder_path")
                self.txt_folder_path_set2 = settings.get("txt_folder_path_set2")
                self.txt_folder_path_set3 = settings.get("txt_folder_path_set3")
                self.txt_folder_path_set4 = settings.get("txt_folder_path_set4")
                self.txt_folder_path_set5 = settings.get("txt_folder_path_set5")

                # --- Load the three separate config lists ---
                self.txt_mapping_config = settings.get("txt_mapping_config", self.txt_mapping_config)
                self.generated_fields_config = settings.get("generated_fields_config", self.generated_fields_config)
                self.static_field_configs = settings.get("static_field_configs", [])

                # --- Handle backwards compatibility for old settings files ---
                if "txt_field_columns_config" in settings and not settings.get("txt_mapping_config"):
                    all_configs = settings["txt_field_columns_config"]
                    print("Old settings format detected. Migrating to new format...")
                    generated_fields_set = {"Date-Time", "Local Time", "Event", "Code", "KP Ref."}
                    self.generated_fields_config = [c for c in all_configs if c.get("field") in generated_fields_set]
                    self.txt_mapping_config = [c for c in all_configs if c.get("field") not in generated_fields_set and not str(c.get("column_name", "")).startswith('=')]
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
                combined_configs = self.txt_mapping_config + self.generated_fields_config + self.static_field_configs
                self.txt_field_columns = {cfg["field"]: cfg["column_name"] for cfg in combined_configs}
                self.txt_field_skips = {cfg["field"]: cfg.get("skip", False) for cfg in combined_configs}

                # --- Load Remaining Settings (Folder, Button, etc.) ---
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
                self.txt_source_aliases = settings.get("txt_source_aliases", self.txt_source_aliases)

                print("Settings loaded successfully")
                self.update_status("Settings loaded.")
            else:
                self.update_status("Settings file not found. Using defaults.")
                print("Settings file not found, using defaults.")
                # When no file is found, derive the lookup dictionaries from defaults
                combined_configs = self.txt_mapping_config + self.generated_fields_config + self.static_field_configs
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
            # FIX: Add this line to update the text of the newly created monitor status label
            self.update_monitor_indicator_text() 
            self.master.update_idletasks()

    def toggle_always_on_top(self):
        """Toggles the 'always on top' state of the main window based on the checkbox."""
        is_on_top = self.always_on_top_var.get()
        self.master.wm_attributes("-topmost", is_on_top)

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
                return None # <<< CHANGE THIS
            os.listdir(folder_path) # Check permissions
        except Exception as e: 
            print(f"Error accessing directory '{folder_path}' for monitoring '{folder_name}': {e}")
            return None # <<< CHANGE THIS
        
        try:
            event_handler = FolderMonitor(folder_path, folder_name, self, file_extension)
            observer = PollingObserver(timeout=1)
            observer.schedule(event_handler, folder_path, recursive=True)
            observer.start()
            self.monitors[folder_name] = observer
            print(f"Successfully started recursive monitoring for {folder_name} at {folder_path} (ext: {file_extension}).")
            return event_handler # <<< CHANGE THIS to return the handler object
        except Exception as e: 
            print(f"Failed to start watchdog monitor for {folder_name} at {folder_path}: {e}")
            return None # <<< CHANGE THIS

    def stop_monitoring(self):
        """Public method to stop all monitoring."""
        if hasattr(self, 'progress_bar') and self.progress_bar.winfo_ismapped():
             self.hide_progress_bar() # <<< ADD THIS
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
        # Get the current source key for the button, defaulting to "Main TXT"
        current_source_key = button_config.get("txt_source_key", "Main TXT")
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
        
       
        # Create a list of "Code - Description" strings for the dropdown
        event_code_display_list = [""] # Start with a blank option
        for code, desc in sorted(self.event_codes.items()):
            event_code_display_list.append(f"{code} - {desc}")
        
        event_code_combobox = ttk.Combobox(frame, textvariable=event_code_display_var, # Use the new display variable
                                           values=event_code_display_list,             # Use the new display list
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
        ToolTip(source_combobox, "Select which data source this button should use. Names are configured in Settings -> File Paths.")
        
        
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
            internal_to_display_map = {internal: display for display, internal in zip(display_names, internal_keys)}
            selected_display_name = source_combobox.get()  # Use source_combobox.get()
            selected_source_key = next((key for key, value in internal_to_display_map.items() if value == selected_display_name), "None")
            self.main_button_configs[button_name]['txt_source_key'] = selected_source_key
            

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
        
        
        current_event_code = button_config.get("event_code", "")
        # Find the full "Code - Description" string for the current code
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
        
        
        # Create a list of "Code - Description" strings
        event_code_display_list = [""]
        for code, desc in sorted(self.event_codes.items()):
            event_code_display_list.append(f"{code} - {desc}")
            
        event_code_combobox = ttk.Combobox(frame, textvariable=event_code_display_var, # Use display var
                                           values=event_code_display_list,             # Use display list
                                           state="readonly", width=27)
        
        
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
            
            # Get the full "Code - Description" string and parse it
            selected_display_string = event_code_display_var.get()
            code_to_save = ""
            if " - " in selected_display_string:
                code_to_save = selected_display_string.split(" - ", 1)[0]
            button_config["event_code"] = code_to_save
            
            
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
        self.create_txt_mapping_tab()
        self.create_generated_fields_tab()
        self.create_static_fields_tab()
        self.create_button_configuration_tab()
        self.create_event_codes_tab()
        self.create_monitored_folders_tab()
        self.create_auto_events_tab()
        self.create_timezone_tab()

        # --- Bottom Buttons (remain in the main_frame) ---
        button_frame = ttk.Frame(self.main_frame)
        # Span both columns (canvas and scrollbar)
        button_frame.grid(row=1, column=0, columnspan=2, pady=(10, 0), sticky="e")
        ttk.Button(button_frame, text="Save and Close", command=self.save_and_close, style="Accent.TButton").pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.master.destroy).pack(side=tk.RIGHT)

    

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
            "skip": False  # <-- ADD THIS LINE
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
        self.txt_name_set4_var = tk.StringVar()
        self.txt_path_set4_var = tk.StringVar()
        self.txt_name_set5_var = tk.StringVar()
        self.txt_path_set5_var = tk.StringVar()

        # Create the three source blocks using the helper
        create_txt_source_frame(txt_sources_container, "Main Vehicle Navigation (Main TXT Data)", self.txt_name_main_var, self.txt_path_main_var)
        create_txt_source_frame(txt_sources_container, "Additional Vehicle Navigation Data (TXT Source 2)", self.txt_name_set2_var, self.txt_path_set2_var)
        create_txt_source_frame(txt_sources_container, "Additional Vehicle Navigation Data (TXT Source 3)", self.txt_name_set3_var, self.txt_path_set3_var)
        create_txt_source_frame(txt_sources_container, "Additional Vehicle Navigation Data (TXT Source 4)", self.txt_name_set4_var, self.txt_path_set4_var)
        create_txt_source_frame(txt_sources_container, "Additional Vehicle Navigation Data (TXT Source 5)", self.txt_name_set5_var, self.txt_path_set5_var)
   
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

    def create_txt_mapping_tab(self):
        tab = ttk.Frame(self.notebook, padding=20)
        self.notebook.add(tab, text="TXT File Mapping")
        
        ttk.Label(tab, text="Define the structure of your navigation text file. The order must match the order of columns in the file. Use the preview button to verify.", wraplength=900, justify=tk.LEFT).pack(pady=(0, 10), anchor='w')

        controls_frame = ttk.Frame(tab)
        controls_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Button(controls_frame, text="Preview Latest Data", command=self.parent_gui.preview_data_file).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(controls_frame, text="Clear Preview", command=self.parent_gui.clear_data_preview).pack(side=tk.LEFT, padx=(0, 20))

        spacer = ttk.Frame(controls_frame)
        spacer.pack(side=tk.LEFT, expand=True, fill='x')

        self.txt_move_up_btn = ttk.Button(controls_frame, text="Move Up", command=lambda: self.move_selected_txt_field("up"))
        self.txt_move_up_btn.pack(side=tk.RIGHT, padx=5)

        self.txt_move_down_btn = ttk.Button(controls_frame, text="Move Down", command=lambda: self.move_selected_txt_field("down"))
        self.txt_move_down_btn.pack(side=tk.RIGHT, padx=5)

        ttk.Button(controls_frame, text="Add New Field", command=self.add_txt_field_row).pack(side=tk.RIGHT, padx=5)
        
        self.txt_fields_canvas = tk.Canvas(tab, borderwidth=0, background="#ffffff")
        txt_scrollbar = ttk.Scrollbar(tab, orient="vertical", command=self.txt_fields_canvas.yview)
        self.txt_fields_scrollable_frame = ttk.Frame(self.txt_fields_canvas)
        self.txt_fields_scrollable_frame.bind("<Configure>", lambda e: self.txt_fields_canvas.configure(scrollregion=self.txt_fields_canvas.bbox("all")))
        self.txt_fields_canvas.create_window((0, 0), window=self.txt_fields_scrollable_frame, anchor="nw")
        self.txt_fields_canvas.configure(yscrollcommand=txt_scrollbar.set)
        self.txt_fields_canvas.pack(side="left", fill="both", expand=True)
        txt_scrollbar.pack(side="right", fill="y")
        
        self.txt_field_row_widgets = []
        self.add_txt_field_header(self.txt_fields_scrollable_frame)
        # The load_settings method will call recreate_txt_mapping_rows
        
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

    def recreate_txt_field_rows(self, reselect_index=None):
        # Clear existing widgets except the header
        for widget in self.txt_fields_scrollable_frame.winfo_children():
            if int(widget.grid_info()["row"]) > 0:
                widget.destroy()
        self.txt_field_row_widgets.clear()

        # CORRECTED: Use the new txt_mapping_config attribute
        for i, config in enumerate(self.parent_gui.txt_mapping_config):
            grid_row_index = i + 1
            parent_frame = self.txt_fields_scrollable_frame
            widgets_in_row = []

            # Order Label
            order_label = ttk.Label(parent_frame, text=str(i + 1), anchor='center')
            order_label.grid(row=grid_row_index, column=0, padx=5, pady=2, sticky='ew')
            widgets_in_row.append(order_label)

            # TXT Field Entry
            field_widget = ttk.Entry(parent_frame)
            field_widget.insert(0, config["field"])
            field_widget.grid(row=grid_row_index, column=1, padx=5, pady=2, sticky='ew')
            widgets_in_row.append(field_widget)

            # Preview Data Label
            preview_label = ttk.Label(parent_frame, text="", anchor='w', foreground="blue")
            preview_label.grid(row=grid_row_index, column=2, padx=5, pady=2, sticky='ew')
            widgets_in_row.append(preview_label)
            
            # Excel Column Entry
            column_entry = ttk.Entry(parent_frame)
            column_entry.insert(0, config.get("column_name", config["field"]))
            column_entry.grid(row=grid_row_index, column=3, padx=5, pady=2, sticky="ew")
            widgets_in_row.append(column_entry)
            
            # Skip Checkbox
            skip_var = tk.BooleanVar(value=config.get("skip", False))
            skip_checkbox = ttk.Checkbutton(parent_frame, variable=skip_var)
            skip_checkbox.grid(row=grid_row_index, column=5, padx=(15,5), pady=2, sticky='w')
            widgets_in_row.append(skip_checkbox)

            # Remove Button
            remove_btn = ttk.Button(parent_frame, text="Remove", width=8, style="Toolbutton",
                                      command=lambda idx=i: self.remove_txt_field_row(idx))
            remove_btn.grid(row=grid_row_index, column=6, padx=5, pady=2, sticky='w')
            widgets_in_row.append(remove_btn)

            click_handler = lambda e, idx=i: self._select_txt_row(idx)
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
            self._select_txt_row(reselect_index)
        else:
            self._select_txt_row(-1)

        self.master.after_idle(lambda: self.txt_fields_canvas.config(scrollregion=self.txt_fields_canvas.bbox("all")))

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
        # CORRECTED: Use the new txt_mapping_config attribute
        can_move_up = hasattr(self, 'selected_txt_row_index') and self.selected_txt_row_index > 0
        can_move_down = hasattr(self, 'selected_txt_row_index') and self.selected_txt_row_index != -1 and self.selected_txt_row_index < len(self.parent_gui.txt_mapping_config) - 1

        if self.txt_move_up_btn:
            self.txt_move_up_btn.config(state=tk.NORMAL if can_move_up else tk.DISABLED)
        if self.txt_move_down_btn:
            self.txt_move_down_btn.config(state=tk.NORMAL if can_move_down else tk.DISABLED)

    def move_selected_txt_field(self, direction):
        current_index = self.selected_txt_row_index
        if current_index == -1: return

        # CORRECTED: Use the new txt_mapping_config attribute
        config_list = self.parent_gui.txt_mapping_config
        total_items = len(config_list)

        if direction == "up" and current_index > 0:
            config_list[current_index], config_list[current_index - 1] = config_list[current_index - 1], config_list[current_index]
            self.recreate_txt_field_rows(reselect_index=current_index - 1)
        elif direction == "down" and current_index < total_items - 1:
            config_list[current_index], config_list[current_index + 1] = config_list[current_index + 1], config_list[current_index]
            self.recreate_txt_field_rows(reselect_index=current_index + 1)

    def add_txt_field_row(self):
        # CORRECTED: Add to the new txt_mapping_config attribute
        new_field_index = len(self.parent_gui.txt_mapping_config) + 1
        self.parent_gui.txt_mapping_config.append({
            "field": f"Custom_Field_{new_field_index}",
            "column_name": f"Custom_Col_{new_field_index}",
            "skip": False
        })
        self.recreate_txt_field_rows(reselect_index=len(self.parent_gui.txt_mapping_config) - 1)

    def remove_txt_field_row(self, index_to_remove):
        # CORRECTED: Remove from the new txt_mapping_config attribute
        if not (0 <= index_to_remove < len(self.parent_gui.txt_mapping_config)):
            return
        
        config_to_remove = self.parent_gui.txt_mapping_config[index_to_remove]
        if messagebox.askyesno("Confirm Deletion", f"Are you sure you want to remove field '{config_to_remove['field']}'?", parent=self.master):
            del self.parent_gui.txt_mapping_config[index_to_remove]
            
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
        ToolTip(threshold_spinbox, "A file is considered 'active' if it was modified within this many seconds.\nIf inactive, the cell will be left blank.")

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
        self.folder_row_widgets = {}
        self.add_folder_header(self.scrollable_frame)

    def add_folder_header(self, parent):
        # Configure the grid columns on the parent frame
        parent.columnconfigure(0, weight=2, minsize=140)  # Folder Type
        parent.columnconfigure(1, weight=4, minsize=250)  # Monitor Path
        parent.columnconfigure(2, weight=0)               # ... button
        parent.columnconfigure(3, weight=2, minsize=150)  # Excel Column
        parent.columnconfigure(5, weight=1, minsize=80)   # File Ext.
        parent.columnconfigure(6, weight=0, minsize=50)   # Skip?
        parent.columnconfigure(7, weight=0, minsize=70)   # Log 'X'?

        # 1. Create a header frame to contain the labels
        header_frame = ttk.Frame(parent, style="Header.TFrame")
        header_frame.grid(row=0, column=0, columnspan=8, sticky="ew")

        # 2. Add header labels to the new frame
        ttk.Label(header_frame, text="Folder Type", font=("Arial", 10, "bold"), style="Header.TLabel").grid(row=0, column=0, sticky='w', padx=(15, 5))
        ttk.Label(header_frame, text="Monitor Path", font=("Arial", 10, "bold"), style="Header.TLabel").grid(row=0, column=1, sticky='w', padx=5)
        # Empty label for browse button column to maintain spacing
        ttk.Label(header_frame, text="", style="Header.TLabel").grid(row=0, column=2)
        ttk.Label(header_frame, text="Excel Column", font=("Arial", 10, "bold"), style="Header.TLabel").grid(row=0, column=3, sticky='w', padx=5)
        ttk.Label(header_frame, text="", style="Header.TLabel", padding=5).grid(row=0, column=4, sticky='w')
        ttk.Label(header_frame, text="File Ext.", font=("Arial", 10, "bold"), style="Header.TLabel").grid(row=0, column=5, sticky='w', padx=5)
        ttk.Label(header_frame, text="Skip?", font=("Arial", 10, "bold"), style="Header.TLabel").grid(row=0, column=6, sticky='w', padx=5)
        ttk.Label(header_frame, text="Log 'X'?", font=("Arial", 10, "bold"), style="Header.TLabel").grid(row=0, column=7, sticky='w', padx=5)

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

        # --- Place Widgets on the Shared Grid ---
        label.grid(row=row_index, column=0, padx=5, pady=2, sticky="ew")
        entry.grid(row=row_index, column=1, padx=5, pady=2, sticky="ew")
        button.grid(row=row_index, column=2, padx=(0,5), pady=2, sticky='w')
        column_entry.grid(row=row_index, column=3, padx=5, pady=2, sticky="ew")
        extension_entry.grid(row=row_index, column=5, padx=5, pady=2, sticky="ew")
        skip_checkbox.grid(row=row_index, column=6, padx=(15, 5), pady=2, sticky='w')
        log_x_checkbox.grid(row=row_index, column=7, padx=(15, 5), pady=2, sticky='w')

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
    

        # Store references for selection, saving, and removal
        self.folder_entries[folder_name] = entry
        self.folder_column_entries[folder_name] = column_entry
        self.file_extension_entries[folder_name] = extension_entry
        self.folder_skip_vars[folder_name] = skip_var
        self.folder_log_x_vars[folder_name] = log_x_var
        # Store all widgets in the row for highlighting
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
        Creates the tab for configuring automatic timed events with an improved layout.
        """
        tab = ttk.Frame(self.notebook, padding=20)
        self.notebook.add(tab, text="Programmed Events")

        # Use a main grid to structure the tab content
        tab.columnconfigure(0, weight=1)
        
        # 1. Midnight 'New Day' Event Configuration
        new_day_frame = ttk.LabelFrame(tab, text="Midnight 'New Day' Event", padding=15)
        new_day_frame.grid(row=0, column=0, sticky='ew', pady=(0, 15))
        new_day_frame.columnconfigure(1, weight=1) # Allow second column to expand

        # Row 0: Enable Checkbox
        new_day_check = ttk.Checkbutton(new_day_frame, text="Enable this automatic event", 
                                        variable=self.parent_gui.new_day_event_enabled_var,
                                        style="Large.TCheckbutton")
        new_day_check.grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 10))
        ToolTip(new_day_check, "If checked, an event will be logged automatically at midnight.")

        # Rows 1-2: Color Pickers
        ttk.Label(new_day_frame, text="Excel Row Colors:").grid(row=1, column=0, sticky='w', padx=5, pady=(2, 0))
        self._create_color_picker_widgets(new_day_frame, 1, "New Day")


        # 2. Hourly KP Log Event Configuration
        hourly_frame = ttk.LabelFrame(tab, text="Hourly KP Log Event", padding=15)
        hourly_frame.grid(row=1, column=0, sticky='ew', pady=5)
        hourly_frame.columnconfigure(1, weight=1)

        # Row 0: Enable Checkbox
        hourly_check = ttk.Checkbutton(hourly_frame, text="Enable this automatic event",
                                    variable=self.parent_gui.hourly_event_enabled_var,
                                    style="Large.TCheckbutton")
        hourly_check.grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 10))
        ToolTip(hourly_check, "If checked, the current KP will be logged automatically every hour.")
        
        # Rows 1-2: Color Pickers
        ttk.Label(hourly_frame, text="Excel Row Colors:").grid(row=1, column=0, sticky='w', padx=5, pady=(2, 0))
        self._create_color_picker_widgets(hourly_frame, 1, "Hourly KP Log")


        # 3. Log off Distance/Speed Calculation
        logoff_frame = ttk.LabelFrame(tab, text="Log off Distance/Speed Calculation", padding=15)
        logoff_frame.grid(row=2, column=0, sticky='ew', pady=5)
        logoff_frame.columnconfigure(1, weight=1)
        
        # Row 0: Enable Checkbox
        logoff_check = ttk.Checkbutton(logoff_frame, text="Calculate distance & speed on Log off",
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
        self.parent_gui.txt_source_aliases["TXT Source 4"] = self.txt_name_set4_var.get().strip()
        self.parent_gui.txt_folder_path_set4 = self.txt_path_set4_var.get().strip()
        self.parent_gui.txt_source_aliases["TXT Source 5"] = self.txt_name_set5_var.get().strip()
        self.parent_gui.txt_folder_path_set5 = self.txt_path_set5_var.get().strip()
        
        # --- TXT File Mapping Tab ---
        new_txt_mapping_configs = []
        for i, row_info in enumerate(self.txt_field_row_widgets):
            field_name = row_info["field_entry_widget"].get().strip() or f"Custom_Field_{i+1}"
            column_name = row_info["column_entry"].get().strip() or field_name
            skip_value = row_info["skip_var"].get()
            new_txt_mapping_configs.append({
                "field": field_name, "column_name": column_name, "skip": skip_value
            })
        self.parent_gui.txt_mapping_config = new_txt_mapping_configs

        # --- Generated Fields Tab ---
        new_generated_configs = []
        for i, widget_info in enumerate(self.generated_field_widgets):
            original_config = self.parent_gui.generated_fields_config[i]
            original_config["column_name"] = widget_info["entry"].get().strip()
            original_config["skip"] = widget_info["skip_var"].get()
            new_generated_configs.append(original_config)
        self.parent_gui.generated_fields_config = new_generated_configs

        # --- Static Fields Tab ---
        new_static_configs = []
        for i, row_info in enumerate(self.static_field_row_widgets):
            field_name = row_info["column_entry"].get().strip()
            description = row_info["description_entry"].get().strip()
            cell_ref = row_info["cell_entry"].get().strip()
            skip_value = row_info["skip_var"].get()
            new_static_configs.append({
                "field": field_name, "description": description, "column_name": cell_ref, "skip": skip_value
            })
        self.parent_gui.static_field_configs = new_static_configs

        # --- Monitored Folders Tab ---
        # NEWLY ADDED: This block reads the data from the Monitored Folders UI
        parent_folder_paths = {}
        parent_folder_cols = {}
        parent_folder_exts = {}
        parent_folder_skips = {}
        parent_folder_log_x_instead = {}
        for folder_name in self.folder_entries.keys():
            folder_path = self.folder_entries[folder_name].get().strip()
            if folder_path: # Only save configurations that have a path
                parent_folder_paths[folder_name] = folder_path
                parent_folder_cols[folder_name] = self.folder_column_entries[folder_name].get().strip()
                parent_folder_exts[folder_name] = self.file_extension_entries[folder_name].get().strip().lstrip('.')
                parent_folder_skips[folder_name] = self.folder_skip_vars[folder_name].get()
                parent_folder_log_x_instead[folder_name] = self.folder_log_x_vars[folder_name].get()
        
        self.parent_gui.folder_paths = parent_folder_paths
        self.parent_gui.folder_columns = parent_folder_cols
        self.parent_gui.file_extensions = parent_folder_exts
        self.parent_gui.folder_skips = parent_folder_skips
        self.parent_gui.folder_log_x_instead = parent_folder_log_x_instead

        # --- Button Configuration Tab ---
        # (Assuming your button saving logic is here and correct)

        # --- Final Actions ---
        self.parent_gui.save_settings()
        self.parent_gui.update_custom_buttons()
        
        

    def load_settings(self):
        """Loads settings from the parent DataLoggerGUI instance and populates the UI."""
        
        # --- File Paths Tab ---
        self.log_file_entry.delete(0, tk.END)
        self.log_file_entry.insert(0, self.parent_gui.log_file_path or "")
        
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

        

        # --- Data Columns Tab ---
        self.parent_gui.load_settings()
        self.recreate_txt_field_rows()
        self.master.after_idle(lambda: self.txt_fields_canvas.config(scrollregion=self.txt_fields_canvas.bbox("all")))

        # --- Monitored Folders Tab ---
        for name in list(self.folder_row_widgets.keys()):
            if name not in self.parent_gui.folder_paths:
                widgets_to_destroy = self.folder_row_widgets.pop(name, [])
                for widget in widgets_to_destroy:
                    if widget and widget.winfo_exists():
                        widget.destroy()
        self.folder_entries.clear()
        self.folder_column_entries.clear()
        self.folder_db_column_entries.clear()
        self.file_extension_entries.clear()
        self.folder_skip_vars.clear()
        self.folder_log_x_vars.clear()
        self.add_initial_folder_rows()
        self.master.after_idle(self.update_scroll_region)

        # --- Button Configuration Tab ---
        self.num_buttons_entry.delete(0, tk.END)
        self.num_buttons_entry.insert(0, str(self.parent_gui.num_custom_buttons))
        self.recreate_custom_button_settings()

        # This logic is handled inside the recreate_txt_field_rows and recreate_static_field_rows methods now.
        self.recreate_txt_field_rows()
        self.recreate_static_field_rows()
        self.master.after_idle(lambda: self.txt_fields_canvas.config(scrollregion=self.txt_fields_canvas.bbox("all")))
        self.master.after_idle(lambda: self.static_fields_canvas.config(scrollregion=self.static_fields_canvas.bbox("all")))
        
        # --- Programmed Events Tab ---
        self.parent_gui.new_day_event_enabled_var.set(self.parent_gui.new_day_event_enabled_var.get())
        self.parent_gui.hourly_event_enabled_var.set(self.parent_gui.hourly_event_enabled_var.get())
        self.parent_gui.calculate_logoff_values.set(self.parent_gui.calculate_logoff_values.get())
        
        # Load the color values into the UI variables
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