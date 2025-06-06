import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser, Toplevel, Label, Frame, Entry, Button, StringVar, BooleanVar
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
        # self.widget.bind("<Destroy>", self.on_leave, add='+') # Might cause issues

    def on_enter(self, event=None):
        self.cancel_scheduled_hide()
        self.schedule_show()

    def on_leave(self, event=None):
        self.cancel_scheduled_show()
        self.schedule_hide()

    def schedule_show(self):
        self.cancel_scheduled_show()
        self.show_id = self.widget.after(self.show_delay, self.show_tooltip)

    def schedule_hide(self):
        self.cancel_scheduled_hide()
        self.hide_id = self.widget.after(max(100, self.hide_delay // 5) , self.hide_tooltip)

    def cancel_scheduled_show(self):
        if self.show_id:
            try:
                self.widget.after_cancel(self.show_id)
            except ValueError: pass
            self.show_id = None

    def cancel_scheduled_hide(self):
        if self.hide_id:
            try:
                self.widget.after_cancel(self.hide_id)
            except ValueError: pass
            self.hide_id = None

    def show_tooltip(self):
        if not self.widget.winfo_exists() or not self.widget.winfo_ismapped():
            self.hide_tooltip()
            return

        self.hide_tooltip() 

        try:
            x, y, _, _ = self.widget.bbox("insert")
            if x is None or y is None: x = y = 0 
        except tk.TclError: 
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
            self.hide_id = self.widget.after(5000, self.hide_tooltip)
        except tk.TclError: 
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
                        if self.gui_instance and hasattr(self.gui_instance, 'master') and self.gui_instance.master.winfo_exists():
                           self.gui_instance.master.after(0, self.gui_instance.update_status, f"Latest {self.folder_name} file: {file_name}")
        except Exception as e:
            print(f"Error updating cache for {self.folder_name}: {e}")
            traceback.print_exc()


# --- Button Editor Dialog ---
class ButtonEditorDialog(Toplevel):
    def __init__(self, parent, gui_instance, config_index=None, config_set=1): # Added config_set
        super().__init__(parent)
        self.gui_instance = gui_instance
        self.config_index = config_index
        self.config_set = config_set # 1 for primary, 2 for secondary
        self.is_edit_mode = config_index is not None
        self.pastel_colors = ["#FFB3BA", "#FFDFBA", "#FFFFBA", "#BAFFC9", "#BAE1FF", "#E0BBE4", "#FFC8A2", "#D4A5A5", "#A2D4AB", "#A2C4D4"]

        self.title(f"Edit Custom Button (Set {config_set})" if self.is_edit_mode else f"Add Custom Button (Set {config_set})")
        self.geometry("450x250")
        self.minsize(400, 230)
        self.transient(parent)
        self.grab_set()

        main_frame = ttk.Frame(self, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="Button Text:").grid(row=0, column=0, sticky=tk.W, pady=(0,5))
        self.text_entry = ttk.Entry(main_frame, width=40)
        self.text_entry.grid(row=0, column=1, columnspan=2, sticky=tk.EW, pady=(0,5))
        ToolTip(self.text_entry, "Text displayed on the button.")

        ttk.Label(main_frame, text="Event Text (for Log):").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.event_entry = ttk.Entry(main_frame, width=40)
        self.event_entry.grid(row=1, column=1, columnspan=2, sticky=tk.EW, pady=5)
        ToolTip(self.event_entry, "Text written to the 'Event' column in the log when this button is pressed.")

        ttk.Label(main_frame, text="Button/Row Color:").grid(row=2, column=0, sticky=tk.W, pady=5)
        color_widget_frame = ttk.Frame(main_frame)
        color_widget_frame.grid(row=2, column=1, columnspan=2, sticky=tk.EW, pady=5)
        self.selected_color_var = StringVar()
        self.color_display_label = tk.Label(color_widget_frame, width=4, relief="solid", borderwidth=1)
        self.color_display_label.pack(side=tk.LEFT, padx=(0, 5))
        clear_btn = ttk.Button(color_widget_frame, text="X", width=2, style="Toolbutton", command=lambda: self._set_color_in_dialog(None))
        clear_btn.pack(side=tk.LEFT, padx=1); ToolTip(clear_btn, "Clear color.")
        presets_frame = ttk.Frame(color_widget_frame)
        presets_frame.pack(side=tk.LEFT, padx=(2, 2))
        for p_color in self.pastel_colors[:5]:
            try:
                b = tk.Button(presets_frame, bg=p_color, width=1, height=1, relief="raised", bd=1, command=lambda c=p_color: self._set_color_in_dialog(c))
                b.pack(side=tk.LEFT, padx=1)
            except tk.TclError: pass
        choose_btn = ttk.Button(color_widget_frame, text="...", width=3, style="Toolbutton", command=self._choose_color_in_dialog)
        choose_btn.pack(side=tk.LEFT, padx=1); ToolTip(choose_btn, "Choose a custom color.")

        button_bar = ttk.Frame(main_frame)
        button_bar.grid(row=3, column=0, columnspan=3, sticky=tk.E, pady=(20,0))
        ttk.Button(button_bar, text="Save", command=self.save_button_config, style="Accent.TButton").pack(side=tk.LEFT, padx=5)
        ttk.Button(button_bar, text="Cancel", command=self.destroy).pack(side=tk.LEFT)

        self.load_initial_values()
        self.text_entry.focus_set()
        self.protocol("WM_DELETE_WINDOW", self.destroy)

    def _set_color_in_dialog(self, color_hex):
        valid_color = None
        if color_hex:
            temp_label = None
            try: temp_label = tk.Label(self); temp_label.config(background=color_hex); valid_color = color_hex
            except tk.TclError: print(f"Warning (ButtonEditorDialog): Invalid color code '{color_hex}'."); valid_color = None
            finally:
                if temp_label is not None:
                    try: temp_label.destroy()
                    except tk.TclError: pass
        self.selected_color_var.set(valid_color if valid_color else "")
        try: self.color_display_label.config(background=valid_color if valid_color else 'SystemButtonFace')
        except tk.TclError: self.color_display_label.config(background='SystemButtonFace')

    def _choose_color_in_dialog(self):
        current_color = self.selected_color_var.get()
        btn_name_for_title = self.text_entry.get().strip() or ("New Button" if not self.is_edit_mode else f"Button (Set {self.config_set})")
        color_code = colorchooser.askcolor(color=current_color if current_color else None, title=f"Choose Color for {btn_name_for_title}", parent=self)
        if color_code and color_code[1]: self._set_color_in_dialog(color_code[1])

    def load_initial_values(self):
        initial_text = ""
        initial_event = ""
        initial_color_hex = None
        
        configs_list = self.gui_instance.custom_button_configs if self.config_set == 1 else self.gui_instance.custom_button_configs_set2

        if self.is_edit_mode and self.config_index < len(configs_list):
            config = configs_list[self.config_index]
            initial_text = config.get("text", f"Custom {self.config_index + 1}")
            initial_event = config.get("event_text", f"{initial_text} Triggered")
            _, initial_color_hex = self.gui_instance.button_colors.get(initial_text, (None, None))
        else: 
            num_existing = len(configs_list)
            initial_text = f"Custom {num_existing + 1}"
            initial_event = f"{initial_text} Triggered"

        self.text_entry.insert(0, initial_text)
        self.event_entry.insert(0, initial_event)
        if initial_color_hex: self._set_color_in_dialog(initial_color_hex)

    def save_button_config(self):
        new_text = self.text_entry.get().strip()
        new_event_text = self.event_entry.get().strip()
        new_color_hex = self.selected_color_var.get() if self.selected_color_var.get() else None

        if not new_text: messagebox.showerror("Input Error", "Button text cannot be empty.", parent=self); return
        if not new_event_text: new_event_text = f"{new_text} Triggered"
        
        old_text_for_color_key = None
        configs_list_ref = self.gui_instance.custom_button_configs if self.config_set == 1 else self.gui_instance.custom_button_configs_set2
        num_buttons_attr_name = "num_custom_buttons" if self.config_set == 1 else "num_custom_buttons_set2"

        if self.is_edit_mode:
            if self.config_index < len(configs_list_ref):
                old_text_for_color_key = configs_list_ref[self.config_index].get('text')
                configs_list_ref[self.config_index] = {"text": new_text, "event_text": new_event_text}
            else: messagebox.showerror("Error", "Could not find button to edit.", parent=self); return
        else: # Add mode
            max_buttons = 10 # Hard limit for now
            if len(configs_list_ref) >= max_buttons:
                 if not messagebox.askyesno("Limit Reached", f"You have reached the limit of {max_buttons} custom buttons for this set. Add anyway?\n(This may not display correctly if the 'Number of Custom Buttons' in main Settings is not also updated).", parent=self):
                     return
            configs_list_ref.append({"text": new_text, "event_text": new_event_text})
            setattr(self.gui_instance, num_buttons_attr_name, len(configs_list_ref))

        if old_text_for_color_key and old_text_for_color_key != new_text:
            if old_text_for_color_key in self.gui_instance.button_colors:
                color_tuple = self.gui_instance.button_colors.pop(old_text_for_color_key)
                self.gui_instance.button_colors[new_text] = (color_tuple[0], new_color_hex)
            else: self.gui_instance.button_colors[new_text] = (None, new_color_hex)
        else: self.gui_instance.button_colors[new_text] = (None, new_color_hex)

        self.gui_instance.create_main_buttons(self.gui_instance.button_frame)
        self.gui_instance.save_settings()
        self.destroy()

# --- Main Application GUI Class ---
class DataLoggerGUI:
    def __init__(self, master):
        self.master = master
        master.title("Data Acquisition Logger (SQLite Mode)")
        master.geometry("550x600") # Increased height for second button set
        master.minsize(450, 450) # Increased min height
        self.settings_file = "logger_settings.json"
        self.init_styles()
        self.init_variables()
        self.load_settings()

        self.main_frame = ttk.Frame(self.master, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.master.columnconfigure(0, weight=1)
        self.master.rowconfigure(0, weight=1)

        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.rowconfigure(0, weight=1) 
        self.main_frame.rowconfigure(1, weight=0) 
        self.main_frame.rowconfigure(2, weight=0) 

        self.button_frame = ttk.Frame(self.main_frame, padding="10")
        self.button_frame.grid(row=0, column=0, sticky=(tk.N, tk.W, tk.E, tk.S))
        self.create_main_buttons(self.button_frame) # Create buttons

        self.create_status_indicators(self.main_frame) 
        self.create_status_bar(self.main_frame)     

        self.schedule_new_day()
        self.start_monitoring() 

    def init_styles(self):
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
        self.style.configure("Row0.TFrame", background="#ffffff")
        self.style.configure("Row1.TFrame", background="#f5f5f5")
        self.style.configure("TLabelframe", background="#f0f0f0", padding=5)
        self.style.configure("TLabelframe.Label", background="#f0f0f0", font=("Arial", 10, "bold"))
        self.style.configure("Large.TCheckbutton", font=("Arial", 11))
        self.style.configure("Toolbutton", padding=2) 
        self.style.configure("Accent.TButton", font=("Arial", 10, "bold"), foreground="white", background="#0078D4")
        self.style.map("TButton",
                       foreground=[('pressed', 'darkblue'), ('active', 'blue'), ('disabled', '#999999')],
                       background=[('pressed', '!disabled', '#c0c0c0'), ('active', '#e0e0e0')]
                       )

    def init_variables(self):
        # Primary TXT and Custom Buttons (Set 1)
        self.log_file_path = None
        self.txt_folder_path = None
        self.txt_file_path = None
        self.txt_field_columns = {"Event": "Event"}
        self.txt_field_skips = {}
        self.num_custom_buttons = 3
        self.custom_button_configs = [
            {"text": "Custom Event 1", "event_text": "Custom Event 1 Triggered"},
            {"text": "Custom Event 2", "event_text": "Custom Event 2 Triggered"},
            {"text": "Custom Event 3", "event_text": "Custom Event 3 Triggered"},
        ]

        # Secondary TXT and Custom Buttons (Set 2) - NEW
        self.txt_folder_path_set2 = None
        self.txt_file_path_set2 = None # To store latest found file for set 2
        self.txt_field_columns_set2 = {"Event": "Event"} # Default Event column for set 2
        self.txt_field_skips_set2 = {}
        self.num_custom_buttons_set2 = 0 # Default to 0 for the second set
        self.custom_button_configs_set2 = []


        # Shared variables
        self.folder_paths = {}
        self.folder_columns = {}
        self.file_extensions = {}
        self.folder_skips = {}
        self.monitors = {}
        
        self.button_colors = { 
            "Log on": (None, "#90EE90"), "Log off": (None, "#FFB6C1"),
            "Event": (None, "#FFFFE0"), "SVP": (None, "#ADD8E6"),
            "New Day": (None, "#FFFF99")
        }
        for i in range(10): # Placeholders for up to 10 buttons in set 1
            self.button_colors[f"Custom Event {i+1}"] = (None, None)
        # No separate button_colors for set 2 initially; colors are keyed by button text.
        # If Set 2 buttons have unique names, they'll get their own entries.
        # If names overlap, they'll share color settings (simplification for now).

        self.sqlite_enabled = False
        self.sqlite_db_path = None
        self.sqlite_table = "EventLog"

        self.status_var = tk.StringVar()
        self.monitor_status_label = None
        self.db_status_label = None
        self.settings_window_instance = None

    def create_main_buttons(self, parent_frame):
        for widget in parent_frame.winfo_children(): widget.destroy()

        logging_frame = ttk.LabelFrame(parent_frame, text="Logging Actions", padding=10)
        logging_frame.grid(row=0, column=0, padx=5, pady=(0, 5), sticky="nsew")
        logging_frame.columnconfigure(0, weight=1)

        # --- Custom Events Area with Notebook ---
        custom_events_outer_frame = ttk.LabelFrame(parent_frame, text="Custom Events", padding=10)
        custom_events_outer_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
        custom_events_outer_frame.columnconfigure(0, weight=1)
        custom_events_outer_frame.rowconfigure(0, weight=1)

        custom_notebook = ttk.Notebook(custom_events_outer_frame)
        custom_notebook.pack(expand=True, fill='both', pady=5)

        # --- Tab 1 for Custom Buttons (Set 1) ---
        tab1_frame = ttk.Frame(custom_notebook, padding=5)
        custom_notebook.add(tab1_frame, text="Custom Set 1")
        tab1_frame.bind("<Button-3>", lambda e: self.show_add_button_context_menu(e, config_set=1))
        ToolTip(tab1_frame, "Right-click here to add a new button to Set 1.")
        
        # --- Tab 2 for Custom Buttons (Set 2) ---
        tab2_frame = ttk.Frame(custom_notebook, padding=5)
        custom_notebook.add(tab2_frame, text="Custom Set 2")
        tab2_frame.bind("<Button-3>", lambda e: self.show_add_button_context_menu(e, config_set=2))
        ToolTip(tab2_frame, "Right-click here to add a new button to Set 2.")

        num_custom_cols = 2 
        tab1_frame.columnconfigure(list(range(num_custom_cols)), weight=1)
        tab2_frame.columnconfigure(list(range(num_custom_cols)), weight=1)
        # Row configure handled dynamically

        other_frame = ttk.LabelFrame(parent_frame, text="Other Actions", padding=10)
        other_frame.grid(row=0, column=1, padx=5, pady=(0, 5), sticky="nsew")
        other_frame.columnconfigure(0, weight=1)

        parent_frame.columnconfigure(0, weight=1); parent_frame.columnconfigure(1, weight=1)
        parent_frame.rowconfigure(0, weight=0); parent_frame.rowconfigure(1, weight=1) 

        all_buttons_data = {
            "Log on":  {"command_ref": self.log_event, "frame": logging_frame, "tooltip": "Record 'Log on' marker."},
            "Log off": {"command_ref": self.log_event, "frame": logging_frame, "tooltip": "Record 'Log off' marker."},
            "Event":   {"command_ref": self.log_event, "frame": logging_frame, "tooltip": "Record data from TXT."},
            "SVP":     {"command_ref": self.apply_svp, "frame": other_frame, "tooltip": "Record data with SVP."},
            "New Day": {"command_ref": self.log_new_day, "frame": other_frame, "tooltip": "Trigger 'New Day' log."},
            "Settings":{"command_ref": self.open_settings, "frame": other_frame, "tooltip": "Open configuration."},
            "Sync Excel->DB":{"command_ref": self.sync_excel_to_sqlite_triggered, "frame": other_frame, "tooltip": "Sync Excel to SQLite."}
        }
        
        # --- Populate Standard Buttons ---
        buttons_dict = {}
        btn_row_log, btn_row_other = 0, 0
        for text, data in all_buttons_data.items():
            button = ttk.Button(data["frame"], text=text, style="TButton")
            buttons_dict[text] = {"widget": button, "data": data} # Store for command assignment
            # Grid standard buttons
            current_frame = data["frame"]
            if current_frame == logging_frame:
                button.grid(row=btn_row_log, column=0, padx=5, pady=4, sticky="ew"); current_frame.rowconfigure(btn_row_log, weight=1); btn_row_log += 1
            elif current_frame == other_frame:
                button.grid(row=btn_row_other, column=0, padx=5, pady=4, sticky="ew"); current_frame.rowconfigure(btn_row_other, weight=1); btn_row_other += 1
            ToolTip(button, data["tooltip"])
            # Assign command (standard buttons)
            cmd_ref = data["command_ref"]
            if text in ["Log on", "Log off", "Event"]: cmd = lambda t=text, b=button, ref=cmd_ref: ref(t, b)
            elif text in ["SVP", "New Day"]: cmd = lambda b=button, ref=cmd_ref: ref(b)
            else: cmd = cmd_ref # For Settings, Sync Excel->DB
            button.config(command=cmd)


        # --- Populate Custom Buttons for Set 1 ---
        self._populate_custom_button_set(tab1_frame, self.custom_button_configs, self.num_custom_buttons, num_custom_cols, config_set=1)
        
        # --- Populate Custom Buttons for Set 2 ---
        self._populate_custom_button_set(tab2_frame, self.custom_button_configs_set2, self.num_custom_buttons_set2, num_custom_cols, config_set=2)

        parent_frame.update_idletasks()

    def _populate_custom_button_set(self, target_tab_frame, configs_list, num_buttons, num_cols, config_set):
        """Helper to populate a tab with its custom buttons."""
        # Ensure configs_list has enough placeholders if num_buttons is greater
        while len(configs_list) < num_buttons:
            idx = len(configs_list) + 1
            default_text = f"Custom {idx}"
            if config_set == 2: default_text = f"CustomS2 {idx}" # Distinguish default names slightly
            configs_list.append({"text": default_text, "event_text": f"{default_text} Triggered"})
        
        valid_configs_for_set = configs_list[:num_buttons] # Use only up to num_buttons

        custom_idx_set = 0
        for i, config_data in enumerate(valid_configs_for_set):
            button_text = config_data.get("text", f"Custom {i+1}" if config_set==1 else f"CustomS2 {i+1}")
            event_desc = config_data.get("event_text", f"{button_text} Triggered")
            tooltip_text = f"Log '{event_desc}'. Right-click to edit/remove."
            
            # Button styling (optional, if specific GUI colors are desired for buttons themselves)
            # For now, using default ttk.Button style
            button = ttk.Button(target_tab_frame, text=button_text, style="TButton")
            
            cmd = lambda cfg=config_data, b=button, cs=config_set: self.log_custom_event(cfg, b, txt_source_set=cs)
            button.config(command=cmd)
            
            # Bind right-click for edit/remove to this specific button, passing its set and index
            button.bind("<Button-3>", lambda e, b=button, idx=i, cs=config_set: self.show_edit_remove_button_context_menu(e, b, idx, cs))
            
            custom_row = custom_idx_set // num_cols
            custom_col = custom_idx_set % num_cols
            button.grid(row=custom_row, column=custom_col, padx=5, pady=4, sticky="nsew")
            target_tab_frame.rowconfigure(custom_row, weight=1) # Allow rows to expand if needed
            custom_idx_set += 1
            ToolTip(button, tooltip_text)

    def show_add_button_context_menu(self, event, config_set):
        """Shows context menu for adding a new custom button to a specific set."""
        context_menu = tk.Menu(self.master, tearoff=0)
        context_menu.add_command(label=f"Add New Button to Set {config_set}...", 
                                 command=lambda cs=config_set: self.open_button_editor_dialog(config_set=cs))
        try: context_menu.tk_popup(event.x_root, event.y_root)
        finally: context_menu.grab_release()

    def show_edit_remove_button_context_menu(self, event, button_widget, config_index, config_set):
        """Shows context menu for editing/removing a button from a specific set."""
        context_menu = tk.Menu(self.master, tearoff=0)
        context_menu.add_command(label="Edit This Button...", 
                                 command=lambda idx=config_index, cs=config_set: self.open_button_editor_dialog(config_index=idx, config_set=cs))
        context_menu.add_command(label="Remove This Button...", 
                                 command=lambda idx=config_index, cs=config_set: self.remove_custom_button_action(idx, cs))
        try: context_menu.tk_popup(event.x_root, event.y_root)
        finally: context_menu.grab_release()

    def open_button_editor_dialog(self, config_index=None, config_set=1):
        if hasattr(self, '_button_editor_active') and self._button_editor_active and self._button_editor_active.winfo_exists():
            self._button_editor_active.lift(); return
        editor_dialog = ButtonEditorDialog(self.master, self, config_index=config_index, config_set=config_set)
        self._button_editor_active = editor_dialog
        self.master.wait_window(editor_dialog)
        if hasattr(self, '_button_editor_active'): del self._button_editor_active

    def remove_custom_button_action(self, config_index, config_set):
        configs_list_ref = self.custom_button_configs if config_set == 1 else self.custom_button_configs_set2
        num_buttons_attr_name = "num_custom_buttons" if config_set == 1 else "num_custom_buttons_set2"

        if not (0 <= config_index < len(configs_list_ref)):
            messagebox.showerror("Error", "Invalid button index for removal.", parent=self.master); return
        
        button_config = configs_list_ref[config_index]
        button_text = button_config.get("text", f"Button at index {config_index} (Set {config_set})")

        if messagebox.askyesno("Confirm Removal", f"Remove button '{button_text}' from Set {config_set}?", parent=self.master):
            del configs_list_ref[config_index]
            setattr(self, num_buttons_attr_name, len(configs_list_ref))
            if button_text in self.button_colors: del self.button_colors[button_text]
            self.create_main_buttons(self.button_frame)
            self.save_settings()
            self.update_status(f"Button '{button_text}' (Set {config_set}) removed.")

    def create_status_indicators(self, parent_frame):
        indicator_frame = ttk.Frame(parent_frame, padding="5 0")
        indicator_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        indicator_frame.columnconfigure(1, weight=0)
        indicator_frame.columnconfigure(3, weight=0)
        indicator_frame.columnconfigure(4, weight=1) 

        ttk.Label(indicator_frame, text="Monitoring:", font=("Arial", 9, "bold")).grid(row=0, column=0, sticky=tk.W, padx=(0, 2))
        self.monitor_status_label = ttk.Label(indicator_frame, text="Initializing...", foreground="orange", font=("Arial", 9))
        self.monitor_status_label.grid(row=0, column=1, sticky=tk.W)

        ttk.Label(indicator_frame, text="SQLite:", font=("Arial", 9, "bold")).grid(row=0, column=2, sticky=tk.W, padx=(15, 2))
        self.db_status_label = ttk.Label(indicator_frame, text="Initializing...", foreground="orange", font=("Arial", 9))
        self.db_status_label.grid(row=0, column=3, sticky=tk.W)
        ttk.Frame(indicator_frame).grid(row=0, column=4) 
        self.update_db_indicator()

    def sync_excel_to_sqlite_triggered(self):
        print("Sync Excel -> SQLite button pressed.")
        if not self.sqlite_enabled:
            messagebox.showwarning("Sync Skipped", "SQLite logging is not enabled in Settings.", parent=self.master)
            self.update_status("Sync Error: SQLite disabled.")
            return
        if not self.log_file_path or not os.path.exists(self.log_file_path):
            messagebox.showerror("Sync Error", "Excel log file path is not set or the file does not exist.", parent=self.master)
            self.update_status("Sync Error: Excel file path invalid.")
            return
        if not self.sqlite_db_path:
            messagebox.showerror("Sync Error", "SQLite database path is not set.", parent=self.master)
            self.update_status("Sync Error: SQLite DB path missing.")
            return

        sync_button = None
        target_button_text = "Sync Excel->DB"
        try:
            if hasattr(self, 'button_frame') and self.button_frame:
                for child in self.button_frame.winfo_children(): 
                    if isinstance(child, ttk.LabelFrame): # Check outer frames first
                        for btn_candidate in child.winfo_children(): # Iterate through widgets in LabelFrame
                             if isinstance(btn_candidate, ttk.Button) and btn_candidate.cget('text') == target_button_text:
                                sync_button = btn_candidate
                                break
                    if sync_button: break
                    # Also check if buttons are directly in button_frame (though current structure uses LabelFrames)
                    elif isinstance(child, ttk.Button) and child.cget('text') == target_button_text:
                        sync_button = child
                        break

            else:
                print("Debug: button_frame attribute not found during sync button search.")
        except Exception as e:
            print(f"Debug: Could not find sync button to disable: {e}")

        original_text = None
        if sync_button:
            try:
                if sync_button.winfo_exists():
                    original_text = sync_button['text']
                    sync_button.config(state=tk.DISABLED, text="Syncing...")
            except tk.TclError:
                print("Warning: Could not disable sync button (already destroyed?).")
                sync_button = None 
        else:
            print("Debug: Sync button widget not found, cannot disable.")

        self.update_status("Starting sync from Excel to SQLite...")

        def _sync_worker():
            nonlocal original_text 
            success, message = self.perform_excel_to_sqlite_sync()
            print(f"Sync thread finished. Status: {message}")
            if not success:
                self.master.after(0, lambda m=message: messagebox.showerror("Sync Failed", m, parent=self.master))
            self.master.after(0, self.update_status, message)

            if sync_button:
                def re_enable_sync_button(btn=sync_button, txt=original_text):
                    try:
                        if btn and btn.winfo_exists():
                            btn.config(state=tk.NORMAL)
                            if txt: btn.config(text=txt)
                    except tk.TclError:
                        print("Warning: Could not re-enable sync button.")
                self.master.after(0, re_enable_sync_button)
        sync_thread = threading.Thread(target=_sync_worker, daemon=True)
        sync_thread.start()

    def perform_excel_to_sqlite_sync(self):
        # This method remains unchanged from the previous version as it's not directly
        # affected by the addition of a second custom button set.
        # ... (Keep existing perform_excel_to_sqlite_sync logic) ...
        print("\n--- Starting perform_excel_to_sqlite_sync ---")
        excel_file = self.log_file_path
        db_file = self.sqlite_db_path
        db_table = self.sqlite_table
        record_id_column = "RecordID"
        date_col_name = self.txt_field_columns.get("Date", "Date") 
        time_col_name = self.txt_field_columns.get("Time", "Time") 
        print(f"Sync Params: Excel='{excel_file}', DB='{db_file}', Table='{db_table}', ID Col='{record_id_column}'")

        if not excel_file or not db_file or not db_table:
            print("Sync Error: Missing file paths or table name.")
            return False, "Sync Error: Configuration paths or table missing."
        
        excel_data = {}
        app = None; wb = None; sheet = None; header = None; df_excel = None
        try:
            print("Sync Step 1: Reading Excel file...")
            app = xw.App(visible=False, add_book=False)
            wb = app.books.open(excel_file, update_links=False, read_only=True)
            sheet = wb.sheets[0]
            header_range = sheet.range('A1').expand('right')
            header = header_range.value
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
                except Exception: last_row = 1
            
            if last_row <= 1:
                print("Sync Info: Excel sheet appears empty or has only header.")
                if wb: wb.close(); wb = None
                if app: app.quit(); app = None
                return True, "Sync Info: Excel sheet is empty, nothing to sync."

            data_range = sheet.range((2, 1), (last_row, len(header)))
            df_excel = pd.DataFrame(data_range.value, columns=header)

            if date_col_name in df_excel.columns:
                df_excel[date_col_name] = pd.to_datetime(df_excel[date_col_name], errors='coerce')
            if time_col_name in df_excel.columns: 
                def format_excel_time(excel_time_float):
                    if pd.isna(excel_time_float) or not isinstance(excel_time_float, (float, int)): return None
                    try:
                        total_seconds = int(excel_time_float * 24 * 60 * 60)
                        hours, remainder = divmod(total_seconds, 3600)
                        minutes, seconds = divmod(remainder, 60)
                        return f"{hours:02}:{minutes:02}:{seconds:02}"
                    except Exception: return str(excel_time_float)
                df_excel[time_col_name] = df_excel[time_col_name].apply(format_excel_time)

            if record_id_column in df_excel.columns:
                df_excel[record_id_column] = df_excel[record_id_column].astype(str).replace({'nan': '', 'None': '', None: ''})
                df_excel = df_excel[df_excel[record_id_column].str.strip() != '']
                df_excel = df_excel.dropna(subset=[record_id_column])
            else: raise ValueError(f"'{record_id_column}' column disappeared.")

            if df_excel.empty:
                print("Sync Info: No valid rows with RecordIDs found after cleaning Excel data.")
                if wb: wb.close(); wb = None;
                if app: app.quit(); app = None;
                return True, "Sync Info: No valid Excel rows found to sync."

            df_excel = df_excel.set_index(record_id_column, drop=False)
            excel_data = df_excel.to_dict('index')
            print(f"Sync Step 1 Complete: Read {len(excel_data)} rows with valid RecordIDs from Excel.")
        except Exception as e_excel:
            print(f"Sync Error (Step 1 - Excel): {type(e_excel).__name__} - {e_excel}"); traceback.print_exc()
            return False, f"Sync Error: Reading Excel failed ({type(e_excel).__name__})"
        finally:
            if wb: 
                try: wb.close()
                except: pass
            if app: 
                try: app.quit()
                except: pass
            wb = None; app = None
        
        sqlite_data = {}
        conn_sqlite = None; db_cols = []
        try:
            print("Sync Step 2: Reading SQLite database...")
            conn_sqlite = sqlite3.connect(db_file, timeout=10)
            conn_sqlite.row_factory = sqlite3.Row
            cursor = conn_sqlite.cursor()
            cursor.execute(f"SELECT name FROM sqlite_master WHERE type='table' AND name=?;", (db_table,))
            if cursor.fetchone() is None:
                return False, f"Sync Error: SQLite table '{db_table}' does not exist."
            
            cursor.execute(f"PRAGMA table_info([{db_table}])")
            cols_info = cursor.fetchall()
            db_cols = [col['name'] for col in cols_info]
            if record_id_column not in db_cols:
                return False, f"Sync Error: Column '{record_id_column}' not found in SQLite table."

            quoted_db_cols = ", ".join([f"[{c}]" for c in db_cols])
            cursor.execute(f"SELECT {quoted_db_cols} FROM [{db_table}]")
            rows = cursor.fetchall()
            for row in rows:
                row_dict = dict(row)
                rec_id = str(row_dict.get(record_id_column, '')).strip()
                if rec_id: sqlite_data[rec_id] = row_dict
            print(f"Sync Step 2 Complete: Read {len(sqlite_data)} rows from SQLite.")
        except sqlite3.Error as e_sqlite:
            print(f"Sync Error (Step 2 - SQLite): {e_sqlite}"); traceback.print_exc()
            return False, f"Sync Error: Reading SQLite failed - {type(e_sqlite).__name__}"
        finally:
            if conn_sqlite: conn_sqlite.close()

        updates_to_apply = []
        records_processed = 0; records_updated = 0; db_cols_set = set(db_cols)
        print(f"Sync Step 3: Comparing {len(excel_data)} Excel rows to {len(sqlite_data)} SQLite rows...")

        for rec_id, excel_row_dict in excel_data.items():
            records_processed += 1
            sqlite_row_dict = sqlite_data.get(rec_id)
            if not sqlite_row_dict: continue 

            row_needs_update = False
            current_record_updates = {} 

            for excel_col_name, excel_val in excel_row_dict.items():
                if excel_col_name in db_cols_set and excel_col_name != record_id_column:
                    sqlite_val = sqlite_row_dict.get(excel_col_name)
                    formatted_excel_val = excel_val
                    if isinstance(excel_val, pd.Timestamp): 
                        formatted_excel_val = excel_val.strftime('%Y-%m-%d') if not pd.isna(excel_val) else None
                    
                    str_formatted_excel = str(formatted_excel_val) if formatted_excel_val is not None else ""
                    str_sqlite_val = str(sqlite_val) if sqlite_val is not None else ""

                    if str_formatted_excel != str_sqlite_val:
                        current_record_updates[excel_col_name] = formatted_excel_val 
                        row_needs_update = True
            
            if row_needs_update:
                updates_to_apply.append({'id': rec_id, 'changes': current_record_updates})
                records_updated +=1
        
        print(f"Sync Step 3 Complete: Comparison found differences for {records_updated} records.")
        if not updates_to_apply:
            print("Sync Step 4: No differences found requiring update.")
            return True, f"Sync complete. No changes detected in {records_processed} Excel rows."

        print(f"Sync Step 4: Applying updates for {len(updates_to_apply)} records to SQLite...")
        conn_sqlite = None 
        try:
            conn_sqlite = sqlite3.connect(db_file, timeout=10)
            cursor = conn_sqlite.cursor()
            update_statements_run = 0; rows_affected_total = 0
            for update_item in updates_to_apply:
                rec_id = update_item['id']
                col_val_dict = update_item['changes']
                set_clauses = []
                values_for_sql = []
                for col, val in col_val_dict.items():
                    set_clauses.append(f"[{col}] = ?")
                    values_for_sql.append(val)
                if set_clauses:
                    values_for_sql.append(rec_id) 
                    sql_update = f"UPDATE [{db_table}] SET {', '.join(set_clauses)} WHERE [{record_id_column}] = ?"
                    cursor.execute(sql_update, values_for_sql)
                    update_statements_run += 1
                    rows_affected_total += cursor.rowcount
                    if cursor.rowcount == 0: print(f"    - Warning: UPDATE affected 0 rows for RecordID {rec_id}.")
            conn_sqlite.commit()
            print(f"  - Commit successful. {update_statements_run} UPDATEs executed, {rows_affected_total} total rows affected.")
            return True, f"Sync successful. Updated {records_updated} records ({rows_affected_total} rows affected)."
        except sqlite3.Error as e_update:
            print(f"Sync Error (Step 4 - SQLite Update): {e_update}"); traceback.print_exc()
            if conn_sqlite: conn_sqlite.rollback()
            return False, f"Sync Error: Updating SQLite failed - {type(e_update).__name__}"
        finally:
            if conn_sqlite: conn_sqlite.close()

    def create_status_bar(self, parent_frame):
        self.status_var.set("Status: Ready")
        status_bar = ttk.Label(parent_frame, textvariable=self.status_var, style="StatusBar.TLabel", anchor='w')
        status_bar.grid(row=2, column=0, sticky=(tk.W, tk.E), padx=0, pady=(5,0))

    def update_status(self, message):
        def _update():
            timestamp = datetime.datetime.now().strftime("%H:%M:%S")
            max_len = 100
            display_message = message if len(message) <= max_len else message[:max_len-3] + "..."
            try:
                if self.status_var: 
                    self.status_var.set(f"[{timestamp}] {display_message}")
            except tk.TclError: 
                print(f"Status Update Error: Could not set status_var. Message: {message}")
        if hasattr(self, 'master') and self.master.winfo_exists():
            try:
                self.master.after(0, _update)
            except tk.TclError: 
                print(f"Status Update Error: Could not schedule update. Message: {message}")

    def update_db_indicator(self):
        if not hasattr(self, 'db_status_label') or not self.db_status_label: return
        if not self.master.winfo_exists(): return 

        status_text = "Disabled"; status_color = "gray"
        if self.sqlite_enabled:
            if self.sqlite_db_path and os.path.exists(self.sqlite_db_path):
                status_text = "Enabled"; status_color = "green"
            elif self.sqlite_db_path:
                status_text = "File Missing"; status_color = "#E65C00" 
            else:
                status_text = "Path Missing"; status_color = "#E65C00"
        try:
            self.db_status_label.config(text=status_text, foreground=status_color)
        except tk.TclError:
            print("DB Indicator Error: Could not configure label.")

    def log_event(self, event_type, button_widget):
        print(f"'{event_type}' button pressed.")
        event_text_for_excel = None; skip_files = False
        if event_type in ["Log on", "Log off"]:
            event_text_for_excel = f"{event_type} event occurred"
        elif event_type == "Event":
            skip_files = True; event_text_for_excel = ""
            print("Event button: Logging data without event text, skipping filename check.")
        self._perform_log_action(event_type=event_type,
                                 event_text_for_excel=event_text_for_excel,
                                 skip_latest_files=skip_files,
                                 triggering_button=button_widget,
                                 txt_source_set=1) # Standard events use primary TXT source

    def log_custom_event(self, config, button_widget, txt_source_set=1): # Added txt_source_set
        button_text = config.get("text", "Unknown Custom")
        event_text_for_excel = config.get("event_text", f"{button_text} Triggered")
        print(f"'{button_text}' (Set {txt_source_set}) button pressed. Event text: '{event_text_for_excel}'")
        self._perform_log_action(event_type=button_text, 
                                 event_text_for_excel=event_text_for_excel,
                                 triggering_button=button_widget,
                                 txt_source_set=txt_source_set) # Pass the set number

    def log_new_day(self, button_widget=None): 
        print("Logging 'New Day' event.")
        self._perform_log_action(event_type="New Day",
                                 event_text_for_excel="New Day",
                                 triggering_button=button_widget,
                                 txt_source_set=1) # New Day uses primary TXT source

    def apply_svp(self, button_widget):
        print("Applying SVP...")
        # SVP uses primary TXT data and folder monitors, so txt_source_set=1
        if not self.log_file_path or not self.txt_folder_path or "SVP" not in self.folder_paths:
            messagebox.showinfo("Info", "Please select log file, TXT folder (Set 1), and configure SVP folder path/column in Settings.", parent=self.master)
            self.update_status("SVP Error: Configuration missing."); return
        if not self.folder_columns.get("SVP"):
            messagebox.showinfo("Info", "Please configure the 'Target Column' for SVP in Folder Settings.", parent=self.master)
            self.update_status("SVP Error: Target column missing."); return
        if self.log_file_path and not os.path.exists(self.log_file_path):
            messagebox.showerror("Error", f"Excel Log file does not exist:\n{self.log_file_path}", parent=self.master)
            self.update_status("SVP Error: Excel file missing."); return
        self._perform_log_action(event_type="SVP",
                                 event_text_for_excel="SVP applied",
                                 svp_specific_handling=True,
                                 triggering_button=button_widget,
                                 txt_source_set=1) # SVP uses primary TXT source

    def _perform_log_action(self, event_type, event_text_for_excel, skip_latest_files=False, svp_specific_handling=False, triggering_button=None, txt_source_set=1): # Added txt_source_set
        print(f"Queueing log action: Type='{event_type}', TXT Set='{txt_source_set}'")
        self.update_status(f"Processing '{event_type}' (TXT Set {txt_source_set})...")
        original_text = None
        if triggering_button and isinstance(triggering_button, ttk.Button):
            try:
                if triggering_button.winfo_exists(): 
                    original_text = triggering_button['text']
                    triggering_button.config(state=tk.DISABLED, text="Working...")
            except tk.TclError:
                print("Warning: Could not disable button."); triggering_button = None

        def _worker_thread_func():
            nonlocal original_text 
            row_data = {}; excel_success = False; sqlite_logged = False
            excel_save_exception = None; sqlite_save_exception_type = None
            status_msg = f"'{event_type}' processed with errors."
            record_id = str(uuid.uuid4())
            row_data['RecordID'] = record_id
            try:
                # Determine which Event column name to use based on txt_source_set
                current_txt_field_columns = self.txt_field_columns if txt_source_set == 1 else self.txt_field_columns_set2
                event_col_name = current_txt_field_columns.get("Event", "Event")

                row_data["EventType"] = event_type # For SQLite, always include EventType
                if event_text_for_excel is not None:
                    row_data[event_col_name] = event_text_for_excel
                
                try:
                    txt_data = self.insert_txt_data(txt_source_set=txt_source_set) # Pass txt_source_set
                    if txt_data: row_data.update(txt_data)
                except Exception as e_txt:
                    print(f"Thread '{event_type}': Error fetching TXT (Set {txt_source_set}) data: {e_txt}")
                    self.master.after(0, lambda e=e_txt: messagebox.showerror("Error", f"Failed to read TXT (Set {txt_source_set}) data:\n{e}", parent=self.master))
                
                if not skip_latest_files: # Folder monitors are global, not per-set
                    try:
                        latest_files_data = self.get_latest_files_data()
                        if latest_files_data: row_data.update(latest_files_data)
                    except Exception as e_files:
                        print(f"Thread '{event_type}': Error fetching latest file data: {e_files}")
                        self.master.after(0, lambda e=e_files: messagebox.showerror("Error", f"Failed to get latest file data:\n{e}", parent=self.master))
                
                if svp_specific_handling: # SVP logic also global
                    svp_folder_path = self.folder_paths.get("SVP")
                    svp_col_name = self.folder_columns.get("SVP", "SVP")
                    if svp_folder_path and svp_col_name:
                        latest_svp_file = folder_cache.get("SVP")
                        row_data[svp_col_name] = latest_svp_file if latest_svp_file else "N/A"
                    elif svp_col_name:
                        print("SVP folder not defined in settings.")
                        row_data[svp_col_name] = "Config Error"
                
                if row_data:
                    color_tuple = self.button_colors.get(event_type, (None, None))
                    row_color_for_excel = color_tuple[1] if isinstance(color_tuple, tuple) and len(color_tuple) > 1 else None
                    excel_data = {k: v for k, v in row_data.items() if k != 'EventType'}
                    try:
                        if not self.log_file_path: excel_save_exception = ValueError("Excel path missing")
                        elif not os.path.exists(self.log_file_path): excel_save_exception = FileNotFoundError("Excel file missing")
                        else:
                            self.save_to_excel(excel_data, row_color=row_color_for_excel)
                            excel_success = True
                    except Exception as e_excel:
                        excel_save_exception = e_excel
                        print(f"Thread '{event_type}': Error saving to Excel: {e_excel}"); traceback.print_exc()
                        self.master.after(0, lambda e=e_excel: messagebox.showerror("Excel Error", f"Failed to save to Excel:\n{e}", parent=self.master))
                    
                    sqlite_logged, sqlite_save_exception_type = self.log_to_sqlite(row_data)
                    status_parts = []
                    if excel_success: status_parts.append("Excel: OK")
                    elif excel_save_exception: status_parts.append(f"Excel: Fail ({type(excel_save_exception).__name__})")
                    else: status_parts.append("Excel: Fail (Path?)")
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
                    status_msg = f"'{event_type}' pressed, no data collected."
                    print(f"Thread '{event_type}': No data generated.")
            except Exception as thread_ex:
                print(f"!!! Unexpected Error in logging thread for '{event_type}' !!!"); print(traceback.format_exc())
                status_msg = f"'{event_type}' - Thread error: {thread_ex}"
                self.master.after(0, lambda e=thread_ex: messagebox.showerror("Thread Error", f"Critical error logging '{event_type}':\n{e}", parent=self.master))
            finally:
                print(f"Thread '{event_type}': Action finished. Status: {status_msg}")
                self.master.after(0, self.update_status, status_msg)
                if triggering_button and isinstance(triggering_button, ttk.Button):
                    def re_enable_button(btn=triggering_button, txt=original_text):
                        try:
                            if btn and btn.winfo_exists():
                                btn.config(state=tk.NORMAL)
                                if txt: btn.config(text=txt)
                        except tk.TclError: print("Warn: Could not re-enable button.")
                    self.master.after(0, re_enable_button)
        log_thread = threading.Thread(target=_worker_thread_func, daemon=True)
        log_thread.start()

    def insert_txt_data(self, txt_source_set=1): # Added txt_source_set
        row_data = {}; current_dt = datetime.datetime.now(); current_timestamp = time.time()
        use_pc_time = False; reason_for_pc_time = ""

        # Determine which set of TXT configs to use
        if txt_source_set == 1:
            current_txt_folder_path = self.txt_folder_path
            current_txt_file_path_attr = 'txt_file_path' # Attribute name to set on self
            current_txt_field_columns = self.txt_field_columns
            current_txt_field_skips = self.txt_field_skips
            set_label_for_messages = "Set 1"
        elif txt_source_set == 2:
            current_txt_folder_path = self.txt_folder_path_set2
            current_txt_file_path_attr = 'txt_file_path_set2'
            current_txt_field_columns = self.txt_field_columns_set2
            current_txt_field_skips = self.txt_field_skips_set2
            set_label_for_messages = "Set 2"
        else: # Should not happen
            print(f"Error: Invalid txt_source_set: {txt_source_set}")
            return {}


        if not current_txt_folder_path or not os.path.exists(current_txt_folder_path):
            print(f"Warn: TXT folder path (Set {txt_source_set}) missing or invalid."); use_pc_time = True; reason_for_pc_time = f"TXT folder (Set {txt_source_set}) path missing"; setattr(self, current_txt_file_path_attr, None)
        else:
            latest_file_for_set = self.find_latest_file_in_folder(current_txt_folder_path, ".txt")
            setattr(self, current_txt_file_path_attr, latest_file_for_set) # Store it on self
            if not latest_file_for_set:
                print(f"Warn: No TXT file found (Set {txt_source_set})."); use_pc_time = True; reason_for_pc_time = f"No TXT file found (Set {txt_source_set})"
        
        current_txt_file_to_check = getattr(self, current_txt_file_path_attr) # Get the stored path

        if current_txt_file_to_check and not use_pc_time:
            try:
                file_mod_timestamp = os.path.getmtime(current_txt_file_to_check); time_diff = current_timestamp - file_mod_timestamp
                if time_diff > 1.0:
                    print(f"Info: TXT file (Set {txt_source_set}) '{os.path.basename(current_txt_file_to_check)}' modified {time_diff:.2f}s ago.")
                    use_pc_time = True; reason_for_pc_time = f"file (Set {txt_source_set}) modified {time_diff:.2f}s ago"
            except OSError as e_modtime:
                print(f"Warn: Could not get mod time for TXT (Set {txt_source_set}) '{current_txt_file_to_check}': {e_modtime}.")
                use_pc_time = True; reason_for_pc_time = f"failed to get mod time (Set {txt_source_set})"
        
        txt_data_found = False; parse_success = True; temp_txt_data = {}
        if not use_pc_time and current_txt_file_to_check:
            print(f"Info: Reading TXT (Set {txt_source_set}) file: {os.path.basename(current_txt_file_to_check)}")
            try:
                lines = []; encodings_to_try = ['utf-8', 'latin-1', 'cp1252']; read_success = False; last_error = None
                for enc in encodings_to_try:
                    try:
                        for attempt in range(3):
                            try:
                                with open(current_txt_file_to_check, "r", encoding=enc) as file: lines = file.readlines()
                                read_success = True; break 
                            except IOError as e_io:
                                if attempt < 2: time.sleep(0.1); continue
                                else: raise e_io 
                        if read_success: break 
                    except UnicodeDecodeError: last_error = f"UnicodeDecodeError with {enc}"; continue
                    except Exception as e_open: last_error = f"Error reading TXT (Set {txt_source_set}) file {os.path.basename(current_txt_file_to_check)} with {enc}: {e_open}"; print(last_error); lines = []; break
                if not read_success and not lines:
                    print(f"Warn: Could not decode/read TXT (Set {txt_source_set}) file. Last error: {last_error}")
                    use_pc_time = True; reason_for_pc_time = f"failed to read/decode TXT (Set {txt_source_set})"
                if lines:
                    latest_line_str = lines[-1].strip(); latest_line_parts = latest_line_str.split(",")
                    field_keys = ["Date", "Time", "KP", "DCC", "Line name", "Latitude", "Longitude", "Easting", "Northing"] # Standard field keys
                    for i, field_key in enumerate(field_keys):
                        excel_col = current_txt_field_columns.get(field_key); skip_field = current_txt_field_skips.get(field_key, False)
                        if excel_col and not skip_field:
                            try:
                                value = latest_line_parts[i].strip(); temp_txt_data[excel_col] = value
                                if field_key in ["Date", "Time"]: txt_data_found = True 
                            except IndexError:
                                temp_txt_data[excel_col] = None; print(f"Warn: Field '{field_key}' missing in TXT (Set {txt_source_set}) for '{excel_col}'.")
                                if field_key in ["Date", "Time"]: parse_success = False 
                            except Exception as e_parse:
                                temp_txt_data[excel_col] = None; print(f"Warn: Err parsing '{field_key}' for '{excel_col}' (Set {txt_source_set}). Error: {e_parse}")
                                if field_key in ["Date", "Time"]: parse_success = False 
                    if not txt_data_found: print(f"Info: No Date/Time field found/mapped in TXT (Set {txt_source_set})."); use_pc_time = True; reason_for_pc_time = f"Date/Time not found/mapped in TXT (Set {txt_source_set})"
                    elif not parse_success: print(f"Warn: Failed to parse Date/Time from TXT (Set {txt_source_set})."); use_pc_time = True; reason_for_pc_time = f"Date/Time parsing failed in TXT (Set {txt_source_set})"
                else: 
                    if not use_pc_time: print(f"Info: TXT (Set {txt_source_set}) file found but empty/unreadable."); use_pc_time = True; reason_for_pc_time = f"TXT (Set {txt_source_set}) file empty/unreadable"
            except Exception as e: 
                print(f"Error processing TXT (Set {txt_source_set}) file: {e}"); traceback.print_exc()
                if not use_pc_time: use_pc_time = True; reason_for_pc_time = f"unexpected error processing TXT (Set {txt_source_set}): {type(e).__name__}"
        
        if use_pc_time:
            print(f"Info: Using PC Time/Date for {set_label_for_messages}. Reason: {reason_for_pc_time}.")
            date_col = current_txt_field_columns.get("Date"); time_col = current_txt_field_columns.get("Time")
            skip_date = current_txt_field_skips.get("Date", False); skip_time = current_txt_field_skips.get("Time", False)
            if date_col and not skip_date: row_data[date_col] = current_dt.strftime("%Y-%m-%d")
            if time_col and not skip_time: row_data[time_col] = current_dt.strftime("%H:%M:%S")
            for col, val in temp_txt_data.items():
                if col != date_col and col != time_col: 
                    if col not in row_data: row_data[col] = val
        else:
            print(f"Info: Using Date and Time from recent TXT file (Set {txt_source_set}).")
            row_data.update(temp_txt_data)
        return row_data

    def get_latest_files_data(self): # This is global for monitored folders
        latest_files = {}
        for folder_name, folder_path in self.folder_paths.items():
            if not folder_path or self.folder_skips.get(folder_name, False): continue
            latest_file = folder_cache.get(folder_name); column_name = self.folder_columns.get(folder_name)
            if not column_name: continue
            if latest_file: latest_files[column_name] = latest_file
            else: latest_files[column_name] = "N/A"
        return latest_files

    def find_latest_file_in_folder(self, folder_path, extension=".txt"):
        try:
            files = []; ext_lower = extension.lower()
            for f in os.listdir(folder_path):
                f_path = os.path.join(folder_path, f)
                try:
                    if os.path.isfile(f_path) and f.lower().endswith(ext_lower): files.append(f_path)
                except OSError: continue
            return max(files, key=os.path.getmtime) if files else None
        except FileNotFoundError: print(f"Error: Folder not found '{folder_path}'"); return None
        except Exception as e: print(f"Error finding latest file in '{folder_path}': {e}"); return None

    def save_to_excel(self, row_data, row_color=None, next_row=None):
        # ... (Unchanged from previous version)
        if not self.log_file_path: raise ValueError("Excel log file path is missing.")
        if not os.path.exists(self.log_file_path): raise FileNotFoundError(f"Excel log file not found: {self.log_file_path}")
        app = None; workbook = None; opened_new_app = False; opened_workbook = False
        print(f"Debug Excel: Attempting to write {len(row_data)} columns.")
        try:
            try:
                app = xw.apps.active
                if app is None: raise Exception("No active Excel instance")
            except Exception:
                try: app = xw.App(visible=False); opened_new_app = True
                except Exception as e_app: raise ConnectionAbortedError(f"Failed to start/connect to Excel: {e_app}")
            target_norm_path = os.path.normcase(os.path.abspath(self.log_file_path))
            for wb in app.books:
                try:
                    if os.path.normcase(os.path.abspath(wb.fullname)) == target_norm_path:
                        workbook = wb; break
                except Exception as e_fullname: print(f"Warn: Error checking workbook fullname: {e_fullname}")
            if workbook is None:
                try: workbook = app.books.open(self.log_file_path); opened_workbook = True
                except Exception as e_open: raise IOError(f"Failed to open Excel workbook: {e_open}")
            sheet = workbook.sheets[0]
            header_range_obj = sheet.range("A1").expand("right"); header_values = header_range_obj.value
            if not header_values or not any(h is not None for h in header_values):
                raise ValueError("Excel header row (A1) is missing or empty.")
            record_id_col_name = "RecordID" 
            if record_id_col_name not in header_values:
                print(f"Fatal Excel Error: Header row does not contain a '{record_id_col_name}' column.")
                raise ValueError(f"Excel header missing required '{record_id_col_name}' column.")
            header_map_lower = {str(h).lower(): i + 1 for i, h in enumerate(header_values) if h is not None}
            last_header_col_index = max(header_map_lower.values()) if header_map_lower else 1
            if next_row is None:
                try:
                    last_row_a = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
                    check_row = last_row_a + 1; 
                    if check_row < 2: check_row = 2 
                    while sheet.range(f'A{check_row}').value is not None: check_row += 1
                    next_row = check_row
                except Exception as e_row: print(f"Warn Excel: Error finding next empty row ({e_row}). Defaulting to row 2."); next_row = 2
            
            written_cols = []
            for col_name, value in row_data.items():
                col_name_lower = str(col_name).lower()
                if col_name_lower in header_map_lower:
                    col_index = header_map_lower[col_name_lower]
                    try:
                        write_value = value
                        if col_name == record_id_col_name: 
                            sheet.range(next_row, col_index).number_format = '@' 
                            write_value = str(value)
                        elif isinstance(value, datetime.time): 
                             write_value = value.strftime("%H:%M:%S")
                        elif isinstance(value, datetime.date) and not isinstance(value, datetime.datetime):
                            write_value = value.strftime("%Y-%m-%d")
                        
                        sheet.range(next_row, col_index).value = write_value
                        written_cols.append(col_index)
                    except Exception as e_write_cell:
                        print(f"Warn Excel: Failed to write '{value}' to ({next_row},{col_index}) for '{col_name}'. Error: {e_write_cell}")
            if row_color and written_cols:
                try:
                    target_range = sheet.range((next_row, 1), (next_row, last_header_col_index))
                    target_range.color = row_color
                except Exception as e_color: print(f"Warn Excel: Failed to apply color to row {next_row}. Error: {e_color}")
            try: workbook.save()
            except Exception as e_save: print(f"Error saving workbook: {e_save}"); raise IOError(f"Failed to save Excel workbook: {e_save}")
        except Exception as e: print(f"Error during save_to_excel: {e}"); traceback.print_exc(); raise e
        finally:
            if workbook is not None and opened_workbook:
                try: workbook.close(save_changes=False)
                except Exception as e_close: print(f"Warn: Error closing workbook: {e_close}")
            if app is not None and opened_new_app:
                try: app.quit()
                except Exception as e_quit: print(f"Warn: Error quitting Excel: {e_quit}")
            if app is not None and opened_new_app: app = None
            elif app is not None and not opened_new_app: pass

    def log_to_sqlite(self, row_data):
        # ... (Unchanged from previous version, indentation already corrected)
        success = False
        error_type = None
        if not self.sqlite_enabled:
            return False, "Disabled"
        if not self.sqlite_db_path or not self.sqlite_table:
            msg = "SQLite Log Error: DB Path or Table Name missing."
            self.master.after(0, self.update_status, msg)
            print(msg)
            return False, "ConfigurationMissing"
        
        conn = None
        cursor = None
        try:
            conn = sqlite3.connect(self.sqlite_db_path, timeout=5)
            cursor = conn.cursor()
            table_columns_info = {}
            try:
                pragma_sql = f"PRAGMA table_info([{self.sqlite_table}]);"
                cursor.execute(pragma_sql)
                results = cursor.fetchall()
                if not results:
                    raise sqlite3.OperationalError(f"No such table: {self.sqlite_table}")
                table_columns_info = {row[1].lower(): row[1] for row in results}
            except sqlite3.Error as e_meta:
                print(f"SQLite Error: Could not retrieve columns for table '{self.sqlite_table}'. Error: {e_meta}")
                self.master.after(0, self.update_status, f"SQLite Log Error: Cannot get columns for '{self.sqlite_table}'")
                error_type = f"MetadataError_{type(e_meta).__name__}"
                raise e_meta 
            
            data_to_insert = {}
            provided_keys_lower = {str(k).lower(): k for k, v in row_data.items()}
            for lower_key, original_key in provided_keys_lower.items():
                if lower_key in table_columns_info:
                    db_col_name = table_columns_info[lower_key]
                    value = row_data[original_key]
                    if isinstance(value, datetime.time):
                        data_to_insert[db_col_name] = value.strftime("%H:%M:%S")
                    elif isinstance(value, datetime.date) and not isinstance(value, datetime.datetime):
                        data_to_insert[db_col_name] = value.strftime("%Y-%m-%d")
                    else:
                        data_to_insert[db_col_name] = value
            
            if not data_to_insert:
                print("SQLite Log Info: No matching columns between provided data and table. Nothing inserted.")
                self.master.after(0, self.update_status, "SQLite Log Info: No data matched DB columns.")
                return True, None

            cols = list(data_to_insert.keys())
            placeholders = ", ".join(["?"] * len(cols))
            col_name_string = ", ".join([f"[{c}]" for c in cols])
            sql_insert = f"INSERT INTO [{self.sqlite_table}] ({col_name_string}) VALUES ({placeholders})"
            values = [data_to_insert[c] for c in cols]
            
            cursor.execute(sql_insert, values)
            conn.commit()
            print("Debug SQLite: Record inserted and committed.")
            success = True
            error_type = None
    
        except sqlite3.OperationalError as op_err: 
            error_message = str(op_err)
            error_type = "OperationalError"
            print(f"SQLite Log Error (Operational): {error_message}")
            self.master.after(0, self.update_status, f"SQLite Log Error: {error_message}")
            if conn:
                try:
                    conn.rollback()
                    print("SQLite transaction rolled back (Operational Error).")
                except Exception as rb_ex:
                    print(f"SQLite Warning: Rollback failed after OpError: {rb_ex}")
            if "no such table" in error_message.lower():
                error_type = "NoSuchTable"
            elif "has no column named" in error_message.lower():
                error_type = "NoSuchColumn"
            elif "database is locked" in error_message.lower():
                error_type = "DatabaseLocked"
            self.master.after(0, lambda em=error_message, et=error_type: self.show_sqlite_error_message(em, et))
            success = False
    
        except sqlite3.Error as ex: 
            error_message = str(ex)
            error_type = type(ex).__name__
            print(f"SQLite Log Error ({error_type}): {error_message}"); traceback.print_exc()
            self.master.after(0, self.update_status, f"SQLite Log Error: {error_message}")
            if conn:
                try:
                    conn.rollback()
                    print(f"SQLite transaction rolled back ({error_type}).")
                except Exception as rb_ex:
                    print(f"SQLite Warning: Rollback failed after {error_type}: {rb_ex}")
            self.master.after(0, lambda et=error_type, em=error_message: messagebox.showerror("SQLite Error", f"Failed to log to SQLite database.\nType: {et}\nMessage: {em}", parent=self.master))
            success = False
    
        except Exception as e: 
            error_message = str(e)
            error_type = type(e).__name__
            print(f"SQLite Log Error (Unexpected - {error_type}): {error_message}"); traceback.print_exc()
            self.master.after(0, self.update_status, f"SQLite Log Error: Unexpected error ({error_type}).")
            if conn:
                try:
                    conn.rollback()
                    print(f"SQLite transaction rolled back (Unexpected {error_type}).")
                except Exception as rb_ex:
                    print(f"SQLite Warning: Rollback failed after Unexpected Error: {rb_ex}")
            self.master.after(0, lambda em=error_message: messagebox.showerror("Application Error", f"Unexpected error during SQLite logging:\n{em}", parent=self.master))
            success = False
    
        finally: 
            if cursor:
                try:
                    cursor.close()
                except Exception: pass 
            if conn:
                try:
                    conn.close()
                except Exception: pass
        return success, error_type

    def show_sqlite_error_message(self, error_message, error_type):
        # ... (Unchanged from previous version)
        parent_window = self.settings_window_instance if (hasattr(self, 'settings_window_instance') and self.settings_window_instance and self.settings_window_instance.winfo_exists()) else self.master
        if error_type == "NoSuchTable":
            messagebox.showerror("SQLite Error", f"Table '{self.sqlite_table}' not found.\nPlease check table name or create table.\nDB: {self.sqlite_db_path}", parent=parent_window)
        elif error_type == "NoSuchColumn":
            try: missing_col = error_message.split("column named")[-1].strip().split(":")[0].strip().strip("'\"[]")
            except: missing_col = "[unknown]"
            messagebox.showerror("SQLite Error", f"Column '{missing_col}' not found in table '{self.sqlite_table}'.\nCheck Settings vs. DB structure.\n\n(Err: {error_message})", parent=parent_window)
        elif error_type == "DatabaseLocked":
            messagebox.showerror("SQLite Error", f"Database file is locked.\nAnother program may be using it.\nDB: {self.sqlite_db_path}\n\n(Err: {error_message})", parent=parent_window)
        else: 
            messagebox.showerror("SQLite Operational Error", f"Error with database:\n{error_message}", parent=parent_window)


    def save_settings(self):
        colors_to_save = {}
        all_custom_configs = self.custom_button_configs + self.custom_button_configs_set2
        all_button_names_with_color_potential = list(self.button_colors.keys())
        
        for key in all_button_names_with_color_potential:
            color_tuple = self.button_colors.get(key)
            if color_tuple and color_tuple[1]: 
                colors_to_save[key] = color_tuple[1]
            # Clean up colors for custom buttons that no longer exist across both sets
            elif key not in ["Log on", "Log off", "Event", "SVP", "New Day"] and not any(cfg.get("text") == key for cfg in all_custom_configs):
                if key in self.button_colors:
                    print(f"Debug save_settings: Removing color for non-existent button '{key}'")
        settings = {
            "log_file_path": self.log_file_path, "txt_folder_path": self.txt_folder_path,
            "txt_field_columns": self.txt_field_columns, "txt_field_skips": self.txt_field_skips,
            "num_custom_buttons": self.num_custom_buttons, 
            "custom_button_configs": self.custom_button_configs[:self.num_custom_buttons],

            "txt_folder_path_set2": self.txt_folder_path_set2, # NEW
            "txt_field_columns_set2": self.txt_field_columns_set2, # NEW
            "txt_field_skips_set2": self.txt_field_skips_set2, # NEW
            "num_custom_buttons_set2": self.num_custom_buttons_set2, # NEW
            "custom_button_configs_set2": self.custom_button_configs_set2[:self.num_custom_buttons_set2], # NEW

            "folder_paths": self.folder_paths, "folder_columns": self.folder_columns,
            "file_extensions": self.file_extensions, "folder_skips": self.folder_skips,
            "button_colors": colors_to_save, 
            "sqlite_enabled": self.sqlite_enabled,
            "sqlite_db_path": self.sqlite_db_path, "sqlite_table": self.sqlite_table,
        }
        try:
            with open(self.settings_file, 'w') as f: json.dump(settings, f, indent=4)
            print("Settings saved successfully."); self.update_status("Settings saved.")
        except Exception as e:
            print(f"Error saving settings: {e}")
            messagebox.showerror("Save Error", f"Could not save settings to {self.settings_file}:\n{e}", parent=self.master)
            self.update_status("Error saving settings.")

    def load_settings(self):
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r') as f: settings = json.load(f)
                # Set 1
                self.log_file_path = settings.get("log_file_path")
                self.txt_folder_path = settings.get("txt_folder_path")
                self.txt_field_columns = {"Event": "Event"}; self.txt_field_columns.update(settings.get("txt_field_columns", {}))
                self.txt_field_skips.clear(); self.txt_field_skips.update(settings.get("txt_field_skips", {}))
                self.num_custom_buttons = settings.get("num_custom_buttons", 3)
                loaded_configs_s1 = settings.get("custom_button_configs", [])
                self.custom_button_configs = loaded_configs_s1[:self.num_custom_buttons]
                while len(self.custom_button_configs) < self.num_custom_buttons:
                    idx = len(self.custom_button_configs) + 1
                    self.custom_button_configs.append({"text": f"Custom {idx}", "event_text": f"Custom {idx} Event"})

                # Set 2 - NEW
                self.txt_folder_path_set2 = settings.get("txt_folder_path_set2")
                self.txt_field_columns_set2 = {"Event": "Event"}; self.txt_field_columns_set2.update(settings.get("txt_field_columns_set2", {}))
                self.txt_field_skips_set2.clear(); self.txt_field_skips_set2.update(settings.get("txt_field_skips_set2", {}))
                self.num_custom_buttons_set2 = settings.get("num_custom_buttons_set2", 0)
                loaded_configs_s2 = settings.get("custom_button_configs_set2", [])
                self.custom_button_configs_set2 = loaded_configs_s2[:self.num_custom_buttons_set2]
                while len(self.custom_button_configs_set2) < self.num_custom_buttons_set2:
                    idx = len(self.custom_button_configs_set2) + 1
                    self.custom_button_configs_set2.append({"text": f"CustomS2 {idx}", "event_text": f"CustomS2 {idx} Event"})
                
                # Shared
                self.folder_paths.clear(); self.folder_paths.update(settings.get("folder_paths", {}))
                self.folder_columns.clear(); self.folder_columns.update(settings.get("folder_columns", {}))
                self.file_extensions.clear(); self.file_extensions.update(settings.get("file_extensions", {}))
                self.folder_skips.clear(); self.folder_skips.update(settings.get("folder_skips", {}))
                
                loaded_colors_hex = settings.get("button_colors", {})
                default_colors = {
                    "Log on": (None, "#90EE90"), "Log off": (None, "#FFB6C1"),
                    "Event": (None, "#FFFFE0"), "SVP": (None, "#ADD8E6"), "New Day": (None, "#FFFF99")
                }
                self.button_colors = default_colors.copy()
                all_current_custom_configs = self.custom_button_configs + self.custom_button_configs_set2
                for config in all_current_custom_configs: # Ensure placeholders for all current buttons
                    btn_text = config.get("text")
                    if btn_text and btn_text not in self.button_colors: self.button_colors[btn_text] = (None, None)
                for key, color_hex in loaded_colors_hex.items(): # Apply loaded colors
                    if color_hex: self.button_colors[key] = (None, color_hex)

                self.sqlite_enabled = settings.get("sqlite_enabled", False)
                self.sqlite_db_path = settings.get("sqlite_db_path")
                self.sqlite_table = settings.get("sqlite_table", "EventLog")
                print("Settings loaded successfully."); self.update_status("Settings loaded.")
            else: # Settings file does not exist, use defaults from init_variables
                print("Settings file not found. Using default variables."); self.update_status("Settings file not found. Using defaults.")
                self.init_variables() # This will set defaults including for set2 variables

        except json.JSONDecodeError as e:
            print(f"Error loading settings: Invalid JSON. Error: {e}")
            messagebox.showerror("Load Error", f"Settings file '{self.settings_file}' invalid:\n{e}\n\nUsing defaults.", parent=self.master)
            self.update_status("Error loading settings: Invalid format."); self.init_variables() 
        except Exception as e:
            print(f"Error loading settings: {e}"); traceback.print_exc()
            messagebox.showerror("Load Error", f"Could not load settings:\n{e}\n\nUsing defaults.", parent=self.master)
            self.update_status("Error loading settings."); self.init_variables() 
        finally:
            if hasattr(self, 'button_frame') and self.button_frame: self.update_custom_buttons() 
            if hasattr(self, 'db_status_label') and self.db_status_label: self.update_db_indicator()

    def open_settings(self):
        if hasattr(self, 'settings_window_instance') and self.settings_window_instance and self.settings_window_instance.winfo_exists():
            self.settings_window_instance.lift(); self.settings_window_instance.focus_set()
        else:
            settings_top_level = tk.Toplevel(self.master); settings_top_level.title("Settings")
            settings_top_level.transient(self.master); settings_top_level.grab_set()
            self.settings_window_instance = settings_top_level 
            settings_gui = SettingsWindow(settings_top_level, self) # Pass self (DataLoggerGUI instance)
            settings_gui.load_settings() # Load current parent_gui settings into dialog
            self.master.wait_window(settings_top_level) 
            try: del self.settings_window_instance
            except AttributeError: pass

    def update_custom_buttons(self): # This now redraws both sets based on current config
        if hasattr(self, 'button_frame') and self.button_frame:
            print("Redrawing main buttons (including custom sets)...")
            self.create_main_buttons(self.button_frame)
            self.master.update_idletasks()
        else: print("Error: Button frame does not exist when trying to update buttons.")

    def start_monitoring(self):
        # ... (Unchanged from previous version, monitoring is global)
        print("Stopping existing monitors...")
        monitoring_was_active = False
        for name, monitor_observer in list(self.monitors.items()):
            try:
                if monitor_observer.is_alive():
                    monitor_observer.stop(); monitoring_was_active = True
                print(f"Stopped monitor '{name}'.")
            except Exception as e: print(f"Error stopping monitor '{name}': {e}")
        self.monitors.clear(); folder_cache.clear()
        if monitoring_was_active: print("All active monitors stopped.")
        else: print("No active monitors to stop.")

        print("Starting folder monitoring based on settings...")
        count = 0; monitoring_active = False
        for folder_name, folder_path in self.folder_paths.items():
            if folder_path and os.path.isdir(folder_path) and not self.folder_skips.get(folder_name, False):
                file_extension = self.file_extensions.get(folder_name, "")
                success = self.start_folder_monitoring(folder_name, folder_path, file_extension)
                if success: count += 1; monitoring_active = True
            elif self.folder_skips.get(folder_name): pass
            elif folder_path: print(f"Skipping monitor for '{folder_name}': Path ('{folder_path}') invalid.")
        
        print(f"Monitoring {count} folders.")
        self.update_status(f"Monitoring {count} active folders.")
        if hasattr(self, 'monitor_status_label') and self.monitor_status_label:
            if monitoring_active: self.monitor_status_label.config(text="Active", foreground="green")
            else: self.monitor_status_label.config(text="Inactive", foreground="red")
        self.update_db_indicator()


    def start_folder_monitoring(self, folder_name, folder_path, file_extension):
        # ... (Unchanged from previous version)
        try: os.listdir(folder_path) 
        except Exception as e: print(f"Error accessing '{folder_path}' for '{folder_name}': {e}. Monitor not started."); return False
        try:
            event_handler = FolderMonitor(folder_path, folder_name, self, file_extension)
            observer = PollingObserver(timeout=1)
            observer.schedule(event_handler, folder_path, recursive=False)
            observer.start(); self.monitors[folder_name] = observer
            threading.Thread(target=event_handler.update_latest_file, daemon=True).start()
            print(f"Started monitoring '{folder_name}' at '{folder_path}' (Ext: '{file_extension or 'Any'}')")
            return True
        except Exception as e: print(f"Error starting observer for '{folder_path}': {e}"); traceback.print_exc(); return False

    def schedule_new_day(self):
        # ... (Unchanged from previous version)
        now = datetime.datetime.now(); tomorrow = now.date() + datetime.timedelta(days=1)
        midnight = datetime.datetime.combine(tomorrow, datetime.time.min)
        time_until_midnight_ms = int((midnight - now).total_seconds() * 1000)
        trigger_delay_ms = time_until_midnight_ms + 1000 
        print(f"Scheduling next 'New Day' log in {(trigger_delay_ms / 1000 / 3600):.2f} hours.")
        if hasattr(self, '_new_day_timer_id') and self._new_day_timer_id:
            try: self.master.after_cancel(self._new_day_timer_id)
            except: pass
        self._new_day_timer_id = self.master.after(trigger_delay_ms, self.trigger_new_day)

    def trigger_new_day(self):
        # ... (Unchanged from previous version)
        print("--- Triggering Automatic New Day Log ---")
        self.log_new_day(button_widget=None)
        self.schedule_new_day()


# --- Settings Window Class ---
class SettingsWindow:
    def __init__(self, master, parent_gui):
        self.master = master; self.parent_gui = parent_gui
        self.master.title("Settings"); self.master.geometry("1000x700"); self.master.minsize(750, 550) # Adjusted size
        self.style = parent_gui.style 
        self.main_frame = ttk.Frame(self.master); self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.main_frame.rowconfigure(0, weight=1); self.main_frame.columnconfigure(0, weight=1)
        self.notebook = ttk.Notebook(self.main_frame); self.notebook.grid(row=0, column=0, sticky="nsew")
        
        self.create_general_tab()
        self.create_txt_column_mapping_tab_generic(set_num=1) # TXT Columns (Set 1)
        self.create_txt_column_mapping_tab_generic(set_num=2) # TXT Columns (Set 2)
        self.create_folder_selection_tab()
        self.create_custom_buttons_tab_generic(set_num=1) # Custom Buttons (Set 1)
        self.create_custom_buttons_tab_generic(set_num=2) # Custom Buttons (Set 2)
        self.create_sqlite_tab()
        
        button_frame = ttk.Frame(self.main_frame); button_frame.grid(row=1, column=0, pady=(10, 0), sticky="e")
        ttk.Button(button_frame, text="Save and Close", command=self.save_and_close, style="Accent.TButton").pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.master.destroy).pack(side=tk.RIGHT)

    def save_and_close(self):
        print("Saving settings from dialog..."); self.save_settings()
        print("Settings saved. Closing dialog."); self.master.destroy()

    def create_general_tab(self):
        tab = ttk.Frame(self.notebook, padding=20); self.notebook.add(tab, text="General / TXT Paths")
        
        # Excel Log File
        log_frame = ttk.LabelFrame(tab, text="Excel Log File (.xlsx)", padding=15); 
        log_frame.pack(fill="x", pady=(0, 10)); log_frame.columnconfigure(1, weight=1)
        ttk.Label(log_frame, text="Path:", anchor='e').grid(row=0, column=0, padx=(0, 5), pady=5, sticky='w')
        self.log_file_entry = ttk.Entry(log_frame, width=80); self.log_file_entry.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        ttk.Button(log_frame, text="Browse...", command=self.select_excel_file).grid(row=0, column=2, padx=(5, 0), pady=5)

        # TXT Data Folder (Set 1)
        txt_frame_s1 = ttk.LabelFrame(tab, text="Navigation TXT Data Folder (Set 1)", padding=15)
        txt_frame_s1.pack(fill="x", pady=10); txt_frame_s1.columnconfigure(1, weight=1)
        ttk.Label(txt_frame_s1, text="Folder:", anchor='e').grid(row=0, column=0, padx=(0, 5), pady=5, sticky='w')
        self.txt_folder_entry = ttk.Entry(txt_frame_s1, width=80); self.txt_folder_entry.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        ttk.Button(txt_frame_s1, text="Browse...", command=lambda: self.select_txt_folder(set_num=1)).grid(row=0, column=2, padx=(5,0), pady=5)
        ToolTip(self.txt_folder_entry, "Path to the folder for Set 1 TXT files.")
        
        # TXT Data Folder (Set 2) - NEW
        txt_frame_s2 = ttk.LabelFrame(tab, text="Navigation TXT Data Folder (Set 2)", padding=15)
        txt_frame_s2.pack(fill="x", pady=10); txt_frame_s2.columnconfigure(1, weight=1)
        ttk.Label(txt_frame_s2, text="Folder:", anchor='e').grid(row=0, column=0, padx=(0, 5), pady=5, sticky='w')
        self.txt_folder_entry_set2 = ttk.Entry(txt_frame_s2, width=80); self.txt_folder_entry_set2.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        ttk.Button(txt_frame_s2, text="Browse...", command=lambda: self.select_txt_folder(set_num=2)).grid(row=0, column=2, padx=(5,0), pady=5)
        ToolTip(self.txt_folder_entry_set2, "Path to the folder for Set 2 TXT files.")


    def select_excel_file(self):
        initial_dir = os.path.dirname(self.log_file_entry.get()) if self.log_file_entry.get() else "/"
        file_path = filedialog.askopenfilename(initialdir=initial_dir, filetypes=[("Excel files", "*.xlsx")], parent=self.master, title="Select Excel Log File")
        if file_path: self.log_file_entry.delete(0, tk.END); self.log_file_entry.insert(0, file_path)

    def select_txt_folder(self, set_num=1): # Added set_num
        entry_widget = self.txt_folder_entry if set_num == 1 else self.txt_folder_entry_set2
        initial_dir = entry_widget.get() if entry_widget.get() else "/"
        folder_path = filedialog.askdirectory(initialdir=initial_dir, parent=self.master, title=f"Select Navigation TXT Folder (Set {set_num})")
        if folder_path: entry_widget.delete(0, tk.END); entry_widget.insert(0, folder_path)

    def create_txt_column_mapping_tab_generic(self, set_num): # Made generic
        tab_title = f"TXT Columns (Set {set_num})"
        tab = ttk.Frame(self.notebook, padding=20); self.notebook.add(tab, text=tab_title)
        
        # Store widgets based on set_num
        if set_num == 1:
            self.txt_field_column_widgets_s1 = {}
            self.txt_field_skip_vars_s1 = {}
            widgets_dict_ref = self.txt_field_column_widgets_s1
            skip_vars_dict_ref = self.txt_field_skip_vars_s1
        else: # set_num == 2
            self.txt_field_column_widgets_s2 = {}
            self.txt_field_skip_vars_s2 = {}
            widgets_dict_ref = self.txt_field_column_widgets_s2
            skip_vars_dict_ref = self.txt_field_skip_vars_s2

        fields = ["Date", "Time", "KP", "DCC", "Line name", "Latitude", "Longitude", "Easting", "Northing"]; event_field = "Event"
        header_frame = ttk.Frame(tab, style="Header.TFrame", padding=(5,2)); header_frame.pack(fill='x', pady=(0,10))
        ttk.Label(header_frame, text="TXT Field", font=("Arial", 10, "bold"), width=15).pack(side=tk.LEFT, padx=5)
        ttk.Label(header_frame, text="Target Excel/DB Column Name", font=("Arial", 10, "bold"), width=30).pack(side=tk.LEFT, padx=5)
        ttk.Label(header_frame, text="Skip This Field?", font=("Arial", 10, "bold"), width=15).pack(side=tk.LEFT, padx=5)
        rows_frame = ttk.Frame(tab); rows_frame.pack(fill='both', expand=True); rows_frame.columnconfigure(1, weight=1)
        for i, field in enumerate(fields):
            row_f = ttk.Frame(rows_frame); row_f.grid(row=i, column=0, sticky='ew', pady=1)
            ttk.Label(row_f, text=f"{field}:", width=15).pack(side=tk.LEFT, padx=5)
            entry = ttk.Entry(row_f, width=30); entry.pack(side=tk.LEFT, padx=5, fill='x', expand=True)
            ToolTip(entry, f"Excel/DB column for '{field}' data (Set {set_num})."); widgets_dict_ref[field] = entry
            skip_var = tk.BooleanVar(); skip_checkbox = ttk.Checkbutton(row_f, variable=skip_var, text=""); skip_checkbox.pack(side=tk.LEFT, padx=20)
            ToolTip(skip_checkbox, f"Ignore '{field}' from TXT (Set {set_num})."); skip_vars_dict_ref[field] = skip_var
        
        event_row_index = len(fields); row_f = ttk.Frame(rows_frame); row_f.grid(row=event_row_index, column=0, sticky='ew', pady=1)
        ttk.Label(row_f, text=f"{event_field}:", width=15, font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
        entry = ttk.Entry(row_f, width=30); entry.pack(side=tk.LEFT, padx=5, fill='x', expand=True); entry.insert(0, "Event")
        ToolTip(entry, f"Column for event text (Set {set_num})."); widgets_dict_ref[event_field] = entry
        skip_var = tk.BooleanVar(); skip_checkbox = ttk.Checkbutton(row_f, variable=skip_var, text=""); skip_checkbox.pack(side=tk.LEFT, padx=20)
        ToolTip(skip_checkbox, f"Prevent writing event text (Set {set_num})."); skip_vars_dict_ref[event_field] = skip_var

    def create_folder_selection_tab(self):
        # ... (Unchanged)
        tab = ttk.Frame(self.notebook); self.notebook.add(tab, text="Monitored Folders")
        self.folder_canvas = tk.Canvas(tab, borderwidth=0, background="#ffffff")
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=self.folder_canvas.yview)
        self.scrollable_frame = ttk.Frame(self.folder_canvas, style="Row0.TFrame")
        self.scrollable_frame.bind("<Configure>", lambda e: self.folder_canvas.configure(scrollregion=self.folder_canvas.bbox("all")))
        self.folder_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.folder_canvas.configure(yscrollcommand=scrollbar.set)
        self.folder_canvas.pack(side="left", fill="both", expand=True, padx=(10,0), pady=10)
        scrollbar.pack(side="right", fill="y", padx=(0,10), pady=10)
        def _on_mousewheel(event):
            delta = 0
            if hasattr(event, 'delta') and event.delta != 0 : delta = -int(event.delta / abs(event.delta)) 
            elif event.num == 4: delta = -1 
            elif event.num == 5: delta = 1  
            if delta !=0: self.folder_canvas.yview_scroll(delta, "units")
        self.folder_canvas.bind_all("<MouseWheel>", _on_mousewheel, add='+'); self.folder_canvas.bind_all("<Button-4>", _on_mousewheel, add='+'); self.folder_canvas.bind_all("<Button-5>", _on_mousewheel, add='+')
        self.folder_entries = {}; self.folder_column_entries = {}; self.folder_skip_vars = {}; self.file_extension_entries = {}; self.folder_row_widgets = {}
        self.add_folder_header(self.scrollable_frame)


    def add_folder_header(self, parent):
        # ... (Unchanged)
        header_frame = ttk.Frame(parent, style="Header.TFrame", padding=(5,3)); header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 5)); header_frame.columnconfigure(1, weight=1)
        ttk.Label(header_frame, text="Folder Type", width=15, anchor="w", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=(5,0))
        ttk.Label(header_frame, text="Monitor Path", anchor="w", font=("Arial", 10, "bold")).grid(row=0, column=1, padx=5, sticky='w')
        ttk.Label(header_frame, text="...", width=4, anchor="center", font=("Arial", 10, "bold")).grid(row=0, column=2, padx=1)
        ttk.Label(header_frame, text="Target Column", width=20, anchor="w", font=("Arial", 10, "bold")).grid(row=0, column=3, padx=5, sticky='w')
        ttk.Label(header_frame, text="File Ext.", width=10, anchor="w", font=("Arial", 10, "bold")).grid(row=0, column=4, padx=5, sticky='w')
        ttk.Label(header_frame, text="Skip?", width=5, anchor="center", font=("Arial", 10, "bold")).grid(row=0, column=5, padx=(10,5), sticky='w')

    def add_initial_folder_rows(self):
        # ... (Unchanged)
        default_folders = ["Qinsy DB", "Naviscan", "SIS", "SSS", "SBP", "Mag", "Grad", "SVP", "SpintINS", "Video", "Cathx", "Hypack RAW", "Eiva NaviPac"]
        loaded_paths = self.parent_gui.folder_paths; all_folder_names = []; processed = set()
        for name in default_folders:
            if name not in processed: all_folder_names.append(name); processed.add(name)
        for name in loaded_paths:
            if name not in processed: all_folder_names.append(name); processed.add(name)
        for folder_name in all_folder_names:
            self.add_folder_row(folder_name=folder_name, folder_path=loaded_paths.get(folder_name, ""),
                                column_name=self.parent_gui.folder_columns.get(folder_name, folder_name),
                                extension=self.parent_gui.file_extensions.get(folder_name, ""),
                                skip=self.parent_gui.folder_skips.get(folder_name, False))
        self.master.after_idle(self.update_scroll_region)

    def add_folder_row(self, folder_name="", folder_path="", column_name="", extension="", skip=False):
        # ... (Unchanged)
        row_index = len(self.folder_row_widgets) + 1; style_name = f"Row{row_index % 2}.TFrame"
        try: self.style.configure(style_name)
        except tk.TclError: bg = "#ffffff" if row_index % 2 == 0 else "#f5f5f5"; self.style.configure(style_name, background=bg)
        row_frame = ttk.Frame(self.scrollable_frame, style=style_name, padding=(0, 2)); row_frame.grid(row=row_index, column=0, sticky="ew", pady=0); row_frame.columnconfigure(1, weight=1)
        label_style = style_name.replace("Frame", "Label")
        try: self.style.configure(label_style, background=self.style.lookup(style_name, 'background'))
        except: pass
        label = ttk.Label(row_frame, text=f"{folder_name}:", width=15, anchor='w', style=label_style); label.grid(row=0, column=0, padx=(5,0), pady=1, sticky="w")
        entry = ttk.Entry(row_frame, width=50); entry.insert(0, folder_path); entry.grid(row=0, column=1, padx=5, pady=1, sticky="ew"); ToolTip(entry, f"Path to '{folder_name}' folder.")
        def select_folder(e=entry):
            current_path = e.get(); initial = current_path if os.path.isdir(current_path) else (os.path.dirname(current_path) if current_path else "/")
            folder = filedialog.askdirectory(parent=self.master, initialdir=initial, title=f"Select Folder for {folder_name}")
            if folder: e.delete(0, tk.END); e.insert(0, folder)
        button = ttk.Button(row_frame, text="...", width=3, command=select_folder); button.grid(row=0, column=2, padx=(0,5), pady=1); ToolTip(button, "Browse for folder.")
        column_entry = ttk.Entry(row_frame, width=20); column_entry.insert(0, column_name if column_name else folder_name); column_entry.grid(row=0, column=3, padx=5, pady=1, sticky="w"); ToolTip(column_entry, f"Excel/DB column for '{folder_name}' filename.")
        extension_entry = ttk.Entry(row_frame, width=10); extension_entry.insert(0, extension); extension_entry.grid(row=0, column=4, padx=5, pady=1, sticky="w"); ToolTip(extension_entry, f"Monitor specific extension (e.g., 'svp'). Blank for any.")
        skip_var = tk.BooleanVar(value=skip); skip_checkbox = ttk.Checkbutton(row_frame, variable=skip_var); skip_checkbox.grid(row=0, column=5, padx=(15,5), pady=1, sticky="w"); ToolTip(skip_checkbox, f"Disable monitoring for '{folder_name}'.")
        self.folder_entries[folder_name] = entry; self.folder_column_entries[folder_name] = column_entry; self.file_extension_entries[folder_name] = extension_entry; self.folder_skip_vars[folder_name] = skip_var; self.folder_row_widgets[folder_name] = row_frame


    def update_scroll_region(self):
        # ... (Unchanged)
        self.scrollable_frame.update_idletasks(); self.folder_canvas.configure(scrollregion=self.folder_canvas.bbox("all"))

    def create_custom_buttons_tab_generic(self, set_num): # Made generic
        tab_title = f"Custom Buttons (Set {set_num})"
        tab = ttk.Frame(self.notebook, padding=20); self.notebook.add(tab, text=tab_title)
        self.pastel_colors = ["#FFB3BA", "#FFDFBA", "#FFFFBA", "#BAFFC9", "#BAE1FF", "#E0BBE4", "#FFC8A2", "#D4A5A5", "#A2D4AB", "#A2C4D4"]

        # Standard Button Colors (only on the first custom buttons tab for simplicity, or could be its own tab)
        if set_num == 1:
            std_color_frame = ttk.LabelFrame(tab, text="Standard Action Row Colors (Excel Output)", padding=10)
            std_color_frame.pack(pady=(0, 15), fill='x', anchor='n')
            std_header = ttk.Frame(std_color_frame, style="Header.TFrame", padding=(5,3)); std_header.pack(fill='x', pady=(0, 5))
            ttk.Label(std_header, text="Action", font=("Arial", 10, "bold"), width=15).pack(side=tk.LEFT, padx=(5,0))
            ttk.Label(std_header, text="Row Color", font=("Arial", 10, "bold"), width=25).pack(side=tk.LEFT, padx=5)
            self.standard_button_color_widgets = {}
            standard_buttons_to_color = ["Log on", "Log off", "Event", "SVP", "New Day"]
            for i, btn_name in enumerate(standard_buttons_to_color):
                style_name_std = f"Row{i % 2}.TFrame"; row_frame_std = ttk.Frame(std_color_frame, style=style_name_std, padding=(0, 2)); row_frame_std.pack(fill='x', pady=0)
                ttk.Label(row_frame_std, text=f"{btn_name}:", width=15, style=style_name_std.replace("Frame","Label")).pack(side=tk.LEFT, padx=(5,0))
                color_widget_frame_std = ttk.Frame(row_frame_std, style=style_name_std); color_widget_frame_std.pack(side=tk.LEFT, padx=5)
                initial_color_std = self.parent_gui.button_colors.get(btn_name, (None, None))[1]
                selected_color_var_std = tk.StringVar(value=initial_color_std if initial_color_std else "")
                color_display_label_std = tk.Label(color_widget_frame_std, width=4, relief="solid", borderwidth=1)
                color_display_label_std.pack(side="left", padx=(0, 5))
                try: color_display_label_std.config(background=initial_color_std if initial_color_std else 'SystemButtonFace')
                except tk.TclError: color_display_label_std.config(background='SystemButtonFace')
                def _set_color_std(color_hex, var=selected_color_var_std, label=color_display_label_std): # Closure for std
                    valid_color = None; temp_label = None
                    if color_hex:
                        try: temp_label = tk.Label(self.master); temp_label.config(background=color_hex); valid_color = color_hex
                        except tk.TclError: valid_color = None
                        finally: 
                            if temp_label is not None: 
                                try: temp_label.destroy()
                                except tk.TclError: pass
                    var.set(valid_color if valid_color else "")
                    try: label.config(background=valid_color if valid_color else 'SystemButtonFace')
                    except tk.TclError: label.config(background='SystemButtonFace')
                def _choose_color_std(var=selected_color_var_std, label=color_display_label_std, name=btn_name): # Closure for std
                    current_color = var.get()
                    color_code = colorchooser.askcolor(color=current_color if current_color else None, title=f"Row Color for {name}", parent=self.master)
                    if color_code and color_code[1]: _set_color_std(color_code[1], var, label)
                clear_btn_std = ttk.Button(color_widget_frame_std, text="X", width=2, style="Toolbutton", command=lambda v=selected_color_var_std, l=color_display_label_std: _set_color_std(None, v, l))
                clear_btn_std.pack(side="left", padx=1); ToolTip(clear_btn_std, f"Clear row color for {btn_name}.")
                presets_frame_std = ttk.Frame(color_widget_frame_std, style=style_name_std); presets_frame_std.pack(side="left", padx=(2, 2))
                for p_color in self.pastel_colors[:5]:
                    try:
                        b = tk.Button(presets_frame_std, bg=p_color, width=1, height=1, relief="raised", bd=1, command=lambda c=p_color, v=selected_color_var_std, l=color_display_label_std: _set_color_std(c, v, l))
                        b.pack(side=tk.LEFT, padx=1)
                    except tk.TclError: pass
                choose_btn_std = ttk.Button(color_widget_frame_std, text="...", width=3, style="Toolbutton", command=lambda v=selected_color_var_std, l=color_display_label_std, n=btn_name: _choose_color_std(v, l, n))
                choose_btn_std.pack(side="left", padx=1); ToolTip(choose_btn_std, f"Choose custom row color for {btn_name}.")
                self.standard_button_color_widgets[btn_name] = (selected_color_var_std, color_display_label_std)

        custom_frame = ttk.LabelFrame(tab, text=f"Custom Buttons Configuration (Set {set_num})", padding=10)
        custom_frame.pack(pady=10, fill='both', expand=True)
        num_buttons_frame = ttk.Frame(custom_frame); num_buttons_frame.pack(pady=5, anchor='w')
        ttk.Label(num_buttons_frame, text=f"Number of Buttons (Set {set_num}, 0-10):").pack(side='left', padx=5)
        
        num_buttons_entry_ref = ttk.Entry(num_buttons_frame, width=5)
        num_buttons_entry_ref.pack(side='left', padx=5)
        ToolTip(num_buttons_entry_ref, f"Number of custom buttons for Set {set_num} (max 10).")
        
        update_btn = ttk.Button(num_buttons_frame, text="Update Count", command=lambda s=set_num: self.update_num_custom_buttons(s))
        update_btn.pack(side='left', padx=5); ToolTip(update_btn, "Update list below.")

        if set_num == 1: self.num_buttons_entry_s1 = num_buttons_entry_ref
        else: self.num_buttons_entry_s2 = num_buttons_entry_ref

        header_frame = ttk.Frame(custom_frame, style="Header.TFrame", padding=(5,3)); header_frame.pack(fill='x', pady=(15,5))
        ttk.Label(header_frame, text="Btn #", font=("Arial", 10, "bold"), width=7).pack(side=tk.LEFT, padx=(5,0))
        ttk.Label(header_frame, text="Button Text", font=("Arial", 10, "bold"), width=25).pack(side=tk.LEFT, padx=5)
        ttk.Label(header_frame, text="Event Text (for Log)", font=("Arial", 10, "bold"), width=35).pack(side=tk.LEFT, padx=5, expand=True, fill='x')
        ttk.Label(header_frame, text="Button/Row Color", font=("Arial", 10, "bold"), width=25).pack(side=tk.LEFT, padx=5)

        entries_frame_ref = ttk.Frame(custom_frame); entries_frame_ref.pack(pady=0, fill='both', expand=True)
        if set_num == 1:
            self.custom_button_entries_frame_s1 = entries_frame_ref
            self.custom_button_widgets_s1 = []
        else:
            self.custom_button_entries_frame_s2 = entries_frame_ref
            self.custom_button_widgets_s2 = []


    def update_num_custom_buttons(self, set_num): # Made generic
        num_entry_ref = self.num_buttons_entry_s1 if set_num == 1 else self.num_buttons_entry_s2
        num_buttons_attr = "num_custom_buttons" if set_num == 1 else "num_custom_buttons_set2"
        configs_attr = "custom_button_configs" if set_num == 1 else "custom_button_configs_set2"
        
        try:
            num_buttons_val = int(num_entry_ref.get())
            if not (0 <= num_buttons_val <= 10): raise ValueError("Number must be 0-10")
            
            current_num = getattr(self.parent_gui, num_buttons_attr)
            if current_num != num_buttons_val:
                print(f"Updating number of custom buttons for Set {set_num} to {num_buttons_val}")
                setattr(self.parent_gui, num_buttons_attr, num_buttons_val)
                current_configs_list = getattr(self.parent_gui, configs_attr)
                
                if num_buttons_val < len(current_configs_list):
                    setattr(self.parent_gui, configs_attr, current_configs_list[:num_buttons_val])
                else:
                    while len(current_configs_list) < num_buttons_val:
                        idx = len(current_configs_list) + 1
                        default_text = f"Custom {idx}" if set_num == 1 else f"CustomS2 {idx}"
                        current_configs_list.append({"text": default_text, "event_text": f"{default_text} Event"})
                self.recreate_custom_button_settings(set_num)
        except ValueError as e:
            messagebox.showerror("Invalid Number", f"Enter whole number 0-10. Error: {e}", parent=self.master)
            num_entry_ref.delete(0, tk.END); num_entry_ref.insert(0, str(getattr(self.parent_gui, num_buttons_attr)))

    def recreate_custom_button_settings(self, set_num): # Made generic
        entries_frame_ref = self.custom_button_entries_frame_s1 if set_num == 1 else self.custom_button_entries_frame_s2
        widgets_list_ref = self.custom_button_widgets_s1 if set_num == 1 else self.custom_button_widgets_s2
        num_buttons_val = self.parent_gui.num_custom_buttons if set_num == 1 else self.parent_gui.num_custom_buttons_set2
        configs_list = self.parent_gui.custom_button_configs if set_num == 1 else self.parent_gui.custom_button_configs_set2

        for widget in entries_frame_ref.winfo_children(): widget.destroy()
        widgets_list_ref.clear()

        for i in range(num_buttons_val):
            config = configs_list[i] if i < len(configs_list) else {} 
            default_btn_text = f"Custom {i+1}" if set_num == 1 else f"CustomS2 {i+1}"
            initial_text = config.get("text", default_btn_text)
            initial_event = config.get("event_text", f"{initial_text} Event")
            _, initial_color = self.parent_gui.button_colors.get(initial_text, (None, None)) # Uses shared button_colors

            style_name = f"Row{i % 2}.TFrame"; row_frame = ttk.Frame(entries_frame_ref, style=style_name, padding=(0, 2)); row_frame.pack(fill='x', pady=0)
            ttk.Label(row_frame, text=f"{i+1}", width=7, style=style_name.replace("Frame","Label")).pack(side=tk.LEFT, padx=(5,0))
            text_entry = ttk.Entry(row_frame, width=25); text_entry.insert(0, initial_text); text_entry.pack(side=tk.LEFT, padx=5); ToolTip(text_entry, "Button display text.")
            event_entry = ttk.Entry(row_frame); event_entry.insert(0, initial_event); event_entry.pack(side=tk.LEFT, padx=5, fill='x', expand=True); ToolTip(event_entry, "Text for 'Event' column.")
            
            color_frame = ttk.Frame(row_frame, style=style_name); color_frame.pack(side=tk.LEFT, padx=5)
            selected_color_var = tk.StringVar(value=initial_color if initial_color else "")
            color_display_label = tk.Label(color_frame, width=4, relief="solid", borderwidth=1)
            color_display_label.pack(side="left", padx=(0, 5))
            try: color_display_label.config(background=initial_color if initial_color else 'SystemButtonFace')
            except tk.TclError: color_display_label.config(background='SystemButtonFace')
            
            def _set_color_custom(color_hex, var=selected_color_var, label=color_display_label): # Unique name for custom
                valid_color=None; temp_label=None
                if color_hex:
                    try:temp_label=tk.Label(self.master);temp_label.config(background=color_hex);valid_color=color_hex
                    except tk.TclError:valid_color=None
                    finally:
                        if temp_label:
                            try:temp_label.destroy()
                            except:pass
                var.set(valid_color if valid_color else "")
                try:label.config(background=valid_color if valid_color else 'SystemButtonFace')
                except tk.TclError:label.config(background='SystemButtonFace')

            def _choose_color_custom(var=selected_color_var, label=color_display_label, txt_widget=text_entry, current_i=i, current_set_num=set_num): # Unique name for custom
                current_color = var.get()
                btn_name = txt_widget.get().strip() or (f"Button {current_i+1}" if current_set_num==1 else f"ButtonS2 {current_i+1}")
                color_code = colorchooser.askcolor(color=current_color if current_color else None, title=f"Color for {btn_name}", parent=self.master)
                if color_code and color_code[1]: _set_color_custom(color_code[1], var, label)

            clear_btn = ttk.Button(color_frame, text="X", width=2, style="Toolbutton", command=lambda v=selected_color_var, l=color_display_label: _set_color_custom(None, v, l))
            clear_btn.pack(side="left", padx=1); ToolTip(clear_btn, "Clear color.")
            presets_frame = ttk.Frame(color_frame, style=style_name); presets_frame.pack(side="left", padx=(2, 2))
            for p_color in self.pastel_colors[:5]:
                try:
                    b = tk.Button(presets_frame, bg=p_color, width=1, height=1, relief="raised", bd=1, command=lambda c=p_color, v=selected_color_var, l=color_display_label: _set_color_custom(c, v, l))
                    b.pack(side=tk.LEFT, padx=1)
                except tk.TclError: pass
            choose_btn = ttk.Button(color_frame, text="...", width=3, style="Toolbutton", command=lambda v=selected_color_var, l=color_display_label, t=text_entry, ci=i, csn=set_num: _choose_color_custom(v,l,t,ci,csn))
            choose_btn.pack(side="left", padx=1); ToolTip(choose_btn, "Choose custom color.")
            widgets_list_ref.append( (text_entry, event_entry, selected_color_var, color_display_label) )

    def create_sqlite_tab(self):
        # ... (Unchanged)
        tab = ttk.Frame(self.notebook, padding=20); self.notebook.add(tab, text="SQLite Log")
        enable_frame = ttk.Frame(tab); enable_frame.pack(fill='x', pady=(0, 15))
        self.sqlite_enabled_var = tk.BooleanVar()
        enable_check = ttk.Checkbutton(enable_frame, text="Enable SQLite Database Logging", variable=self.sqlite_enabled_var, style="Large.TCheckbutton"); enable_check.pack(side=tk.LEFT, pady=(5, 10)); ToolTip(enable_check, "Enable SQLite logging.")
        config_frame = ttk.LabelFrame(tab, text="SQLite Configuration", padding=15); config_frame.pack(fill='x'); config_frame.columnconfigure(1, weight=1)
        ttk.Label(config_frame, text="Database File (.db):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.sqlite_db_path_entry = ttk.Entry(config_frame, width=70); self.sqlite_db_path_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew"); ToolTip(self.sqlite_db_path_entry, "Path to SQLite DB file.")
        db_browse_btn = ttk.Button(config_frame, text="Browse/Create...", command=self.select_sqlite_file); db_browse_btn.grid(row=0, column=2, padx=5, pady=5); ToolTip(db_browse_btn, "Browse/specify SQLite file.")
        ttk.Label(config_frame, text="Table Name:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.sqlite_table_entry = ttk.Entry(config_frame, width=40); self.sqlite_table_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w"); ToolTip(self.sqlite_table_entry, "Table name (e.g., 'EventLog'). Must exist.")
        test_button = ttk.Button(config_frame, text="Test Connection & Table", command=self.test_sqlite_connection); test_button.grid(row=2, column=1, padx=5, pady=15, sticky="w"); ToolTip(test_button, "Verify DB connection and table existence.")
        self.test_result_label = ttk.Label(config_frame, text="", font=("Arial", 9), wraplength=500); self.test_result_label.grid(row=3, column=0, columnspan=3, padx=5, pady=2, sticky="w")


    def select_sqlite_file(self):
        # ... (Unchanged)
        filetypes = [("SQLite Database", "*.db"), ("SQLite Database", "*.sqlite"), ("SQLite3 Database", "*.sqlite3"), ("All Files", "*.*")]
        current_path = self.sqlite_db_path_entry.get(); initial_dir = os.path.dirname(current_path) if current_path else "."
        filepath = filedialog.asksaveasfilename(parent=self.master, title="Select or Create SQLite Database File", initialdir=initial_dir, initialfile="DataLoggerLog.db", filetypes=filetypes, defaultextension=".db")
        if filepath: self.sqlite_db_path_entry.delete(0, tk.END); self.sqlite_db_path_entry.insert(0, filepath); print(f"SQLite DB file specified: {filepath}");
        if hasattr(self, 'test_result_label'): self.test_result_label.config(text="")


    def test_sqlite_connection(self):
        # ... (Unchanged)
        db_path = self.sqlite_db_path_entry.get().strip(); table_name = self.sqlite_table_entry.get().strip() or "EventLog"
        if not db_path: self.test_result_label.config(text="❌ Error: Database path is empty.", foreground="red"); return
        conn = None; result_text = ""; result_color = "red"
        try:
            print(f"Testing SQLite connection to: {db_path}"); conn = sqlite3.connect(db_path, timeout=3); cursor = conn.cursor(); print("Connection successful.")
            result_text = f"✔️ Connection to '{os.path.basename(db_path)}' successful.\n"
            try:
                cursor.execute(f"SELECT 1 FROM [{table_name}] LIMIT 1;"); print(f"Table '{table_name}' exists.")
                result_text += f"✔️ Table '{table_name}' found."; result_color = "green"
            except sqlite3.OperationalError as e_table:
                if "no such table" in str(e_table).lower(): print(f"Table '{table_name}' not found."); result_text += f"⚠️ Warning: Table '{table_name}' not found."; result_color = "#E67E00"
                else: raise e_table
        except sqlite3.Error as e: print(f"SQLite Test Error: {e}"); result_text = f"❌ Error: {e}"; result_color = "red"
        except Exception as e: print(f"Unexpected Test Error: {e}"); result_text = f"❌ Unexpected Error: {e}"; result_color = "red"
        finally:
            if conn: conn.close(); print("Connection closed.")
            self.test_result_label.config(text=result_text, foreground=result_color)
            self.master.after(15000, lambda: self.test_result_label.config(text=""))


    def save_settings(self):
        # General
        self.parent_gui.log_file_path = self.log_file_entry.get().strip()
        self.parent_gui.txt_folder_path = self.txt_folder_entry.get().strip()
        self.parent_gui.txt_folder_path_set2 = self.txt_folder_entry_set2.get().strip() # NEW

        # TXT Columns Set 1
        parent_txt_cols_s1 = {}; parent_txt_skips_s1 = {}
        for field, entry in self.txt_field_column_widgets_s1.items(): parent_txt_cols_s1[field] = entry.get().strip() or field
        for field, var in self.txt_field_skip_vars_s1.items(): parent_txt_skips_s1[field] = var.get()
        self.parent_gui.txt_field_columns = parent_txt_cols_s1; self.parent_gui.txt_field_skips = parent_txt_skips_s1
        
        # TXT Columns Set 2 - NEW
        parent_txt_cols_s2 = {}; parent_txt_skips_s2 = {}
        for field, entry in self.txt_field_column_widgets_s2.items(): parent_txt_cols_s2[field] = entry.get().strip() or field
        for field, var in self.txt_field_skip_vars_s2.items(): parent_txt_skips_s2[field] = var.get()
        self.parent_gui.txt_field_columns_set2 = parent_txt_cols_s2; self.parent_gui.txt_field_skips_set2 = parent_txt_skips_s2

        # Monitored Folders (global)
        parent_folder_paths = {}; parent_folder_cols = {}; parent_folder_exts = {}; parent_folder_skips = {}
        # ... (rest of folder saving logic unchanged) ...
        for folder_name, entry_widget in self.folder_entries.items():
            folder_path = entry_widget.get().strip()
            if folder_path:
                parent_folder_paths[folder_name] = folder_path; col_entry = self.folder_column_entries.get(folder_name); ext_entry = self.file_extension_entries.get(folder_name); skip_var = self.folder_skip_vars.get(folder_name)
                parent_folder_cols[folder_name] = col_entry.get().strip() if col_entry and col_entry.get().strip() else folder_name
                parent_folder_exts[folder_name] = ext_entry.get().strip().lstrip('.') if ext_entry else ""
                parent_folder_skips[folder_name] = skip_var.get() if skip_var else False
            else:
                for d in [self.parent_gui.folder_paths, self.parent_gui.folder_columns, self.parent_gui.file_extensions, self.parent_gui.folder_skips]:
                    d.pop(folder_name, None)
        self.parent_gui.folder_paths = parent_folder_paths; self.parent_gui.folder_columns = parent_folder_cols; self.parent_gui.file_extensions = parent_folder_exts; self.parent_gui.folder_skips = parent_folder_skips


        # --- Save Button Colors/Configs (for both sets) ---
        new_button_colors_combined = {} # Single dict for all button colors

        # Standard button row colors (from Set 1 tab)
        standard_buttons_to_color = ["Log on", "Log off", "Event", "SVP", "New Day"]
        for std_btn_name in standard_buttons_to_color:
            if std_btn_name in self.standard_button_color_widgets: # Check if widgets exist
                color_var, _ = self.standard_button_color_widgets[std_btn_name]
                color_hex = color_var.get()
                new_button_colors_combined[std_btn_name] = (None, color_hex if color_hex else None)
            elif std_btn_name in self.parent_gui.button_colors: # Preserve old if widget missing
                 new_button_colors_combined[std_btn_name] = self.parent_gui.button_colors[std_btn_name]

        # Custom Buttons Set 1
        parent_custom_configs_s1 = []
        self.parent_gui.num_custom_buttons = int(self.num_buttons_entry_s1.get())
        for i, (text_widget, event_widget, color_var, _) in enumerate(self.custom_button_widgets_s1):
            if i >= self.parent_gui.num_custom_buttons: break # only save up to the count
            text = text_widget.get().strip(); event_text = event_widget.get().strip(); color_hex = color_var.get()
            default_text = f"Custom {i + 1}"; final_text = text if text else default_text; final_event_text = event_text if event_text else f"{final_text} Triggered"
            parent_custom_configs_s1.append({"text": final_text, "event_text": final_event_text})
            new_button_colors_combined[final_text] = (None, color_hex if color_hex else None)
        self.parent_gui.custom_button_configs = parent_custom_configs_s1
        
        # Custom Buttons Set 2 - NEW
        parent_custom_configs_s2 = []
        self.parent_gui.num_custom_buttons_set2 = int(self.num_buttons_entry_s2.get())
        for i, (text_widget, event_widget, color_var, _) in enumerate(self.custom_button_widgets_s2):
            if i >= self.parent_gui.num_custom_buttons_set2: break
            text = text_widget.get().strip(); event_text = event_widget.get().strip(); color_hex = color_var.get()
            default_text = f"CustomS2 {i + 1}"; final_text = text if text else default_text; final_event_text = event_text if event_text else f"{final_text} Triggered"
            parent_custom_configs_s2.append({"text": final_text, "event_text": final_event_text})
            new_button_colors_combined[final_text] = (None, color_hex if color_hex else None) # Add to combined color dict
        self.parent_gui.custom_button_configs_set2 = parent_custom_configs_s2
        
        self.parent_gui.button_colors = new_button_colors_combined # Assign the combined dict

        # SQLite settings
        self.parent_gui.sqlite_enabled = self.sqlite_enabled_var.get()
        self.parent_gui.sqlite_db_path = self.sqlite_db_path_entry.get().strip()
        self.parent_gui.sqlite_table = self.sqlite_table_entry.get().strip() or "EventLog"

        # Apply changes by calling parent's save and update methods
        self.parent_gui.save_settings(); 
        self.parent_gui.update_custom_buttons(); 
        self.parent_gui.start_monitoring(); 
        self.parent_gui.update_db_indicator()

    def load_settings(self):
        # General
        self.log_file_entry.delete(0, tk.END); self.log_file_entry.insert(0, self.parent_gui.log_file_path or "")
        self.txt_folder_entry.delete(0, tk.END); self.txt_folder_entry.insert(0, self.parent_gui.txt_folder_path or "")
        self.txt_folder_entry_set2.delete(0, tk.END); self.txt_folder_entry_set2.insert(0, self.parent_gui.txt_folder_path_set2 or "") # NEW

        # TXT Columns Set 1
        for field, entry in self.txt_field_column_widgets_s1.items(): entry.delete(0, tk.END); default_val = "Event" if field == "Event" else field; entry.insert(0, self.parent_gui.txt_field_columns.get(field, default_val))
        for field, var in self.txt_field_skip_vars_s1.items(): var.set(self.parent_gui.txt_field_skips.get(field, False))
        
        # TXT Columns Set 2 - NEW
        for field, entry in self.txt_field_column_widgets_s2.items(): entry.delete(0, tk.END); default_val = "Event" if field == "Event" else field; entry.insert(0, self.parent_gui.txt_field_columns_set2.get(field, default_val))
        for field, var in self.txt_field_skip_vars_s2.items(): var.set(self.parent_gui.txt_field_skips_set2.get(field, False))

        # Monitored Folders
        for name, frame in list(self.folder_row_widgets.items()):
            if frame and frame.winfo_exists(): frame.destroy()
        self.folder_row_widgets.clear(); self.folder_entries.clear(); self.folder_column_entries.clear(); self.file_extension_entries.clear(); self.folder_skip_vars.clear()
        self.add_initial_folder_rows(); self.master.after_idle(self.update_scroll_region)
        
        # Custom Buttons Set 1
        self.num_buttons_entry_s1.delete(0, tk.END); self.num_buttons_entry_s1.insert(0, str(self.parent_gui.num_custom_buttons))
        self.recreate_custom_button_settings(set_num=1) 
        
        # Custom Buttons Set 2 - NEW
        self.num_buttons_entry_s2.delete(0, tk.END); self.num_buttons_entry_s2.insert(0, str(self.parent_gui.num_custom_buttons_set2))
        self.recreate_custom_button_settings(set_num=2)

        # Standard Button Colors (loaded within recreate_custom_button_settings for set_num=1 implicitly if logic is there, or needs explicit load)
        # Explicitly load standard button colors if they are on the first tab
        if hasattr(self, 'standard_button_color_widgets'):
            standard_buttons_to_color = ["Log on", "Log off", "Event", "SVP", "New Day"]
            for btn_name in standard_buttons_to_color:
                if btn_name in self.standard_button_color_widgets:
                    color_var, display_label = self.standard_button_color_widgets[btn_name]
                    loaded_color_hex = self.parent_gui.button_colors.get(btn_name, (None, None))[1]
                    # Define local helper for this scope
                    def _set_color_load_std(color_hex, var=color_var, label=display_label): 
                        valid_color = None; temp_label = None
                        if color_hex:
                            try: temp_label = tk.Label(self.master); temp_label.config(background=color_hex); valid_color = color_hex
                            except tk.TclError: valid_color = None
                            finally: 
                                if temp_label is not None: 
                                    try: temp_label.destroy()
                                    except tk.TclError: pass
                        var.set(valid_color if valid_color else "")
                        try: label.config(background=valid_color if valid_color else 'SystemButtonFace')
                        except tk.TclError: label.config(background='SystemButtonFace')
                    _set_color_load_std(loaded_color_hex)
        
        # SQLite
        self.sqlite_enabled_var.set(self.parent_gui.sqlite_enabled); self.sqlite_db_path_entry.delete(0, tk.END); self.sqlite_db_path_entry.insert(0, self.parent_gui.sqlite_db_path or "")
        self.sqlite_table_entry.delete(0, tk.END); self.sqlite_table_entry.insert(0, self.parent_gui.sqlite_table or "EventLog")
        if hasattr(self, 'test_result_label'): self.test_result_label.config(text="")


# --- Main Execution ---
if __name__ == "__main__":
    root = tk.Tk()
    gui = DataLoggerGUI(root) 

    def on_closing():
        print("Closing application requested...")
        print("Shutting down monitors...")
        active_monitors = list(gui.monitors.items()) 
        if not active_monitors: print("No monitors were active.")
        else:
            for name, monitor_observer in active_monitors:
                try:
                    if monitor_observer.is_alive(): monitor_observer.stop(); print(f"Stopped monitor '{name}'.")
                except Exception as e: print(f"Error stopping monitor '{name}': {e}")
            for name, monitor_observer in active_monitors:
                try:
                    if monitor_observer.is_alive():
                        monitor_observer.join(timeout=0.5) 
                        if monitor_observer.is_alive(): print(f"Warning: Monitor thread '{name}' did not stop gracefully.")
                except Exception as e: print(f"Error joining monitor thread '{name}': {e}")
                finally: 
                    if name in gui.monitors: del gui.monitors[name] 
        print("Monitors shut down. Exiting.")
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()
