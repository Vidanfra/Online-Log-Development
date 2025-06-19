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
        # Ensure tooltip hides if the widget is destroyed
        # self.widget.bind("<Destroy>", self.on_leave, add='+') # Might cause issues if triggered too often

    def on_enter(self, event=None):
        # When mouse enters, cancel any scheduled hide and schedule a show
        self.cancel_scheduled_hide()
        self.schedule_show()

    def on_leave(self, event=None):
        # When mouse leaves, cancel any scheduled show and schedule a hide
        self.cancel_scheduled_show()
        self.schedule_hide()

    def schedule_show(self):
        # Cancel previous show timer if any
        self.cancel_scheduled_show()
        # Schedule the tooltip to appear after delay
        self.show_id = self.widget.after(self.show_delay, self.show_tooltip)

    def schedule_hide(self):
        # Cancel previous hide timer if any
        self.cancel_scheduled_hide()
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
                        if self.gui_instance and hasattr(self.gui_instance, 'master') and self.gui_instance.master.winfo_exists():
                           self.gui_instance.master.after(0, self.gui_instance.update_status, f"Latest {self.folder_name} file: {file_name}")
        except Exception as e:
            print(f"Error updating cache for {self.folder_name}: {e}")
            traceback.print_exc()


# --- Custom Button Edit/Add Dialog ---
class CustomButtonEditorDialog(Toplevel):
    def __init__(self, master, parent_gui, button_index=None): # button_index is None for "Add" mode
        super().__init__(master)
        self.parent_gui = parent_gui
        self.button_index = button_index
        self.is_add_mode = (button_index is None)

        self.title("Edit Custom Button" if not self.is_add_mode else "Add Custom Button")
        self.geometry("450x280") 
        self.transient(master)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self.cancel)

        # Initial values
        self.initial_button_text = ""
        if not self.is_add_mode:
            config = self.parent_gui.custom_button_configs[self.button_index]
            self.initial_button_text = config.get("text", "")
            initial_event_text = config.get("event_text", "")
            _, initial_color_hex = self.parent_gui.button_colors.get(self.initial_button_text, (None, None))
        else:
            initial_event_text = ""
            initial_color_hex = None 

        # --- Widgets ---
        main_frame = ttk.Frame(self, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="Button Text:").grid(row=0, column=0, padx=5, pady=8, sticky="w")
        self.text_entry = ttk.Entry(main_frame, width=40)
        self.text_entry.grid(row=0, column=1, padx=5, pady=8, sticky="ew")
        self.text_entry.insert(0, self.initial_button_text)

        ttk.Label(main_frame, text="Event Text (for Log):").grid(row=1, column=0, padx=5, pady=8, sticky="w")
        self.event_entry = ttk.Entry(main_frame, width=40)
        self.event_entry.grid(row=1, column=1, padx=5, pady=8, sticky="ew")
        self.event_entry.insert(0, initial_event_text)

        ttk.Label(main_frame, text="Button/Row Color:").grid(row=2, column=0, padx=5, pady=8, sticky="w")
        self.color_var = tk.StringVar(value=initial_color_hex if initial_color_hex else "")
        
        color_picker_frame = ttk.Frame(main_frame)
        color_picker_frame.grid(row=2, column=1, padx=5, pady=8, sticky="ew")

        self.color_display_label = tk.Label(color_picker_frame, width=4, relief="solid", borderwidth=1)
        self.color_display_label.pack(side="left", padx=(0, 5))
        self._update_color_display() 

        clear_btn = ttk.Button(color_picker_frame, text="X", width=2, style="Toolbutton", command=lambda: self._set_color_direct(None))
        clear_btn.pack(side="left", padx=1)
        ToolTip(clear_btn, "Clear color (uses default style).")

        pastel_colors = ["#FFB3BA", "#FFDFBA", "#FFFFBA", "#BAFFC9", "#BAE1FF"] 
        presets_frame = ttk.Frame(color_picker_frame)
        presets_frame.pack(side="left", padx=(2,2))
        for p_color in pastel_colors:
            try:
                b = tk.Button(presets_frame, bg=p_color, width=1, height=1, relief="raised", bd=1,
                              command=lambda c=p_color: self._set_color_direct(c))
                b.pack(side=tk.LEFT, padx=1)
            except tk.TclError: pass 

        choose_btn = ttk.Button(color_picker_frame, text="...", width=3, style="Toolbutton", command=self._choose_color_dialog)
        choose_btn.pack(side="left", padx=1)
        ToolTip(choose_btn, "Choose a custom color from the palette.")

        # --- Buttons ---
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=2, pady=(25,0), sticky="e") 
        ttk.Button(button_frame, text="Save", command=self.save, style="Accent.TButton").pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.cancel).pack(side=tk.LEFT)

        main_frame.columnconfigure(1, weight=1)
        self.text_entry.focus_set()
        self.bind("<Return>", lambda event: self.save())
        self.bind("<Escape>", lambda event: self.cancel())

    def _update_color_display(self):
        color_hex = self.color_var.get()
        try:
            self.color_display_label.config(background=color_hex if color_hex else self.winfo_toplevel().option_get('background', '.')) 
        except tk.TclError:
            self.color_display_label.config(background=self.winfo_toplevel().option_get('background', '.'))


    def _set_color_direct(self, color_hex):
        valid_color = None
        if color_hex:
            temp_label = None
            try:
                temp_label = tk.Label(self) 
                temp_label.config(background=color_hex)
                valid_color = color_hex
            except tk.TclError:
                print(f"Warning: Invalid color code '{color_hex}' in dialog.")
                valid_color = None
            finally:
                if temp_label:
                    try: temp_label.destroy()
                    except tk.TclError: pass
        self.color_var.set(valid_color if valid_color else "")
        self._update_color_display()

    def _choose_color_dialog(self):
        current_color = self.color_var.get()
        color_code = colorchooser.askcolor(color=current_color if current_color else None,
                                           title="Choose Button/Row Color", parent=self) 
        if color_code and color_code[1]:
            self._set_color_direct(color_code[1])

    def save(self):
        new_text = self.text_entry.get().strip()
        new_event_text = self.event_entry.get().strip()
        new_color_hex = self.color_var.get() if self.color_var.get() else None

        if not new_text:
            messagebox.showerror("Error", "Button text cannot be empty.", parent=self)
            self.text_entry.focus_set()
            return

        if self.is_add_mode:
            if any(btn_cfg.get("text") == new_text for btn_cfg in self.parent_gui.custom_button_configs):
                 if not messagebox.askyesno("Warning", f"A custom button with text '{new_text}' already exists. This may cause unexpected behavior if colors are also duplicated. Continue?", parent=self, icon='warning'):
                    self.text_entry.focus_set()
                    return

            if len(self.parent_gui.custom_button_configs) >= 10: 
                messagebox.showerror("Error", "Maximum number of custom buttons (10) reached.", parent=self)
                return

            self.parent_gui.num_custom_buttons += 1
            self.parent_gui.custom_button_configs.append({
                "text": new_text,
                "event_text": new_event_text if new_event_text else f"{new_text} Triggered"
            })
            self.parent_gui.button_colors[new_text] = (None, new_color_hex)
        else: 
            old_config = self.parent_gui.custom_button_configs[self.button_index]
            old_text = old_config.get("text")

            if new_text != old_text and any(
                i != self.button_index and btn_cfg.get("text") == new_text
                for i, btn_cfg in enumerate(self.parent_gui.custom_button_configs)
            ):
                if not messagebox.askyesno("Warning", f"Another custom button with text '{new_text}' already exists. This may cause unexpected behavior. Continue?", parent=self, icon='warning'):
                    self.text_entry.focus_set()
                    return
            
            if old_text and old_text != new_text and old_text in self.parent_gui.button_colors:
                del self.parent_gui.button_colors[old_text]
            
            self.parent_gui.custom_button_configs[self.button_index] = {
                "text": new_text,
                "event_text": new_event_text if new_event_text else f"{new_text} Triggered"
            }
            self.parent_gui.button_colors[new_text] = (None, new_color_hex)

        self.parent_gui.update_custom_buttons() 
        self.parent_gui.save_settings() 
        self.destroy()

    def cancel(self):
        self.destroy()


# --- Main Application GUI Class ---
class DataLoggerGUI:
    def __init__(self, master):
        self.master = master
        master.title("Data Acquisition Logger (SQLite Mode)")
        master.geometry("560x580") # Slightly increased default size
        master.minsize(480, 450) # Slightly increased min size
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

        self.button_frame = ttk.Frame(self.main_frame, padding="5") # Reduced padding for button_frame itself
        self.button_frame.grid(row=0, column=0, sticky=(tk.N, tk.W, tk.E, tk.S))
        
        self.create_main_buttons(self.button_frame) 

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
            elif 'clam' in available_themes: self.style.theme_use('clam')
            else: self.style.theme_use(self.style.theme_names()[0] if self.style.theme_names() else "default")
        except tk.TclError:
            self.style.theme_use("clam") 

        self.style.configure("TFrame", background="#f0f0f0")
        self.style.configure("TLabel", background="#f0f0f0", font=("Arial", 10))
        self.style.configure("TButton", font=("Arial", 10, "bold"), padding=8) # Increased button padding
        self.style.configure("TEntry", font=("Arial", 10), padding=6)
        self.style.configure("StatusBar.TLabel", background="#e0e0e0", font=("Arial", 9), relief=tk.SUNKEN, padding=(5, 3)) # Increased vertical padding
        self.style.configure("Header.TFrame", background="#dcdcdc")
        self.style.configure("Row0.TFrame", background="#ffffff")
        self.style.configure("Row1.TFrame", background="#f5f5f5")
        self.style.configure("TLabelframe", background="#f0f0f0", padding=8) # Increased LabelFrame padding
        self.style.configure("TLabelframe.Label", background="#f0f0f0", font=("Arial", 10, "bold"))
        self.style.configure("Large.TCheckbutton", font=("Arial", 11))
        self.style.configure("Toolbutton", padding=3) 
        self.style.configure("Accent.TButton", font=("Arial", 10, "bold"), foreground="white", background="#0078D4")

        self.style.map("TButton",
                       foreground=[('pressed', 'darkblue'), ('active', 'blue'), ('disabled', '#999999')],
                       background=[('pressed', '!disabled', '#c0c0c0'), ('active', '#e0e0e0')]
                       )

    def init_variables(self):
        self.log_file_path = None
        self.txt_folder_path = None
        self.txt_file_path = None
        self.txt_field_columns = {"Event": "Event"} 
        self.txt_field_skips = {}
        self.folder_paths = {}
        self.folder_columns = {}
        self.file_extensions = {}
        self.folder_skips = {}
        self.monitors = {}
        self.num_custom_buttons = 3 
        self.custom_button_configs = [ 
            {"text": "Custom Event 1", "event_text": "Custom Event 1 Triggered"},
            {"text": "Custom Event 2", "event_text": "Custom Event 2 Triggered"},
            {"text": "Custom Event 3", "event_text": "Custom Event 3 Triggered"},
        ]
        self.custom_buttons_widgets_ref = [] 

        self.button_colors = {
            "Log on": (None, "#90EE90"), "Log off": (None, "#FFB6C1"),
            "Event": (None, "#FFFFE0"), "SVP": (None, "#ADD8E6"),
            "New Day": (None, "#FFFF99")
        }
        for i in range(10): 
            default_text = f"Custom Event {i+1}"
            if default_text not in self.button_colors:
                 self.button_colors[default_text] = (None, None)


        self.sqlite_enabled = False
        self.sqlite_db_path = None
        self.sqlite_table = "EventLog"

        self.status_var = tk.StringVar()
        self.monitor_status_label = None
        self.db_status_label = None
        self.settings_window_instance = None

    def create_main_buttons(self, parent_frame):
        for widget in parent_frame.winfo_children(): widget.destroy()
        self.custom_buttons_widgets_ref = [] 

        logging_frame = ttk.LabelFrame(parent_frame, text="Logging Actions", padding=10)
        logging_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew") # Added pady
        logging_frame.columnconfigure(0, weight=1)

        self.custom_frame = ttk.LabelFrame(parent_frame, text="Custom Events", padding=10) 
        self.custom_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=(10, 5), sticky="nsew") # Increased pady
        num_custom_cols = 2 
        for i in range(num_custom_cols):
            self.custom_frame.columnconfigure(i, weight=1)
        
        self.custom_frame.bind("<Button-3>", self.show_custom_frame_context_menu)


        other_frame = ttk.LabelFrame(parent_frame, text="Other Actions", padding=10)
        other_frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew") # Added pady
        other_frame.columnconfigure(0, weight=1)

        parent_frame.columnconfigure(0, weight=1); parent_frame.columnconfigure(1, weight=1)
        parent_frame.rowconfigure(0, weight=0); parent_frame.rowconfigure(1, weight=1)

        all_buttons_data = {
            "Log on":  {"command_ref": self.log_event, "frame": logging_frame, "tooltip": "Record a 'Log on' marker."},
            "Log off": {"command_ref": self.log_event, "frame": logging_frame, "tooltip": "Record a 'Log off' marker."},
            "Event":   {"command_ref": self.log_event, "frame": logging_frame, "tooltip": "Record data from TXT."},
            "SVP":     {"command_ref": self.apply_svp, "frame": other_frame, "tooltip": "Record data and insert SVP filename."},
            "New Day": {"command_ref": self.log_new_day, "frame": other_frame, "tooltip": "Manually trigger 'New Day' log."},
            "Settings":{"command_ref": self.open_settings, "frame": other_frame, "tooltip": "Open configuration window."},
            "Sync Excel->DB":{"command_ref": self.sync_excel_to_sqlite_triggered, "frame": other_frame, "tooltip": "Sync Excel log to SQLite DB."}
        }

        custom_buttons_data_list = []
        valid_custom_configs = self.custom_button_configs[:self.num_custom_buttons]

        for i, config in enumerate(valid_custom_configs):
            button_text = config.get("text", f"Custom {i+1}")
            event_desc = config.get("event_text", f"{button_text} Triggered")
            custom_buttons_data_list.append({
                "text": button_text,
                "config": config, 
                "frame": self.custom_frame, 
                "tooltip": f"Log '{event_desc}'. Right-click to edit/delete.",
                "index": i 
            })

        buttons_dict = {} 

        for text, data in all_buttons_data.items():
            style_name = f"{text.replace(' ', '').replace('->','')}.TButton"
            _, bg_color = self.button_colors.get(text, (None, None))
            final_style = "TButton" 
            if bg_color:
                try:
                    self.style.configure(style_name, background=bg_color)
                    final_style = style_name 
                except tk.TclError as e_style:
                    print(f"Warning: Style color '{bg_color}' for '{text}'. Error: {e_style}")
            button = ttk.Button(data["frame"], text=text, style=final_style)
            buttons_dict[text] = {"widget": button, "data": data}

        for i, button_data in enumerate(custom_buttons_data_list): 
            button_text = button_data["text"]
            style_name = f"Custom.{button_text.replace(' ', '')}.TButton" 
            _, bg_color = self.parent_gui.button_colors.get(button_text, (None, None)) if hasattr(self, 'parent_gui') else self.button_colors.get(button_text, (None, None))

            final_style = "TButton"
            if bg_color:
                try:
                    self.style.configure(style_name, background=bg_color)
                    final_style = style_name
                except tk.TclError as e_style:
                    print(f"Warning: Style color '{bg_color}' for custom button '{button_text}'. Error: {e_style}")
            
            button = ttk.Button(button_data["frame"], text=button_text, style=final_style)
            buttons_dict[button_text] = {"widget": button, "data": button_data}
            self.custom_buttons_widgets_ref.append(button) 

        btn_row_log, btn_row_other = 0, 0
        custom_idx_grid = 0 

        for text, item in buttons_dict.items():
            button = item["widget"]
            data = item["data"]
            frame = data["frame"]
            cmd = None
            command_ref = data.get("command_ref") 
            config_ref = data.get("config")    
            button_original_index = data.get("index") 

            if command_ref: 
                if text in ["Log on", "Log off", "Event"]:
                    cmd = lambda t=text, b=button, ref=command_ref: ref(t, b)
                elif text in ["SVP", "New Day"]:
                    cmd = lambda b=button, ref=command_ref: ref(b)
                elif text == "Settings" or text == "Sync Excel->DB":
                    cmd = command_ref 
            elif config_ref: 
                cmd = lambda cfg=config_ref, b=button: self.log_custom_event(cfg, b)
                button.bind("<Button-3>", lambda event, b_widget=button, b_idx=button_original_index: \
                            self.show_custom_button_context_menu(event, b_widget, b_idx))
            else:
                cmd = lambda t=text: print(f"Error: No command/config for {t}") 

            button.config(command=cmd)

            if frame == logging_frame:
                button.grid(row=btn_row_log, column=0, padx=5, pady=5, sticky="ew") # Increased pady
                frame.rowconfigure(btn_row_log, weight=1)
                btn_row_log += 1
            elif frame == other_frame:
                button.grid(row=btn_row_other, column=0, padx=5, pady=5, sticky="ew") # Increased pady
                frame.rowconfigure(btn_row_other, weight=1)
                btn_row_other += 1
            elif frame == self.custom_frame: 
                custom_row = custom_idx_grid // num_custom_cols
                custom_col = custom_idx_grid % num_custom_cols
                button.grid(row=custom_row, column=custom_col, padx=5, pady=5, sticky="nsew") # Increased pady
                frame.rowconfigure(custom_row, weight=1)
                custom_idx_grid += 1
            
            ToolTip(button, data["tooltip"])
        parent_frame.update_idletasks()

    # --- Context Menu and Editor Dialog Methods ---
    def show_custom_button_context_menu(self, event, button_widget, button_index):
        """Shows context menu for a specific custom button."""
        context_menu = tk.Menu(self.master, tearoff=0)
        context_menu.add_command(label="Edit Button", command=lambda idx=button_index: self.open_custom_button_editor(idx))
        context_menu.add_command(label="Delete Button", command=lambda idx=button_index: self.delete_custom_button(idx))
        try:
            context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            context_menu.grab_release()

    def show_custom_frame_context_menu(self, event):
        """Shows context menu for the custom_frame (for adding new buttons)."""
        if event.widget == self.custom_frame:
            context_menu = tk.Menu(self.master, tearoff=0)
            if len(self.custom_button_configs) < 10:
                 context_menu.add_command(label="Add New Custom Button", command=self.open_custom_button_editor) 
            else:
                 context_menu.add_command(label="Add New Custom Button (Max Reached)", state="disabled")
            try:
                context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                context_menu.grab_release()

    def open_custom_button_editor(self, button_index=None):
        """Opens the dialog to edit or add a custom button."""
        editor = CustomButtonEditorDialog(self.master, self, button_index)
        self.master.wait_window(editor) 

    def delete_custom_button(self, button_index):
        """Deletes a custom button after confirmation."""
        if not (0 <= button_index < len(self.custom_button_configs)):
            messagebox.showerror("Error", "Invalid button index for deletion.", parent=self.master)
            return

        config_to_delete = self.custom_button_configs[button_index]
        button_text_to_delete = config_to_delete.get("text")

        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete the custom button '{button_text_to_delete}'?", parent=self.master, icon='warning'):
            del self.custom_button_configs[button_index]
            self.num_custom_buttons = len(self.custom_button_configs)
            if button_text_to_delete in self.button_colors:
                del self.button_colors[button_text_to_delete]
            
            self.update_custom_buttons() 
            self.save_settings()         
            self.update_status(f"Custom button '{button_text_to_delete}' deleted.")

    def create_status_indicators(self, parent_frame):
        indicator_frame = ttk.Frame(parent_frame, padding="5 0")
        indicator_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(8, 0)) # Increased pady
        indicator_frame.columnconfigure(1, weight=0) 
        indicator_frame.columnconfigure(3, weight=0) 
        indicator_frame.columnconfigure(4, weight=1) 

        ttk.Label(indicator_frame, text="Monitoring:", font=("Arial", 9, "bold")).grid(row=0, column=0, sticky=tk.W, padx=(0, 5)) # Increased padx
        self.monitor_status_label = ttk.Label(indicator_frame, text="Initializing...", foreground="orange", font=("Arial", 9))
        self.monitor_status_label.grid(row=0, column=1, sticky=tk.W, padx=(0, 15)) # Increased padx

        ttk.Label(indicator_frame, text="SQLite:", font=("Arial", 9, "bold")).grid(row=0, column=2, sticky=tk.W, padx=(0, 5)) # Increased padx
        self.db_status_label = ttk.Label(indicator_frame, text="Initializing...", foreground="orange", font=("Arial", 9))
        self.db_status_label.grid(row=0, column=3, sticky=tk.W)
        ttk.Frame(indicator_frame).grid(row=0, column=4) 

        self.update_db_indicator()


    def sync_excel_to_sqlite_triggered(self):
        """Handles the 'Sync Excel->DB' button press."""
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
                for child_frame in self.button_frame.winfo_children(): 
                    if isinstance(child_frame, ttk.LabelFrame): 
                        for btn in child_frame.winfo_children():
                            if isinstance(btn, ttk.Button) and btn.cget('text') == target_button_text:
                                sync_button = btn
                                break
                    if sync_button: break
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
            else:
                pass 

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
        """
        Reads Excel, reads SQLite, compares, and updates SQLite.
        Returns (bool_success, status_message).
        MUST RUN IN A BACKGROUND THREAD.
        Requires pandas library.
        """
        print("\n--- Starting perform_excel_to_sqlite_sync ---")
        excel_file = self.log_file_path
        db_file = self.sqlite_db_path
        db_table = self.sqlite_table
        record_id_column = "RecordID" 
        date_col_name = "Date" 
        time_col_name = "Time" 

        excel_date_col_name = self.txt_field_columns.get("Date", "Date")
        excel_time_col_name = self.txt_field_columns.get("Time", "Time")


        print(f"Sync Params: Excel='{excel_file}', DB='{db_file}', Table='{db_table}', ID Col='{record_id_column}'")
        print(f"Excel Date Col: '{excel_date_col_name}', Excel Time Col: '{excel_time_col_name}'")


        if not excel_file or not db_file or not db_table:
            print("Sync Error: Missing file paths or table name.")
            return False, "Sync Error: Configuration paths or table missing."

        excel_data = {}
        app = None; wb = None; sheet = None; header = None; df_excel = None

        try:
            print("Sync Step 1: Reading Excel file...")
            app = xw.App(visible=False, add_book=False)
            if not os.path.exists(excel_file):
                raise FileNotFoundError(f"Excel file not found: {excel_file}")
            wb = app.books.open(excel_file, update_links=False, read_only=True)
            if not wb.sheets: raise ValueError("Workbook has no sheets")
            sheet = wb.sheets[0]
            
            header_range = sheet.range('A1').expand('right')
            if header_range is None or header_range.value is None:
                 raise ValueError("Cannot find header range or header is empty in Excel.")
            header = [str(h) if h is not None else "" for h in header_range.value] 

            if record_id_column not in header:
                raise ValueError(f"Column '{record_id_column}' not found in Excel header: {header}")

            last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
            if last_row <= 1: 
                print("Sync Info: Excel sheet appears empty or has only header.")
                if wb: wb.close()
                if app: app.quit()
                return True, "Sync Info: Excel sheet is empty, nothing to sync."

            data_range = sheet.range((2, 1), (last_row, len(header)))
            df_excel = pd.DataFrame(data_range.value, columns=header)

            if excel_date_col_name in df_excel.columns:
                print(f"  - Converting Excel column '{excel_date_col_name}' to datetime")
                df_excel[excel_date_col_name] = pd.to_datetime(df_excel[excel_date_col_name], errors='coerce').dt.strftime('%Y-%m-%d')
            
            if excel_time_col_name in df_excel.columns:
                print(f"  - Converting Excel column '{excel_time_col_name}' to time string")
                def format_excel_time(t_val):
                    if pd.isna(t_val): return None
                    if isinstance(t_val, (datetime.time, datetime.datetime)):
                        return t_val.strftime('%H:%M:%S')
                    if isinstance(t_val, (float, int)): 
                        try: 
                            if 0 <= t_val < 1: 
                                total_seconds = int(t_val * 24 * 60 * 60)
                                hours = total_seconds // 3600
                                minutes = (total_seconds % 3600) // 60
                                seconds = total_seconds % 60
                                return f"{hours:02}:{minutes:02}:{seconds:02}"
                            else: 
                                return pd.NaT 
                        except: return pd.NaT
                    return str(t_val) 
                df_excel[excel_time_col_name] = df_excel[excel_time_col_name].apply(format_excel_time)


            if record_id_column in df_excel.columns:
                df_excel[record_id_column] = df_excel[record_id_column].astype(str).str.strip()
                df_excel = df_excel[df_excel[record_id_column] != '']
                df_excel = df_excel.dropna(subset=[record_id_column])
            else:
                raise ValueError(f"'{record_id_column}' column disappeared after initial read.")

            if df_excel.empty:
                print("Sync Info: No valid rows with RecordIDs found after cleaning Excel data.")
                if wb: wb.close()
                if app: app.quit()
                return True, "Sync Info: No valid Excel rows found to sync."

            df_excel = df_excel.set_index(record_id_column, drop=False) 
            excel_data = df_excel.to_dict('index')
            print(f"Sync Step 1 Complete: Read {len(excel_data)} rows with valid RecordIDs from Excel.")

        except Exception as e_excel:
            print(f"--- CAUGHT EXCEPTION in Excel Read (Step 1) --- Type: {type(e_excel).__name__}, Args: {e_excel.args}")
            traceback.print_exc()
            return False, f"Sync Error: Reading Excel failed - {type(e_excel).__name__}: {e_excel.args}"
        finally:
            if wb: wb.close()
            if app: app.quit()

        sqlite_data = {}
        conn_sqlite = None
        db_cols = []
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
                return False, f"Sync Error: Column '{record_id_column}' not found in SQLite table '{db_table}'."

            quoted_db_cols = ", ".join([f"[{c}]" for c in db_cols])
            cursor.execute(f"SELECT {quoted_db_cols} FROM [{db_table}]")
            rows = cursor.fetchall()
            for row in rows:
                row_dict = dict(row)
                rec_id = str(row_dict.get(record_id_column, '')).strip()
                if rec_id:
                    sqlite_data[rec_id] = row_dict
            print(f"Sync Step 2 Complete: Read {len(sqlite_data)} rows from SQLite table '{db_table}'.")

        except sqlite3.Error as e_sqlite:
            print(f"Sync Error (Step 2 - SQLite): Failed reading DB. Error: {e_sqlite}")
            traceback.print_exc()
            return False, f"Sync Error: Reading SQLite failed - {type(e_sqlite).__name__}"
        finally:
            if conn_sqlite: conn_sqlite.close()
        
        updates_to_apply = [] 
        records_processed = 0
        records_needing_update_in_db = 0 
        db_cols_set = set(db_cols) 

        print(f"Sync Step 3: Comparing {len(excel_data)} Excel rows to {len(sqlite_data)} SQLite rows...")

        for rec_id, excel_row_dict in excel_data.items():
            records_processed += 1
            if rec_id not in sqlite_data:
                print(f"  - Info: RecordID '{rec_id}' from Excel not found in SQLite. Will be skipped for update (sync is update-only).")
                continue 

            sqlite_row_dict = sqlite_data[rec_id]
            row_has_diff = False

            for excel_col_key, excel_val in excel_row_dict.items():
                if excel_col_key not in db_cols_set:
                    continue
                
                if excel_col_key == record_id_column:
                    continue

                sqlite_val = sqlite_row_dict.get(excel_col_key)
                str_excel_val = str(excel_val).strip() if excel_val is not None else ""
                str_sqlite_val = str(sqlite_val).strip() if sqlite_val is not None else ""
                
                if str_excel_val != str_sqlite_val:
                    print(f"  -> Difference Found! RecordID: {rec_id}, Column: '{excel_col_key}'")
                    print(f"     Excel Value (raw)  : '{excel_val}' (Type: {type(excel_val).__name__})")
                    print(f"     Excel Value (norm) : '{str_excel_val}'")
                    print(f"     SQLite Value (raw) : '{sqlite_val}' (Type: {type(sqlite_val).__name__})")
                    print(f"     SQLite Value (norm): '{str_sqlite_val}'")
                    updates_to_apply.append((rec_id, excel_col_key, excel_val)) 
                    row_has_diff = True
            
            if row_has_diff:
                records_needing_update_in_db +=1

        print(f"Sync Step 3 Complete: Comparison found {len(updates_to_apply)} cell differences across {records_needing_update_in_db} records.")

        if not updates_to_apply:
            print("Sync Step 4: No differences found requiring update in SQLite.")
            return True, f"Sync complete. No changes detected in {records_processed} Excel rows compared to SQLite."

        print(f"Sync Step 4: Applying {len(updates_to_apply)} cell updates to SQLite for {records_needing_update_in_db} records...")
        conn_sqlite = None 
        try:
            conn_sqlite = sqlite3.connect(db_file, timeout=10)
            cursor = conn_sqlite.cursor()
            update_statements_run = 0
            actual_rows_affected_in_db = 0

            updates_by_record = {}
            for r_id, col, val in updates_to_apply:
                if r_id not in updates_by_record:
                    updates_by_record[r_id] = {}
                updates_by_record[r_id][col] = val
            
            print(f"  - Grouped updates for {len(updates_by_record)} unique RecordIDs.")

            for r_id, col_val_dict in updates_by_record.items():
                set_clauses = []
                values_for_sql = []
                for col, val in col_val_dict.items():
                    set_clauses.append(f"[{col}] = ?")
                    if pd.isna(val):
                        values_for_sql.append(None)
                    else:
                        values_for_sql.append(val) 
                
                if set_clauses:
                    values_for_sql.append(r_id) 
                    sql_update = f"UPDATE [{db_table}] SET {', '.join(set_clauses)} WHERE [{record_id_column}] = ?"
                    cursor.execute(sql_update, values_for_sql)
                    update_statements_run += 1
                    actual_rows_affected_in_db += cursor.rowcount 
                    if cursor.rowcount == 0:
                         print(f"  - Warning: UPDATE statement affected 0 rows for RecordID {r_id}. It might have been deleted from DB concurrently.")


            conn_sqlite.commit()
            print(f"  - Commit successful. Ran {update_statements_run} UPDATE statements. Total rows affected in DB: {actual_rows_affected_in_db}.")
            return True, f"Sync successful. Updated {len(updates_by_record)} records ({actual_rows_affected_in_db} rows affected) in SQLite."

        except sqlite3.Error as e_update:
            print(f"Sync Error (Step 4 - SQLite Update): Failed updating DB. Error: {e_update}")
            traceback.print_exc()
            if conn_sqlite: conn_sqlite.rollback()
            return False, f"Sync Error: Updating SQLite failed - {type(e_update).__name__}"
        finally:
            if conn_sqlite: conn_sqlite.close()


    def create_status_bar(self, parent_frame):
        self.status_var.set("Status: Ready")
        status_bar = ttk.Label(parent_frame, textvariable=self.status_var, style="StatusBar.TLabel", anchor='w')
        status_bar.grid(row=2, column=0, sticky=(tk.W, tk.E), padx=0, pady=(8,0)) # Increased pady

    def update_status(self, message):
        def _update():
            timestamp = datetime.datetime.now().strftime("%H:%M:%S")
            max_len = 100 
            display_message = message if len(message) <= max_len else message[:max_len-3] + "..."
            try:
                if self.status_var: 
                    self.status_var.set(f"[{timestamp}] {display_message}")
            except tk.TclError: 
                print(f"Status Update Error (window closing?): {message}")
        if hasattr(self, 'master') and self.master.winfo_exists():
            try:
                self.master.after(0, _update)
            except tk.TclError: 
                print(f"Status Update Error (scheduling failed?): {message}")


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
            print("DB Indicator Error: Could not configure label (window closing?).")

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
                                 triggering_button=button_widget)

    def log_custom_event(self, config, button_widget):
        button_text = config.get("text", "Unknown Custom")
        event_text_for_excel = config.get("event_text", f"{button_text} Triggered")
        print(f"'{button_text}' button pressed. Event text: '{event_text_for_excel}'")
        self._perform_log_action(event_type=button_text, 
                                 event_text_for_excel=event_text_for_excel,
                                 triggering_button=button_widget)

    def log_new_day(self, button_widget=None): 
        print("Logging 'New Day' event.")
        self._perform_log_action(event_type="New Day",
                                 event_text_for_excel="New Day",
                                 triggering_button=button_widget)

    def apply_svp(self, button_widget):
        print("Applying SVP...")
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
                                 triggering_button=button_widget)

    def _perform_log_action(self, event_type, event_text_for_excel, skip_latest_files=False, svp_specific_handling=False, triggering_button=None):
        print(f"Queueing log action: Type='{event_type}'")
        self.update_status(f"Processing '{event_type}'...")
        original_text = None
        if triggering_button and isinstance(triggering_button, ttk.Button):
            try:
                if triggering_button.winfo_exists(): 
                    original_text = triggering_button['text']
                    triggering_button.config(state=tk.DISABLED, text="Working...")
            except tk.TclError:
                print("Warning: Could not disable button (already destroyed?).")
                triggering_button = None 

        def _worker_thread_func():
            nonlocal original_text 
            row_data = {}; excel_success = False; sqlite_logged = False
            excel_save_exception = None; sqlite_save_exception_type = None
            status_msg = f"'{event_type}' processed with errors." 
            record_id = str(uuid.uuid4())
            row_data['RecordID'] = record_id
            try:
                event_col_name = self.txt_field_columns.get("Event", "Event")
                row_data["EventType"] = event_type 
                if event_text_for_excel is not None:
                    row_data[event_col_name] = event_text_for_excel 

                print(f"Thread '{event_type}': Fetching TXT data...")
                try:
                    txt_data = self.insert_txt_data()
                    if txt_data: row_data.update(txt_data)
                except Exception as e_txt:
                    print(f"Thread '{event_type}': Error fetching TXT data: {e_txt}")
                    self.master.after(0, lambda e=e_txt: messagebox.showerror("Error", f"Failed to read TXT data:\n{e}", parent=self.master))

                if not skip_latest_files:
                    print(f"Thread '{event_type}': Fetching latest filenames...")
                    try:
                        latest_files_data = self.get_latest_files_data()
                        if latest_files_data: row_data.update(latest_files_data)
                    except Exception as e_files:
                        print(f"Thread '{event_type}': Error fetching latest file data: {e_files}")
                        self.master.after(0, lambda e=e_files: messagebox.showerror("Error", f"Failed to get latest file data:\n{e}", parent=self.master))

                if svp_specific_handling:
                    print(f"Thread '{event_type}': Performing SVP specific file handling...")
                    svp_folder_path = self.folder_paths.get("SVP")
                    svp_col_name = self.folder_columns.get("SVP", "SVP") 
                    if svp_folder_path and svp_col_name:
                        latest_svp_file = folder_cache.get("SVP")
                        row_data[svp_col_name] = latest_svp_file if latest_svp_file else "N/A"
                    elif svp_col_name: 
                        print("SVP folder not defined in settings, but column is.")
                        row_data[svp_col_name] = "Config Error (Path)"


                if row_data: 
                    color_tuple = self.button_colors.get(event_type, (None, None))
                    row_color_for_excel = color_tuple[1] if isinstance(color_tuple, tuple) and len(color_tuple) > 1 else None
                    
                    excel_data_to_log = {k: v for k, v in row_data.items() if k != 'EventType'} 

                    print(f"Thread '{event_type}': Saving data to Excel...")
                    try:
                        if not self.log_file_path: excel_save_exception = ValueError("Excel path missing")
                        elif not os.path.exists(self.log_file_path): excel_save_exception = FileNotFoundError("Excel file missing")
                        else:
                            self.save_to_excel(excel_data_to_log, row_color=row_color_for_excel)
                            excel_success = True
                    except Exception as e_excel:
                        excel_save_exception = e_excel
                        print(f"Thread '{event_type}': Error saving data to Excel: {e_excel}")
                        traceback.print_exc() 
                        self.master.after(0, lambda e=e_excel: messagebox.showerror("Excel Error", f"Failed to save data to Excel:\n{e}", parent=self.master))
                    
                    print(f"Thread '{event_type}': Saving data to SQLite...")
                    sqlite_logged, sqlite_save_exception_type = self.log_to_sqlite(row_data) 

                    status_parts = []
                    if excel_success: status_parts.append("Excel: OK")
                    elif excel_save_exception: status_parts.append(f"Excel: Fail ({type(excel_save_exception).__name__})")
                    else: status_parts.append("Excel: Fail (Path/Access)") 

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
                    print(f"Thread '{event_type}': No data generated.")

            except Exception as thread_ex:
                print(f"!!! Unexpected Error in logging thread for '{event_type}' !!!")
                print(traceback.format_exc())
                status_msg = f"'{event_type}' - Unexpected thread error: {thread_ex}"
                self.master.after(0, lambda e=thread_ex: messagebox.showerror("Thread Error", f"Critical error during logging action '{event_type}':\n{e}", parent=self.master))
            finally:
                print(f"Thread '{event_type}': Action finished. Status: {status_msg}")
                self.master.after(0, self.update_status, status_msg) 
                if triggering_button and isinstance(triggering_button, ttk.Button):
                    def re_enable_button(btn=triggering_button, txt=original_text):
                        try:
                            if btn and btn.winfo_exists():
                                btn.config(state=tk.NORMAL)
                                if txt: 
                                    btn.config(text=txt)
                        except tk.TclError:
                            print("Warning: Could not re-enable button (already destroyed?).")
                    self.master.after(0, re_enable_button)
        log_thread = threading.Thread(target=_worker_thread_func, daemon=True)
        log_thread.start()

    def insert_txt_data(self):
        row_data = {} 
        current_dt = datetime.datetime.now() 
        current_timestamp = time.time() 
        use_pc_time = False 
        reason_for_pc_time = "" 

        if not self.txt_folder_path or not os.path.exists(self.txt_folder_path):
            print("Warn: TXT folder path missing or invalid.")
            use_pc_time = True
            reason_for_pc_time = "TXT folder path missing or invalid"
            self.txt_file_path = None 
        else:
            self.txt_file_path = self.find_latest_file_in_folder(self.txt_folder_path, ".txt")
            if not self.txt_file_path:
                print("Warn: No TXT file found in the specified folder.")
                use_pc_time = True
                reason_for_pc_time = "No TXT file found"
        
        if self.txt_file_path and not use_pc_time: 
            try:
                file_mod_timestamp = os.path.getmtime(self.txt_file_path)
                time_diff = current_timestamp - file_mod_timestamp
                if time_diff > 1.0: 
                    print(f"Info: TXT file '{os.path.basename(self.txt_file_path)}' modified {time_diff:.2f} seconds ago (> 1s).")
                    use_pc_time = True
                    reason_for_pc_time = f"file modified {time_diff:.2f}s ago"
            except OSError as e_modtime:
                print(f"Warn: Could not get modification time for '{self.txt_file_path}': {e_modtime}.")
                use_pc_time = True 
                reason_for_pc_time = "failed to get file modification time"
        
        txt_data_found = False 
        parse_success = True 
        temp_txt_data = {} 

        if not use_pc_time and self.txt_file_path: 
            print(f"Info: Attempting to read data from recent TXT file: {os.path.basename(self.txt_file_path)}")
            try:
                lines = []; encodings_to_try = ['utf-8', 'latin-1', 'cp1252']; read_success = False; last_error = None
                for enc in encodings_to_try:
                    try:
                        for attempt in range(3): 
                            try:
                                with open(self.txt_file_path, "r", encoding=enc) as file: lines = file.readlines()
                                read_success = True; break 
                            except IOError as e_io:
                                if attempt < 2: time.sleep(0.1); continue 
                                else: raise e_io 
                        if read_success: break 
                    except UnicodeDecodeError: last_error = f"UnicodeDecodeError with {enc}"; continue 
                    except Exception as e_open: last_error = f"Error reading TXT file {os.path.basename(self.txt_file_path)} with {enc}: {e_open}"; print(last_error); lines = []; break 

                if not read_success and not lines:
                    print(f"Warn: Could not decode or read TXT file '{os.path.basename(self.txt_file_path)}'. Last error: {last_error}")
                    use_pc_time = True 
                    reason_for_pc_time = "failed to read/decode TXT file"

                if lines:
                    latest_line_str = lines[-1].strip()
                    latest_line_parts = latest_line_str.split(",")
                    field_keys = ["Date", "Time", "KP", "DCC", "Line name", "Latitude", "Longitude", "Easting", "Northing"]

                    for i, field_key in enumerate(field_keys):
                        excel_col = self.txt_field_columns.get(field_key) 
                        skip_field = self.txt_field_skips.get(field_key, False) 

                        if excel_col and not skip_field:
                            try:
                                value = latest_line_parts[i].strip()
                                temp_txt_data[excel_col] = value
                                if field_key in ["Date", "Time"]: txt_data_found = True 
                            except IndexError:
                                temp_txt_data[excel_col] = None 
                                print(f"Warn: Field '{field_key}' missing in TXT line for column '{excel_col}', assigning NULL.")
                                if field_key in ["Date", "Time"]: parse_success = False 
                            except Exception as e_parse:
                                temp_txt_data[excel_col] = None 
                                print(f"Warn: Err parsing field '{field_key}' for column '{excel_col}', assigning NULL. Error: {e_parse}")
                                if field_key in ["Date", "Time"]: parse_success = False 
                    
                    if not txt_data_found:
                        print("Info: No Date/Time field found or mapped reliably in TXT.")
                        use_pc_time = True; reason_for_pc_time = "Date/Time not found/mapped in TXT"
                    elif not parse_success:
                        print("Warn: Failed to parse required Date/Time from TXT.")
                        use_pc_time = True; reason_for_pc_time = "Date/Time parsing failed in TXT"
                else: 
                    if not use_pc_time: 
                        print("Info: TXT file found but empty or could not be read.")
                        use_pc_time = True; reason_for_pc_time = "TXT file empty or unreadable"
            except Exception as e: 
                print(f"Error processing TXT file '{os.path.basename(self.txt_file_path)}': {e}"); traceback.print_exc()
                if not use_pc_time: use_pc_time = True; reason_for_pc_time = f"unexpected error processing TXT: {type(e).__name__}"

        if use_pc_time:
            print(f"Info: Using PC Time/Date. Reason: {reason_for_pc_time}.")
            date_col = self.txt_field_columns.get("Date"); time_col = self.txt_field_columns.get("Time")
            skip_date = self.txt_field_skips.get("Date", False); skip_time = self.txt_field_skips.get("Time", False)
            if date_col and not skip_date: row_data[date_col] = current_dt.strftime("%Y-%m-%d")
            if time_col and not skip_time: row_data[time_col] = current_dt.strftime("%H:%M:%S")
            for col, val in temp_txt_data.items():
                if col != date_col and col != time_col: 
                    if col not in row_data: row_data[col] = val
        else:
            print("Info: Using Date and Time from recent TXT file.")
            row_data.update(temp_txt_data)
        return row_data 

    def get_latest_files_data(self):
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
        if not self.log_file_path:
            print("Excel Error: Log file path is not set.")
            raise ValueError("Excel log file path is missing.") 
        if not os.path.exists(self.log_file_path):
            print(f"Excel Error: Log file not found at {self.log_file_path}")
            raise FileNotFoundError(f"Excel log file not found: {self.log_file_path}")

        app = None; workbook = None; opened_new_app = False; opened_workbook = False
        print(f"Debug Excel: Attempting to write {len(row_data)} columns.")

        try:
            try:
                app = xw.apps.active
                if app is None: raise Exception("No active Excel instance")
                print("Debug Excel: Using active Excel instance.")
            except Exception:
                try:
                    print("Debug Excel: No active instance, starting new one.")
                    app = xw.App(visible=False) 
                    opened_new_app = True
                except Exception as e_app:
                    print(f"Fatal Excel Error: Could not start or connect to Excel: {e_app}")
                    raise ConnectionAbortedError(f"Failed to start/connect to Excel: {e_app}")

            target_norm_path = os.path.normcase(os.path.abspath(self.log_file_path))
            for wb_iter in app.books: 
                try:
                    if os.path.normcase(os.path.abspath(wb_iter.fullname)) == target_norm_path:
                        workbook = wb_iter
                        print(f"Debug Excel: Found workbook '{os.path.basename(workbook.fullname)}' already open.")
                        break
                except Exception as e_fullname: print(f"Warn: Error checking workbook fullname: {e_fullname}") 

            if workbook is None:
                try:
                    print(f"Debug Excel: Opening workbook '{self.log_file_path}'")
                    workbook = app.books.open(self.log_file_path)
                    opened_workbook = True
                except Exception as e_open:
                    print(f"Fatal Excel Error: Could not open workbook {self.log_file_path}: {e_open}")
                    raise IOError(f"Failed to open Excel workbook: {e_open}")

            sheet = workbook.sheets[0] 
            header_range_obj = sheet.range("A1").expand("right")
            header_values = header_range_obj.value
            if not header_values or not any(h is not None for h in header_values):
                print("Fatal Excel Error: Excel header row (A1) is missing or empty.")
                raise ValueError("Excel header row is missing or empty.")

            record_id_col_name = "RecordID" 
            if record_id_col_name not in header_values:
                print(f"Fatal Excel Error: Header row does not contain a '{record_id_col_name}' column.")
                raise ValueError(f"Excel header missing required '{record_id_col_name}' column.")
            
            header_map_lower = {str(h).lower(): i + 1 for i, h in enumerate(header_values) if h is not None}
            last_header_col_index = max(header_map_lower.values()) if header_map_lower else 1

            if next_row is None:
                try:
                    last_row_a = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
                    check_row = last_row_a + 1
                    if check_row < 2: check_row = 2 
                    while sheet.range(f'A{check_row}').value is not None: check_row += 1
                    next_row = check_row
                    print(f"Debug Excel: Determined next empty row: {next_row}")
                except Exception as e_row:
                    print(f"Warn Excel: Error finding next empty row ({e_row}). Defaulting to row 2.")
                    next_row = 2
            
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
                        elif isinstance(value, datetime.datetime): 
                            write_value = value.strftime("%Y-%m-%d %H:%M:%S")
                        elif isinstance(value, datetime.date):
                            write_value = value.strftime("%Y-%m-%d")
                        elif isinstance(value, datetime.time):
                            write_value = value.strftime("%H:%M:%S")
                        
                        sheet.range(next_row, col_index).value = write_value
                        written_cols.append(col_index)
                    except Exception as e_write_cell:
                        print(f"Warn Excel: Failed to write value '{value}' to cell ({next_row},{col_index}) for column '{col_name}'. Error: {e_write_cell}")

            if row_color and written_cols:
                try:
                    target_range = sheet.range((next_row, 1), (next_row, last_header_col_index))
                    target_range.color = row_color
                    print(f"Debug Excel: Applied color {row_color} to row {next_row}.")
                except Exception as e_color: print(f"Warn Excel: Failed to apply color to row {next_row}. Error: {e_color}")

            try:
                workbook.save()
                print("Debug Excel: Workbook saved.")
            except Exception as e_save:
                print(f"Error saving workbook: {e_save}")
                raise IOError(f"Failed to save Excel workbook: {e_save}")

        except Exception as e:
            print(f"Error during save_to_excel: {e}"); traceback.print_exc(); raise e 
        finally:
            if workbook is not None and opened_workbook: 
                try: workbook.close(save_changes=False); print("Debug Excel: Closed workbook.")
                except Exception as e_close: print(f"Warn: Error closing workbook: {e_close}")
            if app is not None and opened_new_app: 
                try: app.quit(); print("Debug Excel: Quit Excel instance.")
                except Exception as e_quit: print(f"Warn: Error quitting Excel: {e_quit}")
            if app is not None and opened_new_app: app = None
            elif app is not None and not opened_new_app: print("Debug Excel: Did not quit existing Excel instance.")

    def log_to_sqlite(self, row_data):
        success = False
        error_type = None 
        if not self.sqlite_enabled:
            return False, "Disabled"
        if not self.sqlite_db_path or not self.sqlite_table:
            msg = "SQLite Log Error: DB Path or Table Name missing in settings."
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
                if not results: raise sqlite3.OperationalError(f"No such table: {self.sqlite_table}")
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
                    data_to_insert[db_col_name] = row_data[original_key]
            
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
            print(f"SQLite Log Error ({error_type}): {error_message}")
            traceback.print_exc()
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
            print(f"SQLite Log Error (Unexpected - {error_type}): {error_message}")
            traceback.print_exc()
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
                except Exception as e_cursor_close:
                    print(f"Debug SQLite: Error closing cursor: {e_cursor_close}")
            if conn:
                try:
                    conn.close()
                except Exception as e_close:
                    print(f"Debug SQLite: Error closing connection: {e_close}")
        return success, error_type

    def show_sqlite_error_message(self, error_message, error_type):
        parent_window = self.settings_window_instance if (hasattr(self, 'settings_window_instance') and self.settings_window_instance and self.settings_window_instance.winfo_exists()) else self.master
        if error_type == "NoSuchTable":
            messagebox.showerror("SQLite Error", f"Table '{self.sqlite_table}' not found.\nPlease check table name or create table.\nDB: {self.sqlite_db_path}", parent=parent_window)
        elif error_type == "NoSuchColumn":
            try: missing_col = error_message.split("column named")[-1].strip().split(":")[0].strip().strip("'\"[]")
            except: missing_col = "[unknown]"
            messagebox.showerror("SQLite Error", f"Column '{missing_col}' not found in table '{self.sqlite_table}'.\nCheck Settings (TXT Columns / Folder Columns) vs. DB table structure.\n\n(Original error: {error_message})", parent=parent_window)
        elif error_type == "DatabaseLocked":
            messagebox.showerror("SQLite Error", f"Database file is locked.\nAnother program might be using it.\nDB: {self.sqlite_db_path}\n\n(Original error: {error_message})", parent=parent_window)
        else: 
            messagebox.showerror("SQLite Operational Error", f"Error interacting with database:\n{error_message}", parent=parent_window)

    def save_settings(self):
        colors_to_save = {}
        for key, (_, color_hex) in self.button_colors.items():
            if color_hex:  
                colors_to_save[key] = color_hex
        
        settings = {
            "log_file_path": self.log_file_path, "txt_folder_path": self.txt_folder_path,
            "txt_field_columns": self.txt_field_columns, "txt_field_skips": self.txt_field_skips,
            "folder_paths": self.folder_paths, "folder_columns": self.folder_columns,
            "file_extensions": self.file_extensions, "folder_skips": self.folder_skips,
            "num_custom_buttons": self.num_custom_buttons, 
            "custom_button_configs": self.custom_button_configs, 
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
                self.log_file_path = settings.get("log_file_path")
                self.txt_folder_path = settings.get("txt_folder_path")
                self.txt_field_columns = {"Event": "Event"}; self.txt_field_columns.update(settings.get("txt_field_columns", {}))
                self.txt_field_skips.clear(); self.txt_field_skips.update(settings.get("txt_field_skips", {}))
                self.folder_paths.clear(); self.folder_paths.update(settings.get("folder_paths", {}))
                self.folder_columns.clear(); self.folder_columns.update(settings.get("folder_columns", {}))
                self.file_extensions.clear(); self.file_extensions.update(settings.get("file_extensions", {}))
                self.folder_skips.clear(); self.folder_skips.update(settings.get("folder_skips", {}))
                
                self.num_custom_buttons = settings.get("num_custom_buttons", 3) 
                loaded_configs = settings.get("custom_button_configs", [])
                self.custom_button_configs = loaded_configs
                if self.num_custom_buttons > len(self.custom_button_configs):
                    for i in range(len(self.custom_button_configs), self.num_custom_buttons):
                        self.custom_button_configs.append({"text": f"Custom {i+1}", "event_text": f"Custom {i+1} Event"})
                elif self.num_custom_buttons < len(self.custom_button_configs):
                    self.custom_button_configs = self.custom_button_configs[:self.num_custom_buttons]

                loaded_colors_hex = settings.get("button_colors", {}) 
                default_button_colors = {
                    "Log on": (None, "#90EE90"), "Log off": (None, "#FFB6C1"),
                    "Event": (None, "#FFFFE0"), "SVP": (None, "#ADD8E6"),
                    "New Day": (None, "#FFFF99")
                }
                self.button_colors = default_button_colors.copy() 
                for key, color_hex_val in loaded_colors_hex.items():
                    self.button_colors[key] = (None, color_hex_val)
                
                for config in self.custom_button_configs:
                    btn_text = config.get("text")
                    if btn_text and btn_text not in self.button_colors:
                        self.button_colors[btn_text] = (None, None) 

                self.sqlite_enabled = settings.get("sqlite_enabled", False)
                self.sqlite_db_path = settings.get("sqlite_db_path")
                self.sqlite_table = settings.get("sqlite_table", "EventLog")
                print("Settings loaded successfully."); self.update_status("Settings loaded.")
            else:
                print("Settings file not found. Using default variables."); self.update_status("Settings file not found. Using defaults.")
                self.custom_button_configs = self.custom_button_configs[:self.num_custom_buttons]
                while len(self.custom_button_configs) < self.num_custom_buttons:
                    idx = len(self.custom_button_configs) + 1
                    self.custom_button_configs.append({"text": f"Custom {idx}", "event_text": f"Custom {idx} Event"})
                    if f"Custom {idx}" not in self.button_colors: 
                        self.button_colors[f"Custom {idx}"] = (None, None)


        except json.JSONDecodeError as e:
            print(f"Error loading settings: Invalid JSON in {self.settings_file}. Error: {e}")
            messagebox.showerror("Load Error", f"Settings file '{self.settings_file}' has invalid format:\n{e}\n\nUsing default settings.", parent=self.master)
            self.update_status("Error loading settings: Invalid format."); self.init_variables() 
        except Exception as e:
            print(f"Error loading settings: {e}"); traceback.print_exc()
            messagebox.showerror("Load Error", f"Could not load settings from {self.settings_file}:\n{e}\n\nUsing default settings.", parent=self.master)
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
            settings_gui = SettingsWindow(settings_top_level, self)
            settings_gui.load_settings() 
            self.master.wait_window(settings_top_level) 
            try: del self.settings_window_instance
            except AttributeError: pass

    def update_custom_buttons(self):
        """Redraws the main button area, reflecting changes in custom button settings."""
        if hasattr(self, 'button_frame') and self.button_frame:
            print("Redrawing main buttons due to custom button update...")
            self.create_main_buttons(self.button_frame) 
            self.master.update_idletasks()
        else: print("Error: Button frame does not exist when trying to update buttons.")

    def start_monitoring(self):
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
            elif folder_path: print(f"Skipping monitor for '{folder_name}': Path is not a valid directory ('{folder_path}').")
        
        print(f"Monitoring {count} folders.")
        self.update_status(f"Monitoring {count} active folders.")

        if hasattr(self, 'monitor_status_label') and self.monitor_status_label:
            if monitoring_active: self.monitor_status_label.config(text="Active", foreground="green")
            else: self.monitor_status_label.config(text="Inactive", foreground="red")
        self.update_db_indicator() 

    def start_folder_monitoring(self, folder_name, folder_path, file_extension):
        try: os.listdir(folder_path) 
        except Exception as e: print(f"Error accessing '{folder_path}' for '{folder_name}': {e}. Monitor not started."); return False
        try:
            event_handler = FolderMonitor(folder_path, folder_name, self, file_extension)
            observer = PollingObserver(timeout=1) 
            observer.schedule(event_handler, folder_path, recursive=False)
            observer.start()
            self.monitors[folder_name] = observer
            threading.Thread(target=event_handler.update_latest_file, daemon=True).start() 
            print(f"Started monitoring '{folder_name}' at '{folder_path}' (Ext: '{file_extension or 'Any'}')")
            return True
        except Exception as e: print(f"Error starting observer for '{folder_path}': {e}"); traceback.print_exc(); return False

    def schedule_new_day(self):
        now = datetime.datetime.now()
        tomorrow = now.date() + datetime.timedelta(days=1)
        midnight = datetime.datetime.combine(tomorrow, datetime.time.min) 
        time_until_midnight_ms = int((midnight - now).total_seconds() * 1000)
        trigger_delay_ms = time_until_midnight_ms + 1000 
        
        print(f"Scheduling next 'New Day' log in {(trigger_delay_ms / 1000 / 3600):.2f} hours.")
        self._new_day_timer_id = self.master.after(trigger_delay_ms, self.trigger_new_day)

    def trigger_new_day(self):
        print("--- Triggering Automatic New Day Log ---")
        self.log_new_day(button_widget=None) 
        self.schedule_new_day() 


# --- Settings Window Class (incorporates previous improvements) ---
class SettingsWindow:
    def __init__(self, master, parent_gui):
        self.master = master
        self.parent_gui = parent_gui
        self.master.title("Settings")
        self.master.geometry("1000x680") # Slightly increased height
        self.master.minsize(750, 550) # Slightly increased min size
        self.style = parent_gui.style 

        self.main_frame = ttk.Frame(self.master)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.main_frame.rowconfigure(0, weight=1); self.main_frame.columnconfigure(0, weight=1)

        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.grid(row=0, column=0, sticky="nsew", padx=5, pady=5) # Added padding to notebook

        self.create_general_tab()
        self.create_txt_column_mapping_tab()
        self.create_folder_selection_tab()
        self.create_custom_buttons_tab()
        self.create_sqlite_tab()

        button_frame = ttk.Frame(self.main_frame)
        button_frame.grid(row=1, column=0, pady=(15, 5), sticky="e") # Increased top padding
        ttk.Button(button_frame, text="Save and Close", command=self.save_and_close, style="Accent.TButton").pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.master.destroy).pack(side=tk.RIGHT)

    def save_and_close(self):
        print("Saving settings from dialog...")
        self.save_settings() 
        print("Settings saved. Closing dialog.")
        self.master.destroy()

    def create_general_tab(self):
        tab = ttk.Frame(self.notebook, padding=20); self.notebook.add(tab, text="General")
        log_frame = ttk.LabelFrame(tab, text="Excel Log File (.xlsx)", padding=15); log_frame.pack(fill="x", pady=(5, 15)); log_frame.columnconfigure(1, weight=1) # Added top pady
        self.log_file_label = ttk.Label(log_frame, text="Path:", anchor='e'); self.log_file_label.grid(row=0, column=0, padx=(0, 5), pady=5, sticky='w')
        self.log_file_entry = ttk.Entry(log_frame, width=80); self.log_file_entry.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        log_browse_btn = ttk.Button(log_frame, text="Browse...", command=self.select_excel_file); log_browse_btn.grid(row=0, column=2, padx=(5, 0), pady=5)
        ToolTip(log_browse_btn, "Select the main Excel file for logging."); ToolTip(self.log_file_entry, "Path to the .xlsx file where logs will be written.")
        
        txt_frame = ttk.LabelFrame(tab, text="Navigation TXT Data Folder", padding=15); txt_frame.pack(fill="x", pady=5); txt_frame.columnconfigure(1, weight=1) # Added pady
        self.txt_folder_label = ttk.Label(txt_frame, text="Folder:", anchor='e'); self.txt_folder_label.grid(row=0, column=0, padx=(0, 5), pady=5, sticky='w')
        self.txt_folder_entry = ttk.Entry(txt_frame, width=80); self.txt_folder_entry.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        txt_browse_btn = ttk.Button(txt_frame, text="Browse...", command=self.select_txt_folder); txt_browse_btn.grid(row=0, column=2, padx=(5, 0), pady=5)
        ToolTip(txt_browse_btn, "Select the folder containing navigation TXT files (e.g., P190)."); ToolTip(self.txt_folder_entry, "Path to the folder containing navigation TXT files. The latest .txt file in this folder will be read.")

    def select_excel_file(self):
        initial_dir = os.path.dirname(self.log_file_entry.get()) if self.log_file_entry.get() else "/"
        file_path = filedialog.askopenfilename(initialdir=initial_dir, filetypes=[("Excel files", "*.xlsx")], parent=self.master, title="Select Excel Log File")
        if file_path: self.log_file_entry.delete(0, tk.END); self.log_file_entry.insert(0, file_path)

    def select_txt_folder(self):
        initial_dir = self.txt_folder_entry.get() if self.txt_folder_entry.get() else "/"
        folder_path = filedialog.askdirectory(initialdir=initial_dir, parent=self.master, title="Select Navigation TXT Folder")
        if folder_path: self.txt_folder_entry.delete(0, tk.END); self.txt_folder_entry.insert(0, folder_path)

    def create_txt_column_mapping_tab(self):
        tab = ttk.Frame(self.notebook, padding=20); self.notebook.add(tab, text="TXT Columns")
        self.txt_field_column_widgets = {}; self.txt_field_skip_vars = {}
        fields = ["Date", "Time", "KP", "DCC", "Line name", "Latitude", "Longitude", "Easting", "Northing"]; event_field = "Event"
        
        header_frame = ttk.Frame(tab, style="Header.TFrame", padding=(5,3)); header_frame.pack(fill='x', pady=(0,10))
        ttk.Label(header_frame, text="TXT Field", font=("Arial", 10, "bold"), width=15).pack(side=tk.LEFT, padx=5)
        ttk.Label(header_frame, text="Target Excel/DB Column Name", font=("Arial", 10, "bold"), width=30).pack(side=tk.LEFT, padx=5)
        ttk.Label(header_frame, text="Skip This Field?", font=("Arial", 10, "bold"), width=15).pack(side=tk.LEFT, padx=5)
        
        rows_frame = ttk.Frame(tab); rows_frame.pack(fill='both', expand=True); rows_frame.columnconfigure(1, weight=1)
        for i, field in enumerate(fields):
            row_f = ttk.Frame(rows_frame); row_f.grid(row=i, column=0, sticky='ew', pady=3) # Increased pady
            ttk.Label(row_f, text=f"{field}:", width=15).pack(side=tk.LEFT, padx=5)
            entry = ttk.Entry(row_f, width=30); entry.pack(side=tk.LEFT, padx=5, fill='x', expand=True)
            ToolTip(entry, f"Enter the exact column name in your Excel/DB where '{field}' data should be written.")
            self.txt_field_column_widgets[field] = entry
            skip_var = tk.BooleanVar(); skip_checkbox = ttk.Checkbutton(row_f, variable=skip_var, text=""); skip_checkbox.pack(side=tk.LEFT, padx=20)
            ToolTip(skip_checkbox, f"Check this box to ignore the '{field}' field from the TXT file.")
            self.txt_field_skip_vars[field] = skip_var
        
        event_row_index = len(fields); row_f = ttk.Frame(rows_frame); row_f.grid(row=event_row_index, column=0, sticky='ew', pady=3) # Increased pady
        ttk.Label(row_f, text=f"{event_field}:", width=15, font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
        entry = ttk.Entry(row_f, width=30); entry.pack(side=tk.LEFT, padx=5, fill='x', expand=True); entry.insert(0, "Event") 
        ToolTip(entry, "Enter the exact column name where event text (e.g., 'Log on', 'Custom Event X') should be written.")
        self.txt_field_column_widgets[event_field] = entry
        skip_var = tk.BooleanVar(); skip_checkbox = ttk.Checkbutton(row_f, variable=skip_var, text=""); skip_checkbox.pack(side=tk.LEFT, padx=20)
        ToolTip(skip_checkbox, "Check this to prevent writing any event text to the specified column.")
        self.txt_field_skip_vars[event_field] = skip_var

    def create_folder_selection_tab(self):
        tab = ttk.Frame(self.notebook); self.notebook.add(tab, text="Monitored Folders")
        self.folder_canvas = tk.Canvas(tab, borderwidth=0, background="#ffffff")
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=self.folder_canvas.yview)
        self.scrollable_frame = ttk.Frame(self.folder_canvas, style="Row0.TFrame") 
        self.scrollable_frame.bind("<Configure>", lambda e: self.folder_canvas.configure(scrollregion=self.folder_canvas.bbox("all")))
        self.folder_canvas_window = self.folder_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.folder_canvas.configure(yscrollcommand=scrollbar.set)
        self.folder_canvas.pack(side="left", fill="both", expand=True, padx=(10,0), pady=10)
        scrollbar.pack(side="right", fill="y", padx=(0,10), pady=10)
        def _on_mousewheel(event): 
            delta = 0
            if event.num == 4: delta = -1 
            elif event.num == 5: delta = 1  
            elif hasattr(event, 'delta') and event.delta != 0 : delta = -int(event.delta / abs(event.delta) if abs(event.delta) >=120 else event.delta/3) 
            if delta !=0: self.folder_canvas.yview_scroll(delta, "units")
        self.folder_canvas.bind_all("<MouseWheel>", _on_mousewheel); self.folder_canvas.bind_all("<Button-4>", _on_mousewheel); self.folder_canvas.bind_all("<Button-5>", _on_mousewheel)
        self.folder_entries = {}; self.folder_column_entries = {}; self.folder_skip_vars = {}; self.file_extension_entries = {}; self.folder_row_widgets = {}
        self.add_folder_header(self.scrollable_frame)

    def add_folder_header(self, parent):
        header_frame = ttk.Frame(parent, style="Header.TFrame", padding=(5,3)); header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 5)); header_frame.columnconfigure(1, weight=1)
        ttk.Label(header_frame, text="Folder Type", width=15, anchor="w", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=(5,0))
        ttk.Label(header_frame, text="Monitor Path", anchor="w", font=("Arial", 10, "bold")).grid(row=0, column=1, padx=5, sticky='w')
        ttk.Label(header_frame, text="...", width=4, anchor="center", font=("Arial", 10, "bold")).grid(row=0, column=2, padx=1) 
        ttk.Label(header_frame, text="Target Column", width=20, anchor="w", font=("Arial", 10, "bold")).grid(row=0, column=3, padx=5, sticky='w')
        ttk.Label(header_frame, text="File Ext.", width=10, anchor="w", font=("Arial", 10, "bold")).grid(row=0, column=4, padx=5, sticky='w')
        ttk.Label(header_frame, text="Skip?", width=5, anchor="center", font=("Arial", 10, "bold")).grid(row=0, column=5, padx=(10,5), sticky='w')

    def add_initial_folder_rows(self): 
        for name, frame in list(self.folder_row_widgets.items()):
            if frame and frame.winfo_exists(): frame.destroy()
        self.folder_row_widgets.clear(); self.folder_entries.clear(); self.folder_column_entries.clear(); self.file_extension_entries.clear(); self.folder_skip_vars.clear()

        default_folders = ["Qinsy DB", "Naviscan", "SIS", "SSS", "SBP", "Mag", "Grad", "SVP", "SpintINS", "Video", "Cathx", "Hypack RAW", "Eiva NaviPac"]
        loaded_paths = self.parent_gui.folder_paths; all_folder_names = []; processed = set()
        for name in default_folders: 
            if name not in processed: all_folder_names.append(name); processed.add(name)
        for name in loaded_paths: 
            if name not in processed: all_folder_names.append(name); processed.add(name)
        
        for folder_name in all_folder_names:
            self.add_folder_row(folder_name=folder_name, 
                                folder_path=loaded_paths.get(folder_name, ""),
                                column_name=self.parent_gui.folder_columns.get(folder_name, folder_name), 
                                extension=self.parent_gui.file_extensions.get(folder_name, ""),
                                skip=self.parent_gui.folder_skips.get(folder_name, False))
        self.master.after_idle(self.update_scroll_region) 

    def add_folder_row(self, folder_name="", folder_path="", column_name="", extension="", skip=False):
        row_index = len(self.folder_row_widgets) + 1 
        style_name = f"Row{row_index % 2}.TFrame" 
        try: self.style.configure(style_name) 
        except tk.TclError: 
            bg = "#ffffff" if row_index % 2 == 0 else "#f5f5f5"
            self.style.configure(style_name, background=bg)
        
        row_frame = ttk.Frame(self.scrollable_frame, style=style_name, padding=(0, 2))
        row_frame.grid(row=row_index, column=0, sticky="ew", pady=(2,3)) # Changed pady for more space
        row_frame.columnconfigure(1, weight=1) 

        label_style = style_name.replace("Frame", "Label") 
        try: self.style.configure(label_style, background=self.style.lookup(style_name, 'background'))
        except: pass 

        label = ttk.Label(row_frame, text=f"{folder_name}:", width=15, anchor='w', style=label_style)
        label.grid(row=0, column=0, padx=(5,0), pady=1, sticky="w")
        
        entry = ttk.Entry(row_frame, width=50); entry.insert(0, folder_path)
        entry.grid(row=0, column=1, padx=5, pady=1, sticky="ew")
        ToolTip(entry, f"Enter the full path to the '{folder_name}' data folder.")
        
        def select_folder(e=entry, fn=folder_name): 
            current_path = e.get()
            initial = current_path if os.path.isdir(current_path) else (os.path.dirname(current_path) if current_path else "/")
            folder = filedialog.askdirectory(parent=self.master, initialdir=initial, title=f"Select Folder for {fn}")
            if folder: e.delete(0, tk.END); e.insert(0, folder)
        
        button = ttk.Button(row_frame, text="...", width=3, command=select_folder, style="Toolbutton")
        button.grid(row=0, column=2, padx=(0,5), pady=1)
        ToolTip(button, "Browse for the folder.")

        column_entry = ttk.Entry(row_frame, width=20); column_entry.insert(0, column_name if column_name else folder_name)
        column_entry.grid(row=0, column=3, padx=5, pady=1, sticky="w")
        ToolTip(column_entry, f"Enter the Excel/DB column name for the latest '{folder_name}' filename.")

        extension_entry = ttk.Entry(row_frame, width=10); extension_entry.insert(0, extension)
        extension_entry.grid(row=0, column=4, padx=5, pady=1, sticky="w")
        ToolTip(extension_entry, f"Optional: Monitor only files ending with this extension (e.g., 'svp', 'log'). Leave blank for any file.")

        skip_var = tk.BooleanVar(value=skip)
        skip_checkbox = ttk.Checkbutton(row_frame, variable=skip_var) 
        skip_checkbox.grid(row=0, column=5, padx=(15,5), pady=1, sticky="w")
        ToolTip(skip_checkbox, f"Check to disable monitoring for the '{folder_name}' folder.")

        self.folder_entries[folder_name] = entry
        self.folder_column_entries[folder_name] = column_entry
        self.file_extension_entries[folder_name] = extension_entry
        self.folder_skip_vars[folder_name] = skip_var
        self.folder_row_widgets[folder_name] = row_frame 

    def update_scroll_region(self):
        self.scrollable_frame.update_idletasks() 
        self.folder_canvas.configure(scrollregion=self.folder_canvas.bbox("all"))
        self.folder_canvas.itemconfigure(self.folder_canvas_window, width=self.scrollable_frame.winfo_width())


    def create_custom_buttons_tab(self):
        tab = ttk.Frame(self.notebook, padding=20); self.notebook.add(tab, text="Button & Row Colors") 
        self.pastel_colors = ["#FFB3BA", "#FFDFBA", "#FFFFBA", "#BAFFC9", "#BAE1FF", "#E0BBE4", "#FFC8A2", "#D4A5A5", "#A2D4AB", "#A2C4D4"]

        std_color_frame = ttk.LabelFrame(tab, text="Standard Action Row Colors (Excel Output)", padding=10)
        std_color_frame.pack(pady=(5, 15), fill='x', anchor='n') 
        std_header = ttk.Frame(std_color_frame, style="Header.TFrame", padding=(5,3))
        std_header.pack(fill='x', pady=(0, 5))
        ttk.Label(std_header, text="Action", font=("Arial", 10, "bold"), width=15).pack(side=tk.LEFT, padx=(5,0))
        ttk.Label(std_header, text="Row Color", font=("Arial", 10, "bold"), width=25).pack(side=tk.LEFT, padx=5)
        self.standard_button_color_widgets = {} 
        standard_buttons_to_color = ["Log on", "Log off", "Event", "SVP", "New Day"] 
        for i, btn_name in enumerate(standard_buttons_to_color):
            style_name = f"Row{i % 2}.TFrame"; row_frame = ttk.Frame(std_color_frame, style=style_name, padding=(0, 3)); row_frame.pack(fill='x', pady=0) 
            ttk.Label(row_frame, text=f"{btn_name}:", width=15, style=style_name.replace("Frame","Label")).pack(side=tk.LEFT, padx=(5,0))
            color_widget_frame = ttk.Frame(row_frame, style=style_name); color_widget_frame.pack(side=tk.LEFT, padx=5)
            initial_color = self.parent_gui.button_colors.get(btn_name, (None, None))[1]
            selected_color_var = tk.StringVar(value=initial_color if initial_color else "")
            color_display_label = tk.Label(color_widget_frame, width=4, relief="solid", borderwidth=1)
            color_display_label.pack(side="left", padx=(0, 5))
            try: color_display_label.config(background=initial_color if initial_color else self.master.option_get('background', '.'))
            except tk.TclError: color_display_label.config(background=self.master.option_get('background', '.'))
            
            def _local_set_color(color_hex, var=selected_color_var, label=color_display_label):
                _valid_color = None; _temp_label = None
                if color_hex:
                    try: _temp_label = tk.Label(self.master); _temp_label.config(background=color_hex); _valid_color = color_hex
                    except tk.TclError: print(f"Warning: Invalid color code '{color_hex}'."); _valid_color = None
                    finally: 
                        if _temp_label: 
                            try: _temp_label.destroy()
                            except tk.TclError: pass
                var.set(_valid_color if _valid_color else "")
                try: label.config(background=_valid_color if _valid_color else self.master.option_get('background', '.'))
                except tk.TclError: label.config(background=self.master.option_get('background', '.'))

            def _local_choose_color(var=selected_color_var, label=color_display_label, name=btn_name):
                _current_color = var.get()
                _color_code = colorchooser.askcolor(color=_current_color if _current_color else None, title=f"Choose Row Color for {name}", parent=self.master)
                if _color_code and _color_code[1]: _local_set_color(_color_code[1], var, label)

            clear_btn = ttk.Button(color_widget_frame, text="X", width=2, style="Toolbutton", command=lambda v=selected_color_var, l=color_display_label: _local_set_color(None, v, l))
            clear_btn.pack(side="left", padx=1); ToolTip(clear_btn, f"Clear row color for {btn_name}.")
            presets_frame = ttk.Frame(color_widget_frame, style=style_name); presets_frame.pack(side="left", padx=(2, 2))
            for p_color in self.pastel_colors[:5]:
                try:
                    b = tk.Button(presets_frame, bg=p_color, width=1, height=1, relief="raised", bd=1, command=lambda c=p_color, v=selected_color_var, l=color_display_label: _local_set_color(c, v, l))
                    b.pack(side=tk.LEFT, padx=1)
                except tk.TclError: pass
            choose_btn = ttk.Button(color_widget_frame, text="...", width=3, style="Toolbutton", command=lambda v=selected_color_var, l=color_display_label, n=btn_name: _local_choose_color(v, l, n))
            choose_btn.pack(side="left", padx=1); ToolTip(choose_btn, f"Choose a custom row color for {btn_name}.")
            self.standard_button_color_widgets[btn_name] = (selected_color_var, color_display_label)

        custom_frame = ttk.LabelFrame(tab, text="Custom Button Configuration", padding=10)
        custom_frame.pack(pady=10, fill='both', expand=True)
        num_buttons_frame = ttk.Frame(custom_frame); num_buttons_frame.pack(pady=(5,10), anchor='w') 
        ttk.Label(num_buttons_frame, text="Number of Custom Buttons (0-10):").pack(side='left', padx=5)
        self.num_buttons_entry = ttk.Entry(num_buttons_frame, width=5); self.num_buttons_entry.pack(side='left', padx=5); ToolTip(self.num_buttons_entry, "Enter the number of custom event buttons needed (max 10).")
        update_btn = ttk.Button(num_buttons_frame, text="Update Count", command=self.update_num_custom_buttons); update_btn.pack(side='left', padx=5); ToolTip(update_btn, "Update the list below to show the specified number of button configurations.")
        
        header_frame = ttk.Frame(custom_frame, style="Header.TFrame", padding=(5,3)); header_frame.pack(fill='x', pady=(5,5)) 
        ttk.Label(header_frame, text="Button #", font=("Arial", 10, "bold"), width=7).pack(side=tk.LEFT, padx=(5,0))
        ttk.Label(header_frame, text="Button Text", font=("Arial", 10, "bold"), width=25).pack(side=tk.LEFT, padx=5)
        ttk.Label(header_frame, text="Event Text (for Log)", font=("Arial", 10, "bold"), width=35).pack(side=tk.LEFT, padx=5, expand=True, fill='x')
        ttk.Label(header_frame, text="Button/Row Color", font=("Arial", 10, "bold"), width=25).pack(side=tk.LEFT, padx=5)
        
        self.custom_button_entries_frame = ttk.Frame(custom_frame); self.custom_button_entries_frame.pack(pady=0, fill='both', expand=True)
        self.custom_button_widgets = [] 
        
    def update_num_custom_buttons(self):
        try:
            num_buttons = int(self.num_buttons_entry.get())
            if not (0 <= num_buttons <= 10): raise ValueError("Number must be between 0 and 10")
            if self.parent_gui.num_custom_buttons != num_buttons:
                print(f"Updating number of custom buttons to {num_buttons}")
                self.parent_gui.num_custom_buttons = num_buttons
                current_configs = list(self.parent_gui.custom_button_configs) 

                if num_buttons < len(current_configs):
                    self.parent_gui.custom_button_configs = current_configs[:num_buttons]
                else: 
                    for i in range(len(current_configs), num_buttons):
                        new_text = f"Custom {i + 1}"
                        self.parent_gui.custom_button_configs.append({
                            "text": new_text,
                            "event_text": f"{new_text} Event"
                        })
                        if new_text not in self.parent_gui.button_colors:
                            self.parent_gui.button_colors[new_text] = (None, None)
                
                self.recreate_custom_button_settings() 
        except ValueError as e:
            messagebox.showerror("Invalid Number", f"Please enter a whole number between 0 and 10. Error: {e}", parent=self.master)
            self.num_buttons_entry.delete(0, tk.END); self.num_buttons_entry.insert(0, str(self.parent_gui.num_custom_buttons)) 

    def recreate_custom_button_settings(self):
        for widget in self.custom_button_entries_frame.winfo_children(): widget.destroy()
        self.custom_button_widgets = [] 

        num_display_buttons = self.parent_gui.num_custom_buttons 
        configs_to_display = self.parent_gui.custom_button_configs[:num_display_buttons]

        for i in range(num_display_buttons):
            config = configs_to_display[i] if i < len(configs_to_display) else {}
            initial_text = config.get("text", f"Custom {i+1}")
            initial_event = config.get("event_text", f"{initial_text} Event")
            _, initial_color = self.parent_gui.button_colors.get(initial_text, (None, None))

            style_name = f"Row{i % 2}.TFrame"; row_frame = ttk.Frame(self.custom_button_entries_frame, style=style_name, padding=(0, 3)); row_frame.pack(fill='x', pady=0) 
            ttk.Label(row_frame, text=f"{i+1}", width=7, style=style_name.replace("Frame","Label")).pack(side=tk.LEFT, padx=(5,0))
            text_entry = ttk.Entry(row_frame, width=25); text_entry.insert(0, initial_text); text_entry.pack(side=tk.LEFT, padx=5); ToolTip(text_entry, "Text displayed on the button.")
            event_entry = ttk.Entry(row_frame); event_entry.insert(0, initial_event); event_entry.pack(side=tk.LEFT, padx=5, fill='x', expand=True); ToolTip(event_entry, "Text written to the 'Event' column in the log.")
            
            color_frame = ttk.Frame(row_frame, style=style_name); color_frame.pack(side=tk.LEFT, padx=5)
            selected_color_var = tk.StringVar(value=initial_color if initial_color else "")
            color_display_label = tk.Label(color_frame, width=4, relief="solid", borderwidth=1)
            color_display_label.pack(side="left", padx=(0, 5))
            try: color_display_label.config(background=initial_color if initial_color else self.master.option_get('background', '.'))
            except tk.TclError: color_display_label.config(background=self.master.option_get('background', '.'))

            def _local_set_color(color_hex, var=selected_color_var, label=color_display_label):
                _valid_color = None; _temp_label = None
                if color_hex:
                    try: _temp_label = tk.Label(self.master); _temp_label.config(background=color_hex); _valid_color = color_hex
                    except tk.TclError: _valid_color = None
                    finally: 
                        if _temp_label: 
                            try: _temp_label.destroy()
                            except tk.TclError: pass
                var.set(_valid_color if _valid_color else "")
                try: label.config(background=_valid_color if _valid_color else self.master.option_get('background', '.'))
                except tk.TclError: label.config(background=self.master.option_get('background', '.'))

            def _local_choose_color(var=selected_color_var, label=color_display_label, txt_widget=text_entry, btn_idx=i):
                _current_color = var.get()
                _btn_name = txt_widget.get().strip() or f"Button {btn_idx+1}"
                _color_code = colorchooser.askcolor(color=_current_color if _current_color else None, title=f"Choose Color for {_btn_name}", parent=self.master)
                if _color_code and _color_code[1]: _local_set_color(_color_code[1], var, label)

            clear_btn = ttk.Button(color_frame, text="X", width=2, style="Toolbutton", command=lambda v=selected_color_var, l=color_display_label: _local_set_color(None, v, l))
            clear_btn.pack(side="left", padx=1); ToolTip(clear_btn, "Clear color.")
            presets_frame = ttk.Frame(color_frame, style=style_name); presets_frame.pack(side="left", padx=(2, 2))
            for p_color in self.pastel_colors[:5]:
                try:
                    b = tk.Button(presets_frame, bg=p_color, width=1, height=1, relief="raised", bd=1, command=lambda c=p_color, v=selected_color_var, l=color_display_label: _local_set_color(c, v, l))
                    b.pack(side=tk.LEFT, padx=1)
                except tk.TclError: pass
            choose_btn = ttk.Button(color_frame, text="...", width=3, style="Toolbutton", command=lambda v=selected_color_var, l=color_display_label, t=text_entry, b_idx=i: _local_choose_color(v, l, t, b_idx))
            choose_btn.pack(side="left", padx=1); ToolTip(choose_btn, "Choose a custom color.")
            self.custom_button_widgets.append( (text_entry, event_entry, selected_color_var, color_display_label) )

    def create_sqlite_tab(self):
        tab = ttk.Frame(self.notebook, padding=20); self.notebook.add(tab, text="SQLite Log")
        enable_frame = ttk.Frame(tab); enable_frame.pack(fill='x', pady=(5, 20)) # Increased bottom pady
        self.sqlite_enabled_var = tk.BooleanVar()
        enable_check = ttk.Checkbutton(enable_frame, text="Enable SQLite Database Logging", variable=self.sqlite_enabled_var, style="Large.TCheckbutton"); enable_check.pack(side=tk.LEFT, pady=(5, 10)); ToolTip(enable_check, "Check to enable logging events to an SQLite database file.")
        
        config_frame = ttk.LabelFrame(tab, text="SQLite Configuration", padding=15); config_frame.pack(fill='x', pady=5); config_frame.columnconfigure(1, weight=1) 
        ttk.Label(config_frame, text="Database File (.db):").grid(row=0, column=0, padx=5, pady=8, sticky="w") 
        self.sqlite_db_path_entry = ttk.Entry(config_frame, width=70); self.sqlite_db_path_entry.grid(row=0, column=1, padx=5, pady=8, sticky="ew"); ToolTip(self.sqlite_db_path_entry, "Path to the SQLite database file. Will be created if it doesn't exist.")
        db_browse_btn = ttk.Button(config_frame, text="Browse/Create...", command=self.select_sqlite_file); db_browse_btn.grid(row=0, column=2, padx=5, pady=8); ToolTip(db_browse_btn, "Browse for an existing SQLite file or specify a name/location for a new one.")
        
        ttk.Label(config_frame, text="Table Name:").grid(row=1, column=0, padx=5, pady=8, sticky="w") 
        self.sqlite_table_entry = ttk.Entry(config_frame, width=40); self.sqlite_table_entry.grid(row=1, column=1, padx=5, pady=8, sticky="w"); ToolTip(self.sqlite_table_entry, "The name of the table within the database where logs will be written (e.g., 'EventLog'). It must exist.")
        
        test_button = ttk.Button(config_frame, text="Test Connection & Table", command=self.test_sqlite_connection); test_button.grid(row=2, column=1, padx=5, pady=(15,5), sticky="w"); ToolTip(test_button, "Verify connection to the DB file and check if the table exists.") 
        self.test_result_label = ttk.Label(config_frame, text="", font=("Arial", 9), wraplength=500); self.test_result_label.grid(row=3, column=0, columnspan=3, padx=5, pady=2, sticky="w")

    def select_sqlite_file(self):
        filetypes = [("SQLite Database", "*.db"), ("SQLite Database", "*.sqlite"), ("SQLite3 Database", "*.sqlite3"), ("All Files", "*.*")]
        current_path = self.sqlite_db_path_entry.get(); initial_dir = os.path.dirname(current_path) if current_path else "."
        filepath = filedialog.asksaveasfilename(parent=self.master, title="Select or Create SQLite Database File", initialdir=initial_dir, initialfile="DataLoggerLog.db", filetypes=filetypes, defaultextension=".db")
        if filepath: self.sqlite_db_path_entry.delete(0, tk.END); self.sqlite_db_path_entry.insert(0, filepath); print(f"SQLite DB file selected/specified: {filepath}");
        if hasattr(self, 'test_result_label'): self.test_result_label.config(text="") 

    def test_sqlite_connection(self):
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
                if "no such table" in str(e_table).lower(): print(f"Table '{table_name}' not found."); result_text += f"⚠️ Warning: Table '{table_name}' not found. It needs to be created."; result_color = "#E67E00" 
                else: raise e_table 
        except sqlite3.Error as e: print(f"SQLite Test Error: {e}"); result_text = f"❌ Error connecting/checking DB: {e}"; result_color = "red"
        except Exception as e: print(f"Unexpected Test Error: {e}"); result_text = f"❌ Unexpected Error: {e}"; result_color = "red"
        finally:
            if conn: conn.close(); print("Connection closed.")
            self.test_result_label.config(text=result_text, foreground=result_color)
            self.master.after(15000, lambda: self.test_result_label.config(text="")) 

    def save_settings(self): 
        self.parent_gui.log_file_path = self.log_file_entry.get().strip()
        self.parent_gui.txt_folder_path = self.txt_folder_entry.get().strip()
        parent_txt_cols = {}; parent_txt_skips = {}
        for field, entry in self.txt_field_column_widgets.items(): parent_txt_cols[field] = entry.get().strip() or field 
        for field, var in self.txt_field_skip_vars.items(): parent_txt_skips[field] = var.get()
        self.parent_gui.txt_field_columns = parent_txt_cols; self.parent_gui.txt_field_skips = parent_txt_skips
        parent_folder_paths = {}; parent_folder_cols = {}; parent_folder_exts = {}; parent_folder_skips = {}
        for folder_name, entry_widget in self.folder_entries.items():
            folder_path = entry_widget.get().strip()
            if folder_path: 
                parent_folder_paths[folder_name] = folder_path
                col_entry = self.folder_column_entries.get(folder_name); ext_entry = self.file_extension_entries.get(folder_name); skip_var = self.folder_skip_vars.get(folder_name)
                parent_folder_cols[folder_name] = col_entry.get().strip() if col_entry and col_entry.get().strip() else folder_name
                parent_folder_exts[folder_name] = ext_entry.get().strip().lstrip('.') if ext_entry else ""
                parent_folder_skips[folder_name] = skip_var.get() if skip_var else False
            else: 
                for d in [self.parent_gui.folder_paths, self.parent_gui.folder_columns, self.parent_gui.file_extensions, self.parent_gui.folder_skips]:
                    d.pop(folder_name, None)
        self.parent_gui.folder_paths = parent_folder_paths; self.parent_gui.folder_columns = parent_folder_cols; self.parent_gui.file_extensions = parent_folder_exts; self.parent_gui.folder_skips = parent_folder_skips

        parent_custom_configs = []
        new_button_colors_for_parent = self.parent_gui.button_colors.copy() 

        for std_btn_name, (color_var, _) in self.standard_button_color_widgets.items():
            color_hex = color_var.get()
            new_button_colors_for_parent[std_btn_name] = (None, color_hex if color_hex else None)
        
        for i, (text_widget, event_widget, color_var, _) in enumerate(self.custom_button_widgets):
            text = text_widget.get().strip(); event_text = event_widget.get().strip(); color_hex = color_var.get()
            default_text = f"Custom {i + 1}"; final_text = text if text else default_text; final_event_text = event_text if event_text else f"{final_text} Triggered"
            
            if i < len(self.parent_gui.custom_button_configs):
                old_cfg_text = self.parent_gui.custom_button_configs[i].get("text")
                if old_cfg_text and old_cfg_text != final_text and old_cfg_text in new_button_colors_for_parent:
                    del new_button_colors_for_parent[old_cfg_text]
            
            parent_custom_configs.append({"text": final_text, "event_text": final_event_text})
            new_button_colors_for_parent[final_text] = (None, color_hex if color_hex else None)
        
        self.parent_gui.custom_button_configs = parent_custom_configs
        self.parent_gui.button_colors = new_button_colors_for_parent 

        self.parent_gui.sqlite_enabled = self.sqlite_enabled_var.get()
        self.parent_gui.sqlite_db_path = self.sqlite_db_path_entry.get().strip()
        self.parent_gui.sqlite_table = self.sqlite_table_entry.get().strip() or "EventLog"

        self.parent_gui.save_settings() 
        self.parent_gui.update_custom_buttons() 
        self.parent_gui.start_monitoring() 
        self.parent_gui.update_db_indicator()

    def load_settings(self): 
        self.log_file_entry.delete(0, tk.END); self.log_file_entry.insert(0, self.parent_gui.log_file_path or "")
        self.txt_folder_entry.delete(0, tk.END); self.txt_folder_entry.insert(0, self.parent_gui.txt_folder_path or "")
        for field, entry in self.txt_field_column_widgets.items():
            entry.delete(0, tk.END)
            default_val = "Event" if field == "Event" else field 
            entry.insert(0, self.parent_gui.txt_field_columns.get(field, default_val))
        for field, var in self.txt_field_skip_vars.items():
            var.set(self.parent_gui.txt_field_skips.get(field, False))
        self.add_initial_folder_rows() 

        self.num_buttons_entry.delete(0, tk.END)
        self.num_buttons_entry.insert(0, str(self.parent_gui.num_custom_buttons))
        self.recreate_custom_button_settings() 

        for btn_name, (color_var, display_label) in self.standard_button_color_widgets.items():
            loaded_color_hex = self.parent_gui.button_colors.get(btn_name, (None, None))[1]
            color_var.set(loaded_color_hex if loaded_color_hex else "")
            try: display_label.config(background=loaded_color_hex if loaded_color_hex else self.master.option_get('background', '.'))
            except tk.TclError: display_label.config(background=self.master.option_get('background', '.'))
        
        self.sqlite_enabled_var.set(self.parent_gui.sqlite_enabled)
        self.sqlite_db_path_entry.delete(0, tk.END); self.sqlite_db_path_entry.insert(0, self.parent_gui.sqlite_db_path or "")
        self.sqlite_table_entry.delete(0, tk.END); self.sqlite_table_entry.insert(0, self.parent_gui.sqlite_table or "EventLog")
        if hasattr(self, 'test_result_label'): self.test_result_label.config(text="")


# --- Main Execution ---
if __name__ == "__main__":
    root = tk.Tk()
    # Optional: Set App icon (replace 'your_icon.ico' or handle .png)
    # try:
    #     # For Windows .ico
    #     # icon_path_ico = 'your_icon.ico' 
    #     # if os.path.exists(icon_path_ico): root.iconbitmap(icon_path_ico)
    #     
    #     # For cross-platform .png (requires PhotoImage)
    #     # icon_path_png = 'your_icon.png'
    #     # if os.path.exists(icon_path_png):
    #     #     img = tk.PhotoImage(file=icon_path_png)
    #     #     root.tk.call('wm', 'iconphoto', root._w, img)
    # except Exception as e:
    #     print(f"Icon Error: {e}")

    gui = DataLoggerGUI(root) 

    def on_closing():
        print("Closing application requested...")
        print("Shutting down monitors...")
        active_monitors = list(gui.monitors.items()) 
        if not active_monitors:
            print("No monitors were active.")
        else:
            for name, monitor_observer in active_monitors:
                try:
                    if monitor_observer.is_alive():
                        monitor_observer.stop()
                        print(f"Stopped monitor '{name}'.")
                except Exception as e:
                    print(f"Error stopping monitor '{name}': {e}")
            
            for name, monitor_observer in active_monitors: # Separate loop for join
                try:
                    if monitor_observer.is_alive(): 
                        monitor_observer.join(timeout=0.5) # Wait briefly
                        if monitor_observer.is_alive():
                            print(f"Warning: Monitor thread '{name}' did not stop gracefully.")
                except Exception as e:
                    print(f"Error joining monitor thread '{name}': {e}")
                finally: 
                    if name in gui.monitors:
                        del gui.monitors[name] # Clean up dict
        
        print("Monitors shut down. Exiting.")
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()
