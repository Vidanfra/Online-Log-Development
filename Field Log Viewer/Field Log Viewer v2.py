import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser
import sqlite3
import xlwings as xw
import threading
import json
import os
import time

# --- Constants (Modified) ---
EVENT_COLUMN_NAME = 'Event'

# Get the directory where this script is located
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# Settings file should be in the same directory as the script
SETTINGS_FILE = os.path.join(SCRIPT_DIR, 'flv_settings/flv_settings.json')

# TABLE_NAME is now dynamic, but we keep the old name for status/settings defaults
DEFAULT_TABLE_NAME = 'DailyLog-Horizon_v14' 
UNIQUE_ID_COLUMN = 'index' 
MAX_FETCH_RETRIES = 5



class ColumnSelector(tk.Toplevel):
    """A modal dialog window for selecting columns from a list."""
    def __init__(self, parent, all_columns, selected_columns):
        super().__init__(parent)
        self.title("Select Columns to Import")
        self.geometry("400x500")

        self.transient(parent)
        self.grab_set()

        self.result = None
        self.vars = {col: tk.BooleanVar(value=(col in selected_columns)) for col in all_columns}

        # --- Main frame ---
        main_frame = tk.Frame(self, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Controls
        controls_frame = tk.Frame(main_frame)
        controls_frame.pack(fill=tk.X, side=tk.TOP, pady=(0, 10))
        tk.Button(controls_frame, text="Select All", command=self._select_all).pack(side=tk.LEFT)
        tk.Button(controls_frame, text="Deselect All", command=self._deselect_all).pack(side=tk.LEFT, padx=10)

        # Bottom buttons
        bottom_frame = tk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(10, 0))
        tk.Button(bottom_frame, text="OK", command=self._on_ok, width=10).pack(side=tk.RIGHT)
        tk.Button(bottom_frame, text="Cancel", command=self.destroy, width=10).pack(side=tk.RIGHT, padx=10)

        # Scrollable Checkbox Area
        scroll_area_frame = tk.Frame(main_frame)
        scroll_area_frame.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(scroll_area_frame, borderwidth=0)
        checkbox_frame = tk.Frame(canvas)
        scrollbar = tk.Scrollbar(scroll_area_frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        canvas.create_window((4,4), window=checkbox_frame, anchor="nw")

        checkbox_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        for col in all_columns:
            tk.Checkbutton(checkbox_frame, text=col, variable=self.vars[col]).pack(anchor='w')
        
        self.protocol("WM_DELETE_WINDOW", self.destroy)

    def _select_all(self):
        for var in self.vars.values():
            var.set(True)

    def _deselect_all(self):
        for var in self.vars.values():
            var.set(False)

    def _on_ok(self):
        self.result = [col for col, var in self.vars.items() if var.get()]
        if not self.result:
            messagebox.showwarning("No Columns Selected", "You must select at least one column to import.", parent=self)
            return
        self.destroy()

# ==============================================================================
# ----------------------------- ExcelUpdaterApp Class --------------------------
# ==============================================================================

class ExcelUpdaterApp:
    def __init__(self, root):
        """Initialize the GUI application."""
        self.root = root
        self.root.title("Excel Updater from SQLite DB (Index Sync)")
        self.root.geometry("650x700")  # Increased height to ensure status bar is always visible 

        # --- Settings File ---
        self.settings_file = SETTINGS_FILE
        
        # --- Data for column selection ---
        self.all_db_columns = []
        self.selected_db_columns = [] # This is the single source of truth for column selection!
        
        # --- Table Selection Variables ---
        self.available_tables = []
        self.selected_table_name = tk.StringVar(value=DEFAULT_TABLE_NAME) # Will hold the table name

        self.db_path = tk.StringVar()
        self.excel_path = tk.StringVar()
        
        # --- Dynamic keyword management lists ---
        self.keyword_widgets = [] 
        self.selected_colors_rgb = [
            (255, 204, 204), (255, 255, 153), (204, 255, 204),
            (204, 229, 255), (255, 229, 204), (229, 204, 255),
            (204, 255, 255), (255, 204, 255), (255, 240, 200), (200, 240, 255) 
        ]
        
        # --- Tracking largest successful DB pull size ---
        self.max_db_count = 0

        # Create status bar FIRST (before main_frame) so it's always visible
        self.status_label = tk.Label(root, text="Ready. Please select files and enter keywords.", bd=1, relief=tk.SUNKEN, anchor=tk.W, padx=5, pady=3)
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

        # Now create main frame - it will fill remaining space above status bar
        main_frame = tk.Frame(root, padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(0, weight=1)

        # --- File Selection Widgets (Row 0-1) ---
        tk.Label(main_frame, text="1. Select SQLite Database File:", anchor='w').grid(row=0, column=0, columnspan=3, sticky='ew', pady=(0, 5))
        db_entry = tk.Entry(main_frame, textvariable=self.db_path, state='readonly')
        db_entry.grid(row=1, column=0, sticky='ew', ipady=4)
        self.select_columns_button = tk.Button(main_frame, text="Select Columns...", command=self.open_column_selector, state='disabled')
        self.select_columns_button.grid(row=1, column=1, sticky='ew', padx=(10, 0))
        tk.Button(main_frame, text="Browse...", command=self.select_db_file).grid(row=1, column=2, sticky='ew', padx=(10, 0))
        
        # --- Table Selection Widgets (New Row 2) ---
        table_frame = tk.Frame(main_frame)
        table_frame.grid(row=2, column=0, columnspan=3, sticky='ew', pady=(10, 5))
        tk.Label(table_frame, text="2. Select Database Table:", anchor='w').pack(side=tk.LEFT)
        
        self.table_option_menu = tk.OptionMenu(table_frame, self.selected_table_name, DEFAULT_TABLE_NAME, command=self._on_table_selected)
        self.table_option_menu.config(state='disabled')
        self.table_option_menu.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        
        # --- Excel File Selection (Old Row 2, now Row 3) ---
        tk.Label(main_frame, text="3. Select Target Excel File:", anchor='w').grid(row=3, column=0, columnspan=3, sticky='ew', pady=(10, 5))
        excel_entry = tk.Entry(main_frame, textvariable=self.excel_path, state='readonly')
        excel_entry.grid(row=4, column=0, sticky='ew', ipady=4)
        tk.Button(main_frame, text="Browse...", command=self.select_excel_file).grid(row=4, column=2, sticky='ew', padx=(10, 0))

        # --- Status Label for Index Requirement (Old Row 4, now Row 5) ---
        tk.Label(main_frame, 
                 text=f"The program synchronizes by appending new data and deleting obsolete data using the '{UNIQUE_ID_COLUMN}' column.", 
                 anchor='w', fg='blue').grid(row=5, column=0, columnspan=3, sticky='ew', pady=(10, 10))


        # --- KEYWORDS FRAME (Old Row 5, now Row 6) ---
        keywords_container = tk.LabelFrame(main_frame, text="4. Enter Keywords and Choose Highlight Colors", padx=10, pady=10)
        keywords_container.grid(row=6, column=0, columnspan=3, sticky='ew', pady=(20, 10))
        keywords_container.columnconfigure(0, weight=1)

        # Frame to hold the dynamically added keyword rows
        self.keyword_rows_frame = tk.Frame(keywords_container)
        self.keyword_rows_frame.pack(fill='x', expand=True)
        self.keyword_rows_frame.columnconfigure(0, weight=1)
        
        # Add Button to add more rows
        add_button_frame = tk.Frame(keywords_container)
        add_button_frame.pack(fill='x', pady=(10, 0))
        tk.Button(add_button_frame, text="+ Add Rule", command=lambda: self.add_keyword_row(len(self.keyword_widgets))).pack(side=tk.LEFT)
        # --- END KEYWORDS FRAME ---
        
        # --- General Settings/Save Frame (New Row 7) ---
        settings_control_frame = tk.Frame(main_frame)
        settings_control_frame.grid(row=8, column=0, columnspan=3, sticky='ew', pady=(5, 5))

        tk.Button(settings_control_frame, 
                  text="💾 Manually Save Settings", 
                  command=self.manual_save_settings_click,
                  bg="#1E90FF", fg="white", 
                  font=('Helvetica', 10, 'bold')).pack(side=tk.RIGHT, padx=5)


        self.update_button = tk.Button(main_frame, text="Synchronize Excel Sheet", command=self.start_update_thread, bg="#4CAF50", fg="white", font=('Helvetica', 10, 'bold'))
        self.update_button.grid(row=9, column=0, columnspan=3, pady=(15, 10), ipady=8, sticky='ew') # Changed to row 9

        # Status bar was already created earlier (before main_frame)
        # This ensures it's always visible at the bottom

        self.load_settings()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    # ==========================================================================
    # --- New Manual Save Method ---
    # ==========================================================================
    
    def manual_save_settings_click(self):
        """Saves current settings and provides user feedback."""
        try:
            self.save_settings()
            self.status_label.config(text="Settings manually saved.")
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save settings: {e}")

    # ==========================================================================
    # --- Other Methods ---
    # ==========================================================================
    
    def _fetch_table_names(self):
        """Connects to the DB and returns a list of all non-sqlite tables."""
        db_path = self.db_path.get()
        if not db_path:
            return []
        
        conn = None
        try:
            # Added timeout here for good measure, though not strictly required for SELECT on master table
            conn = sqlite3.connect(db_path, timeout=10.0) 
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")
            tables = sorted([row[0] for row in cursor.fetchall()])
            return tables
        except Exception as e:
            print(f"Error fetching tables: {e}")
            return []
        finally:
            if conn:
                conn.close()

    def _update_table_option_menu(self):
        """Updates the OptionMenu with the newly fetched table names."""
        self.available_tables = self._fetch_table_names()
        
        menu = self.table_option_menu["menu"]
        menu.delete(0, "end") # Clear existing options

        if not self.available_tables:
            self.table_option_menu.config(state='disabled')
            self.selected_table_name.set("No Tables Found")
            return
        
        self.table_option_menu.config(state='normal')

        # Add all fetched tables to the menu
        for table in self.available_tables:
            menu.add_command(label=table, command=tk._setit(self.selected_table_name, table, self._on_table_selected))

        # Restore previously selected table or set to the first one found
        current_selection = self.selected_table_name.get()
        if current_selection not in self.available_tables:
            if self.available_tables:
                # Use the default table name if it exists, otherwise use the first table
                if DEFAULT_TABLE_NAME in self.available_tables:
                    self.selected_table_name.set(DEFAULT_TABLE_NAME)
                else:
                    self.selected_table_name.set(self.available_tables[0])
            else:
                self.selected_table_name.set("Select Table")
        
        self._load_and_set_db_columns() # Load columns for the new selection

    def _on_table_selected(self, table_name):
        """Callback when a table is selected from the OptionMenu."""
        self.status_label.config(text=f"Table selected: {table_name}. Loading columns...")
        self.save_settings()
        self._load_and_set_db_columns()

    # ==========================================================================
    # --- Keyword and Color Methods ---
    # ==========================================================================

    def add_keyword_row(self, index, initial_keyword="", initial_rgb=None):
        """Dynamically creates a new keyword entry row."""
        
        # FIX for ZeroDivisionError: Use the fixed default colors list for initial indexing
        default_fixed_colors = [
            (255, 204, 204), (255, 255, 153), (204, 255, 204),
            (204, 229, 255), (255, 229, 204), (229, 204, 255),
            (204, 255, 255), (255, 204, 255), (255, 240, 200), (200, 240, 255) 
        ]
        
        if initial_rgb is None:
            # Use the fixed list length to calculate the default color index
            default_color_index = index % len(default_fixed_colors)
            initial_rgb = default_fixed_colors[default_color_index]
            
        # Append the chosen or default color to the instance's tracking list.
        if index < len(self.selected_colors_rgb):
            self.selected_colors_rgb[index] = initial_rgb
        else:
            self.selected_colors_rgb.append(initial_rgb)


        row_frame = tk.Frame(self.keyword_rows_frame)
        row_frame.grid(row=index, column=0, sticky='ew', pady=4)
        row_frame.columnconfigure(0, weight=1)

        entry = tk.Entry(row_frame)
        entry.insert(0, initial_keyword)
        entry.grid(row=0, column=0, sticky='ew', padx=(0, 10))
        
        color_label = tk.Label(row_frame, width=4, relief='sunken', bg=self.get_hex_from_rgb(initial_rgb))
        color_label.grid(row=0, column=1, padx=(0, 10))
        
        color_button = tk.Button(row_frame, text="Choose Color...", command=lambda idx=index: self.choose_color(idx))
        color_button.grid(row=0, column=2)

        self.keyword_widgets.append({
            'row_frame': row_frame,
            'entry': entry, 
            'label': color_label, 
            'index': index 
        })

    def on_closing(self):
        self.save_settings()
        self.root.destroy()

    def save_settings(self):
        current_keywords = [widget['entry'].get() for widget in self.keyword_widgets]
        colors_to_save = self.selected_colors_rgb[:len(current_keywords)] 
        
        settings = {
            "db_path": self.db_path.get(),
            "excel_path": self.excel_path.get(),
            "selected_table_name": self.selected_table_name.get(), # Save the selected table
            "keywords": current_keywords,
            "colors_rgb": colors_to_save,
            "selected_columns": self.selected_db_columns # This is where the column list is saved
        }
        try:
            with open(self.settings_file, 'w') as f:
                json.dump(settings, f, indent=4)
        except Exception as e:
            print(f"Error saving settings: {e}")

    def load_settings(self):
        
        default_keywords = ["CRITICAL, FATAL", "ERROR", "WARNING, Failure, Low"]
        
        # Reset dynamic components
        for widget in self.keyword_widgets:
            widget['row_frame'].destroy()
        self.keyword_widgets = []
        self.selected_colors_rgb = []
        
        
        if not os.path.exists(self.settings_file):
            # Create default keyword rows if no settings file exists
            for i, keyword in enumerate(default_keywords):
                self.add_keyword_row(i, initial_keyword=keyword)
            # Important: Table name will be the DEFAULT_TABLE_NAME constant
            return

        try:
            with open(self.settings_file, 'r') as f:
                settings = json.load(f)
                
            self.db_path.set(settings.get("db_path", ""))
            self.excel_path.set(settings.get("excel_path", ""))
            self.selected_table_name.set(settings.get("selected_table_name", DEFAULT_TABLE_NAME)) # Load saved table
            self.selected_db_columns = settings.get("selected_columns", []) 
            
            loaded_keywords = settings.get("keywords", [])
            loaded_colors = settings.get("colors_rgb", [])

            if not loaded_keywords:
                loaded_keywords = default_keywords
                
            for i, keyword in enumerate(loaded_keywords):
                color = tuple(loaded_colors[i]) if i < len(loaded_colors) else None
                self.add_keyword_row(i, initial_keyword=keyword, initial_rgb=color)
            
            # After loading DB path, fetch tables and columns
            if self.db_path.get():
                self._update_table_option_menu()
            
            self.status_label.config(text="Settings loaded successfully.")
        except Exception as e:
            self.status_label.config(text=f"Error loading settings: {e}. Using defaults.")
            # Ensure defaults are created if loading fails
            for i, keyword in enumerate(default_keywords):
                self.add_keyword_row(i, initial_keyword=keyword)
            self.selected_table_name.set(DEFAULT_TABLE_NAME)


    def choose_color(self, index):
        while index >= len(self.selected_colors_rgb):
            self.selected_colors_rgb.append((255, 255, 255)) 

        initial_color_rgb = self.selected_colors_rgb[index]
        initial_color_hex = self.get_hex_from_rgb(initial_color_rgb)
        
        color_data = colorchooser.askcolor(title="Choose highlight color", initialcolor=initial_color_hex)
        
        if color_data and color_data[0]:
            new_rgb = tuple(map(int, color_data[0]))
            self.selected_colors_rgb[index] = new_rgb
            
            widget_match = next((w for w in self.keyword_widgets if w['index'] == index), None)
            if widget_match:
                widget_match['label'].config(bg=color_data[1])
            
            self.save_settings()

    def get_hex_from_rgb(self, rgb_tuple):
        return f'#{int(rgb_tuple[0]):02x}{int(rgb_tuple[1]):02x}{int(rgb_tuple[2]):02x}'

    def select_db_file(self):
        path = filedialog.askopenfilename(title="Select a Database File", filetypes=(("SQLite files", "*.db *.sqlite *.sqlite3"), ("All files", "*.*")))
        if path:
            self.db_path.set(path)
            self.status_label.config(text=f"Database selected: {path}. Fetching tables...")
            self._update_table_option_menu() # Fetch tables upon DB selection
            self.save_settings()

    def _load_and_set_db_columns(self):
        """Reads column headers from the selected DB and enables the column selector."""
        db_path = self.db_path.get()
        table_name = self.selected_table_name.get()

        if not db_path or table_name in ("No Tables Found", "Select Table", DEFAULT_TABLE_NAME) and DEFAULT_TABLE_NAME not in self.available_tables:
            self.all_db_columns = []
            self.selected_db_columns = []
            self.select_columns_button.config(state='disabled')
            return

        conn = None
        try:
            conn = sqlite3.connect(db_path, timeout=10.0)
            cursor = conn.cursor()
            
            cursor.execute(f'PRAGMA table_info("{table_name}")') 
            
            self.all_db_columns = [info[1] for info in cursor.fetchall()]
            conn.close()

            # --- FIX: Direct persistence logic using the loaded data ---
            
            # Check if the existing selected_db_columns (loaded from settings) 
            # are still valid for the current table.
            
            # **The fix is primarily in this block:**
            if self.selected_db_columns:
                # Filter the loaded selection against the newly fetched all_db_columns
                valid_saved = [col for col in self.selected_db_columns if col in self.all_db_columns]
                
                # Check if we have at least one valid column from the saved setting.
                if valid_saved:
                    # Crucially, ONLY use the saved columns. We don't need a percentage check 
                    # as the user wants their specific selection.
                    self.selected_db_columns = valid_saved
                else:
                    # Saved data had no columns in common with the new table, default to all.
                    self.selected_db_columns = self.all_db_columns[:]
            else:
                # No selection loaded from settings, default to all.
                self.selected_db_columns = self.all_db_columns[:]

            # End Fix
            
            self.select_columns_button.config(state='normal')
            self.status_label.config(text=f"Table '{table_name}' loaded. {len(self.selected_db_columns)} of {len(self.all_db_columns)} columns selected.")
        except Exception as e:
            messagebox.showerror("Database Error", f"Could not read table info from '{table_name}':\n{e}")
            self.all_db_columns = []
            self.selected_db_columns = []
            self.select_columns_button.config(state='disabled')
        finally:
            if conn:
                conn.close()

    def open_column_selector(self):
        """Opens the Toplevel window for column selection."""
        dialog = ColumnSelector(self.root, self.all_db_columns, self.selected_db_columns)
        self.root.wait_window(dialog)
        if dialog.result is not None:
            self.selected_db_columns = dialog.result
            self.status_label.config(text=f"{len(self.selected_db_columns)} of {len(self.all_db_columns)} columns selected. Saving selection...")
            self.save_settings()

    def select_excel_file(self):
        path = filedialog.askopenfilename(title="Select an Excel File", filetypes=(("Excel files", "*.xlsx *.xls *.xlsm *.xlsb"), ("All files", "*.*")))
        if path:
            self.excel_path.set(path)
            self.status_label.config(text=f"Excel file selected: {path}")
            

    def start_update_thread(self):
        if not self.db_path.get() or not self.excel_path.get():
            messagebox.showerror("Error", "Please select both a database and an Excel file.")
            return
        if not self.selected_db_columns:
            messagebox.showerror("Error", "No columns are selected to import. Please use the 'Select Columns...' button.")
            return
        if UNIQUE_ID_COLUMN not in self.selected_db_columns:
            messagebox.showerror("Error", f"Synchronization requires the '{UNIQUE_ID_COLUMN}' column to be selected.")
            return

        self.update_button.config(state=tk.DISABLED, text="Synchronizing...")
        thread = threading.Thread(target=self.update_sheet)
        thread.daemon = True
        thread.start()

    # --- ID Standardization Helper Function (KEPT) ---
    def standardize_id(self, id_raw):
        """
        Converts ID from database or Excel into a clean string format (e.g., '631').
        Handles float string formats like '631.0' that Excel can produce.
        """
        if id_raw is None:
            return None
        
        id_str = str(id_raw).strip()
        
        # Check if the string can be cleanly converted to an integer
        try:
            # If it's a float-string (e.g., '3242.0'), convert it to int first to drop the .0
            if '.' in id_str and id_str.endswith('.0'):
                return str(int(float(id_str)))
            # If it's just a number string, try to clean it up
            return str(int(id_str))
        except ValueError:
            # If it's not a standard number (e.g., a text-based index), return the original string
            return id_str
        
    def fetch_data_once(self, query):
        """Executes the query once and returns the data."""
        conn = None
        try:
            # --- FIX: Added timeout=10.0 (seconds) ---
            conn = sqlite3.connect(self.db_path.get(), timeout=10.0) 
            cursor = conn.cursor()
            cursor.execute(query)
            data = cursor.fetchall()
            return data
        # IMPORTANT: Catching sqlite3.OperationalError for a clean exit on lock/timeout
        except sqlite3.OperationalError as e:
            # --- Added diagnostic print for clear feedback when a lock occurs ---
            print(f"Database access failed (Lock/Timeout/Operational Error): {e}")
            return []
        finally:
            if conn:
                conn.close()

    def get_unique_id_count(self, data, unique_id_col_index):
        """Helper to get the count of unique, standardized IDs."""
        if not data:
            return 0
        db_ids = set(self.standardize_id(row[unique_id_col_index]) for row in data)
        db_ids.discard(None)
        return len(db_ids)

    def fetch_with_retry(self, query, unique_id_col_index):
        """
        Fetches data from DB with retry logic based on count consistency and maximum count seen.
        """
        
        largest_data_so_far = []
        largest_count_so_far = self.max_db_count
        
        for attempt in range(MAX_FETCH_RETRIES):
            self.root.after(0, lambda: self.status_label.config(text=f"DB Fetch Attempt {attempt + 1}: Retrieving data..."))
            
            # --- Fetch 1 ---
            data1 = self.fetch_data_once(query)
            count1 = self.get_unique_id_count(data1, unique_id_col_index)
            
            # --- CRITICAL FIX: If count is low/zero, wait longer and retry immediately ---
            if count1 < 100 and count1 < largest_count_so_far:
                print(f"DB Fetch Warning: Initial pull failed (Count: {count1}). Waiting 1.0s and retrying...")
                time.sleep(1.0)
                data1 = self.fetch_data_once(query)
                count1 = self.get_unique_id_count(data1, unique_id_col_index)
                
            # --- Fetch 2 (Validation) ---
            time.sleep(0.2) # Increased pause to 200ms
            data2 = self.fetch_data_once(query)
            count2 = self.get_unique_id_count(data2, unique_id_col_index)
            
            # Update largest count seen
            current_max = max(count1, count2)
            if current_max > largest_count_so_far:
                largest_count_so_far = current_max
                largest_data_so_far = data1 if count1 >= count2 else data2

            
            if count1 == count2:
                # Stable pull achieved.
                if count1 >= self.max_db_count:
                    # Stable and >= historical max. SUCCESS condition.
                    print(f"DB Fetch Success: Stable count {count1} confirmed (Attempt {attempt + 1}).")
                    self.max_db_count = count1 # Update the historical max count
                    return data1 # Return the stable data

                elif count1 < self.max_db_count:
                    # Stable but smaller than the historical max (e.g., stable 5874 < max 6126).
                    # This means data was deleted OR the DB is currently only allowing partial pulls.
                    print(f"DB Fetch Warning: Stable count {count1} is lower than historical Max Seen ({self.max_db_count}).")

                    if attempt == MAX_FETCH_RETRIES - 1:
                        # On final attempt, accept the largest dataset found so far.
                        print(f"DB Fetch Returning: Max retries reached. Returning largest dataset seen ({largest_count_so_far}).")
                        return largest_data_so_far
                    
                    # Otherwise, continue retrying, hoping the max count will appear.
            
            else:
                print(f"DB Fetch Warning: Counts mismatched (Attempt {attempt + 1}). Count 1: {count1}, Count 2: {count2}. Largest seen: {largest_count_so_far}. Retrying...")
                
                if attempt == MAX_FETCH_RETRIES - 1:
                    # Last attempt: return the largest dataset retrieved.
                    print(f"DB Fetch Failure: Could not achieve stability. Returning largest dataset fetched: {largest_count_so_far}.")
                    return largest_data_so_far if largest_data_so_far else data1

                # --- MODIFIED LINE: Wait longer between retries (increased from 0.5 to 1.5) ---
                time.sleep(1.5) # Wait before next retry set

        return [] # Should only be reached if MAX_FETCH_RETRIES=0
    
    def _apply_highlights_to_range(self, sheet, data_range, event_col_index, highlight_rules):
        """Applies keyword-based highlighting to a given xlwings Range."""
        if not highlight_rules or event_col_index == -1:
            return

        # Read all event cells once to minimize calls to Excel
        try:
            event_cells = data_range.columns[event_col_index].value
            # Ensure it's a list even if it's a single cell value
            if not isinstance(event_cells, list):
                event_cells = [event_cells] if event_cells is not None else []
        except Exception as e:
            print(f"Error reading event column for highlighting: {e}")
            return

        for i, cell_content_raw in enumerate(event_cells):
            cell_content = str(cell_content_raw).strip().lower()
            
            # Apply highlighting if any phrase matches
            for rule in highlight_rules:
                if any(phrase in cell_content for phrase in rule['phrases']):
                    row_to_format = data_range.rows[i]
                    row_to_format.color = rule['color']
                    
                    # Apply borders (same as your existing logic)
                    for border_id in [7, 8, 9, 10, 11]:
                        border = row_to_format.api.Borders(border_id)
                        border.LineStyle = 1
                        border.Weight = 2
                        border.ColorIndex = 15
                    break

    def update_sheet(self):
        app, wb = None, None
        
        try:
            # 0. --- Dynamic keyword reading ---
            highlight_rules = []
            for i, widget in enumerate(self.keyword_widgets):
                keyword_text = widget['entry'].get().strip().lower()
                color = self.selected_colors_rgb[i] if i < len(self.selected_colors_rgb) else (255, 255, 255)
                if keyword_text:
                    phrases = [p.strip() for p in keyword_text.split(',') if p.strip()]
                    if phrases:
                        highlight_rules.append({'phrases': phrases, 'color': color})
            # --- End Dynamic keyword reading ---

            # 1. Setup Query
            table_name = self.selected_table_name.get() # Get the currently selected table
            self.status_label.config(text=f"Setting up DB query for table '{table_name}'...")
            
            # This line correctly quotes column names
            safe_columns = [f'"{col}"' for col in self.selected_db_columns]
            
            # --- FIX APPLIED HERE: Quote the table name to handle hyphens/special characters ---
            query = f"SELECT {', '.join(safe_columns)} FROM \"{table_name}\"" 
            
            db_headers = self.selected_db_columns
            lower_db_headers = [str(h).lower() for h in db_headers]
            
            if UNIQUE_ID_COLUMN.lower() not in lower_db_headers:
                raise ValueError(f"Required column '{UNIQUE_ID_COLUMN}' not selected in the current table.")

            unique_id_col_index = lower_db_headers.index(UNIQUE_ID_COLUMN.lower())
            event_col_index = lower_db_headers.index(EVENT_COLUMN_NAME.lower()) if EVENT_COLUMN_NAME.lower() in lower_db_headers else -1

            # 2. Fetch ALL Data from DB (Using Retry Logic)
            db_data = self.fetch_with_retry(query, unique_id_col_index)
            raw_db_count = len(db_data)
            
            if not db_data and raw_db_count == 0:
                messagebox.showerror("DB Error", f"Failed to retrieve any data from table '{table_name}' after multiple attempts. Cannot synchronize.")
                return

            # 3. Excel Setup
            self.status_label.config(text="Connecting to Excel...")
            wb = xw.Book(self.excel_path.get())
            app = wb.app
            sheet, app.screen_updating = wb.sheets[0], False
            
            # 4. Read Existing Data and Perform Synchronization
            self.status_label.config(text="Reading existing data in Excel for synchronization...")

            try:
                last_row = sheet.range('A' + str(sheet.api.Rows.Count)).end('up').row
            except Exception:
                last_row = 1
            
            # Write headers if they don't exist
            if last_row == 1 or sheet.range('A1').value is None or sheet.range('A1').value.lower() != db_headers[0].lower():
                sheet.range('A1').value = db_headers
                sheet.range('A1').expand('right').font.bold = True
                last_row = 1 
            
            existing_rows_map = {} # {standardized_id: row_number}
            
            if last_row > 1:
                unique_id_column_letter = chr(ord('A') + unique_id_col_index)
                
                unique_id_col_range = sheet.range(f'{unique_id_column_letter}1:{unique_id_column_letter}{last_row}')
                unique_id_col_range.number_format = '@' 
                
                id_range = sheet.range(f'{unique_id_column_letter}2:{unique_id_column_letter}{last_row}')
                existing_ids_raw = id_range.value
                
                if not isinstance(existing_ids_raw, list):
                     existing_ids_raw = [existing_ids_raw] if existing_ids_raw is not None else []

                for i, row_id_raw in enumerate(existing_ids_raw):
                    if row_id_raw is None: continue 
                    
                    standardized_id = self.standardize_id(row_id_raw)
                    if standardized_id is None: continue 
                    
                    existing_rows_map[standardized_id] = i + 2 

            # Identify rows to ADD to Excel and rows to DELETE from Excel
            db_ids = set(self.standardize_id(row[unique_id_col_index]) for row in db_data)
            db_ids.discard(None) 
            
            excel_ids = set(existing_rows_map.keys())

            ids_to_add = db_ids - excel_ids
            ids_to_delete = excel_ids - db_ids
            
            # --- DIAGNOSTIC PRINTING ---
            print("\n--- Synchronization Diagnostics (Standardized) ---")
            print(f"Table: {table_name}")
            print(f"Raw DB Records Fetched: {raw_db_count}")
            print(f"Unique/Valid DB Records: {len(db_ids)}")
            print(f"Total Excel Records: {len(excel_ids)}")
            print(f"IDs to Add: {len(ids_to_add)} (Example: {list(ids_to_add)[:5]})")
            print(f"IDs to Delete: {len(ids_to_delete)} (Example: {list(ids_to_delete)[:5]})")
            print("--------------------------------------------------")
            # --- END DIAGNOSTIC PRINTING ---

            # Find the actual rows to delete in Excel
            rows_to_delete_nums = sorted([existing_rows_map[id_val] for id_val in ids_to_delete], reverse=True)
            
            # Find the new data rows to add (must use the original DB data list)
            data_to_add = []
            for row in db_data:
                db_id = self.standardize_id(row[unique_id_col_index])
                if db_id in ids_to_add:
                    data_to_add.append(row)

            # 5. Apply Synchronization Changes
            
            # A) Delete rows from Excel
            if rows_to_delete_nums:
                self.status_label.config(text=f"Deleting {len(rows_to_delete_nums)} obsolete rows from Excel...")
                for row_num in rows_to_delete_nums:
                    sheet.range(f'A{row_num}').api.EntireRow.Delete()
                last_row = sheet.range('A' + str(sheet.api.Rows.Count)).end('up').row
                
            # B) Append new rows to Excel
            start_row_new_data = last_row + 1 if last_row > 1 or sheet.range('A1').value else 2
            
            if data_to_add:
                
                self.status_label.config(text=f"Appending {len(data_to_add)} new rows to Excel, starting at row {start_row_new_data}...")
                
                # Write Data
                data_range_new = sheet.range(f'A{start_row_new_data}').resize(len(data_to_add), len(db_headers))
                data_range_new.number_format = '@'
                data_range_new.value = data_to_add
            
            # 6. Apply Highlighting (New and Existing Data)
            
            final_last_row = sheet.range('A' + str(sheet.api.Rows.Count)).end('up').row
            
            if final_last_row > 1 and highlight_rules:
                self.status_label.config(text="Applying/Re-applying highlights to all existing rows...")
                
                # Highlight the entire data region (excluding header row 1)
                data_range_all = sheet.range(f'A2:A{final_last_row}').expand('right')
                
                # Apply the rules to the entire dataset
                self._apply_highlights_to_range(sheet, data_range_all, event_col_index, highlight_rules)

            
            if not data_to_add and not rows_to_delete_nums and not highlight_rules:
                 self.status_label.config(text="Synchronization complete. No differences found.")
                 messagebox.showinfo("Success", "Synchronization complete. No new or missing records found.")
                 return

            sheet.autofit()
            self.status_label.config(text="Success! Excel sheet synchronized. Please save the file.")
            messagebox.showinfo("Success", "The Excel sheet has been synchronized successfully.\n\nPlease save the file in Excel to keep the changes.")
            
        except Exception as e:
            self.status_label.config(text=f"An error occurred: {e}")
            messagebox.showerror("Error", f"An unexpected error occurred during synchronization:\n\n{e}")
        finally:
            if app:
                try: app.screen_updating = True
                except Exception as e: print(f"Could not re-enable screen updating: {e}")
                
            self.root.after(0, lambda: self.update_button.config(state=tk.NORMAL, text="Synchronize Excel Sheet"))

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelUpdaterApp(root)
    root.mainloop()