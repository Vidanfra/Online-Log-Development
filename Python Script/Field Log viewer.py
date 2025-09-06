import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser
import sqlite3
import xlwings as xw
import threading
import json
import os

# --- Constants ---
# The name of the column IN THE EXCEL FILE that will be checked for keywords.
# The search for this column name is case-insensitive.
EVENT_COLUMN_NAME = 'Event'
# The name of the table to query in the database.
TABLE_NAME = 'fieldlog'


class ColumnSelector(tk.Toplevel):
    """A modal dialog window for selecting columns from a list."""
    def __init__(self, parent, all_columns, selected_columns):
        super().__init__(parent)
        self.title("Select Columns to Import")
        self.geometry("400x500")

        # Make the window modal
        self.transient(parent)
        self.grab_set()

        self.result = None
        self.vars = {col: tk.BooleanVar(value=(col in selected_columns)) for col in all_columns}

        # --- Main frame ---
        main_frame = tk.Frame(self, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- Controls on top ---
        controls_frame = tk.Frame(main_frame)
        controls_frame.pack(fill=tk.X, side=tk.TOP, pady=(0, 10))
        tk.Button(controls_frame, text="Select All", command=self._select_all).pack(side=tk.LEFT)
        tk.Button(controls_frame, text="Deselect All", command=self._deselect_all).pack(side=tk.LEFT, padx=10)

        # --- Bottom buttons (Packed before the expanding scroll area) ---
        bottom_frame = tk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(10, 0))
        tk.Button(bottom_frame, text="OK", command=self._on_ok, width=10).pack(side=tk.RIGHT)
        tk.Button(bottom_frame, text="Cancel", command=self.destroy, width=10).pack(side=tk.RIGHT, padx=10)

        # --- Scrollable Checkbox Area (Fills the remaining middle space) ---
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


class ExcelUpdaterApp:
    def __init__(self, root):
        """Initialize the GUI application."""
        self.root = root
        self.root.title("Excel Updater from SQLite DB")
        self.root.geometry("650x650")

        # --- Settings File ---
        self.settings_file = "settings.json"
        
        # --- Data for column selection ---
        self.all_db_columns = []
        self.selected_db_columns = []
        self.saved_selected_columns = []

        self.db_path = tk.StringVar()
        self.excel_path = tk.StringVar()

        self.keyword_entries = []
        self.color_labels = []
        self.selected_colors_rgb = [
            (255, 204, 204), (255, 255, 153), (204, 255, 204),
            (204, 229, 255), (255, 229, 204), (229, 204, 255)
        ]

        main_frame = tk.Frame(root, padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(0, weight=1)

        # --- File Selection Widgets (reconfigured grid) ---
        tk.Label(main_frame, text="1. Select SQLite Database File:", anchor='w').grid(row=0, column=0, columnspan=3, sticky='ew', pady=(0, 5))
        db_entry = tk.Entry(main_frame, textvariable=self.db_path, state='readonly')
        db_entry.grid(row=1, column=0, sticky='ew', ipady=4)
        self.select_columns_button = tk.Button(main_frame, text="Select Columns...", command=self.open_column_selector, state='disabled')
        self.select_columns_button.grid(row=1, column=1, sticky='ew', padx=(10, 0))
        tk.Button(main_frame, text="Browse...", command=self.select_db_file).grid(row=1, column=2, sticky='ew', padx=(10, 0))
        
        tk.Label(main_frame, text="2. Select Target Excel File:", anchor='w').grid(row=2, column=0, columnspan=3, sticky='ew', pady=(10, 5))
        excel_entry = tk.Entry(main_frame, textvariable=self.excel_path, state='readonly')
        excel_entry.grid(row=3, column=0, sticky='ew', ipady=4, columnspan=3) # Span 3 to align with widgets above
        tk.Button(main_frame, text="Browse...", command=self.select_excel_file).grid(row=3, column=2, sticky='ew', padx=(10, 0)) # Overlaps for alignment

        keywords_frame = tk.LabelFrame(main_frame, text="3. Enter Keywords and Choose Highlight Colors", padx=10, pady=10)
        keywords_frame.grid(row=4, column=0, columnspan=3, sticky='ew', pady=(20, 10))
        keywords_frame.columnconfigure(0, weight=1)

        for i in range(6):
            entry = tk.Entry(keywords_frame)
            entry.grid(row=i, column=0, sticky='ew', pady=4, padx=(0, 10))
            self.keyword_entries.append(entry)
            color_label = tk.Label(keywords_frame, width=4, relief='sunken', bg=self.get_hex_from_rgb(self.selected_colors_rgb[i]))
            color_label.grid(row=i, column=1, padx=(0, 10))
            self.color_labels.append(color_label)
            color_button = tk.Button(keywords_frame, text="Choose Color...", command=lambda index=i: self.choose_color(index))
            color_button.grid(row=i, column=2)

        self.update_button = tk.Button(main_frame, text="Update Excel Sheet", command=self.start_update_thread, bg="#4CAF50", fg="white", font=('Helvetica', 10, 'bold'))
        self.update_button.grid(row=5, column=0, columnspan=3, pady=(20, 10), ipady=8, sticky='ew')

        self.status_label = tk.Label(root, text="Ready. Please select files and enter keywords.", bd=1, relief=tk.SUNKEN, anchor=tk.W, padx=5)
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

        self.load_settings()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        self.save_settings()
        self.root.destroy()

    def save_settings(self):
        settings = {
            "db_path": self.db_path.get(),
            "excel_path": self.excel_path.get(),
            "keywords": [entry.get() for entry in self.keyword_entries],
            "colors_rgb": self.selected_colors_rgb,
            "selected_columns": self.selected_db_columns
        }
        try:
            with open(self.settings_file, 'w') as f:
                json.dump(settings, f, indent=4)
        except Exception as e:
            print(f"Error saving settings: {e}")

    def load_settings(self):
        if not os.path.exists(self.settings_file):
            self.keyword_entries[0].insert(0, "CRITICAL, FATAL")
            self.keyword_entries[1].insert(0, "ERROR")
            self.keyword_entries[2].insert(0, "WARNING, Failure, Low")
            return

        try:
            with open(self.settings_file, 'r') as f:
                settings = json.load(f)
            self.db_path.set(settings.get("db_path", ""))
            self.excel_path.set(settings.get("excel_path", ""))
            self.saved_selected_columns = settings.get("selected_columns", [])
            
            # If a DB path was loaded, try to load its columns
            if self.db_path.get():
                self._load_and_set_db_columns()

            loaded_keywords = settings.get("keywords", [])
            for i, keyword in enumerate(loaded_keywords):
                if i < len(self.keyword_entries):
                    self.keyword_entries[i].delete(0, tk.END)
                    self.keyword_entries[i].insert(0, keyword)
            
            loaded_colors = settings.get("colors_rgb", [])
            if loaded_colors and len(loaded_colors) == len(self.color_labels):
                self.selected_colors_rgb = [tuple(c) for c in loaded_colors]
            for i, color_rgb in enumerate(self.selected_colors_rgb):
                if i < len(self.color_labels):
                    self.color_labels[i].config(bg=self.get_hex_from_rgb(color_rgb))
            self.status_label.config(text="Settings loaded successfully.")
        except Exception as e:
            self.status_label.config(text=f"Error loading settings: {e}. Using defaults.")

    def choose_color(self, index):
        color_data = colorchooser.askcolor(title="Choose highlight color", initialcolor=self.get_hex_from_rgb(self.selected_colors_rgb[index]))
        if color_data and color_data[0]:
            self.selected_colors_rgb[index] = tuple(map(int, color_data[0]))
            self.color_labels[index].config(bg=color_data[1])

    def get_hex_from_rgb(self, rgb_tuple):
        return f'#{int(rgb_tuple[0]):02x}{int(rgb_tuple[1]):02x}{int(rgb_tuple[2]):02x}'

    def select_db_file(self):
        path = filedialog.askopenfilename(title="Select a Database File", filetypes=(("SQLite files", "*.db *.sqlite *.sqlite3"), ("All files", "*.*")))
        if path:
            self.db_path.set(path)
            self.status_label.config(text=f"Database selected: {path}")
            self._load_and_set_db_columns()

    def _load_and_set_db_columns(self):
        """Reads column headers from the selected DB and enables the column selector."""
        try:
            conn = sqlite3.connect(self.db_path.get())
            cursor = conn.cursor()
            cursor.execute(f"PRAGMA table_info({TABLE_NAME})")
            self.all_db_columns = [info[1] for info in cursor.fetchall()]
            conn.close()

            if self.saved_selected_columns:
                valid_saved = [col for col in self.saved_selected_columns if col in self.all_db_columns]
                self.selected_db_columns = valid_saved if valid_saved else self.all_db_columns[:]
            else:
                self.selected_db_columns = self.all_db_columns[:]
            
            self.select_columns_button.config(state='normal')
            self.status_label.config(text=f"Database loaded. {len(self.selected_db_columns)} of {len(self.all_db_columns)} columns selected.")
        except Exception as e:
            messagebox.showerror("Database Error", f"Could not read table info from '{TABLE_NAME}':\n{e}")
            self.all_db_columns = []
            self.selected_db_columns = []
            self.select_columns_button.config(state='disabled')

    def open_column_selector(self):
        """Opens the Toplevel window for column selection."""
        dialog = ColumnSelector(self.root, self.all_db_columns, self.selected_db_columns)
        self.root.wait_window(dialog)
        if dialog.result is not None:
            self.selected_db_columns = dialog.result
            self.status_label.config(text=f"{len(self.selected_db_columns)} of {len(self.all_db_columns)} columns selected.")

    def select_excel_file(self):
        path = filedialog.askopenfilename(title="Select an Excel File", filetypes=(("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*")))
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

        self.update_button.config(state=tk.DISABLED, text="Updating...")
        thread = threading.Thread(target=self.update_sheet)
        thread.daemon = True
        thread.start()

    def update_sheet(self):
        app, wb, we_started_app = None, None, False
        try:
            self.status_label.config(text="Fetching data from database...")
            safe_columns = [f'"{col}"' for col in self.selected_db_columns]
            query = f"SELECT {', '.join(safe_columns)} FROM {TABLE_NAME}"
            conn = sqlite3.connect(self.db_path.get())
            cursor = conn.cursor()
            cursor.execute(query)
            data = cursor.fetchall()
            conn.close()
            db_headers = self.selected_db_columns

            if not data:
                messagebox.showinfo("Info", "The database table is empty. Nothing to write.")
                self.root.after(0, lambda: self.update_button.config(state=tk.NORMAL, text="Update Excel Sheet"))
                return

            highlight_rules = []
            for i in range(6):
                keyword_text = self.keyword_entries[i].get().strip().lower()
                if keyword_text:
                    phrases = [p.strip() for p in keyword_text.split(',') if p.strip()]
                    if phrases:
                        highlight_rules.append({'phrases': phrases, 'color': self.selected_colors_rgb[i]})

            self.status_label.config(text="Connecting to Excel...")
            try:
                wb = xw.Book(self.excel_path.get())
                app = wb.app
            except Exception:
                app = xw.App(visible=True)
                we_started_app = True
                wb = app.books.open(self.excel_path.get())
            sheet, app.screen_updating = wb.sheets[0], False

            self.status_label.config(text="Analyzing sheet layout...")

            event_col_index = -1
            lower_db_headers = [str(h).lower() for h in db_headers]
            if EVENT_COLUMN_NAME.lower() in lower_db_headers:
                event_col_index = lower_db_headers.index(EVENT_COLUMN_NAME.lower())
            else:
                messagebox.showwarning("Highlighting Skipped", f"The '{EVENT_COLUMN_NAME}' column was not selected for import. Highlighting is disabled.")
                highlight_rules = []

            self.status_label.config(text="Clearing old sheet content...")
            sheet.used_range.clear() 

            self.status_label.config(text="Writing new headers...")
            sheet.range('A1').value = db_headers
            sheet.range('A1').expand('right').font.bold = True

            self.status_label.config(text="Writing new data to sheet...")
            data_range = sheet.range('A2').resize(len(data), len(db_headers))
            data_range.value = data

            if highlight_rules and event_col_index != -1:
                self.status_label.config(text="Applying highlights...")
                for i, row_data in enumerate(data):
                    cell_content = str(row_data[event_col_index]).lower()
                    for rule in highlight_rules:
                        if any(phrase in cell_content for phrase in rule['phrases']):
                            row_to_format = data_range.rows[i]
                            # Apply background color
                            row_to_format.color = rule['color']
                            
                            # --- FIX: RE-APPLY BORDERS TO SIMULATE GRIDLINES ---
                            # Border constants: 7-11 correspond to left, top, bottom, right, and inside vertical
                            for border_id in [7, 8, 9, 10, 11]:
                                border = row_to_format.api.Borders(border_id)
                                border.LineStyle = 1  # xlContinuous
                                border.Weight = 2     # xlThin
                                border.ColorIndex = 15 # Grey-25%, a good approximation for gridlines
                            # --- END FIX ---
                            break

            sheet.autofit()
            self.status_label.config(text="Success! Your open Excel sheet has been updated. Please save the file.")
            messagebox.showinfo("Success", "The Excel sheet has been updated successfully.\n\nPlease save the file in Excel to keep the changes.")
        except Exception as e:
            self.status_label.config(text=f"An error occurred: {e}")
            messagebox.showerror("Error", f"An unexpected error occurred:\n\n{e}")
        finally:
            if app:
                try: app.screen_updating = True
                except Exception as e: print(f"Could not re-enable screen updating: {e}")
            if we_started_app and app:
                try: app.quit()
                except Exception as e: print(f"Error quitting Excel instance: {e}")
            self.root.after(0, lambda: self.update_button.config(state=tk.NORMAL, text="Update Excel Sheet"))

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelUpdaterApp(root)
    root.mainloop()

