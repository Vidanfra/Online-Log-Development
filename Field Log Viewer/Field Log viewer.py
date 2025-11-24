
import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser
import sqlite3
import xlwings as xw
import threading
import json
import os
import shutil
import tempfile

# --- Constants ---
# The name of the column IN THE EXCEL FILE that will be checked for keywords.
# The search for this column name is case-insensitive.
EVENT_COLUMN_NAME = 'Event'
# The name of the table to query in the database.
TABLE_NAME = 'fieldlog'
# The name of the primary key or rowid column in the database.
INDEX_COLUMN_NAME = 'index'


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
        main_frame = tk.Frame(self, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        controls_frame = tk.Frame(main_frame)
        controls_frame.pack(fill=tk.X, side=tk.TOP, pady=(0, 10))
        tk.Button(controls_frame, text="Select All", command=self._select_all).pack(side=tk.LEFT)
        tk.Button(controls_frame, text="Deselect All", command=self._deselect_all).pack(side=tk.LEFT, padx=10)
        bottom_frame = tk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(10, 0))
        tk.Button(bottom_frame, text="OK", command=self._on_ok, width=10).pack(side=tk.RIGHT)
        tk.Button(bottom_frame, text="Cancel", command=self.destroy, width=10).pack(side=tk.RIGHT, padx=10)
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
        for var in self.vars.values(): var.set(True)

    def _deselect_all(self):
        for var in self.vars.values(): var.set(False)

    def _on_ok(self):
        self.result = [col for col, var in self.vars.items() if var.get()]
        if not self.result:
            messagebox.showwarning("No Columns Selected", "You must select at least one column to import.", parent=self)
            return
        self.destroy()

class ExcelUpdaterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Updater from SQLite DB")
        self.root.geometry("650x700")
        self.settings_file = "settings.json"
        self.all_db_columns = []
        self.selected_db_columns = []
        self.saved_selected_columns = []
        self.db_path = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.keyword_widgets = [] 
        self.selected_colors_rgb = []
        self.default_palette = [(255, 204, 204), (255, 255, 153), (204, 255, 204), (204, 229, 255), (255, 229, 204), (229, 204, 255), (204, 255, 255), (255, 204, 255), (255, 240, 200), (200, 240, 255)]
        main_frame = tk.Frame(root, padx=15, pady=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(0, weight=1)
        tk.Label(main_frame, text="1. Select SQLite Database File:", anchor='w').grid(row=0, column=0, columnspan=3, sticky='ew', pady=(0, 5))
        db_entry = tk.Entry(main_frame, textvariable=self.db_path, state='readonly')
        db_entry.grid(row=1, column=0, sticky='ew', ipady=4)
        self.select_columns_button = tk.Button(main_frame, text="Select Columns...", command=self.open_column_selector, state='disabled')
        self.select_columns_button.grid(row=1, column=1, sticky='ew', padx=(10, 0))
        tk.Button(main_frame, text="Browse...", command=self.select_db_file).grid(row=1, column=2, sticky='ew', padx=(10, 0))
        tk.Label(main_frame, text="2. Select Target Excel File:", anchor='w').grid(row=2, column=0, columnspan=3, sticky='ew', pady=(10, 5))
        excel_entry = tk.Entry(main_frame, textvariable=self.excel_path, state='readonly')
        excel_entry.grid(row=3, column=0, sticky='ew', ipady=4)
        tk.Button(main_frame, text="Browse...", command=self.select_excel_file).grid(row=3, column=2, sticky='ew', padx=(10, 0))
        keywords_container = tk.LabelFrame(main_frame, text="3. Enter Keywords and Choose Highlight Colors", padx=10, pady=10)
        keywords_container.grid(row=4, column=0, columnspan=3, sticky='ew', pady=(20, 10))
        keywords_container.columnconfigure(0, weight=1)
        self.keyword_rows_frame = tk.Frame(keywords_container)
        self.keyword_rows_frame.pack(fill='x', expand=True)
        self.keyword_rows_frame.columnconfigure(0, weight=1)
        add_button_frame = tk.Frame(keywords_container)
        add_button_frame.pack(fill='x', pady=(10, 0))
        tk.Button(add_button_frame, text="+ Add Rule", command=lambda: self.add_keyword_row(len(self.keyword_widgets))).pack(side=tk.LEFT)
        self.update_button = tk.Button(main_frame, text="Update Excel Sheet", command=self.start_update_thread, bg="#4CAF50", fg="white", font=('Helvetica', 10, 'bold'))
        self.update_button.grid(row=5, column=0, columnspan=3, pady=(20, 10), ipady=8, sticky='ew')
        self.status_label = tk.Label(root, text="Ready. Please select files and enter keywords.", bd=1, relief=tk.SUNKEN, anchor=tk.W, padx=5)
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)
        self.load_settings()
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def add_keyword_row(self, index, initial_keyword="", initial_rgb=None):
        if initial_rgb is None:
            initial_rgb = self.default_palette[index % len(self.default_palette)]
        while len(self.selected_colors_rgb) <= index:
            self.selected_colors_rgb.append(initial_rgb)
        self.selected_colors_rgb[index] = initial_rgb
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
        self.keyword_widgets.append({'row_frame': row_frame, 'entry': entry, 'label': color_label, 'index': index})

    def on_closing(self):
        self.save_settings()
        self.root.destroy()

    def save_settings(self):
        settings = {
            "db_path": self.db_path.get(),
            "excel_path": self.excel_path.get(),
            "keywords": [w['entry'].get() for w in self.keyword_widgets],
            "colors": self.selected_colors_rgb[:len(self.keyword_widgets)],
            "selected_columns": self.selected_db_columns
        }
        try:
            with open(self.settings_file, 'w') as f: json.dump(settings, f, indent=4)
        except Exception as e: print(f"Error saving settings: {e}")

    def load_settings(self):
        if not os.path.exists(self.settings_file): self.save_settings()
        else:
            with open(self.settings_file, "r") as f:
                try: settings = json.load(f)
                except json.JSONDecodeError: settings = {}
            self.db_path.set(settings.get("db_path", ""))
            self.excel_path.set(settings.get("excel_path", ""))
            self.saved_selected_columns = settings.get("selected_columns", [])
            self.selected_colors_rgb = settings.get("colors", [])
            for i, keyword in enumerate(settings.get("keywords", [])):
                self.add_keyword_row(i, initial_keyword=keyword, initial_rgb=tuple(self.selected_colors_rgb[i]) if i < len(self.selected_colors_rgb) else None)
            if self.db_path.get(): self._load_and_set_db_columns()
            self.status_label.config(text="Settings loaded.")

    def choose_color(self, index):
        while index >= len(self.selected_colors_rgb): self.selected_colors_rgb.append((255, 255, 255))
        color_data = colorchooser.askcolor(title="Choose highlight color", initialcolor=self.get_hex_from_rgb(self.selected_colors_rgb[index]))
        if color_data and color_data[0]:
            self.selected_colors_rgb[index] = tuple(map(int, color_data[0]))
            widget_match = next((w for w in self.keyword_widgets if w['index'] == index), None)
            if widget_match: widget_match['label'].config(bg=color_data[1])
            self.save_settings()

    def get_hex_from_rgb(self, rgb_tuple):
        return f'#{int(rgb_tuple[0]):02x}{int(rgb_tuple[1]):02x}{int(rgb_tuple[2]):02x}'

    def select_db_file(self):
        path = filedialog.askopenfilename(title="Select Database", filetypes=(("SQLite files", "*.db *.sqlite *.sqlite3"), ("All files", "*.*")))
        if path:
            self.db_path.set(path)
            self._load_and_set_db_columns()
            self.save_settings()

    def _load_and_set_db_columns(self):
        if not self.db_path.get(): return
        temp_db_path = None
        try:
            safe_table_name = f'"{TABLE_NAME}"'
            temp_db_path = tempfile.NamedTemporaryFile(delete=False, suffix=".db").name
            shutil.copy2(self.db_path.get(), temp_db_path)
            conn = sqlite3.connect(temp_db_path)
            cursor = conn.cursor()
            cursor.execute(f"PRAGMA table_info({safe_table_name})")
            self.all_db_columns = [info[1] for info in cursor.fetchall()]
            conn.close()
            if not self.all_db_columns:
                self.select_columns_button.config(state='disabled')
                messagebox.showwarning("DB Warning", f"No columns in table '{TABLE_NAME}'.")
                return
            valid_saved = [c for c in self.saved_selected_columns if c in self.all_db_columns]
            self.selected_db_columns = valid_saved if valid_saved else self.all_db_columns[:]
            self.select_columns_button.config(state='normal')
            self.status_label.config(text=f"DB loaded. {len(self.selected_db_columns)}/{len(self.all_db_columns)} columns selected.")
        except Exception as e:
            messagebox.showerror("DB Error", f"Could not read table '{TABLE_NAME}':\n{e}")
            self.all_db_columns, self.selected_db_columns = [], []
            self.select_columns_button.config(state='disabled')
        finally:
            if temp_db_path and os.path.exists(temp_db_path):
                try: os.remove(temp_db_path)
                except: pass

    def open_column_selector(self):
        dialog = ColumnSelector(self.root, self.all_db_columns, self.selected_db_columns)
        self.root.wait_window(dialog)
        if dialog.result is not None:
            self.selected_db_columns = dialog.result
            self.status_label.config(text=f"{len(self.selected_db_columns)}/{len(self.all_db_columns)} columns selected.")
            self.save_settings()

    def select_excel_file(self):
        path = filedialog.askopenfilename(title="Select Excel File", filetypes=(("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*")))
        if path:
            self.excel_path.set(path)
            self.save_settings()

    def start_update_thread(self):
        if not all([self.db_path.get(), self.excel_path.get()]):
            messagebox.showerror("Error", "Database and Excel file must be selected.")
            return
        if not self.selected_db_columns:
            messagebox.showerror("Error", "No columns selected for import.")
            return
        if INDEX_COLUMN_NAME not in self.selected_db_columns:
            messagebox.showerror("Error", f"Index column '{INDEX_COLUMN_NAME}' must be selected.")
            return
        self.update_button.config(state=tk.DISABLED, text="Updating...")
        threading.Thread(target=self.update_sheet, daemon=True).start()

    def update_sheet(self):
        app, wb = None, None
        try:
            self.status_label.config(text="Connecting to Excel...")
            wb = xw.Book(self.excel_path.get())
            app = wb.app
            sheet, app.screen_updating = wb.sheets[0], False
            
            excel_headers = sheet.range('A1').expand('right').value
            excel_headers = [excel_headers] if not isinstance(excel_headers, list) else (excel_headers or [])
            
            # Default to Full Refresh (safe default)
            perform_full_refresh = True
            last_excel_index = 0
            
            safe_table = f'"{TABLE_NAME}"'
            safe_index = f'"{INDEX_COLUMN_NAME}"'
            
            # --- 1. Determine if we can do an Incremental Update ---
            if INDEX_COLUMN_NAME in excel_headers:
                # Find the column letter for the Index
                col_letter = chr(ord('A') + excel_headers.index(INDEX_COLUMN_NAME))
                
                # Find the last row in Excel
                last_row_obj = sheet.range(f'{col_letter}{sheet.cells.last_cell.row}').end('up')
                last_row = last_row_obj.row
                
                if last_row > 1: 
                    val = last_row_obj.value
                    try: 
                        last_excel_index = int(val)
                    except (ValueError, TypeError): 
                        last_excel_index = 0

                # Check Database Max Index
                conn = sqlite3.connect(self.db_path.get())
                max_db_index = conn.execute(f"SELECT MAX({safe_index}) FROM {safe_table}").fetchone()[0] or 0
                conn.close()

                # LOGIC FIX:
                # If Excel is up to date (indices match), stop. 
                # If Excel has data (index > 0) and is behind DB, append (Incremental).
                if max_db_index == last_excel_index:
                    self.status_label.config(text="Database and Excel are in sync.")
                    messagebox.showinfo("Info", "No new data to update.")
                    return # STOP HERE - Don't reload everything
                elif last_excel_index > 0 and last_excel_index < max_db_index:
                    perform_full_refresh = False
                # Else (Excel is empty, or DB was reset and is smaller than Excel): Full Refresh

            # --- 2. Fetch Data ---
            conn = sqlite3.connect(self.db_path.get())
            safe_cols = [f'"{c}"' for c in self.selected_db_columns]
            
            if perform_full_refresh:
                self.status_label.config(text="Performing FULL refresh...")
                query = f"SELECT {', '.join(safe_cols)} FROM {safe_table}"
                sheet.range('A1').expand('table').clear()
                sheet.range('A1').value = self.selected_db_columns
                sheet.range('A1').expand('right').font.bold = True
                start_cell = 'A2'
            else:
                # Incremental: Only get rows GREATER THAN the last one in Excel
                count_new = max_db_index - last_excel_index
                self.status_label.config(text=f"Appending {count_new} new rows...")
                query = f"SELECT {', '.join(safe_cols)} FROM {safe_table} WHERE {safe_index} > {last_excel_index} ORDER BY {safe_index} ASC"
                start_cell = f'A{sheet.range("A" + str(sheet.cells.last_cell.row)).end("up").row + 1}'

            data = conn.execute(query).fetchall()
            conn.close()

            if not data:
                messagebox.showinfo("Info", "No new data found.")
                return

            # --- 3. Write Data ---
            self.status_label.config(text=f"Writing {len(data)} rows...")
            # Defines a range covering ONLY the NEW data
            data_range = sheet.range(start_cell).resize(len(data), len(self.selected_db_columns))
            data_range.number_format = '@' 
            data_range.value = data
            
            # --- 4. Highlight Logic (Optimized) ---
            rules = []
            for i, w in enumerate(self.keyword_widgets):
                text_val = w['entry'].get().strip()
                if text_val:
                    phrases = [p.strip().lower() for p in text_val.split(',') if p.strip()]
                    rules.append({'phrases': phrases, 'color': self.selected_colors_rgb[i]})

            # Find column index safely
            event_col_idx = -1
            lower_target = EVENT_COLUMN_NAME.lower()
            for i, col_name in enumerate(self.selected_db_columns):
                if col_name.lower() == lower_target:
                    event_col_idx = i
                    break
            
            if rules and event_col_idx != -1:
                self.status_label.config(text="Calculating highlights...")
                
                # Group rows by color to minimize Excel calls
                # Dict format: { (255,0,0): [0, 1, 5], ... }
                color_groups = {} 

                for i, row in enumerate(data):
                    cell_content = str(row[event_col_idx]).lower()
                    for rule in rules:
                        if any(phrase in cell_content for phrase in rule['phrases']):
                            c = rule['color']
                            if c not in color_groups: color_groups[c] = []
                            color_groups[c].append(i)
                            break 
                
                self.status_label.config(text="Applying highlights...")
                
                # Apply formatting in batches to avoid string limits and improve speed
                BATCH_SIZE = 30 
                
                for color, row_indices in color_groups.items():
                    # Process this color in chunks of 30 rows
                    for i in range(0, len(row_indices), BATCH_SIZE):
                        batch = row_indices[i : i + BATCH_SIZE]
                        
                        # Build a comma-separated string of addresses relative to data_range
                        # data_range.rows[idx] gets the specific NEW row
                        addresses = [data_range.rows[r_idx].address for r_idx in batch]
                        address_str = ",".join(addresses)
                        
                        try:
                            # Select the combined range for this batch
                            union_range = sheet.range(address_str)
                            
                            # 1. Set Color
                            union_range.color = color
                            
                            # 2. Set Borders (Optimized: Set all borders at once for the selection)
                            # 7-12 covers all edge and inside borders
                            for border_id in range(7, 13):
                                border = union_range.api.Borders(border_id)
                                border.LineStyle = 1
                                border.Weight = 2
                                border.ColorIndex = 15
                        except Exception as e:
                            print(f"Batch error: {e}")
                            # Fallback: apply one by one if batch fails
                            for r_idx in batch:
                                try:
                                    rng = data_range.rows[r_idx]
                                    rng.color = color
                                except: pass

            sheet.autofit()
            self.status_label.config(text="Success! Sheet updated.")
            self.root.after(0, lambda: messagebox.showinfo("Success", f"Added {len(data)} new rows."))

        except Exception as e:
            error_info = f"{e}"
            self.status_label.config(text=f"Error: {error_info.split(':', 1)[0]}")
            self.root.after(0, lambda e=e: messagebox.showerror("Error", f"An unexpected error occurred:\n\n{e}"))
        finally:
            if app:
                try: app.screen_updating = True
                except: pass
            self.root.after(0, lambda: self.update_button.config(state=tk.NORMAL, text="Update Excel Sheet"))

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelUpdaterApp(root)
    root.mainloop()