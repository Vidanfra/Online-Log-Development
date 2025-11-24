import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser, ttk
import sqlite3
import xlwings as xw
import threading
import json
import os
import time

# --- Constants ---
EVENT_COLUMN_NAME = 'Event'
DEFAULT_TABLE_NAME = 'DailyLog-Horizon_v14'
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SETTINGS_FILE = os.path.join(SCRIPT_DIR, 'flv_settings.json')
DEFAULT_COLUMNS = ["DateTime", "Runline", "KP", "KP ref", "Event", "Latitude", "Longitude"]

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

        # Main Layout
        main_frame = tk.Frame(self, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Controls
        controls_frame = tk.Frame(main_frame)
        controls_frame.pack(fill=tk.X, side=tk.TOP, pady=(0, 10))
        tk.Button(controls_frame, text="Select All", command=self._select_all).pack(side=tk.LEFT)
        tk.Button(controls_frame, text="Deselect All", command=self._deselect_all).pack(side=tk.LEFT, padx=10)

        # Bottom Buttons
        bottom_frame = tk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(10, 0))
        tk.Button(bottom_frame, text="OK", command=self._on_ok, width=10).pack(side=tk.RIGHT)
        tk.Button(bottom_frame, text="Cancel", command=self.destroy, width=10).pack(side=tk.RIGHT, padx=10)

        # Scrollable Area
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
        
    def _select_all(self):
        for var in self.vars.values(): var.set(True)

    def _deselect_all(self):
        for var in self.vars.values(): var.set(False)

    def _on_ok(self):
        self.result = [col for col, var in self.vars.items() if var.get()]
        if not self.result:
            messagebox.showwarning("No Columns Selected", "You must select at least one column.")
            return
        self.destroy()

# ==============================================================================
# ----------------------------- LogViewerApp Class -----------------------------
# ==============================================================================

class LogViewerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Field Log Viewer (Native)")
        self.root.geometry("1100x800") 

        # --- APPLY GRID STYLING ---
        style = ttk.Style()
        style.theme_use("clam") 
        
        style.configure("Treeview", 
                        background="white",
                        foreground="black",
                        rowheight=25,
                        fieldbackground="white")
        
        style.configure("Treeview.Heading", 
                        font=('Arial', 9, 'bold'),
                        background="#d9d9d9",
                        foreground="black")
        
        style.map('Treeview', background=[('selected', '#0078D7')])

        # --- State Variables ---
        self.settings_file = SETTINGS_FILE
        self.all_db_columns = []
        self.selected_db_columns = [] 
        self.available_tables = []
        self.selected_table_name = tk.StringVar(value=DEFAULT_TABLE_NAME)
        self.db_path = tk.StringVar()
        self.keyword_widgets = [] 
        self.selected_colors_rgb = []
        self.temp_cell_value = "" 
        
        # --- Data Cache for Filtering ---
        self.full_dataset = []      # Stores all rows from DB
        self.current_display_cols = []
        self.current_event_idx = -1
        self.active_highlight_rules = []

        # --- Top Control Panel ---
        control_frame = tk.Frame(root, padx=10, pady=10)
        control_frame.pack(side=tk.TOP, fill=tk.X)

        # Row 0: DB Selection
        tk.Label(control_frame, text="Database:").grid(row=0, column=0, sticky='w')
        tk.Entry(control_frame, textvariable=self.db_path, width=60).grid(row=0, column=1, padx=5, sticky='w')
        tk.Button(control_frame, text="Browse...", command=self.select_db_file).grid(row=0, column=2, padx=5, sticky='w')
        
        # Row 1: Table & Columns
        tk.Label(control_frame, text="Table:").grid(row=1, column=0, sticky='w', pady=5)
        self.table_option_menu = tk.OptionMenu(control_frame, self.selected_table_name, DEFAULT_TABLE_NAME, command=self._on_table_selected)
        self.table_option_menu.grid(row=1, column=1, sticky='ew', padx=5)
        self.select_columns_button = tk.Button(control_frame, text="Select Columns...", command=self.open_column_selector, state='disabled')
        self.select_columns_button.grid(row=1, column=2, padx=5, sticky='w')

        # Row 2: Keywords Frame
        kw_frame = tk.LabelFrame(control_frame, text="Highlight Rules (Comma separated keywords)")
        kw_frame.grid(row=2, column=0, columnspan=3, sticky='ew', pady=10)
        
        self.keyword_scroll_frame = tk.Frame(kw_frame)
        self.keyword_scroll_frame.pack(fill='x', expand=True, padx=5, pady=5)
        
        tk.Button(kw_frame, text="+ Add Rule", command=lambda: self.add_keyword_row(len(self.keyword_widgets))).pack(anchor='w', padx=5, pady=(0,5))

        # Row 3: Action Buttons AND Filter
        btn_frame = tk.Frame(control_frame)
        btn_frame.grid(row=3, column=0, columnspan=3, pady=10)
        
        self.load_btn = tk.Button(btn_frame, text="🔄 Load Data to Viewer", command=self.start_load_thread, bg="#4CAF50", fg="white", font=('Arial', 10, 'bold'), padx=15, pady=5)
        self.load_btn.pack(side=tk.LEFT, padx=(0, 20))

        tk.Button(btn_frame, text="📤 Export to Excel", command=self.export_to_excel, bg="#2196F3", fg="white", font=('Arial', 10), padx=15, pady=5).pack(side=tk.LEFT, padx=(0, 40))

        # --- FILTER UI ---
        tk.Label(btn_frame, text="🔍 Filter Events:", font=('Arial', 10, 'bold')).pack(side=tk.LEFT)
        self.filter_var = tk.StringVar()
        # Bind the key release event to filter instantly as you type
        self.filter_entry = tk.Entry(btn_frame, textvariable=self.filter_var, width=30)
        self.filter_entry.pack(side=tk.LEFT, padx=5)
        self.filter_entry.bind("<KeyRelease>", self.apply_filter)
        
        # --- Bottom: Data Viewer (Treeview) ---
        tree_frame = tk.Frame(root)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Scrollbars
        tree_scroll_y = tk.Scrollbar(tree_frame)
        tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        tree_scroll_x = tk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        self.tree = ttk.Treeview(tree_frame, yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        tree_scroll_y.config(command=self.tree.yview)
        tree_scroll_x.config(command=self.tree.xview)

        # --- COPY FUNCTIONALITY ---
        self.tree.bind("<Control-c>", lambda e: self.copy_selection_to_clipboard())
        
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Copy Cell Value", command=self.copy_clicked_cell)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Copy Selected Row(s)", command=self.copy_selection_to_clipboard)
        
        self.tree.bind("<Button-3>", self.show_context_menu) # Windows/Linux
        self.tree.bind("<Button-2>", self.show_context_menu) # MacOS

        # Status Bar
        self.status_label = tk.Label(root, text="Ready.", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

        self.load_settings()

    # ==========================================================================
    # --- Copy/Paste Logic ---
    # ==========================================================================
    def show_context_menu(self, event):
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        
        if row_id:
            if row_id not in self.tree.selection():
                self.tree.selection_set(row_id)
            
            try:
                col_index = int(col_id.replace('#', '')) - 1
                values = self.tree.item(row_id)['values']
                if 0 <= col_index < len(values):
                    self.temp_cell_value = str(values[col_index])
                else:
                    self.temp_cell_value = ""
            except Exception:
                self.temp_cell_value = ""

            state = "normal" if self.temp_cell_value else "disabled"
            self.context_menu.entryconfig("Copy Cell Value", state=state)
            self.context_menu.post(event.x_root, event.y_root)

    def copy_clicked_cell(self):
        if self.temp_cell_value:
            self.root.clipboard_clear()
            self.root.clipboard_append(self.temp_cell_value)
            self.root.update() 

    def copy_selection_to_clipboard(self):
        selected_items = self.tree.selection()
        if not selected_items: return

        text_to_copy = ""
        for item_id in selected_items:
            values = self.tree.item(item_id)['values']
            line = "\t".join(str(v) if v is not None else "" for v in values)
            text_to_copy += line + "\n"

        self.root.clipboard_clear()
        self.root.clipboard_append(text_to_copy)
        self.root.update()
        self.status_label.config(text=f"Copied {len(selected_items)} rows to clipboard.")

    # ==========================================================================
    # --- Settings & UI Logic ---
    # ==========================================================================
    def add_keyword_row(self, index, initial_keyword="", initial_rgb=None):
        default_fixed_colors = [
            (255, 204, 204), (255, 255, 153), (204, 255, 204),
            (204, 229, 255), (255, 229, 204)
        ]
        
        if initial_rgb is None:
            initial_rgb = default_fixed_colors[index % len(default_fixed_colors)]
            
        if index < len(self.selected_colors_rgb):
            self.selected_colors_rgb[index] = initial_rgb
        else:
            self.selected_colors_rgb.append(initial_rgb)

        row_frame = tk.Frame(self.keyword_scroll_frame)
        row_frame.pack(fill='x', pady=2)
        
        entry = tk.Entry(row_frame)
        entry.insert(0, initial_keyword)
        entry.pack(side=tk.LEFT, fill='x', expand=True, padx=5)
        
        color_label = tk.Label(row_frame, width=4, relief='sunken', bg=self.get_hex_from_rgb(initial_rgb))
        color_label.pack(side=tk.LEFT, padx=5)
        
        tk.Button(row_frame, text="Color...", command=lambda idx=index: self.choose_color(idx)).pack(side=tk.LEFT)
        
        self.keyword_widgets.append({'row_frame': row_frame, 'entry': entry, 'label': color_label, 'index': index})

    def choose_color(self, index):
        color_data = colorchooser.askcolor(initialcolor=self.get_hex_from_rgb(self.selected_colors_rgb[index]))
        if color_data and color_data[0]:
            self.selected_colors_rgb[index] = tuple(map(int, color_data[0]))
            self.keyword_widgets[index]['label'].config(bg=color_data[1])

    def get_hex_from_rgb(self, rgb): return f'#{int(rgb[0]):02x}{int(rgb[1]):02x}{int(rgb[2]):02x}'
    
    def save_settings(self):
        settings = {
            "db_path": self.db_path.get(),
            "selected_table_name": self.selected_table_name.get(),
            "keywords": [w['entry'].get() for w in self.keyword_widgets],
            "colors_rgb": self.selected_colors_rgb[:len(self.keyword_widgets)],
            "selected_columns": self.selected_db_columns
        }
        try:
            with open(self.settings_file, 'w') as f: json.dump(settings, f, indent=4)
        except: pass

    def load_settings(self):
        for w in self.keyword_widgets: w['row_frame'].destroy()
        self.keyword_widgets = []
        self.selected_colors_rgb = []

        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, 'r') as f:
                    data = json.load(f)
                    self.db_path.set(data.get("db_path", ""))
                    self.selected_table_name.set(data.get("selected_table_name", DEFAULT_TABLE_NAME))
                    self.selected_db_columns = data.get("selected_columns", [])
                    kw = data.get("keywords", ["Error", "Warning"])
                    cl = data.get("colors_rgb", [])
                    for i, k in enumerate(kw):
                        c = tuple(cl[i]) if i < len(cl) else None
                        self.add_keyword_row(i, k, c)
                if self.db_path.get(): self._update_table_option_menu()
            except:
                 self.add_keyword_row(0, "Error")
        else:
            self.add_keyword_row(0, "Error")

    # ==========================================================================
    # --- DB & Column Logic ---
    # ==========================================================================
    def select_db_file(self):
        path = filedialog.askopenfilename(filetypes=(("SQLite DB", "*.db *.sqlite *.sqlite3"), ("All Files", "*.*")))
        if path:
            self.db_path.set(path)
            self._update_table_option_menu()
            self.save_settings()

    def _fetch_table_names(self):
        if not self.db_path.get(): return []
        try:
            conn = sqlite3.connect(self.db_path.get())
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")
            res = sorted([row[0] for row in cursor.fetchall()])
            conn.close()
            return res
        except: return []

    def _update_table_option_menu(self):
        self.available_tables = self._fetch_table_names()
        menu = self.table_option_menu["menu"]
        menu.delete(0, "end")
        
        if not self.available_tables:
            self.table_option_menu.config(state='disabled')
            self.selected_table_name.set("No Tables Found")
            return
            
        self.table_option_menu.config(state='normal')
        for t in self.available_tables:
            menu.add_command(label=t, command=tk._setit(self.selected_table_name, t, self._on_table_selected))
            
        if self.selected_table_name.get() not in self.available_tables:
            if DEFAULT_TABLE_NAME in self.available_tables:
                self.selected_table_name.set(DEFAULT_TABLE_NAME)
            else:
                self.selected_table_name.set(self.available_tables[0])
                
        self._load_and_set_db_columns()

    def _on_table_selected(self, val):
        self._load_and_set_db_columns()
        self.save_settings()

    def _load_and_set_db_columns(self):
        try:
            conn = sqlite3.connect(self.db_path.get())
            cursor = conn.cursor()
            cursor.execute(f'PRAGMA table_info("{self.selected_table_name.get()}")')
            self.all_db_columns = [info[1] for info in cursor.fetchall()]
            conn.close()
            
            valid_existing_selection = [c for c in self.selected_db_columns if c in self.all_db_columns]
            
            if valid_existing_selection:
                self.selected_db_columns = valid_existing_selection
            else:
                default_matches = [c for c in DEFAULT_COLUMNS if c in self.all_db_columns]
                if default_matches:
                    self.selected_db_columns = default_matches
                else:
                    self.selected_db_columns = self.all_db_columns[:]
            
            self.select_columns_button.config(state='normal')
        except: pass

    def open_column_selector(self):
        dialog = ColumnSelector(self.root, self.all_db_columns, self.selected_db_columns)
        self.root.wait_window(dialog)
        if dialog.result:
            self.selected_db_columns = dialog.result
            self.save_settings()

    # ==========================================================================
    # --- DATA LOADING & FILTERING ---
    # ==========================================================================
    def start_load_thread(self):
        if not self.db_path.get():
            messagebox.showerror("Error", "Select a database file first.")
            return
        self.save_settings()
        self.load_btn.config(state='disabled', text="Loading...")
        # Clear current filter when reloading data
        self.filter_var.set("") 
        threading.Thread(target=self.load_data_to_treeview, daemon=True).start()

    def load_data_to_treeview(self):
        try:
            cols = self.selected_db_columns
            if not cols: return
            
            conn = sqlite3.connect(self.db_path.get())
            cursor = conn.cursor()
            safe_cols = [f'"{c}"' for c in cols]
            query = f'SELECT {", ".join(safe_cols)} FROM "{self.selected_table_name.get()}"'
            cursor.execute(query)
            
            # --- Store data in Memory for filtering ---
            self.full_dataset = cursor.fetchall()
            self.current_display_cols = cols
            conn.close()

            # Prepare Rules
            self.active_highlight_rules = []
            for i, widget in enumerate(self.keyword_widgets):
                txt = widget['entry'].get().strip().lower()
                if txt:
                    self.active_highlight_rules.append({
                        'phrases': [p.strip() for p in txt.split(',') if p.strip()],
                        'tag': f'rule_{i}',
                        'hex': self.get_hex_from_rgb(self.selected_colors_rgb[i])
                    })
            
            # Find Event Index
            self.current_event_idx = -1
            lower_cols = [c.lower() for c in cols]
            if EVENT_COLUMN_NAME.lower() in lower_cols:
                self.current_event_idx = lower_cols.index(EVENT_COLUMN_NAME.lower())

            # Render Full Dataset initially
            self.root.after(0, lambda: self._render_tree(self.full_dataset))
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            self.root.after(0, lambda: self.load_btn.config(state='normal', text="🔄 Load Data to Viewer"))

    def apply_filter(self, event=None):
        """Filters the stored self.full_dataset based on entry box."""
        if not self.full_dataset: return

        search_term = self.filter_var.get().strip().lower()
        
        if not search_term:
            # If empty, show everything
            self._render_tree(self.full_dataset)
            return

        filtered_rows = []
        
        for row in self.full_dataset:
            # If we know the Event column, search ONLY there
            if self.current_event_idx != -1:
                cell_val = str(row[self.current_event_idx]).lower() if row[self.current_event_idx] else ""
                if search_term in cell_val:
                    filtered_rows.append(row)
            else:
                # Fallback: Search the whole row if "Event" column not selected
                if any(search_term in str(cell).lower() for cell in row if cell is not None):
                    filtered_rows.append(row)

        self._render_tree(filtered_rows)

    def _render_tree(self, rows_to_show):
        self.tree.delete(*self.tree.get_children())
        
        self.tree["columns"] = self.current_display_cols
        self.tree["show"] = "headings"
        
        # Zebra Stripes
        self.tree.tag_configure('oddrow', background='#F2F2F2') 
        self.tree.tag_configure('evenrow', background='white')
        
        for col in self.current_display_cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120)

        # Highlight Rules
        for rule in self.active_highlight_rules:
            self.tree.tag_configure(rule['tag'], background=rule['hex'])

        # Update Status
        if self.filter_var.get():
             self.status_label.config(text=f"Filtered: {len(rows_to_show)} of {len(self.full_dataset)} records.")
        else:
             self.status_label.config(text=f"Showing {len(rows_to_show)} records.")

        # Batch Insert
        for i, row in enumerate(rows_to_show):
            tags_to_apply = []
            
            # Determine Highlight
            highlight_tag = None
            if self.current_event_idx != -1 and row[self.current_event_idx]:
                content = str(row[self.current_event_idx]).lower()
                for rule in self.active_highlight_rules:
                    if any(p in content for p in rule['phrases']):
                        highlight_tag = rule['tag']
                        break
            
            if highlight_tag:
                tags_to_apply = [highlight_tag]
            else:
                tags_to_apply = ['oddrow' if i % 2 else 'evenrow']
            
            display_row = [str(item) if item is not None else "" for item in row]
            self.tree.insert("", "end", values=display_row, tags=tuple(tags_to_apply))

    # ==========================================================================
    # --- Excel Export ---
    # ==========================================================================
    def export_to_excel(self):
        if not self.tree.get_children():
            messagebox.showinfo("Empty", "No data to export. Load data first.")
            return
            
        try:
            self.status_label.config(text="Exporting to Excel (opening new instance)...")
            
            app = xw.App(visible=True)
            wb = app.books.add()
            ws = wb.sheets[0]
            
            headers = self.tree["columns"]
            ws.range('A1').value = headers
            ws.range('A1').expand('right').font.bold = True
            
            data = []
            for child in self.tree.get_children():
                data.append(self.tree.item(child)['values'])
                
            if data:
                ws.range('A2').value = data
                ws.autofit()
                
            self.status_label.config(text="Export Complete.")
            
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export: {e}")
            self.status_label.config(text="Export Failed.")

if __name__ == "__main__":
    root = tk.Tk()
    app = LogViewerApp(root)
    root.mainloop()