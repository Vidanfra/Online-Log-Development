import tkinter as tk
from tkinter import ttk  # For the Treeview widget
import customtkinter as ctk
import sqlite3
import pandas as pd
import os
from tkinter import filedialog, messagebox

# --- Main Application Class ---
class SQLiteViewerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- App State ---
        self.current_df = None # Holds the currently viewed data in memory

        # --- Window Setup ---
        self.title("SQLite Table Viewer")
        self.geometry("1000x650")
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        # --- App State Variables ---
        self.source_db_path_var = ctk.StringVar()
        self.table_var = ctk.StringVar(value="Select a table")
        self.status_var = ctk.StringVar(value="Ready")

        # --- Main Layout ---
        self.grid_columnconfigure(0, weight=1, minsize=380)
        self.grid_columnconfigure(1, weight=3)
        self.grid_rowconfigure(0, weight=1)

        self._create_widgets()

    def _create_widgets(self):
        """Creates and places all the UI widgets."""
        # --- Left Control Panel ---
        left_panel = ctk.CTkFrame(self, width=380)
        left_panel.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        left_panel.grid_propagate(False)
        
        # 1. Source Frame
        source_frame = ctk.CTkFrame(left_panel)
        source_frame.pack(pady=10, padx=10, fill="x", expand=True)
        ctk.CTkLabel(source_frame, text="1. Source Database", font=ctk.CTkFont(weight="bold")).pack(pady=(5,10), anchor="w")
        
        ctk.CTkLabel(source_frame, text="Source SQLite File:").pack(anchor="w")
        db_entry_frame = ctk.CTkFrame(source_frame, fg_color="transparent")
        db_entry_frame.pack(fill="x", expand=True)
        ctk.CTkEntry(db_entry_frame, textvariable=self.source_db_path_var).pack(side="left", fill="x", expand=True)
        ctk.CTkButton(db_entry_frame, text="...", width=30, command=self._select_source_db).pack(side="left", padx=(5,0))
        
        ctk.CTkButton(source_frame, text="Fetch Tables", command=self._fetch_tables).pack(pady=10, fill="x")

        ctk.CTkLabel(source_frame, text="Select Table:").pack(anchor="w", pady=(10,0))
        self.table_combobox = ctk.CTkComboBox(source_frame, variable=self.table_var, state="readonly")
        self.table_combobox.pack(fill="x", expand=True, pady=(0, 20))
        self.table_combobox.set("Select a DB and click Fetch...")

        # 2. Action Frame
        action_frame = ctk.CTkFrame(left_panel)
        action_frame.pack(pady=10, padx=10, fill="x", side="bottom")
        ctk.CTkButton(action_frame, text="Load Table to Viewer", height=40, font=ctk.CTkFont(size=14, weight="bold"), command=self._load_table_to_viewer).pack(fill="x", ipady=5)

        # --- Right Panel (Data Viewer) ---
        right_panel = ctk.CTkFrame(self)
        right_panel.grid(row=0, column=1, padx=(0,10), pady=10, sticky="nsew")
        right_panel.grid_rowconfigure(1, weight=1)
        right_panel.grid_columnconfigure(0, weight=1)

        viewer_controls_frame = ctk.CTkFrame(right_panel)
        viewer_controls_frame.grid(row=0, column=0, pady=(0,10), padx=10, sticky="ew")
        ctk.CTkLabel(viewer_controls_frame, text="SQL Data Viewer", font=ctk.CTkFont(weight="bold")).pack(side="left")
        ctk.CTkButton(viewer_controls_frame, text="Refresh View from Source", command=self._load_table_to_viewer).pack(side="right")

        tree_frame = ctk.CTkFrame(right_panel)
        tree_frame.grid(row=1, column=0, padx=10, sticky="nsew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        # Create Treeview with scrollbars
        self.tree = ttk.Treeview(tree_frame, show="headings")
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        # --- Status Bar ---
        self.status_bar = ctk.CTkLabel(self, textvariable=self.status_var, anchor="w", font=ctk.CTkFont(size=12))
        self.status_bar.grid(row=1, column=0, columnspan=2, padx=10, pady=(5,5), sticky="ew")

    def _update_status(self, text, color="white"):
        self.status_var.set(text)
        self.status_bar.configure(text_color=color)
        self.update_idletasks()

    # --- UI Event Handlers ---
    def _select_source_db(self):
        path = filedialog.askopenfilename(title="Select Source SQLite DB", filetypes=(("DB files", "*.db *.sqlite *.sqlite3"), ("All files", "*.*")))
        if path: 
            self.source_db_path_var.set(path)
            self.table_var.set("Click Fetch Tables...")
            self._clear_treeview() # Clear viewer if a new DB is selected

    def _fetch_tables(self):
        db_path = self.source_db_path_var.get()
        if not db_path: 
            messagebox.showerror("Error", "Please select a source SQLite database first.")
            return
        try:
            self._update_status("Fetching tables...")
            with sqlite3.connect(db_path) as conn:
                query = "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%';"
                tables = [tbl[0] for tbl in conn.execute(query).fetchall()]
            
            if tables:
                self.table_combobox.configure(values=tables)
                self.table_var.set(tables[0])
                self._update_status("Tables fetched successfully.")
            else:
                self.table_combobox.configure(values=[])
                self.table_var.set("No tables found")
                messagebox.showinfo("Info", "No user tables found in the database.")
        except Exception as e: 
            messagebox.showerror("Database Error", f"An error occurred: {e}")
            self._update_status("Error fetching tables.", "red")

    # --- Data Viewer Logic ---
    def _load_table_to_viewer(self):
        db_path = self.source_db_path_var.get()
        table_name = self.table_var.get()

        if not all([db_path, table_name]) or "..." in table_name or "No tables" in table_name:
            messagebox.showerror("Input Error", "Please select a valid database and table first.")
            return

        self._update_status(f"Loading data from '{table_name}'...")
        try:
            with sqlite3.connect(db_path) as conn:
                self.current_df = pd.read_sql_query(f'SELECT * FROM "{table_name}"', conn)
            
            if self.current_df is None:
                raise ValueError("Failed to load data into DataFrame.")

            self._populate_treeview(self.current_df)
            self._update_status(f"Successfully loaded {len(self.current_df)} rows from '{table_name}'.", "green")

        except Exception as e:
            messagebox.showerror("Viewer Error", f"Could not load data from the database:\n{e}")
            self._update_status("Failed to load data.", "red")
            
    def _clear_treeview(self):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = []

    def _populate_treeview(self, df):
        self._clear_treeview()
        cols = list(df.columns)
        self.tree["columns"] = cols
        for col in cols:
            self.tree.heading(col, text=col, anchor='w')
            # Basic auto-sizing
            col_width = max(df[col].astype(str).map(len).max(), len(col)) * 7 + 10
            self.tree.column(col, width=min(max(col_width, 80), 400), anchor='w')
        
        for index, row in df.iterrows():
            self.tree.insert("", "end", values=list(row))
        
        if len(df) == 0:
            self._update_status("Table loaded, but it is empty.", "yellow")


if __name__ == "__main__":
    app = SQLiteViewerApp()
    app.mainloop()