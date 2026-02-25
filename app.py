import sys
import traceback
import os
import json
import re
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# --- CONFIGURATION ---
DATABASE_FILE = "shipping_database.json"

TARGET_MAP = {
    'Seal': 6,
    'PKG': 11,
    'Weight': 12,
    'Description': 13,
    'Broker': 20,
    'VGM': 21
}

MASTER_KEYWORDS = {
    'Seal': ['SEAL'],
    'PKG': ['PKG', 'PACKAGE', 'QTY', 'QUANTITY'],
    'Weight': ['WEIGHT', 'KGS', 'G.W', 'GROSS'],
    'Description': ['DESCRIPTION', 'COMMODITY', 'GOODS', 'DESC'],
    'Broker': ['BROKER', 'AGENT'],
    'VGM': ['VGM', 'VERIFIED']
}

class ShippingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Shipping App - Native Excel")
        self.root.geometry("900x700")
        self.db = {} 
        self.load_database()
        self.build_ui()

    def load_database(self):
        if getattr(sys, 'frozen', False):
            app_path = os.path.dirname(sys.executable)
        else:
            app_path = os.path.dirname(os.path.abspath(__file__))
        self.db_path = os.path.join(app_path, DATABASE_FILE)
        
        if os.path.exists(self.db_path):
            try:
                with open(self.db_path, 'r') as f: self.db = json.load(f)
            except: self.db = {}

    def save_database(self):
        try:
            with open(self.db_path, 'w') as f: json.dump(self.db, f, indent=4)
        except: pass

    def normalize(self, val):
        if pd.isna(val) or val == "": return ""
        return re.sub(r'[^A-Z0-9]', '', str(val).upper())

    def build_ui(self):
        p = ttk.Frame(self.root, padding=20)
        p.pack(fill="both", expand=True)

        lf1 = ttk.LabelFrame(p, text="Step 1: Import Source", padding=10)
        lf1.pack(fill="x", pady=10)
        ttk.Button(lf1, text="Import Master Data", command=self.import_data).pack(anchor="w")
        self.lbl_stats = ttk.Label(lf1, text=f"Records: {len(self.db)}")
        self.lbl_stats.pack(anchor="w")

        lf2 = ttk.LabelFrame(p, text="Step 2: Fill Target", padding=10)
        lf2.pack(fill="x", pady=10)
        ttk.Button(lf2, text="Select Target File & Run", command=self.process_native).pack(fill="x", pady=5)
        
        self.log_text = tk.Text(p, height=15)
        self.log_text.pack(fill="both", expand=True, pady=10)
        self.log("Ready.")

    def log(self, msg):
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")

    def import_data(self):
        path = filedialog.askopenfilename()
        if not path: return
        try:
            self.log(f"Reading {os.path.basename(path)}...")
            if path.lower().endswith('.csv'):
                try: df = pd.read_csv(path, header=None, encoding='utf-8')
                except: df = pd.read_csv(path, header=None, encoding='latin1')
            else:
                df = pd.read_excel(path, header=None)

            header_idx = None
            cont_col_idx = None
            for r_idx, row in df.head(30).iterrows():
                row_vals = [self.normalize(x) for x in row.values]
                if any("CONTAINER" in str(x) for x in row_vals):
                    header_idx = r_idx
                    for c_idx, val in enumerate(row_vals):
                        if "CONTAINER" in str(val) or "CNTR" in str(val):
                            cont_col_idx = c_idx
                            break
                    if cont_col_idx is not None: break
            
            if cont_col_idx is None:
                messagebox.showerror("Error", "No 'Container' column found.")
                return

            count = 0
            header_row = df.iloc[header_idx]
            col_name_map = {}
            for c_idx, val in enumerate(header_row):
                col_name_map[c_idx] = self.normalize(val)

            for r_idx, row in df.iloc[header_idx+1:].iterrows():
                k = self.normalize(row.iloc[cont_col_idx])
                if not k: continue
                if k not in self.db: self.db[k] = {}
                for c_idx, val in enumerate(row):
                    if pd.notna(val):
                        header_name = col_name_map.get(c_idx, f"COL_{c_idx}")
                        self.db[k][header_name] = str(val).strip()
                count += 1
                
            self.save_database()
            self.lbl_stats.config(text=f"Records: {len(self.db)}")
            self.log(f"Success! Imported {count} records.")

        except Exception as e:
            self.log(f"Error: {e}")
            messagebox.showerror("Error", str(e))

    def process_native(self):
        target_path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not target_path: return

        # --- LIBRARY CHECK ---
        try:
            import win32com.client as win32
        except ImportError:
            # Show exact path where to install
            py_path = sys.executable
            messagebox.showerror("Library Missing", 
                f"Python cannot find 'pywin32'.\n\nPlease run this command in your terminal:\n\n{py_path} -m pip install pywin32")
            return
        # ---------------------

        excel = None
        wb = None
        try:
            self.log("Launching Excel...")
            # Dispatch Excel
            excel = win32.Dispatch("Excel.Application")
            excel.Visible = True  
            
            self.log(f"Opening {os.path.basename(target_path)}...")
            wb = excel.Workbooks.Open(target_path)
            
            total_updates = 0
            sheet = excel.ActiveSheet # Use active sheet
            
            # Find Used Range
            try:
                used_range = sheet.UsedRange
                max_row = used_range.Rows.Count + used_range.Row - 1
            except:
                max_row = 1000 # Fallback

            self.log(f"Scanning {max_row} rows...")

            # Iterate rows starting at 9
            for r in range(9, max_row + 1):
                # Read Column 3 (C)
                try:
                    c_val = sheet.Cells(r, 3).Value
                except: continue

                if not c_val: continue
                
                k = self.normalize(c_val)
                if k in self.db:
                    db_rec = self.db[k]
                    
                    for field, col_idx in TARGET_MAP.items():
                        val_to_write = None
                        field_kws = MASTER_KEYWORDS[field]
                        
                        for db_key, db_val in db_rec.items():
                            if any(kw in db_key for kw in field_kws):
                                val_to_write = db_val
                                break
                        
                        if val_to_write:
                            sheet.Cells(r, col_idx).Value = val_to_write
                            total_updates += 1

            self.log(f"Finished. Total updates: {total_updates}")
            
            if total_updates > 0:
                new_file = os.path.splitext(target_path)[0] + "_Filled_Native.xlsx"
                # Remove old if exists
                if os.path.exists(new_file):
                    try: os.remove(new_file)
                    except: pass
                    
                wb.SaveAs(new_file)
                self.log(f"Saved to: {new_file}")
                messagebox.showinfo("Success", f"Updated {total_updates} fields!\nFile saved.")
            else:
                self.log("No matches found.")
                messagebox.showwarning("Result", "No updates made.")

            # Close
            try: wb.Close(SaveChanges=False)
            except: pass
            try: excel.Quit()
            except: pass

        except Exception as e:
            self.log(f"Excel Error: {e}")
            messagebox.showerror("Excel Error", str(e))
            if wb: 
                try: wb.Close(SaveChanges=False)
                except: pass
            if excel: 
                try: excel.Quit()
                except: pass

if __name__ == "__main__":
    root = tk.Tk()
    app = ShippingApp(root)
    root.mainloop()
