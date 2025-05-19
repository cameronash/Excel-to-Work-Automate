#!/usr/bin/env python3
"""
A minimal Tkinter UI to pick Excel + Word files, run the G-Slide script,
and show a simple progress bar.
"""
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import sys
import os

# Import your main logic
from run_value_into_word import main as run_script

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("G-Slide Runner")

        # Excel picker
        tk.Label(self, text="Excel file:").grid(row=0, column=0, sticky="e", padx=4, pady=4)
        self.excel_var = tk.StringVar()
        tk.Entry(self, textvariable=self.excel_var, width=50).grid(row=0, column=1, padx=4)
        tk.Button(self, text="Browse", command=self.browse_excel).grid(row=0, column=2, padx=4)

        # Word picker
        tk.Label(self, text="Word file:").grid(row=1, column=0, sticky="e", padx=4, pady=4)
        self.word_var = tk.StringVar()
        tk.Entry(self, textvariable=self.word_var, width=50).grid(row=1, column=1, padx=4)
        tk.Button(self, text="Browse", command=self.browse_word).grid(row=1, column=2, padx=4)

        # Progress bar
        self.progress = ttk.Progressbar(self, orient="horizontal", length=400, mode="indeterminate")
        self.progress.grid(row=2, column=0, columnspan=3, pady=10)

        # Run / Exit buttons
        tk.Button(self, text="Run",   width=12, command=self.start).grid(row=3, column=1, sticky="e", pady=8)
        tk.Button(self, text="Exit",  width=12, command=self.destroy).grid(row=3, column=2, sticky="w", pady=8)

        # Make the window non-resizable
        self.resizable(False, False)
        
        # Default to known good files if they exist
        default_excel = os.path.abspath("tests/assets/Sample.xlsx")
        default_word = os.path.abspath("tests/assets/Template.docm")
        
        if os.path.exists(default_excel):
            self.excel_var.set(default_excel)
        
        if os.path.exists(default_word):
            self.word_var.set(default_word)

    def browse_excel(self):
        path = filedialog.askopenfilename(
            title="Select Excel file", filetypes=[("Excel files","*.xlsx;*.xlsm")]
        )
        if path:
            self.excel_var.set(path)

    def browse_word(self):
        path = filedialog.askopenfilename(
            title="Select Word file", filetypes=[("Word docs","*.doc;*.docx;*.docm")]
        )
        if path:
            self.word_var.set(path)

    def start(self):
        excel_path = self.excel_var.get().strip()
        word_path  = self.word_var.get().strip()
        if not excel_path or not word_path:
            messagebox.showwarning("Missing file", "Please select both Excel and Word files first.")
            return

        # Disable UI and start progress
        self.progress.start(10)
        self.update_idletasks()

        # Run the heavy work in a separate thread so the UI stays responsive
        thread = threading.Thread(target=self._run_script, args=(excel_path, word_path), daemon=True)
        thread.start()

    def _run_script(self, excel_path, word_path):
        # Make sure the paths exist and are accessible
        if not os.path.exists(excel_path):
            messagebox.showerror("Error", f"Excel file not found: {excel_path}")
            self.progress.stop()
            return
        
        if not os.path.exists(word_path):
            messagebox.showerror("Error", f"Word file not found: {word_path}")
            self.progress.stop()
            return
        
        # Print paths for debugging
        print(f"\nUsing Excel file: {excel_path}")
        print(f"Using Word file: {word_path}")
        
        # We'll simulate passing CLI args by setting sys.argv
        sys_argv_backup = sys.argv.copy()
        sys.argv = [sys.argv[0],
                    "--config", "config/g-slide mapping.xlsx",
                    "--excel", excel_path,
                    "--word", word_path]
        try:
            # Initialize COM in this thread with STA model
            import pythoncom
            import time
            pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
            
            # Modify run_script to create our own instance with delay
            from gslide.excel_reader import _open_excel, _safe_close
            from run_value_into_word import load_mappings_from_excel, format_number, format_number_as_words
            from gslide.word_writer import _open_word
            
            # Load mappings
            config_path = "config/g-slide mapping.xlsx"
            mappings = load_mappings_from_excel(config_path)
            print(f"🔢 Loaded {len(mappings)} mappings from '{config_path}'")
            
            # --- open Excel with delay ---
            xl, wb = _open_excel(excel_path)
            print("📊 Excel opened, waiting for workbook to fully load...")
            time.sleep(1)  # Give Excel a second to fully load
            
            try:
                raw_values = {}
                for sheet, cell, bookmark, fmt in mappings:
                    try:
                        # Add retry logic for reading cells
                        max_retries = 3
                        for attempt in range(max_retries):
                            try:
                                val = wb.Worksheets(sheet).Range(cell).Value
                                if attempt > 0:
                                    print(f"✓ Successfully read {sheet}!{cell} on attempt {attempt+1}")
                                break
                            except Exception as e:
                                if attempt < max_retries - 1:
                                    print(f"⚠️ Retry {attempt+1} for {sheet}!{cell}: {e}")
                                    time.sleep(0.2)  # Short delay before retry
                                else:
                                    print(f"❌ Failed to read {sheet}!{cell} after {max_retries} attempts: {e}")
                                    val = None
                    except Exception:
                        val = None
                    raw_values[(sheet, cell)] = val
            finally:
                _safe_close(xl, wb)
            
            # --- open Word once ---
            word_app, doc = _open_word(word_path)
            try:
                for sheet, cell, bookmark, fmt in mappings:
                    raw = raw_values.get((sheet, cell))
                    
                    # Format based on whether formatting is specified
                    if fmt is None or fmt == "":
                        # Handle as text (words)
                        formatted = format_number_as_words(raw)
                        format_type = "text"
                    else:
                        # Handle as numeric with the specified format
                        formatted = format_number(raw, fmt)
                        format_type = "numeric"
                    
                    # Write to bookmark
                    try:
                        rng = doc.Bookmarks(bookmark).Range
                        rng.Text = formatted
                        doc.Bookmarks.Add(bookmark, rng)
                        print(f"✅ Wrote {format_type} '{formatted}' into '{bookmark}'")
                    except Exception as e:
                        print(f"❌ Error writing to bookmark '{bookmark}': {e}")
                
                doc.Save()
            finally:
                doc.Close(False)
                word_app.Quit()
                
            messagebox.showinfo("Done", "All bookmarks updated successfully!")
        except Exception as e:
            import traceback
            error_msg = traceback.format_exc()
            print(f"Error details:\n{error_msg}")
            messagebox.showerror("Error", f"Something went wrong:\n\n{str(e)}")
        finally:
            # Clean up COM
            try:
                pythoncom.CoUninitialize()
            except:
                pass
            sys.argv = sys_argv_backup
            self.progress.stop()

if __name__ == "__main__":
    App().mainloop()