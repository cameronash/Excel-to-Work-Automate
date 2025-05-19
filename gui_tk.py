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
import shutil
import time
import traceback

# Determine application base path (works for both script and frozen exe)
def get_base_path():
    if getattr(sys, 'frozen', False):
        # If running as compiled executable
        return os.path.dirname(sys.executable)
    else:
        # If running as script
        return os.path.dirname(os.path.abspath(__file__))

# Make sure we can import run_value_into_word regardless of how we're running
base_path = get_base_path()
if base_path not in sys.path:
    sys.path.insert(0, base_path)

# Import your main logic
try:
    from run_value_into_word import main as run_script
    from run_value_into_word import load_mappings_from_excel, format_number, format_number_as_words
except ImportError as e:
    messagebox.showerror("Import Error", f"Could not import required modules: {e}")
    sys.exit(1)

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("G-Slide Runner")
        self.base_path = get_base_path()

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
        tk.Button(self, text="Run", width=12, command=self.start).grid(row=3, column=1, sticky="e", pady=8)
        tk.Button(self, text="Exit", width=12, command=self.destroy).grid(row=3, column=2, sticky="w", pady=8)

        # Make the window non-resizable
        self.resizable(False, False)
        
        # Default to known good files if they exist
        self.set_default_files()

    def set_default_files(self):
        """Set default Excel and Word files if they exist"""
        # Try multiple possible locations for the sample files
        possible_paths = [
            # Relative to executable or script
            os.path.join(self.base_path, "tests", "assets"),
            # One directory up (if in dist folder)
            os.path.join(os.path.dirname(self.base_path), "tests", "assets"),
            # Absolute path if app was previously in C:\g-slide-next
            r"C:\g-slide-next\tests\assets"
        ]
        
        for path in possible_paths:
            excel_path = os.path.join(path, "Sample.xlsx")
            word_path = os.path.join(path, "Template.docm")
            
            if os.path.exists(excel_path) and os.path.exists(word_path):
                self.excel_var.set(excel_path)
                self.word_var.set(word_path)
                print(f"Found default files in: {path}")
                break
        
        if not self.excel_var.get():
            print("No default files found. Please browse for files.")

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

    def get_config_path(self):
        """Find the config file, considering both script and executable modes"""
        # List of possible config file locations
        possible_paths = [
            # Direct relative to exe or script
            os.path.join(self.base_path, "config", "g-slide mapping.xlsx"),
            # One folder up (if in dist folder)
            os.path.join(os.path.dirname(self.base_path), "config", "g-slide mapping.xlsx"),
            # Inside a temporary _MEI folder (PyInstaller's temp folder)
            os.path.join(getattr(sys, '_MEIPASS', self.base_path), "config", "g-slide mapping.xlsx"),
            # Absolute path if app was previously in C:\g-slide-next
            r"C:\g-slide-next\config\g-slide mapping.xlsx"
        ]
        
        # Try each path
        for path in possible_paths:
            if os.path.exists(path):
                print(f"Found config file: {path}")
                return path
        
        # If no config file found, ask user to select it
        messagebox.showinfo("Config File Not Found", 
                           "Please select the 'g-slide mapping.xlsx' config file.")
        path = filedialog.askopenfilename(
            title="Select g-slide mapping.xlsx", 
            filetypes=[("Excel files","*.xlsx")]
        )
        
        if path:
            # Try to copy the selected config file to the expected location
            try:
                target_dir = os.path.join(self.base_path, "config")
                os.makedirs(target_dir, exist_ok=True)
                target_path = os.path.join(target_dir, "g-slide mapping.xlsx")
                shutil.copy2(path, target_path)
                print(f"Copied config file to: {target_path}")
                return target_path
            except Exception as e:
                print(f"Could not copy config file: {e}")
                return path
        
        return None  # No config file found or selected

    def start(self):
        excel_path = self.excel_var.get().strip()
        word_path = self.word_var.get().strip()
        
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
        
        # Get the config file path
        config_path = self.get_config_path()
        if not config_path:
            messagebox.showerror("Error", "Config file not found and not selected.")
            self.progress.stop()
            return
        
        # Print paths for debugging
        print(f"\nOriginal Excel file: {excel_path}")
        print(f"Original Word file: {word_path}")
        print(f"Original config file: {config_path}")
        
        # Create a temporary directory for local processing
        import tempfile
        temp_dir = tempfile.mkdtemp()
        print(f"Created temporary directory: {temp_dir}")
        
        # Create local copies of files
        local_excel = os.path.join(temp_dir, os.path.basename(excel_path))
        local_word = os.path.join(temp_dir, os.path.basename(word_path))
        local_config = os.path.join(temp_dir, os.path.basename(config_path))
        
        try:
            # Copy files locally
            print(f"Copying Excel file to local temp directory...")
            shutil.copy2(excel_path, local_excel)
            
            print(f"Copying Word file to local temp directory...")
            shutil.copy2(word_path, local_word)
            
            print(f"Copying config file to local temp directory...")
            shutil.copy2(config_path, local_config)
            
            print(f"\nUsing local Excel file: {local_excel}")
            print(f"Using local Word file: {local_word}")
            print(f"Using local config file: {local_config}")
            
            # Initialize COM in this thread with STA model
            import pythoncom
            pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
            
            # Import needed modules here to ensure proper COM initialization
            try:
                from gslide.excel_reader import _open_excel, _safe_close
                from gslide.word_writer import _open_word
            except ImportError as e:
                # If running as exe and can't find gslide module
                if getattr(sys, 'frozen', False):
                    # Try to add src to path
                    for src_path in [
                        os.path.join(self.base_path, "src"),
                        os.path.join(os.path.dirname(self.base_path), "src"),
                        r"C:\g-slide-next\src"
                    ]:
                        if os.path.exists(src_path):
                            sys.path.insert(0, src_path)
                            print(f"Added to path: {src_path}")
                            break
                    
                    # Try import again
                    from gslide.excel_reader import _open_excel, _safe_close
                    from gslide.word_writer import _open_word
            
            # Load mappings from the local config
            mappings = load_mappings_from_excel(local_config)
            print(f"🔢 Loaded {len(mappings)} mappings from '{local_config}'")
            
            # --- open Excel with delay ---
            xl, wb = _open_excel(local_excel)
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
            word_app, doc = _open_word(local_word)
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
            
            # Copy the updated Word document back to the original location
            print(f"Copying updated Word file back to original location: {word_path}")
            shutil.copy2(local_word, word_path)
                
            messagebox.showinfo("Done", "All bookmarks updated successfully!")
        except Exception as e:
            error_msg = traceback.format_exc()
            print(f"Error details:\n{error_msg}")
            messagebox.showerror("Error", f"Something went wrong:\n\n{str(e)}")
        finally:
            # Clean up COM
            try:
                pythoncom.CoUninitialize()
            except:
                pass
            
            # Clean up temporary directory
            try:
                print(f"Cleaning up temporary directory: {temp_dir}")
                shutil.rmtree(temp_dir)
            except Exception as e:
                print(f"Warning: Could not remove temp directory: {e}")
            
            self.progress.stop()

if __name__ == "__main__":
    App().mainloop()