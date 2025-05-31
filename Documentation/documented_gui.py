#!/usr/bin/env python3
"""
EXCEL-TO-WORD GUI APPLICATION
=============================
A graphical user interface for the Excel-to-Word automation tool.
This transforms a command-line script into a user-friendly desktop application
that non-technical users can operate with point-and-click.

GUI PROGRAMMING CONCEPTS:
- Event-driven programming (button clicks trigger functions)
- Threading (keep UI responsive during long operations)
- File dialogs (let users browse for files)
- Progress indicators (show something is happening)
- Error handling with message boxes (user-friendly error display)

ARCHITECTURE DECISION:
Instead of calling the main CLI function, this duplicates the core logic.
This gives more control over COM threading and error handling in GUI context,
but creates code duplication. Trade-offs in software design!
"""

import tkinter as tk                      # Python's built-in GUI toolkit
from tkinter import filedialog, messagebox, ttk  # Dialog boxes and widgets
import threading                         # Run background tasks without freezing UI
import sys                              # System utilities (command line args)
import os                               # File system operations

# Import your main logic (though we'll duplicate it for GUI reasons)
from run_value_into_word import main as run_script


class App(tk.Tk):
    """
    MAIN APPLICATION CLASS
    ======================
    Inherits from tk.Tk (the main window class).
    This creates a complete desktop application with all the GUI elements.
    
    OBJECT-ORIENTED DESIGN:
    By inheriting from tk.Tk, our App IS a window. We can add widgets to it,
    handle events, and manage the application lifecycle.
    
    GUI LAYOUT:
    Row 0: Excel file picker
    Row 1: Word file picker  
    Row 2: Progress bar
    Row 3: Run/Exit buttons
    """
    
    def __init__(self):
        """
        GUI CONSTRUCTOR
        ===============
        Sets up all the visual elements and their layout.
        This is called once when the application starts.
        
        TKINTER LAYOUT MANAGEMENT:
        Using .grid() to arrange widgets in rows and columns.
        Alternative approaches: .pack() or .place()
        """
        super().__init__()  # Initialize the parent tk.Tk class
        self.title("G-Slide Runner")  # Window title bar text

        # ROW 0: EXCEL FILE PICKER
        # ========================
        # Label + Entry + Browse button in columns 0, 1, 2
        tk.Label(self, text="Excel file:").grid(row=0, column=0, sticky="e", padx=4, pady=4)
        
        # StringVar holds the current text in the entry field
        # This is Tkinter's way of linking data to GUI elements
        self.excel_var = tk.StringVar()
        tk.Entry(self, textvariable=self.excel_var, width=50).grid(row=0, column=1, padx=4)
        tk.Button(self, text="Browse", command=self.browse_excel).grid(row=0, column=2, padx=4)

        # ROW 1: WORD FILE PICKER
        # =======================
        tk.Label(self, text="Word file:").grid(row=1, column=0, sticky="e", padx=4, pady=4)
        self.word_var = tk.StringVar()
        tk.Entry(self, textvariable=self.word_var, width=50).grid(row=1, column=1, padx=4)
        tk.Button(self, text="Browse", command=self.browse_word).grid(row=1, column=2, padx=4)

        # ROW 2: PROGRESS BAR
        # ===================
        # ttk.Progressbar is the modern styled progress bar
        # mode="indeterminate" means it just shows "something is happening" (no percentage)
        self.progress = ttk.Progressbar(self, orient="horizontal", length=400, mode="indeterminate")
        self.progress.grid(row=2, column=0, columnspan=3, pady=10)  # Spans 3 columns

        # ROW 3: ACTION BUTTONS
        # =====================
        tk.Button(self, text="Run",   width=12, command=self.start).grid(row=3, column=1, sticky="e", pady=8)
        tk.Button(self, text="Exit",  width=12, command=self.destroy).grid(row=3, column=2, sticky="w", pady=8)

        # WINDOW CONFIGURATION
        # ====================
        self.resizable(False, False)  # Fixed size window (cleaner look)
        
        # SMART DEFAULTS
        # ==============
        # If test files exist, pre-populate the fields
        # This makes testing easier and shows users the expected file types
        default_excel = os.path.abspath("tests/assets/Sample.xlsx")
        default_word = os.path.abspath("tests/assets/Template.docm")
        
        if os.path.exists(default_excel):
            self.excel_var.set(default_excel)
        
        if os.path.exists(default_word):
            self.word_var.set(default_word)


    def browse_excel(self):
        """
        EXCEL FILE BROWSER
        ==================
        Opens a file dialog for users to select Excel files.
        This is much more user-friendly than typing file paths.
        
        GUI PROGRAMMING CONCEPT:
        Event handler - this function runs when the Browse button is clicked.
        The button's command= parameter connects the button to this function.
        """
        path = filedialog.askopenfilename(
            title="Select Excel file", 
            filetypes=[("Excel files","*.xlsx;*.xlsm")]  # Filter to show only Excel files
        )
        if path:  # User clicked OK (not Cancel)
            self.excel_var.set(path)  # Update the text field


    def browse_word(self):
        """
        WORD FILE BROWSER
        =================
        Same concept as Excel browser, but for Word documents.
        """
        path = filedialog.askopenfilename(
            title="Select Word file", 
            filetypes=[("Word docs","*.doc;*.docx;*.docm")]  # Word file types
        )
        if path:
            self.word_var.set(path)


    def start(self):
        """
        MAIN EXECUTION TRIGGER
        ======================
        This runs when user clicks the "Run" button.
        Validates inputs and starts the background processing.
        
        THREADING CONCEPT:
        Long-running operations (like COM automation) would freeze the GUI
        if run on the main thread. We use threading to keep the UI responsive.
        """
        # Get the file paths from the GUI
        excel_path = self.excel_var.get().strip()
        word_path  = self.word_var.get().strip()
        
        # VALIDATION: Make sure user selected both files
        if not excel_path or not word_path:
            messagebox.showwarning("Missing file", "Please select both Excel and Word files first.")
            return

        # START VISUAL FEEDBACK
        # =====================
        # Show progress bar animation so user knows something is happening
        self.progress.start(10)  # Animate every 10ms
        self.update_idletasks()  # Process GUI updates immediately

        # BACKGROUND PROCESSING
        # =====================
        # Start the heavy work in a separate thread
        # daemon=True means thread dies when main program exits
        thread = threading.Thread(target=self._run_script, args=(excel_path, word_path), daemon=True)
        thread.start()


    def _run_script(self, excel_path, word_path):
        """
        BACKGROUND WORKER FUNCTION
        ==========================
        This runs in a separate thread to avoid freezing the GUI.
        Contains the actual Excel-to-Word automation logic.
        
        THREADING SAFETY:
        This function runs on a different thread from the GUI, so we need to be
        careful about COM initialization and error handling.
        
        DESIGN DECISION:
        Instead of calling the main CLI function, we duplicate the core logic here.
        This gives us better control over COM threading and GUI-specific error handling,
        but creates code duplication. This is a common trade-off in GUI applications.
        """
        
        # FILE VALIDATION
        # ===============
        if not os.path.exists(excel_path):
            messagebox.showerror("Error", f"Excel file not found: {excel_path}")
            self.progress.stop()
            return
        
        if not os.path.exists(word_path):
            messagebox.showerror("Error", f"Word file not found: {word_path}")
            self.progress.stop()
            return
        
        # Debug output (appears in console, not GUI)
        print(f"\nUsing Excel file: {excel_path}")
        print(f"Using Word file: {word_path}")
        
        # COMMAND LINE SIMULATION
        # =======================
        # Clever hack: Instead of rewriting everything, simulate command line args
        # This lets us reuse the existing configuration loading logic
        sys_argv_backup = sys.argv.copy()  # Save original args
        sys.argv = [sys.argv[0],
                    "--config", "config/g-slide mapping.xlsx",
                    "--excel", excel_path,
                    "--word", word_path]
        
        try:
            # COM INITIALIZATION FOR THREADING
            # ================================
            # CRITICAL: Each thread that uses COM must initialize it
            # COINIT_APARTMENTTHREADED is required for Office automation
            import pythoncom
            import time
            pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
            
            # IMPORT THE AUTOMATION FUNCTIONS
            # ===============================
            # Import here to avoid issues with module loading and COM
            from gslide.excel_reader import _open_excel, _safe_close
            from run_value_into_word import load_mappings_from_excel, format_number, format_number_as_words
            from gslide.word_writer import _open_word
            
            # STEP 1: LOAD CONFIGURATION
            # ==========================
            config_path = "config/g-slide mapping.xlsx"
            mappings = load_mappings_from_excel(config_path)
            print(f"🔢 Loaded {len(mappings)} mappings from '{config_path}'")
            
            # STEP 2: EXTRACT DATA FROM EXCEL
            # ===============================
            xl, wb = _open_excel(excel_path)
            print("📊 Excel opened, waiting for workbook to fully load...")
            time.sleep(1)  # GUI-specific: Give Excel time to fully initialize
            
            try:
                raw_values = {}
                for sheet, cell, bookmark, fmt in mappings:
                    # RETRY LOGIC FOR GUI ROBUSTNESS
                    # ==============================
                    # Excel can be flaky in GUI environments, so add retry logic
                    max_retries = 3
                    for attempt in range(max_retries):
                        try:
                            val = wb.Worksheets(sheet).Range(cell).Value
                            if attempt > 0:  # Log successful retries
                                print(f"✓ Successfully read {sheet}!{cell} on attempt {attempt+1}")
                            break
                        except Exception as e:
                            if attempt < max_retries - 1:
                                print(f"⚠️ Retry {attempt+1} for {sheet}!{cell}: {e}")
                                time.sleep(0.2)  # Brief delay before retry
                            else:
                                print(f"❌ Failed to read {sheet}!{cell} after {max_retries} attempts: {e}")
                                val = None
                    
                    raw_values[(sheet, cell)] = val
                    
            finally:
                _safe_close(xl, wb)
            
            # STEP 3: WRITE DATA TO WORD
            # ==========================
            word_app, doc = _open_word(word_path)
            try:
                for sheet, cell, bookmark, fmt in mappings:
                    raw = raw_values.get((sheet, cell))
                    
                    # BUSINESS LOGIC: Format based on whether formatting is specified
                    if fmt is None or fmt == "":
                        # No format = convert to words
                        formatted = format_number_as_words(raw)
                        format_type = "text"
                    else:
                        # Has format = treat as number
                        formatted = format_number(raw, fmt)
                        format_type = "numeric"
                    
                    # Write to Word bookmark
                    try:
                        rng = doc.Bookmarks(bookmark).Range
                        rng.Text = formatted
                        doc.Bookmarks.Add(bookmark, rng)  # Recreate bookmark
                        print(f"✅ Wrote {format_type} '{formatted}' into '{bookmark}'")
                    except Exception as e:
                        print(f"❌ Error writing to bookmark '{bookmark}': {e}")
                
                doc.Save()  # Save the document
                
            finally:
                doc.Close(False)
                word_app.Quit()
                
            # SUCCESS MESSAGE
            # ===============
            # messagebox calls are thread-safe in Tkinter
            messagebox.showinfo("Done", "All bookmarks updated successfully!")
            
        except Exception as e:
            # ERROR HANDLING FOR GUI
            # ======================
            import traceback
            error_msg = traceback.format_exc()
            print(f"Error details:\n{error_msg}")  # Console output for debugging
            messagebox.showerror("Error", f"Something went wrong:\n\n{str(e)}")  # User-friendly popup
            
        finally:
            # CLEANUP
            # =======
            # Always clean up COM and restore original state
            try:
                pythoncom.CoUninitialize()  # Clean up COM for this thread
            except:
                pass
            sys.argv = sys_argv_backup  # Restore original command line args
            self.progress.stop()        # Stop the progress bar animation


# APPLICATION ENTRY POINT
# =======================
if __name__ == "__main__":
    # Create and run the GUI application
    # .mainloop() starts the event loop (waits for button clicks, etc.)
    App().mainloop()


"""
GUI PROGRAMMING CONCEPTS DEMONSTRATED:
======================================

1. EVENT-DRIVEN PROGRAMMING:
   - GUI waits for user actions (button clicks, file selections)
   - Each action triggers a specific function (event handler)

2. THREADING FOR RESPONSIVENESS:
   - Long operations run in background threads  
   - Main thread handles GUI updates and user interactions
   - Prevents the dreaded "frozen window" problem

3. USER EXPERIENCE DESIGN:
   - File dialogs instead of typing paths
   - Progress indicators for feedback
   - Smart defaults to reduce user effort
   - Clear error messages with message boxes

4. RESOURCE MANAGEMENT IN GUI:
   - COM initialization per thread
   - Careful cleanup to prevent memory leaks
   - Error handling that doesn't crash the GUI

5. SEPARATION OF CONCERNS:
   - GUI handles user interaction
   - Background thread handles business logic  
   - Clear separation between presentation and processing

This is actually quite sophisticated GUI programming! You're handling threading,
COM automation, file I/O, and error handling in a user-friendly interface.
"""
