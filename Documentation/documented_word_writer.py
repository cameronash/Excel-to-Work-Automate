"""
WORD AUTOMATION MODULE
======================
This module handles Word document automation using Windows COM.
It provides a simple interface for writing text to Word bookmarks.

WORD BOOKMARKS EXPLAINED:
Word bookmarks are like invisible anchors in documents. You can place them
anywhere (like after "Total Revenue:" in a report) and then programmatically
insert text at that exact location. They're perfect for document templates.

WHY THIS IS USEFUL:
Instead of manually opening Word, finding the right spot, and typing numbers,
this lets you automatically populate Word templates with data from Excel,
databases, or calculations.

PROGRAMMING CONCEPTS DEMONSTRATED:
- COM automation (controlling Word applications)
- Flexible resource management (reuse existing sessions or create new ones)
- Defensive programming (handle missing bookmarks gracefully)
- Context-aware operations (readonly vs editable documents)
- Proper cleanup of external resources
"""

from __future__ import annotations  # Modern type hints
import logging                      # Professional logging
import win32com.client as win32     # Windows COM automation
import pythoncom                    # Core COM support

# Set up simple logging configuration
# This replaces print() statements with proper logging levels
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


def _open_word(doc_path: str, readonly: bool = True):
    """
    WORD LAUNCHER
    =============
    Opens Microsoft Word and loads a specific document.
    This is the COM automation entry point for Word.
    
    Args:
        doc_path (str): Full path to Word document (.docx, .doc)
        readonly (bool): If True, opens in read-only mode (safer for templates)
        
    Returns:
        tuple: (word_application, document) - both are COM objects
        
    PROGRAMMING CONCEPTS:
    - COM initialization: Required before automating Office apps
    - Application control: Starting and configuring Word
    - Document opening: Loading files programmatically
    - Read-only safety: Protecting template files from accidental changes
    
    WHY READ-ONLY BY DEFAULT:
    If you're working with templates, you usually want to open them read-only,
    make your changes, then save as a new file. This protects the original
    template from being accidentally modified.
    """
    
    # Initialize COM for this thread
    # Every thread that uses COM must call this first
    pythoncom.CoInitialize()
    
    # Connect to Word application (starts Word if not running)
    word_app = win32.Dispatch("Word.Application")
    
    try:
        # Try to make Word invisible (run in background)
        # Some security policies might prevent this - that's OK
        word_app.Visible = False
    except AttributeError:
        # If we can't control visibility, just continue
        # Word might be visible, but that's fine
        pass
    
    # Open the document with specific settings
    doc = word_app.Documents.Open(
        doc_path,                    # Path to document
        ReadOnly=readonly,           # Protect template from changes
        AddToRecentFiles=False,      # Don't clutter Word's recent files list
    )
    
    return word_app, doc


def write_to_bookmark(
    doc_path: str,
    bookmark: str,
    text,
    *,  # Everything after * must be passed as keyword arguments
    word_app: win32.Dispatch | None = None,
    doc=None,
    readonly_template: bool = True,
) -> None:
    """
    BOOKMARK TEXT WRITER
    ====================
    Writes text to a specific bookmark in a Word document.
    This is the main function that other modules use.
    
    FLEXIBLE DESIGN:
    This function can work in two modes:
    1. STANDALONE: Opens Word, writes text, closes Word (simple)
    2. BATCH: Uses existing Word session for multiple operations (efficient)
    
    Args:
        doc_path (str): Path to Word document
        bookmark (str): Name of bookmark to write to
        text (Any): Text to insert (converted to string)
        word_app (optional): Existing Word application (for batch operations)
        doc (optional): Existing document handle (for batch operations)
        readonly_template (bool): If True, treat as template and SaveAs
        
    Returns:
        None: Function performs side effects (modifies document)
        
    PROGRAMMING CONCEPTS:
    - Flexible resource management: Handle both standalone and batch modes
    - Keyword-only arguments: Using * to force named parameters
    - Exception handling: Graceful degradation on errors
    - COM object manipulation: Working with Word's object model
    - Bookmark recreation: Word's quirky behavior with bookmarks
    
    WORD BOOKMARK QUIRK:
    When you change the text at a bookmark, Word automatically DELETES the 
    bookmark! So we have to recreate it after inserting text. This is just
    how Word works - annoying but we deal with it.
    """
    
    # Determine if we need to manage Word session ourselves
    own_session = word_app is None
    
    if own_session:
        # STANDALONE MODE: Open Word just for this operation
        word_app, doc = _open_word(doc_path, readonly_template)

    try:
        # STEP 1: Check if bookmark exists
        # Better to check first than handle the error later
        if not doc.Bookmarks.Exists(bookmark):
            logging.warning(f"Bookmark '{bookmark}' not found → skipping")
            return  # Exit gracefully, don't crash the whole operation

        # STEP 2: Get the bookmark's location in the document
        # Bookmarks in Word are "ranges" - they mark a position or selection
        rng = doc.Bookmarks(bookmark).Range
        
        # STEP 3: Replace the text at that location
        rng.Text = str(text)  # Convert anything to string (numbers, dates, etc.)

        # STEP 4: CRITICAL - Recreate the bookmark!
        # Word deletes bookmarks when you change their text (crazy, right?)
        # So we immediately recreate it at the same location
        doc.Bookmarks.Add(bookmark, rng)
        
        # Log success for debugging/monitoring
        logging.info(f"Wrote {text!r} into bookmark '{bookmark}'")

    except pythoncom.com_error as e:
        # Handle COM-specific errors (Word crashes, document corruption, etc.)
        # Log the error but don't crash - other bookmarks might still work
        logging.warning(f"COM error on bookmark '{bookmark}': {e}")
        
    except Exception as e:
        # Handle any other unexpected errors
        # This is defensive programming - don't let one bad bookmark kill everything
        logging.warning(f"Unexpected error on bookmark '{bookmark}': {e}")

    finally:
        # CLEANUP: Only if we opened Word ourselves
        if own_session:
            # This is the cleanup section - always runs, even if errors occurred
            
            try:
                # SAVE LOGIC: Different behavior for templates vs regular docs
                if readonly_template:
                    # For templates: Save as new file (preserve original template)
                    doc.SaveAs2(doc_path)
                else:
                    # For regular docs: Just save changes to same file
                    doc.Save()
            except Exception:
                # If saving fails, just continue - at least close Word properly
                pass
            
            try:
                # Close the document
                doc.Close(False)  # False = don't save again
            except Exception:
                pass
            
            try:
                # Quit Word application
                word_app.Quit()
            except Exception:
                pass


# BACKWARDS COMPATIBILITY
# =======================
# Provide old function name in case other code is using it
# This is good practice when refactoring - don't break existing code
write_value_to_bookmark = write_to_bookmark


"""
USAGE EXAMPLES
==============

SIMPLE USAGE (one bookmark):
    write_to_bookmark("report_template.docx", "total_revenue", "1,500,000")

BATCH USAGE (multiple bookmarks - more efficient):
    word_app, doc = _open_word("report_template.docx")
    try:
        write_to_bookmark("", "revenue", "1,500,000", word_app=word_app, doc=doc)
        write_to_bookmark("", "profit", "350,000", word_app=word_app, doc=doc) 
        write_to_bookmark("", "growth", "15%", word_app=word_app, doc=doc)
        doc.Save()
    finally:
        doc.Close(False)
        word_app.Quit()

CREATING BOOKMARKS IN WORD:
1. Open your Word document
2. Place cursor where you want the bookmark
3. Insert → Bookmark → Type name → Add
4. Now you can reference it by name in your code
"""
