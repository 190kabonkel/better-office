================================================================================
                better office - README
================================================================================

VERSION: 1.0
DATE: 2026

================================================================================
INTRODUCTION
================================================================================

better office is an easy-to-use office application that includes three
main tools:
  1. Word Processor - Create and edit text documents
  2. Spreadsheet - Work with tables and data
  3. Presentation - Create slide presentations

================================================================================
SYSTEM REQUIREMENTS
================================================================================

- Python 3.6 or higher
- tkinter (usually included with Python)
- Works on Windows, Mac, and Linux

================================================================================
INSTALLATION & RUNNING
================================================================================

1. Make sure Python is installed on your computer
2. Double-click the office_app.py file, OR
3. Open a terminal/command prompt and run:
   python office_app.py
   OR
   python3 office_app.py

================================================================================
FEATURES
================================================================================

--------------------------
WORD PROCESSOR FEATURES
--------------------------

✓ Text Editing:
  - Type and edit text freely
  - Undo/Redo support
  - Copy, Cut, Paste (use Ctrl+C, Ctrl+X, Ctrl+V)

✓ Text Formatting:
  - Change font family (Arial, Times, etc.)
  - Change font size (8-32 points)
  - Bold, Italic, Underline
  - Text color picker
  - Text alignment (Left, Center, Right)

✓ File Operations:
  - Save as .txt or .rtf format
  - Open existing text files
  - Save with custom filename

✓ Statistics:
  - Live word count
  - Character count

--------------------------
SPREADSHEET FEATURES
--------------------------

✓ Grid Operations:
  - 20 rows x 10 columns default grid
  - Add new rows
  - Add new columns
  - Delete rows
  - Delete columns
  - Scrollable grid for large data

✓ Data Entry:
  - Enter text or numbers in cells
  - Column headers (A, B, C...)
  - Row numbers (1, 2, 3...)

✓ Calculations:
  - Calculate sum of all numeric values
  - Calculate average

✓ File Operations:
  - Save as CSV (Comma Separated Values)
  - Save as TSV (Tab Separated Values)
  - Open existing CSV files
  - Excel-compatible format

--------------------------
PRESENTATION FEATURES
--------------------------

✓ Slide Management:
  - Create multiple slides
  - Delete slides
  - Navigate between slides (Previous/Next)
  - View slide list

✓ Slide Content:
  - Add title to each slide
  - Add content/body text
  - Edit slide title and content
  - Slide counter (shows current slide number)

✓ File Operations:
  - Save presentation as JSON format
  - Open existing presentations
  - Preserve all slides and content

================================================================================
HOW TO USE
================================================================================

--------------------------
WORD PROCESSOR
--------------------------

1. Click on the "Word Processor" tab
2. Start typing your document
3. Use the toolbar to format text:
   - Select text, then click Bold (B), Italic (I), or Underline (U)
   - Select text, then click "Color" to change color
   - Select text, then click alignment buttons (←, ↔, →)
4. To save:
   - Click File → Save (if previously saved)
   - Click File → Save As to choose name and location
5. To open a file:
   - Click File → Open
   - Choose your .txt file

--------------------------
SPREADSHEET
--------------------------

1. Click on the "Spreadsheet" tab
2. Click on any cell and type data
3. Use toolbar buttons:
   - "Add Row" - adds a new row at the bottom
   - "Add Column" - adds a new column on the right
   - "Delete Row" - removes the last row
   - "Delete Column" - removes the last column
   - "Calculate Sum" - shows total and average of all numbers
4. To save:
   - Click File → Save As
   - Choose location and filename
   - Select CSV or TSV format
5. To open:
   - Click File → Open
   - Choose your CSV file

--------------------------
PRESENTATION
--------------------------

1. Click on the "Presentation" tab
2. A default slide is created automatically
3. Edit the slide:
   - Type a title in the "Title" field
   - Type content in the "Content" area
4. To add more slides:
   - Click "New Slide" button
5. To navigate:
   - Click "← Prev" or "Next →" buttons
   - OR click on slides in the slide list (right side)
6. To delete a slide:
   - Navigate to the slide
   - Click "Delete Slide" button
7. To save:
   - Click File → Save As
   - Choose location and filename (saves as .json)
8. To open:
   - Click File → Open
   - Choose your .json presentation file

================================================================================
MENU BAR OPTIONS
================================================================================

FILE MENU:
- New: Creates a new blank document/spreadsheet/presentation
- Open: Opens an existing file
- Save: Saves the current file (if previously saved)
- Save As: Saves with a new name or location
- Exit: Closes the application

HELP MENU:
- About: Shows application information

================================================================================
FILE FORMATS
================================================================================

Word Processor:
- .txt (Plain Text) - Simple text format, works everywhere
- .rtf (Rich Text Format) - Preserves some formatting

Spreadsheet:
- .csv (Comma Separated Values) - Standard spreadsheet format
- .tsv (Tab Separated Values) - Tab-delimited format
- Can be opened in Excel, Google Sheets, etc.

Presentation:
- .json (JSON format) - Stores slides and content
- Can only be opened by this application

================================================================================
KEYBOARD SHORTCUTS
================================================================================

Word Processor:
- Ctrl+Z: Undo
- Ctrl+Y: Redo
- Ctrl+C: Copy
- Ctrl+X: Cut
- Ctrl+V: Paste
- Ctrl+A: Select All

================================================================================
TIPS & TRICKS
================================================================================

1. SAVING YOUR WORK:
   - Save frequently to avoid losing data
   - Use "Save As" to create backup copies
   - Give files descriptive names

2. FORMATTING TEXT:
   - Select text first, then apply formatting
   - Try different fonts and sizes
   - Use alignment for better layout

3. SPREADSHEET:
   - Keep your data organized in columns
   - Use the first row for headers
   - Save as CSV to open in other programs

4. PRESENTATIONS:
   - Keep slide titles short and clear
   - Use bullet points in content
   - Save after creating each slide

5. FILE ORGANIZATION:
   - Create separate folders for documents, spreadsheets, and presentations
   - Use consistent naming conventions
   - Include dates in filenames (e.g., "Report_2026-03-01.txt")

================================================================================
TROUBLESHOOTING
================================================================================

Problem: Application won't start
Solution: Make sure Python is installed. Try running from terminal/command prompt.

Problem: Can't save files
Solution: Make sure you have write permissions in the folder you're saving to.

Problem: File won't open
Solution: Make sure the file format matches the module:
  - Open .txt files in Word Processor
  - Open .csv files in Spreadsheet
  - Open .json files in Presentation

Problem: Lost formatting after saving
Solution: Text files (.txt) don't preserve formatting. The formatting is visual only
in the current session.

Problem: Spreadsheet calculations not working
Solution: Make sure cells contain only numbers (no text) for calculations.

================================================================================
LIMITATIONS
================================================================================

- Word Processor formatting is visual only and not saved to .txt files
- Spreadsheet has basic features (no advanced formulas)
- Presentations are saved in custom JSON format
- No spell-checking or grammar checking
- No image insertion (text only)
- No printing functionality (use system print after saving)

================================================================================
FUTURE ENHANCEMENTS
================================================================================

Planned features for future versions:
- Print support
- More font options
- Image insertion
- Advanced spreadsheet formulas
- Presentation themes
- Export to PDF
- Cloud storage integration

================================================================================
SUPPORT & CONTACT
================================================================================

Email kabonkel@proton.me

================================================================================
LICENSE
================================================================================

Free to use, modify, and distribute.

================================================================================
VERSION HISTORY
================================================================================

Version 1.0 (2026-03-01):
- Initial release
- Word Processor with formatting
- Spreadsheet with grid
- Presentation with slides
- Save/Open functionality
- Multiple file format support

================================================================================
CREDITS
================================================================================

Built with:
- Python 3
- tkinter (GUI framework)
- Standard Python libraries

================================================================================
                            ENJOY USING BETTER OFFICE!
================================================================================
