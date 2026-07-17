[README.md](https://github.com/user-attachments/files/30142916/README.md)
# Donut Data Cleaner

A cute, candy-colored desktop app that cleans up messy spreadsheets in one click. Built for small business owners, students, and anyone drowning in unorganized CSV or Excel files who just wants clean, usable data without writing a single line of code.

## What It Does

Donut Data Cleaner takes a messy spreadsheet and automatically:

- Removes exact duplicate rows
- Strips extra whitespace from text columns
- Standardizes column names into clean snake_case format
- Detects and reformats date columns
- Detects and converts number-looking columns (even ones with $ or % symbols) into proper numeric values
- Flags any missing values left over, so nothing silently disappears
- Lets you preview both the raw and cleaned data side by side before exporting

No formulas, no scripting, no manual find-and-replace. Load a file, click Clean Data, review the result, and export.

## Features

- One-click cleanup pipeline: load a file and get a cleaned version in seconds
- Side-by-side preview tabs: compare Raw Data and Cleaned Data before exporting
- Activity log: a running list of exactly what was changed (duplicates removed, columns renamed, etc.)
- Flexible file support: works with .csv, .xlsx, and .xls files
- Export anywhere: save your cleaned result back out as CSV or Excel
- Chunky, candy-colored pixel-art interface with textured buttons and a donut icon
- Runs as a standalone app: no Python installation required once packaged

## Getting Started

### Option 1: Run the pre-built app (recommended for most users)
If you were given a packaged installer or standalone executable:

- Windows: run the installer (DonutDataCleaner_Setup.exe) or double-click DonutDataCleaner.exe
- Mac: open the .dmg file and drag Donut Data Cleaner into your Applications folder

No Python, no terminal, no setup required.

### Option 2: Run from source (for developers)
If you have Python 3.9+ installed:

    pip install -r requirements.txt
    python donut_data_cleaner.py

Dependencies: pandas, openpyxl (installed automatically via requirements.txt above). Tkinter comes bundled with most standard Python installs.

## How to Use It

1. Click Load File and select a CSV or Excel file.
2. Review the Raw Data tab to see what you started with.
3. Click Clean Data — the app automatically removes duplicates, fixes column names, trims whitespace, and reformats dates/numbers.
4. Switch to the Cleaned Data tab to review the result and check the activity log for a summary of changes.
5. Click Export File and choose CSV or Excel to save your cleaned spreadsheet.
6. Use Reset anytime to clear everything and start over with a new file.

## What Gets Cleaned Automatically

| Issue in raw data | What Donut Data Cleaner does |
|---|---|
| Duplicate rows | Removed entirely, keeping only one copy |
| Inconsistent column names (e.g. "Customer Name ") | Converted to clean snake_case (e.g. customer_name) |
| Extra spaces in text cells | Trimmed automatically |
| Dates stored as text | Detected and converted to proper date format |
| Numbers stored as text (e.g. "$1,200" or "45%") | Detected and converted to usable numeric values |
| Empty or blank cells | Flagged in the activity log so you know where gaps remain |

## Project Structure

    donut_data_cleaner.py     - the main application
    app_icon.ico              - Windows app icon
    donut_icon.png            - source icon image
    requirements.txt          - Python dependencies

## Building a Standalone Installer

This project supports building native installers for both Windows and macOS via PyInstaller and GitHub Actions, so end users never need Python installed. See README_GITHUB_BUILD.md for full instructions on setting up automated cloud builds, or the individual build_windows-1.bat / build_mac_app-1.sh / build_mac_installer-1.sh scripts for manual local builds.

## Notes

- First launch of a packaged app may take a few extra seconds while it unpacks — this is normal for onefile builds.
- Unsigned builds may trigger a security warning on first launch (Windows SmartScreen or macOS Gatekeeper) since the app isn't code-signed. This is expected for indie apps — right-click and choose Open (Mac) or "More info -> Run anyway" (Windows) to bypass it.
- The app never sends your data anywhere — all cleaning happens locally on your machine.

## Credits

Built by Adrian Mohammed.
