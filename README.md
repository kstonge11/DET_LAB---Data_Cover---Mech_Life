# DET Lab Cover Sheet Tool

A PyQt6 desktop application for HP printer validation testing. This tool streamlines the process of logging print errors and generating cover sheets for defect tracking during DET (Design Evaluation Testing) lab sessions.

## Overview

The DET Lab Cover Sheet Tool allows test technicians to:
- Load test data from Excel workbooks (one sheet per printer)
- Navigate through test lines and view test parameters
- Log errors with specific error types, affected printers, and page numbers
- Generate and print cover sheets for physical defect samples
- Save cover sheet files to a specified directory

## Features

### Test Info Panel
- **Printer Selection**: Switch between 8 printer units (TH5BA5X058 - TH5BA5X101)
- **Line Navigation**: Browse through test lines with auto-populated test parameters
- **Auto-loaded Fields**: Climate, Unit, Media #, Script #, Plexity, Load #, Run #, Phase, Firmware
- **Page Count Display**: Reads total page count from cell N6 of each printer sheet

### Session Settings
- **KaMP #**: Test campaign identifier
- **Operator**: Technician name
- **Firmware**: Current firmware version
- **Phase #**: Test phase number

All settings are edited via a single dialog and persist throughout the session.

### Error Logging
- **Error Type Selection**: 60+ predefined error categories including:
  - Pick Issues (skew, no pick, multipick, etc.)
  - Print Quality (smears, banding, ink drops, etc.)
  - Jams (duplex/simplex, input/output, damage variants)
  - Stacking Errors
  - Media Damage
  - Unit Errors
- **Printer Selection**: Select one or multiple affected printers
- **Page Tally**: Click to record each affected page number
- **Notes**: Optional notes field for additional context

### Queue Management
- View all queued errors before processing
- **Save Covers**: Generate cover sheet Excel files to a chosen directory
- **Print Covers**: Send cover sheets directly to the default printer
- **Clear Queue**: Remove all queued entries

## Installation

### Requirements
- Python 3.10+
- Windows (for printing via COM automation) or macOS (for printing via AppleScript)

### Dependencies
```
PyQt6>=6.4.0
pandas>=1.5.0
openpyxl>=3.0.0
pywin32>=305  # Windows only, for printing
```

### Setup
```bash
# Clone or download the project
cd "DET LAB"

# Install dependencies
pip install -r requirements.txt

# Run the application
python main.py
```

## Usage

### Getting Started
1. **Launch the app**: `python main.py`
2. **Load Data File**: Click "Browse Data" and select your test Excel workbook
3. **Load Cover Template**: Click "Browse Cover" and select your cover sheet template
4. **Set Session Info**: Click "Edit Settings" to enter KaMP#, Operator, Firmware, and Phase#

### Logging Errors
1. Select the **Printer** from the dropdown
2. Navigate to the **Line #** where the error occurred
3. Click **Load** to populate test parameters
4. In the Add Error panel:
   - Select the **Error Type** from the dropdown
   - Check the affected **Printer(s)**
   - Click the **Tally** button for each affected page
5. Click **Add to Queue**

### Generating Cover Sheets
1. After adding errors to the queue, click **"Set Save Location..."** to choose an output folder
2. Click **"Save Covers"** to generate Excel cover sheets without printing
3. Or click **"Print Covers"** to generate and send directly to the printer

## Project Structure

```
DET LAB/
├── main.py                     # Application entry point
├── requirements.txt            # Python dependencies
├── README.md                   # This file
└── app/
    ├── __init__.py
    ├── constants.py            # Configuration constants
    ├── main_window.py          # Main application window
    ├── models/
    │   ├── __init__.py
    │   └── error_entry.py      # Error data model
    ├── services/
    │   ├── __init__.py
    │   └── cover_sheet_service.py  # Cover sheet generation & printing
    ├── utils/
    │   ├── __init__.py
    │   └── file_utils.py       # File utility functions
    └── widgets/
        ├── __init__.py
        ├── add_error_box.py    # Error input widget
        ├── placeholder_box.py  # Generic placeholder widget
        ├── queue_box.py        # Error queue widget
        ├── settings_dialog.py  # Session settings dialog
        └── test_info_box.py    # Test info display widget
```

## Configuration

### Printers (constants.py)
Default printer list can be modified in `app/constants.py`:
```python
PRINTERS = [
    "TH5BA5X058",
    "TH5BA5X059",
    # ... etc
]
```

### Excel Data Format
The tool expects Excel workbooks with:
- **One sheet per printer** (sheet names matching printer IDs)
- **Header row at row 9** (0-indexed row 8)
- **Page count in cell N6** of each sheet

Expected columns (after header row):
| Column Name | Maps To |
|-------------|---------|
| `Environment` | Climate |
| `Unit` | Unit |
| `Media_Number*` | Media # |
| `Suite_ID*` | Script # |
| `Print_Options*` | Plexity |
| `Load_NR*` | Load # |
| `Run_NR*` | Run # |
| `Load_Phase` | Phase |
| `Firmware` | Firmware |

### Cover Sheet Template
The cover sheet template should be an Excel file with the following cell mappings:

| Cell | Content |
|------|---------|
| A1   | KaMP # |
| I1   | Date |
| D4   | Operator |
| D5   | Climate |
| D6   | Event Name (Error Type) |
| D7   | Unit # |
| D8   | Media # |
| D9   | Script # |
| D10  | Plexity |
| D11  | Load |
| D12  | Run |
| D13  | Pages |
| D14  | Phase |
| D15  | Firmware |

## Printing

### Windows
Uses COM automation via `pywin32` to print through Microsoft Excel. Requires Excel to be installed.

### macOS
Uses AppleScript to automate printing through Microsoft Excel (preferred) or Numbers (fallback).

### Linux
Direct printing not supported. Use "Save Covers" to generate files, then print manually.

## Troubleshooting

### Slow Loading
If switching printers feels slow, ensure you're using the latest version which uses pandas (fast) instead of openpyxl (slow) for reading the page count.

### Print Not Working
- **Windows**: Ensure `pywin32` is installed: `pip install pywin32`
- **macOS**: Grant accessibility permissions to Terminal/IDE for AppleScript automation
- **All**: Verify Microsoft Excel is installed and the cover template path is valid

### Missing Data
- Verify your Excel workbook has sheets named exactly like the printer IDs
- Check that the header row is at row 9 (1-indexed)
- Ensure column names match the expected format (see Configuration section)

## Version History

- **v1.0.1** - Current version
  - Settings dialog for KaMP#, Operator, Firmware, Phase#
  - Page count display from cell N6
  - Save covers to custom directory
  - Optimized sheet loading performance

## License

Internal HP tool - Not for distribution.

## Author

DET Lab Team, Kelly St.Onge
