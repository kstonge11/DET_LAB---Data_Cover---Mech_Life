"""
Constants and configuration for DET Cover/Entry Tool
"""

APP_NAME = "DET Cover/Entry Tool"
APP_VERSION = "0.0.8"
WINDOW_TITLE = f"{APP_NAME} V.{APP_VERSION}"

# Default printer list
PRINTERS = [
    "TH5BA5X058",
    "TH5BA5X059",
    "TH5BA5X068",
    "TH5BA5X069",
    "TH5BA5X098",
    "TH5BA5X099",
    "TH5BA5X100",
    "TH5BA5X101",
]

# Short names for printer checkboxes
PRINTER_SHORT_NAMES = ["P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8"]

# Excel data row where headers are located (0-indexed)
EXCEL_HEADER_ROW = 8

# Field definitions for Test Info box
# Format: (display_name, [possible_column_names])
TEST_INFO_FIELDS = {
    "Test Setup": [
        ("Operator", ["_STORED_OPERATOR_"]),
        ("Climate", ["Environment"]),
        ("Unit", ["Unit"]),
        ("Media #", ["Media_Number*"]),
        ("Script #", ["Suite_ID*"]),
        ("Plexity", ["Print_Options*"]),
        ("Load #", ["Load_NR*"]),
        ("Run #", ["Run_NR*"]),
        ("Phase", ["Load_Phase"]),
        ("Firmware", ["Firmware"]),
    ]
}

# Error types from Excel columns
ERROR_TYPES = [
    # Pick Issues
    "Pick Skew",
    "Carraige Induced Skew",
    "No Pick Clearable",
    "No Pick",
    "Partial Pick",
    "Trailing Pick",
    "Failure To Advance",
    "Multipick",
    "Multipick - over 5 sheets",
    "No Pick, Last Sheet",
    # Print Quality
    "Back of Page Rib Smear",
    "Massive Missing Nozzles",
    "Incomplete Print",
    "Vertical Smear",
    "Ink Redeposit",
    "Blurred Print",
    "Carriage Smear",
    "Horizontal Banding",
    "Ink Drops on Page",
    "Margin Shift Up",
    "Margin Shift Right",
    "Margin Shift Down",
    "Margin Shift Left",
    "TOF Misprint",
    # Stacking/Output
    "Stacking Error - Bulldoze",
    "Stacking Error - Rollover",
    # Jams - Duplex
    "Duplex Output Suckback Jam",
    "Duplex Output Suckback Jam - Damage",
    "Duplex Output Jam",
    "Duplex Output Jam - Damage",
    "Duplex Printzone Jam",
    "Duplex Printzone Jam - Damage",
    "Duplex Input Jam",
    "Duplex Input Jam - Damage",
    "Duplex Turnroller Jam",
    "Duplex Turnroller Jam - Damage",
    "Duplex on One Side",
    # Jams - Simplex
    "Simplex Input Jam",
    "Simplex Input Jam - Damage",
    "Simplex Output Jam",
    "Simplex Output Jam - Damage",
    "Simplex Output Suckback Jam",
    "Simplex Output Suckback Jam - Damage",
    "Simplex Printzone Jam",
    "Simplex Printzone Jam - Damage",
    "Simplex Turn Jam",
    "Simplex Turn Zone Jam - Damage",
    # Media Damage
    "Edge Damage",
    "Dog Ear",
    "Creased Media",
    "Media Corruption",
    # Unit Errors
    "Media Size Mismatch",
    "Page Stutter",
    "Paper Stalled While Printing",
    "Unit Stops With Error",
    "Unit Stops without Error",
    "Unit Ejects Blank Page",
    "Pen Failure",
    "Intermittent Pen Failure",
    "Unit Powers Itself Off",
]
