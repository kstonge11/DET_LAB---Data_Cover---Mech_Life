"""
Constants and configuration for DET Cover/Entry Tool
"""

APP_NAME = "DET Cover/Entry Tool"
APP_VERSION = "0.0.7"
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
