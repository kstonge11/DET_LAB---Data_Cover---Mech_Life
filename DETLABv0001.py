'''
Cover Sheet / Data Insert Tool V.0.0.7

    Created By Kelly St.Onge March 24th 2026
    kelly.st-onge@hp.com
    kelly.stonge@us.beyondsoft.com
    kellywstonge@gmail.com
    ---Purpose---
    To help automate the flow of DET test labs cover sheet and data entry.
    ---Dependencies---
    PyQt6, pandas, pyqtdarktheme

'''

import pandas as pd
import sys
import os
import subprocess
import platform
from pathlib import Path
from PyQt6.QtWidgets import (
    QMainWindow, QApplication, QWidget,
    QLineEdit, QPushButton, QVBoxLayout, QHBoxLayout,
    QLabel, QFileDialog, QFrame, QSpacerItem, QSizePolicy,
    QSpinBox, QGroupBox, QGridLayout, QScrollArea, QComboBox
)
from PyQt6.QtCore import Qt
import qdarktheme


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        # ===APP WINDOW========
        self.setWindowTitle("DET Cover/Entry Tool V.0.0.7")
        self.resize(1200, 700)

        # Track current files
        self.data_file_path = None
        self.cover_file_path = None
        self.data_df = None  # Current DataFrame
        self.excel_file = None  # Excel file object for multi-sheet access
        self.data_df = None  # Pandas DataFrame for loaded data

        self._setup_ui()
        self._connect_signals()

    def _setup_ui(self):

        # ====MAIN CONTAINER========
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # ====TOP BAR========
        top_bar = QFrame()
        top_bar.setObjectName("topBar")
        top_bar.setStyleSheet("""
            #topBar {
                background-color: #2a2a2a;
                border-bottom: 1px solid #444;
            }
            QLabel#pathLabel {
                color: #888;
                font-size: 11px;
            }
        """)

        top_bar.setFixedHeight(120)  # Taller for 3 rows
        top_bar_layout = QHBoxLayout(top_bar)
        top_bar_layout.setContentsMargins(12, 8, 12, 8)

        # ====LEFT SIDE - FILE BROWSE BUTTONS====
        left_section = QVBoxLayout()
        left_section.setSpacing(6)

        # --- Data File Row ---
        data_row = QHBoxLayout()
        data_row.setSpacing(8)

        self.browse_data_btn = QPushButton("📂 Browse Data")
        self.browse_data_btn.setFixedSize(130, 28)

        self.open_data_btn = QPushButton("📊 Open in Excel")
        self.open_data_btn.setFixedSize(130, 28)
        self.open_data_btn.setEnabled(False)

        self.data_path_label = QLabel("No data file selected")
        self.data_path_label.setObjectName("pathLabel")
        self.data_path_label.setMinimumWidth(300)

        data_row.addWidget(self.browse_data_btn)
        data_row.addWidget(self.open_data_btn)
        data_row.addWidget(self.data_path_label)
        data_row.addStretch()

        # --- Cover Sheet Row ---
        cover_row = QHBoxLayout()
        cover_row.setSpacing(8)

        self.browse_cover_btn = QPushButton("📂 Browse Cover")
        self.browse_cover_btn.setFixedSize(130, 28)

        self.open_cover_btn = QPushButton("📊 Open in Excel")
        self.open_cover_btn.setFixedSize(130, 28)
        self.open_cover_btn.setEnabled(False)

        self.cover_path_label = QLabel("No cover sheet selected")
        self.cover_path_label.setObjectName("pathLabel")
        self.cover_path_label.setMinimumWidth(300)

        cover_row.addWidget(self.browse_cover_btn)
        cover_row.addWidget(self.open_cover_btn)
        cover_row.addWidget(self.cover_path_label)
        cover_row.addStretch()

        left_section.addLayout(data_row)
        left_section.addLayout(cover_row)
        top_bar_layout.addLayout(left_section)

        # Expanding spacer between left and right
        top_bar_layout.addSpacerItem(
            QSpacerItem(40, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)
        )

        # ====RIGHT SIDE - KAMP # AND OPERATOR WITH EDIT BUTTONS====
        right_section = QVBoxLayout()
        right_section.setSpacing(6)
        right_section.setContentsMargins(0, 0, 10, 0)  # Add right padding

        # Kamp # row
        kamp_row = QHBoxLayout()
        kamp_row.setSpacing(6)
        kamp_label = QLabel("KaMP #:")
        kamp_label.setFixedWidth(70)
        self.kamp_input = QLineEdit()
        self.kamp_input.setPlaceholderText("Enter Kamp #")
        self.kamp_input.setFixedWidth(140)
        self.kamp_edit_btn = QPushButton("✏️ Set")
        self.kamp_edit_btn.setFixedSize(65, 26)
        self.kamp_display = QLabel("")
        self.kamp_display.setStyleSheet("color: #4CAF50; font-weight: bold;")
        self.kamp_display.setMinimumWidth(80)

        kamp_row.addWidget(kamp_label)
        kamp_row.addWidget(self.kamp_input)
        kamp_row.addWidget(self.kamp_edit_btn)
        kamp_row.addWidget(self.kamp_display)

        # Operator row
        operator_row = QHBoxLayout()
        operator_row.setSpacing(6)
        operator_label = QLabel("Operator:")
        operator_label.setFixedWidth(70)
        self.operator_input = QLineEdit()
        self.operator_input.setPlaceholderText("Enter Operator")
        self.operator_input.setFixedWidth(140)
        self.operator_edit_btn = QPushButton("✏️ Set")
        self.operator_edit_btn.setFixedSize(65, 26)
        self.operator_display = QLabel("")
        self.operator_display.setStyleSheet("color: #4CAF50; font-weight: bold;")
        self.operator_display.setMinimumWidth(80)

        operator_row.addWidget(operator_label)
        operator_row.addWidget(self.operator_input)
        operator_row.addWidget(self.operator_edit_btn)
        operator_row.addWidget(self.operator_display)

        # Firmware row
        firmware_row = QHBoxLayout()
        firmware_row.setSpacing(6)
        firmware_label = QLabel("Firmware:")
        firmware_label.setFixedWidth(70)
        self.firmware_input = QLineEdit()
        self.firmware_input.setPlaceholderText("Enter Firmware")
        self.firmware_input.setFixedWidth(140)
        self.firmware_edit_btn = QPushButton("✏️ Set")
        self.firmware_edit_btn.setFixedSize(65, 26)
        self.firmware_display = QLabel("")
        self.firmware_display.setStyleSheet("color: #4CAF50; font-weight: bold;")
        self.firmware_display.setMinimumWidth(80)

        firmware_row.addWidget(firmware_label)
        firmware_row.addWidget(self.firmware_input)
        firmware_row.addWidget(self.firmware_edit_btn)
        firmware_row.addWidget(self.firmware_display)

        right_section.addLayout(kamp_row)
        right_section.addLayout(operator_row)
        right_section.addLayout(firmware_row)
        top_bar_layout.addLayout(right_section)

        main_layout.addWidget(top_bar)

        # ====MAIN CONTENT AREA========
        content_area = QWidget()
        content_area.setObjectName("contentArea")
        content_layout = QHBoxLayout(content_area)
        content_layout.setContentsMargins(15, 15, 15, 15)
        content_layout.setSpacing(15)

        # --- BOX 1: Test Info ---
        self.test_info_box = self._create_test_info_box()
        content_layout.addWidget(self.test_info_box, stretch=2)

        # --- BOX 2: Printer ---
        self.printer_box = self._create_placeholder_box("Printer", "🖨️")
        content_layout.addWidget(self.printer_box, stretch=1)

        # --- BOX 3: Add Error ---
        self.error_box = self._create_placeholder_box("Add Error", "⚠️")
        content_layout.addWidget(self.error_box, stretch=1)

        # --- BOX 4: Send ---
        self.send_box = self._create_placeholder_box("Send", "📤")
        content_layout.addWidget(self.send_box, stretch=1)

        main_layout.addWidget(content_area)

        # ====STORED VALUES====
        self.stored_kamp = ""
        self.stored_operator = ""
        self.stored_firmware = ""

    def _create_test_info_box(self):
        """Create the Test Info box with Printer selector, Line # selector and data fields."""
        box = QFrame()
        box.setObjectName("testInfoBox")
        box.setStyleSheet("""
            #testInfoBox {
                background-color: #333;
                border: 1px solid #555;
                border-radius: 8px;
            }
            QLabel#boxTitle {
                font-size: 14px;
                font-weight: bold;
                color: #fff;
            }
            QLabel#sectionTitle {
                font-size: 11px;
                font-weight: bold;
                color: #888;
                padding-top: 6px;
            }
            QLabel#fieldLabel {
                color: #aaa;
                font-size: 10px;
            }
            QLabel#fieldValue {
                color: #4CAF50;
                font-size: 11px;
                font-weight: bold;
            }
            QComboBox {
                min-width: 120px;
            }
        """)

        layout = QVBoxLayout(box)
        layout.setContentsMargins(12, 10, 12, 12)
        layout.setSpacing(8)

        # Title
        title = QLabel("📋 Test Info")
        title.setObjectName("boxTitle")
        layout.addWidget(title)

        # Printer & Line selector row
        selector_row = QHBoxLayout()
        selector_row.setSpacing(15)

        # Printer selector
        printer_section = QHBoxLayout()
        printer_section.setSpacing(6)
        printer_label = QLabel("Printer:")
        printer_label.setStyleSheet("color: #fff;")

        self.printer_combo = QComboBox()
        self.printer_combo.addItems([
            "TH5BA5X058",
            "TH5BA5X059",
            "TH5BA5X068",
            "TH5BA5X069",
            "TH5BA5X098",
            "TH5BA5X099",
            "TH5BA5X100",
            "TH5BA5X101"
        ])
        self.printer_combo.setFixedWidth(130)

        printer_section.addWidget(printer_label)
        printer_section.addWidget(self.printer_combo)
        selector_row.addLayout(printer_section)

        # Separator line
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.VLine)
        sep.setStyleSheet("background-color: #555;")
        sep.setFixedWidth(1)
        selector_row.addWidget(sep)

        # Line # selector
        line_section = QHBoxLayout()
        line_section.setSpacing(6)

        line_label = QLabel("Line #:")
        line_label.setStyleSheet("color: #fff;")

        #self.line_down_btn = QPushButton("◀")
        #self.line_down_btn.setFixedSize(30, 26)

        self.line_spinbox = QSpinBox()
        self.line_spinbox.setMinimum(1)
        self.line_spinbox.setMaximum(9999)
        self.line_spinbox.setValue(1)
        self.line_spinbox.setFixedWidth(70)

        #self.line_up_btn = QPushButton("▶")
        #self.line_up_btn.setFixedSize(30, 26)

        self.load_line_btn = QPushButton("Load")
        self.load_line_btn.setFixedSize(60, 26)

        line_section.addWidget(line_label)
        #line_section.addWidget(self.line_down_btn)
        line_section.addWidget(self.line_spinbox)
        #line_section.addWidget(self.line_up_btn)
        line_section.addWidget(self.load_line_btn)

        selector_row.addLayout(line_section)
        selector_row.addStretch()

        layout.addLayout(selector_row)

        # Separator
        separator = QFrame()
        separator.setFrameShape(QFrame.Shape.HLine)
        separator.setStyleSheet("background-color: #555;")
        separator.setFixedHeight(1)
        layout.addWidget(separator)

        # Scrollable content area for all the fields
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("""
            QScrollArea { 
                border: none; 
                background-color: transparent;
            }
            QScrollArea > QWidget > QWidget { 
                background-color: transparent;
            }
        """)

        scroll_content = QWidget()
        fields_layout = QVBoxLayout(scroll_content)
        fields_layout.setContentsMargins(0, 0, 0, 0)
        fields_layout.setSpacing(4)

        # Store all field value labels
        self.test_info_values = {}

        # Define field groups with only the requested fields
        field_groups = {
            "Test Setup": [
                ("Operator", ["_STORED_OPERATOR_"]),  # Special: from stored value
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

        for group_name, fields in field_groups.items():
            # Section header
            section_label = QLabel(group_name)
            section_label.setObjectName("sectionTitle")
            fields_layout.addWidget(section_label)

            # Single column grid for fields
            grid = QGridLayout()
            grid.setSpacing(4)
            grid.setColumnMinimumWidth(0, 70)
            grid.setColumnStretch(1, 1)  # Value column stretches

            for i, (display_name, column_names) in enumerate(fields):
                label = QLabel(f"{display_name}:")
                label.setObjectName("fieldLabel")

                value = QLabel("—")
                value.setObjectName("fieldValue")

                grid.addWidget(label, i, 0)
                grid.addWidget(value, i, 1)

                # Store with display name as key
                self.test_info_values[display_name] = {
                    "label": value,
                    "columns": column_names
                }

            fields_layout.addLayout(grid)

        fields_layout.addStretch()
        scroll.setWidget(scroll_content)
        layout.addWidget(scroll)

        return box

    def _create_placeholder_box(self, title, icon):
        """Create a placeholder box for future functionality."""
        box = QFrame()
        box.setObjectName("placeholderBox")
        box.setStyleSheet("""
            #placeholderBox {
                background-color: #333;
                border: 1px solid #555;
                border-radius: 8px;
            }
            QLabel#boxTitle {
                font-size: 14px;
                font-weight: bold;
                color: #fff;
            }
            QLabel#placeholder {
                color: #666;
                font-size: 12px;
            }
        """)

        layout = QVBoxLayout(box)
        layout.setContentsMargins(15, 12, 15, 15)
        layout.setSpacing(10)

        # Title
        title_label = QLabel(f"{icon} {title}")
        title_label.setObjectName("boxTitle")
        layout.addWidget(title_label)

        # Placeholder text
        placeholder = QLabel("Coming soon...")
        placeholder.setObjectName("placeholder")
        placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(placeholder)

        layout.addStretch()

        return box

    def _connect_signals(self):
        """Connect all signals and slots."""
        # File browse/open buttons
        self.browse_data_btn.clicked.connect(self.browse_data_file)
        self.open_data_btn.clicked.connect(lambda: self.open_in_excel(self.data_file_path))
        self.browse_cover_btn.clicked.connect(self.browse_cover_file)
        self.open_cover_btn.clicked.connect(lambda: self.open_in_excel(self.cover_file_path))

        # Kamp/Operator/Firmware set buttons
        self.kamp_edit_btn.clicked.connect(self.set_kamp)
        self.operator_edit_btn.clicked.connect(self.set_operator)
        self.firmware_edit_btn.clicked.connect(self.set_firmware)

        # Allow Enter key to set values
        self.kamp_input.returnPressed.connect(self.set_kamp)
        self.operator_input.returnPressed.connect(self.set_operator)
        self.firmware_input.returnPressed.connect(self.set_firmware)

        # Printer selector
        self.printer_combo.currentTextChanged.connect(self.on_printer_changed)

        # Line # controls
        #self.line_up_btn.clicked.connect(lambda: self.line_spinbox.setValue(self.line_spinbox.value() + 1))
        #self.line_down_btn.clicked.connect(lambda: self.line_spinbox.setValue(max(1, self.line_spinbox.value() - 1)))
        self.load_line_btn.clicked.connect(self.load_line_data)
        self.line_spinbox.valueChanged.connect(self.on_line_changed)

    def browse_data_file(self):
        """Open file dialog to select Data Excel file."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Data File",
            "",
            "Excel Files (*.xlsx *.xls *.xlsm);;CSV Files (*.csv);;All Files (*)")

        if file_path:
            self.data_file_path = file_path
            # Show truncated path if too long
            display_path = self._truncate_path(file_path, 50)
            self.data_path_label.setText(display_path)
            self.data_path_label.setToolTip(file_path)  # Full path on hover
            self.open_data_btn.setEnabled(True)

            # Load the Excel file
            self._load_data_file(file_path)

    def _load_data_file(self, file_path):
        """Load Excel file and store reference for sheet access."""
        try:
            if file_path.endswith('.csv'):
                self.data_df = pd.read_csv(file_path)
                self.excel_file = None
            else:
                # Store the Excel file object for multi-sheet access
                self.excel_file = pd.ExcelFile(file_path)
                print(f"Available sheets: {self.excel_file.sheet_names}")

                # Load the currently selected printer's sheet
                self._load_printer_sheet()

        except Exception as e:
            print(f"Error loading file: {e}")
            self.data_df = None
            self.excel_file = None

    def _load_printer_sheet(self):
        """Load the data from the currently selected printer's sheet."""
        if self.excel_file is None:
            return

        selected_printer = self.printer_combo.currentText()

        # Check if the sheet exists
        if selected_printer not in self.excel_file.sheet_names:
            print(f"Sheet '{selected_printer}' not found in Excel file")
            print(f"Available sheets: {self.excel_file.sheet_names}")
            self.data_df = None
            self._clear_test_info()
            return

        try:
            # Read the sheet with header on row 8 (0-indexed)
            self.data_df = pd.read_excel(
                self.excel_file, 
                sheet_name=selected_printer,
                header=8  # Row 9 in Excel (0-indexed = 8)
            )

            # Update max line number based on data rows
            max_rows = len(self.data_df)
            self.line_spinbox.setMaximum(max_rows)
            self.line_spinbox.setValue(1)

            print(f"Loaded {len(self.data_df)} rows from sheet '{selected_printer}'")
            print(f"Columns: {list(self.data_df.columns)[:10]}...")  # First 10 columns

            # Auto-load first data row
            self.load_line_data()

        except Exception as e:
            print(f"Error loading sheet '{selected_printer}': {e}")
            self.data_df = None

    def on_printer_changed(self, printer_name):
        """Called when printer selection changes."""
        print(f"Printer changed to: {printer_name}")
        if self.excel_file is not None:
            self._load_printer_sheet()

    def load_line_data(self):
        """Load data from the selected line number."""
        if self.data_df is None:
            print("No data file loaded")
            return

        # Line 1 = index 0 (data starts at row 1)
        row_index = self.line_spinbox.value() - 1

        if row_index < 0 or row_index >= len(self.data_df):
            print(f"Line {self.line_spinbox.value()} out of range")
            self._clear_test_info()
            return

        row = self.data_df.iloc[row_index]

        # Update each field
        for display_name, field_info in self.test_info_values.items():
            value_label = field_info["label"]
            possible_columns = field_info["columns"]

            # Special handling for stored operator
            if "_STORED_OPERATOR_" in possible_columns:
                value = self.stored_operator if self.stored_operator else "—"
            else:
                value = "—"
                for col in possible_columns:
                    if col in row.index:
                        val = row[col]
                        # Handle NaN values
                        if pd.notna(val):
                            value = str(val)
                        break
            value_label.setText(value)

        print(f"Loaded line {self.line_spinbox.value()} (Id: {row.get('Id', 'N/A')})")

    def _clear_test_info(self):
        """Clear all test info fields."""
        for field_info in self.test_info_values.values():
            field_info["label"].setText("—")

    def on_line_changed(self, value):
        """Called when line spinbox value changes."""
        # Optional: auto-load on change (uncomment if desired)
        self.load_line_data()
        pass

    def browse_cover_file(self):
        """Open file dialog to select Cover Sheet Excel file."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Cover Sheet",
            "",
            "Excel Files (*.xlsx *.xls *.xlsm);;CSV Files (*.csv);;All Files (*)"
        )
        if file_path:
            self.cover_file_path = file_path
            # Show truncated path if too long
            display_path = self._truncate_path(file_path, 50)
            self.cover_path_label.setText(display_path)
            self.cover_path_label.setToolTip(file_path)  # Full path on hover
            self.open_cover_btn.setEnabled(True)

    def _truncate_path(self, path, max_length):
        """Truncate path for display, keeping filename visible."""
        if len(path) <= max_length:
            return path
        filename = os.path.basename(path)
        if len(filename) >= max_length - 5:
            return "..." + filename[-(max_length - 3):]
        remaining = max_length - len(filename) - 4  # 4 for ".../"
        return path[:remaining] + ".../" + filename

    def set_kamp(self):
        """Store the Kamp # value."""
        value = self.kamp_input.text().strip()
        if value:
            self.stored_kamp = value
            self.kamp_display.setText(f"✓ {value}")
            print(f"Kamp # set to: {self.stored_kamp}")

    def set_operator(self):
        """Store the Operator value."""
        value = self.operator_input.text().strip()
        if value:
            self.stored_operator = value
            self.operator_display.setText(f"✓ {value}")
            print(f"Operator set to: {self.stored_operator}")
            # Update the Test Info display if data is loaded
            if self.data_df is not None and "Operator" in self.test_info_values:
                self.test_info_values["Operator"]["label"].setText(value)

    def set_firmware(self):
        """Store the Firmware value."""
        value = self.firmware_input.text().strip()
        if value:
            self.stored_firmware = value
            self.firmware_display.setText(f"✓ {value}")
            print(f"Firmware set to: {self.stored_firmware}")

    def open_in_excel(self, file_path=None):
        """Open the specified file in the system's default Excel application."""
        if not file_path or not os.path.exists(file_path):
            return

        try:
            if platform.system() == "Windows":
                os.startfile(file_path)
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", file_path])
            else:  # Linux
                subprocess.run(["xdg-open", file_path])
        except Exception as e:
            print(f"Error opening file: {e}")


def main():
    app = QApplication(sys.argv)
    qdarktheme.setup_theme()
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
