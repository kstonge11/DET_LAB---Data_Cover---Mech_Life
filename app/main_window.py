"""
Main application window for DET Cover/Entry Tool
"""

import pandas as pd
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLineEdit, QPushButton, QLabel, QFileDialog,
    QFrame, QSpacerItem, QSizePolicy
)

from .constants import WINDOW_TITLE, EXCEL_HEADER_ROW
from .widgets import TestInfoBox, PlaceholderBox
from .utils import truncate_path, open_in_default_app


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(WINDOW_TITLE)
        self.resize(1200, 700)

        # File tracking
        self.data_file_path = None
        self.cover_file_path = None
        self.excel_file = None
        self.data_df = None

        # Stored values
        self.stored_kamp = ""
        self.stored_operator = ""
        self.stored_firmware = ""

        self._setup_ui()
        self._connect_signals()

    def _setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # Top bar
        main_layout.addWidget(self._create_top_bar())

        # Content area
        main_layout.addWidget(self._create_content_area())

    def _create_top_bar(self) -> QFrame:
        """Create the top bar with file browsers and settings inputs."""
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
        top_bar.setFixedHeight(120)

        layout = QHBoxLayout(top_bar)
        layout.setContentsMargins(12, 8, 12, 8)

        # Left: File browsers
        layout.addLayout(self._create_file_browser_section())

        # Spacer
        layout.addSpacerItem(
            QSpacerItem(40, 20, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Minimum)
        )

        # Right: Settings inputs
        layout.addLayout(self._create_settings_section())

        return top_bar

    def _create_file_browser_section(self) -> QVBoxLayout:
        """Create the file browser buttons and labels."""
        section = QVBoxLayout()
        section.setSpacing(6)

        # Data file row
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
        section.addLayout(data_row)

        # Cover sheet row
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
        section.addLayout(cover_row)

        return section

    def _create_settings_section(self) -> QVBoxLayout:
        """Create the Kamp/Operator/Firmware input section."""
        section = QVBoxLayout()
        section.setSpacing(6)
        section.setContentsMargins(0, 0, 10, 0)

        # Helper to create each row
        def make_row(label_text, placeholder):
            row = QHBoxLayout()
            row.setSpacing(6)

            label = QLabel(f"{label_text}:")
            label.setFixedWidth(70)

            input_field = QLineEdit()
            input_field.setPlaceholderText(placeholder)
            input_field.setFixedWidth(140)

            btn = QPushButton("✏️ Set")
            btn.setFixedSize(65, 26)

            display = QLabel("")
            display.setStyleSheet("color: #4CAF50; font-weight: bold;")
            display.setMinimumWidth(80)

            row.addWidget(label)
            row.addWidget(input_field)
            row.addWidget(btn)
            row.addWidget(display)

            return row, input_field, btn, display

        # Create rows
        kamp_row, self.kamp_input, self.kamp_edit_btn, self.kamp_display = \
            make_row("KaMP #", "Enter Kamp #")
        operator_row, self.operator_input, self.operator_edit_btn, self.operator_display = \
            make_row("Operator", "Enter Operator")
        firmware_row, self.firmware_input, self.firmware_edit_btn, self.firmware_display = \
            make_row("Firmware", "Enter Firmware")

        section.addLayout(kamp_row)
        section.addLayout(operator_row)
        section.addLayout(firmware_row)

        return section

    def _create_content_area(self) -> QWidget:
        """Create the main content area with the four boxes."""
        content = QWidget()
        content.setObjectName("contentArea")
        layout = QHBoxLayout(content)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(15)

        # Test Info box (custom widget)
        self.test_info_box = TestInfoBox()
        layout.addWidget(self.test_info_box, stretch=2)

        # Placeholder boxes
        self.printer_box = PlaceholderBox("Printer", "🖨️")
        layout.addWidget(self.printer_box, stretch=1)

        self.error_box = PlaceholderBox("Add Error", "⚠️")
        layout.addWidget(self.error_box, stretch=1)

        self.send_box = PlaceholderBox("Send", "📤")
        layout.addWidget(self.send_box, stretch=1)

        return content

    def _connect_signals(self):
        """Connect all signals to their handlers."""
        # File browsers
        self.browse_data_btn.clicked.connect(self._browse_data_file)
        self.open_data_btn.clicked.connect(lambda: open_in_default_app(self.data_file_path))
        self.browse_cover_btn.clicked.connect(self._browse_cover_file)
        self.open_cover_btn.clicked.connect(lambda: open_in_default_app(self.cover_file_path))

        # Settings inputs
        self.kamp_edit_btn.clicked.connect(self._set_kamp)
        self.operator_edit_btn.clicked.connect(self._set_operator)
        self.firmware_edit_btn.clicked.connect(self._set_firmware)
        self.kamp_input.returnPressed.connect(self._set_kamp)
        self.operator_input.returnPressed.connect(self._set_operator)
        self.firmware_input.returnPressed.connect(self._set_firmware)

        # Test info box signals
        self.test_info_box.printer_changed.connect(self._on_printer_changed)
        self.test_info_box.line_changed.connect(self._on_line_changed)
        self.test_info_box.load_requested.connect(self._load_line_data)

    # ========== FILE HANDLING ==========

    def _browse_data_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Data File", "",
            "Excel Files (*.xlsx *.xls *.xlsm);;CSV Files (*.csv);;All Files (*)"
        )
        if file_path:
            self.data_file_path = file_path
            self.data_path_label.setText(truncate_path(file_path, 50))
            self.data_path_label.setToolTip(file_path)
            self.open_data_btn.setEnabled(True)
            self._load_data_file(file_path)

    def _browse_cover_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Cover Sheet", "",
            "Excel Files (*.xlsx *.xls *.xlsm);;CSV Files (*.csv);;All Files (*)"
        )
        if file_path:
            self.cover_file_path = file_path
            self.cover_path_label.setText(truncate_path(file_path, 50))
            self.cover_path_label.setToolTip(file_path)
            self.open_cover_btn.setEnabled(True)

    def _load_data_file(self, file_path: str):
        try:
            if file_path.endswith('.csv'):
                self.data_df = pd.read_csv(file_path)
                self.excel_file = None
            else:
                self.excel_file = pd.ExcelFile(file_path)
                print(f"Available sheets: {self.excel_file.sheet_names}")
                self._load_printer_sheet()
        except Exception as e:
            print(f"Error loading file: {e}")
            self.data_df = None
            self.excel_file = None

    def _load_printer_sheet(self):
        if self.excel_file is None:
            return

        selected_printer = self.test_info_box.current_printer()

        if selected_printer not in self.excel_file.sheet_names:
            print(f"Sheet '{selected_printer}' not found")
            self.data_df = None
            self.test_info_box.clear_all_fields()
            return

        try:
            self.data_df = pd.read_excel(
                self.excel_file,
                sheet_name=selected_printer,
                header=EXCEL_HEADER_ROW
            )
            self.test_info_box.set_max_lines(len(self.data_df))
            self.test_info_box.set_line(1)
            print(f"Loaded {len(self.data_df)} rows from '{selected_printer}'")
            self._load_line_data()
        except Exception as e:
            print(f"Error loading sheet: {e}")
            self.data_df = None

    # ========== DATA LOADING ==========

    def _on_printer_changed(self, printer_name: str):
        print(f"Printer changed to: {printer_name}")
        if self.excel_file is not None:
            self._load_printer_sheet()

    def _on_line_changed(self, value: int):
        self._load_line_data()

    def _load_line_data(self):
        if self.data_df is None:
            return

        row_index = self.test_info_box.current_line() - 1
        if row_index < 0 or row_index >= len(self.data_df):
            self.test_info_box.clear_all_fields()
            return

        row = self.data_df.iloc[row_index]

        for field_name in self.test_info_box.get_all_field_names():
            columns = self.test_info_box.get_field_columns(field_name)

            if "_STORED_OPERATOR_" in columns:
                value = self.stored_operator if self.stored_operator else "—"
            else:
                value = "—"
                for col in columns:
                    if col in row.index and pd.notna(row[col]):
                        value = str(row[col])
                        break

            self.test_info_box.set_field_value(field_name, value)

    # ========== SETTINGS ==========

    def _set_kamp(self):
        value = self.kamp_input.text().strip()
        if value:
            self.stored_kamp = value
            self.kamp_display.setText(f"✓ {value}")

    def _set_operator(self):
        value = self.operator_input.text().strip()
        if value:
            self.stored_operator = value
            self.operator_display.setText(f"✓ {value}")
            # Update test info if loaded
            if self.data_df is not None:
                self.test_info_box.set_field_value("Operator", value)

    def _set_firmware(self):
        value = self.firmware_input.text().strip()
        if value:
            self.stored_firmware = value
            self.firmware_display.setText(f"✓ {value}")
