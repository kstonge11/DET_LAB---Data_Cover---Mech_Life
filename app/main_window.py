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
from .widgets import TestInfoBox, PlaceholderBox, AddErrorBox, QueueBox, SettingsDialog
from .services import CoverSheetService
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
        self.stored_phase = ""

        # Page count from sheet metadata (cell N6)
        self.page_count = ""

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
        """Create the compact settings display with edit button."""
        section = QVBoxLayout()
        section.setSpacing(4)
        section.setContentsMargins(0, 0, 10, 0)

        # Settings display frame
        settings_frame = QFrame()
        settings_frame.setStyleSheet("""
            QFrame {
                background-color: #3a3a3a;
                border: 1px solid #555;
                border-radius: 6px;
                padding: 4px;
            }
            QLabel#settingLabel {
                color: #888;
                font-size: 10px;
            }
            QLabel#settingValue {
                color: #4CAF50;
                font-size: 11px;
                font-weight: bold;
            }
        """)

        frame_layout = QVBoxLayout(settings_frame)
        frame_layout.setContentsMargins(8, 6, 8, 6)
        frame_layout.setSpacing(2)

        # Create compact display rows
        def make_setting_row(label_text):
            row = QHBoxLayout()
            row.setSpacing(4)

            label = QLabel(f"{label_text}:")
            label.setObjectName("settingLabel")
            label.setFixedWidth(55)

            value = QLabel("—")
            value.setObjectName("settingValue")
            value.setMinimumWidth(80)

            row.addWidget(label)
            row.addWidget(value)
            row.addStretch()

            return row, value

        kamp_row, self.kamp_display = make_setting_row("KaMP #")
        operator_row, self.operator_display = make_setting_row("Operator")
        firmware_row, self.firmware_display = make_setting_row("Firmware")
        phase_row, self.phase_display = make_setting_row("Phase #")

        frame_layout.addLayout(kamp_row)
        frame_layout.addLayout(operator_row)
        frame_layout.addLayout(firmware_row)
        frame_layout.addLayout(phase_row)

        section.addWidget(settings_frame)

        # Edit button
        self.edit_settings_btn = QPushButton("✏️ Edit Settings")
        self.edit_settings_btn.setFixedHeight(28)
        self.edit_settings_btn.setStyleSheet("""
            QPushButton {
                background-color: #4a4a4a;
                color: #fff;
                border: 1px solid #666;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #555;
            }
        """)
        section.addWidget(self.edit_settings_btn)

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

        '''# Printer box (placeholder for now)
        self.printer_box = PlaceholderBox("Printer", "🖨️")
        layout.addWidget(self.printer_box, stretch=1)'''

        # Add Error box
        self.error_box = AddErrorBox()
        layout.addWidget(self.error_box, stretch=2)

        # Queue box
        self.queue_box = QueueBox()
        layout.addWidget(self.queue_box, stretch=2)

        return content

    def _connect_signals(self):
        """Connect all signals to their handlers."""
        # File browsers
        self.browse_data_btn.clicked.connect(self._browse_data_file)
        self.open_data_btn.clicked.connect(lambda: open_in_default_app(self.data_file_path))
        self.browse_cover_btn.clicked.connect(self._browse_cover_file)
        self.open_cover_btn.clicked.connect(lambda: open_in_default_app(self.cover_file_path))

        # Settings edit button
        self.edit_settings_btn.clicked.connect(self._open_settings_dialog)

        # Test info box signals
        self.test_info_box.printer_changed.connect(self._on_printer_changed)
        self.test_info_box.line_changed.connect(self._on_line_changed)
        self.test_info_box.load_requested.connect(self._load_line_data)

        # Error box signals
        self.error_box.error_added.connect(self._on_error_added)

        # Queue box signals
        self.queue_box.choose_path_btn.clicked.connect(self._choose_save_path)
        self.queue_box.save_covers_requested.connect(self._save_cover_sheets)
        self.queue_box.print_requested.connect(self._print_cover_sheets)

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

            # Read page count from cell N6 using openpyxl
            self._load_page_count(selected_printer)

            self._load_line_data()
        except Exception as e:
            print(f"Error loading sheet: {e}")
            self.data_df = None

    def _load_page_count(self, sheet_name: str):
        """Load page count from cell N6 of the specified sheet using pandas (fast)."""
        try:
            # Read just row 6 (0-indexed row 5), no header, single row
            # Column N is index 13 (0-indexed)
            df_row = pd.read_excel(
                self.excel_file,
                sheet_name=sheet_name,
                header=None,
                skiprows=5,  # Skip rows 1-5 (0-indexed 0-4)
                nrows=1      # Read only 1 row (row 6)
            )

            # Column N = index 13
            if len(df_row.columns) > 13:
                page_count_value = df_row.iloc[0, 13]
                if pd.notna(page_count_value):
                    self.page_count = str(int(page_count_value) if isinstance(page_count_value, float) else page_count_value)
                    self.test_info_box.set_page_count(self.page_count)
                    print(f"Page count from N6: {self.page_count}")
                    return

            self.page_count = "—"
            self.test_info_box.set_page_count("—")
        except Exception as e:
            print(f"Error reading page count: {e}")
            self.page_count = "—"
            self.test_info_box.set_page_count("—")

    # ========== DATA LOADING ==========

    def _on_printer_changed(self, printer_name: str):
        print(f"Printer changed to: {printer_name}")
        if self.excel_file is not None:
            self._load_printer_sheet()

    def _on_line_changed(self, value: int):
        self._load_line_data()
        self._update_error_context()

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

    def _open_settings_dialog(self):
        """Open the settings dialog to edit session values."""
        dialog = SettingsDialog(
            self,
            kamp=self.stored_kamp,
            operator=self.stored_operator,
            firmware=self.stored_firmware,
            phase=self.stored_phase
        )

        if dialog.exec():
            values = dialog.get_values()
            self._apply_settings(values)

    def _apply_settings(self, values: dict):
        """Apply settings from the dialog."""
        # Update stored values
        self.stored_kamp = values["kamp"]
        self.stored_operator = values["operator"]
        self.stored_firmware = values["firmware"]
        self.stored_phase = values["phase"]

        # Update display labels
        self.kamp_display.setText(self.stored_kamp if self.stored_kamp else "—")
        self.operator_display.setText(self.stored_operator if self.stored_operator else "—")
        self.firmware_display.setText(self.stored_firmware if self.stored_firmware else "—")
        self.phase_display.setText(self.stored_phase if self.stored_phase else "—")

        # Update test info if loaded
        if self.data_df is not None:
            if self.stored_operator:
                self.test_info_box.set_field_value("Operator", self.stored_operator)

        # Update error context
        self._update_error_context()

        print(f"Settings updated: KaMP={self.stored_kamp}, Op={self.stored_operator}, FW={self.stored_firmware}, Phase={self.stored_phase}")

    # ========== ERROR HANDLING ==========

    def _update_error_context(self):
        """Update the error box with current context including test info."""
        # Get test info values from the test info box
        test_values = self.test_info_box.get_all_values()

        # Use stored_phase if set, otherwise fall back to test_values
        phase_value = self.stored_phase if self.stored_phase else test_values.get("Phase", "")

        self.error_box.set_context(
            kamp=self.stored_kamp,
            operator=self.stored_operator,
            firmware=self.stored_firmware,
            line_number=self.test_info_box.current_line(),
            # Test info fields - map display names to context keys
            climate=test_values.get("Climate", ""),
            media=test_values.get("Media #", ""),
            script=test_values.get("Script #", ""),
            plexity=test_values.get("Plexity", ""),
            load=test_values.get("Load #", ""),
            run=test_values.get("Run #", ""),
            phase=phase_value,
        )

    def _on_error_added(self, entry):
        """Handle when an error is added from the error box."""
        self.queue_box.add_entry(entry)

    def _choose_save_path(self):
        """Open dialog to choose save path for cover sheets."""
        path = QFileDialog.getExistingDirectory(
            self, 
            "Choose Save Location for Cover Sheets",
            "",
            QFileDialog.Option.ShowDirsOnly
        )
        if path:
            self.queue_box.set_save_path(path)
            print(f"Save path set to: {path}")

    def _save_cover_sheets(self, queue):
        """Generate and save cover sheets to the chosen path."""
        if not queue:
            print("No errors to save")
            return

        if not self.cover_file_path:
            print("No cover sheet template loaded")
            return

        save_path = self.queue_box.get_save_path()
        if not save_path:
            print("No save path selected")
            return

        # Create service and save files
        service = CoverSheetService(self.cover_file_path)

        total_saved = 0
        saved_files = []
        for entry in queue:
            files = service.save_from_entry(entry, save_path)
            saved_files.extend(files)
            total_saved += len(files)

        print(f"Saved {total_saved} cover sheets to: {save_path}")
        for f in saved_files:
            print(f"  - {f}")

    def _print_cover_sheets(self, queue):
        """Generate and print cover sheets for queued errors."""
        if not queue:
            print("No errors to print")
            return

        if not self.cover_file_path:
            print("No cover sheet template loaded")
            return

        # Create service with the template (files auto-deleted after printing)
        service = CoverSheetService(self.cover_file_path, keep_files=True)

        total_printed = 0
        for entry in queue:
            printed = service.print_from_entry(entry)
            total_printed += printed

        print(f"Printed {total_printed} cover sheets")
