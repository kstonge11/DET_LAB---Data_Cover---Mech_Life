# Cover Sheet / Data Insert Tool V.0.0.1
# Created By Kelly St.Onge March 24th 2026
# kelly.st-onge@hp.com
# kelly.stonge@us.beyondsoft.com
# kellywstonge@gmail.com
# ---Purpose---
# To help automate the flow of DET test labs cover sheet and data entry.
# ---Dependencies---
#   PyQt6, pandas, pyqtdarktheme

import pandas as pd
import sys
import os
import subprocess
import platform
from pathlib import Path
from PyQt6.QtWidgets import (
    QMainWindow, QApplication, QWidget,
    QLineEdit, QPushButton, QVBoxLayout, QHBoxLayout,
    QLabel, QFileDialog, QFrame, QSpacerItem, QSizePolicy
)
from PyQt6.QtCore import Qt
import qdarktheme


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        # ===APP WINDOW========
        self.setWindowTitle("DET Cover/Entry Tool V.0.0.1")
        self.resize(1200, 700)
        
        # Track current files
        self.data_file_path = None
        self.cover_file_path = None
        
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
        top_bar.setFixedHeight(100)
        
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
        
        # Kamp # row
        kamp_row = QHBoxLayout()
        kamp_row.setSpacing(6)
        kamp_label = QLabel("Kamp #:")
        kamp_label.setFixedWidth(55)
        self.kamp_input = QLineEdit()
        self.kamp_input.setPlaceholderText("Enter Kamp #")
        self.kamp_input.setFixedWidth(120)
        self.kamp_edit_btn = QPushButton("✏️ Set")
        self.kamp_edit_btn.setFixedSize(60, 26)
        self.kamp_display = QLabel("")
        self.kamp_display.setStyleSheet("color: #4CAF50; font-weight: bold;")
        
        kamp_row.addWidget(kamp_label)
        kamp_row.addWidget(self.kamp_input)
        kamp_row.addWidget(self.kamp_edit_btn)
        kamp_row.addWidget(self.kamp_display)
        
        # Operator row
        operator_row = QHBoxLayout()
        operator_row.setSpacing(6)
        operator_label = QLabel("Operator:")
        operator_label.setFixedWidth(55)
        self.operator_input = QLineEdit()
        self.operator_input.setPlaceholderText("Enter Operator")
        self.operator_input.setFixedWidth(120)
        self.operator_edit_btn = QPushButton("✏️ Set")
        self.operator_edit_btn.setFixedSize(60, 26)
        self.operator_display = QLabel("")
        self.operator_display.setStyleSheet("color: #4CAF50; font-weight: bold;")
        
        operator_row.addWidget(operator_label)
        operator_row.addWidget(self.operator_input)
        operator_row.addWidget(self.operator_edit_btn)
        operator_row.addWidget(self.operator_display)
        
        right_section.addLayout(kamp_row)
        right_section.addLayout(operator_row)
        top_bar_layout.addLayout(right_section)
        
        main_layout.addWidget(top_bar)
        
        # ====MAIN CONTENT AREA========
        content_area = QWidget()
        content_area.setObjectName("contentArea")
        content_layout = QVBoxLayout(content_area)
        content_layout.setContentsMargins(20, 20, 20, 20)
        
        # Placeholder for future content (data preview, etc.)
        content_layout.addStretch()
        
        main_layout.addWidget(content_area)
        
        # ====STORED VALUES====
        self.stored_kamp = ""
        self.stored_operator = ""

    def _connect_signals(self):
        """Connect all signals and slots."""
        # File browse/open buttons
        self.browse_data_btn.clicked.connect(self.browse_data_file)
        self.open_data_btn.clicked.connect(lambda: self.open_in_excel(self.data_file_path))
        self.browse_cover_btn.clicked.connect(self.browse_cover_file)
        self.open_cover_btn.clicked.connect(lambda: self.open_in_excel(self.cover_file_path))
        
        # Kamp/Operator set buttons
        self.kamp_edit_btn.clicked.connect(self.set_kamp)
        self.operator_edit_btn.clicked.connect(self.set_operator)
        
        # Allow Enter key to set values
        self.kamp_input.returnPressed.connect(self.set_kamp)
        self.operator_input.returnPressed.connect(self.set_operator)

    def browse_data_file(self):
        """Open file dialog to select Data Excel file."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Data File",
            "",
            "Excel Files (*.xlsx *.xls *.xlsm);;CSV Files (*.csv);;All Files (*)"
        )
        if file_path:
            self.data_file_path = file_path
            # Show truncated path if too long
            display_path = self._truncate_path(file_path, 50)
            self.data_path_label.setText(display_path)
            self.data_path_label.setToolTip(file_path)  # Full path on hover
            self.open_data_btn.setEnabled(True)

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
