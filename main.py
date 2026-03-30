#!/usr/bin/env python3
"""
Cover Sheet / Data Insert Tool V.1.0.2

Created By Kelly St.Onge March 24th 2026
kelly.st-onge@hp.com
kelly.stonge@us.beyondsoft.com
kellywstonge@gmail.com

Purpose:
    To help automate the flow of DET test labs cover sheet and data entry.

Dependencies:
    PyQt6, pandas, pyqtdarktheme
"""

import sys
from PyQt6.QtWidgets import QApplication
import qdarktheme

from app import MainWindow


def main():
    app = QApplication(sys.argv)
    qdarktheme.setup_theme()

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
