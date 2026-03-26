'''
Cover Sheet / Data Insert Tool V.0.0.1
Created By Kelly St.Onge March 24th 2026
kelly.st-onge@hp.com
kelly.stonge@us.beyondsoft.com
kellywstonge@gmail.com
---Purpose--
To help automate the flow of DET test labs cover sheet and data entry.
---Dependancies---
PyQt6 , pandas , pyqtdarktheme'''

import PyQt6
import pandas as PD
import sys
import os
import socket
import time
import traceback
from pathlib import Path

from PyQt6.QtWidgets import (QMainWindow, QApplication, QComboBox, QWidget, QLineEdit, QPushButton, QVBoxLayout, QHBoxLayout,QLabel, QFileDialog, QSpinBox, QCheckBox, QStackedLayout)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PIL import Image, ImageOps

import qdarktheme


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("DET Cover/Entry Tool V.0.0.1")
        self.resize(1200, 700)
        self._setup_ui()
        self._connect_signals()

    def _setup_ui(self):

        #====TOP BAR WIDGET========

        pageLayout = QVBoxLayout()
        button_layout = QHBoxLayout()
        self.stackLayout = QStackedLayout()

        pageLayout.addLayout(button_layout)
        pageLayout.addLayout(self.stackLayout)

        browse_dataBtn = QPushButton("Browse/Data")
        browse_dataBtn.pressed.connect(self.browse_file)
        button_layout.addWidget(browse_dataBtn)

        open_nativeBtn = QPushButton("Open in Xcl")
        open_nativeBtn.pressed.connect(self.open)
        button_layout.addWidget(browse_dataBtn)



        widget = QWidget()
        widget.setLayout(layout)
        self.setCentralWidget(widget)


    def browse_file(self):
        """Open file dialog to select image or PDF."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select File",
            "",
            "Images (*.png *.jpg *.jpeg);;PDF Files (*.pdf);;All Files (*)"
        )

        if file_path:
            self.file_path_input.setText(file_path)
            self.print_button.setEnabled(True)

    def openNativelyInXcel(self):

    	#open the excell program with this file.


    def _connect_signals(self):
        """Connect all signals and slots."""
        #self.submit_button.clicked.connect(self.on_connect)
        #self.text_input.returnPressed.connect(self.on_connect)
        self.browse_dataBtn.clicked.connect(self.browse_file)
        self.open_nativeBtn.clicked.connect(self.openNativelyInXcel)
        #self.print_button.clicked.connect(self.on_print)


def main():
    app = QApplication(sys.argv)
    qdarktheme.setup_theme()
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()