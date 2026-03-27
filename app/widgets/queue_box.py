"""
Queue box widget for displaying and managing queued errors.
"""

from typing import List
from PyQt6.QtWidgets import (
    QFrame, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QScrollArea, QWidget
)
from PyQt6.QtCore import Qt, pyqtSignal

from ..models import ErrorEntry


class QueueBox(QFrame):
    """Widget for displaying the error queue and action buttons."""
    
    # Signals for actions
    save_requested = pyqtSignal(list)  # Emits list of ErrorEntry
    print_requested = pyqtSignal(list)
    queue_cleared = pyqtSignal()
    
    STYLESHEET = """
        #queueBox {
            background-color: #333;
            border: 1px solid #555;
            border-radius: 8px;
        }
        QLabel#boxTitle {
            font-size: 14px;
            font-weight: bold;
            color: #fff;
        }
        QLabel#queueItem {
            color: #ccc;
            font-size: 10px;
            padding: 4px;
            background-color: #2a2a2a;
            border-radius: 3px;
        }
        QLabel#emptyQueue {
            color: #666;
            font-size: 11px;
        }
    """
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("queueBox")
        self.setStyleSheet(self.STYLESHEET)
        
        self._queue: List[ErrorEntry] = []
        self._setup_ui()
        self._connect_signals()
    
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 10, 12, 12)
        layout.setSpacing(8)
        
        # Title with count
        title_row = QHBoxLayout()
        title = QLabel("📤 Queue")
        title.setObjectName("boxTitle")
        self.count_label = QLabel("(0)")
        self.count_label.setStyleSheet("color: #888; font-size: 12px;")
        title_row.addWidget(title)
        title_row.addWidget(self.count_label)
        title_row.addStretch()
        layout.addLayout(title_row)
        
        # Queue list (scrollable)
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setStyleSheet("""
            QScrollArea { 
                border: none; 
                background-color: transparent;
            }
        """)
        
        self.scroll_content = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_content)
        self.scroll_layout.setContentsMargins(0, 0, 0, 0)
        self.scroll_layout.setSpacing(4)
        
        self._show_empty_state()
        
        self.scroll.setWidget(self.scroll_content)
        layout.addWidget(self.scroll)
        
        # Action buttons
        btn_layout = QVBoxLayout()
        btn_layout.setSpacing(4)
        
        self.save_btn = QPushButton("💾 Save to Excel")
        self.save_btn.setFixedHeight(28)
        self.save_btn.setEnabled(False)
        
        self.print_btn = QPushButton("🖨️ Print Covers")
        self.print_btn.setFixedHeight(28)
        self.print_btn.setEnabled(False)
        
        self.clear_btn = QPushButton("🗑️ Clear")
        self.clear_btn.setFixedHeight(28)
        self.clear_btn.setEnabled(False)
        
        btn_layout.addWidget(self.save_btn)
        btn_layout.addWidget(self.print_btn)
        btn_layout.addWidget(self.clear_btn)
        layout.addLayout(btn_layout)
    
    def _connect_signals(self):
        self.save_btn.clicked.connect(self._on_save)
        self.print_btn.clicked.connect(self._on_print)
        self.clear_btn.clicked.connect(self.clear)
    
    def _show_empty_state(self):
        label = QLabel("No errors queued")
        label.setObjectName("emptyQueue")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.scroll_layout.addWidget(label)
        self.scroll_layout.addStretch()
    
    def _clear_layout(self):
        while self.scroll_layout.count() > 0:
            item = self.scroll_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
    
    def _refresh_display(self):
        self._clear_layout()
        self.count_label.setText(f"({len(self._queue)})")
        
        if not self._queue:
            self._show_empty_state()
            self.save_btn.setEnabled(False)
            self.print_btn.setEnabled(False)
            self.clear_btn.setEnabled(False)
            return
        
        for entry in self._queue:
            label_text = f"{entry.error_type}\n{entry.printer_summary} • {entry.page_summary}"
            item = QLabel(label_text)
            item.setObjectName("queueItem")
            item.setWordWrap(True)
            self.scroll_layout.addWidget(item)
        
        self.scroll_layout.addStretch()
        
        self.save_btn.setEnabled(True)
        self.print_btn.setEnabled(True)
        self.clear_btn.setEnabled(True)
    
    def add_entry(self, entry: ErrorEntry):
        """Add an error entry to the queue."""
        self._queue.append(entry)
        self._refresh_display()
        print(f"Queued: {entry.error_type} on {entry.printer_summary}")
    
    def clear(self):
        """Clear all entries from the queue."""
        self._queue.clear()
        self._refresh_display()
        self.queue_cleared.emit()
        print("Queue cleared")
    
    def get_queue(self) -> List[ErrorEntry]:
        """Return a copy of the current queue."""
        return list(self._queue)
    
    def _on_save(self):
        self.save_requested.emit(list(self._queue))
    
    def _on_print(self):
        self.print_requested.emit(list(self._queue))
