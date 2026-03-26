"""
File utility functions for DET Cover/Entry Tool
"""

import os
import platform
import subprocess


def truncate_path(path: str, max_length: int = 50) -> str:
    """Truncate path for display, keeping filename visible."""
    if len(path) <= max_length:
        return path
    
    filename = os.path.basename(path)
    if len(filename) >= max_length - 5:
        return "..." + filename[-(max_length - 3):]
    
    remaining = max_length - len(filename) - 4  # 4 for ".../"
    return path[:remaining] + ".../" + filename


def open_in_default_app(file_path: str) -> bool:
    """
    Open the specified file in the system's default application.
    Returns True on success, False on failure.
    """
    if not file_path or not os.path.exists(file_path):
        return False

    try:
        if platform.system() == "Windows":
            os.startfile(file_path)
        elif platform.system() == "Darwin":  # macOS
            subprocess.run(["open", file_path], check=True)
        else:  # Linux
            subprocess.run(["xdg-open", file_path], check=True)
        return True
    except Exception as e:
        print(f"Error opening file: {e}")
        return False
