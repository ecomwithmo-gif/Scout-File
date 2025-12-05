"""
Excel Formatter Pro - Main Application
Advanced Excel processing with analytics, streaming, and modern UI.
"""

import customtkinter as ctk
import logging
import sys
import os
from pathlib import Path

# Add project root to Python path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from ui.main_window import MainWindow
from utils.config_manager import config_manager
from utils.exceptions import handle_exception

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('excel_formatter.log'),
        logging.StreamHandler(sys.stdout)
    ]
)

logger = logging.getLogger(__name__)


def main():
    """Main application entry point."""
    try:
        # Initialize configuration
        config_manager.reload_config()
        
        # Create main window
        root = ctk.CTk()
        app = MainWindow(root)
        
        # Set window icon (if available)
        try:
            icon_path = project_root / "assets" / "icon.ico"
            if icon_path.exists():
                root.iconbitmap(str(icon_path))
        except Exception as e:
            logger.warning(f"Could not set window icon: {e}")
        
        # Start application
        logger.info("Starting Excel Formatter Pro")
        root.mainloop()
        
    except Exception as e:
        error = handle_exception(e)
        logger.error(f"Application failed to start: {error.message}")
        
        # Show error dialog
        try:
            import tkinter as tk
            from tkinter import messagebox
            
            root = tk.Tk()
            root.withdraw()  # Hide main window
            messagebox.showerror("Application Error", error.user_friendly_message)
            root.destroy()
        except Exception:
            print(f"Application Error: {error.user_friendly_message}")
        
        sys.exit(1)


if __name__ == "__main__":
    main()
