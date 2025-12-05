"""
Test script to verify Excel Formatter Pro installation and dependencies.
"""

import sys
import importlib
from pathlib import Path

def test_imports():
    """Test all required imports."""
    print("Testing imports...")
    
    try:
        import customtkinter as ctk
        print("✓ CustomTkinter imported successfully")
    except ImportError as e:
        print(f"✗ CustomTkinter import failed: {e}")
        return False
    
    try:
        import pandas as pd
        print("✓ Pandas imported successfully")
    except ImportError as e:
        print(f"✗ Pandas import failed: {e}")
        return False
    
    try:
        import numpy as np
        print("✓ NumPy imported successfully")
    except ImportError as e:
        print(f"✗ NumPy import failed: {e}")
        return False
    
    try:
        import openpyxl
        print("✓ OpenPyXL imported successfully")
    except ImportError as e:
        print(f"✗ OpenPyXL import failed: {e}")
        return False
    
    try:
        import matplotlib.pyplot as plt
        print("✓ Matplotlib imported successfully")
    except ImportError as e:
        print(f"✗ Matplotlib import failed: {e}")
        return False
    
    try:
        import seaborn as sns
        print("✓ Seaborn imported successfully")
    except ImportError as e:
        print(f"✗ Seaborn import failed: {e}")
        return False
    
    return True

def test_project_structure():
    """Test project structure and files."""
    print("\nTesting project structure...")
    
    project_root = Path(__file__).parent
    required_files = [
        "main.py",
        "requirements.txt",
        "setup.py",
        "config/settings.json",
        "config/header_mappings.json",
        "utils/config_manager.py",
        "utils/exceptions.py",
        "utils/data_validator.py",
        "core/excel_processor.py",
        "analytics/data_analytics.py",
        "ui/main_window.py"
    ]
    
    all_exist = True
    for file_path in required_files:
        full_path = project_root / file_path
        if full_path.exists():
            print(f"✓ {file_path}")
        else:
            print(f"✗ {file_path} - Missing!")
            all_exist = False
    
    return all_exist

def test_config_loading():
    """Test configuration loading."""
    print("\nTesting configuration loading...")
    
    try:
        from utils.config_manager import config_manager
        
        # Test settings loading
        settings = config_manager.load_settings()
        print("✓ Settings loaded successfully")
        
        # Test header mappings loading
        mappings = config_manager.load_header_mappings()
        print("✓ Header mappings loaded successfully")
        
        return True
    except Exception as e:
        print(f"✗ Configuration loading failed: {e}")
        return False

def test_ui_components():
    """Test UI component imports."""
    print("\nTesting UI components...")
    
    try:
        from ui.main_window import MainWindow
        print("✓ MainWindow imported successfully")
        
        from ui.components.file_upload import FileUploadSection
        print("✓ FileUploadSection imported successfully")
        
        from ui.components.settings_panel import SettingsPanel
        print("✓ SettingsPanel imported successfully")
        
        from ui.components.progress_panel import ProgressPanel
        print("✓ ProgressPanel imported successfully")
        
        from ui.components.analytics_panel import AnalyticsPanel
        print("✓ AnalyticsPanel imported successfully")
        
        from ui.components.status_bar import StatusBar
        print("✓ StatusBar imported successfully")
        
        return True
    except Exception as e:
        print(f"✗ UI components import failed: {e}")
        return False

def main():
    """Run all tests."""
    print("Excel Formatter Pro - Installation Test")
    print("=" * 50)
    
    tests = [
        ("Dependencies", test_imports),
        ("Project Structure", test_project_structure),
        ("Configuration", test_config_loading),
        ("UI Components", test_ui_components)
    ]
    
    all_passed = True
    
    for test_name, test_func in tests:
        print(f"\n{test_name} Test:")
        print("-" * 20)
        if not test_func():
            all_passed = False
    
    print("\n" + "=" * 50)
    if all_passed:
        print("🎉 All tests passed! Excel Formatter Pro is ready to use.")
        print("\nTo run the application:")
        print("python main.py")
    else:
        print("❌ Some tests failed. Please check the errors above.")
        print("\nTo install missing dependencies:")
        print("pip install -r requirements.txt")
    
    return all_passed

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
