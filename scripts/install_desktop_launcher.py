"""
Copy the packaged launcher build to the user's desktop and create a shortcut.
"""

from __future__ import annotations

import shutil
from pathlib import Path

import win32com.client  # type: ignore[import]


def get_desktop_path() -> Path:
    return Path.home() / "Desktop"


def copy_launcher_bundle(source_dir: Path, target_dir: Path, icon_path: Path) -> None:
    if target_dir.exists():
        shutil.rmtree(target_dir)
    shutil.copytree(source_dir, target_dir)
    # Ensure the icon is available next to the executable
    shutil.copy2(icon_path, target_dir / icon_path.name)


def create_shortcut(target_dir: Path, exe_name: str, icon_path: Path) -> None:
    desktop = get_desktop_path()
    shortcut_path = desktop / "Excel Formatter Pro.lnk"

    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(str(shortcut_path))
    shortcut.TargetPath = str(target_dir / exe_name)
    shortcut.WorkingDirectory = str(target_dir)
    shortcut.IconLocation = str(icon_path)
    shortcut.Description = "Launch Excel Formatter Pro"
    shortcut.Save()


def main() -> None:
    project_root = Path(__file__).resolve().parent.parent
    source_dir = project_root / "build" / "dist" / "ExcelFormatterProLauncher"
    if not source_dir.exists():
        raise FileNotFoundError("Launcher build not found. Run PyInstaller first.")

    icon_path = project_root / "assets" / "launcher.ico"
    if not icon_path.exists():
        raise FileNotFoundError("Icon file missing at assets/launcher.ico.")

    target_dir = get_desktop_path() / "Excel Formatter Pro"
    copy_launcher_bundle(source_dir, target_dir, icon_path)
    create_shortcut(target_dir, "ExcelFormatterProLauncher.exe", icon_path)
    print(f"Desktop launcher installed at {target_dir}")


if __name__ == "__main__":
    main()







