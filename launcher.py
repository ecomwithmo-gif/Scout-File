"""
Excel Formatter Pro Launcher

Provides a lightweight launcher window with a single button to start the
main Excel Formatter Pro UI.
"""

import customtkinter as ctk

from excel_formatter_app import ExcelFormatterApp


def launch_excel_formatter(root: ctk.CTk) -> None:
    """Swap the launcher layout for the main Excel Formatter UI."""
    # Clean up launcher widgets
    for widget in root.winfo_children():
        widget.destroy()

    # Delegate to the main application (it will resize/configure the window)
    ExcelFormatterApp(root)


def build_launcher_ui(root: ctk.CTk) -> None:
    """Create the minimal launcher UI with a single launch button."""
    root.title("Formatter Launcher")
    root.geometry("520x360")
    root.resizable(False, False)
    root.configure(fg_color="#0b1220")

    container = ctk.CTkFrame(root, fg_color="#ffffff", corner_radius=18)
    container.pack(expand=True, fill="both", padx=28, pady=28)

    header = ctk.CTkFrame(container, fg_color="#f8fafc", corner_radius=12)
    header.pack(fill="x", padx=16, pady=(16, 12))

    badge = ctk.CTkLabel(
        header,
        text="LAUNCHER",
        font=ctk.CTkFont(family="Inter", size=12, weight="bold"),
        text_color="#1d4ed8",
        fg_color="#e0ecff",
        corner_radius=8,
        padx=12,
        pady=6,
    )
    badge.pack(side="left")

    status = ctk.CTkLabel(
        header,
        text="Ready to start",
        font=ctk.CTkFont(family="Inter", size=12),
        text_color="#475569",
    )
    status.pack(side="right")

    title = ctk.CTkLabel(
        container,
        text="Formatter workspace",
        font=ctk.CTkFont(family="Inter", size=30, weight="bold"),
        text_color="#0f172a",
    )
    title.pack(pady=(4, 4))

    subtitle = ctk.CTkLabel(
        container,
        text="A clean, focused way to jump into your formatting workspace.",
        font=ctk.CTkFont(family="Inter", size=14),
        text_color="#475569",
        wraplength=440,
        justify="center",
    )
    subtitle.pack(pady=(0, 12))

    detail_block = ctk.CTkFrame(container, fg_color="#f8fafc", corner_radius=12)
    detail_block.pack(fill="x", padx=16, pady=(4, 18))

    for line in (
        "- Launches the full formatter window instantly.",
        "- Keeps your taskbar shortcut light and distraction-free.",
        "- Safe to close anytime; your project loads remain in the main app.",
    ):
        label = ctk.CTkLabel(
            detail_block,
            text=line,
            font=ctk.CTkFont(family="Inter", size=13),
            text_color="#334155",
            anchor="w",
            justify="left",
        )
        label.pack(fill="x", padx=12, pady=(8, 0))

    launch_button = ctk.CTkButton(
        container,
        text="Open Formatter",
        command=lambda: launch_excel_formatter(root),
        width=240,
        height=50,
        corner_radius=14,
        font=ctk.CTkFont(family="Inter", size=17, weight="bold"),
        fg_color="#1d4ed8",
        hover_color="#1e40af",
        text_color="#ffffff",
    )
    launch_button.pack(pady=(4, 18))

    helper = ctk.CTkLabel(
        container,
        text="Tip: Pin this window to your taskbar so the formatter is always one click away.",
        font=ctk.CTkFont(family="Inter", size=12),
        text_color="#475569",
        wraplength=440,
        justify="center",
    )
    helper.pack(pady=(0, 12))


def main() -> None:
    """Entry point for the launcher."""
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    build_launcher_ui(root)
    root.mainloop()


if __name__ == "__main__":
    main()


