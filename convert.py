import tkinter as tk
from tkinter import ttk


class ConvertFrame(ttk.Frame):
    """Placeholder for future conversion workflow."""

    def __init__(self, parent: tk.Widget, controller) -> None:
        super().__init__(parent)
        self.controller = controller

        header = ttk.Frame(self)
        header.pack(fill=tk.X, pady=(10, 20))

        ttk.Label(header, text="Convert Data", font=("Segoe UI", 16, "bold")).pack(side=tk.LEFT)
        ttk.Button(header, text="Home", command=lambda: controller.show_frame("home")).pack(side=tk.RIGHT)

        body = ttk.Frame(self)
        body.pack(fill=tk.BOTH, expand=True)
        ttk.Label(
            body,
            text="Conversion workflow coming soon. Define steps and wire business logic here.",
            wraplength=500,
            justify=tk.LEFT,
        ).pack(pady=20)
