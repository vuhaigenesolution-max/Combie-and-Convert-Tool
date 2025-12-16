import tkinter as tk
from tkinter import ttk

from Fontend.combie import CombineFrame
from Fontend.convert import ConvertFrame


class App(tk.Tk):
    """Main application window hosting multiple screens."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Metadata Tool")
        self.geometry("820x560")
        self.resizable(True, True)

        self._configure_styles()

        container = ttk.Frame(self, padding=28, style="Root.TFrame")
        container.pack(fill=tk.BOTH, expand=True)

        self.frames: dict[str, ttk.Frame] = {}
        self._register_frames(container)

        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)

        self.show_frame("home")

    def _register_frames(self, container: ttk.Frame) -> None:
        self.frames["home"] = HomeFrame(parent=container, controller=self)
        self.frames["combine"] = CombineFrame(parent=container, controller=self)
        self.frames["convert"] = ConvertFrame(parent=container, controller=self)

        for frame in self.frames.values():
            frame.grid(row=0, column=0, sticky="nsew")

    def show_frame(self, key: str) -> None:
        frame = self.frames[key]
        frame.tkraise()

    def _configure_styles(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")

        bg = "#0c1f2b"
        card_bg = "#122b38"
        accent = "#0fa958"  # requested green accent
        text_primary = "#e9f4ff"

        style.configure("Root.TFrame", background=bg)
        style.configure("Card.TFrame", background=card_bg)
        style.configure("Title.TLabel", background=bg, foreground=text_primary, font=("Segoe UI", 22, "bold"))
        style.configure("Subtitle.TLabel", background=bg, foreground="#9fb6c3", font=("Segoe UI", 12))
        style.configure("CardTitle.TLabel", background=card_bg, foreground=text_primary, font=("Segoe UI", 14, "bold"))
        style.configure(
            "Primary.TButton",
            background=accent,
            foreground="#ffffff",
            padding=(18, 12),
            font=("Segoe UI Semibold", 12),
            borderwidth=0,
        )
        style.map(
            "Primary.TButton",
            background=[("active", "#0c8f4a"), ("pressed", "#0b7d40")],
            relief=[("pressed", "sunken"), ("!pressed", "flat")],
        )
        # Smaller secondary buttons (Home, Browse).
    
        style.configure("Card.TButton", padding=(8, 5), font=("Segoe UI", 10))


class HomeFrame(ttk.Frame):
    """Landing screen to choose a workflow."""

    def __init__(self, parent: tk.Widget, controller: App) -> None:
        super().__init__(parent)
        self.controller = controller
        self.configure(style="Root.TFrame")

        # Hero section
        hero = ttk.Frame(self, style="Root.TFrame")
        hero.pack(fill=tk.X, pady=(10, 30), padx=8)

        # icon = ttk.Label(hero, text="🍃", font=("Segoe UI", 38), background="#0c1f2b", foreground="#0fa958")
        # icon.pack()

        title = ttk.Label(hero, text="Metadata Tool", style="Title.TLabel")
        title.pack(pady=(6, 6))

        subtitle = ttk.Label(hero, text="Choose a workflow to begin", style="Subtitle.TLabel")
        subtitle.pack()

        # Cards for workflows
        cards = ttk.Frame(self, style="Root.TFrame")
        cards.pack(expand=True, fill=tk.BOTH, padx=12)

        card = ttk.Frame(cards, padding=20, style="Card.TFrame")
        card.pack(fill=tk.X, expand=True, padx=40, pady=10)

        ttk.Label(card, text="Combine Data", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(
            card,
            text="Merge data using a source, template, and output location.",
            background="#122b38",
            foreground="#b6c7d4",
            font=("Segoe UI", 11),
        ).pack(anchor="w", pady=(6, 12))

        btn_combine = ttk.Button(
            card,
            text="Start Combine",
            style="Primary.TButton",
            command=lambda: controller.show_frame("combine"),
            width=22,
        )
        btn_combine.pack(anchor="w")

        card2 = ttk.Frame(cards, padding=20, style="Card.TFrame")
        card2.pack(fill=tk.X, expand=True, padx=40, pady=10)

        ttk.Label(card2, text="Convert Data", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(
            card2,
            text="Planned workflow for data conversion.",
            background="#122b38",
            foreground="#b6c7d4",
            font=("Segoe UI", 11),
        ).pack(anchor="w", pady=(6, 12))

        btn_convert = ttk.Button(
            card2,
            text="Start Convert",
            style="Primary.TButton",
            command=lambda: controller.show_frame("convert"),
            width=22,
        )
        btn_convert.pack(anchor="w")


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
