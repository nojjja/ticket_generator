"""Shared visual tokens and ttk styling for the ticket generator UI."""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk


COLORS = {
    "background": "#101521",
    "surface": "#171F2E",
    "surface_border": "#29364A",
    "elevated": "#202B3D",
    "text": "#F2F5FB",
    "secondary": "#AAB7CA",
    "disabled": "#738198",
    "accent": "#705CFF",
    "accent_hover": "#8575FF",
    "accent_pressed": "#5D49E8",
    "white": "#FFFFFF",
    "success": "#3DDC97",
    "warning": "#FFB547",
    "error": "#FF5C7A",
    "input": "#131C2A",
}

SPACE = {"xs": 4, "sm": 8, "md": 12, "lg": 16, "xl": 24, "xxl": 32}
SIZE = {
    "main_width": 740,
    "main_height": 600,
    "main_min_width": 600,
    "main_min_height": 540,
    "card_width": 520,
    "main_card_height": 380,
    "result_width": 680,
    "result_height": 600,
    "result_min_width": 560,
    "result_min_height": 460,
    "lock_dialog_width": 440,
    "lock_dialog_height": 200,
    "button_height": 48,
    "badge_size": 32,
    "corner": 14,
    "outline": 1,
}
_BASE_SPACE = SPACE.copy()
_BASE_SIZE = SIZE.copy()
_BASE_TK_SCALING = 96 / 72
TIMING = {"ui_poll_ms": 100}
FONT = {
    "family": "Segoe UI",
    "title": 27,
    "heading": 18,
    "body": 14,
    "caption": 12,
    "hero": 44,
}


def _scale_layout(scale: float) -> None:
    for dimensions, base_dimensions in ((SPACE, _BASE_SPACE), (SIZE, _BASE_SIZE)):
        dimensions.update(
            {name: max(1, round(value * scale)) for name, value in base_dimensions.items()}
        )


def _font(size: int, weight: str = "normal") -> tuple[str, int, str]:
    return FONT["family"], size, weight


def apply_theme(root: tk.Tk) -> ttk.Style:
    """Install the dark clam theme and shared widget styles."""
    tk_scaling = float(root.tk.call("tk", "scaling"))
    _scale_layout(tk_scaling / _BASE_TK_SCALING)
    style = ttk.Style(root)
    style.theme_use("clam")
    families = set(root.tk.call("font", "families"))
    FONT["family"] = next((family for family in ("Segoe UI", "Arial") if family in families), "TkDefaultFont")
    root.configure(background=COLORS["background"])
    root.option_add("*TCombobox*Listbox.background", COLORS["elevated"])
    root.option_add("*TCombobox*Listbox.foreground", COLORS["text"])
    root.option_add("*TCombobox*Listbox.selectBackground", COLORS["accent"])
    root.option_add("*TCombobox*Listbox.selectForeground", COLORS["white"])
    root.option_add("*TCombobox*Listbox.font", _font(FONT["body"]))

    style.configure("App.TFrame", background=COLORS["background"])
    style.configure("Surface.TFrame", background=COLORS["surface"])
    style.configure("Question.TFrame", background=COLORS["elevated"])
    style.configure("Header.TLabel", background=COLORS["background"], foreground=COLORS["text"], font=_font(FONT["title"], "bold"))
    style.configure("Subtitle.TLabel", background=COLORS["background"], foreground=COLORS["secondary"], font=_font(FONT["body"]))
    style.configure("Card.TLabel", background=COLORS["surface"], foreground=COLORS["text"], font=_font(FONT["body"]))
    style.configure("MutedCard.TLabel", background=COLORS["surface"], foreground=COLORS["secondary"], font=_font(FONT["caption"]))
    style.configure("Question.TLabel", background=COLORS["elevated"], foreground=COLORS["text"], font=_font(FONT["body"]))
    style.configure("Badge.TLabel", background=COLORS["accent"], foreground=COLORS["white"], font=_font(FONT["body"], "bold"), anchor="center")
    style.configure("Status.TLabel", background=COLORS["background"], foreground=COLORS["secondary"], font=_font(FONT["caption"]))
    style.configure("Success.TLabel", background=COLORS["background"], foreground=COLORS["success"], font=_font(FONT["caption"]))
    style.configure("WarningCard.TLabel", background=COLORS["surface"], foreground=COLORS["warning"], font=_font(FONT["caption"]))
    style.configure("ErrorCard.TLabel", background=COLORS["surface"], foreground=COLORS["error"], font=_font(FONT["caption"]))
    style.configure("Warning.TLabel", background=COLORS["background"], foreground=COLORS["warning"], font=_font(FONT["caption"]))
    style.configure("Error.TLabel", background=COLORS["background"], foreground=COLORS["error"], font=_font(FONT["caption"]))
    style.configure("Pill.TLabel", background=COLORS["elevated"], foreground=COLORS["secondary"], padding=(SPACE["md"], SPACE["sm"]), font=_font(FONT["caption"]))
    style.configure("Hero.TLabel", background=COLORS["background"], foreground=COLORS["accent_hover"], font=_font(FONT["hero"], "bold"))
    style.configure("Banner.TLabel", background=COLORS["elevated"], foreground=COLORS["warning"], padding=SPACE["md"], font=_font(FONT["body"]))

    style.configure("Primary.TButton", background=COLORS["accent"], foreground=COLORS["white"], font=_font(FONT["body"], "bold"), padding=(SPACE["lg"], SPACE["md"]), borderwidth=0, focusthickness=SIZE["outline"], focuscolor=COLORS["accent_hover"])
    style.map("Primary.TButton", background=[("disabled", COLORS["elevated"]), ("pressed", COLORS["accent_pressed"]), ("active", COLORS["accent_hover"])], foreground=[("disabled", COLORS["disabled"]), ("!disabled", COLORS["white"])])
    style.configure("Ghost.TButton", background=COLORS["background"], foreground=COLORS["secondary"], font=_font(FONT["caption"]), padding=(SPACE["sm"], SPACE["xs"]), borderwidth=0, focusthickness=SIZE["outline"], focuscolor=COLORS["accent"])
    style.map("Ghost.TButton", background=[("pressed", COLORS["surface"]), ("active", COLORS["elevated"])], foreground=[("active", COLORS["text"])])
    style.configure("Secondary.TButton", background=COLORS["elevated"], foreground=COLORS["text"], font=_font(FONT["body"]), padding=(SPACE["lg"], SPACE["sm"]), borderwidth=0, focusthickness=SIZE["outline"], focuscolor=COLORS["accent"])
    style.map("Secondary.TButton", background=[("pressed", COLORS["surface_border"]), ("active", COLORS["surface_border"])])

    style.configure("TCombobox", fieldbackground=COLORS["input"], background=COLORS["elevated"], foreground=COLORS["text"], insertcolor=COLORS["text"], arrowcolor=COLORS["secondary"], bordercolor=COLORS["surface_border"], lightcolor=COLORS["surface_border"], darkcolor=COLORS["surface_border"], padding=(SPACE["md"], SPACE["md"]), focusthickness=SIZE["outline"], focuscolor=COLORS["accent"])
    style.map("TCombobox", fieldbackground=[("disabled", COLORS["surface"]), ("readonly", COLORS["input"])], foreground=[("disabled", COLORS["disabled"]), ("readonly", COLORS["text"])], selectbackground=[("readonly", COLORS["input"])], selectforeground=[("readonly", COLORS["text"])], bordercolor=[("focus", COLORS["accent"])])
    return style


def rounded_panel(parent: tk.Misc, *, fill: str, outline: str | None = None, radius: int | None = None, content_style: str = "Surface.TFrame") -> tk.Canvas:
    """Canvas-backed rounded panel with an embedded themed frame."""
    radius = radius or SIZE["corner"]
    panel = tk.Canvas(parent, background=COLORS["background"], highlightthickness=0, bd=0)
    panel.content = ttk.Frame(panel, style=content_style)
    panel.minimum_height = 0
    padding = radius

    def draw(_event: tk.Event | None = None) -> None:
        panel.delete("panel")
        width, height = panel.winfo_width(), panel.winfo_height()
        if width <= 1 or height <= 1:
            return
        r = min(radius, width // 2, height // 2)
        panel.create_polygon(
            (r, 0, width-r, 0, width, 0, width, r, width, height-r, width, height,
             width-r, height, r, height, 0, height, 0, height-r, 0, r, 0, 0),
            smooth=True, splinesteps=12, fill=fill, outline=outline or fill,
            width=SIZE["outline"], tags="panel",
        )
        panel.coords("content", padding, padding)
        panel.itemconfigure("content", width=max(1, width-padding*2))
        panel.content.update_idletasks()
        panel.configure(height=max(panel.minimum_height, panel.content.winfo_reqheight() + padding*2))
        panel.tag_lower("panel")

    panel.create_window(padding, padding, anchor="nw", window=panel.content, tags="content")
    panel.bind("<Configure>", draw)
    return panel
