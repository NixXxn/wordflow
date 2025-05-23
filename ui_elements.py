import tkinter as tk
from tkinter import ttk
from constants import FONTS

class ToolTip:
    """
    Provides a tooltip for a given widget.
    The tooltip appears after a short delay when the mouse enters the widget.
    """
    def __init__(self, widget, text, theme):
        self.widget = widget
        self.text = text
        self.theme = theme
        self.tooltip_window = None
        self.id = None
        self.widget.bind("<Enter>", self.schedule_show, add="+")
        self.widget.bind("<Leave>", self.hide, add="+")
        self.widget.bind("<ButtonPress>", self.hide, add="+")

    def schedule_show(self, event=None):
        """Schedules the tooltip to be shown after a delay."""
        self.unschedule()
        if self.tooltip_window:
            return
        self.id = self.widget.after(600, lambda: self.show(event))

    def unschedule(self):
        """Cancels any pending tooltip show operations."""
        if self.id:
            self.widget.after_cancel(self.id)
            self.id = None

    def show(self, event=None):
        """Displays the tooltip window."""
        if self.tooltip_window:
            return

        x = y = 0
        if event:
            x = event.x_root + 10
            y = event.y_root + 10
        else:
            x, y_rel, _, _ = self.widget.bbox("insert")
            x += self.widget.winfo_rootx() + 25
            y_rel += self.widget.winfo_rooty() + 25
            y = y_rel

        self.tooltip_window = tk.Toplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True)
        self.tooltip_window.wm_geometry(f"+{x}+{y}")

        bg_color = self.theme.get("secondary_bg", "#FFFFE0")
        fg_color = self.theme.get("fg", "#000000")
        border_color = self.theme.get("border", "#AAAAAA")

        frame = tk.Frame(
            self.tooltip_window,
            background=bg_color,
            highlightbackground=border_color,
            highlightthickness=1,
        )
        frame.pack(fill="both", expand=True)

        label = tk.Label(
            frame,
            text=self.text,
            justify=tk.LEFT,
            background=bg_color,
            foreground=fg_color,
            wraplength=250,
            font=FONTS.get("small", ("Segoe UI", 9)),
        )
        label.pack(padx=5, pady=3)

    def hide(self, event=None):
        """Hides the tooltip window."""
        self.unschedule()
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None

def make_draggable(widget, window):
    """
    Makes a Tkinter window draggable by clicking and dragging a specified widget.

    Args:
        widget: The widget that will act as the handle for dragging (e.g., a title bar).
        window: The Toplevel window to be dragged.
    """
    widget._drag_start_x = 0
    widget._drag_start_y = 0

    def on_drag_start(event):
        widget._drag_start_x = event.x
        widget._drag_start_y = event.y

    def on_drag_motion(event):
        dx = event.x - widget._drag_start_x
        dy = event.y - widget._drag_start_y
        new_x = window.winfo_x() + dx
        new_y = window.winfo_y() + dy
        window.geometry(f"+{new_x}+{new_y}")

    widget.bind("<ButtonPress-1>", on_drag_start)
    widget.bind("<B1-Motion>", on_drag_motion)