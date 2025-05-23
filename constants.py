import os
import sys
import datetime

# --- Application Directory and Logging ---
def get_app_dir():
    """Returns the path to the application's data directory."""
    app_dir = os.path.join(os.path.expanduser("~"), "TextExpander")
    os.makedirs(app_dir, exist_ok=True)
    return app_dir

APP_DIR = get_app_dir()
LOG_FILE = os.path.join(APP_DIR, "text_expander.log")


# --- Optional Module Availability ---
KEYBOARD_AVAILABLE = False
CLIPBOARD_AVAILABLE = False
SYSTRAY_AVAILABLE = False

try:
    import pynput.keyboard
    import pynput.mouse
    KEYBOARD_AVAILABLE = True
except ImportError:
    pass

try:
    import pyperclip
    CLIPBOARD_AVAILABLE = True
except ImportError:
    pass

try:
    import pystray
    from PIL import Image, ImageDraw, ImageFont
    SYSTRAY_AVAILABLE = True
except ImportError:
    pass


# --- Theme Colors (Only Light Theme) ---
LIGHT_THEME = {
    "bg": "#ffffff",
    "fg": "#333333",          # Dark grey for text
    "accent": "#0071e3",
    "hover": "#005bb5",       # Slightly darker accent for hover
    "secondary_bg": "#f0f0f0", # Lighter grey for secondary backgrounds (like editor bg, combobox)
    "border": "#d1d1d6",
    "success": "#34c759",
    "warning": "#ff9500",
    "error": "#ff3b30",
    "highlight": "#e6f0ff",   # Light blue highlight
    "button_text": "#333333", # Black text for buttons in light mode
    "active_tab_text": "#333333" # Changed to black for active tab in light mode
}

# --- Fonts ---
FONTS = {
    "heading": ("Segoe UI", 12, "bold"),
    "subheading": ("Segoe UI", 10, "bold"),
    "body": ("Segoe UI", 10),
    "small": ("Segoe UI", 9),
    "button": ("Segoe UI", 9, "bold"),
    "monospace": ("Consolas", 10),
}