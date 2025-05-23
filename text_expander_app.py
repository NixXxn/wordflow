# text_expander_app.py

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import os
import datetime
import threading
import sys
import time
import re

# Import local modules
from config_manager import ConfigManager
from snippet_manager import SnippetManager
from ui_elements import ToolTip, make_draggable
from constants import LIGHT_THEME, FONTS, APP_DIR, SYSTRAY_AVAILABLE, KEYBOARD_AVAILABLE, CLIPBOARD_AVAILABLE
from utils import log

# Optional imports for external modules
try:
    from pynput import keyboard, mouse
except ImportError:
    pass

try:
    import pyperclip
except ImportError:
    pass

try:
    import pystray
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    pass


class ModernTextExpander:
    """
    The main application class for the Text Expander.
    Manages UI, data, and keyboard/mouse listeners.
    """
    def __init__(self, root):
        self.root = root
        self.root.geometry("800x600")
        self.root.minsize(700, 550)

        # File paths
        self.snippets_file = os.path.join(APP_DIR, "snippets.json")
        self.config_file = os.path.join(APP_DIR, "config.json")
        self.icon_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "icon.ico"
        )

        # Default configuration (only light theme)
        self.default_config = {
            "theme": "light", # Default to light theme
            "minimize_to_tray": True,
            "start_minimized": False,
            "show_tooltips": True,
            "auto_backup": True,
            "backup_interval_days": 7,
            "last_backup": None,
            "default_category": "General",
        }

        # Initialize managers
        self.config_manager = ConfigManager(APP_DIR, self.default_config)
        self.config = self.config_manager.config
        self.snippet_manager = SnippetManager(APP_DIR)
        self.snippets = self.snippet_manager.get_all_snippets()

        # Initialize core state variables early
        self.snippet_history = []
        self.current_input = ""
        self.listener = None
        self.is_listening = False
        self.mouse_listener = None
        self.last_mouse_pos = (0, 0)
        self.tray_icon = None
        self.status_var = tk.StringVar(value="Ready")

        # Initialize tooltips dictionary very early, before any UI elements are created that might need it
        self.tooltips = {}

        # Theme detection and application (will now always be light theme)
        self._detect_and_set_theme()
        self.apply_theme()

        # Create icon (definition placed before this call)
        self.create_default_icon()

        # Create main frame
        self.main_frame = ttk.Frame(root)
        self.main_frame.pack(fill="both", expand=True)

        # Create notebook (tabs)
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # Create tabs (definitions placed before these calls)
        self.create_snippets_tab()
        self.create_settings_tab()
        self.create_help_tab()

        # Now, with all UI widgets created, set up tooltips if enabled in config.
        if self.config_manager.get("show_tooltips", True):
            self.setup_tooltips()

        # Status bar with enhanced visuals
        self.status_frame = ttk.Frame(self.main_frame)
        self.status_frame.pack(fill="x", side="bottom", padx=10, pady=5)

        self.status_indicator = ttk.Label(
            self.status_frame, text="●", foreground=self.theme["success"]
        )
        self.status_indicator.pack(side="left", padx=(0, 5))

        ttk.Label(self.status_frame, textvariable=self.status_var).pack(side="left")

        self.toggle_btn = ttk.Button(
            self.status_frame,
            text="Pause Listening" if KEYBOARD_AVAILABLE else "Keyboard Not Available",
            command=self.toggle_listener,
            style="Accent.TButton" if KEYBOARD_AVAILABLE else "TButton",
        )
        self.toggle_btn.pack(side="right")

        # Set up window icon
        if os.path.exists(self.icon_path):
            try:
                self.root.iconbitmap(self.icon_path)
            except Exception as e:
                log(f"Error setting window icon: {e}")

        # Setup systray if available and configured (definition placed before this call)
        if SYSTRAY_AVAILABLE and self.config_manager.get("minimize_to_tray", True):
            self.setup_systray()

        # Handle window close protocol
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # Start keyboard listener if available
        if KEYBOARD_AVAILABLE:
            self.start_listener()
            self.start_mouse_tracking()

        # Check for backup if auto-backup enabled
        if self.config_manager.get("auto_backup", True):
            self.check_auto_backup()

        # Start minimized if configured
        if self.config_manager.get("start_minimized", False):
            self.root.after(500, self.root.withdraw)

        # Schedule a check for dependencies with a slight delay
        self.root.after(1000, self.check_dependencies)

    def _detect_and_set_theme(self):
        """Sets the internal theme to LIGHT_THEME as Dark Mode is removed."""
        self.theme = LIGHT_THEME
        self.theme_var = tk.StringVar(value="light") # Always set to 'light' for consistency

    def create_default_icon(self):
        """Creates a default icon file if none exists or loading fails, suitable for tray and window."""
        try:
            icon_file = os.path.join(APP_DIR, "icon.ico")
            if os.path.exists(icon_file):
                self.icon_path = icon_file
                return

            if SYSTRAY_AVAILABLE:
                img = Image.new("RGBA", (64, 64), color=(0, 0, 0, 0))
                draw = ImageDraw.Draw(img)

                background_color = self.theme["accent"]
                draw.rectangle([(0, 0), (63, 63)], fill=background_color)

                try:
                    font = ImageFont.truetype("arial.ttf", 24)
                except IOError:
                    try: # Fallback to a common system font if Arial isn't found
                        font = ImageFont.truetype("segoe ui bold.ttf", 24)
                    except:
                        font = ImageFont.load_default() # Fallback to default PIL font

                text = "TE"
                text_color = (255, 255, 255)

                try:
                    left, top, right, bottom = draw.textbbox((0, 0), text, font=font)
                    text_width = right - left
                    text_height = bottom - top
                    position = ((64 - text_width) // 2, (64 - text_height) // 2)
                except AttributeError:
                    text_width, text_height = draw.textsize(text, font=font)
                    position = ((64 - text_width) // 2, (64 - text_height) // 2)

                draw.text(position, text, font=font, fill=text_color)
                img.save(icon_file, format="ICO")
                self.icon_path = icon_file
                log(f"Created default icon at {icon_file}")
        except Exception as e:
            log(f"Error creating default icon: {e}")

    def apply_theme(self):
        """Applies the current theme colors and styles to the UI."""
        style = ttk.Style()

        bg_color = self.theme["bg"]
        fg_color = self.theme["fg"]
        accent = self.theme["accent"]
        hover = self.theme["hover"]
        secondary_bg = self.theme["secondary_bg"]
        border_color = self.theme["border"]
        button_text_color = self.theme["button_text"]
        active_tab_text_color = self.theme["active_tab_text"]


        style.configure("TFrame", background=bg_color)
        style.configure("TLabel", background=bg_color, foreground=fg_color)
        style.configure("TLabelframe", background=bg_color, foreground=fg_color)
        style.configure("TLabelframe.Label", background=bg_color, foreground=fg_color)

        style.configure(
            "TButton",
            background=secondary_bg,
            foreground=fg_color,
            relief="raised",
            borderwidth=1,
            padding=(10, 5),
            font=FONTS["button"],
        )
        style.map(
            "TButton",
            background=[("active", hover), ("pressed", border_color)],
            foreground=[("active", fg_color)],
            relief=[("pressed", "sunken")],
        )

        style.configure(
            "Accent.TButton",
            background=accent,
            foreground=button_text_color,
            borderwidth=0,
            padding=(10, 5),
            font=FONTS["button"],
        )
        style.map(
            "Accent.TButton",
            background=[("active", hover), ("pressed", accent)],
            foreground=[("active", button_text_color), ("pressed", button_text_color)],
            relief=[("pressed", "sunken")],
        )

        style.configure(
            "Danger.TButton",
            background=self.theme["error"],
            foreground=button_text_color,
            borderwidth=0,
            padding=(10, 5),
            font=FONTS["button"],
        )
        style.map(
            "Danger.TButton",
            background=[("active", "#ff6b64"), ("pressed", self.theme["error"])],
            foreground=[("active", button_text_color)],
            relief=[("pressed", "sunken")],
        )

        style.configure(
            "Secondary.TButton",
            background=secondary_bg,
            foreground=fg_color,
            borderwidth=1,
            padding=(8, 3),
            font=FONTS["button"],
        )
        style.map(
            "Secondary.TButton",
            background=[("active", border_color), ("pressed", secondary_bg)],
            foreground=[("active", fg_color)],
            relief=[("pressed", "sunken")],
        )

        style.configure(
            "Link.TButton",
            background=bg_color,
            foreground=accent,
            borderwidth=0,
            padding=0,
            font=FONTS["body"],
        )
        style.map(
            "Link.TButton",
            foreground=[("active", hover)],
            background=[("active", bg_color)],
        )

        style.configure(
            "TEntry",
            fieldbackground=secondary_bg,
            foreground=fg_color,
            bordercolor=border_color,
            font=FONTS["body"],
        )
        style.map(
            "TEntry",
            fieldbackground=[("disabled", secondary_bg)],
            bordercolor=[("focus", accent)],
        )

        style.configure(
            "TCombobox",
            background=secondary_bg,
            fieldbackground=secondary_bg,
            foreground=fg_color,
            arrowcolor=fg_color,
            font=FONTS["body"],
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", secondary_bg)],
            background=[("readonly", secondary_bg)],
        )

        style.configure("TCheckbutton", background=bg_color, foreground=fg_color, font=FONTS["body"])
        style.map("TCheckbutton", background=[("active", bg_color)])

        style.configure("TRadiobutton", background=bg_color, foreground=fg_color, font=FONTS["body"])
        style.map("TRadiobutton", background=[("active", bg_color)])

        style.configure("TNotebook", background=bg_color, borderwidth=0)
        style.configure(
            "TNotebook.Tab",
            background=secondary_bg,
            foreground=fg_color,
            padding=(15, 5),
            borderwidth=0,
            font=FONTS["subheading"],
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", accent)],
            foreground=[("selected", active_tab_text_color)],
        )

        style.configure(
            "Treeview",
            background=bg_color,
            fieldbackground=bg_color,
            foreground=fg_color,
            borderwidth=0,
            font=FONTS["body"],
        )
        style.map(
            "Treeview",
            background=[("selected", accent)],
            foreground=[("selected", active_tab_text_color)],
        )

        style.configure(
            "Treeview.Heading",
            background=secondary_bg,
            foreground=fg_color,
            relief="flat",
            font=FONTS["subheading"],
        )
        style.map("Treeview.Heading", background=[("active", secondary_bg)])

        style.configure(
            "Vertical.TScrollbar",
            background=secondary_bg,
            arrowcolor=fg_color,
            borderwidth=0,
            relief="flat",
        )
        style.map(
            "Vertical.TScrollbar",
            background=[("active", border_color), ("pressed", border_color)],
        )

        style.configure("TSeparator", background=border_color)

        style.configure(
            "TSpinbox",
            background=secondary_bg,
            fieldbackground=secondary_bg,
            foreground=fg_color,
            arrowcolor=fg_color,
            font=FONTS["body"],
        )

        if hasattr(self, "line_numbers"):
            self.line_numbers.configure(bg=secondary_bg)

        if hasattr(self, "editor"):
            self.editor.configure(bg=bg_color, fg=fg_color, insertbackground=fg_color)
            self.editor.tag_configure(
                "placeholder", foreground=accent, font=FONTS["monospace"] + ("bold",)
            )
            self.editor.tag_configure(
                "input_placeholder",
                foreground=self.theme["success"],
                font=FONTS["monospace"] + ("bold",),
            )
            self.editor.tag_configure(
                "unknown_placeholder",
                foreground=self.theme["error"],
                font=FONTS["monospace"] + ("italic",),
            )

        self.root.configure(bg=bg_color)

        if hasattr(self, "status_indicator") and hasattr(self, "status_var"):
            current_status_color = self.theme["success"] if self.is_listening and KEYBOARD_AVAILABLE else self.theme["warning"]
            if not KEYBOARD_AVAILABLE:
                current_status_color = self.theme["error"]
            self.status_indicator.configure(foreground=current_status_color)

        if hasattr(self, "editor"):
            self.highlight_placeholders()

    def setup_systray(self):
        """Sets up the system tray icon."""
        if not SYSTRAY_AVAILABLE:
            return

        try:
            icon_image = Image.open(self.icon_path) if os.path.exists(self.icon_path) else self._create_simple_tray_image()

            menu = pystray.Menu(
                pystray.MenuItem("Text Expander", None, enabled=False),
                pystray.MenuItem(
                    "Listening: " + ("Active" if self.is_listening else "Paused"),
                    None,
                    enabled=False,
                ),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("Show Window", self.show_window),
                pystray.MenuItem("Toggle Listener", self.toggle_listener_from_tray),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("Exit", self.quit_app),
            )

            self.tray_icon = pystray.Icon("TextExpander", icon_image, "Text Expander", menu)
            threading.Thread(target=self.tray_icon.run, daemon=True).start()
            log("System tray icon initialized")
        except Exception as e:
            log(f"Error setting up system tray: {e}")

    def _create_simple_tray_image(self):
        """Creates a simple PIL image for the system tray if no icon file is found."""
        img = Image.new("RGBA", (64, 64), color=(0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        draw.rectangle([(0, 0), (63, 63)], fill=self.theme["accent"])

        try:
            font = ImageFont.truetype("arial.ttf", 24)
        except IOError:
            font = ImageFont.load_default()

        text = "TE"
        text_width, text_height = draw.textsize(text, font=font)
        position = ((64 - text_width) // 2, (64 - text_height) // 2)
        draw.text(position, text, font=font, fill="white")
        return img

    def toggle_listener_from_tray(self, icon=None, item=None):
        """Toggles the keyboard listener and updates tray menu."""
        self.toggle_listener()
        if self.tray_icon and hasattr(self.tray_icon, "update_menu"):
            menu_items = list(self.tray_icon.menu.items)
            for i, menu_item_obj in enumerate(menu_items):
                if str(menu_item_obj.text).startswith("Listening:"):
                    menu_items[i] = pystray.MenuItem(
                        "Listening: " + ("Active" if self.is_listening else "Paused"),
                        None,
                        enabled=False,
                    )
                    break
            self.tray_icon.menu = pystray.Menu(*menu_items)

    def show_window(self, icon=None, item=None):
        """Restores the main window from minimized state."""
        self.root.deiconify()
        self.root.lift()
        self.root.attributes("-topmost", True)
        self.root.update_idletasks()
        self.root.focus_force()
        self.root.after(100, lambda: self.root.attributes("-topmost", False))

    def quit_app(self, icon=None, item=None):
        """Exits the application gracefully."""
        if hasattr(self, "listener") and self.listener and self.listener.is_alive():
            try:
                self.listener.stop()
            except Exception as e:
                log(f"Error stopping keyboard listener on exit: {e}")

        if hasattr(self, "mouse_listener") and self.mouse_listener and self.mouse_listener.is_alive():
            try:
                self.mouse_listener.stop()
            except Exception as e:
                log(f"Error stopping mouse listener on exit: {e}")

        if self.tray_icon:
            try:
                self.tray_icon.stop()
            except Exception as e:
                log(f"Error stopping tray icon on exit: {e}")

        self.snippet_manager.save_snippets()
        self.config_manager.save_config()
        log("Application exiting")
        self.root.destroy()

    def on_close(self):
        """Handles window close event, minimizing to tray if configured."""
        if (
            self.config_manager.get("minimize_to_tray", True)
            and SYSTRAY_AVAILABLE
            and self.tray_icon
        ):
            self.root.withdraw()
            if sys.platform == "win32" and hasattr(self.tray_icon, "notify"):
                try:
                    self.tray_icon.notify(
                        "Text Expander is still running in the background. Click the tray icon to restore.",
                        "Text Expander",
                    )
                except Exception as e:
                    log(f"Error showing tray notification: {e}")
        else:
            self.quit_app()

    def create_snippets_tab(self):
        """Creates the snippets management tab with improved layout."""
        snippets_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(snippets_frame, text="  Snippets  ")

        toolbar = ttk.Frame(snippets_frame)
        toolbar.pack(fill="x", pady=(0, 10))

        new_btn = ttk.Button(
            toolbar,
            text="New Snippet",
            command=self.create_new_snippet,
            style="Accent.TButton",
        )
        new_btn.pack(side="left", padx=(0, 5))
        self.create_tooltip(new_btn, "Create a new blank snippet (Ctrl+N)")


        search_frame = ttk.Frame(toolbar)
        search_frame.pack(side="right", fill="x", expand=True)

        ttk.Label(search_frame, text="Search:").pack(side="left", padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *args: self.refresh_snippet_list())

        search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        search_entry.pack(side="left", fill="x", expand=True)
        self.create_tooltip(search_entry, "Search snippets by shortcut, content, or description")


        ttk.Label(search_frame, text="Category:").pack(side="left", padx=(10, 5))
        self.category_filter_var = tk.StringVar(value="All Categories")
        self.category_dropdown = ttk.Combobox(
            search_frame,
            textvariable=self.category_filter_var,
            state="readonly",
            width=15,
        )
        self.category_dropdown.pack(side="left", padx=(0, 5))
        self.category_dropdown.bind(
            "<<ComboboxSelected>>", lambda e: self.refresh_snippet_list()
        )
        self.create_tooltip(self.category_dropdown, "Filter snippets by category")


        # Main split view
        paned = ttk.PanedWindow(snippets_frame, orient="horizontal")
        paned.pack(fill="both", expand=True)

        # Left panel - Snippet list with headers
        left_frame = ttk.Frame(paned)
        paned.add(left_frame, weight=1)

        columns = ("shortcut", "category")
        self.snippet_tree = ttk.Treeview(
            left_frame, columns=columns, show="headings", selectmode="browse"
        )

        self.snippet_tree.heading("shortcut", text="Shortcut")
        self.snippet_tree.heading("category", text="Category")

        self.snippet_tree.column("shortcut", width=120, anchor="w")
        self.snippet_tree.column("category", width=120, anchor="w")

        tree_scroll_y = ttk.Scrollbar(
            left_frame, orient="vertical", command=self.snippet_tree.yview
        )
        self.snippet_tree.configure(yscrollcommand=tree_scroll_y.set)

        self.snippet_tree.pack(side="left", fill="both", expand=True)
        tree_scroll_y.pack(side="right", fill="y")

        self.snippet_tree.bind("<<TreeviewSelect>>", self.on_select_snippet)
        self.create_tooltip(self.snippet_tree, "Select a snippet to view or edit its details")


        # Right panel - Editor with improved layout
        right_frame = ttk.Frame(paned)
        paned.add(right_frame, weight=2)

        # Editor section with title
        editor_frame = ttk.LabelFrame(right_frame, text="Snippet Editor", padding=10)
        editor_frame.pack(fill="both", expand=True) # IMPORTANT: Allows editor to expand and push content down

        # Form layout
        form_frame = ttk.Frame(editor_frame)
        form_frame.pack(fill="x", pady=(0, 10))

        form_frame.columnconfigure(0, weight=0)
        form_frame.columnconfigure(1, weight=1)

        row = 0

        ttk.Label(form_frame, text="Shortcut:").grid(
            row=row, column=0, sticky="w", padx=(0, 10), pady=5
        )
        self.shortcut_var = tk.StringVar()
        self.shortcut_entry_widget = ttk.Entry(form_frame, textvariable=self.shortcut_var)
        self.shortcut_entry_widget.grid(row=row, column=1, sticky="ew", pady=5)
        self.create_tooltip(self.shortcut_entry_widget, "The abbreviation that will trigger the snippet expansion")
        row += 1

        ttk.Label(form_frame, text="Category:").grid(
            row=row, column=0, sticky="w", padx=(0, 10), pady=5
        )
        category_frame = ttk.Frame(form_frame)
        category_frame.grid(row=row, column=1, sticky="ew", pady=5)

        self.category_var = tk.StringVar(
            value=self.config_manager.get("default_category", "General")
        )
        self.category_box = ttk.Combobox(category_frame, textvariable=self.category_var)
        self.category_box.pack(side="left", fill="x", expand=True)
        self.create_tooltip(self.category_box, "Assign a category to this snippet for organization")


        self.refresh_categories()
        row += 1

        ttk.Label(form_frame, text="Description:").grid(
            row=row, column=0, sticky="w", padx=(0, 10), pady=5
        )
        self.description_var = tk.StringVar()
        desc_entry = ttk.Entry(form_frame, textvariable=self.description_var)
        desc_entry.grid(row=row, column=1, sticky="ew", pady=5)
        self.create_tooltip(desc_entry, "A short description to remember what this snippet is for")
        row += 1

        ttk.Label(editor_frame, text="Content:").pack(anchor="w")

        editor_content_frame = ttk.Frame(editor_frame)
        editor_content_frame.pack(fill="both", expand=True)

        self.editor = scrolledtext.ScrolledText(
            editor_content_frame,
            wrap="word",
            font=FONTS["monospace"],
            undo=True,
        )
        self.editor.pack(side="left", fill="both", expand=True)
        self.create_tooltip(self.editor, "Enter the full text content of your snippet here. Use placeholders for dynamic content.")


        self.line_numbers = tk.Canvas(
            editor_content_frame,
            width=30,
            bg=self.theme["secondary_bg"],
            highlightthickness=0,
        )
        self.line_numbers.pack(side="left", fill="y")

        self.editor.bind("<KeyRelease>", self.update_line_numbers)
        self.editor.bind("<MouseWheel>", self.update_line_numbers)
        self.editor.bind("<KeyRelease>", self.highlight_placeholders, add="+")

        # Placeholders section with improved organization
        ph_frame = ttk.LabelFrame(right_frame, text="Insert Placeholders")
        ph_frame.pack(fill="x", pady=10) # IMPORTANT: Ensure this takes its space, but doesn't vertically expand unnecessarily

        placeholder_categories = {
            "Date & Time": [
                ("Date", "{date}", "Insert current date (e.g., 05/08/2025)"),
                ("Time", "{time}", "Insert current time (24-hour format, e.g., 14:30)"),
                ("Long Date", "{date_long}", "Insert date with month name (e.g., May 8, 2025)"),
                ("Weekday", "{weekday}", "Insert current day of week (e.g., Thursday)"),
            ],
            "Input & Data": [
                ("Named Input", "{input:name}", "Ask user for input with custom prompt (e.g., {input:Enter your name})"),
                ("Clipboard", "{clipboard}", "Insert current clipboard content"),
                ("Random Number", "{random:1-100}", "Generate random number in a specified range (e.g., {random:1-100})"),
            ],
        }

        ph_notebook = ttk.Notebook(ph_frame)
        ph_notebook.pack(fill="x", padx=5, pady=5)

        for category, placeholders in placeholder_categories.items():
            tab_frame = ttk.Frame(ph_notebook, padding=5)
            ph_notebook.add(tab_frame, text=category)

            for i, (label, placeholder, tooltip) in enumerate(placeholders):
                row, col = divmod(i, 4)

                btn = ttk.Button(
                    tab_frame,
                    text=label,
                    command=lambda p=placeholder: self.insert_placeholder(p),
                    style="Secondary.TButton",
                )
                btn.grid(row=row, column=col, padx=3, pady=3, sticky="ew")
                self.create_tooltip(btn, tooltip)


            for i in range(4):
                tab_frame.columnconfigure(i, weight=1)

        sample_btn = ttk.Button(
            ph_frame, text="Preview Result", command=self.test_snippet
        )
        sample_btn.pack(pady=5)
        self.create_tooltip(sample_btn, "See the final output of your snippet after all placeholders are processed")


        # Action buttons - ensure this frame is packed at the bottom of right_frame
        btn_frame = ttk.Frame(right_frame)
        btn_frame.pack(fill="x", pady=10) # IMPORTANT: Pack this at the end of right_frame, to always be visible


        save_btn = ttk.Button(
            btn_frame, text="Save", command=self.save_snippet, style="Accent.TButton"
        )
        save_btn.pack(side="left", padx=5)
        self.create_tooltip(save_btn, "Save changes to the current snippet")

        delete_btn = ttk.Button(
            btn_frame,
            text="Delete",
            command=self.delete_snippet,
            style="Danger.TButton",
        )
        delete_btn.pack(side="left", padx=5)
        self.create_tooltip(delete_btn, "Delete the selected snippet")


        clear_btn = ttk.Button(btn_frame, text="Clear", command=self.clear_editor)
        clear_btn.pack(side="left", padx=5)
        self.create_tooltip(clear_btn, "Clear all fields in the editor to create a new snippet")


        test_btn = ttk.Button(btn_frame, text="Test", command=self.test_snippet)
        test_btn.pack(side="right", padx=5)
        self.create_tooltip(test_btn, "Preview how the snippet will expand with placeholders processed")


        self.refresh_snippet_list()
        self.update_line_numbers()

    def create_settings_tab(self):
        """Creates the application settings tab."""
        settings_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(settings_frame, text="  Settings  ")

        canvas = tk.Canvas(settings_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(settings_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        general_frame = ttk.LabelFrame(
            scrollable_frame, text="General Settings", padding=10
        )
        general_frame.pack(fill="x", pady=10)

        self.tray_var = tk.BooleanVar(value=self.config_manager.get("minimize_to_tray", True))
        tray_check = ttk.Checkbutton(
            general_frame,
            text="Minimize to system tray",
            variable=self.tray_var,
            command=lambda: self.config_manager.set("minimize_to_tray", self.tray_var.get()),
            state="normal" if SYSTRAY_AVAILABLE else "disabled",
        )
        tray_check.pack(anchor="w", pady=5)
        self.create_tooltip(tray_check, "When checked, the application will minimize to the system tray instead of closing.")

        self.minimized_var = tk.BooleanVar(
            value=self.config_manager.get("start_minimized", False)
        )
        minimized_check = ttk.Checkbutton(
            general_frame,
            text="Start minimized",
            variable=self.minimized_var,
            command=lambda: self.config_manager.set("start_minimized", self.minimized_var.get()),
        )
        minimized_check.pack(anchor="w", pady=5)
        self.create_tooltip(minimized_check, "When checked, the application will start minimized to the system tray on launch.")


        self.tooltips_var = tk.BooleanVar(value=self.config_manager.get("show_tooltips", True))
        tooltips_check = ttk.Checkbutton(
            general_frame,
            text="Show tooltips",
            variable=self.tooltips_var,
            command=self.toggle_tooltips,
        )
        tooltips_check.pack(anchor="w", pady=5)
        self.create_tooltip(tooltips_check, "Toggle visibility of helpful tooltips across the application.")


        default_cat_frame = ttk.Frame(general_frame)
        default_cat_frame.pack(fill="x", pady=5)

        ttk.Label(default_cat_frame, text="Default category:").pack(
            side="left", padx=(0, 10)
        )

        self.default_category_var = tk.StringVar(
            value=self.config_manager.get("default_category", "General")
        )
        self.default_category_box = ttk.Combobox(
            default_cat_frame, textvariable=self.default_category_var, width=20
        )
        self.default_category_box.pack(side="left")
        self.create_tooltip(self.default_category_box, "Set the default category that is automatically assigned to new snippets.")


        self.default_category_var.trace_add(
            "write",
            lambda *args: self.config_manager.set(
                "default_category", self.default_category_var.get()
            ),
        )

        backup_frame = ttk.LabelFrame(
            scrollable_frame, text="Backup Settings", padding=10
        )
        backup_frame.pack(fill="x", pady=10)

        self.auto_backup_var = tk.BooleanVar(value=self.config_manager.get("auto_backup", True))
        auto_backup_check = ttk.Checkbutton(
            backup_frame,
            text="Automatically backup data",
            variable=self.auto_backup_var,
            command=lambda: self.config_manager.set("auto_backup", self.auto_backup_var.get()),
        )
        auto_backup_check.pack(anchor="w", pady=5)
        self.create_tooltip(auto_backup_check, "Enable or disable automatic backups of your snippets and settings.")


        interval_frame = ttk.Frame(backup_frame)
        interval_frame.pack(fill="x", pady=5)

        ttk.Label(interval_frame, text="Backup every:").pack(side="left", padx=(0, 10))

        self.backup_interval_var = tk.StringVar(
            value=str(self.config_manager.get("backup_interval_days", 7))
        )
        interval_spin = ttk.Spinbox(
            interval_frame,
            textvariable=self.backup_interval_var,
            from_=1,
            to=30,
            width=5,
            command=lambda: self.update_backup_interval_from_spinbox()
        )
        interval_spin.pack(side="left")
        self.create_tooltip(interval_spin, "Set the number of days between automatic backups.")


        ttk.Label(interval_frame, text="days").pack(side="left", padx=(5, 0))

        self.backup_interval_var.trace_add("write", lambda *args: self.update_backup_interval_from_var())

        location_frame = ttk.Frame(backup_frame)
        location_frame.pack(fill="x", pady=5)

        ttk.Label(location_frame, text="Backup location:").pack(
            side="left", padx=(0, 10)
        )

        self.backup_location_var = tk.StringVar(
            value=self.config_manager.get("backup_location", APP_DIR)
        )
        location_entry = ttk.Entry(
            location_frame, textvariable=self.backup_location_var
        )
        location_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.create_tooltip(location_entry, "The folder where backup files will be saved.")


        browse_btn = ttk.Button(
            location_frame, text="Browse", command=self.browse_backup_location
        )
        browse_btn.pack(side="left")
        self.create_tooltip(browse_btn, "Choose a different folder for backups.")


        backup_now_btn = ttk.Button(
            backup_frame,
            text="Backup Now",
            command=self.backup_data,
            style="Accent.TButton",
        )
        backup_now_btn.pack(anchor="w", pady=10)
        self.create_tooltip(backup_now_btn, "Manually trigger a backup of your snippets and configuration immediately.")


        last_backup = self.config_manager.get("last_backup")
        last_backup_text = (
            "Never"
            if not last_backup
            else datetime.datetime.fromisoformat(last_backup).strftime("%Y-%m-%d %H:%M")
        )
        self.last_backup_var = tk.StringVar(value=f"Last backup: {last_backup_text}")
        ttk.Label(backup_frame, textvariable=self.last_backup_var).pack(anchor="w")

        theme_frame = ttk.LabelFrame(scrollable_frame, text="Theme", padding=10)
        theme_frame.pack(fill="x", pady=10)

        light_radio = ttk.Radiobutton(
            theme_frame,
            text="Light",
            value="light",
            variable=self.theme_var,
            command=lambda: self.change_theme("light"),
        )
        light_radio.pack(anchor="w", pady=5)
        self.create_tooltip(light_radio, "Switch to the light color scheme for the application.")


        io_frame = ttk.LabelFrame(scrollable_frame, text="Import/Export", padding=10)
        io_frame.pack(fill="x", pady=10)

        btn_frame_io = ttk.Frame(io_frame)
        btn_frame_io.pack(fill="x", pady=5)

        import_btn = ttk.Button(
            btn_frame_io, text="Import Snippets", command=self.import_snippets
        )
        import_btn.pack(side="left", padx=(0, 10))
        self.create_tooltip(import_btn, "Import snippets and settings from a previously exported JSON backup file.")


        export_btn = ttk.Button(
            btn_frame_io, text="Export Snippets", command=self.export_snippets
        )
        export_btn.pack(side="left")
        self.create_tooltip(export_btn, "Export all your snippets and current settings to a JSON file for backup or transfer.")

    def create_help_tab(self):
        """Creates the help tab with guides and troubleshooting."""
        help_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(help_frame, text="  Help  ")

        canvas = tk.Canvas(help_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(help_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        welcome_frame = ttk.Frame(scrollable_frame)
        welcome_frame.pack(fill="x", pady=10)

        ttk.Label(
            welcome_frame, text="Welcome to Text Expander", font=FONTS["heading"]
        ).pack(anchor="w")

        ttk.Label(
            welcome_frame,
            text="Text Expander helps you save time by replacing short abbreviations with longer text snippets.",
            wraplength=600,
        ).pack(anchor="w", pady=5)

        start_frame = ttk.LabelFrame(
            scrollable_frame, text="Getting Started", padding=10
        )
        start_frame.pack(fill="x", pady=10)

        steps = [
            "1. Create a new snippet by clicking the 'New Snippet' button.",
            "2. Enter a shortcut (e.g., '/sig' for a signature) and your desired text.",
            "3. Use placeholders like {date} to insert dynamic content.",
            "4. Click 'Save' to store your snippet.",
            "5. Type your shortcut in any application to automatically expand it.",
        ]

        for step in steps:
            ttk.Label(start_frame, text=step, wraplength=600).pack(anchor="w", pady=2)

        ph_frame = ttk.LabelFrame(
            scrollable_frame, text="Placeholders Guide", padding=10
        )
        ph_frame.pack(fill="x", pady=10)

        ttk.Label(
            ph_frame,
            text="Placeholders are special codes that get replaced with dynamic content when your snippet expands.",
            wraplength=600,
        ).pack(anchor="w", pady=5)

        ph_table_frame = ttk.Frame(ph_frame)
        ph_table_frame.pack(fill="x", pady=5)

        ttk.Label(
            ph_table_frame, text="Placeholder", font=FONTS["subheading"], width=20
        ).grid(row=0, column=0, sticky="w", padx=5, pady=5)
        ttk.Label(
            ph_table_frame, text="Description", font=FONTS["subheading"], width=30
        ).grid(row=0, column=1, sticky="w", padx=5, pady=5)
        ttk.Label(
            ph_table_frame, text="Example", font=FONTS["subheading"], width=20
        ).grid(row=0, column=2, sticky="w", padx=5, pady=5)

        ttk.Separator(ph_table_frame, orient="horizontal").grid(
            row=1, column=0, columnspan=3, sticky="ew", pady=5
        )

        placeholders = [
            ("{date}", "Current date (e.g., MM/DD/YYYY)", "05/08/2025"),
            ("{time}", "Current time (24h format)", "14:30"),
            ("{date_long}", "Date with month name (e.g., May 8, 2025)", "May 8, 2025"),
            ("{weekday}", "Current day of the week", "Thursday"),
            (
                "{input:prompt}",
                "Asks user for input with custom prompt",
                "{input:Enter your name}",
            ),
            (
                "{clipboard}",
                "Inserts current clipboard content",
                "Any text you've copied",
            ),
            ("{random:min-max}", "Random number in specified range", "{random:1-100}"),
        ]

        for i, (placeholder, desc, example) in enumerate(placeholders):
            row = i + 2
            ttk.Label(ph_table_frame, text=placeholder, font=FONTS["monospace"]).grid(
                row=row, column=0, sticky="w", padx=5, pady=5
            )
            ttk.Label(ph_table_frame, text=desc, wraplength=250).grid(
                row=row, column=1, sticky="w", padx=5, pady=5
            )
            ttk.Label(ph_table_frame, text=example).grid(
                row=row, column=2, sticky="w", padx=5, pady=5
            )

        tips_frame = ttk.LabelFrame(scrollable_frame, text="Tips & Tricks", padding=10)
        tips_frame.pack(fill="x", pady=10)

        tips = [
            (
                "Use categories to organize your snippets",
                "Group related snippets together for easy management.",
            ),
            (
                "Test before saving",
                "Use the 'Test' button to preview exactly what your snippet will output.",
            ),
            (
                "Use descriptive shortcuts",
                "Choose shortcuts that are easy to remember but unlikely to be typed accidentally.",
            ),
            (
                "Named inputs",
                "Use {input:Name?} to show a specific prompt when asking for input.",
            ),
            ("Backup regularly", "Export your snippets occasionally as a backup."),
        ]

        for i, (title, desc) in enumerate(tips):
            tip_frame = ttk.Frame(tips_frame)
            tip_frame.pack(fill="x", pady=5)

            ttk.Label(tip_frame, text=title, font=FONTS["subheading"]).pack(anchor="w")
            ttk.Label(tip_frame, text=desc, wraplength=600).pack(
                anchor="w", padx=(20, 0)
            )

        trouble_frame = ttk.LabelFrame(
            scrollable_frame, text="Troubleshooting", padding=10
        )
        trouble_frame.pack(fill="x", pady=10)

        issues = [
            (
                "Snippets don't expand",
                [
                    "Make sure the keyboard listener is active (status shows 'Listening')",
                    "Check that you're typing the exact shortcut (case-sensitive)",
                    "Ensure pynput module is installed.",
                ],
            ),
            (
                "Missing placeholder data",
                [
                    "Clipboard placeholder requires the pyperclip module",
                    "Input placeholders need pynput module",
                ],
            ),
            (
                "Application crashes",
                [
                    f"Check the log file at: {APP_DIR}",
                    "Try resetting to default settings if problems persist",
                ],
            ),
        ]

        for i, (issue, solutions) in enumerate(issues):
            issue_frame = ttk.Frame(trouble_frame)
            issue_frame.pack(fill="x", pady=5)

            ttk.Label(issue_frame, text=issue, font=FONTS["subheading"]).pack(
                anchor="w"
            )

            for solution in solutions:
                ttk.Label(issue_frame, text="• " + solution, wraplength=600).pack(
                    anchor="w", padx=(20, 0)
                )

        support_frame = ttk.LabelFrame(
            scrollable_frame, text="Support & Resources", padding=10
        )
        support_frame.pack(fill="x", pady=10)

        ttk.Label(
            support_frame,
            text="For more help, check out these resources:",
            wraplength=600,
        ).pack(anchor="w", pady=5)

        resources = [
            ("About", "View application information"),
            ("Check for Updates", "Make sure you have the latest version"),
        ]

        for title, desc in resources:
            resource_frame = ttk.Frame(support_frame)
            resource_frame.pack(fill="x", pady=2)

            resource_btn = ttk.Button(
                resource_frame,
                text=title,
                command=lambda t=title: self.open_resource(t),
                style="Link.TButton",
            )
            resource_btn.pack(side="left")
            self.create_tooltip(resource_btn, desc)


            ttk.Label(resource_frame, text=desc).pack(side="left", padx=(10, 0))

    def refresh_snippet_list(self):
        """Refreshes the snippet tree view with current data, applying filters."""
        for item in self.snippet_tree.get_children():
            self.snippet_tree.delete(item)

        search = self.search_var.get().lower()
        category_filter = self.category_filter_var.get()

        sorted_shortcuts = sorted(self.snippet_manager.get_all_snippets().keys())

        for shortcut in sorted_shortcuts:
            snippet_data = self.snippet_manager.get_snippet(shortcut)

            if search:
                if (
                    search not in shortcut.lower()
                    and search not in snippet_data["text"].lower()
                    and search not in snippet_data.get("description", "").lower()
                ):
                    continue

            if (
                category_filter != "All Categories"
                and snippet_data.get("category") != category_filter
            ):
                continue

            self.snippet_tree.insert(
                "",
                "end",
                values=(shortcut, snippet_data.get("category", "General")),
                tags=("snippet",),
            )

        count = len(self.snippet_tree.get_children())
        if count == 0 and (search or category_filter != "All Categories"):
            self.status_var.set(f"No snippets match the current filter")
        else:
            self.status_var.set(f"Showing {count} snippets")

    def refresh_categories(self):
        """Refreshes the category lists in all relevant dropdowns."""
        categories = self.snippet_manager.get_all_categories()
        self.category_dropdown["values"] = ["All Categories"] + categories
        self.category_box["values"] = categories
        if hasattr(self, "default_category_box"):
            self.default_category_box["values"] = categories

    def update_line_numbers(self, event=None):
        """Updates line numbers in the scrolled text editor."""
        if not hasattr(self, "editor") or not hasattr(self, "line_numbers"):
            return

        text = self.editor.get("1.0", "end-1c")
        line_count = text.count("\n") + 1 if text else 1

        self.line_numbers.delete("all")

        first_line_idx = self.editor.index("@0,0")
        first_line = int(first_line_idx.split(".")[0])

        last_line_idx = self.editor.index(f"@0,{self.editor.winfo_height()}")
        last_line = int(last_line_idx.split(".")[0])

        for i in range(first_line, min(last_line + 2, line_count + 1)):
            dline = self.editor.dlineinfo(f"{i}.0")
            if dline:
                y = dline[1]
                self.line_numbers.create_text(
                    15, y, text=str(i), anchor="center", fill=self.theme["fg"]
                )

    def highlight_placeholders(self, event=None):
        """Highlights known and unknown placeholders in the editor."""
        if not hasattr(self, "editor"):
            return

        for tag in ["placeholder", "input_placeholder", "unknown_placeholder"]:
            self.editor.tag_remove(tag, "1.0", "end")

        text_content = self.editor.get("1.0", "end-1c")

        placeholder_pattern = r"\{[^}]*\}"

        for match in re.finditer(placeholder_pattern, text_content):
            full_placeholder = match.group(0)
            placeholder_content = match.group(0)[1:-1]
            start_idx_char = match.start()
            end_idx_char = match.end()

            start_pos = f"1.0 + {start_idx_char} chars"
            end_pos = f"1.0 + {end_idx_char} chars"

            if full_placeholder.startswith("{input:"):
                self.editor.tag_add("input_placeholder", start_pos, end_pos)
            elif full_placeholder in [
                "{date}", "{time}", "{date_long}", "{weekday}", "{clipboard}",
            ] or full_placeholder.startswith("{random:"):
                self.editor.tag_add("placeholder", start_pos, end_pos)
            else:
                self.editor.tag_add("unknown_placeholder", start_pos, end_pos)

    def insert_placeholder(self, placeholder):
        """Inserts a placeholder at the current cursor position."""
        if not hasattr(self, "editor"):
            return

        if placeholder == "{input:name}":
            prompt = self.get_simple_input(
                "Define User Input Prompt",
                "What question should the user be asked when this snippet runs?\n(e.g., 'Enter client name:' or 'Project due date?')"
            )
            if prompt:
                prompt = prompt.strip()
                if not prompt.endswith("?") and not prompt.endswith(":") and not prompt.endswith("."):
                    placeholder = f"{{input:{prompt}?}}"
                else:
                    placeholder = f"{{input:{prompt}}}"
            else:
                placeholder = "{input:?}"
        elif placeholder == "{random:1-100}":
            range_input = self.get_simple_input(
                "Random Number Range", "Enter the range (e.g., 1-100 or 50-200):"
            )
            if range_input and "-" in range_input:
                try:
                    min_val, max_val = map(str.strip, range_input.split("-", 1))
                    int(min_val)
                    int(max_val)
                    placeholder = f"{{random:{min_val}-{max_val}}}"
                except ValueError:
                    placeholder = "{random:1-100}"
            else:
                placeholder = "{random:1-100}"

        self.editor.insert("insert", placeholder)
        self.highlight_placeholders()

    def get_simple_input(self, title, prompt):
        """
        Shows a simple, modal input dialog.
        This is for general app inputs (e.g., defining placeholder prompt).
        """
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("400x150")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.attributes("-topmost", True)
        dialog.focus_force()

        dialog.configure(bg=self.theme["bg"])

        frame = ttk.Frame(dialog, padding=15)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text=prompt).pack(pady=(0, 10))

        input_var = tk.StringVar()
        entry = ttk.Entry(frame, textvariable=input_var)
        entry.pack(fill="x", pady=(0, 20))
        entry.focus_set()

        result = [None]

        def on_ok():
            result[0] = input_var.get()
            dialog.destroy()

        def on_cancel():
            dialog.destroy()

        ttk.Button(frame, text="OK", command=on_ok, style="Accent.TButton").pack(side="right", padx=5)
        ttk.Button(frame, text="Cancel", command=on_cancel).pack(side="right", padx=5)

        entry.bind("<Return>", lambda e: on_ok())
        dialog.bind("<Escape>", lambda e: on_cancel())

        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        x = (dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (dialog.winfo_screenheight() // 2) - (height // 2)
        dialog.geometry(f"{width}x{height}+{x}+{y}")

        self.root.wait_window(dialog)
        return result[0]

    def on_select_snippet(self, event):
        """Handles selection of a snippet in the Treeview."""
        selection = self.snippet_tree.selection()
        if not selection:
            return

        item = selection[0]
        shortcut = self.snippet_tree.item(item, "values")[0]

        snippet_data = self.snippet_manager.get_snippet(shortcut)
        if not snippet_data:
            return

        self.shortcut_var.set(shortcut)
        self.category_var.set(snippet_data.get("category", "General"))
        self.description_var.set(snippet_data.get("description", ""))

        self.editor.delete("1.0", tk.END)
        self.editor.insert("1.0", snippet_data["text"])

        self.highlight_placeholders()
        self.update_line_numbers()

    def create_new_snippet(self):
        """Clears the editor to create a new snippet."""
        self.clear_editor()
        if hasattr(self, "shortcut_entry_widget"):
            self.shortcut_entry_widget.focus_set()

    def save_snippet(self):
        """Saves the current snippet from the editor."""
        shortcut = self.shortcut_var.get().strip()
        text = self.editor.get("1.0", "end-1c").strip()
        category = self.category_var.get().strip() or "General"
        description = self.description_var.get().strip()

        if not shortcut:
            self.show_error("Validation Error", "Shortcut cannot be empty")
            return
        if not text:
            self.show_error("Validation Error", "Content cannot be empty")
            return

        invalid_placeholders = self.validate_placeholders(text)
        if invalid_placeholders:
            if not self.show_confirm(
                "Invalid Placeholders",
                f"The following placeholders appear invalid: {', '.join(invalid_placeholders)}\n\nSave anyway?",
            ):
                return

        if self.snippet_manager.add_update_snippet(shortcut, text, category, description):
            self.status_var.set(f"Saved: {shortcut}")
            if shortcut not in self.snippet_history:
                self.snippet_history.append(shortcut)
                if len(self.snippet_history) > 10:
                    self.snippet_history.pop(0)

            self.refresh_snippet_list()
            self.refresh_categories()
            self.show_success_indicator()
        else:
            self.show_error("Save Error", "Failed to save snippet")

    def validate_placeholders(self, text):
        """Validates placeholder syntax within the text."""
        invalid = []
        open_count = text.count("{")
        close_count = text.count("}")

        if open_count != close_count:
            invalid.append(f"Unmatched braces: {open_count} opening, {close_count} closing")

        placeholder_pattern = r"\{([^}]*)\}"
        for match in re.finditer(placeholder_pattern, text):
            placeholder_content = match.group(1)
            full_placeholder = match.group(0)

            if placeholder_content in ["date", "time", "date_long", "weekday", "clipboard"]:
                continue
            elif placeholder_content.startswith("input:"):
                prompt = placeholder_content[6:]
                if not prompt:
                    invalid.append(f"{full_placeholder} (missing prompt)")
            elif placeholder_content.startswith("random:"):
                range_str = placeholder_content[7:]
                if not range_str or "-" not in range_str:
                    invalid.append(f"{full_placeholder} (invalid range format, e.g., 1-100)")
                else:
                    try:
                        min_val_str, max_val_str = range_str.split("-", 1)
                        min_val = int(min_val_str.strip())
                        max_val = int(max_val_str.strip())
                        if min_val >= max_val:
                            invalid.append(f"{full_placeholder} (min value >= max value)")
                    except ValueError:
                        invalid.append(f"{full_placeholder} (range values not numeric)")
            else:
                invalid.append(f"{full_placeholder} (unknown placeholder)")
        return invalid

    def delete_snippet(self):
        """Deletes the selected snippet with confirmation."""
        selection = self.snippet_tree.selection()
        if not selection:
            # Fallback to the shortcut field if nothing is selected in the tree
            shortcut = self.shortcut_var.get().strip()
        else:
            item = selection[0]
            shortcut = self.snippet_tree.item(item, "values")[0]

        if not shortcut or not self.snippet_manager.get_snippet(shortcut):
            self.show_info("No Selection", "Please select a snippet to delete or ensure shortcut field is filled.")
            return

        if self.show_confirm("Confirm Delete", f"Delete snippet '{shortcut}'?"):
            if self.snippet_manager.delete_snippet(shortcut):
                if shortcut in self.snippet_history:
                    self.snippet_history.remove(shortcut)
                self.status_var.set(f"Deleted: {shortcut}")
                self.refresh_snippet_list()
                self.refresh_categories()
                self.clear_editor()
            else:
                self.show_error("Delete Error", "Failed to delete snippet.")

    def clear_editor(self):
        """Clears all fields in the snippet editor."""
        self.shortcut_var.set("")
        self.category_var.set(self.config_manager.get("default_category", "General"))
        self.description_var.set("")
        self.editor.delete("1.0", tk.END)
        if hasattr(self, "snippet_tree"):
            for item in self.snippet_tree.selection():
                self.snippet_tree.selection_remove(item)

    def test_snippet(self):
        """Tests the current snippet content and shows a preview."""
        text = self.editor.get("1.0", "end-1c")
        if not text:
            self.show_info("Empty Snippet", "There is no content to preview.")
            return

        processed = self.process_placeholders(text)

        preview = tk.Toplevel(self.root)
        preview.title("Snippet Preview")
        preview.geometry("600x400")
        preview.transient(self.root)
        preview.grab_set()
        preview.configure(bg=self.theme["bg"])

        preview_notebook = ttk.Notebook(preview)
        preview_notebook.pack(fill="both", expand=True, padx=10, pady=10)

        preview_tab = ttk.Frame(preview_notebook, padding=10)
        preview_notebook.add(preview_tab, text="Preview")

        ttk.Label(preview_tab, text="This is how your snippet will appear when used:").pack(padx=10, pady=5, anchor="w")

        preview_frame = ttk.Frame(preview_tab, relief="solid", borderwidth=1)
        preview_frame.pack(fill="both", expand=True, padx=10, pady=10)

        result_text_widget = scrolledtext.ScrolledText(
            preview_frame,
            wrap="word",
            font=("Segoe UI", 10),
            bg=self.theme["bg"],
            fg=self.theme["fg"],
        )
        result_text_widget.pack(fill="both", expand=True, padx=10, pady=10)
        result_text_widget.insert("1.0", processed)
        result_text_widget.configure(state="disabled")

        compare_tab = ttk.Frame(preview_notebook, padding=10)
        preview_notebook.add(compare_tab, text="Compare")

        compare_pane = ttk.PanedWindow(compare_tab, orient="horizontal")
        compare_pane.pack(fill="both", expand=True)

        before_frame = ttk.LabelFrame(compare_pane, text="Original")
        compare_pane.add(before_frame, weight=1)

        before_text_widget = scrolledtext.ScrolledText(
            before_frame,
            wrap="word",
            font=("Consolas", 10),
            bg=self.theme["bg"],
            fg=self.theme["fg"],
        )
        before_text_widget.pack(fill="both", expand=True, padx=5, pady=5)
        before_text_widget.insert("1.0", text)
        before_text_widget.configure(state="disabled")

        before_text_widget.tag_configure(
            "placeholder_preview", foreground=self.theme["accent"], font=("Consolas", 10, "bold")
        )
        before_text_widget.tag_configure(
            "input_placeholder_preview", foreground=self.theme["success"], font=("Consolas", 10, "bold")
        )
        before_text_widget.tag_configure(
            "unknown_placeholder_preview", foreground=self.theme["error"], font=("Consolas", 10, "italic")
        )

        content_before = before_text_widget.get("1.0", "end-1c")
        placeholder_pattern = r"\{[^}]*\}"
        for match in re.finditer(placeholder_pattern, content_before):
            start_idx_char = match.start()
            end_idx_char = match.end()
            start_pos = f"1.0 + {start_idx_char} chars"
            end_pos = f"1.0 + {end_idx_char} chars"
            full_placeholder = match.group(0)

            if full_placeholder.startswith("{input:"):
                before_text_widget.tag_add("input_placeholder_preview", start_pos, end_pos)
            elif full_placeholder in ["{date}", "{time}", "{date_long}", "{weekday}", "{clipboard}"] or full_placeholder.startswith("{random:"):
                before_text_widget.tag_add("placeholder_preview", start_pos, end_pos)
            else:
                before_text_widget.tag_add("unknown_placeholder_preview", start_pos, end_pos)

        after_frame = ttk.LabelFrame(compare_pane, text="Expanded")
        compare_pane.add(after_frame, weight=1)

        after_text_widget = scrolledtext.ScrolledText(
            after_frame,
            wrap="word",
            font=("Segoe UI", 10),
            bg=self.theme["bg"],
            fg=self.theme["fg"],
        )
        after_text_widget.pack(fill="both", expand=True, padx=5, pady=5)
        after_text_widget.insert("1.0", processed)
        after_text_widget.configure(state="disabled")

        btn_frame_preview = ttk.Frame(preview)
        btn_frame_preview.pack(fill="x", padx=10, pady=10)

        if CLIPBOARD_AVAILABLE:
            copy_to_clipboard_btn = ttk.Button(
                btn_frame_preview,
                text="Copy to Clipboard",
                command=lambda: [self.copy_to_clipboard(processed), preview.destroy()],
                style="Accent.TButton",
            )
            copy_to_clipboard_btn.pack(side="left", padx=5)
            self.create_tooltip(copy_to_clipboard_btn, "Copy the expanded snippet to your clipboard.")


        close_btn = ttk.Button(btn_frame_preview, text="Close", command=preview.destroy)
        close_btn.pack(side="right", padx=5)
        self.create_tooltip(close_btn, "Close the preview window.")


        preview.focus_set()

        preview.update_idletasks()
        width = preview.winfo_width()
        height = preview.winfo_height()
        x = (preview.winfo_screenwidth() // 2) - (width // 2)
        y = (preview.winfo_screenheight() // 2) - (height // 2)
        preview.geometry(f"+{x}+{y}")

    def copy_to_clipboard(self, text):
        """Copies text to the system clipboard."""
        if CLIPBOARD_AVAILABLE:
            try:
                pyperclip.copy(text)
                self.status_var.set("Copied to clipboard")
                self.show_success_indicator()
                return True
            except Exception as e:
                log(f"Error copying to clipboard: {e}")
                self.show_error("Clipboard Error", f"Could not copy to clipboard: {e}")
        else:
            self.show_error("Not Available", "Clipboard functionality is not available.")
        return False

    def show_success_indicator(self):
        """Shows a brief visual success indicator in the status bar."""
        if not hasattr(self, "status_indicator"):
            return

        self.status_indicator.config(foreground=self.theme["success"])

        def reset_color():
            if hasattr(self, "status_indicator"):
                current_indicator_color = self.theme["warning"]
                if self.is_listening and KEYBOARD_AVAILABLE:
                    current_indicator_color = self.theme["success"]
                elif not KEYBOARD_AVAILABLE:
                    current_indicator_color = self.theme["error"]
                self.status_indicator.config(foreground=current_indicator_color)

        self.root.after(1500, reset_color)

    def process_placeholders(self, text):
        """Processes placeholders in a given text string."""
        now = datetime.datetime.now()
        # Locale-aware date/time formatting could be added here if needed,
        # but for simplicity and removal of language setting, fixed formats are used.
        date_format_str = "%m/%d/%Y" # e.g., 05/20/2025
        month_format_str = "%B %d, %Y" # e.g., May 20, 2025
        weekdays = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        current_weekday = weekdays[now.weekday()]

        def get_clipboard_content():
            if CLIPBOARD_AVAILABLE:
                try:
                    return pyperclip.paste() or "" # Return empty string for empty clipboard
                except Exception as e:
                    log(f"Error accessing clipboard: {e}")
                    return "[Clipboard Error]"
            return "[Clipboard N/A]"

        result = text
        result = result.replace("{date}", now.strftime(date_format_str))
        result = result.replace("{time}", now.strftime("%H:%M"))
        result = result.replace("{date_long}", now.strftime(month_format_str))
        result = result.replace("{weekday}", current_weekday)

        if "{clipboard}" in result:
            result = result.replace("{clipboard}", get_clipboard_content())

        def replace_random(match_obj):
            try:
                import random
                min_val_str, max_val_str = match_obj.group(1).split("-", 1)
                min_val = int(min_val_str.strip())
                max_val = int(max_val_str.strip())
                if min_val <= max_val:
                    return str(random.randint(min_val, max_val))
                else:
                    return f"[Random Error: min({min_val}) > max({max_val})]"
            except Exception as e:
                log(f"Error processing random placeholder {match_obj.group(0)}: {e}")
                return f"[Random Error: {match_obj.group(0)}]"

        result = re.sub(r"\{random:([0-9]+\s*-\s*[0-9]+)\}", replace_random, result)

        input_pattern = r"\{input:([^}]*)\}"
        input_matches = list(re.finditer(input_pattern, result))
        input_values = []

        # Collect all input values first
        for match in input_matches:
            full_placeholder = match.group(0)
            # Use _get_expansion_input for runtime dialogs
            user_input = self._get_expansion_input(prompt_text=match.group(1))
            input_values.append((full_placeholder, user_input))

        # Then replace them
        for original_ph, new_value in input_values:
            result = result.replace(original_ph, new_value, 1)

        return result

    def _get_expansion_input(self, prompt_text=""):
        """
        Shows a streamlined input dialog specifically for snippet expansion.
        Ensures it appears on top and grabs focus, positioned near the mouse.
        """
        if not prompt_text:
            prompt_text = "Enter text:"
        elif not prompt_text.endswith("?") and not prompt_text.endswith(":"):
            prompt_text += ":"

        input_dialog = tk.Toplevel(self.root)
        input_dialog.title("Input Required")
        input_dialog.resizable(False, False)

        # Essential for always-on-top and focus
        input_dialog.attributes("-topmost", True)
        input_dialog.transient(self.root) # Link to main window
        input_dialog.grab_set() # Make modal
        # No focus_force here initially; it's done after geometry is set.

        # Bring to front on Windows/macOS if it exists
        # This is often needed with overrideredirect
        if sys.platform == "win32" or sys.platform == "darwin":
             input_dialog.lift()

        # Optional: remove window decorations for a cleaner look
        input_dialog.overrideredirect(True)

        # Apply theme colors
        bg_color = self.theme["bg"]
        fg_color = self.theme["fg"]
        accent_color = self.theme["accent"]
        border_color = self.theme["border"]
        button_text_color = self.theme["button_text"]

        input_dialog.configure(bg=bg_color)

        main_dialog_frame = tk.Frame(
            input_dialog,
            bg=bg_color,
            highlightbackground=border_color,
            highlightthickness=1,
        )
        main_dialog_frame.pack(fill="both", expand=True, padx=1, pady=1)

        # Custom title bar for dragging when overrideredirect is True
        title_bar_frame = tk.Frame(
            main_dialog_frame, bg=accent_color, height=30
        )
        title_bar_frame.pack(fill="x")

        tk.Label(
            title_bar_frame,
            text="Input Required",
            bg=accent_color,
            fg=button_text_color,
            font=FONTS["subheading"],
        ).pack(side="left", padx=10, pady=4)

        close_button_label = tk.Label(
            title_bar_frame,
            text="×", # Unicode multiplication sign for a close 'x'
            bg=accent_color,
            fg=button_text_color,
            font=("Arial", 16, "bold"),
        )
        close_button_label.pack(side="right", padx=8, pady=0)

        content_dialog_frame = tk.Frame(
            main_dialog_frame, bg=bg_color, padx=15, pady=15
        )
        content_dialog_frame.pack(fill="both", expand=True)

        tk.Label(
            content_dialog_frame,
            text=prompt_text, # Use prompt_text directly
            bg=bg_color,
            fg=fg_color,
            font=FONTS["body"],
            wraplength=350,
            justify=tk.LEFT,
        ).pack(anchor="w", pady=(0, 10))

        input_var = tk.StringVar()
        entry_field = tk.Entry(
            content_dialog_frame,
            textvariable=input_var,
            bg=self.theme["secondary_bg"],
            fg=fg_color,
            font=FONTS["body"],
            relief="solid",
            borderwidth=1,
            highlightbackground=border_color,
            highlightthickness=1,
            highlightcolor=accent_color,
            insertbackground=fg_color,
        )
        entry_field.pack(fill="x", pady=(0, 15))
        # Focus is set AFTER positioning for better reliability

        buttons_dialog_frame = tk.Frame(content_dialog_frame, bg=bg_color)
        buttons_dialog_frame.pack(fill="x", anchor="e")

        user_input_result = [None] # Use a list to hold the result

        def on_dialog_ok():
            user_input_result[0] = input_var.get()
            input_dialog.destroy()

        def on_dialog_cancel():
            user_input_result[0] = "" # Return empty string on cancel
            input_dialog.destroy()

        close_button_label.bind("<Button-1>", lambda e: on_dialog_cancel()) # Close acts as cancel

        ok_dialog_button = tk.Button(
            buttons_dialog_frame,
            text="OK",
            bg=accent_color,
            fg=button_text_color,
            font=FONTS["button"],
            relief="flat",
            padx=15,
            pady=3,
            borderwidth=0,
            activebackground=self.theme["hover"],
            activeforeground=button_text_color,
            command=on_dialog_ok,
        )
        ok_dialog_button.pack(side="right", padx=(5, 0))

        cancel_dialog_button = tk.Button(
            buttons_dialog_frame,
            text="Cancel",
            bg=self.theme["secondary_bg"],
            fg=fg_color,
            font=FONTS["button"],
            relief="flat",
            padx=15,
            pady=3,
            borderwidth=0,
            activebackground=self.theme["border"],
            activeforeground=fg_color,
            command=on_dialog_cancel,
        )
        cancel_dialog_button.pack(side="right", padx=(0, 5))

        # Make dialog draggable via title bar
        make_draggable(title_bar_frame, input_dialog)

        # Position dialog near cursor or center if no cursor info
        input_dialog.update_idletasks() # Update to get actual size
        width = input_dialog.winfo_width()
        height = input_dialog.winfo_height()

        x_pos, y_pos = self.last_mouse_pos # Use last known mouse position for initial placement
        if x_pos == 0 and y_pos == 0: # Fallback to center screen if no mouse data
            x_pos = (input_dialog.winfo_screenwidth() // 2) - (width // 2)
            y_pos = (input_dialog.winfo_screenheight() // 2) - (height // 2)
        else: # Position near mouse, ensuring it's on screen
            # Add small offset from mouse position to avoid immediately clicking the dialog
            offset_x = 20
            offset_y = 20
            x_pos = x_pos + offset_x
            y_pos = y_pos + offset_y

            # Ensure it stays within screen bounds
            screen_width = input_dialog.winfo_screenwidth()
            screen_height = input_dialog.winfo_screenheight()

            if x_pos + width > screen_width:
                x_pos = screen_width - width
            if y_pos + height > screen_height:
                y_pos = screen_height - height
            if x_pos < 0: x_pos = 0
            if y_pos < 0: y_pos = 0


        input_dialog.geometry(f"+{x_pos}+{y_pos}")

        # Set focus after positioning is final
        entry_field.focus_set()

        # Bind Enter and Escape keys for dialog interaction
        entry_field.bind("<Return>", lambda e: on_dialog_ok())
        input_dialog.bind("<Escape>", lambda e: on_dialog_cancel())

        self.root.wait_window(input_dialog) # Wait for dialog to close
        return user_input_result[0] if user_input_result[0] is not None else ""


    def start_listener(self):
        """Starts the keyboard listener thread."""
        if not KEYBOARD_AVAILABLE:
            self.status_var.set("Keyboard module not available")
            self.status_indicator.config(foreground=self.theme["error"])
            return

        try:
            if self.listener and self.listener.is_alive():
                self.listener.stop()

            self.listener = keyboard.Listener(on_press=self.on_key_press)
            self.listener.start()
            self.is_listening = True
            log("Keyboard listener started")
            self.status_var.set("Listening for shortcuts")
            self.status_indicator.config(foreground=self.theme["success"])
            if hasattr(self, "toggle_btn"):
                self.toggle_btn.config(text="Pause Listening")
        except Exception as e:
            log(f"Error starting keyboard listener: {e}")
            self.status_var.set("Keyboard error")
            if hasattr(self, "status_indicator"):
                self.status_indicator.config(foreground=self.theme["error"])

    def toggle_listener(self):
        """Toggles the keyboard listener active state."""
        if not KEYBOARD_AVAILABLE:
            self.show_info(
                "Not Available",
                "Keyboard module is not available. Please install the required dependencies (pynput)."
            )
            return

        if self.is_listening:
            self.is_listening = False
            if self.listener and self.listener.is_alive():
                log("Keyboard listener paused (flag set).")
            self.status_var.set("Paused")
            self.toggle_btn.config(text="Resume Listening")
            self.status_indicator.config(foreground=self.theme["warning"])
        else:
            self.is_listening = True
            if not (self.listener and self.listener.is_alive()):
                self.start_listener()
            else:
                self.status_var.set("Listening for shortcuts")
                self.toggle_btn.config(text="Pause Listening")
                self.status_indicator.config(foreground=self.theme["success"])

    def start_mouse_tracking(self):
        """Starts a mouse listener to track cursor position."""
        if not KEYBOARD_AVAILABLE:
            return

        try:
            if self.mouse_listener and self.mouse_listener.is_alive():
                self.mouse_listener.stop()

            self.mouse_listener = mouse.Listener(on_move=self.on_mouse_move)
            self.mouse_listener.start()
            log("Mouse tracking started")
        except Exception as e:
            log(f"Error starting mouse tracking: {e}")

    def on_mouse_move(self, x, y):
        """Updates the last known mouse position."""
        self.last_mouse_pos = (x, y)

    def on_key_press(self, key):
        """
        Processes key press events to detect and expand shortcuts.
        Improved to handle complex keys and select/delete shortcuts robustly.
        """
        if not self.is_listening:
            return

        try:
            char_pressed = None
            if hasattr(key, 'char') and key.char is not None:
                char_pressed = key.char # Regular character
            elif key == keyboard.Key.space:
                char_pressed = " " # Space character
            elif key == keyboard.Key.tab: # Tab should reset input and not be processed as part of shortcut
                self.current_input = ""
                return
            elif key == keyboard.Key.enter: # Enter should reset input and not be processed
                self.current_input = ""
                return
            elif key == keyboard.Key.esc: # Escape should reset input
                self.current_input = ""
                return

            # For backspace, delete last char
            if key == keyboard.Key.backspace:
                self.current_input = self.current_input[:-1]
            elif char_pressed: # Append only if it's a character or space
                self.current_input += char_pressed

            # Keep buffer a reasonable size
            max_buffer = 60 # Max length of a shortcut plus some context
            if len(self.current_input) > max_buffer:
                self.current_input = self.current_input[-max_buffer:]

            current_snippets = self.snippet_manager.get_all_snippets()

            # Iterate through sorted snippet keys (longest first)
            # This handles cases where one shortcut is a prefix of another (e.g., "sig" and "signow").
            possible_shortcuts = sorted(current_snippets.keys(), key=len, reverse=True)

            for shortcut in possible_shortcuts:
                # Check if the current input ends with the shortcut
                if self.current_input.endswith(shortcut):
                    is_command_like = not shortcut[0].isalnum() if shortcut and shortcut[0] else False

                    should_expand = False
                    if is_command_like:
                        # If it's a command-like shortcut (e.g., starts with '/'), expand immediately
                        should_expand = True
                    else:
                        # For regular word-like shortcuts, check for a word boundary before it
                        # E.g., "mytext" should not expand if typed as "somemytext"
                        # The character before the shortcut must not be alphanumeric.
                        idx = self.current_input.rfind(shortcut)
                        if idx == 0: # Shortcut is at the very beginning of the buffer
                            should_expand = True
                        elif idx > 0 and not self.current_input[idx - 1].isalnum():
                            should_expand = True

                    if should_expand:
                        snippet_data = current_snippets[shortcut]
                        text_to_expand = snippet_data["text"]

                        # Schedule placeholder processing and text replacement on the main Tkinter thread
                        # This is crucial for UI interactions (like input dialogs) to happen on the main thread.
                        self.root.after(0, lambda s=shortcut, t=text_to_expand: self._process_and_replace_on_main_thread(s, t))

                        self.current_input = "" # Reset buffer after expansion
                        return # Stop checking other shortcuts, one expansion per trigger
        except Exception as e:
            log(f"Error in on_key_press: {e}")
            self.current_input = "" # Reset buffer on any error


    def _process_and_replace_on_main_thread(self, shortcut, text_to_expand):
        """Helper to run placeholder processing and replacement on the main thread."""
        try:
            processed_text = self.process_placeholders(text_to_expand)
            self.replace_text(shortcut, processed_text)

            if shortcut not in self.snippet_history:
                self.snippet_history.append(shortcut)
                if len(self.snippet_history) > 10:
                    self.snippet_history.pop(0)
            self.status_var.set(f"Expanded: {shortcut}")
        except Exception as e:
            log(f"Error processing/replacing on main thread for shortcut '{shortcut}': {e}")
            self.status_var.set(f"Error expanding: {shortcut}")
            self.status_indicator.config(foreground=self.theme["error"])


    def replace_text(self, shortcut, replacement):
        """
        Replaces the typed shortcut with the expanded text.
        Uses Shift+Left to select the shortcut, then Delete/Backspace, then paste.
        """
        if not KEYBOARD_AVAILABLE:
            return

        kb_controller = keyboard.Controller()

        # Step 1: Select the typed shortcut
        # Simulate pressing Left_arrow `len(shortcut)` times while holding Shift
        with kb_controller.pressed(keyboard.Key.shift):
            for _ in range(len(shortcut)):
                kb_controller.press(keyboard.Key.left)
                kb_controller.release(keyboard.Key.left)
                time.sleep(0.005) # Small delay for each key press

        time.sleep(0.02) # Give system time to register selection

        # Step 2: Delete the selected shortcut
        kb_controller.press(keyboard.Key.backspace) # or keyboard.Key.delete
        kb_controller.release(keyboard.Key.backspace)
        time.sleep(0.02) # Give system time to register deletion

        # Step 3: Paste the replacement text
        if CLIPBOARD_AVAILABLE:
            original_clipboard_content = None
            try:
                # Save original clipboard content if possible
                original_clipboard_content = pyperclip.paste()
            except Exception as e:
                log(f"Could not read original clipboard: {e}")

            try:
                pyperclip.copy(replacement)
                time.sleep(0.05) # Ensure clipboard has the new content

                # Simulate Ctrl+V (Cmd+V on macOS)
                if sys.platform == "darwin": # macOS
                    with kb_controller.pressed(keyboard.Key.cmd):
                        kb_controller.press('v')
                        kb_controller.release('v')
                else: # Windows/Linux
                    with kb_controller.pressed(keyboard.Key.ctrl):
                        kb_controller.press('v')
                        kb_controller.release('v')

                time.sleep(0.05) # Allow paste to happen

                # Restore original clipboard content
                if original_clipboard_content is not None:
                    try:
                        pyperclip.copy(original_clipboard_content)
                    except Exception as e:
                        log(f"Could not restore clipboard: {e}")

            except Exception as e:
                log(f"Clipboard operation failed during replacement: {e}")
                self.root.after(0, lambda: self.status_var.set(f"Clipboard error expanding {shortcut}"))
                self.root.after(0, lambda: self.status_indicator.config(foreground=self.theme["error"]))
        else:
            log("Clipboard module not available for replacement. Typing instead (slow).")
            # Fallback to typing if clipboard is unavailable. This is generally slow and unreliable for long text.
            kb_controller.type(replacement)
            self.root.after(0, lambda: self.status_var.set("Clipboard N/A, typed expansion"))
            self.root.after(0, lambda: self.status_indicator.config(foreground=self.theme["warning"]))

        log(f"Expanded shortcut: {shortcut}")


    def show_error(self, title, message):
        """Displays an error message box."""
        messagebox.showerror(title, message, parent=self.root)

    def show_info(self, title, message):
        """Displays an informational message box."""
        messagebox.showinfo(title, message, parent=self.root)

    def show_confirm(self, title, message):
        """Displays a confirmation message box and returns True if confirmed."""
        return messagebox.askyesno(title, message, parent=self.root)

    def check_auto_backup(self):
        """Checks if automatic backup is due and performs it if necessary."""
        if not self.config_manager.get("auto_backup", True):
            return

        last_backup_str = self.config_manager.get("last_backup")
        if last_backup_str:
            last_backup_date = datetime.datetime.fromisoformat(last_backup_str)
            interval_days = self.config_manager.get("backup_interval_days", 7)
            if (datetime.datetime.now() - last_backup_date).days >= interval_days:
                log("Auto backup due. Performing backup.")
                self.backup_data(silent=True)
        else:
            log("No last backup recorded. Performing initial auto backup.")
            self.backup_data(silent=True)
        self.root.after(86400000, self.check_auto_backup)


    def setup_tooltips(self):
        """Initializes tooltips for relevant UI elements."""
        # This function is called when tooltips are enabled or theme changes.
        # It needs to ensure all current widgets have their tooltips set up/updated.
        # Iterating through notebook tabs and their children is necessary here.

        # Clear existing tooltips explicitly to prevent lingering issues before re-creating
        for widget, tooltip_obj in list(self.tooltips.items()):
            if tooltip_obj:
                tooltip_obj.hide()
            try: # Attempt to unbind old events
                widget.unbind("<Enter>")
                widget.unbind("<Leave>")
                widget.unbind("<ButtonPress>")
            except tk.TclError: # Widget might have been destroyed
                pass
            del self.tooltips[widget] # Remove from tracking list

        # Re-create tooltips for all relevant widgets
        # (These are typically created during the tab creation methods, but this ensures they're hooked up if state changes)
        # Re-run tooltip creation for existing widgets by calling the tab creation methods (if safe)
        # OR, manually identify key widgets and call create_tooltip on them.
        # Given the previous structure, the `create_snippets_tab`, etc. already call `self.create_tooltip`.
        # The key is to ensure `self.create_tooltip` is robust to being called multiple times.
        # The current `create_tooltip` does check `if widget not in self.tooltips`.

        # Instead of calling entire tab creation, which can be complex and recreate widgets,
        # we can target specific elements. For now, assume widgets that need tooltips are already created
        # and create_tooltip is called on them when they are built.
        # The `toggle_tooltips` method now drives the `create_tooltip` directly.
        pass # The individual `create_tooltip` calls within tab creation are sufficient.


    def create_tooltip(self, widget, text):
        """Creates and stores a tooltip for a given widget."""
        if self.config_manager.get("show_tooltips", True):
            if widget not in self.tooltips:
                self.tooltips[widget] = ToolTip(widget, text, self.theme)
            else: # If tooltip already exists for this widget, just update its properties and re-bind
                self.tooltips[widget].text = text
                self.tooltips[widget].theme = self.theme
                self.tooltips[widget].widget.bind("<Enter>", self.tooltips[widget].schedule_show, add="+")
                self.tooltips[widget].widget.bind("<Leave>", self.tooltips[widget].hide, add="+")
                self.tooltips[widget].widget.bind("<ButtonPress>", self.tooltips[widget].hide, add="+")


    def check_dependencies(self):
        """Checks if optional dependencies are installed and updates status."""
        log(f"Dependency check: Keyboard available={KEYBOARD_AVAILABLE}, Clipboard available={CLIPBOARD_AVAILABLE}, Systray available={SYSTRAY_AVAILABLE}")

    def import_snippets(self):
        """Handles importing snippets from a JSON file."""
        filepath = filedialog.askopenfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Import Snippets from JSON",
        )
        if not filepath:
            return

        try:
            with open(filepath, "r", encoding="utf-8") as f:
                imported_data = json.load(f)

            if isinstance(imported_data, dict) and "snippets" in imported_data:
                imported_snippets = imported_data["snippets"]
            elif isinstance(imported_data, dict):
                imported_snippets = imported_data
            else:
                self.show_error("Import Error", "Invalid file format. Expected a JSON object with 'snippets' key or direct snippet mapping.")
                return

            num_imported = 0
            num_overwritten = 0
            for shortcut, data in imported_snippets.items():
                if isinstance(data, dict) and "text" in data:
                    if self.snippet_manager.get_snippet(shortcut):
                        num_overwritten += 1
                    else:
                        num_imported += 1
                    self.snippet_manager.add_update_snippet(
                        shortcut, data["text"], data.get("category", "General"), data.get("description", "")
                    )
                elif isinstance(data, str):
                    if self.snippet_manager.get_snippet(shortcut):
                        num_overwritten += 1
                    else:
                        num_imported += 1
                    self.snippet_manager.add_update_snippet(
                        shortcut, data, "General", ""
                    )
                else:
                    log(f"Skipping invalid snippet during import: {shortcut} - not a dict or string.")

            if self.snippet_manager.save_snippets():
                self.refresh_snippet_list()
                self.refresh_categories()
                self.show_info(
                    "Import Successful",
                    f"Imported {num_imported + num_overwritten} snippets.\n({num_imported} new, {num_overwritten} overwritten).",
                )
            else:
                self.show_error("Import Error", "Failed to save snippets after import.")

        except Exception as e:
            log(f"Error importing snippets: {e}")
            self.show_error("Import Error", f"Could not import snippets: {e}")

    def export_snippets(self):
        """Handles exporting all snippets to a JSON file."""
        filepath = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="Export Snippets to JSON",
            initialfile="text_expander_snippets.json",
        )
        if not filepath:
            return

        try:
            data_to_export = {
                "config": self.config_manager.config,
                "snippets": self.snippet_manager.get_all_snippets()
            }
            with open(filepath, "w", encoding="utf-8") as f:
                json.dump(data_to_export, f, ensure_ascii=False, indent=4)
            self.show_info("Export Successful", f"All data exported to {filepath}")
        except Exception as e:
            log(f"Error exporting snippets: {e}")
            self.show_error("Export Error", f"Could not export snippets: {e}")

    def backup_data(self, silent=False):
        """Performs a backup of both snippets and configuration."""
        backup_dir = self.config_manager.get("backup_location", APP_DIR)
        os.makedirs(backup_dir, exist_ok=True)
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_filename = f"text_expander_backup_{timestamp}.json"
        filepath = os.path.join(backup_dir, backup_filename)

        try:
            data_to_backup = {
                "config": self.config_manager.config,
                "snippets": self.snippet_manager.get_all_snippets()
            }
            with open(filepath, "w", encoding="utf-8") as f:
                json.dump(data_to_backup, f, ensure_ascii=False, indent=4)

            self.config_manager.set("last_backup", datetime.datetime.now().isoformat())
            if hasattr(self, "last_backup_var"):
                self.last_backup_var.set(
                    f"Last backup: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}"
                )
            if not silent:
                self.show_info("Backup Successful", f"Data backed up to {filepath}")
            log(f"Data backed up to {filepath}")
        except Exception as e:
            log(f"Error backing up data: {e}")
            if not silent:
                self.show_error("Backup Error", f"Could not backup data: {e}")

    def restore_backup(self):
        """Restores data from a selected backup file."""
        filepath = filedialog.askopenfilename(
            defaultextension=".json",
            filetypes=[("JSON backup files", "*.json"), ("All files", "*.*")],
            title="Restore from Backup",
        )
        if not filepath:
            return

        if not self.show_confirm(
            "Confirm Restore",
            "Restoring from backup will overwrite current snippets and settings. Continue?",
        ):
            return

        try:
            with open(filepath, "r", encoding="utf-8") as f:
                backup_data = json.load(f)

            if "snippets" in backup_data:
                self.snippet_manager.snippets = backup_data["snippets"]
                self.snippet_manager.save_snippets()

                if "config" in backup_data:
                    current_theme = self.config_manager.get("theme")
                    self.config_manager.config.update(backup_data["config"])
                    self.config_manager.set("theme", current_theme)
                    self.config = self.config_manager.config
                    self.change_theme(current_theme)

                self.refresh_snippet_list()
                self.refresh_categories()
                self.show_info("Restore Successful", "Data restored from backup.")
                self._update_settings_ui_from_config()
            else:
                self.show_error("Restore Error", "Invalid backup file format. Missing 'snippets' key.")
        except Exception as e:
            log(f"Error restoring from backup: {e}")
            self.show_error("Restore Error", f"Could not restore data: {e}")

    def _update_settings_ui_from_config(self):
        """Updates the settings tab UI elements to reflect current config."""
        self.tray_var.set(self.config_manager.get("minimize_to_tray"))
        self.minimized_var.set(self.config_manager.get("start_minimized"))
        self.tooltips_var.set(self.config_manager.get("show_tooltips"))
        self.default_category_var.set(self.config_manager.get("default_category"))
        self.auto_backup_var.set(str(self.config_manager.get("auto_backup")))
        self.backup_interval_var.set(str(self.config_manager.get("backup_interval_days")))
        self.backup_location_var.set(self.config_manager.get("backup_location", APP_DIR))
        last_backup = self.config_manager.get("last_backup")
        last_backup_text = ("Never" if not last_backup else datetime.datetime.fromisoformat(last_backup).strftime("%Y-%m-%d %H:%M"))
        self.last_backup_var.set(f"Last backup: {last_backup_text}")
        self.theme_var.set(self.config_manager.get("theme"))
        self.refresh_categories()

    def toggle_tooltips(self):
        """Toggles the visibility of tooltips."""
        show = self.tooltips_var.get()
        self.config_manager.set("show_tooltips", show)
        if show:
            for widget_key in self.tooltips.keys():
                self.create_tooltip(widget_key, self.tooltips[widget_key].text)
            self.status_var.set("Tooltips enabled.")
        else:
            for widget, tooltip_obj in list(self.tooltips.items()):
                if tooltip_obj:
                    tooltip_obj.hide()
                try:
                    widget.unbind("<Enter>")
                    widget.unbind("<Leave>")
                    widget.unbind("<ButtonPress>")
                except tk.TclError:
                    pass
            self.status_var.set("Tooltips disabled.")

    def open_resource(self, resource_title):
        """Opens various resources based on title."""
        log(f"Opening resource: {resource_title}")
        if resource_title == "About":
            self.show_about()
        elif resource_title == "Check for Updates":
            self.show_info("Check for Updates", "This feature is not yet implemented.")

    def show_about(self):
        """Displays the 'About' dialog."""
        messagebox.showinfo(
            "About Text Expander",
            "Text Expander v2.0\n\n"
            "A modern text expansion tool.\n"
            "Developed with Python and Tkinter.",
            parent=self.root,
        )

    def update_backup_interval_from_spinbox(self):
        """Updates backup interval from spinbox, triggered by value change."""
        try:
            val = int(self.backup_interval_var.get())
            if val < 1: val = 1
            if val > 30: val = 30
            self.config_manager.set("backup_interval_days", val)
            self.backup_interval_var.set(str(val))
        except ValueError:
            self.backup_interval_var.set(str(self.config_manager.get("backup_interval_days", 7)))

    def update_backup_interval_from_var(self, *args):
        """Updates backup interval from stringvar, triggered by trace."""
        self.update_backup_interval_from_spinbox()

    def browse_backup_location(self):
        """Opens a dialog to select a backup directory."""
        directory = filedialog.askdirectory(
            title="Select Backup Location",
            initialdir=self.config_manager.get("backup_location", APP_DIR),
        )
        if directory:
            self.backup_location_var.set(directory)
            self.config_manager.set("backup_location", directory)