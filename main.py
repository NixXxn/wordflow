import tkinter as tk
from text_expander_app import ModernTextExpander
from utils import log, APP_DIR, LOG_FILE
import sys
import os
import ctypes
from tkinter import messagebox

def main():
    try:
        log("Application starting from main.py")
        root = tk.Tk()
        app = ModernTextExpander(root)

        root.title("Text Expander")
        if os.path.exists(app.icon_path):
            try:
                root.iconbitmap(app.icon_path)
            except Exception as e:
                log(f"Main: Error setting window icon: {e}")

        # Apply platform-specific UI enhancements
        try:
            if sys.platform == "win32":
                myappid = "textexpander.app.2.0"  # arbitrary string
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except Exception as e:
            log(f"Platform-specific UI enhancement error: {e}")

        # Warn about missing modules on startup
        from constants import KEYBOARD_AVAILABLE, CLIPBOARD_AVAILABLE, SYSTRAY_AVAILABLE
        missing = []
        if not KEYBOARD_AVAILABLE:
            missing.append("Keyboard module (pynput)")
        if not CLIPBOARD_AVAILABLE:
            missing.append("Clipboard module (pyperclip)")
        if not SYSTRAY_AVAILABLE:
            missing.append("System tray support (pystray, PIL)")

        if missing:
            message = "The following features are not available:\n\n" + "\n".join(
                f"• {m}" for m in missing
            )
            message += "\n\nRefer to documentation for installation instructions."
            messagebox.showwarning("Missing Dependencies", message)

        root.mainloop()
    except Exception as e:
        log(f"Critical error in main application loop: {e}")
        try:
            tk.messagebox.showerror("Error", f"An error occurred: {e}")
        except:
            print(f"Error: {e}")


if __name__ == "__main__":
    main()