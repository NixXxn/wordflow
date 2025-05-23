# utils.py
import datetime
import os
import sys

# Assume APP_DIR and LOG_FILE are defined in constants.py and imported
# For a truly independent utils, these would be passed or determined here,
# but for this project, they are part of common constants.
from constants import APP_DIR, LOG_FILE

def log(message):
    """
    Logs a timestamped message to the application's log file.
    Includes basic error handling for logging itself.
    """
    try:
        # Ensure the directory exists before attempting to write the log file
        os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"{timestamp} - {message}\n")
    except Exception as e:
        print(f"Error writing to log file ({LOG_FILE}): {e}", file=sys.stderr)

# Placeholder for dependency check, actual check moved to constants for clarity
# of module availability at import time.
def check_dependencies_status():
    """
    Returns a dictionary indicating the availability of key optional modules.
    (This is now largely handled by constants.py at import time, but this
    function could be expanded for more granular checks or versioning).
    """
    from constants import KEYBOARD_AVAILABLE, CLIPBOARD_AVAILABLE, SYSTRAY_AVAILABLE
    return {
        "keyboard_available": KEYBOARD_AVAILABLE,
        "clipboard_available": CLIPBOARD_AVAILABLE,
        "systray_available": SYSTRAY_AVAILABLE,
    }