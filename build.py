import os
import sys
import subprocess
import shutil
import platform

def build_app():
    """
    Builds the Text Expander application into a standalone executable using PyInstaller.
    """
    print(f"Detected OS: {platform.system()}")

    # Define the main script
    main_script = "main.py"
    app_name = "Wordflow"

    # Base PyInstaller command
    pyinstaller_cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconsole",          # Don't show a terminal window
        "--onefile",            # Bundle everything into a single file
        "--name", app_name,     # Name of the executable
        "--clean",              # Clean cache before building
    ]

    # Handle imports that might be hidden or dynamic
    # pynput and pystray often need specific backend handling
    hidden_imports = [
        "pynput.keyboard._xorg",
        "pynput.mouse._xorg",
        "pynput.keyboard._win32",
        "pynput.mouse._win32",
        "pynput.keyboard._darwin",
        "pynput.mouse._darwin",
        "pystray._util",
        "pystray._win32",
        "pystray._appindicator",
        "pystray._gtk",
        "pystray._darwin",
        "PIL",
        "PIL._tkinter_finder"
    ]

    for imp in hidden_imports:
        pyinstaller_cmd.extend(["--hidden-import", imp])

    # Add data files if necessary (though the app generates most at runtime in APP_DIR)
    # If we had a default config or initial icon to bundle, we'd use --add-data
    # pyinstaller_cmd.extend(["--add-data", "icon.ico;."])

    # Main script
    pyinstaller_cmd.append(main_script)

    print("Running PyInstaller...")
    print("Command:", " ".join(pyinstaller_cmd))

    try:
        subprocess.check_call(pyinstaller_cmd)
        print("\n" + "="*50)
        print("Build Successful!")
        print("="*50)

        dist_dir = os.path.join(os.getcwd(), "dist")
        if platform.system() == "Windows":
            executable = os.path.join(dist_dir, f"{app_name}.exe")
        else:
            executable = os.path.join(dist_dir, app_name)

        print(f"Your standalone application is located at:\n{executable}")
        print("You can move this file anywhere and run it.")

    except subprocess.CalledProcessError as e:
        print("\n" + "="*50)
        print("Build Failed!")
        print(f"Error: {e}")
        print("="*50)
        sys.exit(1)

if __name__ == "__main__":
    # Ensure dependencies are installed
    try:
        import PyInstaller
    except ImportError:
        print("PyInstaller not found. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])

    build_app()
