# Wordflow

Wordflow is a modern, Python-based text expander application designed to boost your productivity. It allows you to create abbreviations (shortcuts) that automatically expand into longer text snippets, saving you time and reducing repetitive typing. It supports dynamic placeholders for dates, times, clipboard content, and even custom input fields.

## Key Features

- **Text Expansion**: Automatically replaces short abbreviations with full text snippets.
- **Dynamic Placeholders**: Insert current date, time, clipboard content, or random numbers.
- **Custom Input Forms**: Create snippets that ask for user input (e.g., client names) before expanding.
- **Snippet Management**: Organize snippets into categories, search, and filter easily.
- **Modern UI**: Clean, light-themed interface built with Tkinter.
- **System Tray Support**: Minimizes to the system tray to run quietly in the background.
- **Backup & Restore**: Auto-backup feature and manual import/export options.
- **Cross-Platform**: Works on Windows and macOS.

## Installation Instructions

### Prerequisites

- **Python 3.x**: Ensure you have Python installed on your system. You can download it from [python.org](https://www.python.org/downloads/).

### Windows

1.  **Download the Source Code**:
    Clone this repository or download the ZIP file and extract it to a folder.

2.  **Open Command Prompt**:
    Navigate to the project folder where you extracted the files.
    ```cmd
    cd path\to\wordflow
    ```

3.  **Install Dependencies**:
    Run the following command to install the required libraries:
    ```cmd
    pip install -r requirements.txt
    ```

4.  **Run the Application**:
    Start the application by running:
    ```cmd
    python main.py
    ```

### macOS

1.  **Download the Source Code**:
    Clone this repository or download the ZIP file and extract it.

2.  **Open Terminal**:
    Navigate to the project folder.
    ```bash
    cd path/to/wordflow
    ```

3.  **Install Dependencies**:
    Run the following command (you might need to use `pip3` depending on your setup):
    ```bash
    pip3 install -r requirements.txt
    ```

4.  **Run the Application**:
    Start the application by running:
    ```bash
    python3 main.py
    ```

    **Important Note for macOS Users**:
    Because Wordflow monitors keyboard input to detect shortcuts, you must grant it **Accessibility permissions**.
    - When you first run the app, macOS may prompt you to allow "Terminal" (or your Python environment) to control your computer.
    - Go to **System Settings > Privacy & Security > Accessibility** and enable the switch for your terminal or Python application.
    - If the app doesn't expand text, restart it after granting permissions.

## Usage

1.  **Create a Snippet**: Click "New Snippet", enter a shortcut (e.g., `/email`) and the content (e.g., `my.email@example.com`).
2.  **Save**: Click "Save".
3.  **Type**: Open any other application (Notepad, Browser, etc.) and type `/email`. It will instantly change to your email address.

## Building a Standalone Application

If you prefer a standalone executable (an `.exe` file on Windows or an App file on macOS) instead of running the source code, you can build it yourself using the included script.

1.  **Install Dependencies**:
    Ensure you have installed the requirements as described above:
    ```bash
    pip install -r requirements.txt
    ```

2.  **Run the Build Script**:
    Run the following command:
    ```bash
    python build.py
    ```

3.  **Locate the Application**:
    Once the build finishes successfully, check the `dist/` folder in your project directory.
    -   **Windows**: You will see `Wordflow.exe`.
    -   **macOS/Linux**: You will see a `Wordflow` executable.

    You can move this file anywhere on your computer and run it directly.
