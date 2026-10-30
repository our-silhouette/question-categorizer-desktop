"""
Desktop Application Launcher for Question Categorizer
-----------------------------------------------------
This script acts as the entry point when packaged as a standalone desktop executable (PyInstaller).
It performs the following:
1. Disables Windows QuickEdit mode to prevent accidental mouse clicks from freezing the process.
2. Locates the bundled Streamlit application files (app.py, pipeline.py) inside the standalone package.
3. Finds an open network port (default: 8501).
4. Spawns a background thread that polls for server readiness.
5. Automatically opens the user's default browser (Chrome, Edge, Safari) at http://localhost:<port>.
6. Displays a friendly, helpful console status banner so users know not to close it prematurely.
7. Boots Streamlit in headless, non-interactive mode (disabling analytics, telemetry, and email prompts).
"""

import os
import sys
import time
import socket
import threading
import webbrowser

def disable_windows_quickedit():
    """
    On Windows, clicking inside a console window can enter 'QuickEdit Mode',
    which pauses console output and suspends Python until the user presses Enter.
    This function programmatically disables QuickEdit mode so user mouse clicks never freeze the app.
    """
    if sys.platform == "win32":
        try:
            import ctypes
            kernel32 = ctypes.windll.kernel32
            h_stdin = kernel32.GetStdHandle(-10)  # STD_INPUT_HANDLE = -10
            mode = ctypes.c_ulong()
            if kernel32.GetConsoleMode(h_stdin, ctypes.byref(mode)):
                # ENABLE_QUICK_EDIT_MODE = 0x0040, ENABLE_EXTENDED_FLAGS = 0x0080
                new_mode = (mode.value & ~0x0040) | 0x0080
                kernel32.SetConsoleMode(h_stdin, new_mode)
        except Exception:
            pass

def find_free_port(preferred_port=8501):
    """Checks if the preferred port is free; if occupied, finds the next available port."""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.bind(("127.0.0.1", preferred_port))
            return preferred_port
    except OSError:
        # Preferred port in use, let the OS allocate an ephemeral port
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.bind(("127.0.0.1", 0))
            return s.getsockname()[1]

def wait_for_server_and_open_browser(port, timeout=60):
    """Polls the local server until it accepts TCP connections, then opens the default web browser."""
    start_time = time.time()
    url = f"http://localhost:{port}"
    while time.time() - start_time < timeout:
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=0.5):
                # Server is accepting connections; give Streamlit a tiny moment to settle
                time.sleep(0.5)
                webbrowser.open_new_tab(url)
                return
        except (OSError, ConnectionRefusedError):
            time.sleep(0.3)
    
    # Fallback: attempt to open even if health check timed out
    webbrowser.open_new_tab(url)

def get_base_dir():
    """Returns the base directory containing the application assets.
    When running as a PyInstaller bundle, sys._MEIPASS holds the extracted resources.
    When running directly from source, returns the directory containing this script.
    """
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))

def print_banner(port):
    banner = f"""
================================================================================
  📚  QUESTION CATEGORIZER - DESKTOP SERVER
================================================================================
  [✓] Local server initialized on: http://localhost:{port}
  [✓] Your default web browser will open automatically in a moment.

  IMPORTANT TIPS FOR USERS:
  • Please KEEP THIS CONSOLE WINDOW OPEN while using the application.
  • You can minimize this window to your taskbar/dock.
  • To exit the application: Close your browser tab and close this window,
    or press Ctrl + C here.
================================================================================
"""
    print(banner, flush=True)

def main():
    # 1. Prevent Windows mouse clicks from pausing the console
    disable_windows_quickedit()

    base_dir = get_base_dir()
    
    # 2. Locate app.py
    app_path = os.path.join(base_dir, "app.py")
    if not os.path.exists(app_path):
        dev_path = os.path.join(base_dir, "..", "Streamlit approach", "app.py")
        if os.path.exists(dev_path):
            app_path = os.path.abspath(dev_path)
        else:
            print(f"Error: Could not locate app.py in {base_dir}")
            sys.exit(1)

    port = find_free_port(8501)
    
    # 3. Configure Streamlit environment variables for a seamless desktop experience
    os.environ["STREAMLIT_SERVER_PORT"] = str(port)
    os.environ["STREAMLIT_SERVER_ADDRESS"] = "127.0.0.1"
    os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
    os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
    os.environ["STREAMLIT_SERVER_ENABLE_CORS"] = "false"
    os.environ["STREAMLIT_SERVER_ENABLE_XSRF_PROTECTION"] = "false"
    os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"

    # 4. Display user-friendly banner
    print_banner(port)

    # 5. Start browser launcher thread in background
    browser_thread = threading.Thread(
        target=wait_for_server_and_open_browser,
        args=(port,),
        daemon=True
    )
    browser_thread.start()

    # 6. Import Streamlit CLI and run the application
    from streamlit.web import cli as stcli
    sys.argv = [
        "streamlit",
        "run",
        app_path,
        f"--server.port={port}",
        "--server.address=127.0.0.1",
        "--server.headless=true",
        "--browser.gatherUsageStats=false",
        "--server.enableCORS=false",
        "--server.enableXsrfProtection=false",
        "--global.developmentMode=false"
    ]
    stcli.main()

if __name__ == "__main__":
    main()
