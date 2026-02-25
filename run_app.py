import streamlit.web.cli as stcli
import os, sys
import subprocess
import threading
import time
import webbrowser
import signal

def resolve_path(path):
    if getattr(sys, "frozen", False):
        basedir = sys._MEIPASS
    else:
        basedir = os.path.dirname(__file__)
    return os.path.join(basedir, path)

def launch_app_window():
    """Waits for server, then opens a 'No-UI' window"""
    time.sleep(3)  # Wait for Streamlit to start
    url = "http://localhost:8501"
    
    # Try Chrome in App Mode
    try:
        # We use a subprocess that we can track if needed
        subprocess.Popen(f'start chrome --app="{url}"', shell=True)
    except:
        webbrowser.open(url)

def kill_system():
    """Forcefully kills the entire app process tree when closing"""
    print("Shutting down system...")
    # This kills the current process and all its children (the server)
    os.kill(os.getpid(), signal.SIGTERM)

if __name__ == "__main__":
    # 1. SETUP
    os.environ["STREAMLIT_SERVER_HEADLESS"] = "true"
    os.environ["STREAMLIT_BROWSER_GATHER_USAGE_STATS"] = "false"
    
    # 2. START BROWSER
    threading.Thread(target=launch_app_window, daemon=True).start()
    
    # 3. RUN STREAMLIT
    # We wrap this in a try/finally to ensure if the script ends, we kill everything
    try:
        app_path = resolve_path("app.py")
        sys.argv = ["streamlit", "run", app_path, "--global.developmentMode=false"]
        stcli.main()
    finally:
        kill_system()
