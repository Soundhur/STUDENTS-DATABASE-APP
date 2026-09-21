import webview
import subprocess
import time
import socket
import sys
import os

def get_free_port():
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.bind(("", 0))
    port = s.getsockname()[1]
    s.close()
    return port

if __name__ == '__main__':
    port = get_free_port()
    
    # Hide the console window on Windows
    startupinfo = None
    if os.name == 'nt':
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    
    # Start streamlit server in background
    process = subprocess.Popen(
        [sys.executable, "-m", "streamlit", "run", "app.py", "--server.port", str(port), "--server.headless", "true"],
        startupinfo=startupinfo,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL
    )
    
    # Give the server a moment to spin up
    time.sleep(2.5)
    
    # Create the native desktop window
    window = webview.create_window(
        "MCA Student Portal", 
        f"http://localhost:{port}",
        width=1200,
        height=800,
        min_size=(800, 600)
    )
    
    webview.start(private_mode=False)
    
    # Cleanup: When the user closes the window, kill the Streamlit server
    process.kill()
