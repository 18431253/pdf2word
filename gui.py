import sys
import os
import threading
import webbrowser
import time

# When packaged by PyInstaller, resources are in sys._MEIPASS
if getattr(sys, 'frozen', False):
    os.chdir(sys._MEIPASS)

from app import app

def _open_browser():
    time.sleep(1.5)
    webbrowser.open("http://127.0.0.1:8765")

if __name__ == "__main__":
    threading.Thread(target=_open_browser, daemon=True).start()
    app.run(debug=False, host="127.0.0.1", port=8765)
