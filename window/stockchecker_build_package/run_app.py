import os
import sys
from streamlit.web import cli as stcli

# Make paths work both when running as .py and when frozen by PyInstaller
if getattr(sys, "frozen", False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

APP_FILE = os.path.join(BASE_DIR, "Stockchecker.py")

sys.argv = [
    "streamlit",
    "run",
    APP_FILE,
    "--global.developmentMode=false",
    "--server.headless=true",
]

sys.exit(stcli.main())
