import os
import sys
import streamlit.web.cli as stcli
import subprocess
import time
import webbrowser

def main():
    if getattr(sys, 'frozen', False):
        bundle_dir = sys._MEIPASS
    else:
        bundle_dir = os.path.dirname(os.path.abspath(__file__))
    
    app_path = os.path.join(bundle_dir, "app.py")
    
    # We set sys.argv so streamlit picks it up
    sys.argv = ["streamlit", "run", app_path, "--server.headless=false", "--server.port=8599", "--global.developmentMode=false"]
    
    # Exit using streamlit main
    sys.exit(stcli.main())

if __name__ == "__main__":
    main()
