"""Used by Start_Order_Support_App.bat: opens the browser once the local app answers on its port."""
import socket
import sys
import time
import webbrowser

port = int(sys.argv[1]) if len(sys.argv) > 1 else 8501
for _ in range(120):
    try:
        socket.create_connection(("localhost", port), timeout=1).close()
    except OSError:
        time.sleep(0.5)
    else:
        webbrowser.open(f"http://localhost:{port}")
        break
