#!/bin/sh
# Horizon Airways - macOS / Linux launcher
# Run:  chmod +x run_mac_linux.sh && ./run_mac_linux.sh
cd "$(dirname "$0")"
python3 -m pip install -r requirements.txt
python3 app.py
