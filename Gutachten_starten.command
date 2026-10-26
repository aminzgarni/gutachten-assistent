#!/bin/bash
cd "$(dirname "$0")"
nohup python3 programm/web.py >/dev/null 2>&1 &
sleep 1
osascript -e 'tell application "Terminal" to close front window' >/dev/null 2>&1 &
exit 0
