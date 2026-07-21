#!/bin/bash
cd "$(dirname "$0")"
pip3 install -q flask openpyxl flask-cors requests 2>/dev/null
echo "🚀 启动 i18n Translator..."
python3 server.py &
sleep 1
open http://localhost:8080
wait
