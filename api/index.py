# -*- coding: utf-8 -*-
import os
import sys
from pathlib import Path
import traceback

root_dir = str(Path(__file__).resolve().parent.parent)
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)
if os.getcwd() not in sys.path:
    sys.path.insert(0, os.getcwd())

try:
    from app.main import app
except Exception as exc:
    startup_traceback = traceback.format_exc()
    print("STARTUP IMPORT ERROR:\n", startup_traceback, file=sys.stderr)

    from flask import Flask
    app = Flask(__name__)

    @app.route('/', defaults={'path': ''})
    @app.route('/<path:path>')
    def catch_all_error(path):
        return f"""<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><title>Vercel Python Diagnostic</title></head>
<body style="font-family: monospace; padding: 24px; background: #fff5f5; color: #b91c1c;">
    <h2 style="color: #dc2626;">🚨 Phát hiện lỗi khi khởi chạy ứng dụng trên Vercel:</h2>
    <p><b>Chi tiết lỗi Python (Traceback):</b></p>
    <pre style="background: #fee2e2; padding: 16px; border-radius: 8px; border: 1px solid #f87171; overflow-x: auto; white-space: pre-wrap;">{startup_traceback}</pre>
</body>
</html>""", 200, {"Content-Type": "text/html; charset=utf-8"}

if __name__ == "__main__":
    app.run()
