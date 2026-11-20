# -*- coding: utf-8 -*-
import os
import sys
import traceback

root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)

try:
    from app.main import app, VercelPathMiddleware

    if not isinstance(app.wsgi_app, VercelPathMiddleware):
        app.wsgi_app = VercelPathMiddleware(app.wsgi_app)

    class VercelErrorCatcher:
        """Bắt ngoại lệ tại tầng WSGI để trả về giao diện lỗi kèm Traceback chi tiết."""
        def __init__(self, wsgi_app):
            self.wsgi_app = wsgi_app

        def __call__(self, environ, start_response):
            try:
                return self.wsgi_app(environ, start_response)
            except Exception:
                err = traceback.format_exc()
                print("CRITICAL WSGI EXCEPTION:\n", err, file=sys.stderr)
                html = f"""<!DOCTYPE html>
<html>
<head><meta charset='utf-8'><title>500 Application Error</title></head>
<body style='font-family: monospace; padding: 24px; background: #fff5f5; color: #b91c1c;'>
    <h2 style='color: #dc2626;'>Lỗi thực thi máy chủ (Server Execution Error):</h2>
    <pre style='background: #fee2e2; padding: 16px; border-radius: 8px; border: 1px solid #f87171; overflow-x: auto;'>{err}</pre>
</body>
</html>""".encode("utf-8")
                start_response("500 Internal Server Error", [
                    ("Content-Type", "text/html; charset=utf-8"),
                    ("Content-Length", str(len(html)))
                ])
                return [html]

    app.wsgi_app = VercelErrorCatcher(app.wsgi_app)

except Exception:
    startup_err = traceback.format_exc()
    print("STARTUP ERROR IN api/index.py:\n", startup_err, file=sys.stderr)
    from flask import Flask
    app = Flask(__name__)

    @app.route("/", defaults={"path": ""})
    @app.route("/<path:path>")
    def error_handler(path):
        return f"""<!DOCTYPE html>
<html>
<head><meta charset='utf-8'><title>500 Startup Error</title></head>
<body style='font-family: monospace; padding: 24px; background: #fff5f5; color: #b91c1c;'>
    <h2 style='color: #dc2626;'>Lỗi khởi động ứng dụng (Startup Error on Vercel):</h2>
    <pre style='background: #fee2e2; padding: 16px; border-radius: 8px; border: 1px solid #f87171; overflow-x: auto;'>{startup_err}</pre>
</body>
</html>""", 500, {"Content-Type": "text/html; charset=utf-8"}
